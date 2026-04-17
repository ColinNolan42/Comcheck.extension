# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Key findings from community research (pyRevit Forums, Jan 2026):

  1. doc.Regenerate() REQUIRED after AddLeader.
     Without this, SetLeader silently fails or uses stale geometry.
     The leader object returned by GetLeader before Regenerate has
     incorrect internal state — elbow moves are ignored or error.

  2. Perpendicular direction from Curve.Direction cross product.
     Instead of hardcoding +X/-X/+Y/-Y per orientation, compute the
     true perpendicular to the grid in the view plane:
       tan  = grid.Curve.Direction.Normalize()
       perp = XYZ(-tan.Y, tan.X, 0)
     This works for vertical, horizontal, and diagonal grids in any
     Revit version without assumptions about direction.

  3. Use leader.Anchor for collision detection, not curve endpoints.
     After AddLeader+Regenerate, Anchor reflects the actual bubble
     position including any existing leader offset. Curve endpoints
     only reflect the raw grid line end, not where the bubble sits.

  4. Sign of perpendicular direction from dot product.
     To move a bubble AWAY from its colliding neighbour, compute:
       clash_vec = anchor_here - anchor_other
       sign = +1 if clash_vec.DotProduct(perp) >= 0 else -1
     This universally picks the correct away-direction regardless of
     grid orientation or which side the neighbour is on.

  5. Reset elbow to grid line before nudging.
     Project the default elbow onto the grid curve to start from a
     clean position: curve.Project(leader.Elbow).XYZPoint
     Then nudge from there.

Workflow per view:
  1. Collect all host grids with visible bubbles
  2. Ensure all colliding grids have leaders (AddLeader)
  3. doc.Regenerate() once after all AddLeader calls
  4. Reset new leader elbows to the grid line
  5. Iteratively detect collisions using Anchor positions
  6. For each collision, nudge the target's Elbow perpendicular
     away from the neighbour — small increment per iteration
  7. Repeat until no collisions remain or max iterations hit

Host grids only. Linked grids excluded.
FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "15.0.0"
__doc__     = ("Separates colliding grid bubbles using leader elbow nudging. "
               "Works for any grid orientation in Revit 2022-2025.")

import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid,
    View,
    ViewSheet,
    ViewType,
    XYZ,
    Transaction,
    DatumExtentType,
    DatumEnds,
)

from pyrevit import forms, script, revit

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()

PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

# Bubble diameter: 1/2" paper space = 0.04167 ft
# Bubble diameter: 2.0 ft directly in Revit model space.
# This is used as-is for collision detection — no view.Scale multiplication.
DEFAULT_BUBBLE_DIAMETER_FT = 2.0

MIN_GRID_LENGTH_FT = 0.01
MAX_ITERATIONS     = 500   # safety cap on nudge loop per view


# =============================================================================
# Pick a grid — calibrate bubble size
# =============================================================================
def pick_reference_grid():
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

        class GridFilter(ISelectionFilter):
            def AllowElement(self, element):
                return isinstance(element, Grid)
            def AllowReference(self, reference, point):
                return False

        forms.alert(
            "Click any grid line to calibrate bubble size.\n"
            "The script will then process all plan views on sheets.",
            title="Separate Grid Bubbles — Pick a Grid",
            ok=True,
        )
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, GridFilter(), "Click any grid line")
        element = doc.GetElement(ref.ElementId)
        if isinstance(element, Grid):
            return element
        forms.alert("Selected element is not a grid. Cancelled.",
                    title="Invalid Selection")
        return None
    except Exception:
        return None


def read_bubble_diameter_ft(grid):
    """Read bubble diameter from grid head annotation family.

    Tries common parameter names on the grid head symbol.
    If unreadable, prompts user to enter the paper-space diameter in inches.
    Returns diameter in Revit internal feet.
    """
    try:
        grid_type = doc.GetElement(grid.GetTypeId())
        if grid_type is not None:
            for param_name in ("End 1 Default Grid Head",
                               "End 2 Default Grid Head",
                               "Default Grid Head"):
                p = grid_type.LookupParameter(param_name)
                if p is not None and p.HasValue:
                    head_sym = doc.GetElement(p.AsElementId())
                    if head_sym is None:
                        continue
                    for radius_name in ("Circle Radius", "Head Radius",
                                        "Radius", "Bubble Radius"):
                        rp = head_sym.LookupParameter(radius_name)
                        if rp is not None and rp.HasValue:
                            diameter_ft = rp.AsDouble() * 2.0
                            if 0.01 < diameter_ft < 10.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} ft ({:.4f} in)**".format(
                                        diameter_ft, diameter_ft * 12.0))
                                return diameter_ft
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    # Could not read from family — ask user
    output.print_md("Could not read bubble diameter from annotation family.")
    try:
        raw = forms.ask_for_string(
            default="2.0",
            prompt=("Enter the grid bubble diameter in MODEL SPACE FEET.\n"
                    "This is how large the bubble appears in the Revit model.\n"
                    "Common value: 2.0 ft (displays as 1/4\" at 1/8\" scale)."),
            title="Grid Bubble Diameter (Model Space Feet)",
        )
        if raw:
            val = float(raw)
            if 0.01 < val < 100.0:
                output.print_md(
                    "User-entered bubble diameter: **{} ft**".format(val))
                return val
    except Exception:
        pass

    output.print_md("Using default: **{} ft**".format(DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold
# =============================================================================
def collision_threshold(view, bubble_diameter_ft):
    """Return the collision threshold in model-space feet.

    bubble_diameter_ft is already in model space (2.0 ft) so no
    view.Scale multiplication is needed. Returned directly.
    """
    return bubble_diameter_ft


# =============================================================================
# View collection
# =============================================================================
def get_sheet_view_ids(document):
    placed_ids = set()
    for sheet in FilteredElementCollector(document).OfClass(ViewSheet).ToElements():
        try:
            for vid in sheet.GetAllPlacedViews():
                placed_ids.add(vid.IntegerValue)
        except Exception:
            pass
    return placed_ids


def collect_plan_views_on_sheets(document):
    sheet_view_ids = get_sheet_view_ids(document)
    result = []
    for v in FilteredElementCollector(document).OfClass(View):
        if v.IsTemplate:
            continue
        if v.ViewType not in PLAN_VIEW_TYPES:
            continue
        if v.Id.IntegerValue not in sheet_view_ids:
            continue
        result.append(v)
    return result


# =============================================================================
# Grid curve helper
# =============================================================================
def get_grid_curve_in_view(grid, view):
    for extent_type in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(extent_type, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None


# =============================================================================
# Bubble helpers
# =============================================================================
def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True


def grid_has_leader_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.GetLeader(end, view) is not None
    except Exception:
        return False


# =============================================================================
# Alphanumeric sort key — used to enforce name-order nudge guard
# =============================================================================
def name_sort_key(name):
    """Sort key: integers by value (4<5<10), letters alphabetically (A<B<C)."""
    import re
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def higher_name(name_a, name_b):
    """Return True if name_a sorts higher than name_b."""
    return name_sort_key(name_a) > name_sort_key(name_b)


# =============================================================================
# Perpendicular direction — universal, works for any grid orientation
# =============================================================================
def get_perp_direction(grid, view):
    """Return the 2D unit vector perpendicular to the grid line.

    Uses Curve.Direction cross product in the view plane:
      tan  = grid direction (normalised)
      perp = XYZ(-tan.Y, tan.X, 0)

    This works for vertical, horizontal, and diagonal grids in any
    Revit version without hardcoding +X/-X/+Y/-Y assumptions.
    """
    try:
        curve = get_grid_curve_in_view(grid, view)
        if curve is None:
            return XYZ(1.0, 0.0, 0.0)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length = (dx * dx + dy * dy) ** 0.5
        if length < 1e-9:
            return XYZ(1.0, 0.0, 0.0)
        # Normalise then rotate 90 degrees in XY plane
        tan_x = dx / length
        tan_y = dy / length
        return XYZ(-tan_y, tan_x, 0.0)
    except Exception:
        return XYZ(1.0, 0.0, 0.0)


# =============================================================================
# Bubble position collection — uses Anchor after leaders exist
# =============================================================================
def collect_bubble_positions(grids, view):
    """Return list of (grid, datum_end, end_index, anchor_pt) for all
    visible bubbles. Uses leader.Anchor if a leader exists (reflects true
    bubble position), otherwise falls back to curve endpoint.
    """
    positions = []
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        for end_index in (0, 1):
            datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            try:
                leader = g.GetLeader(datum_end, view)
                if leader and leader.Anchor:
                    pt = leader.Anchor
                else:
                    pt = curve.GetEndPoint(end_index)
                positions.append((g, datum_end, end_index, pt))
            except Exception:
                continue
    return positions


# =============================================================================
# Collision detection on Anchor positions
# =============================================================================
def find_colliding_anchor_pairs(positions, threshold):
    """Return colliding (pos_a, pos_b) pairs using Anchor XY distance."""
    pairs = []
    threshold_sq = threshold * threshold
    n = len(positions)
    for i in range(n):
        for j in range(i + 1, n):
            g1 = positions[i][0]
            g2 = positions[j][0]
            if g1.Id == g2.Id:
                continue
            p1 = positions[i][3]
            p2 = positions[j][3]
            dx = p1.X - p2.X
            dy = p1.Y - p2.Y
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((positions[i], positions[j]))
    return pairs


# =============================================================================
# Per-view processing
# =============================================================================
def process_view(view, bubble_diam_ft, threshold):
    """Full processing pipeline for one view.

    Returns (leaders_added, errors) tuple.
    """
    leaders_added = 0
    errors        = []

    # --- Collect host grids --------------------------------------------------
    try:
        grids = list(FilteredElementCollector(doc, view.Id)
                     .OfClass(Grid).ToElements())
    except Exception as ex:
        errors.append("Collect grids: {}".format(ex))
        return leaders_added, errors

    if len(grids) < 2:
        return leaders_added, errors

    # --- Step 1: Ensure all grids with visible bubbles have leaders ----------
    # We add leaders to ALL grids that have visible bubbles so Anchor
    # positions are available for accurate collision detection.
    new_leader_keys = set()
    for g in grids:
        for end_index in (0, 1):
            datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            if grid_has_leader_at_end(g, view, end_index):
                continue
            try:
                g.AddLeader(datum_end, view)
                new_leader_keys.add((g.Id.IntegerValue, end_index))
                leaders_added += 1
            except Exception as ex:
                logger.debug("AddLeader grid {} end {}: {}".format(
                    g.Id.IntegerValue, end_index, ex))

    # --- Step 2: doc.Regenerate() REQUIRED before SetLeader ------------------
    # Without this, leader geometry is stale and SetLeader will fail or
    # produce incorrect results. This is the critical missing step.
    doc.Regenerate()

    # --- Step 3: Reset new leader elbows to sit on the grid line -------------
    # Project the default elbow onto the grid curve so we start from a
    # clean baseline position before nudging.
    for g in grids:
        for end_index in (0, 1):
            datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
            key = (g.Id.IntegerValue, end_index)
            if key not in new_leader_keys:
                continue
            try:
                leader = g.GetLeader(datum_end, view)
                if leader is None:
                    continue
                curve = get_grid_curve_in_view(g, view)
                if curve is None:
                    continue
                projected = curve.Project(leader.Elbow)
                if projected is not None:
                    leader.Elbow = projected.XYZPoint
                    g.SetLeader(datum_end, view, leader)
            except Exception as ex:
                logger.debug("Reset elbow grid {} end {}: {}".format(
                    g.Id.IntegerValue, end_index, ex))

    # --- Step 4: Iteratively nudge colliding bubbles apart -------------------
    # Nudge increment = 1/8 bubble diameter per step.
    nudge_step = threshold / 8.0

    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_anchor_pairs(positions, threshold)

        if not pairs:
            break  # all clear

        # Build a dict of unique targets to nudge this iteration.
        # Key: (grid_id, end_index)
        # Value: [grid, datum_end, nudge_x, nudge_y]
        #
        # Direction rule — name order determines movement direction:
        #   Higher name moves in +perp direction (right for vertical,
        #   up for horizontal based on standard grid layout).
        #   Lower name moves in -perp direction (left / down).
        #
        # This guarantees bubbles always diverge from each other —
        # higher numbers/letters go one way, lower go the other.
        # No bubble can cross toward a lower-named neighbour because
        # the direction is fixed by name order, not by dot product.
        #
        # Example: 8, 9, 10 all colliding
        #   10 always moves +perp (right)
        #   9  always moves +perp (right) relative to 8, -perp relative to 10
        #      -> net: determined by its neighbours' names
        #   8  always moves -perp (left)
        targets = {}

        name_map = {g.Id.IntegerValue: g.Name for g in grids}

        for pos_a, pos_b in pairs:
            g_a, end_a, idx_a, anchor_a = pos_a
            g_b, end_b, idx_b, anchor_b = pos_b

            name_a = name_map.get(g_a.Id.IntegerValue, "")
            name_b = name_map.get(g_b.Id.IntegerValue, "")

            perp_a = get_perp_direction(g_a, view)
            perp_b = get_perp_direction(g_b, view)

            # Determine sign for each grid based purely on name order.
            # Higher name gets +1 (moves in +perp direction).
            # Lower name gets -1 (moves in -perp direction).
            # Both grids move — they diverge from each other.
            if higher_name(name_a, name_b):
                sign_a = 1.0   # a is higher — moves +perp
                sign_b = -1.0  # b is lower  — moves -perp
            else:
                sign_a = -1.0  # a is lower  — moves -perp
                sign_b = 1.0   # b is higher — moves +perp

            # Accumulate nudge for g_a
            key_a = (g_a.Id.IntegerValue, idx_a)
            if key_a not in targets:
                targets[key_a] = [g_a, end_a, 0.0, 0.0]
            targets[key_a][2] += perp_a.X * sign_a
            targets[key_a][3] += perp_a.Y * sign_a

            # Accumulate nudge for g_b
            key_b = (g_b.Id.IntegerValue, idx_b)
            if key_b not in targets:
                targets[key_b] = [g_b, end_b, 0.0, 0.0]
            targets[key_b][2] += perp_b.X * sign_b
            targets[key_b][3] += perp_b.Y * sign_b

        # Apply one nudge per unique target this iteration
        for key, (move_grid, move_end, net_x, net_y) in targets.items():
            net_len = (net_x * net_x + net_y * net_y) ** 0.5
            if net_len < 1e-9:
                continue
            # Normalise the net direction
            nx = net_x / net_len
            ny = net_y / net_len

            try:
                leader = move_grid.GetLeader(move_end, view)
                if leader is None:
                    continue
                elbow = leader.Elbow
                new_elbow = XYZ(
                    elbow.X + nx * nudge_step,
                    elbow.Y + ny * nudge_step,
                    elbow.Z,
                )
                leader.Elbow = new_elbow
                move_grid.SetLeader(move_end, view, leader)

            except Exception as ex:
                errors.append("Nudge grid {} iter {}: {}".format(
                    move_grid.Id.IntegerValue, iteration, ex))
                logger.debug(traceback.format_exc())

    return leaders_added, errors


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Check active view ----------------------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. Pick a grid ----------------------------------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)

    # ---- 3. Collect views --------------------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ---- 4. Stats ----------------------------------------------------------
    views_processed = 0
    total_leaders   = 0
    all_errors      = []

    # ---- 5. Single transaction — one Ctrl+Z undoes everything --------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold = collision_threshold(view, bubble_diam_ft)
                added, errors = process_view(view, bubble_diam_ft, threshold)
                total_leaders += added
                for err in errors:
                    all_errors.append((view.Name, err))
                views_processed += 1

            except Exception as ex:
                all_errors.append((view.Name, str(ex)))
                logger.debug(traceback.format_exc())
                continue

        t.Commit()

    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Transaction failed and was rolled back.\n\n{}".format(ex),
            title="Error",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    # ---- 6. Results --------------------------------------------------------
    summary = "\n".join([
        "Views processed: {}".format(views_processed),
        "Leaders added:   {}".format(total_leaders),
        "Errors:          {}".format(len(all_errors)),
    ])

    output.print_md("### Results\n```\n{}\n```".format(summary))

    if all_errors:
        output.print_md("### Errors")
        for vname, err in all_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee pyRevit output for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


if __name__ == "__main__":
    main()