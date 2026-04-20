# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Current state (handoff to Opus for continued debugging):

WHAT WORKS:
  - Collision detection using leader.Anchor positions (2D XY only)
  - AddLeader correctly adds elbow break to colliding grids
  - doc.Regenerate() after AddLeader required before SetLeader works
  - Name-order logic: higher name (5>4, E>D) moves, lower stays
  - Process order: highest-named first so it clears space before lower moves
  - Nudge direction math verified:
      Vertical grid   tan=(0,1): perp=(tan_y,-tan_x)=(1,0)  = RIGHT (+X) ✓
      Horizontal grid tan=(1,0): perp=(tan_y,-tan_x)=(0,-1) = DOWN  (-Y) ✓
  - Bubble size: 2.0 ft model space (1/4" at 1/8" scale), user-enterable
  - Views: FloorPlan/CeilingPlan/AreaPlan/EngineeringPlan on sheets only
  - Host grids only (linked grids excluded)

OUTSTANDING ISSUES:
  - SetLeader throws "Elbow is between End and Anchor" on some grids
    even after Regenerate. The elbow reset step (Project onto curve) was
    removed because it caused this error. Need to find the valid Elbow
    range and constrain nudges to stay within it.
  - 1500 errors reported on last run — SetLeader failing consistently.
    Root cause: after AddLeader+Regenerate, the Anchor and End define a
    valid range for Elbow. Moving Elbow outside [Anchor..End] throws.
    Need to: (a) read Anchor and End after Regenerate, (b) clamp nudge
    direction so Elbow stays between Anchor and End at all times.
  - Script needs to read leader.Anchor and leader.End AFTER Regenerate
    and verify Elbow stays on the line segment between them.

KEY REVIT API FACTS (proven by diagnostic runs):
  - leader.Anchor is READ-ONLY (computed from Elbow and End)
  - leader.End must stay ON the grid's infinite axis line
  - leader.Elbow must be geometrically BETWEEN Anchor and End
  - SetLeader signature: grid.SetCurveInView(DatumExtentType, View, Curve)
    NO — correct: grid.SetLeader(DatumEnds, View, Leader)
  - AddLeader signature: grid.AddLeader(DatumEnds, View)
  - After AddLeader, MUST call doc.Regenerate() before GetLeader/SetLeader

BUBBLE SIZE:
  - 2.0 ft in model space at 1/8" scale
  - Collision threshold = 2.0 ft (no view.Scale multiplication needed)
  - Nudge step = threshold / 8.0 = 0.25 ft per iteration
  - MAX_ITERATIONS = 50
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "15.0.0"
__doc__     = ("Separates colliding grid bubbles using leader elbow nudging. "
               "Works for any grid orientation in Revit 2022-2025.")

import re
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

# Bubble diameter: 2.0 ft directly in Revit model space (no view.Scale multiply)
DEFAULT_BUBBLE_DIAMETER_FT = 2.0

MIN_GRID_LENGTH_FT = 0.01
MAX_ITERATIONS     = 50   # nudge step = threshold/8 = 0.25 ft, 50 steps = 12.5 ft max


# =============================================================================
# Name sort — higher number/letter = further in sequence = higher sort value
# =============================================================================
def name_sort_key(name):
    """Sort key: 4<5<6<10, A<B<C<D<E. Further in sequence = higher value."""
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def higher_name(name_a, name_b):
    """Return True if name_a is further in the counting/alphabet sequence."""
    return name_sort_key(name_a) > name_sort_key(name_b)


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
    Falls back to user input, then default 2.0 ft."""
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
                                    "**{:.4f} ft**".format(diameter_ft))
                                return diameter_ft
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    output.print_md("Could not read bubble diameter from annotation family.")
    try:
        raw = forms.ask_for_string(
            default="2.0",
            prompt=("Enter the grid bubble diameter in MODEL SPACE FEET.\n"
                    "Common value: 2.0 ft (displays as 1/4\" at 1/8\" scale)."),
            title="Grid Bubble Diameter (Model Space Feet)",
        )
        if raw:
            val = float(raw)
            if 0.01 < val < 100.0:
                output.print_md(
                    "User-entered: **{} ft**".format(val))
                return val
    except Exception:
        pass

    output.print_md("Using default: **{} ft**".format(DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold — model space, no scaling
# =============================================================================
def collision_threshold(view, bubble_diameter_ft):
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
# Nudge direction — verified math:
#   Vertical   tan=(0,1): (tan_y,-tan_x)=(1, 0) = RIGHT (+X) ✓
#   Horizontal tan=(1,0): (tan_y,-tan_x)=(0,-1) = DOWN  (-Y) ✓
# =============================================================================
def get_nudge_direction(grid, view):
    """Unit vector that higher-named grids move toward (RIGHT or DOWN)."""
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
        tan_x = dx / length
        tan_y = dy / length
        # Verified: vertical→RIGHT, horizontal→DOWN
        return XYZ(tan_y, -tan_x, 0.0)
    except Exception:
        return XYZ(1.0, 0.0, 0.0)


# =============================================================================
# Bubble position collection — uses Anchor after leaders exist
# =============================================================================
def collect_bubble_positions(grids, view):
    """(grid, datum_end, end_index, anchor_pt) for all visible bubbles.
    Uses leader.Anchor if available, else curve endpoint."""
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
# Collision detection
# =============================================================================
def find_colliding_anchor_pairs(positions, threshold):
    """Colliding (pos_a, pos_b) pairs using Anchor XY distance. Pure 2D."""
    pairs = []
    threshold_sq = threshold * threshold
    n = len(positions)
    for i in range(n):
        for j in range(i + 1, n):
            if positions[i][0].Id == positions[j][0].Id:
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
    leaders_added = 0
    errors        = []

    try:
        grids = list(FilteredElementCollector(doc, view.Id)
                     .OfClass(Grid).ToElements())
    except Exception as ex:
        errors.append("Collect grids: {}".format(ex))
        return leaders_added, errors

    if len(grids) < 2:
        return leaders_added, errors

    # --- Step 1: AddLeader on all grids with visible bubbles -----------------
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

    # --- Step 2: REQUIRED — Regenerate before any SetLeader calls ------------
    # Without this, leader geometry is stale and SetLeader fails.
    doc.Regenerate()

    # --- Step 3: Iteratively nudge colliding bubbles apart -------------------
    # nudge_step = threshold / 8 = 0.25 ft per step
    # Only higher-named grid moves per pair. Lower stays.
    # Process highest-named first so it clears space before lower moves.
    #
    # KNOWN ISSUE: SetLeader throws "Elbow is between End and Anchor"
    # when Elbow is nudged outside the valid range defined by Anchor and End.
    # TODO: After Regenerate, read leader.Anchor and leader.End, then
    # clamp each nudge so Elbow stays on the segment between them.
    nudge_step = threshold / 8.0

    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_anchor_pairs(positions, threshold)

        if not pairs:
            break

        targets  = {}
        name_map = {g.Id.IntegerValue: g.Name for g in grids}

        for pos_a, pos_b in pairs:
            g_a, end_a, idx_a, anchor_a = pos_a
            g_b, end_b, idx_b, anchor_b = pos_b

            name_a = name_map.get(g_a.Id.IntegerValue, "")
            name_b = name_map.get(g_b.Id.IntegerValue, "")

            # Only higher-named moves
            if higher_name(name_a, name_b):
                move_g, move_end, move_idx, move_name = g_a, end_a, idx_a, name_a
            else:
                move_g, move_end, move_idx, move_name = g_b, end_b, idx_b, name_b

            nudge_dir = get_nudge_direction(move_g, view)
            key = (move_g.Id.IntegerValue, move_idx)
            if key not in targets:
                targets[key] = [move_g, move_end, 0.0, 0.0, move_name]
            targets[key][2] += nudge_dir.X
            targets[key][3] += nudge_dir.Y

        # Highest-named first
        sorted_targets = sorted(
            targets.items(),
            key=lambda item: name_sort_key(item[1][4]),
            reverse=True
        )

        for key, target_data in sorted_targets:
            move_grid = target_data[0]
            move_end  = target_data[1]
            net_x     = target_data[2]
            net_y     = target_data[3]

            net_len = (net_x * net_x + net_y * net_y) ** 0.5
            if net_len < 1e-9:
                continue
            nx = net_x / net_len
            ny = net_y / net_len

            try:
                leader = move_grid.GetLeader(move_end, view)
                if leader is None:
                    continue

                # Read the current leader geometry. Anchor is read-only and
                # computed from Elbow + End; we use it only for clamping.
                anchor = leader.Anchor
                elbow  = leader.Elbow
                end    = leader.End

                # Proposed new Elbow before clamping.
                prop_x = elbow.X + nx * nudge_step
                prop_y = elbow.Y + ny * nudge_step

                # Clamp the proposed Elbow into the axis-aligned bounding
                # box of Anchor and End so the
                #     "Elbow is between End and Anchor"
                # constraint always holds. A 1/16" inset keeps us strictly
                # off the boundary (Revit's "between" check can reject
                # exact-equal values).
                margin = 1.0 / 12.0 / 16.0   # ~1/16" in feet

                min_x = min(anchor.X, end.X) + margin
                max_x = max(anchor.X, end.X) - margin
                min_y = min(anchor.Y, end.Y) + margin
                max_y = max(anchor.Y, end.Y) - margin

                # If Anchor and End coincide on an axis, margin can collapse
                # the range; pin to the midpoint rather than flip order.
                if max_x < min_x:
                    min_x = max_x = (anchor.X + end.X) / 2.0
                if max_y < min_y:
                    min_y = max_y = (anchor.Y + end.Y) / 2.0

                if   prop_x < min_x: prop_x = min_x
                elif prop_x > max_x: prop_x = max_x
                if   prop_y < min_y: prop_y = min_y
                elif prop_y > max_y: prop_y = max_y

                # If clamping left Elbow unchanged, this grid has no room
                # to move further in the requested direction — skip the
                # SetLeader call rather than burning an iteration on a
                # guaranteed no-op.
                if (abs(prop_x - elbow.X) < 1e-9 and
                        abs(prop_y - elbow.Y) < 1e-9):
                    continue

                leader.Elbow = XYZ(prop_x, prop_y, elbow.Z)
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
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)

    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    views_processed = 0
    total_leaders   = 0
    all_errors      = []

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