# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets using
Revit's leader (elbow/break) feature.

How it works:
  When two grid bubbles overlap, Revit's manual fix is to use the "break"
  handle which creates a leader on the bubble end — a bent elbow line that
  offsets just the bubble circle away from the grid line endpoint, leaving
  the grid line itself completely untouched.

  In the API this is done with:
    grid.AddLeader(DatumEnds.End1, view)        — creates the leader
    leader = grid.GetLeader(DatumEnds.End1, view)
    leader.Elbow = XYZ(elbow point)             — sets the bend point
    leader.End   = XYZ(bubble offset point)     — sets where bubble lands
    grid.SetLeader(DatumEnds.End1, view, leader) — writes it back

  The bubble is moved perpendicular to the grid line (away from its
  colliding neighbour) while the grid line endpoint stays fixed.
  This is exactly what Revit's manual break/elbow handle does.

Collision detection:
  Pure 2D (X, Y only). Z stripped at collection, never used again.
  Threshold = one bubble diameter in model space (view-scale-aware).

Scope:
  FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.

Calibration:
  User picks one grid. Bubble diameter is read from that grid's
  annotation family. Falls back to 3/8" default if unreadable.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "7.0.0"
__doc__     = ("Pick a grid, then automatically adds elbow leaders to "
               "separate colliding grid bubbles on all plan views on sheets. "
               "Grid lines are never moved — only the bubble annotation.")

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
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
    RevitLinkInstance,
)

from pyrevit import forms, script, revit

# -----------------------------------------------------------------------------
# Handles
# -----------------------------------------------------------------------------
doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

DEFAULT_BUBBLE_DIAMETER_INCHES = 0.375  # 3/8" standard Revit grid head
OFFSET_MULTIPLIER = 1.25               # 1x clears overlap + 0.25x gap
MIN_GRID_LENGTH_FT = 0.01              # skip degenerate grids


# =============================================================================
# Pick a grid — calibrate bubble size from annotation family
# =============================================================================
def pick_reference_grid():
    """Prompt user to click a grid in the active view.

    Returns the Grid element, or None if cancelled.
    """
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


def read_bubble_diameter_inches(grid):
    """Read grid head annotation bubble diameter from the picked grid's type.

    Checks the grid type's head family symbol for a radius parameter.
    Falls back to DEFAULT_BUBBLE_DIAMETER_INCHES if not found.
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
                            diameter_in = rp.AsDouble() * 2.0 * 12.0
                            if 0.1 < diameter_in < 2.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} in**".format(diameter_in))
                                return diameter_in
    except Exception as ex:
        logger.debug("read_bubble_diameter_inches: {}".format(ex))

    output.print_md("Using default bubble diameter: "
                    "**{} in**".format(DEFAULT_BUBBLE_DIAMETER_INCHES))
    return DEFAULT_BUBBLE_DIAMETER_INCHES


# =============================================================================
# Scale helpers
# =============================================================================
def bubble_diameter_model_units(view, bubble_inches):
    """Bubble diameter in decimal feet at this view's print scale."""
    try:
        scale = float(view.Scale)
        if scale <= 0:
            scale = 96.0
    except Exception:
        scale = 96.0
    return (bubble_inches / 12.0) * scale


def offset_distance_model_units(view, bubble_inches):
    return OFFSET_MULTIPLIER * bubble_diameter_model_units(view, bubble_inches)


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
# Grid collection — host + linked (linked for detection only)
# =============================================================================
def collect_all_grids_in_view(document, view):
    results = []
    try:
        for g in (FilteredElementCollector(document, view.Id)
                  .OfClass(Grid).ToElements()):
            results.append({
                'grid':      g,
                'grid_id':   "host:{}".format(g.Id.IntegerValue),
                'is_linked': False,
            })
    except Exception as ex:
        logger.debug("Host grids: {}".format(ex))

    try:
        for link in (FilteredElementCollector(document)
                     .OfClass(RevitLinkInstance).ToElements()):
            try:
                link_doc = link.GetLinkDocument()
                if link_doc is None:
                    continue
                for g in (FilteredElementCollector(link_doc)
                          .OfClass(Grid).ToElements()):
                    results.append({
                        'grid':      g,
                        'grid_id':   "link_{}:{}".format(
                            link_doc.Title, g.Id.IntegerValue),
                        'is_linked': True,
                    })
            except Exception:
                pass
    except Exception:
        pass

    return results


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
# Bubble visibility
# =============================================================================
def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True


def grid_already_has_leader(grid, view, end_index):
    """Return True if a leader already exists on this end in this view.

    Prevents adding a second leader on repeated runs (idempotent).
    """
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        leader = grid.GetLeader(end, view)
        return leader is not None
    except Exception:
        return False


# =============================================================================
# Entry collection — pure 2D, one entry per visible bubble
# Deduplication: one entry per (grid_id, end_index) — no duplicates
# =============================================================================
def collect_bubble_entries(document, view):
    """One dict per visible bubble endpoint. Coordinates are 2D only."""
    entries = []
    seen_keys = set()
    grid_infos = collect_all_grids_in_view(document, view)

    for info in grid_infos:
        g = info['grid']
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (info['grid_id'], end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':        g,
                'grid_id':     info['grid_id'],
                'is_linked':   info['is_linked'],
                'end_index':   end_index,
                'x':           pt.X,   # 2D only — Z stripped here
                'y':           pt.Y,
                'curve':       curve,  # kept for direction vector calculation
            })
    return entries


# =============================================================================
# Collision detection — pure 2D, deduplicated grid pairs
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Return one (entry_a, entry_b) per colliding grid pair.

    Pure 2D — only X and Y used. Z never referenced.
    Deduplication by frozenset of grid_ids prevents counting the same
    grid pair multiple times when both their ends are within threshold.
    """
    pairs = []
    seen_pairs = set()
    n = len(entries)
    threshold_sq = threshold * threshold

    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue

            # Deduplicate: each grid pair counted once regardless of
            # which endpoint combination triggered the detection
            pair_key = frozenset([entries[i]['grid_id'],
                                  entries[j]['grid_id']])
            if pair_key in seen_pairs:
                continue

            dx = entries[i]['x'] - entries[j]['x']
            dy = entries[i]['y'] - entries[j]['y']
            if (dx * dx + dy * dy) <= threshold_sq:
                seen_pairs.add(pair_key)
                pairs.append((entries[i], entries[j]))

    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Pick the host grid to add a leader to.

    If one is linked (read-only), always pick the host.
    If both linked, return None (skip).
    If both host, pick the one with the higher grid_id (deterministic).
    """
    if entry_a['is_linked'] and entry_b['is_linked']:
        return None
    if entry_a['is_linked']:
        return entry_b
    if entry_b['is_linked']:
        return entry_a
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


def get_other_entry(entry_a, entry_b, target):
    """Return the entry that is NOT the target (the neighbour to move away from)."""
    return entry_b if target is entry_a else entry_a


# =============================================================================
# Leader (elbow/break) application
# =============================================================================
def perpendicular_direction_2d(curve, end_index):
    """Return a 2D unit vector perpendicular to the grid, pointing outward.

    Perpendicular to the grid direction means the bubble slides sideways
    away from its colliding neighbour — exactly what Revit's manual break
    handle does. The direction is 90 degrees to the grid line in XY.
    """
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    length = (dx * dx + dy * dy) ** 0.5
    if length < MIN_GRID_LENGTH_FT:
        return XYZ(1.0, 0.0, 0.0)

    # Normalise grid direction then rotate 90 degrees in XY
    # Rotate CCW: (dx, dy) -> (-dy, dx)
    perp_x = -dy / length
    perp_y =  dx / length
    return XYZ(perp_x, perp_y, 0.0)


def apply_leader(target_entry, neighbour_entry, view,
                 bubble_diam_model, already_done_keys):
    """Add an elbow leader to the target grid's bubble end so it moves
    perpendicular to the grid, away from the colliding neighbour.

    Leader geometry:
      - The grid line endpoint stays fixed (never moved).
      - Elbow point = grid endpoint + (perp_vector * 0.5 * offset)
        This is the bend/kink point of the leader line.
      - End (bubble centre) = grid endpoint + (perp_vector * offset)
        This is where the bubble circle actually lands.

    The perpendicular direction is chosen to move away from the neighbour
    by checking which side of the grid the neighbour's bubble is on.

    Returns True if leader was added, False if skipped.
    """
    key = (target_entry['grid_id'], target_entry['end_index'])
    if key in already_done_keys:
        return False

    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
    curve     = target_entry['curve']

    # Skip if leader already exists on this end (idempotent on repeat runs)
    if grid_already_has_leader(grid, view, end_index):
        already_done_keys.add(key)
        return False

    # Get the bubble endpoint position (2D, Z from original curve)
    bubble_pt = curve.GetEndPoint(end_index)
    z = bubble_pt.Z  # preserve level elevation

    # Compute perpendicular unit vector
    perp = perpendicular_direction_2d(curve, end_index)

    # Determine which side of the grid the neighbour is on so we move AWAY
    # Compare the perpendicular component of (neighbour_pos - bubble_pt)
    neighbour_dx = neighbour_entry['x'] - target_entry['x']
    neighbour_dy = neighbour_entry['y'] - target_entry['y']
    dot = neighbour_dx * perp.X + neighbour_dy * perp.Y

    # If neighbour is on the perp side (dot > 0), flip direction to move away
    if dot > 0:
        perp = XYZ(-perp.X, -perp.Y, 0.0)

    offset = bubble_diam_model * OFFSET_MULTIPLIER

    # Elbow = halfway along the leader (the bend point)
    elbow_pt = XYZ(
        bubble_pt.X + perp.X * offset * 0.5,
        bubble_pt.Y + perp.Y * offset * 0.5,
        z,
    )

    # End = where the bubble circle lands
    end_pt = XYZ(
        bubble_pt.X + perp.X * offset,
        bubble_pt.Y + perp.Y * offset,
        z,
    )

    # Add the leader to this end in this view
    grid.AddLeader(datum_end, view)

    # Read back the newly created leader and set elbow + end positions
    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("AddLeader succeeded but GetLeader returned None")

    leader.Elbow = elbow_pt
    leader.End   = end_pt
    grid.SetLeader(datum_end, view, leader)

    already_done_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Verify active view ---------------------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. Pick a grid to calibrate ---------------------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_inches = read_bubble_diameter_inches(ref_grid)

    # ---- 3. Collect views --------------------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ---- 4. Stats ----------------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    leaders_added    = 0
    skipped_linked   = 0
    per_view_errors  = []

    # ---- 5. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                bubble_diam_model = bubble_diameter_model_units(
                    view, bubble_inches)
                threshold = bubble_diam_model  # 1x diameter = collision zone

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                done_keys = set()
                for a, b in pairs:
                    target   = choose_entry_to_move(a, b)
                    if target is None:
                        skipped_linked += 1
                        continue
                    neighbour = get_other_entry(a, b, target)
                    try:
                        if apply_leader(target, neighbour, view,
                                        bubble_diam_model, done_keys):
                            leaders_added += 1
                    except Exception as ex:
                        per_view_errors.append((
                            view.Name,
                            "Grid {}: {}".format(target['grid_id'], ex),
                        ))
                        logger.debug(traceback.format_exc())

                views_processed += 1

            except Exception as ex:
                per_view_errors.append((view.Name, str(ex)))
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
        "Views processed:  {}".format(views_processed),
        "Collisions found: {}".format(collisions_found),
        "Leaders added:    {}".format(leaders_added),
        "Skipped (linked): {}".format(skipped_linked),
        "Errors:           {}".format(len(per_view_errors)),
    ])

    output.print_md("### Results\n```\n{}\n```".format(summary))

    if per_view_errors:
        output.print_md("### Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee pyRevit output for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


if __name__ == "__main__":
    main()