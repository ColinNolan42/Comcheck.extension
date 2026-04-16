# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Offset direction — based purely on bubble position in the view, NO name parsing:

  The script looks at WHERE the bubble physically sits relative to the
  view's crop box center:

  Bubble is at the BOTTOM of the view  (Y < view center Y):
    -> shift RIGHT (+X). Clears by moving further right along the page.

  Bubble is on the SIDE of the view (LEFT or RIGHT)
  (Y is near center, X is near left or right edge):
    -> shift UP (+Y). Clears by moving further up along the page.

  This works regardless of whether architects put numbers or letters
  on the bottom, side, or any combination. No grid names are read.
  The decision is made purely from the bubble's XY position vs the
  view crop box center.

  "Higher" grid in a collision = whichever bubble is further in the
  offset direction already. That one moves further in that direction.
  If tied, the one with the higher ElementId moves (deterministic).

Default offset: 4'-0" in model space (Revit internal decimal feet).

Leader geometry (proven by diagnostic):
  1. AddLeader if no leader exists, else reuse existing
  2. Read default Anchor/Elbow/End
  3. Extend End by full offset in the chosen direction
  4. Extend Elbow by half offset in the same direction
  5. SetLeader

Host grids only. Linked grids excluded.
FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "10.0.0"
__doc__     = ("Separates colliding grid bubbles. Bottom bubbles shift right, "
               "side bubbles shift up. Based on bubble position, not grid name.")

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

# Offset in Revit internal units (decimal feet) — 4'-0"
DEFAULT_OFFSET_FT = 4.0

# Fallback bubble diameter for collision threshold (inches)
DEFAULT_BUBBLE_DIAMETER_INCHES = 0.375

MIN_GRID_LENGTH_FT = 0.01


# =============================================================================
# View crop box center — used to determine bubble position on page
# =============================================================================
def get_view_center(view):
    """Return the XY center of the view's crop box in model coordinates.

    This is the reference point for deciding if a bubble is at the
    bottom (Y < center.Y) or on the side (Y near center.Y).

    Falls back to the view's origin if crop box is unavailable.
    """
    try:
        if view.CropBoxActive and view.CropBox is not None:
            bb = view.CropBox
            cx = (bb.Min.X + bb.Max.X) * 0.5
            cy = (bb.Min.Y + bb.Max.Y) * 0.5
            cz = (bb.Min.Z + bb.Max.Z) * 0.5
            return XYZ(cx, cy, cz)
    except Exception:
        pass
    try:
        return view.Origin
    except Exception:
        return XYZ(0, 0, 0)


def get_view_bounds(view):
    """Return (min_x, max_x, min_y, max_y) of the view crop box.

    Used to determine what fraction of the view height a bubble sits at.
    """
    try:
        if view.CropBoxActive and view.CropBox is not None:
            bb = view.CropBox
            return bb.Min.X, bb.Max.X, bb.Min.Y, bb.Max.Y
    except Exception:
        pass
    return None, None, None, None


# =============================================================================
# Bubble position classifier — bottom vs side
# =============================================================================
def get_bubble_offset_direction(entry, view):
    """Return the XYZ direction the bubble should move to separate.

    Logic — purely positional, no name parsing:

    1. Get the view crop box bounds.
    2. Compute what fraction of the view height the bubble sits at:
         frac_y = (bubble.Y - min_y) / (max_y - min_y)
    3. If frac_y < 0.35 -> bubble is in the bottom 35% -> shift RIGHT (+X)
       If frac_y > 0.65 -> bubble is in the top 35% -> shift RIGHT (+X)
         (top bubbles also shift right — same axis as bottom)
       Otherwise (frac_y between 0.35 and 0.65) -> bubble is on a side
         -> shift UP (+Y)

    Why top and bottom both shift right:
      Vertical gridlines have bubbles at top AND bottom. Both ends of
      a vertical grid should shift the same direction (right) so the
      grid doesn't twist.

    Why side bubbles shift up:
      Horizontal gridlines have bubbles on the left or right side.
      Shifting up separates them cleanly without crossing other grids.
    """
    min_x, max_x, min_y, max_y = get_view_bounds(view)

    # Fallback if crop box unavailable — use grid orientation from curve
    if min_y is None or max_y is None or (max_y - min_y) < 0.01:
        # Fall back to grid direction: vertical grid -> right, horizontal -> up
        if entry['is_vertical']:
            return XYZ(1.0, 0.0, 0.0)
        else:
            return XYZ(0.0, 1.0, 0.0)

    bub_y = entry['y']
    height = max_y - min_y
    frac_y = (bub_y - min_y) / height

    if frac_y < 0.35 or frac_y > 0.65:
        # Bubble is near top or bottom of view — this is a vertical grid
        # Shift RIGHT (+X)
        return XYZ(1.0, 0.0, 0.0)
    else:
        # Bubble is on the left or right side — this is a horizontal grid
        # Shift UP (+Y)
        return XYZ(0.0, 1.0, 0.0)


# =============================================================================
# Pick a grid — calibrate bubble size for collision threshold
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
            "Click any grid line to calibrate bubble size for collision "
            "detection.\nThe script will then process all plan views on sheets.",
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
# Scale helpers — collision threshold only
# =============================================================================
def bubble_diameter_model_units(view, bubble_inches):
    """Bubble diameter in model feet at view scale — for collision threshold."""
    try:
        scale = float(view.Scale)
        if scale <= 0:
            scale = 96.0
    except Exception:
        scale = 96.0
    return (bubble_inches / 12.0) * scale


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
# Bubble and leader state
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
# Entry collection — HOST GRIDS ONLY, pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
    """One deduplicated entry per visible bubble on host grids only."""
    entries = []
    seen_keys = set()

    try:
        host_grids = (FilteredElementCollector(document, view.Id)
                      .OfClass(Grid).ToElements())
    except Exception:
        return entries

    for g in host_grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length_2d = (dx * dx + dy * dy) ** 0.5
        if length_2d < MIN_GRID_LENGTH_FT:
            continue

        # Classify orientation from actual curve geometry
        is_vertical = abs(dy) >= abs(dx)

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (g.Id.IntegerValue, end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':        g,
                'grid_id':     g.Id.IntegerValue,
                'end_index':   end_index,
                'x':           pt.X,
                'y':           pt.Y,
                'z':           pt.Z,
                'is_vertical': is_vertical,
            })

    return entries


# =============================================================================
# Collision detection — pure 2D, deduplicated by grid pair
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """(entry_a, entry_b) pairs within threshold. Pure 2D. One per grid pair."""
    pairs = []
    seen_pairs = set()
    n = len(entries)
    threshold_sq = threshold * threshold

    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue

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


def choose_entry_to_move(entry_a, entry_b, view):
    """Choose which bubble to move based on its position in the view.

    For vertical grids (bottom/top bubbles, shift right):
      Move the bubble that is further LEFT (lower X) — it needs to go
      right to clear its neighbour which is already to the right.
      If tied, higher grid_id moves.

    For horizontal grids (side bubbles, shift up):
      Move the bubble that is further DOWN (lower Y) — it needs to go
      up to clear its neighbour which is already higher.
      If tied, higher grid_id moves.

    This ensures consistent movement: always toward the standard direction,
    never away from it.
    """
    # Use entry_a's classification (both should be same orientation)
    is_vertical = entry_a['is_vertical']

    if is_vertical:
        # Move the one further left (it shifts right to clear)
        if abs(entry_a['x'] - entry_b['x']) > 0.001:
            return entry_a if entry_a['x'] < entry_b['x'] else entry_b
    else:
        # Move the one further down (it shifts up to clear)
        if abs(entry_a['y'] - entry_b['y']) > 0.001:
            return entry_a if entry_a['y'] < entry_b['y'] else entry_b

    # Tied position — higher grid_id moves (deterministic)
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Leader application
# =============================================================================
def apply_leader(target_entry, view, already_done_keys):
    """Add or reposition a leader to move the bubble in the correct direction.

    Offset direction is determined by bubble position in view:
      Bottom/top of view -> RIGHT (+X)
      Left/right of view -> UP (+Y)

    Leader extension:
      End   moves full DEFAULT_OFFSET_FT in the offset direction
      Elbow moves half DEFAULT_OFFSET_FT (stays between Anchor and End)

    Returns True if applied, False if skipped.
    """
    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
    z         = target_entry['z']

    key = (target_entry['grid_id'], end_index)
    if key in already_done_keys:
        return False

    # Determine offset direction from bubble's position in this view
    offset_dir = get_bubble_offset_direction(target_entry, view)

    has_leader = grid_has_leader_at_end(grid, view, end_index)
    if not has_leader:
        grid.AddLeader(datum_end, view)

    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("GetLeader returned None")

    current_end   = leader.End
    current_elbow = leader.Elbow

    # Extend End and Elbow in the offset direction
    # Both must move together — Revit requires Elbow between Anchor and End
    new_end = XYZ(
        current_end.X   + offset_dir.X * DEFAULT_OFFSET_FT,
        current_end.Y   + offset_dir.Y * DEFAULT_OFFSET_FT,
        z,
    )
    new_elbow = XYZ(
        current_elbow.X + offset_dir.X * (DEFAULT_OFFSET_FT * 0.5),
        current_elbow.Y + offset_dir.Y * (DEFAULT_OFFSET_FT * 0.5),
        z,
    )

    leader.End   = new_end
    leader.Elbow = new_elbow
    grid.SetLeader(datum_end, view, leader)

    already_done_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Check active view type -----------------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. Pick a grid to calibrate bubble size ---------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))
    output.print_md("Offset distance: **{} ft**".format(DEFAULT_OFFSET_FT))

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
    per_view_errors  = []

    # ---- 5. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold = bubble_diameter_model_units(view, bubble_inches)

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                done_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b, view)
                    try:
                        if apply_leader(target, view, done_keys):
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