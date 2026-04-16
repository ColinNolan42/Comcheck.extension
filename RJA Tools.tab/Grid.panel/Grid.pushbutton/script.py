# -*- coding: utf-8 -*-
"""Detects colliding grid bubbles on top-down plan views that are placed on
sheets and automatically offsets one bubble of each colliding pair
perpendicular to the gridline so the annotations no longer overlap.

Scope:
  - Only FloorPlan, CeilingPlan, AreaPlan, and EngineeringPlan view types.
  - Only views that are actually placed on at least one sheet.
  - Sections, elevations, details, 3D views, and unplaced views are skipped.

Collision detection:
  - Threshold is computed automatically per view from view.Scale and the
    standard 3/8-inch Revit grid bubble diameter — no user input required.
  - Offset distance is also computed from view scale (1.5x bubble diameter)
    so bubbles land clearly clear of each other at any sheet scale.

Curve strategy:
  - GetCurvesInView(ViewSpecific) is tried first.
  - If it returns nothing (grid has never been manually adjusted in that view,
    which is the default state for most grids), falls back to
    GetCurvesInView(Model) to read the position.
  - Before writing back, the grid is promoted to ViewSpecific extents using
    SetDatumExtentType() on both ends so SetCurveInView() (singular — the
    correct IronPython API method name) does not throw.

Key API corrections vs previous versions:
  - SetCurvesInView (plural, wrong) -> SetCurveInView (singular, correct)
  - GetCurvesInView (plural) is correct for reading — returns IList
  - SetDatumExtentType must be called before SetCurveInView when the grid
    has not yet been given a ViewSpecific override in that view.

Other behaviour:
  - Only view-specific 2D extents are written — model geometry untouched.
  - For each colliding pair the grid with the higher ElementId is moved so
    the operation is deterministic across repeated runs.
  - Single Transaction wraps everything — one Ctrl+Z undoes the whole run.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "3.0.0"
__doc__     = ("Automatically separates colliding grid bubbles on all plan "
               "views placed on sheets. No user input required.")

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
    Line,
    Transaction,
    DatumExtentType,
    DatumEnds,
)

from pyrevit import forms, script, revit

# -----------------------------------------------------------------------------
# Document handles
# -----------------------------------------------------------------------------
doc    = revit.doc
logger = script.get_logger()
output = script.get_output()

# -----------------------------------------------------------------------------
# View-type filter — top-down plan views only
# -----------------------------------------------------------------------------
PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

# -----------------------------------------------------------------------------
# Scale constants
# Revit grid bubble = 3/8" diameter at plot scale.
# Offset = 1.5x bubble diameter so separated bubbles have clear daylight.
# -----------------------------------------------------------------------------
BUBBLE_DIAMETER_INCHES = 0.375
OFFSET_MULTIPLIER      = 1.5


# =============================================================================
# Scale helpers
# =============================================================================
def bubble_diameter_model_units(view):
    """Bubble diameter in Revit internal feet, scaled to this view's print scale.

    model diameter = (0.375 inches / 12) x view.Scale
    e.g. Scale=96 (1/8"=1'-0") -> 3.0 ft in model space
         Scale=48 (1/4"=1'-0") -> 1.5 ft in model space
    """
    try:
        scale = float(view.Scale)
    except Exception:
        scale = 96.0
    return (BUBBLE_DIAMETER_INCHES / 12.0) * scale


def offset_distance_model_units(view):
    """Offset distance in model feet = OFFSET_MULTIPLIER x bubble diameter."""
    return OFFSET_MULTIPLIER * bubble_diameter_model_units(view)


# =============================================================================
# View collection — plans on sheets only
# =============================================================================
def get_sheet_view_ids(document):
    """Return set of ElementId integer values for every view placed on a sheet."""
    placed_ids = set()
    sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    for sheet in sheets:
        try:
            for vid in sheet.GetAllPlacedViews():
                placed_ids.add(vid.IntegerValue)
        except Exception as ex:
            logger.debug("Sheet {}: {}".format(sheet.Name, ex))
    return placed_ids


def collect_plan_views_on_sheets(document):
    """Return all non-template plan views placed on at least one sheet."""
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
# Grid curve retrieval — ViewSpecific with Model fallback
# =============================================================================
def get_grid_curve_in_view(grid, view):
    """Return the first available curve for a grid in a view, or None.

    Tries ViewSpecific first (2D override). Falls back to Model if the grid
    has never been manually adjusted in this view (the default state).
    """
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
    """Return True if the bubble at end_index (0 or 1) is visible in view."""
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True  # assume present if check fails


# =============================================================================
# Entry collection
# =============================================================================
def collect_bubble_entries(document, view):
    """Return one dict per visible bubble endpoint in this view.

    Keys: grid, grid_id, curve, end_index, point
    """
    entries = []
    grids = FilteredElementCollector(document, view.Id).OfClass(Grid).ToElements()
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            entries.append({
                'grid':      g,
                'grid_id':   g.Id.IntegerValue,
                'curve':     curve,
                'end_index': end_index,
                'point':     curve.GetEndPoint(end_index),
            })
    return entries


# =============================================================================
# Collision detection
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Return (entry_a, entry_b) pairs within threshold feet (2D, Z ignored)."""
    pairs = []
    n = len(entries)
    threshold_sq = threshold * threshold
    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue
            p1 = entries[i]['point']
            p2 = entries[j]['point']
            dx = p1.X - p2.X
            dy = p1.Y - p2.Y
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((entries[i], entries[j]))
    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Move the entry with the higher ElementId — deterministic across runs."""
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Offset application
# =============================================================================
def perpendicular_offset_vector(curve, view, distance):
    """Vector perpendicular to the grid direction in the view plane."""
    try:
        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    except Exception:
        direction = view.RightDirection
    try:
        perp = view.ViewDirection.CrossProduct(direction).Normalize()
    except Exception:
        perp = view.RightDirection
    return perp.Multiply(distance)


def promote_to_view_specific(grid, view):
    """Ensure the grid has ViewSpecific extents in this view.

    SetCurveInView() (singular — the correct IronPython method) will throw
    if the grid is still using Model extents. Calling SetDatumExtentType()
    on both ends first promotes the grid to a ViewSpecific override so the
    subsequent SetCurveInView() call is accepted by Revit.
    """
    try:
        grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
        grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
    except Exception as ex:
        logger.debug("SetDatumExtentType failed for grid {}: {}".format(
            grid.Id.IntegerValue, ex))


def apply_offset(entry, view, offset_distance, already_moved_keys):
    """Move the bubble endpoint perpendicular to the grid and write it back.

    CRITICAL FIX: The correct IronPython Revit API method is SetCurveInView
    (singular), NOT SetCurvesInView (plural). The plural form does not exist
    on the Grid class binding in IronPython and raises AttributeError.

    Signature: grid.SetCurveInView(DatumExtentType, View, Curve)
    — takes a single Curve directly, not an IList.

    Also calls promote_to_view_specific() first so the grid accepts the
    ViewSpecific write even if it was previously on Model extents.

    Returns True if a change was written, False if skipped.
    """
    key = (entry['grid_id'], entry['end_index'])
    if key in already_moved_keys:
        return False

    grid      = entry['grid']
    curve     = entry['curve']
    end_index = entry['end_index']

    offset_vec = perpendicular_offset_vector(curve, view, offset_distance)
    bubble_pt  = curve.GetEndPoint(end_index)
    fixed_pt   = curve.GetEndPoint(1 - end_index)
    new_bubble = bubble_pt + offset_vec

    if end_index == 0:
        new_curve = Line.CreateBound(new_bubble, fixed_pt)
    else:
        new_curve = Line.CreateBound(fixed_pt, new_bubble)

    # Promote to ViewSpecific BEFORE writing — required when the grid
    # was previously using Model extents (the default state).
    promote_to_view_specific(grid, view)

    # SetCurveInView — SINGULAR. Takes (DatumExtentType, View, Curve).
    # This is the correct IronPython binding. SetCurvesInView (plural)
    # does not exist on the Grid class.
    grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_curve)

    already_moved_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Collect qualifying views ----------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert(
            "No floor plan, ceiling plan, area plan, or engineering plan "
            "views placed on sheets were found in this project.",
            title="Nothing to do",
        )
        script.exit()

    # ---- 2. Stats -----------------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    collisions_fixed = 0
    per_view_errors  = []

    # ---- 3. Single transaction — one Ctrl+Z undoes everything ---------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold   = bubble_diameter_model_units(view)
                offset_dist = offset_distance_model_units(view)

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                moved_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    try:
                        if apply_offset(target, view, offset_dist, moved_keys):
                            collisions_fixed += 1
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
            title="Error — Grid Bubble Separation",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    # ---- 4. Results ---------------------------------------------------------
    summary = "\n".join([
        "Views processed:  {}".format(views_processed),
        "Collisions found: {}".format(collisions_found),
        "Collisions fixed: {}".format(collisions_fixed),
        "Errors:           {}".format(len(per_view_errors)),
    ])

    if per_view_errors:
        output.print_md("### Grid Bubble Separation — Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee the pyRevit output window for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()