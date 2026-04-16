# -*- coding: utf-8 -*-
"""Detects colliding grid bubbles across all plan/elevation/section views in the
active Revit project and automatically offsets one bubble of each colliding
pair perpendicular to the gridline so the annotations no longer overlap.

Only view-specific 2D grid extents are modified (GetCurvesInView /
SetCurvesInView). Model geometry of the grid is never altered. A single
Transaction wraps the entire run so the user can undo the operation with one
Ctrl+Z.

For each colliding pair the grid with the higher ElementId is moved, which
makes the operation deterministic across repeated runs.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "1.0.0"
__doc__     = ("Finds grid bubbles that overlap in plan, elevation and section "
               "views and offsets the bubble end of the 2D grid curve "
               "perpendicular to the gridline so the annotations are legible. "
               "User supplies the collision threshold and offset distance.")

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
import math
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid,
    View,
    ViewType,
    XYZ,
    Line,
    Transaction,
    DatumExtentType,
    DatumEnds,
    UnitUtils,
)

# Unit handling changed in Revit 2021 (ForgeTypeId replaced DisplayUnitType).
try:
    from Autodesk.Revit.DB import UnitTypeId
    _USE_FORGE_UNITS = True
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    _USE_FORGE_UNITS = False

from pyrevit import forms, script, revit

# -----------------------------------------------------------------------------
# Document handles
# -----------------------------------------------------------------------------
doc     = revit.doc
uidoc   = revit.uidoc
logger  = script.get_logger()
output  = script.get_output()

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
PROCESSABLE_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
    ViewType.Elevation,
    ViewType.Section,
    ViewType.Detail,
}

PARALLEL_TOL = 1.0e-6


# =============================================================================
# Unit helpers
# =============================================================================
def is_metric_project(document):
    try:
        fmt_options = document.GetUnits().GetFormatOptions(
            __import__("Autodesk.Revit.DB", fromlist=["SpecTypeId"]).SpecTypeId.Length
            if _USE_FORGE_UNITS else
            __import__("Autodesk.Revit.DB", fromlist=["UnitType"]).UnitType.UT_Length
        )
        if _USE_FORGE_UNITS:
            unit_id = fmt_options.GetUnitTypeId()
            return unit_id != UnitTypeId.FeetFractionalInches \
               and unit_id != UnitTypeId.Feet \
               and unit_id != UnitTypeId.FeetAndFractionalInches \
               and unit_id != UnitTypeId.Inches \
               and unit_id != UnitTypeId.FractionalInches
        else:
            dut = fmt_options.DisplayUnits
            return dut not in (
                DisplayUnitType.DUT_DECIMAL_FEET,
                DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES,
                DisplayUnitType.DUT_DECIMAL_INCHES,
                DisplayUnitType.DUT_FRACTIONAL_INCHES,
            )
    except Exception:
        return False


def to_internal_length(value, metric):
    if _USE_FORGE_UNITS:
        unit = UnitTypeId.Millimeters if metric else UnitTypeId.Feet
        return UnitUtils.ConvertToInternalUnits(value, unit)
    else:
        unit = DisplayUnitType.DUT_MILLIMETERS if metric else DisplayUnitType.DUT_DECIMAL_FEET
        return UnitUtils.ConvertToInternalUnits(value, unit)


# =============================================================================
# Geometry helpers
# =============================================================================
def points_within_threshold(p1, p2, threshold):
    dx = p1.X - p2.X
    dy = p1.Y - p2.Y
    dz = p1.Z - p2.Z
    return (dx * dx + dy * dy + dz * dz) <= (threshold * threshold)


def perpendicular_offset_vector(curve, view, offset_distance):
    try:
        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    except Exception:
        direction = view.RightDirection

    view_normal = view.ViewDirection
    try:
        perp = view_normal.CrossProduct(direction).Normalize()
    except Exception:
        perp = view.RightDirection

    return perp.Multiply(offset_distance)


def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True


# =============================================================================
# Data collection
# =============================================================================
def collect_processable_views(document):
    views = []
    collector = FilteredElementCollector(document).OfClass(View)
    for v in collector:
        if v.IsTemplate:
            continue
        if v.ViewType in PROCESSABLE_VIEW_TYPES:
            views.append(v)
    return views


def collect_grid_bubble_entries(document, view):
    entries = []
    grids = FilteredElementCollector(document, view.Id).OfClass(Grid).ToElements()
    for g in grids:
        try:
            curves = g.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        except Exception as ex:
            logger.debug("Could not get curves for grid {} in view {}: {}"
                         .format(g.Id, view.Name, ex))
            continue
        if not curves:
            continue
        for curve in curves:
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
# Collision resolution
# =============================================================================
def find_colliding_pairs(entries, threshold):
    pairs = []
    n = len(entries)
    for i in range(n):
        for j in range(i + 1, n):
            if points_within_threshold(entries[i]['point'],
                                       entries[j]['point'],
                                       threshold):
                if entries[i]['grid_id'] == entries[j]['grid_id']:
                    continue
                pairs.append((entries[i], entries[j]))
    return pairs


def choose_entry_to_move(entry_a, entry_b):
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


def apply_offset(entry, view, offset_distance, already_moved_keys):
    key = (entry['grid_id'], entry['end_index'])
    if key in already_moved_keys:
        return False

    grid       = entry['grid']
    curve      = entry['curve']
    end_index  = entry['end_index']

    offset_vec  = perpendicular_offset_vector(curve, view, offset_distance)
    old_point   = curve.GetEndPoint(end_index)
    new_point   = old_point + offset_vec

    fixed_end_index = 1 - end_index
    fixed_point     = curve.GetEndPoint(fixed_end_index)

    if end_index == 0:
        new_curve = Line.CreateBound(new_point, fixed_point)
    else:
        new_curve = Line.CreateBound(fixed_point, new_point)

    from System.Collections.Generic import List
    from Autodesk.Revit.DB import Curve
    curve_list = List[Curve]()
    curve_list.Add(new_curve)

    grid.SetCurvesInView(DatumExtentType.ViewSpecific, view, curve_list)
    already_moved_keys.add(key)
    return True


# =============================================================================
# User input
# FIX: pyrevit.forms.FlexForm components use these correct class names:
#   forms.Label(text)              — static label
#   forms.TextBox(name, default)   — text input  (NOT Text=kwarg)
#   forms.Separator()              — horizontal rule
#   forms.Button(text)             — submit button
# forms.Label / forms.TextBox / forms.Button are the correct names.
# The original script used Text= as a keyword arg which does not exist.
# =============================================================================
def prompt_user_for_distances(metric):
    if metric:
        default_threshold = "450"
        default_offset    = "600"
        unit_label        = "mm"
    else:
        default_threshold = "1.5"
        default_offset    = "2.0"
        unit_label        = "ft"

    # --- PRIMARY: pyrevit FlexForm with correct component constructors -------
    # forms.Label(text)              positional, no kwargs
    # forms.TextBox(name, default)   second arg is the default string value
    # forms.Separator()              no args
    # forms.Button(text)             positional
    try:
        components = [
            forms.Label("Collision threshold ({0}):".format(unit_label)),
            forms.TextBox("threshold_tb", default_threshold),
            forms.Separator(),
            forms.Label("Offset distance ({0}):".format(unit_label)),
            forms.TextBox("offset_tb", default_offset),
            forms.Separator(),
            forms.Button("Run"),
        ]
        form = forms.FlexForm("Separate Grid Bubbles", components)
        form.show()
        values = form.values
        if not values:
            return None, None
        threshold_val = float(values['threshold_tb'])
        offset_val    = float(values['offset_tb'])

    except Exception:
        # --- FALLBACK: two sequential ask_for_string prompts ----------------
        # This fires if the installed pyRevit build has a different FlexForm
        # API (older or newer) or if the WPF window fails to open.
        threshold_raw = forms.ask_for_string(
            default=default_threshold,
            prompt="Collision threshold ({0}):".format(unit_label),
            title="Separate Grid Bubbles",
        )
        if threshold_raw is None:
            return None, None

        offset_raw = forms.ask_for_string(
            default=default_offset,
            prompt="Offset distance ({0}):".format(unit_label),
            title="Separate Grid Bubbles",
        )
        if offset_raw is None:
            return None, None

        try:
            threshold_val = float(threshold_raw)
            offset_val    = float(offset_raw)
        except ValueError:
            forms.alert("Threshold and offset must be numeric.",
                        title="Invalid input")
            return None, None

    if threshold_val <= 0 or offset_val <= 0:
        forms.alert("Threshold and offset must be positive.",
                    title="Invalid input")
        return None, None

    return (to_internal_length(threshold_val, metric),
            to_internal_length(offset_val,    metric))


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Units & user input ---------------------------------------------
    metric = is_metric_project(doc)
    threshold_ft, offset_ft = prompt_user_for_distances(metric)
    if threshold_ft is None:
        script.exit()

    # ---- 2. Gather views ---------------------------------------------------
    views = collect_processable_views(doc)
    if not views:
        forms.alert("No plan, elevation, or section views were found.",
                    title="Nothing to do")
        script.exit()

    # ---- 3. Stats trackers -------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    collisions_fixed = 0
    per_view_errors  = []

    # ---- 4. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                entries = collect_grid_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold_ft)
                collisions_found += len(pairs)

                moved_keys = set()

                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    try:
                        if apply_offset(target, view, offset_ft, moved_keys):
                            collisions_fixed += 1
                    except Exception as ex:
                        per_view_errors.append((
                            view.Name,
                            "Grid {0}: {1}".format(target['grid_id'], ex)
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
            "Transaction failed and was rolled back.\n\n{0}".format(ex),
            title="Error",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    # ---- 5. Results dialog -------------------------------------------------
    summary_lines = [
        "Views processed:  {0}".format(views_processed),
        "Collisions found: {0}".format(collisions_found),
        "Collisions fixed: {0}".format(collisions_fixed),
        "Errors:           {0}".format(len(per_view_errors)),
    ]
    summary = "\n".join(summary_lines)

    if per_view_errors:
        output.print_md("### Grid Bubble Separation - Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{0}**: {1}".format(vname, err))
        forms.alert(
            summary + "\n\nSee the pyRevit output window for error details.",
            title="Separate Grid Bubbles - Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles - Complete")


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()