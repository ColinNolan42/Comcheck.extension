# -*- coding: utf-8 -*-
"""DIAGNOSTIC VERSION — prints full details about grid 3707753 and any
other grids that fail, so we can understand why SetCurveInView rejects them.
"""

__title__   = "Grid\nDiagnostic"
__author__  = "MEP Tools"
__version__ = "diag"
__doc__     = "Diagnostic: prints geometry details for problem grids."

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
    ElementId,
)

from pyrevit import forms, script, revit

doc    = revit.doc
output = script.get_output()

PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

BUBBLE_DIAMETER_INCHES = 0.375
OFFSET_MULTIPLIER      = 1.5

def bubble_diameter_model_units(view):
    try:
        scale = float(view.Scale)
    except Exception:
        scale = 96.0
    return (BUBBLE_DIAMETER_INCHES / 12.0) * scale

def get_sheet_view_ids(document):
    placed_ids = set()
    sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    for sheet in sheets:
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

def get_grid_curve_in_view(grid, view):
    for extent_type in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(extent_type, view)
            if curves:
                return curves[0], extent_type
        except Exception:
            continue
    return None, None

def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True

def main():
    # Find grid 3707753 directly
    target_id = ElementId(3707753)
    target_grid = doc.GetElement(target_id)

    output.print_md("## Grid 3707753 Diagnostic")

    if target_grid is None:
        output.print_md("**ERROR: Grid 3707753 not found in document.**")
        return

    output.print_md("- **Type**: {}".format(type(target_grid)))
    output.print_md("- **Category**: {}".format(target_grid.Category.Name if target_grid.Category else "None"))

    # Print the grid's model curve
    try:
        output.print_md("- **IsCurved**: {}".format(target_grid.IsCurved))
    except Exception as ex:
        output.print_md("- **IsCurved**: N/A ({})".format(ex))

    # Get one plan view on a sheet to test with
    views = collect_plan_views_on_sheets(doc)
    test_view = views[0] if views else None

    if test_view is None:
        output.print_md("**No plan views on sheets found.**")
        return

    output.print_md("- **Testing in view**: {}".format(test_view.Name))
    output.print_md("- **View Scale**: {}".format(test_view.Scale))

    # Read curve before promotion
    curve_before, extent_before = get_grid_curve_in_view(target_grid, test_view)
    if curve_before:
        p0 = curve_before.GetEndPoint(0)
        p1 = curve_before.GetEndPoint(1)
        output.print_md("- **Curve extent type (before)**: {}".format(extent_before))
        output.print_md("- **End0 before**: ({:.4f}, {:.4f}, {:.4f})".format(p0.X, p0.Y, p0.Z))
        output.print_md("- **End1 before**: ({:.4f}, {:.4f}, {:.4f})".format(p1.X, p1.Y, p1.Z))
        output.print_md("- **Curve type**: {}".format(type(curve_before)))
        try:
            output.print_md("- **Length**: {:.4f} ft".format(curve_before.Length))
        except Exception:
            pass
    else:
        output.print_md("- **No curve found before promotion**")

    # Check bubble visibility
    for end_idx in (0, 1):
        vis = grid_has_bubble_at_end(target_grid, test_view, end_idx)
        output.print_md("- **Bubble at end {}**: {}".format(end_idx, vis))

    # Try promotion
    output.print_md("\n### Attempting SetDatumExtentType promotion...")
    t = Transaction(doc, "Grid Diag Test")
    try:
        t.Start()
        try:
            target_grid.SetDatumExtentType(DatumEnds.End0, test_view, DatumExtentType.ViewSpecific)
            target_grid.SetDatumExtentType(DatumEnds.End1, test_view, DatumExtentType.ViewSpecific)
            output.print_md("- Promotion: **SUCCESS**")
        except Exception as ex:
            output.print_md("- Promotion: **FAILED** — {}".format(ex))

        # Re-read after promotion
        try:
            curves_after = target_grid.GetCurvesInView(DatumExtentType.ViewSpecific, test_view)
            if curves_after:
                c = curves_after[0]
                p0a = c.GetEndPoint(0)
                p1a = c.GetEndPoint(1)
                output.print_md("- **End0 after promotion**: ({:.4f}, {:.4f}, {:.4f})".format(p0a.X, p0a.Y, p0a.Z))
                output.print_md("- **End1 after promotion**: ({:.4f}, {:.4f}, {:.4f})".format(p1a.X, p1a.Y, p1a.Z))

                # Try building and writing a test line extended by 1ft along axis
                try:
                    direction = (p1a - p0a).Normalize()
                    new_p1 = p1a + direction.Multiply(1.0)
                    test_line = Line.CreateBound(p0a, new_p1)
                    target_grid.SetCurveInView(DatumExtentType.ViewSpecific, test_view, test_line)
                    output.print_md("- SetCurveInView test extension: **SUCCESS**")
                except Exception as ex:
                    output.print_md("- SetCurveInView test extension: **FAILED** — {}".format(ex))
            else:
                output.print_md("- **No ViewSpecific curve after promotion**")
        except Exception as ex:
            output.print_md("- Re-read after promotion: **FAILED** — {}".format(ex))

        t.RollBack()
        output.print_md("- Transaction rolled back (no permanent changes)")

    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        output.print_md("**Transaction error**: {}".format(ex))

    output.print_md("\n---\n### All grids in first view")
    grids_in_view = FilteredElementCollector(doc, test_view.Id).OfClass(Grid).ToElements()
    output.print_md("Total grids visible: **{}**".format(len(grids_in_view)))
    for g in grids_in_view:
        try:
            is_curved = g.IsCurved
        except Exception:
            is_curved = "?"
        curve, ext = get_grid_curve_in_view(g, test_view)
        pt = ""
        if curve:
            try:
                p = curve.GetEndPoint(0)
                pt = "End0=({:.2f},{:.2f})".format(p.X, p.Y)
            except Exception:
                pt = "curve read error"
        output.print_md("- ID: **{}** | Name: {} | IsCurved: {} | Extent: {} | {}".format(
            g.Id.IntegerValue, g.Name, is_curved, ext, pt))

if __name__ == "__main__":
    main()