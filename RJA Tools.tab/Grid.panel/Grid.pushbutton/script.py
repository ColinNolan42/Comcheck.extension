# -*- coding: utf-8 -*-
__title__   = "Leader\nGeom Diag"
__author__  = "MEP Tools"
__version__ = "diag3"
__doc__     = "Prints exact Anchor/Elbow/End geometry after AddLeader. Rolls back."

import traceback
from Autodesk.Revit.DB import (
    FilteredElementCollector, Grid, View, ViewSheet, ViewType,
    XYZ, Transaction, DatumExtentType, DatumEnds, ElementId,
)
from pyrevit import forms, script, revit

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

PLAN_VIEW_TYPES = {
    ViewType.FloorPlan, ViewType.CeilingPlan,
    ViewType.AreaPlan,  ViewType.EngineeringPlan,
}

def get_sheet_view_ids():
    ids = set()
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements():
        try:
            for vid in sheet.GetAllPlacedViews():
                ids.add(vid.IntegerValue)
        except Exception:
            pass
    return ids

def get_curve(grid, view):
    for et in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(et, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None

def fmt(pt):
    if pt is None:
        return "None"
    return "({:.3f}, {:.3f}, {:.3f})".format(pt.X, pt.Y, pt.Z)

def test_grid(g, view, label):
    output.print_md("\n## {} (ID {})".format(label, g.Id.IntegerValue))
    curve = get_curve(g, view)
    if curve:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = abs(p1.X - p0.X)
        dy = abs(p1.Y - p0.Y)
        orientation = "VERTICAL (dy>=dx)" if dy >= dx else "HORIZONTAL (dx>dy)"
        output.print_md("- Orientation: **{}**".format(orientation))
        output.print_md("- Curve End0: {}".format(fmt(p0)))
        output.print_md("- Curve End1: {}".format(fmt(p1)))
    for end_index in (0, 1):
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        try:
            vis = g.IsBubbleVisibleInView(end, view)
            has = g.GetLeader(end, view) is not None
            output.print_md("- End{}: bubble={} leader={}".format(end_index, vis, has))
        except Exception as ex:
            output.print_md("- End{}: ERROR {}".format(end_index, ex))

def main():
    sheet_ids = get_sheet_view_ids()
    test_view = None
    for v in FilteredElementCollector(doc).OfClass(View):
        if not v.IsTemplate and v.ViewType in PLAN_VIEW_TYPES:
            if v.Id.IntegerValue in sheet_ids:
                test_view = v
                break

    if test_view is None:
        forms.alert("No plan views on sheets found.")
        script.exit()

    output.print_md("# Leader Geometry Diagnostic")
    output.print_md("**View:** {} (Scale: {})".format(test_view.Name, test_view.Scale))

    # Test all 6 failing grids
    grid_ids = [3707749, 3707752, 3707753, 3707756, 3707761, 3707763]
    grids = []
    for gid in grid_ids:
        g = doc.GetElement(ElementId(gid))
        if g:
            grids.append(g)
            test_grid(g, test_view, "Grid {} ({})".format(g.Name, gid))

    output.print_md("\n---\n# AddLeader Test on each grid (ROLLED BACK)")

    t = Transaction(doc, "Leader Geom Diag - ROLLBACK")
    try:
        t.Start()

        for g in grids:
            output.print_md("\n## Testing AddLeader on Grid {} ({})".format(
                g.Name, g.Id.IntegerValue))

            # Find which end has the bubble
            bubble_end = None
            bubble_end_index = None
            for ei in (0, 1):
                end = DatumEnds.End0 if ei == 0 else DatumEnds.End1
                try:
                    if g.IsBubbleVisibleInView(end, test_view):
                        bubble_end = end
                        bubble_end_index = ei
                        break
                except Exception:
                    pass

            if bubble_end is None:
                output.print_md("- No visible bubble — skipping")
                continue

            output.print_md("- Bubble at End{}".format(bubble_end_index))

            # Check if leader already exists
            existing = g.GetLeader(bubble_end, test_view)
            if existing:
                output.print_md("- Already has leader:")
                output.print_md("  - Anchor: {}".format(fmt(existing.Anchor)))
                output.print_md("  - Elbow:  {}".format(fmt(existing.Elbow)))
                output.print_md("  - End:    {}".format(fmt(existing.End)))
            else:
                try:
                    g.AddLeader(bubble_end, test_view)
                    ldr = g.GetLeader(bubble_end, test_view)
                    if ldr:
                        output.print_md("- AddLeader SUCCESS — default geometry:")
                        output.print_md("  - Anchor: {}".format(fmt(ldr.Anchor)))
                        output.print_md("  - Elbow:  {}".format(fmt(ldr.Elbow)))
                        output.print_md("  - End:    {}".format(fmt(ldr.End)))

                        # Compute direction info
                        anchor = ldr.Anchor
                        elbow  = ldr.Elbow
                        end    = ldr.End
                        if anchor and elbow and end:
                            ae_dx = elbow.X - anchor.X
                            ae_dy = elbow.Y - anchor.Y
                            ef_dx = end.X   - elbow.X
                            ef_dy = end.Y   - elbow.Y
                            output.print_md("  - Anchor->Elbow delta: ({:.3f}, {:.3f})".format(ae_dx, ae_dy))
                            output.print_md("  - Elbow->End delta:    ({:.3f}, {:.3f})".format(ef_dx, ef_dy))

                            # Try moving Elbow only in +X direction
                            new_elbow_x = XYZ(elbow.X + 4.0, elbow.Y, elbow.Z)
                            ldr.Elbow = new_elbow_x
                            try:
                                g.SetLeader(bubble_end, test_view, ldr)
                                output.print_md("  - SetLeader Elbow+X: **SUCCESS**")
                            except Exception as ex:
                                output.print_md("  - SetLeader Elbow+X: **FAILED** — {}".format(ex))
                                # Reset and try Elbow in +Y
                                ldr2 = g.GetLeader(bubble_end, test_view)
                                new_elbow_y = XYZ(ldr2.Elbow.X, ldr2.Elbow.Y + 4.0, ldr2.Elbow.Z)
                                ldr2.Elbow = new_elbow_y
                                try:
                                    g.SetLeader(bubble_end, test_view, ldr2)
                                    output.print_md("  - SetLeader Elbow+Y: **SUCCESS**")
                                except Exception as ex2:
                                    output.print_md("  - SetLeader Elbow+Y: **FAILED** — {}".format(ex2))

                except Exception as ex:
                    output.print_md("- AddLeader FAILED: {}".format(ex))

        t.RollBack()
        output.print_md("\n---\nTransaction rolled back — no permanent changes.")

    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        output.print_md("**Transaction error**: {}".format(ex))
        output.print_md(traceback.format_exc())

if __name__ == "__main__":
    main()