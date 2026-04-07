# encoding: utf-8
"""Pipe Takeoffs - Toggle tool for automated domestic water stub-outs.

Automates branch pipe takeoff placement from domestic water mains.
Click 1: Select existing CW/HW/HWC main pipe
Click 2: Pick destination point (fixture/wall location)
Script builds: tee, 6in rise, elbow, horizontal run, elbow, drop to AFF,
               elbow turning toward main, 6in stub-in.

Assembly: 4 pipe segments + 1 tee + 3 elbows
All branch pipe properties copied from the clicked main pipe.

On activation a fixture picker dialog appears. ESC or re-click to deactivate.
"""

# ============================================================================
# IMPORTS
# ============================================================================
import clr
import math
from collections import OrderedDict

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')

from Autodesk.Revit.DB import (
    XYZ,
    ElementId,
    BuiltInParameter,
    BuiltInCategory,
    Transaction
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import (
    OperationCanceledException,
    InvalidOperationException
)

from pyrevit import revit, DB, UI, script, forms

# ============================================================================
# CONSTANTS
# ============================================================================
ENVVAR_ACTIVE  = "PIPE_TAKEOFFS_ACTIVE"
ENVVAR_FIXTURE = "PIPE_TAKEOFFS_FIXTURE"

RISE_HEIGHT       = 0.5    # 6 inches in feet
STUB_LENGTH       = 0.5    # 6 inches in feet
DIAGONAL_WARN_DEG = 5.0

DEFAULT_FIXTURE = "Lavatory"

# Fixture presets: label -> (nominal diameter inches, AFF inches)
FIXTURES = OrderedDict([
    ('WC - Tank',         (0.5,   19.0)),
    ('WC - Valve',        (1.5,   17.0)),
    ('Lavatory',          (0.5,   34.0)),
    ('Hand Sink',         (0.5,   34.0)),
    ('Shower',            (0.5,   48.0)),
    ('Mop Sink',          (0.75,  24.0)),
    ('Drinking Fountain', (0.5,   36.0)),
    ('Urinal',            (0.75,  24.0)),
])

VALID_SYSTEMS = ["CW", "HW", "HWC"]

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()


# ============================================================================
# FIXTURE PICKER DIALOG
# ============================================================================
def pick_fixture():
    """Show a SelectFromList dialog. Returns fixture label or None if cancelled."""
    saved = script.get_envvar(ENVVAR_FIXTURE) or DEFAULT_FIXTURE
    if saved not in FIXTURES:
        saved = DEFAULT_FIXTURE

    selected = forms.SelectFromList.show(
        list(FIXTURES.keys()),
        title="Pipe Takeoffs - Select Fixture",
        button_name="Start",
        default=saved,
    )
    return selected  # None if user closed/cancelled


# ============================================================================
# SELECTION FILTER - CW / HW / HWC pipes only
# ============================================================================
class WaterPipeFilter(ISelectionFilter):

    def AllowElement(self, element):
        cat = element.Category
        if cat is None:
            return False
        if cat.Id.IntegerValue != int(BuiltInCategory.OST_PipeCurves):
            return False

        sys_param = element.get_Parameter(
            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
        )
        if sys_param is None:
            return False

        sys_type_id = sys_param.AsElementId()
        if sys_type_id == ElementId.InvalidElementId:
            return False

        sys_type = doc.GetElement(sys_type_id)
        if sys_type is None:
            return False

        sys_name   = ""
        abbr_param = sys_type.get_Parameter(
            BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM
        )
        if abbr_param and abbr_param.AsString():
            sys_name = abbr_param.AsString().strip().upper()
        else:
            name_param = sys_type.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            )
            if name_param and name_param.AsString():
                sys_name = name_param.AsString().strip().upper()

        return any(v in sys_name for v in VALID_SYSTEMS)

    def AllowReference(self, reference, position):
        return False


# ============================================================================
# GEOMETRY HELPERS
# ============================================================================
def project_point_onto_line(point, line_start, line_end):
    line_vec       = line_end - line_start
    point_vec      = point - line_start
    line_length_sq = line_vec.DotProduct(line_vec)
    if line_length_sq < 1e-10:
        return line_start
    t = max(0.0, min(1.0, point_vec.DotProduct(line_vec) / line_length_sq))
    return XYZ(
        line_start.X + t * line_vec.X,
        line_start.Y + t * line_vec.Y,
        line_start.Z + t * line_vec.Z
    )


def get_perpendicular_toward_target(main_dir, tee_point, target_point):
    perp_a    = XYZ(-main_dir.Y,  main_dir.X, 0)
    perp_b    = XYZ( main_dir.Y, -main_dir.X, 0)
    to_target = XYZ(
        target_point.X - tee_point.X,
        target_point.Y - tee_point.Y,
        0
    )
    return perp_a.Normalize() if perp_a.DotProduct(to_target) >= 0 else perp_b.Normalize()


def check_diagonal_main(main_dir):
    abs_x = abs(main_dir.X)
    abs_y = abs(main_dir.Y)
    off_axis_deg = math.degrees(
        math.atan2(abs_y, abs_x) if abs_x >= abs_y else math.atan2(abs_x, abs_y)
    )
    if off_axis_deg > DIAGONAL_WARN_DEG:
        return (
            "Main pipe is {:.1f} deg off-axis.\n"
            "Branch will run perpendicular to the main, "
            "which may not be parallel to walls.\n\nContinue anyway?"
        ).format(off_axis_deg)
    return None


def get_open_connector_closest_to(element, target_point):
    best      = None
    best_dist = float('inf')
    for conn in element.ConnectorManager.Connectors:
        if conn.IsConnected:
            continue
        d = conn.Origin.DistanceTo(target_point)
        if d < best_dist:
            best_dist = d
            best      = conn
    return best


# ============================================================================
# PROPERTY COPY FROM MAIN PIPE
# ============================================================================
def copy_main_properties(pipe):
    location = pipe.Location
    if location is None:
        raise ValueError("Selected pipe has no location curve.")

    curve  = location.Curve
    start  = curve.GetEndPoint(0)
    end_pt = curve.GetEndPoint(1)

    raw_dir  = end_pt - start
    main_dir = XYZ(raw_dir.X, raw_dir.Y, 0).Normalize()

    if abs(end_pt.Z - start.Z) > 0.1:
        raise ValueError(
            "Selected pipe is not horizontal. "
            "Z difference: {:.2f} ft. Select a horizontal main."
            .format(abs(end_pt.Z - start.Z))
        )

    pipe_type_id   = pipe.GetTypeId()
    sys_param      = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
    system_type_id = sys_param.AsElementId() if sys_param else None

    lvl_param = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)
    level_id  = lvl_param.AsElementId() if lvl_param else None
    if level_id is None or level_id == ElementId.InvalidElementId:
        level_id = pipe.ReferenceLevel.Id

    dia_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if dia_param is None:
        raise ValueError("Cannot read pipe diameter.")
    diameter = dia_param.AsDouble()

    pipe_type_elem = doc.GetElement(pipe_type_id)
    pipe_type_name = "Unknown"
    if pipe_type_elem:
        name_param = pipe_type_elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if name_param and name_param.AsString():
            pipe_type_name = name_param.AsString()

    return {
        'pipe_type_id':   pipe_type_id,
        'system_type_id': system_type_id,
        'level_id':       level_id,
        'start':          start,
        'end':            end_pt,
        'direction':      main_dir,
        'centerline_z':   start.Z,
        'diameter':       diameter,
        'radius':         diameter / 2.0,
        'pipe_type_name': pipe_type_name,
        'element':        pipe,
    }


# ============================================================================
# ROUTING PREFERENCE PRE-CHECK
# ============================================================================
def check_routing_preferences(pipe_type_id, branch_dia_ft, pipe_type_name):
    warnings = []
    pipe_type_elem = doc.GetElement(pipe_type_id)
    if pipe_type_elem is None:
        warnings.append("Cannot read pipe type from main pipe.")
        return warnings

    rpm = None
    try:
        rpm = pipe_type_elem.RoutingPreferenceManager
    except Exception:
        pass

    if rpm is None:
        warnings.append(
            "Pipe type '{}' has no routing preferences. "
            "Fittings may not place correctly.".format(pipe_type_name)
        )
        return warnings

    try:
        from Autodesk.Revit.DB.Plumbing import RoutingPreferenceRuleGroupType
    except ImportError:
        return warnings

    checks = [
        (RoutingPreferenceRuleGroupType.Junctions, "tee"),
        (RoutingPreferenceRuleGroupType.Elbows,    "elbow"),
        (RoutingPreferenceRuleGroupType.Segments,  "segment"),
    ]
    for group, label in checks:
        try:
            if rpm.GetNumberOfRules(group) == 0:
                warnings.append(
                    "No {} rule in routing preferences for '{}'.".format(label, pipe_type_name)
                )
        except Exception:
            if label != "segment":
                warnings.append(
                    "Cannot read {} rules for '{}'.".format(label, pipe_type_name)
                )
    return warnings


# ============================================================================
# GEOMETRY CALCULATION
# ============================================================================
def calculate_takeoff_geometry(props, click1, click2, aff_height):
    cl_z        = props['centerline_z']
    main_radius = props['radius']
    main_dir    = props['direction']

    tee_center = project_point_onto_line(
        XYZ(click1.X, click1.Y, cl_z), props['start'], props['end']
    )

    rise_start = XYZ(tee_center.X, tee_center.Y, cl_z)
    rise_end   = XYZ(tee_center.X, tee_center.Y, cl_z + main_radius + RISE_HEIGHT)

    perp_dir       = get_perpendicular_toward_target(main_dir, tee_center, click2)
    to_target_xy   = XYZ(click2.X - tee_center.X, click2.Y - tee_center.Y, 0)
    horiz_distance = abs(to_target_xy.DotProduct(perp_dir))

    horiz_end  = XYZ(
        tee_center.X + perp_dir.X * horiz_distance,
        tee_center.Y + perp_dir.Y * horiz_distance,
        rise_end.Z
    )
    drop_start = XYZ(horiz_end.X, horiz_end.Y, horiz_end.Z)
    drop_end   = XYZ(horiz_end.X, horiz_end.Y, aff_height)
    stub_dir   = XYZ(-perp_dir.X, -perp_dir.Y, 0)
    stub_start = XYZ(drop_end.X, drop_end.Y, drop_end.Z)
    stub_end   = XYZ(
        drop_end.X + stub_dir.X * STUB_LENGTH,
        drop_end.Y + stub_dir.Y * STUB_LENGTH,
        drop_end.Z
    )

    if aff_height >= rise_end.Z:
        raise ValueError(
            "Drop terminus ({:.0f}\" AFF) is at or above the horizontal run. "
            "Main pipe is too low or AFF is too high.".format(aff_height * 12.0)
        )
    if horiz_distance < 0.083:
        raise ValueError("Destination point is too close to the main. Pick further away.")
    if (rise_end.Z - rise_start.Z) < 0.01:
        raise ValueError("Rise pipe has zero length. Check main pipe geometry.")
    if (drop_start.Z - drop_end.Z) < 0.01:
        raise ValueError("Drop pipe has zero length. Check AFF height.")

    return {
        'tee_center':     tee_center,
        'rise_start':     rise_start,
        'rise_end':       rise_end,
        'horiz_start':    rise_end,
        'horiz_end':      horiz_end,
        'drop_start':     drop_start,
        'drop_end':       drop_end,
        'stub_start':     stub_start,
        'stub_end':       stub_end,
        'perp_direction': perp_dir,
        'stub_direction': stub_dir,
        'horiz_distance': horiz_distance,
    }


# ============================================================================
# PIPE AND FITTING CREATION
# ============================================================================
def create_pipe_segment(sys_id, type_id, lvl_id, start_pt, end_pt, diameter_ft):
    new_pipe  = Pipe.Create(doc, sys_id, type_id, lvl_id, start_pt, end_pt)
    dia_param = new_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if dia_param:
        dia_param.Set(diameter_ft)
    return new_pipe


def split_main_and_place_tee(main_pipe, tee_center, rise_pipe):
    try:
        from Autodesk.Revit.DB.Plumbing import PlumbingUtils
        new_segment_id = PlumbingUtils.BreakCurve(doc, main_pipe.Id, tee_center)
    except ImportError:
        raise ValueError("PlumbingUtils not available in this Revit version.")

    new_segment = doc.GetElement(new_segment_id)
    if new_segment is None:
        raise ValueError(
            "BreakCurve failed. Click may be too close to a fitting or pipe end."
        )

    doc.Regenerate()

    conn_main_a = get_open_connector_closest_to(main_pipe,   tee_center)
    conn_main_b = get_open_connector_closest_to(new_segment, tee_center)
    conn_branch = get_open_connector_closest_to(rise_pipe,   tee_center)

    if conn_main_a is None:
        raise ValueError("No open connector on main segment A at split point.")
    if conn_main_b is None:
        raise ValueError("No open connector on main segment B at split point.")
    if conn_branch is None:
        raise ValueError("No open connector on rise pipe at tee point.")

    return doc.Create.NewTeeFitting(conn_main_a, conn_main_b, conn_branch)


def place_elbow(pipe_a, point_a, pipe_b, point_b):
    conn_a = get_open_connector_closest_to(pipe_a, point_a)
    conn_b = get_open_connector_closest_to(pipe_b, point_b)
    if conn_a is None:
        raise ValueError("No open connector on pipe near {}".format(point_a))
    if conn_b is None:
        raise ValueError("No open connector on pipe near {}".format(point_b))
    return doc.Create.NewElbowFitting(conn_a, conn_b)


# ============================================================================
# MAIN BUILD FUNCTION
# ============================================================================
def build_takeoff(main_pipe, click1, click2, branch_dia_ft, aff_height):
    props    = copy_main_properties(main_pipe)
    warnings = check_routing_preferences(
        props['pipe_type_id'], branch_dia_ft, props['pipe_type_name']
    )
    if warnings:
        msg = "Routing preference warnings for '{}':\n\n".format(props['pipe_type_name'])
        msg += "".join("- {}\n".format(w) for w in warnings)
        msg += "\nContinue anyway? Fittings may fail to place."
        if not forms.alert(msg, yes=True, no=True):
            return

    diag_warning = check_diagonal_main(props['direction'])
    if diag_warning:
        if not forms.alert(diag_warning, yes=True, no=True):
            return

    geo     = calculate_takeoff_geometry(props, click1, click2, aff_height)
    sys_id  = props['system_type_id']
    type_id = props['pipe_type_id']
    lvl_id  = props['level_id']

    rise_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['rise_start'], geo['rise_end'], branch_dia_ft
    )
    doc.Regenerate()
    split_main_and_place_tee(props['element'], geo['tee_center'], rise_pipe)

    horiz_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['horiz_start'], geo['horiz_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(rise_pipe, geo['rise_end'], horiz_pipe, geo['horiz_start'])

    drop_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['drop_start'], geo['drop_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(horiz_pipe, geo['horiz_end'], drop_pipe, geo['drop_start'])

    stub_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['stub_start'], geo['stub_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(drop_pipe, geo['drop_end'], stub_pipe, geo['stub_start'])

    logger.debug(
        "Takeoff complete: {} - {:.3f} ft dia, {:.1f} in AFF"
        .format(props['pipe_type_name'], branch_dia_ft, aff_height * 12.0)
    )


# ============================================================================
# VIEW VALIDATION
# ============================================================================
def validate_view():
    view = doc.ActiveView
    if view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.EngineeringPlan]:
        forms.alert(
            "Pipe Takeoffs requires a Floor Plan or Engineering Plan view.\n\n"
            "Current view: {}\n\nSwitch to a plan view and try again."
            .format(view.ViewType),
            title="Wrong View Type"
        )
        return False
    return True


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    is_active = script.get_envvar(ENVVAR_ACTIVE)

    if is_active:
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)
        return

    if not validate_view():
        return

    # Show fixture picker - single dialog on activation only
    fixture_name = pick_fixture()
    if fixture_name is None:
        return  # user cancelled

    # Persist selection for next activation
    script.set_envvar(ENVVAR_FIXTURE, fixture_name)

    dia_in, aff_in = FIXTURES[fixture_name]
    branch_dia_ft  = dia_in / 12.0
    aff_height_ft  = aff_in / 12.0

    script.set_envvar(ENVVAR_ACTIVE, True)
    script.toggle_icon(True)

    pipe_filter = WaterPipeFilter()

    try:
        while True:
            try:
                # Click 1: pick main pipe
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    pipe_filter,
                    "Pick CW/HW/HWC main pipe  |  Fixture: {}  (ESC to exit)"
                    .format(fixture_name)
                )
                main_pipe = doc.GetElement(ref.ElementId)
                click1    = ref.GlobalPoint

                if main_pipe is None:
                    forms.alert("Could not read selected pipe. Try again.")
                    continue

                # Click 2: pick destination
                click2 = uidoc.Selection.PickPoint(
                    "Pick destination point  |  Fixture: {}  (ESC to cancel)"
                    .format(fixture_name)
                )

                with revit.Transaction("Pipe Takeoff"):
                    build_takeoff(
                        main_pipe, click1, click2,
                        branch_dia_ft, aff_height_ft
                    )

            except OperationCanceledException:
                break

            except InvalidOperationException as ex:
                forms.alert(
                    "Operation interrupted:\n{}\n\nTool will deactivate.".format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                break

            except ValueError as ex:
                forms.alert(str(ex), title="Pipe Takeoff - Invalid Geometry")
                continue

            except Exception as ex:
                forms.alert(
                    "Unexpected error:\n{}\n\nTakeoff rolled back. You can try again."
                    .format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                logger.error("Pipe Takeoff error: {}".format(str(ex)))
                continue

    finally:
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)


# ============================================================================
# RUN
# ============================================================================
if __name__ == "__main__" or True:
    main()