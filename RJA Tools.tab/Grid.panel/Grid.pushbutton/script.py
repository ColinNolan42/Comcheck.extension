# Quick API probe — check what elbow/break methods exist on Grid
__title__ = "Grid\nAPI Probe"
__author__ = "MEP Tools"
__doc__    = "Probe Grid API for elbow/break methods"

from Autodesk.Revit.DB import FilteredElementCollector, Grid, ViewType, DatumExtentType, DatumEnds
from pyrevit import script, revit, forms

doc   = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# Get one grid
grids = FilteredElementCollector(doc).OfClass(Grid).ToElements()
if not grids:
    forms.alert("No grids found")
    script.exit()

g = grids[0]
output.print_md("## Grid API Methods containing 'elbow', 'break', 'kink', 'offset', 'bend', 'split'")
keywords = ['elbow','break','kink','offset','bend','split','head','bubble','datum','extent','curve','end']
for name in sorted(dir(g)):
    low = name.lower()
    if any(k in low for k in keywords):
        output.print_md("- `{}`".format(name))

output.print_md("\n## All Grid methods")
for name in sorted(dir(g)):
    if not name.startswith('_'):
        output.print_md("- `{}`".format(name))