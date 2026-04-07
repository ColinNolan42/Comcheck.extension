# encoding: utf-8
"""ui.py - Runs at pyRevit load time.
Adds a Fixture ComboBox to the Pipes panel on the Comcheck tab.
The selected fixture name is stored in an envvar that script.py reads.
"""
import clr
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import ComboBoxData, ComboBoxMemberData
from pyrevit import script, HOST_APP
from collections import OrderedDict

ENVVAR_FIXTURE     = "PIPE_TAKEOFFS_FIXTURE"
COMBO_FIXTURE_NAME = "PipeTakeoffsFixtureCombo"
DEFAULT_FIXTURE    = "Lavatory"

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


class FixtureComboEventHandler(object):
    """Called by Revit when the user changes the combo selection."""

    def __init__(self):
        self.__implementation__ = self

    def Execute(self, sender, args):
        current = args.NewValue
        if current is not None:
            script.set_envvar(ENVVAR_FIXTURE, current.Name)

    def GetName(self):
        return "PipeTakeoffsFixtureCombo"


def create_fixture_combo(panel):
    """Add the fixture ComboBox to the given RibbonPanel."""
    # Check if already exists (pyRevit hot-reload guard)
    try:
        existing_items = panel.GetItems()
        if existing_items:
            for item in existing_items:
                if item.Name == COMBO_FIXTURE_NAME:
                    return item
    except Exception:
        pass

    combo_data = ComboBoxData(COMBO_FIXTURE_NAME)
    combo      = panel.AddItem(combo_data)

    members = [ComboBoxMemberData(label, label) for label in FIXTURES.keys()]
    combo.AddItems(members)

    # Set default selection
    try:
        for item in combo.GetItems():
            if item.Name == DEFAULT_FIXTURE:
                combo.Current = item
                break
    except Exception:
        pass

    # Write default to envvar immediately
    saved = script.get_envvar(ENVVAR_FIXTURE)
    if not saved or saved not in FIXTURES:
        script.set_envvar(ENVVAR_FIXTURE, DEFAULT_FIXTURE)

    # Wire up change handler
    try:
        handler = FixtureComboEventHandler()
        combo.CurrentChanged += handler.Execute
    except Exception:
        pass

    return combo


# pyRevit calls setup() automatically from ui.py at extension load
def setup(panel):
    create_fixture_combo(panel)
