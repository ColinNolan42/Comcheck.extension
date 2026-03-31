# encoding: utf-8
# COMCHECK PDF PLACER - pyRevit Script
# DIAGNOSTIC VERSION - prints all available constructors then exits

import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, revit
import System
from System import Array

doc = revit.doc
uidoc = revit.uidoc

# Print all available constructors for ImageTypeOptions
ito_type = clr.GetClrType(ImageTypeOptions)
ctors = ito_type.GetConstructors()

msg = "ImageTypeOptions constructors found:\n\n"
for c in ctors:
    params = c.GetParameters()
    param_str = ", ".join(["{} {}".format(p.ParameterType, p.Name) for p in params])
    msg += "({})".format(param_str) + "\n"

forms.alert(msg, title="Constructor Diagnostics")