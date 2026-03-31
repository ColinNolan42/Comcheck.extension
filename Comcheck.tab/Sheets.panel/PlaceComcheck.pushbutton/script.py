# encoding: utf-8
# COMCHECK PDF PLACER - pyRevit Script
# Place Comcheck PDF pages on Revit sheets in a 3x2 grid

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

ito_type = clr.GetClrType(ImageTypeOptions)
ctor = ito_type.GetConstructor(
    Array[System.Type]([
        clr.GetClrType(System.String),
        clr.GetClrType(System.Boolean),
        clr.GetClrType(ImageTypeSource)
    ])
)

# 1. User picks PDF
pdf_path = forms.pick_file(file_ext='pdf', title='Select Comcheck PDF')
if not pdf_path:
    script.exit()

# 2. User enters total page count
page_count = forms.ask_for_string(
    prompt='How many pages is your Comcheck PDF?',
    title='Page Count',
    default='6'
)
if not page_count:
    script.exit()
page_count = int(page_count)

# 3. Ask for sheet number prefix
sheet_prefix = forms.ask_for_string(
    prompt='Enter sheet number prefix (e.g. M, E, P)',
    title='Sheet Prefix',
    default='M'
)
if not sheet_prefix:
    script.exit()

# 4. Ask for starting sheet number
sheet_start = forms.ask_for_string(
    prompt='Enter starting sheet number (e.g. 5 will create M005, M006...)',
    title='Starting Sheet Number',
    default='5'
)
if not sheet_start:
    script.exit()
sheet_start = int(sheet_start)

# 5. Ask for sheet name
sheet_name = forms.ask_for_string(
    prompt='Enter sheet name (e.g. COMCHECK, ENERGY COMPLIANCE)',
    title='Sheet Name',
    default='COMCHECK'
)
if not sheet_name:
    script.exit()

# 6. Titleblock picker
tb_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()
tb_types = list(tb_collector)
if not tb_types:
    forms.alert("No titleblock types found in project.", exitscript=True)

tb_dict = {}
for tb in tb_types:
    family_name = tb.Family.Name
    type_name = tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    display_name = "{} : {}".format(family_name, type_name)
    tb_dict[display_name] = tb

selected_tb_name = forms.SelectFromList.show(
    sorted(tb_dict.keys()),
    title='Select Titleblock',
    prompt='Choose a titleblock for the Comcheck sheets:',
    multiselect=False
)
if not selected_tb_name:
    script.exit()

selected_tb = tb_dict[selected_tb_name]
tb_id = selected_tb.Id

# 7. Auto detect sheet size and set layout accordingly
# Read width and height from titleblock type parameters (values are in feet)
tb_width_param  = selected_tb.get_Parameter(BuiltInParameter.SHEET_WIDTH)
tb_height_param = selected_tb.get_Parameter(BuiltInParameter.SHEET_HEIGHT)

# WARNING: if these parameters return None your titleblock may store
# dimensions differently - in that case we fall back to 24x36 defaults
if tb_width_param and tb_height_param:
    sheet_w = tb_width_param.AsDouble()   # in feet
    sheet_h = tb_height_param.AsDouble()  # in feet
else:
    # WARNING: fallback to 24x36 if parameters not found
    sheet_w = 3.0   # 36 inches
    sheet_h = 2.0   # 24 inches
    forms.alert(
        "Could not read sheet size from titleblock. Defaulting to 24x36.",
        title="Sheet Size Warning"
    )

PAGES_PER_SHEET = 6
COLS = 3
ROWS = 2

# Margins and gaps in feet
MARGIN_LEFT   = 0.05
MARGIN_TOP    = 0.20
MARGIN_RIGHT  = 0.75   # right side reserved for titleblock border
MARGIN_BOTTOM = 0.30   # bottom reserved for titleblock info strip
GAP_COL       = 0.06   # gap between columns
GAP_ROW       = 0.08   # gap between rows

# Calculate available space and auto size each cell
available_w = sheet_w - MARGIN_LEFT - MARGIN_RIGHT - (GAP_COL * (COLS - 1))
available_h = sheet_h - MARGIN_TOP - MARGIN_BOTTOM - (GAP_ROW * (ROWS - 1))

CELL_W = available_w / COLS
CELL_H = available_h / ROWS

# Origin is top left of the grid
SHEET_ORIGIN_X = MARGIN_LEFT
SHEET_ORIGIN_Y = sheet_h - MARGIN_TOP

# 8. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 9. Create Sheets and Place Pages
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        sheet = ViewSheet.Create(doc, tb_id)

        sheet_number = "{}{}".format(
            sheet_prefix,
            str(sheet_start + sheet_idx).zfill(3)
        )
        sheet.SheetNumber = sheet_number
        sheet.Name = sheet_name

        comments_param = sheet.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if comments_param:
            comments_param.Set("MECHANICAL")

        start_page = sheet_idx * PAGES_PER_SHEET
        end_page = min(start_page + PAGES_PER_SHEET, page_count)

        for i, page_num in enumerate(range(start_page, end_page)):
            col = i % COLS
            row = i // COLS

            x = SHEET_ORIGIN_X + col * (CELL_W + GAP_COL)
            y = SHEET_ORIGIN_Y - row * (CELL_H + GAP_ROW)
            origin = XYZ(x, y, 0)

            img_opts = ctor.Invoke(
                Array[System.Object]([pdf_path, False, ImageTypeSource.Import])
            )
            img_opts.PageNumber = page_num + 1
            img_opts.Resolution = 150

            img_type = ImageType.Create(doc, img_opts)

            place_opts = ImagePlacementOptions()
            place_opts.PlacementPoint = BoxPlacement.TopLeft
            place_opts.Location = origin

            ImageInstance.Create(doc, sheet, img_type.Id, place_opts)

forms.alert(
    "Done! {} sheet(s) created: {}{} to {}{}\nSheet size detected: {:.0f} x {:.0f} inches".format(
        num_sheets,
        sheet_prefix, str(sheet_start).zfill(3),
        sheet_prefix, str(sheet_start + num_sheets - 1).zfill(3),
        sheet_w * 12, sheet_h * 12
    ),
    title="Comcheck Importer"
)