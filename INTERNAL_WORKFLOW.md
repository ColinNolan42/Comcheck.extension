# COMCHECK PLACEMENT TOOL — WORKFLOW

---

**RAMIREZ, JOHNSON, AND ASSOCIATES, LLC**

3301 Lawrence St, Suite 2
Denver, CO 80205

Phone: 720.598.0774

---

## WHEN PLACING COMCHECK SHEETS:

### Step 1: Prepare Your Revit Model

Ensure your project file is open and contains at least one titleblock type. The tool will only proceed if titleblocks are available. If you've added new titleblock families to the project, close and reopen the file to refresh the titleblock list.

[SCREENSHOT: Open Revit model with a project file active]

### Step 2: Launch the Comcheck Tool

Navigate to the **RJA Tools** tab in the Revit ribbon, click the **Sheets** panel, and click the **Place Comcheck** button.

[SCREENSHOT: Ribbon showing RJA Tools > Sheets > Place Comcheck button]

### Step 3: Select the PDF File

The file browser will open. Navigate to your COMcheck PDF file and click **Open**. The tool will read the PDF to auto-detect the page count and page size (usually 8.5×11" for standard COMcheck reports).

[SCREENSHOT: Windows file browser with COMcheck PDF selected]

### Step 4: Complete the Placement Dialog

After selecting the PDF, the Comcheck Sheet Placement dialog appears. Fill in the required fields:

**Sheet Prefix** — Enter the letter prefix for your sheets. Examples:
- "M" for Mechanical (typical)
- "E" for Electrical
- "P" for Plumbing

Default value: "M"

**Sheet Number** — Enter the starting sheet number. The tool recognizes several formats:
- "005" → generates M005, M006, M007, etc.
- "0.4" → generates M0.4, M0.5, M0.6, etc. (preserves the decimal format)
- "04" → generates M04, M05, M06, etc. (preserves padding)

The last numeric portion is auto-incremented across sheets. Everything before it is preserved.

**Titleblock** — Click the dropdown to select a titleblock family and type from your project. You can see both the family name and type name in the list. Once selected, all generated sheets will use this titleblock.

[SCREENSHOT: Dialog showing Sheet Prefix, Sheet Number, and Titleblock fields filled in]

**Sheet Size** — Choose from two preset options:
- "24 x 36" — Standard sheet size with optimized default layout
- "30 x 42" — Large sheet size with defaults tuned for the larger area

The tool will automatically populate the Advanced section with spacing, margin, and resolution defaults for your chosen sheet size when you select it or when the form first loads.

[SCREENSHOT: Dialog showing Sheet Size dropdown with 24 x 36 and 30 x 42 options]

### Step 5: Configure Advanced Options (Optional)

Expand the **Advanced (optional overrides)** section only if you need to adjust layout parameters. All fields are optional — if left blank, the tool uses the default values for your selected sheet size.

**To Customize Spacing:**

- **Margin Left** — Distance in feet from the left edge of the sheet to the first column. Example: 0.20 ft (about 2.4 inches)
- **Margin Top** — Distance in feet from the top edge of the sheet to the first row
- **Margin Right** — Distance in feet from the right edge of the sheet (reserves space for title block trim)
- **Margin Bottom** — Distance in feet from the bottom edge of the sheet

**To Adjust Grid Gaps:**

- **Gap Col** — Horizontal spacing (in feet) between images in the grid. Example: 0.05 ft (about 0.6 inches)
- **Gap Row** — Vertical spacing (in feet) between rows of images

**To Change Grid Structure:**

- **Columns** — Number of columns in the grid. Default: 4
- **Rows** — Number of rows in the grid. Default: 2 (total 8 images per sheet)

**To Control Image Quality and Fit:**

- **Resolution (DPI)** — PDF rasterization quality. Default: 600 DPI. Higher values (e.g., 1200) produce sharper images but increase file size and processing time.
- **Image Scale** — Multiplier for image size within each grid cell:
  - 1.0 = fill the entire cell (default behavior)
  - 0.9 = 90% of cell size (leaves small border)
  - 0.8 = 80% of cell size (larger border)
  - Use smaller values if images are overlapping cell boundaries

**To Override Page Count (Edge Cases Only):**

- **Page Count Override** — If the tool cannot auto-detect your PDF's page count, enter the number manually here and rerun. Leave blank for normal operation (auto-detection).

**Important**: When you change the **Sheet Size** dropdown, all Advanced fields automatically repopulate with the defaults for that size. Your overrides are not saved, so reselect the sheet size only when you intend to reset to those defaults.

[SCREENSHOT: Dialog with Advanced section expanded, showing all parameter fields]

### Step 6: Review and Place

Review your selections one more time:
- Sheet prefix and starting number
- Titleblock selection
- Sheet size

If any existing sheets in your project already use the sheet numbers you're about to create, the tool will alert you and ask you to choose different sheet numbers. Adjust the Sheet Number field and rerun if this happens.

Click **Place Comcheck Sheets** to begin placement.

[SCREENSHOT: Dialog with all fields completed, Place Comcheck Sheets button highlighted]

### Step 7: Completion and Verification

After the tool finishes, an alert will show:
- Number of sheets created (e.g., "Done! 2 sheet(s) created: M005 to M006")
- The sheet size used

Click **OK** to close the alert. New sheets will now be visible in your Project Browser under **Sheets**.

[SCREENSHOT: Completion dialog showing number of sheets created and sheet number range]

Open one of the newly created sheets to verify:
- All PDF pages are placed in a grid layout
- Images are properly sized and aligned
- The titleblock is correct
- The Comments parameter shows "MECHANICAL" (check via Sheet Properties)

[SCREENSHOT: Revit view showing a COMCHECK sheet with 8 PDF pages in a 4×2 grid]

### Step 8: Iterating and Tuning Layout (Optional)

If the layout does not fit your expectations:

1. Note which parameters need adjustment (e.g., images are too large, gaps too small, margins off-center)
2. Delete the generated sheets (right-click sheet in Project Browser → Delete)
3. Rerun the tool with the same PDF and adjusted Advanced parameters
4. Examples:
   - If images touch or overlap: lower the **Image Scale** to 0.9 or 0.8
   - If gaps between images are too large: reduce **Gap Col** or **Gap Row**
   - If the grid is off-center on the sheet: adjust **Margin Left** or **Margin Top**
5. Generate new sheets and verify until layout is correct

Once satisfied, you can delete old test sheets and keep the final version.

---

**Further instructions and updated screenshots will be added as needed.**
