# Comcheck Extension

A pyRevit extension that automates the placement of COMcheck PDF pages as images onto Revit sheets. Each PDF is automatically rasterized and arranged in a configurable grid layout, with support for multiple sheet sizes and customizable spacing.

## Features

- **Automatic PDF Detection**: Detects PDF page count and page size from the file without external libraries
- **Configurable Grid Layout**: Arrange pages in a 4×2 grid (8 pages per sheet) or customize rows, columns, and spacing
- **Multiple Sheet Sizes**: Built-in presets for 24×36" and 30×42" sheets with optimized default layouts
- **Auto-Incrementing Sheet Numbers**: Intelligently increments sheet numbers across multiple generated sheets (e.g., M005, M006, M007...)
- **Advanced Control**: Optional overrides for margins, gaps, resolution (DPI), image scaling, and page count for edge cases
- **Automatic Titleblock Assignment**: Select from all titleblock types in your project; custom titleblocks are supported

## Requirements

- **Revit 2021 or newer**
- **pyRevit** installed and configured

## Installation

### Method 1: Install via pyRevit Extensions Manager (Recommended)

![pyRevit dropdown menu showing the Extensions button](images/00-pyrevit-extensions-menu.png)

1. In Revit, open the **pyRevit** tab and click the **pyRevit** button (top-left of the pyRevit tab, next to Reload) to open the dropdown shown above.
2. Click **Extensions** (the gear icon). The **pyRevit Extension Manager** window opens.

![pyRevit Extension Manager — Git URL and Add and install](images/00b-pyrevit-extension-manager.png)

3. At the bottom of the window, paste the repository URL into the **Git URL** field:
   ```
   https://github.com/ColinNolan42/ComCheck.extension.git
   ```
4. Leave **Path** and **Token** at their defaults, then click **Add and install**.

![Extension installed successfully confirmation](images/00c-install-success.png)

5. A confirmation message appears: *Extension "ComCheck.extension" installed successfully! Revit will reload to apply changes.* Click **OK** and restart Revit if it doesn't reload automatically.
6. The button will appear in the Revit ribbon under **RJA Tools → Sheets → Place Comcheck**.

### Method 2: Manual Copy

1. Clone or download this repository, then copy the `ComCheck.extension` folder to your pyRevit Extensions directory:
   - On Windows: `C:\Users\<username>\AppData\Roaming\pyRevit\Extensions\`
   - On Mac: `~/Library/Application Support/pyRevit/Extensions/`

2. If pyRevit is running, reload extensions (pyRevit → Reload) or restart Revit.

3. The button will appear in the Revit ribbon under **RJA Tools → Sheets → Place Comcheck**.

## Usage

### Basic Workflow

**Steps 1–2: Locate and launch the tool**

![Ribbon location of the Place Comcheck button](images/01-ribbon-placecomcheck.png)

1. **Open a Revit Model** with at least one titleblock type available, and locate the **RJA Tools** tab in the ribbon.
2. **Click the Place Comcheck Button** (Sheets panel) — this immediately opens a file browser.

3. **Select a PDF File** using the file browser dialog. The tool reads the PDF to auto-detect page count and page size.

**Steps 3–7: Fill out the placement dialog**

![Comcheck Sheet Placement dialog — main fields](images/02-dialog-main.png)

4. **Fill Out the Placement Dialog**:
   - **3. Sheet Prefix**: Letter prefix for sheet numbers (e.g., M, E, P) — defaults to "M"
   - **4. Sheet Number**: Starting sheet number (e.g., 005, 0.4, or 04) — the tool preserves the format and auto-increments the last number
   - **5. Titleblock**: Choose from all available titleblocks in your project
   - **6. Sheet Size**: Select 24×36" or 30×42" (other sizes can be configured via Advanced options)
   - **7. Advanced**: Optional overrides — see below. Leave collapsed for normal use.

**Step 8: Generate the sheets**

5. **Click "Place Comcheck Sheets"** to generate sheets with all PDF pages placed.

### Advanced Options (Optional)

![Advanced overrides panel](images/03-advanced-settings.png)

Expand the **Advanced** section to customize layout parameters. All fields are optional — if left blank, defaults for the selected sheet size apply automatically:

- **Margin Left/Top/Right/Bottom** (feet): Space between the sheet edge and the grid
- **Gap Col/Row** (feet): Horizontal and vertical spacing between images in the grid
- **Columns/Rows**: Grid dimensions (default: 4 columns × 2 rows)
- **Resolution (DPI)**: PDF rasterization quality (default: 600 DPI)
- **Image Scale**: Multiplier for image size within each cell (1.0 = fill cell; < 1.0 = smaller images)
- **Page Count Override**: Manually specify PDF page count if auto-detection fails

**Note**: When you change the sheet size dropdown, the Advanced fields automatically repopulate with the default values for that sheet size.

### Handling Errors

**Sheet Number Already Exists**: If you attempt to use sheet numbers that already exist in the project, the tool will alert you and request a different starting sheet number.

**Page Count Auto-Detection Failed**: If the tool cannot read the page count from your PDF, enter a value in the **Page Count Override** field and rerun.

## Output

The tool creates new ViewSheets named **COMCHECK** with the following properties:

- **Sheet Number**: Auto-incremented from your input (e.g., M005, M006, M007...)
- **Sheet Name**: Always "COMCHECK"
- **Titleblock**: The type you selected in the dialog
- **Comments Parameter**: Set to "MECHANICAL" (visible in sheet properties)
- **Images**: All PDF pages placed in the specified grid layout, scaled to fit without distortion

Each page image is placed precisely at calculated XY coordinates within the grid, ensuring consistent, professional layout across all sheets.

## Technical Details

- **Grid Layout**: Calculates cell dimensions based on sheet size, margins, and gaps
- **Image Fitting**: Scales each PDF page to fit its grid cell while preserving aspect ratio
- **Supported Sheet Sizes**: 24×36" and 30×42" (custom sizes supported via Advanced options)
- **Default Resolution**: 600 DPI (adjustable per sheet generation)

## File Structure

```
ComCheck.extension/
├── extension.json
├── README.md
├── INTERNAL_WORKFLOW.md
└── RJA Tools.tab/
    └── Sheets.panel/
        └── PlaceComcheck.pushbutton/
            ├── script.py
            └── icon.png
```
