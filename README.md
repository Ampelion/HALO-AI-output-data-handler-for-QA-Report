# Data Handler for rare cell detction assay (intended for output from HALO AI)

A VBA macro for automated processing, analysis, and quality control reporting of nuclear immunohistochemistry (IHC) assay data from image analysis platforms.

## Overview

This macro transforms raw cellular analysis data into a comprehensive, regulatory-compliant report with automated calculations, quality control metrics, reference range validation, and electronic signature fields. Designed for pluripotent stem cell characterization assays in GMP/GLP environments.

## Key Features

### Automated Analysis
- **Summary Statistics** - Total slides, nuclei counts, and percentage calculations by sample type
- **LLOD Calculations** - Automatic flagging of results below Lower Limit of Detection (<0.0004)
- **Reference Range Validation** - Checks if 0.01% control values fall within established limits
- **Quality Control Metrics** - Validates test nuclei count >1 million requirement

### Sample Classification & Color Coding
Automatically categorizes and color-codes samples:
- **Isotype Control** (light blue) - Negative control samples
- **Test Samples** (dark blue) - Experimental OCT4 test samples  
- **0.01% Control** (light lavender) - Low-level positive control
- **Control Test** (dark lavender) - Control validation samples
- **1% Control** (gray) - High-level positive control

### Regulatory Compliance
- **Pass/Fail Criteria** - Clear YES/NO determinations for:
  - Test nuclei count >1 million
  - 0.01% control within reference range
  - OCT4 test result vs. LLOD
- **Signature Block** - Electronic signature fields for:
  - Performed By (with date)
  - Data Verification (with date)
  - Quality Review (with date)
- **Data Integrity** - Bordered summary tables for formal reporting

### Report Elements
- Individual sample data with percentage calculations
- Subtotal summaries filtered by sample type
- Reference range specifications
- Reportable results with acceptance criteria
- Professional formatting with borders and color coding

## Installation

### Option 1: Personal Macro Workbook (Recommended)
Makes the macro available in all Excel workbooks.

1. Open Excel
2. Press `Alt + F11` to open VBA Editor
3. If you don't see PERSONAL.XLSB:
   - Go back to Excel
   - Click **Developer** → **Record Macro**
   - Set "Store macro in:" to **Personal Macro Workbook**
   - Click OK, then immediately click **Stop Recording**
4. Press `Alt + F11` again to return to VBA Editor
5. Press `Ctrl + R` to show Project Explorer panel
6. Find **VBAProject (PERSONAL.XLSB)** in the left panel
7. Right-click → **Insert** → **Module**
8. Copy the contents of `OCT4DataHandler.bas` and paste into the module
9. Close VBA Editor
10. File → Save (saves to PERSONAL.XLSB automatically)

### Option 2: Individual Workbook
Makes the macro available only in a specific workbook.

1. Open your Excel workbook with OCT4 data
2. Press `Alt + F11` to open VBA Editor
3. Click **Insert** → **Module**
4. Copy the contents of `OCT4DataHandler.bas` and paste into the module
5. Close VBA Editor
6. Save workbook as **Excel Macro-Enabled Workbook (.xlsm)**

## Input Data Format

### Required File Format
The macro expects data exported from TissueStudio or compatible image analysis software with these columns:

**Column Structure (before macro runs):**
- **A**: Image Location (full file path)
- **B**: Image Tag (slide identifier)
- **C**: Algorithm Name (e.g., "OCT4 Nuc")
- **D**: Start Time
- **E**: Job Id
- **F**: Analysis Region (e.g., "Aggregate")
- **G**: Total Cells (total nuclei counted)
- **H**: Htox Positive Cells
- **I**: IHC Positive Cells
- **J**: Negative Cells
- **K**: Htox Positive Nuclei
- **L**: IHC Positive Nuclei (key data for OCT4 positivity)
- **M-AA**: Additional analysis parameters (deleted by macro)

**Sample Naming Convention:**
The macro uses filename patterns to categorize samples:
- Contains "ISO" → Isotype Control
- Contains "OCT4" or "OCPT" → Test Sample
- Contains "_0.01-", "_0_01", or "_0-01" → 0.01% Control
- Contains "test" (in controls) → Control Test
- Contains "_1-" or "_1_" → 1% Control

See `sample_input.csv` for a complete example.

## Usage

### Step-by-Step Workflow

1. **Export data from image analysis software**
   - Use standard cellular analysis export
   - Save as CSV file

2. **Open in Excel**
   - Excel will parse the CSV into columns
   - Data should appear in columns A through AA (or more)

3. **Run the macro**
   - Press `Alt + F8` to open Macro dialog
   - Select `OCT4DataHandler`
   - Click **Run**

4. **Review the generated report**
   - Check summary statistics
   - Verify reference range compliance
   - Review reportable results
   - Add electronic signatures

5. **Save the report**
   - Save as `.xlsx` or `.xlsm` for archival

### What the Macro Produces

**Main Data Table:**
- Column A: Slide name (color-coded by type)
- Column B: Algorithm
- Column C: Analysis ROI
- Column D: Total Nuclei (formatted with thousands separator)
- Column E: IHC Positive nuclei count
- Column F: # Confirmed (for manual review - initially empty)
- Column G: % IHC Positive (calculated: E/D × 100)

**Summary Section (below main table):**
- Overall totals across all slides
- Filtered subtotals by sample type:
  - Isotype Control
  - Test
  - 0.01% Control  
  - Control Test

**Reference Range Specifications:**
- Field for entering control reference range (lower and upper limits)

**Reportable Results:**
- ✓/✗ Test Nuclei > 1M
- ✓/✗ 0.01% Control in Reference Range
- OCT4 Test Result (with LLOD flagging)

**Signature Block:**
- Performed By / Date
- Data Verification / Date
- Quality Review / Date

## Quality Control & Acceptance Criteria

### Test Validity Requirements
1. **Minimum Cell Count**: Test samples must have >1,000,000 nuclei analyzed
2. **Control Performance**: 0.01% control must fall within established reference range
3. **LLOD Threshold**: Results <0.0004 are reported as "<LLOD" (below detection limit)
4. **Secondary Threshold**: Results between 0.0004 and 0.0011 are reported as "<0.0011"

### Reference Range Entry
After running the macro, manually enter the reference range limits:
1. Locate "Reference Range for 0.01% Control" section
2. Enter lower limit value
3. Enter upper limit value
4. The macro automatically validates control against these limits

## Technical Details

### Data Processing Steps
1. Deletes unnecessary columns (A, D, E, H-K, M-AA)
2. Renames column headers for clarity
3. Adds AutoFilter to enable data sorting
4. Calculates % IHC Positive for each sample
5. Creates SUBTOTAL formulas for filtered summaries
6. Applies conditional formatting for color coding
7. Generates filtered subtotals by sample type
8. Creates bordered summary tables
9. Adds signature block with merged cells
10. Auto-fits all columns for readability

### Formula Reference
- **% IHC Positive**: `=F2/D2*100` (IHC Positive ÷ Total Nuclei × 100)
- **Total Slides**: `=SUBTOTAL(103, R2C3:R[-3]C)` (counts non-empty cells)
- **Sum Total Nuclei**: `=SUBTOTAL(109, R2C4:R[-3]C)` (sum of filtered values)
- **Test Result Logic**: `=IF(r[-7]c[5]<0.0004,"<LLOD",IF(r[-7]c[5]<0.0011,"<0.0011",r[-7]c[5]))`
- **Reference Range Check**: `=IF(AND(r[-5]c[5]>r[0]c[3],r[-5]c[5]<r[0]c[4]),"YES","NO")`

### Color Coding (RGB Values)
- Isotype Control: RGB(217, 225, 242) - Light blue
- Test: RGB(180, 198, 231) - Dark blue  
- 0.01% Control: RGB(206, 205, 233) - Light lavender
- Control Test: RGB(183, 181, 221) - Dark lavender
- 1% Control: RGB(166, 166, 166) - Gray

### Number Formatting
- Total Nuclei: `#,##0` (thousands separator, no decimals)
- % IHC Positive: `.0000` (four decimal places)

## Sample Data

This repository includes `sample_input.csv` - actual TissueStudio export data containing:
- 18 OCT4 analysis samples
- Mix of test samples, isotype controls, and 0.01% controls
- Real cellular analysis data from OCPT-CAG71 study
- Representative of typical OCT4 assay output

Use this file to:
- Test the macro functionality
- Understand expected input format  
- See a complete before/after transformation
- Validate installation

## System Requirements

- Microsoft Excel 2010 or later
- Windows or Mac with VBA support
- Macros must be enabled in Excel security settings

## Regulatory & Compliance Notes

### 21 CFR Part 11 Considerations
This macro facilitates electronic record creation for regulated environments:
- **Audit Trail**: Signature block captures operator, verifier, and reviewer with dates
- **Data Integrity**: Automated calculations reduce transcription errors
- **Validation**: Clear pass/fail criteria based on pre-defined acceptance limits

### Best Practices
- Always work on a COPY of raw data files
- Archive original CSV exports separately
- Document reference range values and their justification
- Maintain macro version control
- Validate macro performance during initial deployment

## Limitations & Assumptions

- Designed for TissueStudio OCT4 Nuclear Analysis export format
- Sample type classification based on filename patterns
- Reference range limits must be entered manually
- Assumes specific column arrangement in source data
- Maximum tested dataset: ~100 samples per file

## Troubleshooting

**Problem**: Macro produces errors or columns misaligned  
**Solution**: Verify input file matches expected column structure (see Input Data Format section)

**Problem**: Sample colors not applying correctly  
**Solution**: Check that slide names contain expected patterns (ISO, OCT4, _0.01-, etc.)

**Problem**: Subtotals showing incorrect values  
**Solution**: Ensure AutoFilter is working; try clearing and re-running macro

**Problem**: Reference range validation not working  
**Solution**: Verify lower and upper limit values are entered as numbers, not text

**Problem**: Signature block misaligned  
**Solution**: Re-run macro on fresh data; manual edits may have disrupted cell merging

## Author

Josie Babin
