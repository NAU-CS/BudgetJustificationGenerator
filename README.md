# Budget Justification Generator

Automatically generates budget justification documents from NAU Excel budget templates in both LaTeX and Microsoft Word (.docx) formats.

## Download

**[Download the latest release](https://github.com/NAU-CS/BudgetJustificationGenerator/releases/latest)**

| Platform | Download | Notes |
|----------|----------|-------|
| **macOS** | `BudgetJustificationGenerator-macOS.zip` | Signed and notarized, runs immediately |
| **Windows** | `BudgetJustificationGenerator-Windows.zip` | Click "More info" → "Run anyway" on first launch |

## Installation

### macOS

1. Download `BudgetJustificationGenerator-macOS.zip` from the [latest release](https://github.com/NAU-CS/BudgetJustificationGenerator/releases/latest)
2. Unzip the file
3. Drag **Budget Justification Generator.app** to your Applications folder
4. Double-click to run

The app is signed and notarized by Apple, so it will open without any security warnings.

### Windows

1. Download `BudgetJustificationGenerator-Windows.zip` from the [latest release](https://github.com/NAU-CS/BudgetJustificationGenerator/releases/latest)
2. Extract the zip file to a folder (e.g., `C:\Program Files\Budget Justification Generator`)
3. Double-click **Budget Justification Generator.exe** to run
4. On first launch, Windows SmartScreen may appear:
   - Click **"More info"**
   - Click **"Run anyway"**
   - This only happens once
  
## Budget Requirements

1. Do not modify the template structure (e.g., removing rows, renaming sheets)
2. Use the "Travel Calculator" sheet if you want travel tables to be generated

## Usage

1. **Select Excel File**: Click "Browse" to select your NAU budget Excel file
2. **Select Output Location**: Choose where to save the generated files (defaults to Desktop)
3. **Generate**: Click "Generate Budget Justification"
4. **Done**: The app creates both `.tex` (LaTeX) and `.docx` (Word) files

### Try It Out

An example budget file (`ExampleBudget.xlsx`) is included in this repository for you to test the application.

### NAU Budget Templates

Download official NAU budget templates from the [Office of Sponsored Projects](https://nau.edu/osp/proposal-support/budget-preparation/).

## Output Files

The generator creates two files:

| Format | Use Case |
|--------|----------|
| **LaTeX (.tex)** | Upload to [Overleaf](https://www.overleaf.com) for professional PDF output |
| **Word (.docx)** | Edit directly in Microsoft Word or Google Docs |

Both files contain the same content with proper formatting, TODO highlights for sections requiring customization, and all standard budget justification sections (A-J).

## Features

- Generates professional budget justifications from NAU Excel templates
- Supports 3, 5, and 10-year project budgets
- Outputs both LaTeX (.tex) and Word (.docx) formats
- Automatic fringe benefit calculations with individual ERE rates
- Smart formatting with year ranges (e.g., "2.0 in Years 2-5")
- Professional tables and structured sections
- TODO items highlighted in red (PDF) and yellow (Word) for easy identification
- Bundled Pandoc - no additional software installation required

## Generated Sections

The output includes all standard NIH budget justification sections:

- **A. Senior Personnel** - PI, Co-PIs, Senior Personnel
- **B. Other Personnel** - GRAs, Postdocs, Students, Staff
- **C. Fringe Benefits** - Individual ERE rates and totals
- **D. Equipment** - Equipment over $5,000
- **E. Travel** - Domestic and international travel with tables
- **F. Participant Support Costs** - Trainee support
- **G. Other Direct Costs** - Materials, consultants, publication costs, tuition, subawards
- **H. Total Direct Costs** - Sum of A-G
- **I. Indirect Costs** - F&A costs with MTDC calculation
- **J. Total Project Costs** - Overall budget total

## Command Line Usage (Advanced)

For automation or scripting, you can use the Python script directly:

```bash
# Install dependencies
pip3 install openpyxl python-docx

# Generate budget justification
python3 generate_budget_justification.py MyBudget.xlsx

# With options
python3 generate_budget_justification.py MyBudget.xlsx -o output_folder -v
```

**Note:** Command line usage requires Python 3.6+ and Pandoc for DOCX generation.

## Compiling LaTeX to PDF

### Option 1: Overleaf (Recommended)

1. Go to [Overleaf](https://www.overleaf.com) and create an account
2. Create a new project → Upload Project
3. Upload the generated `.tex` file
4. Click "Recompile" to generate PDF

### Option 2: Local Compilation

If you have LaTeX installed locally:

```bash
xelatex MyBudget_BudgetJustification.tex
xelatex MyBudget_BudgetJustification.tex
```
*(Run twice to resolve cross-references)*

## Troubleshooting

### macOS: "App is damaged" or won't open
This shouldn't happen with the signed release. If it does:
```bash
xattr -cr "/Applications/Budget Justification Generator.app"
```

### Windows: SmartScreen blocks the app
Click "More info" → "Run anyway". This is normal for new applications without an expensive code signing certificate.

### Excel file not recognized
Ensure your Excel file uses the standard NAU budget template structure with these sheets:
- Budget Details
- Summary_of_Personnel Costs
- Cumulative
- Travel Calculator
- Rates

## Building from Source

For developers who want to build the application themselves:

```bash
# Clone the repository
git clone https://github.com/NAU-CS/BudgetJustificationGenerator.git
cd BudgetJustificationGenerator

# Install dependencies
pip install -r requirements.txt

# Build with PyInstaller
python -m PyInstaller budget_justification_gui.spec
```

## Support

For questions about budget preparation, contact the [NAU Office of Sponsored Projects](https://nau.edu/osp/).

For issues with this application, [open an issue](https://github.com/NAU-CS/BudgetJustificationGenerator/issues) on GitHub.

## License

MIT License - Created by Jared Duval
