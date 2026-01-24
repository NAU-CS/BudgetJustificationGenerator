# Budget Justification Generator

Automatically generates budget justification documents from NAU Excel templates in both LaTeX/PDF and Microsoft Word (.docx) formats.

## 🎯 Quick Start for Non-Technical Users

**Looking for a simple, user-friendly application?** Use the **GUI version**:

1. **Download**: Get `BudgetJustificationGenerator-macOS-v1.0.dmg`
2. **Install**: Double-click the DMG and drag the app to Applications
3. **Run**: Double-click "Budget Justification Generator" like any other Mac app
4. **Use**: Click buttons to select your Excel file and generate documents

**No command line or coding required!** The GUI provides:
- 📁 Visual file picker for selecting Excel files
- 📊 Real-time system requirements check
- ✨ Progress indicators
- ✅ Clear success/error messages
- 🚀 One-click generation

**For Technical Users**: Continue reading for command-line usage and advanced options.

---

## Features

- ✅ Generates professional budget justifications from NAU Excel templates
- ✅ Supports 3, 5, and 10-year project budgets
- ✅ Outputs both LaTeX (.tex) and Word (.docx) formats
- ✅ Automatic fringe benefit calculations
- ✅ Smart formatting with year ranges
- ✅ Professional tables and structured sections
- ✅ Includes 3% annual salary escalation text
- ✅ TODO items highlighted in red (PDF) and yellow/red (DOCX) for easy identification

## Requirements

### Python Dependencies
- Python 3.6+
- openpyxl
- python-docx (for .docx formatting)

Install with:
```bash
pip3 install openpyxl python-docx
```

### LaTeX (for PDF generation)
- XeLaTeX (part of TeX Live, MacTeX, or MiKTeX)

**macOS:**
```bash
brew install --cask mactex
```

**Ubuntu/Debian:**
```bash
sudo apt-get install texlive-xetex texlive-latex-extra
```

### Pandoc (for DOCX generation)
Pandoc is required to generate Microsoft Word (.docx) files.

**IMPORTANT:** Pandoc is NOT a Python package and cannot be installed via `pip`. It's a standalone application that must be installed using system package managers or direct download.

**macOS (with Homebrew):**
```bash
brew install pandoc
```

**macOS (without Homebrew - Direct Download):**
1. Download the .pkg installer from https://github.com/jgm/pandoc/releases/latest
2. Look for `pandoc-X.X.X-macOS.pkg` (where X.X.X is the version)
3. Double-click to install

**Ubuntu/Debian:**
```bash
sudo apt-get install pandoc
```

**Windows:**
Download installer from https://pandoc.org/installing.html

**Verify installation:**
```bash
pandoc --version
```

**Note:** The script will work without Pandoc, but will only generate LaTeX/PDF files.

## Usage

### Basic Usage
```bash
python3 generate_budget_justification.py MyBudget.xlsx
```

This generates:
- `MyBudget_BudgetJustification.tex` - LaTeX source file
- `MyBudget_BudgetJustification.docx` - Word document (if Pandoc is installed)

### Compile PDF
```bash
xelatex MyBudget_BudgetJustification.tex
xelatex MyBudget_BudgetJustification.tex
```
*(Run twice to resolve cross-references)*

### Advanced Options
```bash
# Specify output directory
python3 generate_budget_justification.py MyBudget.xlsx -o output_folder

# Verbose output with detailed extraction information
python3 generate_budget_justification.py MyBudget.xlsx -v
```

## Output Formats

### LaTeX/PDF
- Fully formatted with proper typography
- Professional tables and section headers
- Cross-references for travel tables
- Red TODO highlights for required edits

### Word (.docx)
- Automatically generated from LaTeX using Pandoc
- Preserves structure and formatting
- Matches PDF formatting: 0.5" margins, Arial 10pt, justified text
- Yellow-highlighted TODO items with red text for easy identification
- Editable in Microsoft Word or Google Docs
- Ideal for collaboration and final editing

## File Structure

The script expects NAU budget Excel templates with the following sheets:
- **Budget Details** - Senior and other personnel, salary data
- **Summary_of_Personnel Costs** - Fringe calculations
- **Cumulative** - Year-over-year totals
- **Travel Calculator** - Domestic and international travel
- **Rates** - Indirect cost rates

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

## Features

### Smart Formatting
- **Year Ranges**: "2.0 in Years 2-5" instead of "Y2: 2.0, Y3: 2.0, Y4: 2.0, Y5: 2.0"
- **Salary Escalation**: Automatically includes "A 3% annual salary increase is included"
- **Person Months**: Simplified headers when all years are identical

### Accurate Calculations
- **Fringe rates** extracted from Summary sheet for each person
- **Senior Personnel**: Individual ERE rates (35.3%, 15.3%, etc.)
- **Other Personnel**: Position-specific rates (15.1%, 8.1%, etc.)

### Professional Output
- University-mandated verbiage for fringe benefits and indirect costs
- Proper LaTeX formatting with escaped special characters
- TODO highlights for sections requiring customization
- Structured subsections for Other Direct Costs

## Troubleshooting

### Pandoc Not Found
If you see: `⚠ Warning: Could not generate DOCX file. Pandoc may not be installed.`

Install Pandoc using the instructions above. The LaTeX/PDF generation will still work.

### LaTeX Compilation Errors
- Make sure XeLaTeX is installed
- Run compilation **twice** to resolve cross-references
- Check that all special characters are properly escaped

### Excel Template Issues
- Ensure the Excel file uses the standard NAU template structure
- Check that sheet names match exactly (including spaces)
- Verify person months and salary data are in expected columns

## Examples

```bash
# Generate from EdgeCase budget
python3 generate_budget_justification.py EdgeCase.xlsx

# Output to specific folder with verbose logging
python3 generate_budget_justification.py NetGauge.xlsx -o output/ -v

# Compile the resulting PDF
cd output/
xelatex NetGauge_BudgetJustification.tex
xelatex NetGauge_BudgetJustification.tex
```

## Building Standalone Executable

You can package the script as a standalone executable that doesn't require Python to be installed.

### Requirements for Building
- PyInstaller: `pip3 install pyinstaller`

### Quick Build

```bash
./build_executable.sh
```

This creates a single executable file: `dist/budget-justification` (~14MB)

### Manual Build

```bash
# Install PyInstaller
pip3 install pyinstaller

# Build the executable
python3 -m PyInstaller budget_justification.spec --clean
```

### Using the Executable

```bash
# Basic usage
./dist/budget-justification MyBudget.xlsx

# With output directory
./dist/budget-justification MyBudget.xlsx -o output_folder

# Verbose mode
./dist/budget-justification MyBudget.xlsx -v
```

### Distribution

The executable can be distributed to other users **on the same OS** without requiring:
- ✅ Python installation
- ✅ Python packages (openpyxl, python-docx)

**Still Required** on the target system:
- ⚠️ Pandoc (for .docx generation) - [Installation instructions](#pandoc-for-docx-generation)
- ⚠️ XeLaTeX (for PDF compilation) - [Installation instructions](#latex-for-pdf-generation)

**Platform Notes:**
- macOS executable works on macOS only (Intel or Apple Silicon)
- Build separately for Windows/Linux if distribution to those platforms is needed
- Cross-platform: Consider distributing the Python script instead

### What Gets Packaged

The executable includes:
- Complete Python interpreter
- All Python dependencies (openpyxl, python-docx, lxml, etc.)
- The budget justification script

The executable does NOT include:
- Pandoc (system tool, must be installed separately)
- XeLaTeX (system tool, must be installed separately)
- Input Excel files (provided by user)

## Building GUI Application (Recommended for Non-Technical Users)

The GUI application provides a much better experience for non-technical users.

### Quick Build

```bash
./build_gui_app.sh
```

This creates: `dist/Budget Justification Generator.app` (~16MB)

### Create Distributable DMG

```bash
./create_dmg.sh
```

This creates: `BudgetJustificationGenerator-macOS-v1.0.dmg`

**The DMG file is perfect for distribution because:**
- ✅ Users just double-click to open
- ✅ Familiar "drag to Applications" installation
- ✅ Includes a README with instructions
- ✅ Works like any other Mac application
- ✅ No terminal or coding knowledge required

### What Users Get

**Visual Interface:**
- Browse button to select Excel file
- System requirements status (shows if Pandoc/XeLaTeX are installed)
- Progress bar during generation
- Clear error messages
- Success notification with option to open output folder

**Better User Experience:**
- No command line needed
- Visual feedback at every step
- Automatic detection of missing dependencies
- One-click file generation

### GUI vs. Command Line

| Feature | GUI App | Command Line |
|---------|---------|--------------|
| User-friendly | ✅ Very easy | ❌ Technical |
| Installation | Drag & drop | None needed |
| File selection | Click to browse | Type path |
| Error messages | Clear dialogs | Terminal text |
| Best for | Non-technical users | Automation, scripts |

## Notes

- The script automatically detects 3, 5, or 10-year project templates
- All monetary values are formatted with proper currency symbols
- Travel tables include detailed breakdowns by trip
- Section numbering follows NIH standards (A-J)
- Both .tex and .docx files are standalone and self-contained

## Support

For issues or questions, consult the NAU Office of Sponsored Projects or Research Development.
