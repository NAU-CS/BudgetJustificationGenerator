# Budget Justification Generator - User Guide

## For Non-Technical Users

This guide will help you use the Budget Justification Generator application to automatically create budget justification documents from your NAU Excel budget templates.

---

## Installation (One-Time Setup)

### Step 1: Install the Application

1. **Download** the file: `BudgetJustificationGenerator-macOS-v1.0.dmg`
2. **Double-click** the DMG file to open it
3. **Drag** the "Budget Justification Generator" app to the Applications folder
4. **Eject** the DMG by clicking the eject button in Finder

![Installation](Installation.png)

### Step 2: First Launch (macOS Security)

The first time you open the app, macOS may show a security warning.

**If you see "cannot be opened because the developer cannot be verified":**

1. **Don't click "Move to Trash"**
2. Close the dialog
3. Go to **Applications** folder
4. **Right-click** (or Control-click) on "Budget Justification Generator"
5. Select **"Open"**
6. Click **"Open"** in the new dialog

This is a one-time step. The app will open normally from now on.

### Step 3: Install Required Software (Optional but Recommended)

The app needs two additional programs to work fully:

#### **Pandoc** (for creating Word documents)
1. Visit: https://github.com/jgm/pandoc/releases/latest
2. Download the file that ends with `-macOS.pkg`
3. Double-click to install
4. Follow the installation wizard

#### **MacTeX** (for creating PDFs)
1. Visit: https://www.tug.org/mactex/
2. Download `MacTeX.pkg` (**Note:** Large file, ~4GB)
3. Double-click to install
4. Follow the installation wizard (**Takes 10-15 minutes**)

**Don't have these installed?** The app will still work and create `.tex` files, but won't automatically create Word documents.

---

## Using the Application

### Step 1: Open the Application

1. Go to your **Applications** folder
2. Double-click **"Budget Justification Generator"**

### Step 2: Check System Status

When the app opens, look at the **System Requirements** section:

- ✅ **Green checkmark** = Software is installed and working
- ⚠️ **Orange warning** = Software not found (optional features won't work)

### Step 3: Select Your Excel File

1. Click the **"Browse..."** button next to "Select Excel Budget File"
2. Navigate to your NAU budget Excel file (usually ends in `.xlsx`)
3. Click **"Open"**

The file name will appear in the text box.

### Step 4: Choose Output Location (Optional)

By default, files are saved to your Desktop.

To change this:
1. Click **"Change..."** next to "Select Output Location"
2. Choose a different folder
3. Click **"Select"**

### Step 5: Generate Documents

1. Click the big green **"Generate Budget Justification"** button
2. Wait for the progress bar to complete (usually 5-10 seconds)
3. A success message will appear

### Step 6: View Your Files

When generation is complete, you'll see a dialog asking:
**"Would you like to open the output folder?"**

Click **"Yes"** to see your generated files:

- **`YourBudget_BudgetJustification.tex`** - LaTeX source file
- **`YourBudget_BudgetJustification.docx`** - Word document (if Pandoc is installed)

**The Word document (.docx) is what you'll use most often!**

---

## Understanding the Generated Files

### Word Document (.docx)
- **Purpose**: Ready to edit and customize
- **Formatting**:
  - Arial 10pt font
  - 0.5" margins
  - Justified text
  - TODO items highlighted in yellow
- **What to do**:
  1. Open in Microsoft Word or Google Docs
  2. Find yellow-highlighted TODO sections
  3. Replace with your specific details
  4. Save and submit

### LaTeX File (.tex)
- **Purpose**: Source file for generating PDFs
- **What to do**: If you need a PDF:
  1. Open Terminal
  2. Navigate to the folder with the .tex file
  3. Run: `xelatex filename.tex` (twice)

---

## Troubleshooting

### Problem: "Please select an Excel budget file"
**Solution**: You forgot to select an Excel file. Click "Browse..." first.

### Problem: No .docx file was created
**Cause**: Pandoc is not installed
**Solution**: Install Pandoc using the instructions in Step 3 above.

### Problem: Can't compile PDF from .tex file
**Cause**: XeLaTeX is not installed
**Solution**: Install MacTeX using the instructions in Step 3 above.

### Problem: "Processing budget data..." stays forever
**Possible causes**:
- The Excel file is corrupted
- The Excel file is not a valid NAU budget template
- The Excel file is open in another program

**Solutions**:
1. Close Excel if it's open
2. Try a different Excel file
3. Verify you're using an NAU budget template

### Problem: Generated document has errors
**Check**:
- Is this a valid NAU budget Excel template?
- Are there any blank or missing values in the Excel file?
- Try running with a known working Excel file first

---

## Tips for Best Results

### Before Running the Generator

1. ✅ **Complete your Excel budget** - Fill in all required fields
2. ✅ **Close Excel** - Make sure the file isn't open in Excel
3. ✅ **Use the NAU template** - Don't use custom Excel files
4. ✅ **Save your work** - Save the Excel file before processing

### After Generation

1. **Always check TODO items** - These are marked with yellow highlighting
2. **Review all numbers** - Verify budgets match your Excel file
3. **Customize descriptions** - Add project-specific details
4. **Save your edits** - Keep a copy of the customized version

### Working with Multiple Budgets

- **Same project, different versions**: Use the output directory to organize by date
- **Multiple projects**: Create a folder for each project
- **Backup important files**: Keep copies of both Excel and final Word documents

---

## Getting Help

### For Application Issues
- **Error messages**: Take a screenshot and note what you were doing
- **File generation problems**: Verify your Excel file works with other NAU tools

### For Budget Content Questions
- **Contact**: NAU Office of Sponsored Projects
- **Template issues**: Verify you have the latest NAU budget template
- **Budget requirements**: Consult with your grants officer

### For Software Installation Help
- **Pandoc**: https://pandoc.org/installing.html
- **MacTeX**: https://www.tug.org/mactex/
- **macOS security**: https://support.apple.com/guide/mac-help/

---

## Quick Reference

| Task | Steps |
|------|-------|
| **Generate documents** | 1. Open app<br>2. Click Browse<br>3. Select Excel file<br>4. Click Generate |
| **Find generated files** | Check Desktop (or chosen output folder) |
| **Edit Word doc** | Open .docx in Word, find yellow TODO items, edit |
| **Create PDF** | Open Terminal, run `xelatex filename.tex` twice |
| **Change output folder** | Click "Change..." before generating |

---

## File Naming Convention

Generated files follow this pattern:
- **Input**: `MyProjectBudget.xlsx`
- **Output**:
  - `MyProjectBudget_BudgetJustification.tex`
  - `MyProjectBudget_BudgetJustification.docx`

---

## System Requirements

- **macOS**: 10.13 (High Sierra) or later
- **Disk Space**: 50MB for the app, plus ~4GB for MacTeX (optional)
- **Memory**: No special requirements
- **Internet**: Only needed for downloading Pandoc/MacTeX

---

**Version**: 1.0
**Last Updated**: January 2026
**For**: Northern Arizona University Office of Sponsored Projects
