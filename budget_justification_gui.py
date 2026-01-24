#!/usr/bin/env python3
"""
GUI for Budget Justification Generator
Provides a user-friendly interface for generating budget justifications
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import shutil
from pathlib import Path

# Import the main script functionality
from generate_budget_justification import BudgetExtractor, LaTeXGenerator


def get_bundled_path(filename):
    """Get path to bundled file (works both in development and PyInstaller bundle)"""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in development
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


def get_system_arch():
    """Get system architecture (arm64 or x86_64)"""
    import platform
    machine = platform.machine().lower()
    if machine in ('arm64', 'aarch64'):
        return 'arm64'
    else:
        return 'x86_64'


def get_pandoc_path():
    """Get path to Pandoc executable (bundled for current platform/architecture, or system)"""
    import platform

    # Check platform
    is_windows = platform.system() == 'Windows'

    if is_windows:
        # Windows: look for pandoc.exe
        bundled = get_bundled_path('pandoc.exe')
        if os.path.exists(bundled):
            return bundled
    else:
        # macOS: look for architecture-specific binary
        arch = get_system_arch()
        bundled = get_bundled_path(f'pandoc-{arch}')
        if os.path.exists(bundled) and os.access(bundled, os.X_OK):
            return bundled

    # Try generic bundled Pandoc (for backwards compatibility)
    bundled_generic = get_bundled_path('pandoc.exe' if is_windows else 'pandoc')
    if os.path.exists(bundled_generic):
        if is_windows or os.access(bundled_generic, os.X_OK):
            return bundled_generic

    # Fall back to system Pandoc
    system_pandoc = shutil.which('pandoc')
    return system_pandoc


class BudgetJustificationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Budget Justification Generator")
        self.root.geometry("750x580")
        self.root.resizable(False, False)

        # Variables
        self.excel_file = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.status_text = tk.StringVar(value="Ready to generate budget justification")
        self.processing = False

        # Template filenames
        self.templates = {
            3: "OSP-3-Year-Budget-Template-August-2025.xlsx",
            5: "OSP-5-Year-Budget-Template-August-2025.xlsx",
            10: "OSP-10-Year-Budget-Template-August-2025.xlsx"
        }

        self.create_widgets()

    def create_widgets(self):
        # Configure ttk styles for cross-platform button appearance
        style = ttk.Style()
        style.configure("Blue.TButton", font=("Arial", 10), padding=(12, 5))
        style.configure("Green.TButton", font=("Arial", 14, "bold"), padding=(30, 12))

        # Header
        header = tk.Frame(self.root, bg="#2C5F87", height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title = tk.Label(header, text="NAU Budget Justification Generator",
                        font=("Arial", 20, "bold"), bg="#2C5F87", fg="white")
        title.pack(pady=20)

        # Main content
        content = tk.Frame(self.root, padx=30, pady=15)
        content.pack(fill=tk.BOTH, expand=True)

        # New Budget section
        new_budget_frame = tk.LabelFrame(content, text="Start a New Budget",
                                         font=("Arial", 12, "bold"), padx=15, pady=10)
        new_budget_frame.pack(fill=tk.X, pady=(0, 12))

        # Container for buttons to center them
        btn_container = tk.Frame(new_budget_frame)
        btn_container.pack()

        btn_3yr = ttk.Button(btn_container, text="New 3-Year Budget",
                            command=lambda: self.create_new_budget(3),
                            style="Blue.TButton")
        btn_3yr.pack(side=tk.LEFT, padx=5)

        btn_5yr = ttk.Button(btn_container, text="New 5-Year Budget",
                            command=lambda: self.create_new_budget(5),
                            style="Blue.TButton")
        btn_5yr.pack(side=tk.LEFT, padx=5)

        btn_10yr = ttk.Button(btn_container, text="New 10-Year Budget",
                             command=lambda: self.create_new_budget(10),
                             style="Blue.TButton")
        btn_10yr.pack(side=tk.LEFT, padx=5)

        # Separator
        separator = ttk.Separator(content, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)

        # Excel file selection
        file_frame = tk.LabelFrame(content, text="1. Select Excel Budget File",
                                   font=("Arial", 12, "bold"), padx=15, pady=12)
        file_frame.pack(fill=tk.X, pady=(0, 12))

        excel_entry = tk.Entry(file_frame, textvariable=self.excel_file,
                              font=("Arial", 11), state='readonly', width=50)
        excel_entry.pack(side=tk.LEFT, padx=(0, 10))

        browse_btn = tk.Button(file_frame, text="Browse...",
                              command=self.browse_excel,
                              font=("Arial", 11), bg="#2C5F87", fg="white",
                              activebackground="#1E4A6E", activeforeground="white",
                              highlightbackground="#2C5F87",
                              padx=15, cursor="hand2")
        browse_btn.pack(side=tk.LEFT)

        # Output directory selection
        output_frame = tk.LabelFrame(content, text="2. Select Output Location",
                                     font=("Arial", 12, "bold"), padx=15, pady=12)
        output_frame.pack(fill=tk.X, pady=(0, 12))

        output_entry = tk.Entry(output_frame, textvariable=self.output_dir,
                               font=("Arial", 11), state='readonly', width=50)
        output_entry.pack(side=tk.LEFT, padx=(0, 10))

        output_btn = tk.Button(output_frame, text="Change...",
                              command=self.browse_output,
                              font=("Arial", 11), bg="#2C5F87", fg="white",
                              activebackground="#1E4A6E", activeforeground="white",
                              highlightbackground="#2C5F87",
                              padx=15, cursor="hand2")
        output_btn.pack(side=tk.LEFT)

        # Generate button
        self.generate_btn = tk.Button(content, text="Generate Budget Justification",
                                      command=self.generate,
                                      font=("Arial", 14, "bold"), bg="#1E7A46", fg="white",
                                      activebackground="#165C35", activeforeground="white",
                                      highlightbackground="#1E7A46",
                                      padx=30, pady=12, cursor="hand2")
        self.generate_btn.pack(pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(content, mode='indeterminate', length=650)

        # Status message
        status_label = tk.Label(content, textvariable=self.status_text,
                               font=("Arial", 10), fg="#555", wraplength=650)
        status_label.pack()

        # Footer
        footer = tk.Label(self.root, text="Created By Jared Duval",
                         font=("Arial", 9), fg="#666", bg="#f0f0f0", pady=10)
        footer.pack(side=tk.BOTTOM, fill=tk.X)

    def create_new_budget(self, years):
        """Create a new budget file from template"""
        template_name = self.templates.get(years)
        if not template_name:
            messagebox.showerror("Error", f"No template available for {years}-year budget")
            return

        template_path = get_bundled_path(template_name)

        if not os.path.exists(template_path):
            messagebox.showerror("Error", f"Template file not found: {template_name}\n\nPlease ensure the template is bundled with the application.")
            return

        # Ask user where to save the new budget
        save_path = filedialog.asksaveasfilename(
            title=f"Save New {years}-Year Budget As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"My_{years}_Year_Budget.xlsx",
            initialdir=str(Path.home() / "Desktop")
        )

        if save_path:
            try:
                shutil.copy2(template_path, save_path)
                self.excel_file.set(save_path)
                self.status_text.set(f"Created new {years}-year budget: {os.path.basename(save_path)}")
                messagebox.showinfo("Success", f"New {years}-year budget created!\n\nFile: {os.path.basename(save_path)}\n\nOpen this file in Excel to fill in your budget details, then return here to generate the justification.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create budget file:\n\n{e}")

    def browse_excel(self):
        """Open file dialog to select Excel file"""
        filename = filedialog.askopenfilename(
            title="Select NAU Budget Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.status_text.set(f"Selected: {os.path.basename(filename)}")

    def browse_output(self):
        """Open dialog to select output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir.get()
        )
        if directory:
            self.output_dir.set(directory)

    def generate(self):
        """Generate budget justification"""
        if self.processing:
            return

        # Validate inputs
        if not self.excel_file.get():
            messagebox.showerror("Error", "Please select an Excel budget file")
            return

        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("Error", "Selected Excel file does not exist")
            return

        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Error", "Output directory does not exist")
            return

        # Start generation in background thread
        self.processing = True
        self.generate_btn.config(state=tk.DISABLED, bg="#cccccc", fg="#666666")
        self.progress.pack(pady=(0, 10))
        self.progress.start(10)
        self.status_text.set("Processing budget data...")

        thread = threading.Thread(target=self.generate_thread)
        thread.daemon = True
        thread.start()

    def generate_thread(self):
        """Background thread for generation"""
        try:
            excel_path = self.excel_file.get()
            output_path = self.output_dir.get()

            # Extract data
            self.root.after(0, self.status_text.set, "Extracting budget data from Excel...")
            extractor = BudgetExtractor(excel_path)
            extractor.extract_senior_personnel()
            extractor.extract_other_personnel()
            extractor.extract_travel()
            extractor.extract_cumulative()
            extractor.extract_subaward_names()
            extractor.extract_other_direct_costs_items()

            # Generate LaTeX
            self.root.after(0, self.status_text.set, "Generating LaTeX document...")
            generator = LaTeXGenerator(extractor)

            # Save files
            basename = os.path.splitext(os.path.basename(excel_path))[0]
            tex_path = os.path.join(output_path, f"{basename}_BudgetJustification.tex")

            generator.save_standalone(tex_path)

            files_generated = [tex_path]

            # Generate DOCX if Pandoc is available
            self.root.after(0, self.status_text.set, "Generating Word document...")
            docx_path = os.path.join(output_path, f"{basename}_BudgetJustification.docx")

            try:
                import subprocess
                from docx import Document
                from docx.shared import Inches, Pt, RGBColor
                from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
                from generate_budget_justification import format_docx_file

                pandoc_path = get_pandoc_path()
                if pandoc_path:
                    pandoc_cmd = [pandoc_path, tex_path, '-f', 'latex', '-t', 'docx',
                                 '-o', docx_path, '--standalone']
                    subprocess.run(pandoc_cmd, check=True, capture_output=True)

                    # Format the DOCX
                    if format_docx_file(docx_path):
                        files_generated.append(docx_path)
            except:
                pass  # Silently skip DOCX if Pandoc not available

            # Success!
            self.root.after(0, self.show_success, files_generated, basename)

        except Exception as e:
            self.root.after(0, self.show_error, str(e))

    def show_success(self, files, basename):
        """Show success message"""
        self.progress.stop()
        self.progress.pack_forget()
        self.generate_btn.config(state=tk.NORMAL, bg="#1E7A46", fg="white")
        self.processing = False

        files_list = "\n".join([f"• {os.path.basename(f)}" for f in files])
        message = f"Successfully generated:\n\n{files_list}\n\nLocation: {self.output_dir.get()}"

        if len(files) == 1:
            message += "\n\nTip: Upload the .tex file to Overleaf to compile to PDF, or install Pandoc to also generate a Word document."

        result = messagebox.showinfo("Success", message)
        self.status_text.set("Generation complete!")

        # Ask if user wants to open the output folder
        if messagebox.askyesno("Open Folder", "Would you like to open the output folder?"):
            import subprocess
            import platform
            if platform.system() == 'Windows':
                os.startfile(self.output_dir.get())
            else:
                subprocess.run(['open', self.output_dir.get()])

    def show_error(self, error_msg):
        """Show error message"""
        self.progress.stop()
        self.progress.pack_forget()
        self.generate_btn.config(state=tk.NORMAL, bg="#1E7A46", fg="white")
        self.processing = False

        messagebox.showerror("Error", f"Failed to generate budget justification:\n\n{error_msg}")
        self.status_text.set("Error occurred. Please try again.")


def main():
    root = tk.Tk()
    app = BudgetJustificationGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()
