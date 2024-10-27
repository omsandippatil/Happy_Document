import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import os

class ExcelUtilityApp(tk.Tk):
    def __init__(self , window):
        super().__init__()
        self.title("Excel Utility")
        self.geometry("800x600")

        # Main container
        self.main_container = ttk.Frame(self)
        self.main_container.pack(fill='both', expand=True, padx=10, pady=10)

        # Create a style for the notebook and tabs
        style = ttk.Style()
        style.configure('TabNotebook.TNotebook', background='#6a6a6a')
        style.configure('TabNotebook.TFrame', background='white')
        style.configure('TNotebook.Tab', background='white', borderwidth=1)
        style.map('TNotebook.Tab', background=[('selected', 'black')])
        
        # ** New Style for Control Panel Frame **
        style.configure('Dark.TFrame', background='#333333')

        # Create Notebook for tabs
        self.notebook = ttk.Notebook(self.main_container, style='TabNotebook.TNotebook')
        self.notebook.pack(fill='both', expand=True, pady=(10, 0))

        # Initialize variables
        self.selected_file = ""
        self.df = None

        # Create tabs
        self.create_tabs()

        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, style='Status.TLabel')
        self.status_bar.pack(fill='x', pady=5)
        self.status_var.set("Ready")

    def create_tabs(self):
        # Create and add all tabs
        self.tabs = {
            "Extract Data": self.create_extract_tab,
            "Merge Excel Files": self.create_merge_tab,
            "Generate Report": self.create_report_tab,
            "Validate Data": self.create_validation_tab,
            "Convert Excel to PDF": self.create_convert_tab
        }

        for tab_name, tab_function in self.tabs.items():
            frame = ttk.Frame(self.notebook, padding=(10, 20))
            self.notebook.add(frame, text=tab_name)
            tab_function(frame)

    def create_extract_tab(self, frame):
        ttk.Label(frame, text="Extract Data from Excel", font=('Helvetica', 16)).pack(pady=10)
        ttk.Button(frame, text="Extract All Data", command=self.extract_data).pack(pady=10)
        ttk.Button(frame, text="Load Excel File", command=self.browse_file).pack(pady=10)
        self.file_label = ttk.Label(frame, text="No file selected")
        self.file_label.pack(pady=5)

    def create_merge_tab(self, frame):
        ttk.Label(frame, text="Merge Excel Files", font=('Helvetica', 16)).pack(pady=10)
        ttk.Button(frame, text="Add Files", command=self.add_files).pack(pady=5)
        ttk.Button(frame, text="Merge Files", command=self.merge_files).pack(pady=5)
        self.files_to_merge = []

    def create_report_tab(self, frame):
        ttk.Label(frame, text="Generate Simple Report", font=('Helvetica', 16)).pack(pady=10)
        ttk.Button(frame, text="Load Excel File", command=self.browse_file).pack(pady=5)
        self.report_file_label = ttk.Label(frame, text="No file selected")
        self.report_file_label.pack(pady=5)
        
        ttk.Label(frame, text="Column for Report Metric:").pack(pady=5)
        self.report_column_var = tk.StringVar()
        self.report_column_combo = ttk.Combobox(frame, textvariable=self.report_column_var)
        self.report_column_combo.pack(pady=5)
        ttk.Button(frame, text="Generate Report", command=self.generate_report).pack(pady=10)

    def create_validation_tab(self, frame):
        ttk.Label(frame, text="Data Validation", font=('Helvetica', 16)).pack(pady=10)
        ttk.Button(frame, text="Load Excel File", command=self.browse_file).pack(pady=5)
        self.validation_file_label = ttk.Label(frame, text="No file selected")
        self.validation_file_label.pack(pady=5)
        
        ttk.Label(frame, text="Select Column to Validate:").pack(pady=5)
        self.validation_column_var = tk.StringVar()
        self.validation_column_combo = ttk.Combobox(frame, textvariable=self.validation_column_var)
        self.validation_column_combo.pack(pady=5)
        ttk.Button(frame, text="Validate Data", command=self.validate_data).pack(pady=10)

    def create_convert_tab(self, frame):
        ttk.Label(frame, text="Convert Excel to PDF", font=('Helvetica', 16)).pack(pady=10)
        ttk.Button(frame, text="Load Excel File", command=self.browse_file).pack(pady=5)
        self.convert_file_label = ttk.Label(frame, text="No file selected")
        self.convert_file_label.pack(pady=5)
        
        control_panel_frame = ttk.Frame(frame, padding=(10, 10), relief="groove")
        control_panel_frame.pack(pady=10, fill='x')
        control_panel_frame.configure(style='Dark.TFrame')
        ttk.Button(control_panel_frame, text="Convert First Sheet to PDF", command=self.convert_to_pdf).pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.report_file_label.config(text=os.path.basename(file_path))
            self.validation_file_label.config(text=os.path.basename(file_path))
            self.convert_file_label.config(text=os.path.basename(file_path))
            self.df = pd.read_excel(self.selected_file, sheet_name=None)
            self.update_combo_boxes()
            self.status_var.set("Loaded file: " + self.selected_file)

    def update_combo_boxes(self):
        if self.df is not None and len(self.df) > 0:
            columns = list(self.df.values())[0].columns.tolist()
            self.report_column_combo['values'] = columns
            self.validation_column_combo['values'] = columns

    def extract_data(self):
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return
        extracted_data = pd.concat(self.df.values())
        messagebox.showinfo("Extracted Data", extracted_data.to_string())

    def add_files(self):
        file_paths = filedialog.askopenfilenames(title="Select Excel files to merge", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_paths:
            self.files_to_merge.extend(file_paths)
            messagebox.showinfo("Files Added", f"{len(file_paths)} files added for merging.")

    def merge_files(self):
        if not self.files_to_merge:
            messagebox.showwarning("Warning", "No files added for merging.")
            return
        merged_df = pd.concat([pd.read_excel(file) for file in self.files_to_merge], ignore_index=True)
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            merged_df.to_excel(output_path, index=False)
            messagebox.showinfo("Success", f"Merged files saved to: {output_path}")

    def convert_to_pdf(self):
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return
        first_sheet_name = list(self.df.keys())[0]
        sheet_df = self.df[first_sheet_name]
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            for index, row in sheet_df.iterrows():
                pdf.cell(0, 10, ', '.join(map(str, row.values)), ln=True)
            pdf.output(output_path)
            messagebox.showinfo("Success", f"PDF saved to: {output_path}")

    def generate_report(self):
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return
        column = self.report_column_var.get()
        if not column:
            messagebox.showwarning("Warning", "Please select a column for the report.")
            return
        report_data = self.df[list(self.df.keys())[0]][column].describe()
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            for line in report_data.to_string().split('\n'):
                pdf.cell(0, 10, line, ln=True)
            pdf.output(output_path)
            messagebox.showinfo("Success", f"Report saved to: {output_path}")

    def validate_data(self):
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return
        column = self.validation_column_var.get()
        if not column:
            messagebox.showwarning("Warning", "Please select a column to validate.")
            return
        missing_values = self.df[list(self.df.keys())[0]][column].isnull().sum()
        messagebox.showinfo("Validation Results", f"Missing values in column '{column}': {missing_values}")

if __name__ == "__main__":
    app = ExcelUtilityApp()
    app.mainloop()
