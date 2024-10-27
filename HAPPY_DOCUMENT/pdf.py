import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, Listbox, Toplevel
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pytesseract
from PIL import Image
import fitz
import os
import pdf2docx
import pandas as pd
from pdf2image import convert_from_path
import tabula
import time  # For simulating long operations (remove in production)

# Set the Tesseract executable path for Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\HP\scoop\apps\tesseract\current\tesseract.exe'  # Update with your Tesseract path

# Function to extract text using OCR from the PDF pages (image-based PDFs)
def extract_text_ocr(file_path):
    doc = fitz.open(file_path)  # Open the PDF using PyMuPDF
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)  # Load a page
        pix = page.get_pixmap()  # Convert page to an image (for OCR)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  # Convert to PIL Image

        # Use pytesseract to extract text from the image
        page_text = pytesseract.image_to_string(img)

        text += f"Page {page_num + 1}:\n{page_text}\n\n"  # Add page number for clarity
    return text

# Function to merge multiple PDFs
def merge_pdfs(file_paths, output_path, progress_var):
    merger = fitz.open()  # Using fitz for handling PDFs
    total_files = len(file_paths)
    for index, pdf in enumerate(file_paths):
        merger.insert_pdf(fitz.open(pdf))
        progress_var.set((index + 1) / total_files * 100)  # Update progress
        time.sleep(0.5)  # Simulate processing time (remove in production)
    merger.save(output_path)
    merger.close()

# Function to convert PDF to Word
def pdf_to_word(pdf_file, output_path, progress_var):
    try:
        pdf2docx.PdfToDocx().convert(pdf_file, output_path)
        progress_var.set(100)  # Set progress to 100% after completion
    except Exception as e:
        print(f"An error occurred while converting to Word: {str(e)}")

# Function to convert PDF to Excel using tabula-py
def pdf_to_excel(pdf_file, output_path, progress_var):
    try:
        tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
        with pd.ExcelWriter(output_path) as writer:
            for i, table in enumerate(tables):
                table.to_excel(writer, sheet_name=f"Sheet{i + 1}", index=False)
                progress_var.set((i + 1) / len(tables) * 100)  # Update progress
        progress_var.set(100)  # Set progress to 100% after completion
    except Exception as e:
        print(f"An error occurred while converting to Excel: {str(e)}")

class PDFUtilityGUI:
    def __init__(self, master):
        self.master = master
        master.title("PDF Utility")
        master.geometry("900x600")
        master.resizable(True, True)

        # Create a style
        self.style = ttk.Style(theme="darkly")

        # Main frame
        self.main_frame = ttk.Frame(master, padding="20 20 20 20")
        self.main_frame.pack(fill=BOTH, expand=YES)

        # Title
        ttk.Label(self.main_frame, text="PDF Utility", font=("Helvetica", 24, "bold")).pack(pady=20)

        # Operation selection frame
        self.operation_frame = ttk.LabelFrame(self.main_frame, text="Select Operation", padding="20 20 20 20")
        self.operation_frame.pack(fill=X, padx=10, pady=10)

        self.operation = tk.StringVar(value="search")
        operations = [
            ("Extract Text", "extract"),
            ("Merge PDFs", "merge"),
            ("Convert to Images", "convert_images"),
            ("Convert to Word", "convert_word"),
            ("Convert to Excel", "convert_excel"),
            ("Set Password", "set_password"),
            ("Search Text", "search_text"),
        ]

        for text, value in operations:
            ttk.Radiobutton(self.operation_frame, text=text, variable=self.operation, value=value).pack(side=LEFT, padx=10)

        # File selection frame
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=X, padx=10, pady=20)

        self.file_button = ttk.Button(self.file_frame, text="Select PDF File(s)", command=self.browse_files, width=30)
        self.file_button.pack(side=LEFT, padx=30)

        self.file_label = ttk.Label(self.file_frame, text="No files selected")
        self.file_label.pack(side=LEFT, padx=10)

        # Process button
        self.process_button = ttk.Button(self.main_frame, text="Process", command=self.process, style="success.TButton", width=40)
        self.process_button.pack(pady=40)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=X, padx=10, pady=10)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        self.status_bar.pack(fill=X, side=BOTTOM, pady=15)

        self.selected_files = []

    def browse_files(self):
        if self.operation.get() in ["extract", "convert_word", "convert_excel", "set_password", "search_text"]:
            files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        else:
            files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])

        if isinstance(files, str):  # Single file for extract, convert to Word, Excel, set password, or search text
            self.selected_files = [files] if files else []
        else:  # Multiple files for merge or convert to images
            self.selected_files = list(files)

        num_files = len(self.selected_files)
        self.file_label.config(text=f"{num_files} file{'s' if num_files != 1 else ''} selected")

    def process(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select PDF file(s).")
            return

        self.status_var.set("Processing...")
        self.master.update_idletasks()
        self.progress_var.set(0)  # Reset progress bar

        try:
            if self.operation.get() == "extract":
                text = extract_text_ocr(self.selected_files[0])
                self.show_extracted_text(text)
            elif self.operation.get() == "merge":
                if len(self.selected_files) < 2:
                    messagebox.showerror("Error", "Please select at least two PDF files to merge.")
                    return
                self.select_merge_order()
            elif self.operation.get() == "convert_images":
                output_folder = filedialog.askdirectory()
                if output_folder:
                    for pdf_file in self.selected_files:
                        pdf_to_images_with_fitz(pdf_file, output_folder)
                    messagebox.showinfo("Success", "PDF pages converted to images successfully.")
            elif self.operation.get() == "convert_word":
                output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
                if output_path:
                    pdf_to_word(self.selected_files[0], output_path, self.progress_var)
                    messagebox.showinfo("Success", f"Converted to Word successfully. Saved as {output_path}")
            elif self.operation.get() == "convert_excel":
                output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
                if output_path:
                    pdf_to_excel(self.selected_files[0], output_path, self.progress_var)
                    messagebox.showinfo("Success", f"Converted to Excel successfully. Saved as {output_path}")
            elif self.operation.get() == "set_password":
                password = simpledialog.askstring("Password", "Enter a password to protect the PDF:", show='*')
                if password:
                    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
                    if output_path:
                        set_pdf_password(self.selected_files[0], output_path, password)
                        messagebox.showinfo("Success", f"Password set successfully. Saved as {output_path}")
            elif self.operation.get() == "search_text":
                search_text = simpledialog.askstring("Search Text", "Enter text to search for:")
                if search_text:
                    results = search_text_in_pdf(self.selected_files[0], search_text)
                    self.show_search_results(results)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.status_var.set("Ready")
            self.progress_var.set(100)  # Ensure progress bar shows completion

    def select_merge_order(self):
        merge_window = Toplevel(self.master)
        merge_window.title("Select Order of Files")
        merge_window.geometry("800x600")

        listbox = Listbox(merge_window, selectmode='multiple', height=100 , width=500)
        for file in self.selected_files:
            listbox.insert(tk.END, file)
        listbox.pack(pady=25)

        merge_button = ttk.Button(merge_window, text="Merge Selected Order", command=lambda: self.merge_selected_order(listbox.get(listbox.curselection()), merge_window))
        merge_button.pack(pady=20)

    def merge_selected_order(self, selected_files, window):
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_path:
            self.progress_var.set(0)
            merge_pdfs(selected_files, output_path, self.progress_var)
            messagebox.showinfo("Success", f"PDFs merged successfully. Saved as {output_path}")
            window.destroy()

    def show_extracted_text(self, text):
        text_window = Toplevel(self.master)
        text_window.title("Extracted Text")
        text_window.geometry("600x400")

        text_area = scrolledtext.ScrolledText(text_window, wrap=tk.WORD)
        text_area.insert(tk.END, text)
        text_area.pack(fill=BOTH, expand=True)

    def show_search_results(self, results):
        result_window = Toplevel(self.master)
        result_window.title("Search Results")
        result_window.geometry("600x400")

        text_area = scrolledtext.ScrolledText(result_window, wrap=tk.WORD)
        text_area.insert(tk.END, results)
        text_area.pack(fill=BOTH, expand=True)

if __name__ == "__main__":
    root = ttk.Window()
    app = PDFUtilityGUI(root)
    root.mainloop()
