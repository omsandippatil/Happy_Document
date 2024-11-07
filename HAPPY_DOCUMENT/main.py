import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from word import WordUtilityGUI  # Assuming your word utility is imported here
from pdf import PDFUtilityGUI  # Assuming your pdf utility is imported here
from image import ImageUtilityGUI  # Import the ImageUtilityGUI from image.py
from excel import ExcelUtilityApp  # Import the EnhancedExcelUtilityApp from excel.py
import webbrowser  # Import webbrowser to open links
from PIL import Image, ImageTk
from aexcel import EnhancedExcelUtilityApp  # Import the ExcelUtilityApp for advanced utility

class MainApp:
    def __init__(self, master):
        self.master = master
        master.title("HAPPY DOCUMENT")
        master.geometry("500x300")
        master.iconbitmap("icon.ico")

        # Create a custom title bar
        self.title_frame = ttk.Frame(master, padding=(10, 5))
        self.title_frame.pack(fill=X)

        # Load and display the icon image
        self.load_icon()

        # Style
        self.style = ttk.Style(theme="darkly")
        self.style.configure("TButton", background='lightblue', foreground='black', borderwidth=1, font=('Helvetica', 14))

        # Main frame
        self.main_frame = ttk.Frame(master, padding="20 20 20 20")
        self.main_frame.pack(fill=BOTH, expand=YES)

        # Title
        ttk.Label(self.main_frame, text="HAPPY DOCUMENT", font=("Helvetica", 24, "bold")).pack(pady=20)

        # Buttons for PDF, Word, Image, and Excel Utilities
        self.pdf_button = ttk.Button(self.main_frame, text="PDF Utility", command=self.open_pdf_utility, width=30)
        self.pdf_button.pack(pady=10)

        self.word_button = ttk.Button(self.main_frame, text="Word Utility", command=self.open_word_utility, width=30)
        self.word_button.pack(pady=10)

        self.image_button = ttk.Button(self.main_frame, text="Image Utility", command=self.open_image_utility, width=30)
        self.image_button.pack(pady=10)

        self.excel_button = ttk.Button(self.main_frame, text="Excel Utility", command=self.open_excel_utility, width=30)
        self.excel_button.pack(pady=10)

        self.aexcel_button = ttk.Button(self.main_frame, text="Advanced Excel Utility", command=self.open_Aexcel_utility, width=30)
        self.aexcel_button.pack(pady=10)

        # Footer
        self.footer_frame = ttk.Frame(master)
        self.footer_frame.pack(side=BOTTOM, fill=X, pady=(10, 0))

        footer_label = ttk.Label(self.footer_frame, text="Created By : ", font=("Helvetica", 12))
        footer_label.pack(side=LEFT, padx=5)

        link_label = ttk.Label(self.footer_frame, text="Aditya Nalawade", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label.pack(side=LEFT)

        link_label2 = ttk.Label(self.footer_frame, text="GitHub", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label2.pack(side=LEFT, padx=5)
        link_label2.bind("<Button-1>", self.open_link2)

        link_label3 = ttk.Label(self.footer_frame, text="Mail", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label3.pack(side=LEFT, padx=5)
        link_label3.bind("<Button-1>", self.open_link3)

        link_label.bind("<Button-1>", self.open_link)

    def load_icon(self):
        icon_image = Image.open("icon.ico")
        icon_image = icon_image.resize((130, 130))
        self.icon = ImageTk.PhotoImage(icon_image)
        icon_label = ttk.Label(self.title_frame, image=self.icon)
        icon_label.pack(side=TOP, padx=5)

    def open_link(self, event):
        webbrowser.open("https://www.linkedin.com/in/aditya-nalawade-a4b081297?utm_source=share&utm_campaign=share_via&utm_content=profile&utm_medium=android_app")

    def open_link2(self, event):
        webbrowser.open("https://github.com/Adiiiicodes")

    def open_link3(self, event):
        webbrowser.open("adityacodes8@gmail.com")

    def open_pdf_utility(self):
        pdf_window = tk.Toplevel(self.master)
        PDFUtilityGUI(pdf_window)

    def open_word_utility(self):
        word_window = tk.Toplevel(self.master)
        WordUtilityGUI(word_window)

    def open_image_utility(self):
        image_window = tk.Toplevel(self.master)
        ImageUtilityGUI(image_window)

    def open_excel_utility(self):
        excel_window = tk.Toplevel(self.master)
        excel_window.title("Excel Utility")
        ExcelUtilityApp(excel_window)

    def open_Aexcel_utility(self):
        aexcel_window = tk.Toplevel(self.master)
        aexcel_window.title("Advanced Excel Utility")
        EnhancedExcelUtilityApp()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
