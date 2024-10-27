import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import os

class ImageCropper:
    def __init__(self, master, image_path):
        self.master = master
        self.image_path = image_path
        self.image = Image.open(image_path)
        self.tk_image = ImageTk.PhotoImage(self.image)
        
        # Create main frame
        self.main_frame = ttk.Frame(master)
        self.main_frame.pack(fill='both', expand=True)
        
        # Create a scrollable canvas to hold the image
        self.canvas_frame = ttk.Frame(self.main_frame)
        self.canvas_frame.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, width=800, height=600)
        self.canvas.create_image(0, 0, image=self.tk_image, anchor='nw')

        # Add scrollbars
        self.scroll_x = ttk.Scrollbar(self.canvas_frame, orient='horizontal', command=self.canvas.xview)
        self.scroll_y = ttk.Scrollbar(self.canvas_frame, orient='vertical', command=self.canvas.yview)
        self.canvas.config(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)

        # Pack canvas and scrollbars
        self.scroll_y.pack(side='right', fill='y')
        self.canvas.pack(side='top', fill='both', expand=True)
        self.scroll_x.pack(side='bottom', fill='x')

        # Create button frame
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill='x', pady=10)

        # Add buttons
        self.crop_button = ttk.Button(self.button_frame, text="Crop Selection", command=self.crop_image, style="success.TButton")
        self.crop_button.pack(side='left', padx=5)
        
        self.clear_button = ttk.Button(self.button_frame, text="Clear Selection", command=self.clear_selection)
        self.clear_button.pack(side='left', padx=5)

        # Set scroll region for the entire image
        self.canvas.config(scrollregion=self.canvas.bbox('all'))

        self.rect = None
        self.start_x = None
        self.start_y = None

        # Bind mouse events for cropping
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
        # Bind mouse wheel events for scrolling
        self.canvas.bind("<MouseWheel>", self.on_mousewheel_y)  # Windows
        self.canvas.bind("<Shift-MouseWheel>", self.on_mousewheel_x)  # Windows with Shift
        self.canvas.bind("<Button-4>", self.on_mousewheel_up)  # Linux
        self.canvas.bind("<Button-5>", self.on_mousewheel_down)  # Linux
        self.canvas.bind("<Shift-Button-4>", self.on_mousewheel_left)  # Linux with Shift
        self.canvas.bind("<Shift-Button-5>", self.on_mousewheel_right)  # Linux with Shift

    def on_mousewheel_y(self, event):
        # Windows mouse wheel scrolling (vertical)
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_mousewheel_x(self, event):
        # Windows mouse wheel scrolling with Shift (horizontal)
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_mousewheel_up(self, event):
        # Linux mouse wheel up (vertical)
        self.canvas.yview_scroll(-1, "units")

    def on_mousewheel_down(self, event):
        # Linux mouse wheel down (vertical)
        self.canvas.yview_scroll(1, "units")

    def on_mousewheel_left(self, event):
        # Linux mouse wheel with Shift (horizontal left)
        self.canvas.xview_scroll(-1, "units")

    def on_mousewheel_right(self, event):
        # Linux mouse wheel with Shift (horizontal right)
        self.canvas.xview_scroll(1, "units")

    def clear_selection(self):
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None

    def on_button_press(self, event):
        # Remove any existing rectangle
        if self.rect:
            self.canvas.delete(self.rect)

        # Set starting position for the rectangle
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)

        # Create a new rectangle
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline='red', width=2
        )

    def on_mouse_drag(self, event):
        # Update the rectangle's size as the mouse is dragged
        curr_x = self.canvas.canvasx(event.x)
        curr_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, curr_x, curr_y)

    def on_button_release(self, event):
        # The cropping area is finalized when the mouse is released
        pass

    def crop_image(self):
        if not self.rect:
            messagebox.showerror("Error", "Please select a cropping area first.")
            return

        # Get the coordinates of the rectangle
        coords = self.canvas.coords(self.rect)
        # Ensure coordinates are in the correct order (left, top, right, bottom)
        left = min(coords[0], coords[2])
        top = min(coords[1], coords[3])
        right = max(coords[0], coords[2])
        bottom = max(coords[1], coords[3])
        
        crop_box = (int(left), int(top), int(right), int(bottom))

        # Crop the image
        cropped_image = self.image.crop(crop_box)
        
        # Save the cropped image
        save_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Files", "*.png"), ("JPEG Files", "*.jpg;*.jpeg")]
        )
        if save_path:
            cropped_image.save(save_path)
            messagebox.showinfo("Success", "Image cropped and saved successfully.")

class ImageUtilityGUI:
    def __init__(self, master):
        self.master = master
        master.title("Image Utility")
        master.geometry("900x600")
        master.resizable(True, True)

        # Create a style
        self.style = ttk.Style(theme="darkly")

        # Main frame
        self.main_frame = ttk.Frame(master, padding="20 20 20 20")
        self.main_frame.pack(fill=BOTH, expand=YES)

        # Title
        ttk.Label(self.main_frame, text="Image Utility", font=("Helvetica", 24, "bold")).pack(pady=20)

        # Operation selection frame
        self.operation_frame = ttk.LabelFrame(self.main_frame, text="Select Operation", padding="20 20 20 20")
        self.operation_frame.pack(fill=X, padx=10, pady=10)

        self.operation = tk.StringVar(value="convert")
        operations = [
            ("Convert Image Format", "convert"),
            ("Extract Text", "extract"),
            ("Resize Image", "resize"),
            ("Rotate Image", "rotate"),
            ("Compress Image", "compress"),
            ("Crop Image", "crop"),
            ("Batch Processing", "batch"),
            ("Preview Images", "preview"),
        ]

        for text, value in operations:
            ttk.Radiobutton(self.operation_frame, text=text, variable=self.operation, value=value).pack(side=LEFT, padx=10)

        # File selection frame
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=X, padx=10, pady=20)

        self.file_button = ttk.Button(self.file_frame, text="Select Image File(s)", command=self.browse_files, width=30)
        self.file_button.pack(side=LEFT, padx=30)

        self.file_label = ttk.Label(self.file_frame, text="No files selected")
        self.file_label.pack(side=LEFT, padx=10)

        # Process button
        self.process_button = ttk.Button(self.main_frame, text="Process", command=self.process, style="success.TButton", width=40)
        self.process_button.pack(pady=40)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        self.status_bar.pack(fill=X, side=BOTTOM, pady=15)

        self.selected_files = []

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
        self.selected_files = list(files)
        num_files = len(self.selected_files)
        self.file_label.config(text=f"{num_files} file{'s' if num_files != 1 else ''} selected")

    def process(self):
        operation = self.operation.get()
        if operation == "convert":
            self.convert_image_format()
        elif operation == "extract":
            self.extract_text()
        elif operation == "resize":
            self.resize_image()
        elif operation == "rotate":
            self.rotate_image()
        elif operation == "compress":
            self.compress_image()
        elif operation == "crop":
            if self.selected_files:
                # Open a new window for cropping
                crop_window = tk.Toplevel(self.master)
                crop_window.title("Crop Image")
                ImageCropper(crop_window, self.selected_files[0])  # Use the first selected image
            else:
                messagebox.showerror("Error", "Please select an image file to crop.")
        elif operation == "batch":
            self.batch_process()
        elif operation == "preview":
            self.preview_images()

    def convert_image_format(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select an image file to convert.")
            return
        
        output_format = simpledialog.askstring("Convert", "Enter output format (jpg, png, bmp):")
        if not output_format:
            return
        
        for file in self.selected_files:
            image = Image.open(file)
            output_file = os.path.splitext(file)[0] + f".{output_format}"
            image.save(output_file)
        
        messagebox.showinfo("Success", f"Converted {len(self.selected_files)} image(s) to {output_format}.")

    def extract_text(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select an image file to extract text from.")
            return
        
        text = ""
        for file in self.selected_files:
            image = Image.open(file)
            text += pytesseract.image_to_string(image) + "\n"
        
        self.show_text_window(text)

    def resize_image(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select an image file to resize.")
            return
        
        width = simpledialog.askinteger("Resize", "Enter new width:")
        height = simpledialog.askinteger("Resize", "Enter new height:")
        if width is None or height is None:
            return
        
        for file in self.selected_files:
            image = Image.open(file)
            resized_image = image.resize((width, height))
            output_file = os.path.splitext(file)[0] + "_resized.jpg"
            resized_image.save(output_file)
        
        messagebox.showinfo("Success", f"Resized {len(self.selected_files)} image(s).")

    def rotate_image(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select an image file to rotate.")
            return
        
        angle = simpledialog.askinteger("Rotate", "Enter rotation angle (degrees):")
        if angle is None:
            return
        
        for file in self.selected_files:
            image = Image.open(file)
            rotated_image = image.rotate(angle)
            output_file = os.path.splitext(file)[0] + "_rotated.jpg"
            rotated_image.save(output_file)
        
        messagebox.showinfo("Success", f"Rotated {len(self.selected_files)} image(s) by {angle} degrees.")

    def compress_image(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Please select an image file to compress.")
            return
        
        quality = simpledialog.askinteger("Compress", "Enter quality (1-100):", minvalue=1, maxvalue=100)
        if quality is None:
            return
        
        for file in self.selected_files:
            image = Image.open(file)
            output_file = os.path.splitext(file)[0] + "_compressed.jpg"
            image.save(output_file, quality=quality)
        
        messagebox.showinfo("Success", f"Compressed {len(self.selected_files)} image(s).")

    def batch_process(self):
        messagebox.showinfo("Info", "Batch processing is not yet implemented.")

    def preview_images(self):
        preview_window = tk.Toplevel(self.master)
        preview_window.title("Preview Images")
        for file in self.selected_files:
            image = Image.open(file)
            tk_image = ImageTk.PhotoImage(image)
            label = tk.Label(preview_window, image=tk_image)
            label.image = tk_image  # Keep a reference to avoid garbage collection
            label.pack(pady=10)

    def show_text_window(self, text):
        text_window = tk.Toplevel(self.master)
        text_window.title("Extracted Text")
        text_area = scrolledtext.ScrolledText(text_window, wrap=tk.WORD)
        text_area.pack(expand=True, fill='both')
        text_area.insert(tk.END, text)
        text_area.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageUtilityGUI(root)
    root.mainloop()
