import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageTk
import os
import win32com.client  # For converting Word to PDF in Windows

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convert Ease")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        self.root.resizable(True, True)

        # Background images
        self.main_bg_image = "C:\\Users\\Payal\\Downloads\\WhatsApp Image 2024-12-29 at 16.49.18.jpeg"
        self.image_to_pdf_bg_image = "C:\\Users\\Payal\\Downloads\\WhatsApp Image 2024-12-29 at 16.49.17 (1).jpeg"
        self.word_to_pdf_bg_image = "C:\\Users\\Payal\\Downloads\\WhatsApp Image 2024-12-29 at 16.49.17.jpeg"
        self.image_resizer_bg_image = "C:\\Users\\Payal\\Downloads\\WhatsApp Image 2024-12-29 at 16.49.16.jpeg"

        # Default font for the app
        self.font = ("Calibri", 12)

        self.show_main_page()

    def show_main_page(self):
        self.clear_frame()
        self.set_background(self.main_bg_image)

        tk.Label(self.root, text="Convert Ease", font=("Calibri", 24, "bold"), bg="#ffffff").pack(pady=30)

        style = ttk.Style()
        style.configure("RoundedButton.TButton", font=("Calibri", 12, "bold"), padding=(10, 5), 
                        relief="flat", background="#007BFF", foreground="#000000", borderwidth=2)
        style.map("RoundedButton.TButton",
                  background=[("active", "#0056b3")])

        button_options = {'style': "RoundedButton.TButton", 'padding': (10, 10)}

        button1 = ttk.Button(self.root, text="Image to PDF", command=self.image_to_pdf, **button_options)
        button1.pack(fill=tk.NONE, pady=10, padx=20, ipadx=20)

        button2 = ttk.Button(self.root, text="Word to PDF", command=self.word_to_pdf, **button_options)
        button2.pack(fill=tk.NONE, pady=10, padx=20, ipadx=20)

        button3 = ttk.Button(self.root, text="Image Resizer", command=self.image_resizer, **button_options)
        button3.pack(fill=tk.NONE, pady=10, padx=20, ipadx=20)

    def set_background(self, image_path):
        bg = Image.open(image_path)
        bg = bg.resize((self.root.winfo_width(), self.root.winfo_height()), Image.LANCZOS)
        bg_image = ImageTk.PhotoImage(bg)

        bg_label = tk.Label(self.root, image=bg_image)
        bg_label.image = bg_image
        bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Update background dynamically on resize
        self.root.bind("<Configure>", lambda event: self.resize_background(bg_label, image_path))

    def resize_background(self, bg_label, image_path):
        bg = Image.open(image_path)
        bg = bg.resize((self.root.winfo_width(), self.root.winfo_height()), Image.LANCZOS)
        bg_image = ImageTk.PhotoImage(bg)

        bg_label.config(image=bg_image)
        bg_label.image = bg_image

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def image_to_pdf(self):
        self.clear_frame()
        self.set_background(self.image_to_pdf_bg_image)

        def select_images():
            image_files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
            if image_files:
                selected_files.extend(image_files)
                listbox.delete(0, tk.END)
                listbox.insert(tk.END, *selected_files)

        def convert_to_pdf():
            if not selected_files:
                messagebox.showwarning("No Images", "Please select images to convert.")
                return

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not save_path:
                return

            pdf_canvas = canvas.Canvas(save_path, pagesize=letter)
            for img_path in selected_files:
                img = Image.open(img_path)

                # Resize the image to fit the page size (letter size: 612x792)
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                max_width = 600  # max width for the image on the PDF
                max_height = 750  # max height for the image on the PDF

                if img_width > max_width or img_height > max_height:
                    # Scale to fit within the defined dimensions
                    if aspect_ratio > 1:
                        img_width = max_width
                        img_height = int(max_width / aspect_ratio)
                    else:
                        img_height = max_height
                        img_width = int(max_height * aspect_ratio)

                # Draw the image at the center of the page
                x = (612 - img_width) / 2  # Position image horizontally centered
                y = (792 - img_height) / 2  # Position image vertically centered
                pdf_canvas.drawImage(img_path, x, y, width=img_width, height=img_height)
                pdf_canvas.showPage()

            pdf_canvas.save()
            messagebox.showinfo("Success", f"PDF saved at {save_path}")

        selected_files = []

        tk.Label(self.root, text="Image to PDF Converter", font=("Calibri", 16, "bold"), bg="#ffffff").pack(pady=20)

        style = ttk.Style()
        style.configure("BrownButton.TButton", font=("Calibri", 12, "bold"), padding=(10, 5),
                        relief="flat", background="#007BFF", foreground="#000000", borderwidth=2)
        style.map("BrownButton.TButton",
                  background=[("active", "#0056b3")])

        ttk.Button(self.root, text="Select Images", command=select_images, style="BrownButton.TButton").pack(pady=10)
        listbox = tk.Listbox(self.root, width=50, height=10)
        listbox.pack(pady=10)

        ttk.Button(self.root, text="Convert to PDF", command=convert_to_pdf, style="BrownButton.TButton").pack(pady=10)
        ttk.Button(self.root, text="Back", command=self.show_main_page, style="BrownButton.TButton").pack(pady=10)

    def word_to_pdf(self):
        self.clear_frame()
        self.set_background(self.word_to_pdf_bg_image)

        def convert_word_to_pdf():
            word_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx;*.doc")])
            if not word_path:
                return

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not save_path:
                return

            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False  # To ensure Word is not visible during processing
                doc = word.Documents.Open(word_path)
                print(f"Opened Word document: {word_path}")  # Debugging log

                # Save the document as PDF
                doc.SaveAs(save_path, FileFormat=17)  # FileFormat 17 is for PDF
                print(f"PDF saved at: {save_path}")  # Debugging log
                doc.Close()
                word.Quit()
                messagebox.showinfo("Success", f"PDF saved at {save_path}")
            except Exception as e:
                print(f"Error: {e}")  # Debugging log for error
                messagebox.showerror("Error", f"Failed to convert Word to PDF: {e}")

        tk.Label(self.root, text="Word to PDF Converter", font=("Calibri", 16, "bold"), bg="#ffffff").pack(pady=20)

        style = ttk.Style()
        style.configure("LightBrownButton.TButton", font=("Calibri", 12, "bold"), padding=(10, 5),
                        relief="flat", background="#007BFF", foreground="#000000", borderwidth=2)
        style.map("LightBrownButton.TButton",
                  background=[("active", "#0056b3")])

        ttk.Button(self.root, text="Convert Word to PDF", command=convert_word_to_pdf, style="LightBrownButton.TButton").pack(pady=10)
        ttk.Button(self.root, text="Back", command=self.show_main_page, style="LightBrownButton.TButton").pack(pady=10)

    def image_resizer(self):
        self.clear_frame()
        self.set_background(self.image_resizer_bg_image)

        def resize_image():
            img_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
            if not img_path:
                return

            save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
            if not save_path:
                return

            try:
                img = Image.open(img_path)
                new_width = int(width_entry.get())  # No validation on size
                new_height = int(height_entry.get())
                img = img.resize((new_width, new_height), Image.LANCZOS)  # Resize using Lanczos filter for better quality
                img.save(save_path)
                messagebox.showinfo("Success", f"Image resized and saved at {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to resize image: {e}")

        tk.Label(self.root, text="Image Resizer", font=("Calibri", 16, "bold"), bg="#ffffff").pack(pady=20)

        tk.Label(self.root, text="Width:", font=("Calibri", 12), bg="#ffffff").pack(pady=5)
        width_entry = tk.Entry(self.root, font=("Calibri", 12))
        width_entry.pack(pady=5)

        tk.Label(self.root, text="Height:", font=("Calibri", 12), bg="#ffffff").pack(pady=5)
        height_entry = tk.Entry(self.root, font=("Calibri", 12))
        height_entry.pack(pady=5)

        style = ttk.Style()
        style.configure("BlueButton.TButton", font=("Calibri", 12, "bold"), padding=(10, 5),
                        relief="flat", background="#007BFF", foreground="#000000", borderwidth=2)
        style.map("BlueButton.TButton",
                  background=[("active", "#0056b3")])

        ttk.Button(self.root, text="Resize Image", command=resize_image, style="BlueButton.TButton").pack(pady=10)
        ttk.Button(self.root, text="Back", command=self.show_main_page, style="BlueButton.TButton").pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
