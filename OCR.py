import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import pytesseract
from docx import Document

# Set path to tesseract executable if needed (e.g., Windows)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def open_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.png;*.jpeg;*.bmp;*.tiff")])
    if file_path:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img)
        text_box.delete(1.0, tk.END)
        text_box.insert(tk.END, text)

def save_to_word():
    text = text_box.get(1.0, tk.END).strip()
    if not text:
        messagebox.showerror("Error", "No text to save!")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                             filetypes=[("Word Document", "*.docx")])
    if file_path:
        document = Document()
        document.add_paragraph(text)
        document.save(file_path)
        messagebox.showinfo("Success", f"Text saved to {file_path}")

# GUI setup
root = tk.Tk()
root.title("OCR Application")

frame = tk.Frame(root)
frame.pack(pady=10, padx=10)

open_button = tk.Button(frame, text="Open Image", command=open_image)
open_button.pack(side=tk.LEFT, padx=5)

save_button = tk.Button(frame, text="Save as Word", command=save_to_word)
save_button.pack(side=tk.LEFT, padx=5)

text_box = tk.Text(root, wrap='word', width=60, height=20)
text_box.pack(pady=10)

root.mainloop()
