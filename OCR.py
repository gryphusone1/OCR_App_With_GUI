import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QMessageBox
import pytesseract
import cv2
import numpy as np
from docx import Document

# Preprocessing function to enhance image for OCR
def preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)  # Convert to grayscale
    img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)  # Resize to improve text
    img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]  # Apply threshold
    return img

class OCRApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OCR Application")
        self.setGeometry(100, 100, 800, 600)

        # Set up the layout
        self.layout = QVBoxLayout()

        # Create text box for OCR result
        self.text_box = QTextEdit(self)
        self.layout.addWidget(self.text_box)

        # Create open button for selecting files
        self.open_button = QPushButton("Open Images", self)
        self.open_button.clicked.connect(self.open_file)
        self.layout.addWidget(self.open_button)

        # Create save button for exporting to Word
        self.save_button = QPushButton("Save as Word", self)
        self.save_button.clicked.connect(self.save_to_word)
        self.layout.addWidget(self.save_button)

        # Set the layout
        self.setLayout(self.layout)

    def open_file(self):
        # Open multiple image files
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Open Files", "", "Images (*.png *.xpm *.jpg *.jpeg *.bmp *.tiff)")
        if file_paths:
            extracted_text = ""
            page_num = 1

            for file_path in file_paths:
                extracted_text += self.extract_text_from_image(file_path, page_num)
                page_num += 1

            self.text_box.setPlainText(extracted_text)

    def extract_text_from_image(self, file_path, page_num):
        # Preprocess the image for better OCR results
        img = preprocess_image(file_path)
        text = pytesseract.image_to_string(img)
        return f"===<page {page_num}>===\n{text}\n"

    def save_to_word(self):
        text = self.text_box.toPlainText().strip()
        if not text:
            QMessageBox.critical(self, "Error", "No text to save!")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Word Document (*.docx)")
        if file_path:
            # Create a Word document and save the extracted text
            document = Document()
            document.add_paragraph(text)
            document.save(file_path)
            QMessageBox.information(self, "Success", f"Text saved to {file_path}")

# Main entry point to run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OCRApp()
    window.show()
    sys.exit(app.exec_())
