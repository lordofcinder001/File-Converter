from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QComboBox, QLineEdit, QLabel, QMessageBox
from PyQt5.QtGui import QColor, QPalette, QFont, QIcon
from PyQt5.QtCore import Qt
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

class FileConverter(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.resize(500, 200)  # Set window size
        self.setWindowTitle('File Converter')  # Set window title
        self.setWindowIcon(QIcon('icon.png'))  # Set window icon

        # Set window background color
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor('#3498db'))  # You can choose your own color here
        self.setPalette(palette)

        # Set font
        font = QFont("Arial", 12)

        layout = QVBoxLayout()

        self.btn_select_file = QPushButton('Select File', self)
        self.btn_select_file.setFont(font)
        self.btn_select_file.setStyleSheet("background-color: #2ecc71; color: #ffffff;")  # Green button
        self.btn_select_file.clicked.connect(self.select_file)

        self.btn_select_dir = QPushButton('Select Destination Directory', self)
        self.btn_select_dir.setFont(font)
        self.btn_select_dir.setStyleSheet("background-color: #e67e22; color: #ffffff;")  # Orange button
        self.btn_select_dir.clicked.connect(self.select_dir)

        self.combo_box = QComboBox(self)
        self.combo_box.setFont(font)
        self.combo_box.setStyleSheet("background-color: #9b59b6; color: #ffffff;")  # Purple combo box
        self.combo_box.addItem("PDF to Word")
        self.combo_box.addItem("Word to PDF")
        self.combo_box.addItem("JPG to PNG")
        self.combo_box.addItem("PNG to JPG")

        self.label = QLabel("Rename File (optional):", self)
        self.label.setFont(font)
        self.text_field = QLineEdit(self)
        self.text_field.setFont(font)

        self.btn_convert = QPushButton('Convert', self)
        self.btn_convert.setFont(font)
        self.btn_convert.setStyleSheet("background-color: #e74c3c; color: #ffffff;")  # Red button
        self.btn_convert.clicked.connect(self.convert_file)

        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_select_dir)
        layout.addWidget(self.combo_box)
        layout.addWidget(self.label)
        layout.addWidget(self.text_field)
        layout.addWidget(self.btn_convert)

        self.setLayout(layout)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message', "Are you sure you want to quit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def select_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.file_name, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)", options=options)
        if self.file_name:
            print(self.file_name)

    def select_dir(self):
        self.dir_name = QFileDialog.getExistingDirectory(self, "Select Directory")
        if self.dir_name:
            print(self.dir_name)

    def convert_file(self):
        conversion_type = self.combo_box.currentText()
        output_file_name = self.text_field.text() if self.text_field.text() else "output"
        if conversion_type == "PDF to Word":
            cv = Converter(self.file_name)
            cv.convert(f"{self.dir_name}/{output_file_name}.docx", start=0, end=None)
            cv.close()
        elif conversion_type == "Word to PDF":
            convert(self.file_name, f"{self.dir_name}/{output_file_name}.pdf")
        elif conversion_type == "JPG to PNG":
            img = Image.open(self.file_name)
            img.save(f"{self.dir_name}/{output_file_name}.png")
        elif conversion_type == "PNG to JPG":
            img = Image.open(self.file_name)
            img = img.convert("RGB")
            img.save(f"{self.dir_name}/{output_file_name}.jpg")
        QMessageBox.information(self, 'Message', "File converted successfully!")

if __name__ == '__main__':
    app = QApplication([])
    ex = FileConverter()
    ex.show()
    app.exec_()
