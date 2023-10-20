import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from gui import Ui_Form as Ui_MainWindow 
from controller import conversions

class AppWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.selectFileButton.clicked.connect(self.select_file)
        self.convertButton.clicked.connect(self.perform_conversion)
        self.selected_file = ''

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_name:
            self.selected_file = file_name
            self.filePathLabel.setText(file_name)

    def perform_conversion(self):
        conversion_type = self.conversionComboBox.currentText()

        if not self.selected_file:
            QMessageBox.warning(self, "Error", "Please select a file first!")
            return

        try:
            if conversion_type == "PDF to Word":
                output_file = 'output/converted_' + os.path.basename(self.selected_file).split('.')[0] + '.docx'
                conversions.pdf_to_word(self.selected_file, output_file)
            elif conversion_type == "Word to PDF":
                output_file = conversions.word_to_pdf(self.selected_file)
            elif conversion_type == "Excel to Word":
                output_file = 'output/converted_' + os.path.basename(self.selected_file).split('.')[0] + '.docx'
                conversions.xl_to_word(self.selected_file, output_file)
            elif conversion_type == "Image to PDF":
                output_file = 'output/converted_' + os.path.basename(self.selected_file).split('.')[0] + '.pdf'
                conversions.image_to_pdf(self.selected_file)
            elif conversion_type == "Image Text Extract":
                output_file = 'output/extracted_' + os.path.basename(self.selected_file).split('.')[0] + '.txt'
                conversions.image_text_extractor(self.selected_file, output_file)
            elif conversion_type == "Text to Audio":
                output_file = 'output/converted_' + os.path.basename(self.selected_file).split('.')[0] + '.mp3'
                conversions.text_to_audio(self.selected_file, output_file)
            
            self.statusLabel.setText("Conversion successful!")
        except Exception as e:
            self.statusLabel.setText(f"Error: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec_())
