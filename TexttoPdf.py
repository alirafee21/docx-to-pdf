import docx
from docx import Document
import docx2pdf
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QProgressBar
import os

class CoverLetterGUI(QWidget):
    def __init__(self):
        super().__init__()

        # Create a label and line edit for the file name
        self.file_label = QLabel("File name:", self)
        self.file_edit = QLineEdit(self)

        # Create a label and line edit for the new text
        self.text_label = QLabel("New text:", self)
        self.text_edit = QLineEdit(self)

        # Create a label and line edit for new file name 
        self.new_file_label = QLabel("New file name:", self)
        self.new_file_edit = QLineEdit(self)

        # Create a button to start the conversion
        self.convert_button = QPushButton("Convert", self)
        self.convert_button.clicked.connect(self.convert)

        # Create a progress bar to show the conversion progress
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        # Set the layout of the widgets
        self.file_label.move(10, 10)
        self.file_edit.move(120, 10)
        self.text_label.move(10, 40)
        self.text_edit.move(120, 40)
        self.new_file_label.move(10, 70)
        self.new_file_edit.move(120, 70)
        self.convert_button.move(90, 100)
        self.progress_bar.move(15, 130)
        self.progress_bar.resize(290, 20)
        self.setGeometry(300, 300, 300, 180)
        self.setWindowTitle("Cover Letter GUI")
        self.show()
        
    def convert(self):
        # Get the file name and new text from the line edits
        file_name = self.file_edit.text()
        new_text = self.text_edit.text()
        new_file_name = self.new_file_edit.text()

        # Open the Word document
        document = docx.Document(file_name)

        # Iterate through the paragraphs in the document
        for paragraph in document.paragraphs:
            # Iterate through the runs in the paragraph
            for run in paragraph.runs:
                # Check if the run is bolded
                if run.bold:
                    # Replace the bolded text with the new text
                    run.bold = False
                    run.text = new_text


        # Update the progress bar
        self.progress_bar.setValue(25)
        # Save the modified document
        document.save(new_file_name)
        # destination_directory = 'D:/Resume - PEY/' 
        # os.replace(new_file_name, destination_directory)
        # Convert the Word document to a PDF
        self.progress_bar.setValue(50)
        docx2pdf.convert(new_file_name, new_file_name.replace('.docx', '.pdf'))
        self.progress_bar.setValue(75)
        # Remove the Word document
        try:
            os.remove(new_file_name)
        except OSError:
            print('Error: Cannot remove file')
        self.progress_bar.setValue(100)
        # os.replace(new_file_name.replace('.docx', '.pdf'), destination_directory)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = CoverLetterGUI()
    sys.exit(app.exec_())
