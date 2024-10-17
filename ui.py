import os
from PyQt5.QtWidgets import (
    QMainWindow, QMessageBox,
    QLabel, QLineEdit, QTextEdit,
    QPushButton, QVBoxLayout, QWidget, QComboBox, QHBoxLayout, QSizePolicy, QListWidget, QDialog, QTableWidget, QTableWidgetItem
)
from PyQt5 import QtGui
from PyQt5.QtCore import Qt
from document_generator import DocumentGenerator
from folder_generator import FolderGenerator
import pandas as pd

class DocumentTab(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("background-color: #F7F7F7; font-family: Arial; font-size: 14px;")

        # Set the application icon
        self.setWindowIcon(QtGui.QIcon('REC_logo.ico'))

        self.documents = {
            "Certificate of Approval": "Certificate of Approval.docx",
            "Certificate of Exemption": "Certificate of Exemption.docx",
            "Certificate of Acceptance": "Certificate of Acceptance.docx"
        }
        
        self.default_output_dir = os.path.join(os.getcwd(), "REC_CERTIFICATES")
        if not os.path.exists(self.default_output_dir):
            os.makedirs(self.default_output_dir)

        self.input_file = None
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        self.document_dropdown = QComboBox()
        self.document_dropdown.addItems(self.documents.keys())
        layout.addWidget(QLabel("Select Document Type:"))
        layout.addWidget(self.document_dropdown)

        # Add form fields
        self.rec_code_entry = self.create_input_field(layout, "REC CODE:")
        self.principal_investigator_entry = self.create_input_field(layout, "PRINCIPAL INVESTIGATOR:")
        self.protocol_title_text = self.create_text_area(layout, "PROTOCOL TITLE:")
        self.adviser_entry = self.create_input_field(layout, "ADVISER:")

        # Add file paths labels
        self.output_dir_path_label = QLabel(f"Output Directory: {self.default_output_dir}")
        layout.addWidget(self.output_dir_path_label)

        # Add buttons to generate document in a responsive layout
        button_layout = QHBoxLayout()

        self.save_button = QPushButton("Save Document")
        self.save_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.save_button.clicked.connect(self.generate_and_save_document)
        button_layout.addWidget(self.save_button)

        self.print_button = QPushButton("Save and Print Document")
        self.print_button.setStyleSheet("background-color: #FF5722; color: white; font-weight: bold;")
        self.print_button.clicked.connect(self.generate_and_print_document)
        button_layout.addWidget(self.print_button)

        layout.addLayout(button_layout)
        layout.addStretch()

        central_widget.setLayout(layout)

        # Ensure buttons stay the same size
        self.save_button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.print_button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)

    def create_input_field(self, layout, label_text):
        label = QLabel(label_text)
        layout.addWidget(label)
        input_field = QLineEdit()
        input_field.setStyleSheet("padding: 10px; border: 1px solid #ccc; border-radius: 5px;")
        input_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addWidget(input_field)
        return input_field

    def create_text_area(self, layout, label_text):
        label = QLabel(label_text)
        layout.addWidget(label)
        text_area = QTextEdit()
        text_area.setStyleSheet("padding: 10px; border: 1px solid #ccc; border-radius: 5px;")
        text_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.MinimumExpanding)
        layout.addWidget(text_area)
        return text_area

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def generate_and_save_document(self):
        selected_document = self.document_dropdown.currentText()
        self.input_file = self.documents[selected_document]  # Get the input file based on the selected document type

        # Check if input file exists
        if not os.path.exists(self.input_file):
            self.show_error("The selected input file does not exist.")
            return

        # Retrieve user input
        rec_code = self.rec_code_entry.text()
        protocol_title = self.protocol_title_text.toPlainText().strip()
        principal_investigator = self.principal_investigator_entry.text()
        adviser = self.adviser_entry.text()

        # Initialize DocumentGenerator
        doc_gen = DocumentGenerator(self.input_file, self.default_output_dir)
        try:
            doc_gen.save_document(rec_code, protocol_title, principal_investigator, adviser)
            # Show success message
            QMessageBox.information(self, "Success", "Document saved successfully.")
        except Exception as e:
            self.show_error(f"An error occurred while saving the document: {str(e)}")

        # After generating the document, clear the input fields
        self.clear_inputs()

    def generate_and_print_document(self):
        """Generate the document and print it."""
        selected_document = self.document_dropdown.currentText()
        self.input_file = self.documents[selected_document]  # Get the input file based on the selected document type

        # Check if input file exists
        if not os.path.exists(self.input_file):
            self.show_error("Input file does not exist.")
            return

        # Retrieve user input
        rec_code = self.rec_code_entry.text()
        protocol_title = self.protocol_title_text.toPlainText().strip()
        principal_investigator = self.principal_investigator_entry.text()
        adviser = self.adviser_entry.text()

        # Initialize DocumentGenerator
        doc_gen = DocumentGenerator(self.input_file, self.default_output_dir)
        try:
            doc_gen.save_and_print_document(rec_code, protocol_title, principal_investigator, adviser)
            # Show success message
            QMessageBox.information(self, "Success", "Document saved and sent to print successfully.")
        except Exception as e:
            self.show_error(f"An error occurred while saving or printing the document: {str(e)}")

        # After generating and printing the document, clear the input fields
        self.clear_inputs()

    def clear_inputs(self):
        """Clear all input fields after document generation."""
        self.rec_code_entry.clear()
        self.principal_investigator_entry.clear()
        self.protocol_title_text.clear()
        self.adviser_entry.clear()


class FolderTab(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window  # Store the reference to the main window
        self.setStyleSheet("background-color: #F7F7F7; font-family: Arial; font-size: 14px;")
        self.folder_generator = FolderGenerator()
        self.reviewer_list = [
            "M. Carino", "A. De Guzman", "S. Jamilla", "M. Velez", 
            "A. Mendoza", "C. Villanueva", "A. Villegas", "R. Tiro", "P. Rios"
        ]
        self.init_ui()
    def init_ui(self):
        layout = QVBoxLayout()

        # Main Folder Name Entry
        layout.addWidget(QLabel("Enter Main Folder Name:"))
        self.main_folder_name_entry = QLineEdit()
        self.main_folder_name_entry.setStyleSheet("padding: 10px; border: 1px solid #ccc; border-radius: 5px;")
        self.main_folder_name_entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addWidget(self.main_folder_name_entry)

        # Reviewer Dropdowns for each reviewer
        self.document_dropdowns = []  # Track the document dropdowns
        for reviewer in self.reviewer_list:
            h_layout = QHBoxLayout()
            reviewer_label = QLabel(f"{reviewer}:")
            reviewer_label.setMinimumWidth(100)  # Consistent width for the label

            # Document Dropdown for selecting documents
            document_dropdown = QComboBox()
            document_dropdown.addItem("Select Document")
            document_dropdown.addItems([
                "Form 06C ICA",
                "Form 06B1 PRA",
                "Form 06B2 PRA",
                "Form 04A CERF"
            ])
            document_dropdown.setStyleSheet("padding: 10px; border: 1px solid #ccc; border-radius: 5px;")
            document_dropdown.setMinimumHeight(40)  # Make dropdown taller for better visibility
            self.document_dropdowns.append((reviewer, document_dropdown))  # Store with the reviewer

            h_layout.addWidget(reviewer_label)
            h_layout.addWidget(document_dropdown)
            layout.addLayout(h_layout)

        # Output Listbox for Uploaded Files
        self.listbox = QListWidget()
        layout.addWidget(self.listbox)

        # Buttons layout similar to DocumentTab
        button_layout = QHBoxLayout()

        self.upload_button = QPushButton("Upload Files")
        self.upload_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.upload_button.setMinimumHeight(50)
        self.upload_button.clicked.connect(self.upload_files)
        button_layout.addWidget(self.upload_button)

        self.create_structure_button = QPushButton("Create Folder Structure")
        self.create_structure_button.setStyleSheet("background-color: #FF5722; color: white; font-weight: bold;")
        self.create_structure_button.setMinimumHeight(50)
        self.create_structure_button.clicked.connect(self.create_folder_structure)
        button_layout.addWidget(self.create_structure_button)

        layout.addLayout(button_layout)
        layout.addStretch()

        self.setLayout(layout)

    def upload_files(self):
        """Upload files related to the selected documents."""
        self.folder_generator.upload_files_to_list(self.listbox)  # Call the upload logic to populate listbox

    def create_folder_structure(self):
        """Create the folder structure based on user inputs."""
        main_folder_name = self.main_folder_name_entry.text().strip()
        reviewers_documents = []

        # Collect selected documents for each reviewer
        for reviewer, document_dropdown in self.document_dropdowns:
            document_name = document_dropdown.currentText() if document_dropdown.currentIndex() > 0 else None

            # Append the reviewer and document name if the document is selected
            if document_name:
                reviewers_documents.append((reviewer, document_name))

        # Create the folder structure if documents are selected
        if reviewers_documents:
            # Use main_window to access log_tab
            self.folder_generator.handle_create_structure(main_folder_name, reviewers_documents, self, self.listbox, self.main_window.log_tab)

            # Reset the input fields and dropdowns
            self.reset_inputs()  # Call the reset function
        else:
            QMessageBox.warning(self, "Warning", "Please select a document before creating the folder structure.")

    def reset_inputs(self):
        """Reset input fields and dropdowns."""
        self.main_folder_name_entry.clear()  # Clear the main folder name entry
        for _, document_dropdown in self.document_dropdowns:
            document_dropdown.setCurrentIndex(0)  # Reset dropdowns to "Select Document"
        self.listbox.clear()  # Clear the uploaded files list


class LogTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.check_and_create_log_file()  # Check and create the log file if it doesn't exist
        self.display_log_contents()  # Automatically display log contents upon initialization

    def init_ui(self):
        layout = QVBoxLayout()

        # Create a QTableWidget to show the log contents
        self.table_widget = QTableWidget()
        self.table_widget.setSortingEnabled(True)  # Enable sorting on columns
        layout.addWidget(self.table_widget)

        self.setLayout(layout)

    def check_and_create_log_file(self):
        """Check if the log file exists and create it if it doesn't."""
        excel_path = os.path.join("DefaultDirectory", "folder_creation_log.xlsx")

        if not os.path.exists(excel_path):
            # Create a new DataFrame with the desired columns
            df = pd.DataFrame(columns=["Timestamp", "Reviewer", "Document", "Folder Created", "Status"])
            try:
                # Save the empty DataFrame to an Excel file
                df.to_excel(excel_path, index=False)
                QMessageBox.information(self, "Information", "Log file created successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to create log file: {str(e)}")

    def display_log_contents(self):
        """Display the contents of the folder_creation_log.xlsx file."""
        excel_path = os.path.join("DefaultDirectory", "folder_creation_log.xlsx")
        if not os.path.exists(excel_path):
            QMessageBox.warning(self, "Warning", "The log file does not exist.")
            return

        try:
            # Read the Excel file using pandas
            df = pd.read_excel(excel_path)

            # Check if DataFrame is empty
            if df.empty:
                QMessageBox.information(self, "Information", "The log file is empty.")
                return

            # Set the number of rows and columns
            self.table_widget.setRowCount(len(df))
            self.table_widget.setColumnCount(len(df.columns))
            self.table_widget.setHorizontalHeaderLabels(df.columns)

            # Fill the table with data
            for row in range(len(df)):
                for col in range(len(df.columns)):
                    self.table_widget.setItem(row, col, QTableWidgetItem(str(df.iat[row, col])))

            # Resize the columns to fit the content
            self.table_widget.resizeColumnsToContents()
            self.table_widget.resizeRowsToContents()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read the log file: {str(e)}")