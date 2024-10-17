import os
import shutil
import pandas as pd  # Import pandas for Excel handling
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from datetime import datetime

DEFAULT_DIRECTORY = "REC_REVIEWERS"
DEFAULT_SUBDIRECTORIES = ["0 Protocol Files", "1", "2", "3"]

# Map of document names to their file extensions
DOCUMENTS_MAP = {
    "Form 06C ICA": "Form 06C ICA.docx",
    "Form 06B1 PRA": "Form 06B1 PRA.docx",
    "Form 06B2 PRA": "Form 06B2 PRA.docx",
    "Form 04A CERF": "Form 04A CER.docx"
}

class FolderGenerator:
    def __init__(self):
        self.file_paths_to_upload = []
        self.excel_file_path = os.path.join(DEFAULT_DIRECTORY, "folder_creation_log.xlsx")  # Path for Excel log

    def create_directory(self, dir_name):
        """Create a directory if it doesn't exist."""
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)

    def create_empty_document(self, folder_path, document_name):
        """Create an empty document file in the specified folder."""
        document_path = os.path.join(folder_path, document_name)
        try:
            with open(document_path, 'w') as f:
                pass  # Create an empty file
        except Exception as e:
            print(f"Error creating document: {e}")

    def create_subdirectories(self, base_dir, additional_names):
        """Create subdirectories with default names and additional names."""
        sub_dirs = []
        # Create the default directory for "0 Protocol Files"
        folder_name = DEFAULT_SUBDIRECTORIES[0]  # Always the first folder
        path = os.path.join(base_dir, folder_name)
        self.create_directory(path)
        sub_dirs.append(path)

        # Iterate through additional names based on the length of additional_names
        for i in range(1, len(additional_names) + 1):  # Start from 1 to len
            reviewer, selected_document = additional_names[i - 1]
            
            if selected_document:
                folder_name = f"{i} ({reviewer}) {selected_document}"
                # Create the folder for this reviewer with the document name included
                self.create_empty_document(path, DOCUMENTS_MAP[selected_document])  # Create the document in the "0 Protocol Files" folder
            else:
                folder_name = f"{i} ({reviewer})"

            # Create the folder for this reviewer
            folder_path = os.path.join(base_dir, folder_name)
            self.create_directory(folder_path)  # Create the subdirectory
            sub_dirs.append(folder_path)

            # Check if selected_document exists
            if selected_document:
                self.create_empty_document(folder_path, DOCUMENTS_MAP[selected_document])  # Create the document in the reviewer folder

            sub_dirs.append(folder_path)

        return sub_dirs

    def upload_files_to_list(self, listbox_widget):
        """Upload selected files and store their paths in a global list."""
        file_paths, _ = QFileDialog.getOpenFileNames(None, "Select Files", "", "All Files (*)")
        self.file_paths_to_upload = list(file_paths)

        listbox_widget.clear()
        if self.file_paths_to_upload:
            for file_path in self.file_paths_to_upload:
                filename = os.path.basename(file_path)
                listbox_widget.addItem(filename)

    def upload_files_to_folder(self, folder, parent_widget):
        """Upload selected files to the designated folder."""
        if self.file_paths_to_upload:
            for file_path in self.file_paths_to_upload:
                try:
                    shutil.copy(file_path, folder)
                except Exception as e:
                    QMessageBox.critical(parent_widget, "Error", f"Failed to upload the file: {e}")
            QMessageBox.information(parent_widget, "Success", f"{len(self.file_paths_to_upload)} files uploaded to {folder}")

    def log_folder_creation(self, main_folder_name, additional_names, parent_widget):
        """Log the folder creation details to an Excel file."""
        # Create a DataFrame from the folder creation details
        data = {
            'Timestamp': [pd.Timestamp.now()] * len(additional_names),  # Add a timestamp for each entry
            'Main Folder': [main_folder_name] * len(additional_names),
            'Reviewer': [reviewer for reviewer, _ in additional_names],
            'Document': [document for _, document in additional_names],
            'Status': ['Created'] * len(additional_names),  # Add a status if needed
        }
        df = pd.DataFrame(data)

        # Check if the Excel file exists
        if os.path.exists(self.excel_file_path):
            # Read existing data to ensure we don't overwrite
            existing_df = pd.read_excel(self.excel_file_path)

            # Append the new data to existing data
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save the DataFrame to the Excel file
        df.to_excel(self.excel_file_path, index=False)


    def handle_create_structure(self, main_folder_name, additional_names, parent_widget, listbox_widget, log_tab):
        """Create the folder structure and handle file uploads."""
        if not main_folder_name:
            QMessageBox.warning(parent_widget, "Warning", "Please enter a valid folder name.")
            return

        self.create_directory(DEFAULT_DIRECTORY)
        user_folder_path = os.path.join(DEFAULT_DIRECTORY, main_folder_name)
        self.create_directory(user_folder_path)
        self.create_subdirectories(user_folder_path, additional_names)

        # Log the creation of folders to Excel, now passing parent_widget
        self.log_folder_creation(main_folder_name, additional_names, parent_widget)

        # Upload files to the specified folder
        self.upload_files_to_folder(os.path.join(user_folder_path, "0 Protocol Files"), parent_widget)

        # Show success message
        QMessageBox.information(parent_widget, "Success", "Folder created successfully!")

        # Update the log view in LogTab
        log_tab.display_log_contents()
