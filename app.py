import sys
from PyQt5.QtWidgets import QApplication, QTabWidget, QWidget, QVBoxLayout
from PyQt5 import QtGui 
from ui import DocumentTab, FolderTab, LogTab

class MainWindow(QTabWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("REC Document and Folder Generator")
        self.setGeometry(100, 100, 800, 600)
        self.setWindowIcon(QtGui.QIcon('REC_logo.ico'))

        # Add tabs
        self.add_tabs()

    def add_tabs(self):
        # Document Generator Tab
        self.document_tab = QWidget()
        self.folder_tab = QWidget()

        # Create log tab before passing to FolderTab
        self.log_tab = LogTab(self)  # Create the log tab

        # Create layouts for each tab
        self.document_layout = QVBoxLayout()
        self.folder_layout = QVBoxLayout()

        # Add document generator UI
        self.document_ui = DocumentTab(self)
        self.document_layout.addWidget(self.document_ui)
        self.document_tab.setLayout(self.document_layout)

        # **Place it here to pass log_tab to FolderTab**
        self.folder_ui = FolderTab(self)  # Pass the main window as parent
        self.folder_layout.addWidget(self.folder_ui)
        self.folder_tab.setLayout(self.folder_layout)

        # Add the tabs to the QTabWidget
        self.addTab(self.document_tab, "Document Generator")
        self.addTab(self.folder_tab, "Folder Creator")
        self.addTab(self.log_tab, "View Log")  # Add log tab

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
