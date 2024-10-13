import sys
import os
import shutil
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QListWidget, 
                             QInputDialog, QFileDialog, QMessageBox, QStatusBar)

class FileManager(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("File Manager")
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()
        self.setStyleSheet("background-color: #f0f0f0;")

        # Location input layout
        location_layout = QHBoxLayout()
        self.location_label = QLabel("Location:")
        self.location_input = QLineEdit()
        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_location)
        location_layout.addWidget(self.location_label)
        location_layout.addWidget(self.location_input)
        location_layout.addWidget(self.browse_button)

        # Search inputs
        self.extension_label = QLabel("File Extension:")
        self.extension_input = QLineEdit()
        self.search_label = QLabel("Search Unique Number:")
        self.search_input = QLineEdit()

        # Search button
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_files)

        # File list display
        self.file_list = QListWidget()
        self.file_list.setStyleSheet("background-color: white;")

        # Action buttons layout
        action_layout = QHBoxLayout()
        self.copy_button = QPushButton("Copy")
        self.copy_button.clicked.connect(self.copy_file)
        self.move_button = QPushButton("Move")
        self.move_button.clicked.connect(self.move_file)
        self.open_button = QPushButton("Open")
        self.open_button.clicked.connect(self.open_file)
        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_file)
        self.create_folder_button = QPushButton("Create Folder")
        self.create_folder_button.clicked.connect(self.create_folder)

        # Adding widgets to layout
        layout.addLayout(location_layout)
        layout.addWidget(self.extension_label)
        layout.addWidget(self.extension_input)
        layout.addWidget(self.search_label)
        layout.addWidget(self.search_input)
        layout.addWidget(self.search_button)
        layout.addWidget(self.file_list)
        
        # Action Buttons
        action_layout.addWidget(self.copy_button)
        action_layout.addWidget(self.move_button)
        action_layout.addWidget(self.open_button)
        action_layout.addWidget(self.delete_button)
        action_layout.addWidget(self.create_folder_button)
        layout.addLayout(action_layout)

        # Status Bar
        self.status_bar = QStatusBar()
        layout.addWidget(self.status_bar)

        self.setLayout(layout)

        # Apply styles to buttons
        self.apply_styles()

    def apply_styles(self):
        button_style = """
            QWidget {
                background-color: white;
            }
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #e53935;
            }
            QPushButton:pressed {
                background-color: #d32f2f;
            }
        """
        self.copy_button.setStyleSheet(button_style)
        self.move_button.setStyleSheet(button_style)
        self.open_button.setStyleSheet(button_style)
        self.delete_button.setStyleSheet(button_style)
        self.create_folder_button.setStyleSheet(button_style)
        self.browse_button.setStyleSheet(button_style)
        self.search_button.setStyleSheet(button_style)

    def browse_location(self):
        """Open a dialog to select a directory and update the input field."""
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.location_input.setText(directory)

    def search_files(self):
        """Search for files based on the given location, extension, and unique number."""
        location = self.location_input.text()
        extension = self.extension_input.text().strip()
        unique_number = self.search_input.text().strip()

        if unique_number and extension:
            results = self.search_files_in_excel_and_csv(location, extension, unique_number)
            self.file_list.clear()
            if results:
                self.file_list.addItems(results)
                self.status_bar.showMessage("Data found!", 5000)
            else:
                self.status_bar.showMessage("Data not found!", 5000)
        else:
            self.status_bar.showMessage("Please enter both a unique number and a file extension.", 5000)

    def search_files_in_excel_and_csv(self, root_dir, extension, unique_number):
        """Search for the unique number in specified files."""
        results = []
        for root, _, files in os.walk(root_dir):
            for file in files:
                if file.endswith(extension):
                    file_path = os.path.join(root, file)
                    if self.search_in_file(file_path, unique_number):
                        results.append(file_path)
        return results

    def search_in_file(self, file_path, unique_number):
        """Check if the unique number is present in the file."""
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)

            # Attempt to convert the unique_number to string for comparison
            if df.astype(str).isin([str(unique_number)]).any().any():
                return True
            return False
        except Exception as e:
            print(f"Error occurred when searching in {file_path}: {e}")
            return False
        

    def copy_file(self):
        """Copy selected files to the chosen destination."""
        selected_files = self.file_list.selectedItems()
        if selected_files:
            destination = QFileDialog.getExistingDirectory(self, "Select Destination")
            if destination:
                for item in selected_files:
                    file_path = item.text()
                    shutil.copy(file_path, os.path.join(destination, os.path.basename(file_path)))
                self.status_bar.showMessage("Files copied successfully!", 5000)

    def move_file(self):
        """Move selected files to the chosen destination."""
        selected_files = self.file_list.selectedItems()
        if selected_files:
            destination = QFileDialog.getExistingDirectory(self, "Select Destination")
            if destination:
                for item in selected_files:
                    file_path = item.text()
                    shutil.move(file_path, os.path.join(destination, os.path.basename(file_path)))
                self.status_bar.showMessage("Files moved successfully!", 5000)

    def open_file(self):
        """Open selected files."""
        selected_files = self.file_list.selectedItems()
        if selected_files:
            for item in selected_files:
                os.startfile(item.text())

    def delete_file(self):
        """Delete selected files."""
        selected_files = self.file_list.selectedItems()
        if selected_files:
            for item in selected_files:
                os.remove(item.text())
            self.file_list.clear()
            self.status_bar.showMessage("Selected files deleted!", 5000)

    def create_folder(self):
        """Create a new folder in the selected directory."""
        folder_name, ok = QInputDialog.getText(self, "Create Folder", "Enter folder name:")
        if ok and folder_name:
            try:
                os.mkdir(os.path.join(self.location_input.text(), folder_name))
                self.status_bar.showMessage("Folder created successfully!", 5000)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to create folder: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    file_manager = FileManager()
    file_manager.show()
    sys.exit(app.exec_())
