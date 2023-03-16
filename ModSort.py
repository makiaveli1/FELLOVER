import os
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QFileDialog, QLabel, QMessageBox, QProgressBar, QPushButton,
                             QVBoxLayout, QWidget, QCheckBox, QHBoxLayout, QLineEdit, QDialog, QDialogButtonBox, QGroupBox)
from PyQt5.QtCore import Qt
import re
import shutil
import zipfile
import rarfile
from pyunpack import Archive
from datetime import datetime
import fnmatch
import json
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QTreeView, QTreeWidgetItem

# Add a class for handling custom criteria


class CustomCriteriaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Criteria")

        layout = QVBoxLayout()

        self.name_label = QLabel("Archistack")
        layout.addWidget(self.name_label)

        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)

        self.extensions_label = QLabel("Extensions (comma separated):")
        layout.addWidget(self.extensions_label)

        self.extensions_input = QLineEdit()
        layout.addWidget(self.extensions_input)

        self.pattern_label = QLabel("Filename pattern (wildcards/regex):")
        layout.addWidget(self.pattern_label)

        self.pattern_input = QLineEdit()
        layout.addWidget(self.pattern_input)

        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        example_label = QLabel(
            "Example: Name: My Criterion, Extensions: .txt, .log, Pattern: *log*")
        layout.addWidget(example_label)

        self.destination_folders = []

        self.setLayout(layout)

    def get_values(self):
        return self.name_input.text(), [ext.strip() for ext in self.extensions_input.text().split(',')], self.pattern_input.text()


class Extractor(QWidget):
    def __init__(self):
        super().__init__()
        # Define the criteria attribute
        self.criteria = {}

        self.setWindowTitle("ArchiStack")
        self.setFixedSize(500, 300)

        # Set the font
        font = QFont("Noto Sans")
        self.setFont(font)

        self.setWindowIcon(
            QIcon("C: \\Users\\likwi\\OneDrive\\Pictures\\brand images\\ARCHISTACK LOGO.png"))

        # Set color palette
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#4e4d42"))
        palette.setColor(QPalette.Button, QColor("#c7d0b6"))
        palette.setColor(QPalette.WindowText, QColor("#ffffff"))
        self.setPalette(palette)

        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        # Add title label
        title_label = QLabel("ArchiStack - YOUR MOD FOLDER SOLUTION")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Noto Sans", 14, QFont.Bold))
        layout.addWidget(title_label)

        self.criteria_description_label = QLabel()
        layout.addWidget(self.criteria_description_label)

        # Initialize custom criteria
        self.custom_criteria = {}  # Add this line
        self.load_custom_criteria()

        self.destination_folders = []
        self.categories = {'Audio': {
            'extensions': ['*.mp3', '*.wav', '*.ogg', '*.flac', '*.m4a', '*.aac', '*.wma'],
            'pattern': 'audio*'
        },
            'Images': {
            'extensions': ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tiff', '*.gif', '*.webp', '*.ico'],
            'pattern': 'image*'
        },
            'Documents': {
            'extensions': ['*.pdf', '*.txt', '*.doc', '*.docx', '*.xls', '*.xlsx', '*.ppt', '*.pptx', '*.odt', '*.ods', '*.odp', '*.rtf', '*.csv'],
            'pattern': 'doc*'
        },
            'Videos': {
            'extensions': ['*.mp4', '*.avi', '*.mkv', '*.mov', '*.wmv', '*.flv', '*.m4v', '*.h264'],
            'pattern': 'video*'
        },
            'Compressed': {
            'extensions': ['*.zip', '*.rar', '*.7z', '*.tar', '*.gz', '*.bz2', '*.xz', '*.iso'],
            'pattern': 'compressed*'
        },
            'Code': {
            'extensions': ['*.py', '*.js', '*.html', '*.css', '*.php', '*.java', '*.cpp', '*.c', '*.cs', '*.sh', '*.bat', '*.swift', '*.go', '*.rb'],
            'pattern': 'code*'
        },
            'Database': {
            'extensions': ['*.db', '*.sql', '*.sqlite', '*.sqlite3', '*.mdb', '*.accdb'],
            'pattern': 'database*'
        },
            'Fonts': {
            'extensions': ['*.ttf', '*.otf', '*.woff', '*.woff2', '*.eot', '*.fon'],
            'pattern': 'font*'
        },
            'Executable': {
            'extensions': ['*.exe', '*.app', '*.bin', '*.msi', '*.dmg', '*.apk', '*.ipa'],
            'pattern': 'executable*'
        },
            'Sims 4 Mods': {
            'extensions': ['*.ts4script', '*.package', '*.trayitem', '*.blueprint', '*.room', '*.householdbinary', '*.bpi'],
            'pattern': 'sims4mod*'
        }
        }
        main_layout = QVBoxLayout()

        # Initialize treeview
        tree_view = self.init_tree_view()
        layout.addWidget(tree_view)

        self.tree_view.clicked.connect(self.update_criteria_description)

        # Group related elements
        sort_group = QGroupBox("Sort Files")
        sort_group_layout = QVBoxLayout()
        sort_group_layout.setSpacing(5)

        # Add checkboxes for sorting criteria
        checkboxes_layout = self.init_checkboxes()
        sort_group_layout.addLayout(checkboxes_layout)

        self.sort_button = QPushButton("Sort Files")
        self.sort_button.setToolTip(
            "Sort files in a selected folder based on specific criteria")
        self.sort_button.clicked.connect(self.sort_files)
        sort_group_layout.addWidget(self.sort_button)

        sort_group.setLayout(sort_group_layout)
        main_layout.addWidget(sort_group)

        # Edit and remove custom criteria buttons
        criteria_buttons_layout = QHBoxLayout()
        self.edit_custom_criteria_button = QPushButton(
            "Edit Custom Criteria")
        self.edit_custom_criteria_button.clicked.connect(
            self.edit_custom_criteria)
        criteria_buttons_layout.addWidget(self.edit_custom_criteria_button)

        self.remove_custom_criteria_button = QPushButton(
            "Remove Custom Criteria")
        self.remove_custom_criteria_button.clicked.connect(
            self.remove_custom_criteria)
        criteria_buttons_layout.addWidget(
            self.remove_custom_criteria_button)

        main_layout.addLayout(criteria_buttons_layout)

        # Extract, undo, and add criterion buttons
        action_buttons_layout = QHBoxLayout()

        self.extract_button = QPushButton("Extract Files")
        self.extract_button.setToolTip(
            "Extract selected archive files to a destination folder")
        self.extract_button.clicked.connect(self.extract_files)
        action_buttons_layout.addWidget(self.extract_button)

        self.undo_button = QPushButton("Undo Last Action")
        self.undo_button.setEnabled(False)
        self.undo_button.clicked.connect(self.undo_process)
        action_buttons_layout.addWidget(self.undo_button)

        self.undo_specific_button = QPushButton("Undo Specific Folder")
        self.undo_specific_button.setEnabled(False)
        self.undo_specific_button.clicked.connect(self.undo_process)
        action_buttons_layout.addWidget(self.undo_specific_button)

        self.add_criterion_button = QPushButton("Add Criterion")
        self.add_criterion_button.clicked.connect(
            self.add_custom_criterion)
        action_buttons_layout.addWidget(self.add_criterion_button)

        main_layout.addLayout(action_buttons_layout)

        self.checkboxes_layout = QVBoxLayout()
        self.init_checkboxes()

        # Initialize status label
        self.status_label = QLabel()
        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)
        self.extracted_files = []
        self.show()

    def init_checkboxes(self):
        """Initialize checkboxes for criteria"""
        for i in reversed(range(self.checkboxes_layout.count())):
            widget = self.checkboxes_layout.takeAt(i).widget()
            if widget:
                widget.setParent(None)

        criteria = {
            'Audio': {
                'extensions': ['*.mp3', '*.wav', '*.ogg', '*.flac', '*.m4a', '*.aac', '*.wma'],
                'pattern': 'audio*'
            },
            'Images': {
                'extensions': ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tiff', '*.gif', '*.webp', '*.ico'],
                'pattern': 'image*'
            },
            'Documents': {
                'extensions': ['*.pdf', '*.txt', '*.doc', '*.docx', '*.xls', '*.xlsx', '*.ppt', '*.pptx', '*.odt', '*.ods', '*.odp', '*.rtf', '*.csv'],
                'pattern': 'doc*'
            },
            'Videos': {
                'extensions': ['*.mp4', '*.avi', '*.mkv', '*.mov', '*.wmv', '*.flv', '*.m4v', '*.h264'],
                'pattern': 'video*'
            },
            'Compressed': {
                'extensions': ['*.zip', '*.rar', '*.7z', '*.tar', '*.gz', '*.bz2', '*.xz', '*.iso'],
                'pattern': 'compressed*'
            },
            'Code': {
                'extensions': ['*.py', '*.js', '*.html', '*.css', '*.php', '*.java', '*.cpp', '*.c', '*.cs', '*.sh', '*.bat', '*.swift', '*.go', '*.rb'],
                'pattern': 'code*'
            },
            'Database': {
                'extensions': ['*.db', '*.sql', '*.sqlite', '*.sqlite3', '*.mdb', '*.accdb'],
                'pattern': 'database*'
            },
            'Fonts': {
                'extensions': ['*.ttf', '*.otf', '*.woff', '*.woff2', '*.eot', '*.fon'],
                'pattern': 'font*'
            },
            'Executable': {
                'extensions': ['*.exe', '*.app', '*.bin', '*.msi', '*.dmg', '*.apk', '*.ipa'],
                'pattern': 'executable*'
            },
            'Sims 4 Mods': {
                'extensions': ['*.ts4script', '*.package', '*.trayitem', '*.blueprint', '*.room', '*.householdbinary', '*.bpi'],
                'pattern': 'sims4mod*'
            },
            'Custom Content': {
                'extensions': ['*.package'],
                'pattern': 'cc_*'
            },
            'Script Mods': {
                'extensions': ['*.ts4script'],
                'pattern': 'scriptmod_*'
            },
            'Build Mode Objects': {
                'extensions': ['*.package'],
                'pattern': 'buildmode_*'
            },
        }
        # Add checkboxes
        for key in self.criteria:
            checkbox = QCheckBox(key)
            self.criteria_checkboxes.append(checkbox)
            self.checkboxes_layout.addWidget(checkbox)

        # Initialize treeview
        tree_view = self.init_tree_view()
        layout.addWidget(tree_view)

        self.tree_view.clicked.connect(self.update_criteria_description)
        # Group related elements
        sort_group = QGroupBox("Sort Files")
        sort_group_layout = QVBoxLayout()
        sort_group_layout.setSpacing(5)

        # Add checkboxes for sorting criteria
        checkboxes_layout = self.init_checkboxes()
        sort_group_layout.addLayout(checkboxes_layout)

        self.sort_button = QPushButton("Sort Files")
        self.sort_button.setToolTip(
            "Sort files in a selected folder based on specific criteria")
        self.sort_button.clicked.connect(self.sort_files)
        sort_group_layout.addWidget(self.sort_button)

        sort_group.setLayout(sort_group_layout)
        layout.addWidget(sort_group)

        self.edit_custom_criteria_button = QPushButton(
            "Edit Custom Criteria")
        self.edit_custom_criteria_button.clicked.connect(
            self.edit_custom_criteria)
        layout.addWidget(self.edit_custom_criteria_button)

        self.remove_custom_criteria_button = QPushButton(
            "Remove Custom Criteria")
        self.remove_custom_criteria_button.clicked.connect(
            self.remove_custom_criteria)
        layout.addWidget(self.remove_custom_criteria_button)

        # Add Extract button
        self.extract_button = QPushButton("Extract Files")
        self.extract_button.setToolTip(
            "Extract selected archive files to a destination folder")
        self.extract_button.clicked.connect(self.extract_files)
        layout.addWidget(self.extract_button)

        # Add Undo button
        self.undo_button = QPushButton("Undo Last Action")
        self.undo_button.setEnabled(False)
        self.undo_button.clicked.connect(self.undo_process)
        layout.addWidget(self.undo_button)

        # Add Undo Specific Folder button
        self.undo_specific_button = QPushButton("Undo Specific Folder")
        self.undo_specific_button.setEnabled(False)
        self.undo_specific_button.clicked.connect(self.undo_process)
        layout.addWidget(self.undo_specific_button)
        self.undo_button.setEnabled(bool(self.destination_folders))

        # Initialize status label
        self.status_label = QLabel()
        layout.addWidget(self.status_label)

        # Add "Add Criterion" button
        self.add_criterion_button = QPushButton("Add Criterion")
        self.add_criterion_button.clicked.connect(
            self.add_custom_criterion)
        layout.addWidget(self.add_criterion_button)

        self.setLayout(layout)
        self.extracted_files = []
        self.show()

    def init_checkboxes(self):
        checkboxes_layout = QVBoxLayout()

        for key in self.criteria:
            checkbox = QCheckBox(key)
            self.criteria_checkboxes.append(checkbox)
            checkboxes_layout.addWidget(checkbox)

        return checkboxes_layout

    def add_criterion(self, name, extensions, pattern, description):
        self.criteria[name] = {
            "extensions": extensions,
            "pattern": pattern,
            "description": description}

    def add_custom_criterion(self):
        dialog = CustomCriteriaDialog(self)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            name, extensions, pattern, description = dialog.get_values()
            self.add_criterion(name, extensions, pattern, description)
            QMessageBox.information(self, "Criterion Added",
                                    f"Criterion '{name}' added successfully.")
            # Update checkboxes
            self.init_checkboxes()

    def find_best_group(self, file_name, groups):
        highest_score = 0
        best_group = None

        for group_name in groups:
            score = fuzz.token_set_ratio(file_name, group_name)
            if score > highest_score:
                highest_score = score
                best_group = group_name

        return best_group

    def init_tree_view(self):
        self.tree_view = QTreeView()
        self.tree_view.setHeaderHidden(True)

        self.tree_view_model = QStandardItemModel()
        self.tree_view.setModel(self.tree_view_model)

        for category, criteria in self.categories.items():
            category_item = QStandardItem(category)
            category_item.setEditable(False)

            for criterion in criteria:
                criterion_item = QStandardItem(criterion)
                criterion_item.setEditable(False)
                criterion_item.setCheckable(True)
                category_item.appendRow(criterion_item)

            self.tree_view_model.appendRow(category_item)

        return self.tree_view

    def edit_custom_criteria(self):
        # Get the selected criterion
        selected_indexes = self.tree_view.selectedIndexes()

        if not selected_indexes:
            QMessageBox.warning(self, "Invalid Selection",
                                "Please select a custom criterion to edit.")
            return

        selected_index = selected_indexes[0]
        selected_item = self.tree_view_model.itemFromIndex(selected_index)

        # Check if the selected item is a custom criterion
        if selected_item.text() in self.custom_criteria:
            dialog = CustomCriteriaDialog(self, selected_item.text())
            result = dialog.exec_()

        if result == QDialog.Accepted:
            name, extensions, pattern = dialog.get_values()
            self.custom_criteria[name] = {
                "extensions": extensions,
                "pattern": pattern
            }
            self.save_custom_criteria()
            QMessageBox.information(
                self, "Custom Criteria Edited", f"Custom criteria '{name}' edited successfully.")
        else:
            QMessageBox.warning(self, "Invalid Selection",
                                "Please select a custom criterion to edit.")

    def remove_custom_criteria(self):
        # Get the selected criterion
        selected_index = self.tree_view.selectedIndexes()[0]
        selected_item = self.tree_view_model.itemFromIndex(selected_index)

        # Check if the selected item is a custom criterion
        if selected_item.text() in self.custom_criteria:
            self.custom_criteria.pop(selected_item.text())
            self.save_custom_criteria()
            QMessageBox.information(self, "Custom Criteria Removed",
                                    f"Custom criteria '{selected_item.text()}' removed successfully.")
            selected_item.parent().removeRow(selected_item.row())
        else:
            QMessageBox.warning(self, "Invalid Selection",
                                "Please select a custom criterion to remove.")

    def update_criteria_description(self):
        selected_index = self.tree_view.selectedIndexes()[0]
        selected_item = self.tree_view_model.itemFromIndex(selected_index)

        criterion_name = selected_item.text()
        if criterion_name in self.criteria:
            description = self.criteria[criterion_name]["description"]
        elif criterion_name in self.custom_criteria:
            description = f"Custom criterion with extensions {', '.join(self.custom_criteria[criterion_name]['extensions'])} and pattern {self.custom_criteria[criterion_name]['pattern']}"
        else:
            description = ""

        self.criteria_description_label.setText(description)

    def is_supported_file(self, file_path):
        supported_extensions = ['.zip', '.tar', '.7z', '.rar']
        file_extension = os.path.splitext(file_path)[1].lower()
        return file_extension in supported_extensions

    def is_valid_directory(self, directory):
        return os.path.isdir(directory)

    def extract_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        files, _ = QFileDialog.getOpenFileNames(
            self, "Select archive files to extract", "", "All Files (*);;Zip Files (*.zip);;Tar Files (*.tar);;7z Files (*.7z);;Rar Files (*.rar)", options=options)

        valid_files = [f for f in files if self.is_supported_file(f)]

        destination = QFileDialog.getExistingDirectory(
            self, "Select folder to extract files")

        if valid_files and self.is_valid_directory(destination):
            try:
                self._extract_files(valid_files, destination)
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to extract files: {e}")
                self.status_label.setText("Error: Extraction failed.")
        else:
            if not valid_files:
                QMessageBox.warning(self, "Invalid Files",
                                    "No supported archive files were selected.")
            if not self.is_valid_directory(destination):
                QMessageBox.warning(self, "Invalid Directory",
                                    "Please select a valid destination folder.")

    def _extract_files(self, files, destination):
        for index, file in enumerate(files):
            try:
                # Try extracting using pyunpack
                Archive(file).extractall(destination)
            except Exception as e:
                print(f"Failed to extract using pyunpack: {e}")

                # Fallback to zipfile and rarfile libraries
                try:
                    if zipfile.is_zipfile(file):
                        with zipfile.ZipFile(file, 'r') as zip_ref:
                            zip_ref.extractall(destination)
                    elif rarfile.is_rarfile(file):
                        with rarfile.RarFile(file, 'r') as rar_ref:
                            rar_ref.extractall(destination)
                    else:
                        raise Exception("Unsupported file format")
                except Exception as e:
                    print(f"Failed to extract using zipfile and rarfile: {e}")

            # Update progress bar
            total_files = len(files)
            self.progress_bar.setValue(int((index + 1) / total_files * 100))

        self.status_label.setText("Files extracted successfully.")

    def add_custom_criteria(self):
        dialog = CustomCriteriaDialog(self)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            name, extensions, pattern = dialog.get_values()
            self.custom_criteria[name] = {
                "extensions": extensions,
                "pattern": pattern
            }
            self.save_custom_criteria()
            QMessageBox.information(
                self, "Custom Criteria Added", f"Custom criteria '{name}' added successfully.")
            self.extracted_files = []
            self.extracted_files.append((file, destination))

    def load_custom_criteria(self):
        try:
            with open("custom_criteria.json", "r") as file:
                self.custom_criteria = json.load(file)
        except FileNotFoundError:
            pass

    def is_matching_pattern(self, file_name, pattern):
        if "*" in pattern or "?" in pattern:
            return fnmatch.fnmatch(file_name, pattern)
        else:
            return re.match(pattern, file_name)

    def sort_files(self):
        # Get the selected criteria based on checked checkboxes
        selected_criteria = {checkbox.text(): self.criteria[checkbox.text()]
                             for checkbox in self.criteria_checkboxes if checkbox.isChecked()}

        if not selected_criteria:
            QMessageBox.warning(
                self, "No Criteria Selected", "Please select at least one sorting criterion.")
            return

        folder = QFileDialog.getExistingDirectory(
            self, "Select folder to sort")

        if self.is_valid_directory(folder):
            original_locations = {}
            for name, value in list(self.__dict__.items()):
                if isinstance(value, Location):
                    original_locations[name] = value

            # Define a priority list for sorting
            priority_list = ["Mods", "Tray", "Archives"]

            # Sort the priority list based on the selected criteria
            priority_list = [
                key for key in priority_list if key in selected_criteria]

            # Sort files into subfolders based on names
            grouped_files = {}

            for entry in Path(folder).glob('*'):
                if entry.is_file():
                    # Sort by priority, based on the selected criteria
                    for key in priority_list:
                        if any(ext in entry.suffix for ext in selected_criteria[key]):
                            group_name = key
                            break
                    else:
                        group_name = os.path.splitext(entry.name)[0]
                    # Apply custom criteria
                    for custom_name, custom_criteria in self.custom_criteria.items():
                        extensions = custom_criteria["extensions"]
                        pattern = custom_criteria["pattern"]

                        if any(ext in entry.suffix for ext in extensions) or self.is_matching_pattern(entry.name, pattern):
                            group_name = custom_name
                            break
                    # Create a new group if it doesn't exist
                    if group_name not in grouped_files:
                        grouped_files[group_name] = []

                    # Find the best group for the current file
                    best_group = self.find_best_group(
                        entry.name, grouped_files[group_name])

                    if not best_group:
                        best_group = {
                            "name": group_name,
                            "files": []
                        }
                        grouped_files[group_name].append(best_group)

                    best_group["files"].append(entry)

            # Move files to subfolders based on their group
            for group_name, groups in grouped_files.items():
                for group in groups:
                    destination_folder = os.path.join(folder, group["name"])
                    if not os.path.exists(destination_folder):
                        os.makedirs(destination_folder)

                    for file in group["files"]:
                        original_locations[str(file)] = folder
                        shutil.move(str(file), os.path.join(
                            destination_folder, file.name))

            self.status_label.setText("Files sorted successfully.")
            self.undo_button.setEnabled(True)

            self.destination_folders.append(original_locations)
        else:
            QMessageBox.critical(
                self, "Error", "Please select a valid folder to sort.")
            self.status_label.setText("Please select a valid folder to sort.")
    for file in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file)

        if os.path.isfile(file_path):
            category = self.categorize_mods(file_path)
            if category:
                destination_folder = os.path.join(output_folder, category)
                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)

                print(f"Moving file {file} to {destination_folder}")
                shutil.move(file_path, os.path.join(destination_folder, file))

    def categorize_mods(self, file_path):
        for category, criterion in self.criteria.items():
            extensions = criterion['extensions']
            pattern = criterion['pattern']
            file_name = os.path.basename(file_path)

            if fnmatch.fnmatch(file_name, pattern) and any(file_name.endswith(ext) for ext in extensions):
                return category

        return None

    def add_custom_criteria(self):
        dialog = CustomCriteriaDialog(self)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            name, extensions, pattern = dialog.get_values()
            self.custom_criteria[name] = {
                "extensions": extensions,
                "pattern": pattern
            }
            self.save_custom_criteria()
            QMessageBox.information(
                self, "Custom Criteria Added", f"Custom criteria '{name}' added successfully.")

    def load_custom_criteria(self):
        try:
            with open("custom_criteria.json", "r") as file:
                self.custom_criteria = json.load(file)
        except FileNotFoundError:
            pass

    def save_custom_criteria(self):
        with open("custom_criteria.json", "w") as file:
            json.dump(self.custom_criteria, file)

    def is_matching_pattern(self, file_name, pattern):
        if "*" in pattern or "?" in pattern:
            return fnmatch.fnmatch(file_name, pattern)
        else:
            return bool(re.match(pattern, file_name))

    def undo_process(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select folder to undo")

        if self.is_valid_directory(folder):
            files_to_undo = []
            for entry in Path(folder).glob('*'):
                if entry.is_file():
                    for extracted_file, extracted_folder in self.extracted_files:
                        if entry.name == os.path.basename(extracted_file) and extracted_folder == folder:
                            files_to_undo.append(entry)
                            if entry in self.extracted_files:
                                original_file, original_destination = self.extracted_files.pop(
                                    self.extracted_files.index(entry))
                                shutil.move(str(entry), original_file)
                                break
                    for original_locations in self.destination_folders:
                        for file_path, original_location in original_locations.items():
                            if entry.name == Path(file_path).name and Path(original_location).resolve() == Path(folder).resolve():
                                files_to_undo.append(entry)
                                if entry in self.extracted_files:
                                    original_file, original_destination = self.extracted_files.pop(
                                        self.extracted_files.index(entry))
                                    shutil.move(str(entry), original_file)
                                    break

            if not files_to_undo:
                QMessageBox.critical(
                    self, "Error", "Selected folder does not contain any files that were extracted or sorted by the app.")
                self.status_label.setText(
                    "Selected folder does not contain any files that were extracted or sorted by the app.")
                return

            try:
                for entry in files_to_undo:
                    filename = entry.name
                    new_location = os.path.join(folder, filename)
                    shutil.move(str(entry), new_location)
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to undo process: {e}")
                self.status_label.setText(f"Failed to undo process: {e}")
                return

            self._update_status("Process undone successfully.")
            self.destination_folders.pop()
        else:
            QMessageBox.critical(
                self, "Error", "Please select a valid folder to undo.")
            self.status_label.setText("Please select a valid folder to undo.")
        self.destination_folders.pop()
        if not self.destination_folders:
            self.undo_specific_button.setEnabled(False)

    def _update_status(self, message):
        self.status_label.setText(message)
        self.progress_bar.setVisible(False)


if __name__ == '__main__':
    app = QApplication([])
    extractor = Extractor()
    app.exec_()
