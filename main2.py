"""imports"""

import configparser
import copy
import json
import random
import shutil
import sys
import os
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QInputDialog, QAction,
                             QListWidget,QProgressBar,QComboBox,QMessageBox,QDialog, QTextEdit, QCheckBox, QStyledItemDelegate,QPushButton, QLabel, QFileDialog, QSlider, QStyle, QTableWidgetItem, QSpinBox, QHeaderView, QTableWidget)
from PyQt5.QtCore import Qt, QUrl, QSize, QProcess, QTimer, QThread, pyqtSignal, QEvent
from PyQt5.QtGui import QIcon, QFont, QPixmap,QTextCursor, QCursor, QTransform, QColor, QBrush
from PIL import Image, ImageFilter
from openpyxl import Workbook
import openpyxl
import Image_resizer, premiere_export
from concurrent.futures import ThreadPoolExecutor

from EMM_THEMES.theme import set_theme

os.environ["QT_PLUGIN_PATH"] = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages", "PyQt5", "Qt", "plugins")



"""main class"""
class SlideshowCreator(QMainWindow):
    """window creation"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Slideshow Creator") # Set the window title
        self.setGeometry(100, 100, 1200, 800) # Set the window size
        
        # Initialize variables
        self.images = [] # List to store image paths and durations
        self.audio_file = "" # Path to the audio file
        self.output_file = "output.mp4" # Default output file name
        self.button_font = "Segoe UI"
        self.deafult_font = "Segoe UI"
        self.text_font = "Segoe UI"
        self.text_font_size = 10
        self.button_font_size = 9
        self.transitions_types = [
            "fade", "fadeblack", "fadewhite", "distance", 
            "wipeleft", "wiperight", "wipeup", "wipedown",
            "slideleft", "slideright", "slideup", "slidedown",
            "smoothleft", "smoothright", "smoothup", "smoothdown",
            "circlecrop", "rectcrop", "circleclose", "circleopen",
            "horzclose", "horzopen", "vertclose", "vertopen",
            "diagbl", "diagbr", "diagtl", "diagtr","zoomin",
            "hlslice", "hrslice", "vuslice", "vdslice",
            "dissolve", "pixelize", "radial", "hblur",
            "wipetl", "wipetr", "wipebl", "wipebr",
            "fadegrays", "squeezev", "squeezeh",
            "hlwind", "hrwind", "vuwind", "vdwind",
            "coverleft", "coverright", "coverup", "coverdown",
            "revealleft", "revealright", "revealup", "revealdown"
        ]

        self.default_transition_duration = 1
        self.common_width = 1920
        self.common_height = 1080
        self.images_backup = []
        self.backup_state = False
        self.premiere_project_folder = ""

        self.executor = ThreadPoolExecutor(max_workers=os.cpu_count() * 2)  # Use more threads

        self.shortcuts = {
        "save": "Ctrl+S",
        "save_as": "Ctrl+Shift+S",
        "load": "Ctrl+L",
        "easy_text": "Ctrl+T",
        "info": "Alt+I",
        "import_images": "Ctrl+Shift+I",
        "import_audio": "Ctrl+Shift+A",
        "set_image_location": "Ctrl+Q",
        "delete_row": "Delete"
        }
        self.load_shortcuts()  # Load shortcuts from file

        self.loaded_project = ""    
        
        self.create_ui()  # Create the user interface


    
    """User Interface"""
    def create_ui(self):
        # Create the main widget
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)

        # Set dark background for the main widget

        """Left Panel - Image List with Durations"""
        left_panel = QVBoxLayout()


        # Initialize the image_table attribute
        self.image_table = QTableWidget()
        self.image_table.setItemDelegate(CustomDelegate())  # Add this line
        self.image_table.setColumnCount(8)  # Increase the column count
        self.image_table.setHorizontalHeaderLabels([".", "Image", "Duration (sec)", "Transition", "Length (sec)", "Text", "Rotation (deg)", "Second Image"])
        self.image_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        self.image_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)


        self.image_table.itemChanged.connect(self.on_edit_on_table)

        slides_label = QLabel("Slides")
        slides_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))
        left_panel.addWidget(slides_label)
        left_panel.addWidget(self.image_table)


        
        """Right Panel - Audio Files + Preview"""
        self.preview_label = QLabel("Preview")
        self.preview_label.setFont(QFont(self.text_font, 16, QFont.Bold))
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(300)


        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("QProgressBar { background-color: #1E1E1E; color: white; }"
                                        "QProgressBar::chunk { background-color: #0078d4; }")
        
        self.image_progress_bar = QProgressBar()
        self.image_progress_bar.setRange(0, 100)
        self.image_progress_bar.setValue(0)
        self.image_progress_bar.setVisible(False)
        self.image_progress_bar.setStyleSheet("QProgressBar { background-color: #1E1E1E; color: white; }"
                                             "QProgressBar::chunk { background-color: #ff0000; }")  # Red for image exporting
        
        self.image_premiere_progress_bar = QProgressBar()
        self.image_premiere_progress_bar.setRange(0, 100)
        self.image_premiere_progress_bar.setValue(0)
        self.image_premiere_progress_bar.setVisible(False)
        self.image_premiere_progress_bar.setStyleSheet("QProgressBar { background-color: #1E1E1E; color: white; }"
                                             "QProgressBar::chunk { background-color: #800080; }")  # Dark purple for image exporting

        right_panel = QVBoxLayout()
        right_panel.addWidget(self.preview_label)


        
        

        self.audio_table = QTableWidget()
        self.audio_table.setColumnCount(2)
        self.audio_table.setHorizontalHeaderLabels(["Actions", "Audio File"])
        self.audio_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        self.audio_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.audio_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)


        audio_files_label = QLabel("Audio Files:")
        audio_files_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))
        right_panel.addWidget(audio_files_label)
        right_panel.addWidget(self.audio_table)
        right_panel.addWidget(self.progress_bar)
        right_panel.addWidget(self.image_progress_bar)
        right_panel.addWidget(self.image_premiere_progress_bar)


        # Add panels to main layout
        main_layout.addLayout(left_panel, 2)
        main_layout.addLayout(right_panel, 1)

        self.setCentralWidget(main_widget)

        # Initialize audio files list
        self.audio_files = []


    """01_Images Functions"""
    
    """01_01_Duration Functions"""
    def on_edit_on_table(self, item):
        """Handles editing the 'Duration' column."""
        column = item.column()
        row = item.row()
        if column == 2:  # Only handle edits for the 'Duration' column
            try:
                new_duration = int(item.text())
                if new_duration < 2 or new_duration > 600:
                    raise ValueError("Duration out of range (2-600).")
                self.images[row]['duration'] = new_duration  # Update the image data
                if self.images[row]['transition_duration'] > self.images[row]['duration'] -1:
                    self.images[row]['transition_duration'] = self.images[row]['duration'] -1
                    self.update_image_table()
            except ValueError:
                # Revert to the previous value if the input is invalid
                item.setText(str(self.images[row]['duration']))        
        elif column == 5:  # Handle edits for the 'Text' column
            try:
                new_text = str(item.text())
                self.images[row]['text'] = new_text  # Update the image data
                #self.update_image_table()
            except ValueError:
                # Revert to the previous value if the input is invalid
                item.setText(str(self.images[row]['text']))

        elif column == 6:  # Handle edits for the 'Text' column
            try:
                new_rotation = int(item.text())
                if new_rotation < 0 or new_rotation > 359:
                    raise ValueError("Duration out of range (0-359).")
                self.images[row]['rotation'] = new_rotation  # Update the image data
                self.update_preview_with_row(row)
                #self.update_image_table()
            except ValueError:
                # Revert to the previous value if the input is invalid
                item.setText(str(self.images[row]['rotation']))

            
    def set_all_images_duration(self):
        selected_items = self.image_table.selectedItems()
        row = self.image_table.row(selected_items[0]) if selected_items else None
        current_duration = self.images[row]['duration'] if row is not None else 2

        # Create an input dialog instance
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Set Duration")
        dialog.setLabelText("Enter duration in seconds:")
        dialog.setIntValue(current_duration)  # Set default duration
        dialog.setIntRange(2, 600)  # Set valid range


        # Execute the dialog
        if dialog.exec_() == QDialog.Accepted:
            new_duration = dialog.intValue()
            for i in range(len(self.images)):
                self.images[i]['duration'] = new_duration
            self.update_image_table()

    """01_02_Images Implementation Functions"""
    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)")
        if files:
            new_images = [{'path': file, 'duration': 5, 'transition': 'fade', 'transition_duration': self.default_transition_duration, 'text': "", 'rotation': 0, "is_second_image": False} for file in files]
            self.images.extend(new_images)
            self.update_image_table()  # Single update after all images are added

    def set_second_image(self, row, state):
        if row == 0:
            QMessageBox.critical(self, "Error", "Cannot set second image for the first row", QMessageBox.Ok)
            self.update_image_row(row)
            return
        else:
            self.images[row]['is_second_image'] = state == Qt.Checked
            filename_item = self.image_table.item(row, 1)
            if filename_item:
                filename_item.setData(Qt.UserRole, state == Qt.Checked)  # Update the UserRole data
            self.update_image_row(row)

    def update_image_table(self):
        self.image_table.blockSignals(True)  # Disable signals
        self.image_table.setUpdatesEnabled(False)
        self.image_table.setSortingEnabled(False)

        self.image_table.setRowCount(len(self.images))
        for row, img in enumerate(self.images):
            path_img = os.path.basename(img['path'])
            filename_item = QTableWidgetItem(path_img)
            filename_item.setData(Qt.UserRole, img.get('is_second_image', False))  # Set the UserRole data
            duration_item = QTableWidgetItem(str(img.get('duration', 5)))
            self.transition_item = QComboBox()
            self.transition_item.addItems(self.transitions_types)
            self.transition_item.setCurrentText(img.get('transition', 'fade'))
            transition_length_item = QTableWidgetItem(str(img.get('transition_duration', self.default_transition_duration)))
            text_item = QTableWidgetItem(str(img.get('text', "")))
            text_item.setFlags(text_item.flags() | Qt.ItemIsEditable)
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            transition_length_item.setFlags(transition_length_item.flags() & ~Qt.ItemIsEditable)
            rotation_item = QTableWidgetItem(str(img.get('rotation', 0)))



            # Add a checkbox for "Second Image"
            second_image_checkbox = QCheckBox()
            second_image_checkbox.setChecked(img.get('is_second_image', False))
            second_image_checkbox.stateChanged.connect(lambda state, row=row: self.set_second_image(row, state))

            self.image_table.setItem(row, 1, filename_item)
            self.image_table.setItem(row, 2, duration_item)
            self.image_table.setCellWidget(row, 3, self.transition_item)
            self.image_table.setItem(row, 4, transition_length_item)
            self.image_table.setItem(row, 5, text_item)
            self.image_table.setItem(row, 6, rotation_item)
            self.image_table.setCellWidget(row, 7, second_image_checkbox)

            # Add move up, move down, and delete buttons
            move_up_btn = QPushButton("↑")
            # Inside update_image_table:
            move_up_btn.clicked.connect(self.move_image_up)


            move_down_btn = QPushButton("↓")
            move_down_btn.clicked.connect(self.move_image_down)


            delete_btn = QPushButton("✖")
            delete_btn.clicked.connect(self.delete_image)

            button_widget = QWidget()
            button_layout = QHBoxLayout(button_widget)
            button_layout.addWidget(move_up_btn)
            button_layout.addWidget(move_down_btn)
            button_layout.addWidget(delete_btn)
            button_layout.setContentsMargins(0, 0, 0, 0)
            button_widget.setLayout(button_layout)

            self.image_table.setCellWidget(row, 0, button_widget)

            # Connect the QComboBox signal to update the transition in self.images
            self.transition_item.currentTextChanged.connect(lambda text, row=row: self.update_transition(row, text))

        self.image_table.setSortingEnabled(True)
        self.image_table.setUpdatesEnabled(True)
        self.image_table.blockSignals(False)  # Re-enable signals

    def move_image_up(self):
        self.image_table.setSortingEnabled(False)  # Disable sorting
        row = self.image_table.currentRow()
        if row > 0:
            self.images[row], self.images[row - 1] = self.images[row - 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row - 1)
            self.image_table.setCurrentCell(row - 1, 1)
            self.update_preview_with_row(row - 1)
        self.image_table.setSortingEnabled(True)  # Re-enable sorting

    def move_image_down(self):
        self.image_table.setSortingEnabled(False)  # Disable sorting
        row = self.image_table.currentRow()
        if row < len(self.images) - 1:
            self.images[row], self.images[row + 1] = self.images[row + 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row + 1)
            self.image_table.setCurrentCell(row + 1, 1)
            self.update_preview_with_row(row + 1)
        self.image_table.setSortingEnabled(True)  # Re-enable sorting

    def update_image_row(self, row):
        """Update a single row in the image table."""
        if 0 <= row < len(self.images):  # Ensure the row is within bounds
            img = self.images[row]
            path_img = os.path.basename(img['path'])
            duration_item = QTableWidgetItem(str(img.get('duration', 5)))
            transition_length_item = QTableWidgetItem(str(img.get('transition_duration', self.default_transition_duration)))
            text_item = QTableWidgetItem(str(img.get('text', "")))
            text_item.setFlags(text_item.flags() | Qt.ItemIsEditable)
            rotation_item = QTableWidgetItem(str(img.get('rotation', 0)))

            filename_item = QTableWidgetItem(path_img)
            filename_item.setData(Qt.UserRole, img.get('is_second_image', False))  # Add this line
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            transition_length_item.setFlags(transition_length_item.flags() & ~Qt.ItemIsEditable)

            # Add a checkbox for "Second Image"
            second_image_checkbox = QCheckBox()
            second_image_checkbox.setChecked(img.get('is_second_image', False))
            second_image_checkbox.stateChanged.connect(lambda state, row=row: self.set_second_image(row, state))

            """if img.get('is_second_image', False):
                color = QColor(100, 100, 150)  # Darker blue background
                filename_item.setBackground(color)
                duration_item.setBackground(color)
                transition_length_item.setBackground(color)
                text_item.setBackground(color)
                rotation_item.setBackground(color)"""

            self.image_table.setItem(row, 1, filename_item)
            self.image_table.setItem(row, 2, duration_item)
            self.image_table.setItem(row, 4, transition_length_item)
            self.image_table.setItem(row, 5, text_item)
            self.image_table.setItem(row, 6, rotation_item)
            self.image_table.setCellWidget(row, 7, second_image_checkbox)


        

    def delete_image(self):
        # Temporarily disable sorting to ensure correct row indices
        self.image_table.setSortingEnabled(False)

        row = self.image_table.currentRow()
        
        if 0 <= row < len(self.images):  # Ensure the row is within bounds
            del self.images[row]  # Remove the image from the list
            self.image_table.removeRow(row)  # Remove the row from the table
            
            # Update the preview and selection
            if len(self.images) == 0:
                self.preview_label.clear()  # Clear the preview if no images are left
            else:
                # Set the new row to focus on (either the previous row or the next row)
                new_row = max(0, row - 1) if row > 0 else 0
                self.image_table.setCurrentCell(new_row, 1)  # Set focus on the new row
                self.update_preview_with_row(new_row)  # Update the preview
                
        # Re-enable sorting after the operation
        self.image_table.setSortingEnabled(True)


    def set_random_images_order(self):
        random.shuffle(self.images)
        self.update_image_table()

    def update_image_progress(self, value):
        self.image_progress_bar.setValue(value)

    def on_image_processing_finished(self):
        self.image_progress_bar.setVisible(False)
        self.continue_with_video_export()

    def set_image_location(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            current_row = self.image_table.row(selected_items[0])

            # Create the QInputDialog manually
            dialog = QInputDialog(self)

            # Configure dialog properties
            dialog.setWindowTitle("Set Image Location")
            dialog.setLabelText("Enter new position (1-based index):")
            dialog.setInputMode(QInputDialog.IntInput)
            dialog.setIntRange(1, len(self.images))  # Set valid range
            dialog.setIntValue(current_row + 1)  # Default value

            # Show dialog and get user input
            if dialog.exec_() == QDialog.Accepted:
                new_position = dialog.intValue() - 1  # Convert to 0-based index
                image = self.images.pop(current_row)  # Remove the image from the current position
                self.images.insert(new_position, image)  # Insert it at the new position
                self.update_image_table()  # Refresh the table
                self.image_table.setCurrentCell(new_position, 1)  # Set focus on the moved image



    def update_selection_after_operation(self, new_row):
        self.image_table.setCurrentCell(new_row, 1)  # Set focus on the new row
        self.update_preview_with_row(new_row)  # Update the preview






    """02_Audio Functions"""
    def add_audio(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Audio", "", "Audio Files (*.mp3 *.wav *.flac)")
        if file:
            self.audio_files.append({'path': file})  # Default duration 0 or set as needed
            self.update_audio_table()
            print("Audio added:", self.audio_files)  # Debugging output

    def update_audio_table(self):
        self.audio_table.setRowCount(len(self.audio_files))
        for row, audio in enumerate(self.audio_files):
            filename = os.path.basename(audio['path'])
            filename_item = QTableWidgetItem(filename)
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)  # Make the item non-editable
            self.audio_table.setItem(row, 1, filename_item)

            # Add move up, move down, and delete buttons
            move_up_btn = QPushButton("↑")
            move_up_btn.clicked.connect(lambda _, r=row: self.move_audio_up(r))

            move_down_btn = QPushButton("↓")
            move_down_btn.clicked.connect(lambda _, r=row: self.move_audio_down(r))

            delete_btn = QPushButton("✖")
            delete_btn.clicked.connect(lambda _, r=row: self.delete_audio(r))

            # Create a widget to hold the buttons
            button_widget = QWidget()
            button_layout = QHBoxLayout(button_widget)
            button_layout.addWidget(move_up_btn)
            button_layout.addWidget(move_down_btn)
            button_layout.addWidget(delete_btn)
            button_layout.setContentsMargins(0, 0, 0, 0)
            button_widget.setLayout(button_layout)

            # Set the button widget in the table
            self.audio_table.setCellWidget(row, 0, button_widget)

    def move_audio_up(self, row):
        if row > 0:
            self.audio_files[row], self.audio_files[row - 1] = self.audio_files[row - 1], self.audio_files[row]
            self.update_audio_table()
            self.audio_table.setCurrentCell(row -1, 1)

    def move_audio_down(self, row):
        if row < len(self.audio_files) - 1:
            self.audio_files[row], self.audio_files[row + 1] = self.audio_files[row + 1], self.audio_files[row]
            self.update_audio_table()
            self.audio_table.setCurrentCell(row +1, 1)

    def delete_audio(self, row):
        del self.audio_files[row]
        self.update_audio_table()
        if row - 1 <0:
            self.audio_table.setCurrentCell(row +1, 1)
        else:
            self.audio_table.setCurrentCell(row -1, 1)




    """03_Export Functions"""

    def continue_with_video_export(self):
        total_image_duration = 0
        # Get the total image duration
        for i in range(len(self.images)):
            total_image_duration += self.images[i]['duration']

        # Get the total audio duration
        total_audio_duration = 0.0
        for audio in self.audio_files:
            audio_path = audio['path']
            try:
                # Enclose the file path in double quotes
                cmd = [
                    "ffprobe",
                    "-v", "error",
                    "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1",
                    f'"{audio_path}"'  # Quoting the path
                ]
                # Run the ffprobe command
                output = subprocess.check_output(" ".join(cmd), shell=True, universal_newlines=True).strip()
                total_audio_duration += float(output)
            except Exception as e:
                print(f"Error processing file {audio_path}: {e}")

        print("Total image duration:", total_image_duration, "seconds")
        print("Total audio duration:", total_audio_duration, "seconds")

        if total_image_duration > total_audio_duration:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Audio and Video doesn't match")
            msg.setText("The total image duration is bigger than the audio duration. Would you like to change the audio duration to match the image duration?")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            reply = msg.exec_()
            if reply == QMessageBox.Cancel:
                print("Export canceled by user.")
                return
            else:
                total_images = len(self.images)
                new_duration_to_each_image = int(total_audio_duration / total_images)
                if new_duration_to_each_image < 2 or new_duration_to_each_image > 600:
                    msg = QMessageBox()
                    msg.setIcon(QMessageBox.Information)
                    msg.setWindowTitle("Audio and Video doesn't match")
                    msg.setText("Sorry! \n Couldn't match the audio duration to the image duration because of min and max issues. Please add another song or choose a longer one!")
                    msg.setStandardButtons(QMessageBox.Ok)
                    msg.exec_()
                    return
                else:
                    for i in range(len(self.images)):
                        self.images[i]['duration'] = new_duration_to_each_image
                    self.update_image_table()

        # Create FFmpeg command
        command = self.build_ffmpeg_command()
        print("Exporting with command:", command)

        # Execute FFmpeg command
        self.process = QProcess(self)
        self.process.readyReadStandardOutput.connect(self.update_progress)
        self.process.readyReadStandardError.connect(self.update_progress)  # Capture FFmpeg logs
        self.process.finished.connect(self.export_finished)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)  # Reset progress bar
        self.process.start(command)

    def export_slideshow(self):
        if not self.validate_transitions():
            return
        print("Images:", self.images)  # Debugging output
        print("Audio Files:", self.audio_files)  # Debugging output
        if not self.images or not self.audio_files:
            QMessageBox.critical(self, "Error", "Please add images and audio before exporting.", QMessageBox.Ok)
            return
        
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Slideshow", "", "Video Files (*.mp4);;All Files (*)", options=options)
        if file_path:
            self.output_file = file_path
        else:
            QMessageBox.critical(self, "Error", "Please select a location for the exported video.", QMessageBox.Ok)
            return
        
        # Start image processing in a separate thread
        self.image_progress_bar.setVisible(True)
        self.image_progress_bar.setValue(0)

        # Assuming self.images is a list of dictionaries with 'path' and 'duration' as keys
        image_path22 = self.images[0]['path']
        output_folder22 = os.path.join(os.path.dirname(image_path22), "A_Blur")  # Ensure correct folder structure

        self.images_backup = copy.deepcopy(self.images)
        self.backup_state = True

        self.image_worker = ImageProcessingWorker(self.images, output_folder22, self.common_width, self.common_height)
        self.image_worker.progress.connect(self.update_image_progress)
        self.image_worker.finished.connect(self.on_image_processing_finished)
        self.image_worker.start()

    def export_finished(self):
        self.progress_bar.setValue(100)  # Mark as complete
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Export Complete")
        msg.setText("Export finished successfully!")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

        self.progress_bar.setVisible(False)    

    def build_ffmpeg_command(self):
        inputs = []
        filters = []
        total_duration = 0

        common_width = self.common_width
        common_height = self.common_height

        

        #transition_duration = self.transition_duration
            # Add image inputs and scaling filters
        for i, img in enumerate(self.images):
            if i == 0:
                duration = img['duration']
            elif i == len(self.images) - 1:
                duration = img['duration'] + self.images[i-1]['transition_duration']
            else:
                duration = img['duration'] + img['transition_duration']
            inputs.append(f'-loop 1 -t {duration} -i "{img["path"]}"')
            filters.append(f"[{i}:v]scale=1920:1080,setsar=1,format=yuv420p[{i}v]")
            total_duration += duration

        # Add xfade transitions
        
        #transition_type = self.transition_type

        for i in range(len(self.images) - 1):
            # Calculate the offset where the transition starts
            start_offset = sum(img['duration'] for img in self.images[:i + 1]) - self.images[i]['transition_duration']
            if i == 0:
                # Connect the first two video streams
                filters.append(f"[{i}v][{i+1}v]xfade=transition={self.images[i]['transition']}:duration={self.images[i]['transition_duration']}:offset={start_offset}[v{i+1}]")
            else:
                # Chain subsequent xfade transitions
                filters.append(f"[v{i}][{i+1}v]xfade=transition={self.images[i]['transition']}:duration={self.images[i]['transition_duration']}:offset={start_offset}[v{i+1}]")

        # Final video stream
        final_stream = f"[v{len(self.images) - 1}]"

        # Add audio inputs
        audio_inputs = []
        audio_streams = []
        audio_index = len(self.images)

        for i, audio in enumerate(self.audio_files):
            inputs.append(f'-i "{audio["path"]}"')
            audio_streams.append(f"[{audio_index + i}:a]")

        # Concatenate audio files sequentially
        if len(audio_streams) > 1:
            filters.append(f"{''.join(audio_streams)}concat=n={len(audio_streams)}:v=0:a=1[outa]")
            audio_map = "-map [outa]"
        else:
            audio_map = f"-map {audio_index}:a"

        # Build the filter_complex
        filter_complex = ";".join(filters)

        # Construct the FFmpeg command
        command = f'ffmpeg -y {" ".join(inputs)} -filter_complex "{filter_complex}" -map {final_stream} {audio_map} -c:a aac -c:v libx264 -pix_fmt yuv420p -preset ultrafast -shortest "{self.output_file}"'

        if self.backup_state == True:
            self.images = self.images_backup
            self.backup_state = False
            self.update_image_table()

        return command
    




    """04_Progress Functions"""
    def update_progress(self):
        output = self.process.readAllStandardError().data().decode("utf-8")  # FFmpeg logs
        print(output)

        # Extract progress percentage from FFmpeg logs
        for line in output.split("\n"):
            if "time=" in line:
                time_str = line.split("time=")[1].split(" ")[0]
                time_parts = time_str.split(":")
                if len(time_parts) == 3:
                    hours, minutes, seconds = map(float, time_parts)
                    current_time = hours * 3600 + minutes * 60 + seconds

                    # Estimate progress percentage
                    total_duration = sum(img['duration'] for img in self.images)
                    progress = int((current_time / total_duration) * 100)
                    self.progress_bar.setValue(progress)



    """05_Preview Functions"""
    def update_preview(self):
        """Update preview when a slide is selected"""
        selected_items = self.image_table.selectedItems()
        if selected_items and self.images:  # Check if images list is not empty
            row = self.image_table.row(selected_items[0])
            if 0 <= row < len(self.images):  # Ensure the row is valid
                img_data = self.images[row]
                img_path = img_data['path']
                rotation = img_data.get('rotation', 0)  # Default to 0 if 'rotation' is not set
                
                # Load the image
                pixmap = QPixmap(img_path)
                if rotation != 0:
                    # Apply rotation
                    transform = QTransform()
                    transform.rotate(rotation)
                    pixmap = pixmap.transformed(transform, Qt.SmoothTransformation)
                
                # Display the rotated image in the preview
                self.preview_label.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio))


    def update_preview_with_row(self, row):
        """Update preview when a slide is selected"""
        if 0 <= row < len(self.images):  # Ensure the row is valid
            img_data = self.images[row]
            img_path = img_data['path']
            rotation = img_data.get('rotation', 0)  # Default to 0 if 'rotation' is not set
            
            # Load the image
            pixmap = QPixmap(img_path)
            if rotation != 0:
                # Apply rotation
                transform = QTransform()
                transform.rotate(rotation)
                pixmap = pixmap.transformed(transform, Qt.SmoothTransformation)
            
            # Display the rotated image in the preview
            self.preview_label.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio))


    def setup_connections(self):
        """Connect signals to slots"""
        # Connect table selection changes to preview updates
        self.image_table.itemSelectionChanged.connect(self.update_preview)



    """06_Project Functions"""
    def clear_project(self):
        self.images.clear()
        self.audio_files.clear()  # Clear audio files
        self.image_table.setRowCount(0)  # Proper way to clear the table
        self.audio_table.setRowCount(0)  # Clear audio table
        self.preview_label.clear()

        self.loaded_project = ""

    def save_project(self):
        if self.loaded_project != "":
            with open(self.loaded_project, 'w', encoding='utf-8') as f:
                count = 0
                for audio in self.audio_files:
                    count += 1
                f.write(str(count) + "\n")
                for audio in self.audio_files:
                    f.write(f"{audio['path']}\n")  
                for img in self.images:
                    text = img.get('text', '')
                    if '\n' in text:
                        text = text.replace('\n', '\\n')
                    f.write(f"{img['path']},{img.get('duration', 5)},{img.get('transition', 'fade')},{img.get('transition_duration', 1)},{text},{img.get('rotation', '')},{img.get('is_second_image', False)}\n") 
        
        else:
            options = QFileDialog.Options()
            file_name, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
            if file_name:
                with open(file_name, 'w', encoding='utf-8') as f:
                    count = 0
                    for audio in self.audio_files:
                        count += 1
                    f.write(str(count) + "\n")
                    for audio in self.audio_files:
                        f.write(f"{audio['path']}\n")  
                    for img in self.images:
                        text = img.get('text', '')
                        if '\n' in text:
                            text = text.replace('\n', '\\n')
                        f.write(f"{img['path']},{img.get('duration', 5)},{img.get('transition', 'fade')},{img.get('transition_duration', 1)},{text},{img.get('rotation', '')},{img.get('is_second_image', False)}\n") 
                    self.loaded_project = file_name


    def save_project_as(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'w', encoding='utf-8') as f:
                count = 0
                for audio in self.audio_files:
                    count += 1
                f.write(str(count) + "\n")
                for audio in self.audio_files:
                    f.write(f"{audio['path']}\n")  
                for img in self.images:
                    text = img.get('text', '')
                    if '\n' in text:
                        text = text.replace('\n', '\\n')
                    f.write(f"{img['path']},{img.get('duration', 5)},{img.get('transition', 'fade')},{img.get('transition_duration', 1)},{text},{img.get('rotation', '')},{img.get('is_second_image', False)}\n") 
                self.loaded_project = file_name

    def load_project(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                count = int(lines[0].strip())  # Number of audio files
                
                self.audio_files = []  # Clear existing audio files
                
                # Load audio files
                for i in range(1, count + 1):
                    audio_path = lines[i].strip()
                    self.audio_files.append({'path': audio_path})  # Add each audio file to the list
                
                # Load images
                self.images = []
                for line in lines[count + 1:]:  # Start after the audio file lines
                    path, duration, transition, transition_duration, text, rotation, is_second_image = line.strip().split(',')
                    if transition_duration != self.default_transition_duration:
                        transition_duration = self.default_transition_duration
                    self.images.append({
                        'path': path,
                        'duration': int(duration),
                        'transition': transition,
                        'transition_duration': int(transition_duration),
                        'text': text.replace('\\n', '\n'),
                        'rotation': int(rotation),
                        'is_second_image': is_second_image.strip().lower() == 'true'  # Convert to boolea
                    })

                print(self.images)
                
                # Update the UI tables

                self.update_image_table()
                self.update_audio_table()

                self.loaded_project = file_name



    """07_Transition Functions"""
    def update_transition(self, row, transition):
        self.images[row]['transition'] = transition
    
    def set_all_images_transition(self):
        # Create an input dialog instance
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Set Transition")
        dialog.setLabelText("Select transition:")
        dialog.setComboBoxItems(self.transitions_types)  # Set transition options


        # Execute the dialog
        if dialog.exec_() == QDialog.Accepted:
            transition = dialog.textValue()
            for i in range(len(self.images)):
                self.images[i]['transition'] = transition
                self.transition_item.setCurrentText(transition)  # Set current transition
            self.update_image_table()


    def set_random_transition_for_each_image(self):
        
        for i in range(len(self.images)):
            random_i = random.randint(0, len(self.transitions_types) - 1)
            transition = self.transitions_types[random_i]
            self.images[i]['transition'] = transition
            self.transition_item.setCurrentText(transition)  # Set current transition
        self.update_image_table()

    def auto_calc_image_duration(self):
        total_images_count = 0
        total_audio_duration = 0.0
        for audio in self.audio_files:
            audio_path = audio['path']
            try:
                # Enclose the file path in double quotes
                cmd = [
                    "ffprobe",
                    "-v", "error",
                    "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1",
                    f'"{audio_path}"'  # Quoting the path
                ]
                # Run the ffprobe command
                output = subprocess.check_output(" ".join(cmd), shell=True, universal_newlines=True).strip()
                total_audio_duration += float(output)
            except Exception as e:
                print(f"Error processing file {audio_path}: {e}")

        for i in range(len(self.images)):
            total_images_count += 1


        new_image_duration = int((total_audio_duration - 2) / total_images_count)
        for i in range(len(self.images)):
            self.images[i]['duration'] = new_image_duration
        self.update_image_table()

    def validate_transitions(self):
        for img in self.images:
            if img['transition_duration'] >= img['duration']:
                QMessageBox.warning(self, "Invalid Transition Duration", 
                                f"Transition duration for {img['path']} is too long. It must be less than the image duration.")
                return False
        return True


    """08_Premiere_Functions"""
    def export_premiere_slideshow(self):

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Premiere Slideshow", "", "Folder", options=options)
        if file_path:
            self.premiere_project_folder = file_path
        else:
            QMessageBox.critical(self, "Error", "Please select a location for the exported Premiere project.", QMessageBox.Ok)
            return

        self.image_premiere_progress_bar.setVisible(True)
        self.image_premiere_progress_bar.setValue(0)


        self.image_premiere_worker = ImageProcessingPremiereWorker(self.images, self.premiere_project_folder, self.common_width, self.common_height)
        self.image_premiere_worker.progress.connect(self.update_image_premiere_progress)
        self.image_premiere_worker.finished.connect(self.on_image_premiere_processing_finished)
        self.image_premiere_worker.start()

    
    def update_image_premiere_progress(self, value):
        self.image_premiere_progress_bar.setValue(value)

    def on_image_premiere_processing_finished(self):
        self.image_premiere_progress_bar.setVisible(False)
        self.export_premiere_audio()
        self.export_premiere_text()
        self.export_premiere_duration_excel()
        self.copy_premiere_project_file()

    def export_premiere_audio(self):
        # Create the "Audios" folder if it doesn't exist
        premiere_audio_folder = os.path.join(self.premiere_project_folder, "02_אודיו")
        os.makedirs(premiere_audio_folder, exist_ok=True)

        for i, audio_file in enumerate(self.audio_files, start=1):
            audio_path = audio_file['path']
            audio_extension = os.path.splitext(audio_path)[1]  # Get the file extension (e.g., .mp3)
            audio_file_name = os.path.splitext(os.path.basename(audio_path))[0]  # Get the audio file name without extension
            new_audio_name = f"audio{i}_{audio_file_name}{audio_extension}"  # Generate new name (e.g., audio1_filename.mp3)
            premiere_audio_path = os.path.join(premiere_audio_folder, new_audio_name)

            # Copy the audio file to the new location with the new name
            shutil.copy(audio_path, premiere_audio_path)

            print(f"Copied {audio_path} to {premiere_audio_path}")

    def format_time(self, seconds):
        """Convert seconds to SRT time format: HH:MM:SS,mmm."""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        millis = int((seconds - int(seconds)) * 1000)
        return f"{hours:02}:{minutes:02}:{secs:02},{millis:03}"

    def export_premiere_text(self):
        # Create the "Texts" folder if it doesn't exist
        premiere_text_folder = os.path.join(self.premiere_project_folder, "03_טקסט")
        os.makedirs(premiere_text_folder, exist_ok=True)
        style_file_path = r"E:\------ תכנות ------\Even Monatge Maker 2.0\Premiere_Project\טקסט למצגת - עברית.prtextstyle"
        # Copy the original text file to the new location
        original_text_file_path = os.path.join(premiere_text_folder, "טקסט למצגת - בעברית.prtextstyle")
        shutil.copy(style_file_path, original_text_file_path)
        print(f"Copied {style_file_path} to {original_text_file_path}")
        

        # Define the output file path
        srt_file_path = os.path.join(premiere_text_folder, "exported_texts.srt")

        # Initialize time tracking
        current_time = 0
        subtitle_index = 1  # Manually track subtitle index

        # Write the SRT file
        with open(srt_file_path, "w", encoding="utf-8") as srt_file:
            for image in self.images:
                if image['is_second_image']:
                    continue  # Skip second images

                start_time = self.format_time(current_time)
                end_time = self.format_time(current_time + image['duration'])

                # Write the subtitle entry
                srt_file.write(f"{subtitle_index}\n")
                srt_file.write(f"{start_time} --> {end_time}\n")
                srt_file.write(f"{image['text']}\n\n")

                # Update index and time
                subtitle_index += 1
                current_time += image['duration']

        print(f"SRT file created at: {srt_file_path}")

    def export_premiere_duration_excel(self):
        # Create the "Texts" folder if it doesn't exist
        premiere_text_folder = os.path.join(self.premiere_project_folder, "03_טקסט")
        os.makedirs(premiere_text_folder, exist_ok=True)

        # Define the output Excel file path
        excel_file_path = os.path.join(premiere_text_folder, "exported_durations.xlsx")

        # Create a new Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Durations"

        # Write the header row
        ws.append(["Path", "Duration", "Text"])

        # Write the image data
        for image in self.images:
            ws.append([image["path"], image["duration"], image["text"]])

        # Save the workbook to the file
        wb.save(excel_file_path)

        print(f"Excel file created at: {excel_file_path}")

    def copy_premiere_project_file(self):
        # Define the source and destination paths for the Premiere project file
        premiere_project_source = r"E:\------ תכנות ------\Even Monatge Maker 2.0\Premiere_Project\Project.prproj"
        project_destination_folder = os.path.join(self.premiere_project_folder, "04_פרוייקט")
        os.makedirs(project_destination_folder, exist_ok=True)
        project_file_name = os.path.basename(self.premiere_project_folder) + ".prproj"
        project_destination_path = os.path.join(project_destination_folder, project_file_name)

        # Copy the Premiere project file
        shutil.copy(premiere_project_source, project_destination_path)

        print(f"Premiere project file copied to: {project_destination_path}")



    

    def open_easy_text_writing(self):
        if not self.images:
            QMessageBox.warning(self, "No Images", "Please add images before using Easy Text Writing.")
            return
        
        selected_row = self.image_table.currentRow()  # Replace 'self.image_table' with the name of your table widget
        affected_rows = []
        
        self.easy_text_dialog = EasyTextWritingDialog(self.images, affected_rows, start_index=selected_row, parent=self)
        

        self.easy_text_dialog.show()
        if self.easy_text_dialog.exec_():
            affected_rows[:] = self.easy_text_dialog.affected_rows
        
        for row in affected_rows:
            self.update_image_row(row)


    """09_Shortcut Functions"""

    def save_shortcuts(self):
        # Ensure the directory exists
        os.makedirs("C:\\NeriaLTD\\Event_Montage_Maker_2", exist_ok=True)
        
        # Save shortcuts to a file
        with open("C:\\NeriaLTD\\Event_Montage_Maker_2\\shortcuts.txt", "w") as f:
            for action, shortcut in self.shortcuts.items():
                f.write(f"{action}:{shortcut}\n")

    def load_shortcuts(self):
        # Load shortcuts from file
        try:
            with open("C:\\NeriaLTD\\Event_Montage_Maker_2\\shortcuts.txt", "r") as f:
                for line in f:
                    action, shortcut = line.strip().split(":")
                    self.shortcuts[action] = shortcut
        except FileNotFoundError:
            # If the file doesn't exist, use the default shortcuts
            pass

    def update_shortcuts(self):
        # Update the shortcuts in the application
        self.save_action.setShortcut(self.shortcuts.get("save", "Ctrl+S"))
        self.save_as_action.setShortcut(self.shortcuts.get("save_as", "Ctrl+Shif+S"))
        self.load_action.setShortcut(self.shortcuts.get("load", "Ctrl+L"))
        self.easy_text_writing_action.setShortcut(self.shortcuts.get("easy_text", "Ctrl+T"))
        self.show_info_action.setShortcut(self.shortcuts.get("info", "Alt+I"))
        self.import_images.setShortcut(self.shortcuts.get("import_images", "Ctrl+Shift+I"))
        self.import_audio.setShortcut(self.shortcuts.get("import_audio", "Ctrl+Shift+A"))
        self.set_image_location_action.setShortcut(self.shortcuts.get("set_image_location", "Ctrl+Q"))
        self.set_image_location_action.setShortcut(self.shortcuts.get("set_image_location", "Ctrl+Q"))

    def set_shortcut(self, action):
        # Create an input dialog instance
        dialog = QInputDialog(self)
        dialog.setWindowTitle(f"Set {action.capitalize()} Shortcut")
        dialog.setLabelText(f"Enter the new shortcut for {action}:")
        dialog.setTextValue(self.shortcuts.get(action, ""))


        # Execute the dialog
        if dialog.exec_() == QDialog.Accepted:
            shortcut = dialog.textValue()
            if shortcut:
                self.shortcuts[action] = shortcut
                self.save_shortcuts()  # Save the new shortcut
                self.update_shortcuts()  # Update the shortcuts in the application





    def show_info(self):
        
        info_dialog = InfoDialog(self.images, self.audio_files, self)
        info_dialog.exec_()








    """10_Menu Functions"""
    def create_menu(self):
        # Function to create a menu bar
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")

        import_menu = file_menu.addMenu("Import")
        
        self.import_images = QAction("Images", self)
        self.import_images.triggered.connect(self.add_images)
        self.import_images.setShortcut(self.shortcuts.get("import_images", "Ctrl+Shift+I"))
        import_menu.addAction(self.import_images)

        self.import_audio = QAction("Audio", self)
        self.import_audio.triggered.connect(self.add_audio)
        self.import_audio.setShortcut(self.shortcuts.get("import_audio", "Ctrl+Shift+A"))
        import_menu.addAction(self.import_audio)

        self.load_action = QAction("Load Project", self)
        self.load_action.triggered.connect(self.load_project)
        self.load_action.setShortcut(self.shortcuts.get("load", "Ctrl+L"))
        file_menu.addAction(self.load_action)


        self.save_action = QAction("Save Project", self)
        self.save_action.triggered.connect(self.save_project)
        self.save_action.setShortcut(self.shortcuts.get("save", "Ctrl+S"))
        file_menu.addAction(self.save_action)

        self.save_as_action = QAction("Save Project As", self)
        self.save_as_action.triggered.connect(self.save_project_as)
        self.save_as_action.setShortcut(self.shortcuts.get("save_as", "Ctrl+Shift+S"))
        file_menu.addAction(self.save_as_action)

        options_menu = menubar.addMenu("Options")
        
        Img_menu = options_menu.addMenu("Images")


        self.delete_row_action = QAction("Delete Image Row", self)
        self.delete_row_action.triggered.connect(self.delete_image)
        self.delete_row_action.setShortcut(self.shortcuts.get("delete_row", "Delete"))
        Img_menu.addAction(self.delete_row_action) #to change

        clear_action = QAction("Clear Project", self)
        clear_action.triggered.connect(self.clear_project)
        file_menu.addAction(clear_action)

        export_menu = file_menu.addMenu("Export")

        export_slideshow_action = QAction("Export Slideshow", self)
        export_slideshow_action.triggered.connect(self.export_slideshow)
        export_menu.addAction(export_slideshow_action)

        export_premiere_action = QAction("Export To Premiere", self)
        export_premiere_action.triggered.connect(self.export_premiere_slideshow)
        export_menu.addAction(export_premiere_action)

        
        
        set_all_images_duration_action = QAction("Set All Images Duration", self)
        set_all_images_duration_action.triggered.connect(self.set_all_images_duration)
        Img_menu.addAction(set_all_images_duration_action)

        set_random_image_order_action = QAction("Set Random Images Order", self)
        set_random_image_order_action.triggered.connect(self.set_random_images_order)
        Img_menu.addAction(set_random_image_order_action)

        auto_set_images_action = QAction("Auto Calculate Images Duration", self)
        auto_set_images_action.triggered.connect(self.auto_calc_image_duration)
        Img_menu.addAction(auto_set_images_action)

        self.set_image_location_action = QAction("Set Image Location", self)
        self.set_image_location_action.triggered.connect(self.set_image_location)
        self.set_image_location_action.setShortcut(self.shortcuts.get("set_image_location", "Ctrl+Q"))
        Img_menu.addAction(self.set_image_location_action)

        

        Transitions_menu = options_menu.addMenu("Transitions")
        


        set_all_images_transition_type_action = QAction("Set All Images Transition Type", self)
        set_all_images_transition_type_action.triggered.connect(self.set_all_images_transition)
        Transitions_menu.addAction(set_all_images_transition_type_action)
        
        set_random_transition_for_each_image_action = QAction("Set Random Transition For Each Image", self)
        set_random_transition_for_each_image_action.triggered.connect(self.set_random_transition_for_each_image)
        Transitions_menu.addAction(set_random_transition_for_each_image_action)

        Text_menu = options_menu.addMenu("Text")

        # Add the new "Easy Text Writing" option
        self.easy_text_writing_action = QAction("Easy Text Writing", self)
        self.easy_text_writing_action.triggered.connect(self.open_easy_text_writing)
        self.easy_text_writing_action.setShortcut(self.shortcuts.get("easy_text", "Ctrl+T"))
        Text_menu.addAction(self.easy_text_writing_action)

        # Add the Settings menu
        settings_menu = menubar.addMenu("Settings")

        # Add a submenu for keyboard shortcuts
        shortcuts_menu = settings_menu.addMenu("Keyboard Shortcuts")

        # Add actions for setting shortcuts
        set_save_shortcut_action = QAction("Set Save Shortcut", self)
        set_save_shortcut_action.triggered.connect(lambda: self.set_shortcut("save"))
        shortcuts_menu.addAction(set_save_shortcut_action)

        set_save_as_shortcut_action = QAction("Set Save As Shortcut", self)
        set_save_as_shortcut_action.triggered.connect(lambda: self.set_shortcut("save_as"))
        shortcuts_menu.addAction(set_save_as_shortcut_action)

        set_load_shortcut_action = QAction("Set Load Shortcut", self)
        set_load_shortcut_action.triggered.connect(lambda: self.set_shortcut("load"))
        shortcuts_menu.addAction(set_load_shortcut_action)

        set_easy_text_shortcut_action = QAction("Set Easy Text Writing Shortcut", self)
        set_easy_text_shortcut_action.triggered.connect(lambda: self.set_shortcut("easy_text"))
        shortcuts_menu.addAction(set_easy_text_shortcut_action)

        set_show_info_shortcut_action = QAction("Set Show Info Shortcut", self)
        set_show_info_shortcut_action.triggered.connect(lambda: self.set_shortcut("info"))
        shortcuts_menu.addAction(set_show_info_shortcut_action)


        set_delete_row_action = QAction("Set Delete Shortcut", self)
        set_delete_row_action.triggered.connect(lambda: self.set_shortcut("delete_row"))
        shortcuts_menu.addAction(set_delete_row_action)

        set_set_image_location_action = QAction("Set Set Image Location Shortcut", self)
        set_set_image_location_action.triggered.connect(lambda: self.set_shortcut("set_image_location"))
        shortcuts_menu.addAction(set_set_image_location_action)

        # Add the Info menu
        info_menu = menubar.addMenu("Info")

        self.show_info_action = QAction("Show Info", self)
        self.show_info_action.triggered.connect(self.show_info)
        self.show_info_action.setShortcut(self.shortcuts.get("info", "Alt+I"))
        info_menu.addAction(self.show_info_action)











class ImageProcessingWorker(QThread):
    progress = pyqtSignal(int)  # Signal to update the progress bar
    finished = pyqtSignal()  # Signal to indicate that processing is finished

    def __init__(self, images, output_folder, common_width, common_height):
        super().__init__()
        self.images = images
        self.output_folder = output_folder
        self.common_width = common_width
        self.common_height = common_height

    def run(self):
        for i in range(len(self.images)):
            img = self.images[i]['path']
            rotation = self.images[i]['rotation']
            text_on_slide = self.images[i]['text']
            try:
                original_image = Image.open(img)
                if original_image.size != (self.common_width, self.common_height):
                    new_image_path = Image_resizer.process_image(img, self.output_folder, text_on_slide, rotation)
                    if new_image_path:  # Only update path if the new image was created successfully
                        self.images[i]['path'] = new_image_path
                    else:
                        print(f"Failed to process image: {img}")
            except Exception as e:
                print(f"Error opening image {img}: {e}")

            # Emit progress update
            self.progress.emit(int((i + 1) / len(self.images) * 100))

        # Emit finished signal
        self.finished.emit()


class ImageProcessingPremiereWorker(QThread):
    progress = pyqtSignal(int)  # Signal to update the progress bar
    finished = pyqtSignal()  # Signal to indicate that processing is finished

    def __init__(self, images, output_folder, common_width, common_height):
        super().__init__()
        self.images = images
        self.output_folder = output_folder
        self.common_width = common_width
        self.common_height = common_height

    def run(self):
        # Define a progress callback function
        def progress_callback(progress):
            self.progress.emit(progress)  # Emit the progress to the main thread


        # Call the process_images function with the progress callback
        premiere_export.process_images(self.images, self.output_folder, progress_callback)

        # Emit finished signal
        self.finished.emit()


class EasyTextWritingDialog(QDialog):
    def __init__(self, images, affected_rows, start_index=0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Easy Text Writing")
        self.setGeometry(200, 200, 400, 200)
        self.images = images
        self.affected_rows = affected_rows
        self.current_index = start_index


        self.layout = QVBoxLayout(self)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.image_label)

        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("Enter text for the image...")
        self.text_input.setAlignment(Qt.AlignRight)  # Align text to the right
        self.text_input.setLayoutDirection(Qt.RightToLeft)  # Set layout direction to RTL
        self.text_input.setTextInteractionFlags(Qt.TextEditorInteraction)
        self.text_input.installEventFilter(self)
        self.text_input.setPlainText(self.images[self.current_index].get('text', ''))
        self.layout.addWidget(self.text_input)

        # Move the cursor to the start of the document (left side for RTL)
        self.text_input.moveCursor(QTextCursor.Start)  # Use QTextCursor.Start

        self.next_button = QPushButton("Next")
        self.next_button.clicked.connect(self.next_image)
        self.layout.addWidget(self.next_button)

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.close)
        self.layout.addWidget(self.close_button)

        self.update_image()

    def update_image(self):
        if 0 <= self.current_index < len(self.images):
            img_data = self.images[self.current_index]
            pixmap = QPixmap(img_data['path'])
            rotation = img_data.get('rotation', 0)
            if rotation != 0:
                transform = QTransform()
                transform.rotate(rotation)
                pixmap = pixmap.transformed(transform, Qt.SmoothTransformation)
            self.image_label.setPixmap(pixmap.scaled(300, 200, Qt.KeepAspectRatio))
            self.text_input.setPlainText(img_data.get('text', ''))
        else:
            self.image_label.clear()
            self.text_input.clear()

    def next_image(self):
        if 0 <= self.current_index < len(self.images):
            new_text = self.text_input.toPlainText()
            if self.images[self.current_index]['text'] != new_text:
                self.images[self.current_index]['text'] = new_text
                if self.current_index not in self.affected_rows:
                    self.affected_rows.append(self.current_index)
        self.current_index += 1
        if self.current_index >= len(self.images):
            self.current_index = 0
        self.update_image()

    def eventFilter(self, source, event):
        if source is self.text_input and event.type() == QEvent.KeyPress and event.key() == Qt.Key_Tab:
            self.next_image()
            return True
        return super().eventFilter(source, event)



class InfoDialog(QDialog):
    def __init__(self, images, audio_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Info")
        self.setGeometry(300, 300, 300, 150)

        self.layout = QVBoxLayout(self)

        # Calculate total durations
        total_images_duration_with_second = sum(img['duration'] for img in images)
        total_images_duration_without_second = sum(img['duration'] for img in images if not img.get('is_second_image', False))
        total_audio_duration = self.calculate_total_audio_duration(audio_files)

        # Format durations as hh:mm:ss
        total_images_duration_with_second_str = self.format_duration(total_images_duration_with_second)
        total_images_duration_without_second_str = self.format_duration(total_images_duration_without_second)
        total_audio_duration_str = self.format_duration(total_audio_duration)

        # Create labels to display the information
        self.layout.addWidget(QLabel(f"Total Images Duration (with second images): {total_images_duration_with_second_str}"))
        self.layout.addWidget(QLabel(f"Total Images Duration (without second images): {total_images_duration_without_second_str}"))
        self.layout.addWidget(QLabel(f"Total Audio Duration: {total_audio_duration_str}"))

        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close)
        self.layout.addWidget(close_button)

    def calculate_total_audio_duration(self, audio_files):
        total_duration = 0.0
        for audio in audio_files:
            audio_path = audio['path']
            try:
                cmd = [
                    "ffprobe",
                    "-v", "error",
                    "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1",
                    f'"{audio_path}"'
                ]
                output = subprocess.check_output(" ".join(cmd), shell=True, universal_newlines=True).strip()
                total_duration += float(output)
            except Exception as e:
                print(f"Error processing file {audio_path}: {e}")
        return total_duration

    def format_duration(self, seconds):
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02}:{minutes:02}:{secs:02}"




class CustomDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        # Get the model and check if the row is a secondary image
        model = index.model()
        row = index.row()
        col1_index = model.index(row, 1)  # Column 1 holds the filename and our data
        is_secondary = col1_index.data(Qt.UserRole)
        if is_secondary:
            painter.save()
            painter.fillRect(option.rect, QColor(100, 100, 150))  # Dark blue background
            painter.restore()
        super().paint(painter, option, index)




if __name__ == "__main__":
    app = QApplication(sys.argv)
    set_theme(app, theme='dark')
    window = SlideshowCreator()
    window.create_menu()
    window.setup_connections()
    window.show()
    sys.exit(app.exec_())