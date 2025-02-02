import sys
import os
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QInputDialog, QAction,
                             QListWidget, QPushButton, QLabel, QFileDialog, QSlider, QStyle, QTableWidgetItem, QSpinBox, QHeaderView, QTableWidget)
from PyQt5.QtCore import Qt, QUrl, QSize
from PyQt5.QtGui import QIcon, QFont, QPixmap

class SlideshowCreator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Slideshow Creator")
        self.setGeometry(100, 100, 1200, 800)
        
        # Initialize variables
        self.images = []
        self.audio_file = ""
        self.output_file = "output.mp4"
        
        self.create_ui()
        
    def create_ui(self):
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        
        # Left Panel - Image List with Durations
        left_panel = QVBoxLayout()

        btn_set_duration = QPushButton("Set Duration")
        btn_set_duration.clicked.connect(self.set_image_duration)
        
        
        # Initialize the image_table attribute
        self.image_table = QTableWidget()
        self.image_table.setColumnCount(2)
        self.image_table.setHorizontalHeaderLabels(["Image", "Duration (sec)"])
        self.image_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)    
        self.image_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        
        btn_add_images = QPushButton("Add Images")
        btn_add_images.clicked.connect(self.add_images)
        
        btn_add_audio = QPushButton("Add Music")
        btn_add_audio.clicked.connect(self.add_audio)
        
        left_panel.addWidget(QLabel("Slides:"))
        left_panel.addWidget(self.image_table)
        left_panel.addWidget(btn_add_images)
        left_panel.addWidget(btn_add_audio)
        left_panel.addWidget(btn_set_duration)
        
        # Center Panel - Preview
        center_panel = QVBoxLayout()
        self.preview_label = QLabel("Preview")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("border: 1px solid #444; background: #222;")
        
        self.timeline_slider = QSlider(Qt.Horizontal)
        self.timeline_slider.setEnabled(False)
        
        center_panel.addWidget(self.preview_label)
        center_panel.addWidget(self.timeline_slider)
        
        # Right Panel - Settings
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("Settings"))
        
        # New Audio Files Table
        self.audio_table = QTableWidget()
        self.audio_table.setColumnCount(2)
        self.audio_table.setHorizontalHeaderLabels(["Audio File", "Duration (sec)"])
        self.audio_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.audio_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        
        right_panel.addWidget(QLabel("Audio Files:"))
        right_panel.addWidget(self.audio_table)
        
        # Export Button
        btn_export = QPushButton("Export Slideshow")
        btn_export.clicked.connect(self.export_slideshow)
        right_panel.addWidget(btn_export)
        
        # Add panels to main layout
        main_layout.addLayout(left_panel , 1)
        main_layout.addLayout(center_panel, 2)
        main_layout.addLayout(right_panel, 1)
        
        self.setCentralWidget(main_widget)

        # Initialize audio files list
        self.audio_files = []


    def set_image_duration(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            duration, ok = QInputDialog.getInt(self, "Set Duration", "Enter duration in seconds:", self.images[row]['duration'], 1, 600)
            if ok:
                self.images[row]['duration'] = duration
                self.update_image_table()  # Refresh the table to show updated duration



    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)")
        if files:
            for file in files:
                self.images.append({'path': file, 'duration': 5})  # Default duration 5 seconds
            self.update_image_table()
            print("Images added:", self.images)  # Debugging output

    def add_audio(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Audio", "", "Audio Files (*.mp3 *.wav *.flac)")
        if file:
            self.audio_files.append({'path': file, 'duration': 0})  # Default duration 0 or set as needed
            self.update_audio_table()
            print("Audio added:", self.audio_files)  # Debugging output

    def update_audio_table(self):
        self.audio_table.setRowCount(len(self.audio_files))
        for row, audio in enumerate(self.audio_files):
            path_item = QTableWidgetItem(audio['path'])
            duration_item = QTableWidgetItem(str(audio.get('duration', 0)))  # Default duration
            self.audio_table.setItem(row, 0, path_item)
            self.audio_table.setItem(row, 1, duration_item)

    def export_slideshow(self):
        print("Images:", self.images)  # Debugging output
        print("Audio Files:", self.audio_files)  # Debugging output
        if not self.images or not self.audio_files:
            print("Please add images and audio before exporting.")
            return
        
        # Create FFmpeg command
        command = self.build_ffmpeg_command()
        print("Exporting with command:", command)
        
        # Execute FFmpeg command
        subprocess.run(command, shell=True)


    def update_image_table(self):
        self.image_table.setRowCount(len(self.images))
        for row, img in enumerate(self.images):
            # Create NEW QTableWidgetItem instances each time
            path_item = QTableWidgetItem(img['path'])  # Fresh item
            duration_item = QTableWidgetItem(str(img.get('duration', 5)))  # Fresh item
            
            self.image_table.setItem(row, 0, path_item)
            self.image_table.setItem(row, 1, duration_item)

    def build_ffmpeg_command(self):
        inputs = []
        filters = []
        concat_inputs = []
        total_duration = 0

        common_width = 1280
        common_height = 720

        # Handle image inputs with durations
        for i, img in enumerate(self.images):
            duration = img['duration']
            inputs.append(f'-loop 1 -t {duration} -i "{img["path"]}"')
            total_duration += duration

            # Ensure consistent resolution and SAR
            filters.append(f"[{i}:v]scale={common_width}:{common_height},setsar=1,format=yuv420p[{i}v]")
            concat_inputs.append(f"[{i}v]")

        # Concatenate images into a single video stream
        filter_complex = ";".join(filters)
        filter_complex += f";{''.join(concat_inputs)}concat=n={len(self.images)}:v=1:a=0[outv]"

        # Add audio input
        if self.audio_files:
            audio_index = len(self.images)  # The next index after images
            inputs.append(f'-i "{self.audio_files[0]["path"]}"')  # Use the first audio file

        # Construct final command
        command = f'ffmpeg {" ".join(inputs)} -filter_complex "{filter_complex}" -map "[outv]"'
        
        if self.audio_files:
            command += f' -map {audio_index}:a -c:a aac'  # Adjusted index dynamically

        command += ' -c:v libx264 -pix_fmt yuv420p -shortest output.mp4'

        return command


    def update_preview(self):
        """Update preview when a slide is selected"""
        selected_items = self.image_table.selectedItems()  # Changed from image_list to image_table
        if selected_items:
            row = self.image_table.row(selected_items[0])
            img_path = self.images[row]['path']
            # Load and display the selected image in the preview
            pixmap = QPixmap(img_path)
            self.preview_label.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio))

    def setup_connections(self):
        """Connect signals to slots"""
        # Connect table selection changes to preview updates
        self.image_table.itemSelectionChanged.connect(self.update_preview)

    def update_preview(self):
        """Update preview when a slide is selected"""
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            img_path = self.images[row]['path']
            # Load and display the selected image in the preview
            pixmap = QPixmap(img_path)
            self.preview_label.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio))

    def add_text_overlay(self):
        # Function to add text overlay to the selected image
        text, ok = QInputDialog.getText(self, "Add Text Overlay", "Enter text:")
        if ok and text:
            current_row = self.image_table.currentRow()  # Changed from image_list to image_table
            if current_row >= 0:
                # Store the text overlay for the current image
                self.image_table.item(current_row, 0).setText(f"{self.image_table.item(current_row, 0).text()} - {text}")  # Adjusted to access the correct column

    def create_timeline(self):
        # Function to create a visual representation of the timeline
        self.timeline_slider.setEnabled(True)
        self.timeline_slider.setMaximum(len(self.images) * 10)  # Example: 10 seconds per image
        self.timeline_slider.valueChanged.connect(self.update_timeline)

    def update_timeline(self, value):
        # Update the preview based on the timeline slider
        index = value // 10  # Assuming each image is displayed for 10 seconds
        if index < len(self.images):
            self.image_table.setCurrentCell(index, 0)  # Changed from image_list to image_table
            self.update_preview()

    def clear_project(self):
        self.images.clear()
        self.audio_files.clear()  # Clear audio files
        self.image_table.setRowCount(0)  # Proper way to clear the table
        self.audio_table.setRowCount(0)  # Clear audio table
        self.preview_label.clear()
        self.timeline_slider.setEnabled(False)

    def save_project(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(f"{self.audio_files[0]['path']}\n")  # Save the first audio file path
                for img in self.images:
                    f.write(f"{img['path']},{img.get('duration', 5)}\n") 

    def load_project(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                self.audio_files = []  # Clear existing audio files
                self.audio_file = lines[0].strip()  # Load the first audio file path
                self.audio_files.append({'path': self.audio_file, 'duration': 0})  # Add to the list
                self.images = []
                for line in lines[1:]:
                    path, duration = line.strip().split(',')  # Split path and duration
                    self.images.append({'path': path, 'duration': int(duration)})  # Store as dict
                self.update_image_table()
                self.update_audio_table()  # Update the audio table to reflect loaded audio



    def create_menu(self):
        # Function to create a menu bar
        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        
        load_action = QAction("Load Project", self)
        load_action.triggered.connect(self.load_project)
        file_menu.addAction(load_action)

        save_action = QAction("Save Project", self)
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)

        clear_action = QAction("Clear Project", self)
        clear_action.triggered.connect(self.clear_project)
        file_menu.addAction(clear_action)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SlideshowCreator()
    window.create_menu()
    window.setup_connections()
    window.show()
    sys.exit(app.exec_())