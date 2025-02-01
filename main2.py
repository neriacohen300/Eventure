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
        
        # Initialize the image_table attribute
        self.image_table = QTableWidget()  # Ensure this is defined before use
        self.image_table.setColumnCount(2)
        self.image_table.setHorizontalHeaderLabels(["Image", "Duration (sec)"])
        self.image_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.image_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        
        btn_add_images = QPushButton("Add Images")
        btn_add_images.clicked.connect(self.add_images)
        
        btn_add_audio = QPushButton("Add Music")
        btn_add_audio.clicked.connect(self.add_audio)  # Connect to the add_audio method
        
        left_panel.addWidget(QLabel("Slides:"))
        left_panel.addWidget(self.image_table)  # Now this will work
        left_panel.addWidget(btn_add_images)
        left_panel.addWidget(btn_add_audio)
        
        # Center Panel - Preview
        center_panel = QVBoxLayout()
        self.preview_label = QLabel("Preview")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("border: 1px solid #444; background: #222;")
        
        # Timeline
        self.timeline_slider = QSlider(Qt.Horizontal)
        self.timeline_slider.setEnabled(False)
        
        center_panel.addWidget(self.preview_label)
        center_panel.addWidget(self.timeline_slider)
        
        # Right Panel - Settings
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("Settings"))
        
        # Export Button
        btn_export = QPushButton("Export Slideshow")
        btn_export.clicked.connect(self.export_slideshow)
        right_panel.addWidget(btn_export)
        
        # Add panels to main layout
        main_layout.addLayout(left_panel , 1)
        main_layout.addLayout(center_panel, 2)
        main_layout.addLayout(right_panel, 1)
        
        self.setCentralWidget(main_widget)

    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)")
        if files:
            for file in files:
                self.images.append({'path': file, 'duration': 5})  # Default duration 5 seconds
                self.update_image_table()

    def add_audio(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Audio", "", "Audio Files (*.mp3 *.wav *.flac)")
        if file:
            self.audio_file = file

    def export_slideshow(self):
        if not self.images or not self.audio_file:
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
            # Image path
            path_item = QTableWidgetItem(img['path'])
            path_item.setFlags(path_item.flags() ^ Qt.ItemIsEditable)
            
            # Duration spinner
            spin = QSpinBox()
            spin.setMinimum(1)
            spin.setMaximum(60)
            spin.setValue(img['duration'])
            spin.valueChanged.connect(lambda value, row=row: self.update_duration(value, row))
            
            self.image_table.setItem(row, 0, path_item)
            self.image_table.setCellWidget(row, 1, spin)

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
        if self.audio_file:
            audio_index = len(self.images)  # The next index after images
            inputs.append(f'-i "{self.audio_file}"')

        # Construct final command
        command = f'ffmpeg {" ".join(inputs)} -filter_complex "{filter_complex}" -map "[outv]"'
        
        if self.audio_file:
            command += f' -map {audio_index}:a -c:a aac'  # Adjusted index dynamically

        command += ' -c:v libx264 -pix_fmt yuv420p -shortest output.mp4'

        return command


    def update_preview(self):
        # Placeholder for updating the preview based on the current selection
        if self.image_list.count() > 0:
            current_image = self.image_list.currentItem().text()
            self.preview_label.setText(f"Previewing: {current_image}")
            # Here you could load the image into a QLabel or QGraphicsView for a better preview

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
            current_row = self.image_list.currentRow()
            if current_row >= 0:
                # Store the text overlay for the current image
                self.image_list.item(current_row).setText(f"{self.image_list.item(current_row).text()} - {text}")

    def create_timeline(self):
        # Function to create a visual representation of the timeline
        self.timeline_slider.setEnabled(True)
        self.timeline_slider.setMaximum(len(self.images) * 10)  # Example: 10 seconds per image
        self.timeline_slider.valueChanged.connect(self.update_timeline)

    def update_timeline(self, value):
        # Update the preview based on the timeline slider
        index = value // 10  # Assuming each image is displayed for 10 seconds
        if index < len(self.images):
            self.image_list.setCurrentRow(index)
            self.update_preview()

    def clear_project(self):
        # Function to clear the current project
        self.images.clear()
        self.audio_file = ""
        self.image_list.clear()
        self.preview_label.clear()
        self.timeline_slider.setEnabled(False)

    def save_project(self):
        # Function to save the current project state
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'w') as f:
                f.write(f"{self.audio_file}\n")
                for img in self.images:
                    f.write(f"{img}\n")

    def load_project(self):
        # Function to load a previously saved project
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'r') as f:
                lines = f.readlines()
                self.audio_file = lines[0].strip()
                self.images = [line.strip() for line in lines[1:]]
                self.image_list.addItems(self.images)

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