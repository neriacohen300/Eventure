"""imports"""

import sys
import os
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QInputDialog, QAction,
                             QListWidget,QProgressBar,QMessageBox,QDialog, QPushButton, QLabel, QFileDialog, QSlider, QStyle, QTableWidgetItem, QSpinBox, QHeaderView, QTableWidget)
from PyQt5.QtCore import Qt, QUrl, QSize, QProcess, QTimer
from PyQt5.QtGui import QIcon, QFont, QPixmap, QCursor
from PIL import Image, ImageFilter
import Image_resizer

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
        
        self.create_ui()  # Create the user interface
    
    """User Interface"""
    def create_ui(self):
        # Create the main widget
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        
        """Left Panel - Image List with Durations"""
        left_panel = QVBoxLayout()

        btn_set_all_img_duration = QPushButton("Set All Images Duration") # Button to set duration
        btn_set_all_img_duration.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        btn_set_all_img_duration.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        btn_set_all_img_duration.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        btn_set_all_img_duration.clicked.connect(self.set_all_images_duration)
        
        
        # Initialize the image_table attribute
        self.image_table = QTableWidget() 
        self.image_table.setColumnCount(2)
        self.image_table.setHorizontalHeaderLabels(["Image", "Duration (sec)"])
        self.image_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        self.image_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)    
        self.image_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)

        self.image_table.itemChanged.connect(self.on_duration_edit_on_table)
        
        btn_add_images = QPushButton("Add Images") # Button to add images
        btn_add_images.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        btn_add_images.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        btn_add_images.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        btn_add_images.clicked.connect(self.add_images)
    

        # For Images
        move_up_image_btn = QPushButton("Move Up", self) # Button to move image up
        move_up_image_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        move_up_image_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        move_up_image_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        move_down_image_btn = QPushButton("Move Down", self) # Button to move image down
        move_down_image_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        move_down_image_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        move_down_image_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        delete_image_btn = QPushButton("Delete Image", self) # Button to delete image
        delete_image_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        delete_image_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        delete_image_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        """-------------------------------"""
        move_up_image_btn.clicked.connect(self.move_image_up)
        move_down_image_btn.clicked.connect(self.move_image_down)
        delete_image_btn.clicked.connect(self.delete_image)


        slides_label = QLabel("Slides")
        slides_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))
        left_panel.addWidget(slides_label) # Add label
        left_panel.addWidget(self.image_table) # Add image table
        left_panel.addWidget(btn_add_images) # Add button to add images
        left_panel.addWidget(move_up_image_btn) # Add button to move image up
        left_panel.addWidget(move_down_image_btn) # Add button to move image down
        left_panel.addWidget(delete_image_btn) # Add button to delete image
        left_panel.addWidget(btn_set_all_img_duration) # Add button to set duration


        """Center Panel - Preview"""
        center_panel = QVBoxLayout()
        self.preview_label = QLabel("Preview")
        self.preview_label.setFont(QFont(self.text_font, 16, QFont.Bold))

        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("border: 1px solid #444; background: #222;")

        self.progress_bar = QProgressBar() # Progress bar
        self.progress_bar.setRange(0, 100) # Set range
        self.progress_bar.setValue(0)  # Start at 0%
        self.progress_bar.setVisible(False)
        
        
        center_panel.addWidget(self.preview_label)
        center_panel.addWidget(self.progress_bar)

        
        """Right Panel - Audio Files"""
        right_panel = QVBoxLayout()


        # New Audio Files Table
        self.audio_table = QTableWidget()
        self.audio_table.setColumnCount(1)
        self.audio_table.setHorizontalHeaderLabels(["Audio File"])
        self.audio_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        self.audio_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        
        audio_files_label = QLabel("Audio Files:")
        audio_files_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))
        right_panel.addWidget(audio_files_label)
        right_panel.addWidget(self.audio_table)

        # For Audio
        move_up_audio_btn = QPushButton("Move Up", self) # Button to move audio up
        move_up_audio_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        move_up_audio_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        move_up_audio_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        move_down_audio_btn = QPushButton("Move Down", self) # Button to move audio down
        move_down_audio_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        move_down_audio_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        move_down_audio_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        delete_audio_btn = QPushButton("Delete Audio", self) # Button to delete audio
        delete_audio_btn.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        delete_audio_btn.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        delete_audio_btn.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        """-------------------------------"""
        move_up_audio_btn.clicked.connect(self.move_audio_up)
        move_down_audio_btn.clicked.connect(self.move_audio_down)
        delete_audio_btn.clicked.connect(self.delete_audio)

        btn_add_audio = QPushButton("Add Music") # Button to add audio
        btn_add_audio.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        btn_add_audio.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        btn_add_audio.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        btn_add_audio.clicked.connect(self.add_audio)


        
        right_panel.addWidget(btn_add_audio)
        right_panel.addWidget(move_up_audio_btn)
        right_panel.addWidget(move_down_audio_btn)
        right_panel.addWidget(delete_audio_btn)

        # Export Button
        btn_export = QPushButton("Export Slideshow")
        btn_export.clicked.connect(self.export_slideshow)
        btn_export.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        btn_export.setStyleSheet('QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }'
                                    'QPushButton:hover { background-color: #005a9e; }')
        btn_export.setCursor(QCursor(Qt.PointingHandCursor))  # Add hover effect
        right_panel.addWidget(btn_export)
        
        # Add panels to main layout
        main_layout.addLayout(left_panel , 1)
        main_layout.addLayout(center_panel, 2)
        main_layout.addLayout(right_panel, 1)
        
        self.setCentralWidget(main_widget)

        # Initialize audio files list
        self.audio_files = []







    """01_Images Functions"""
    
    """01_01_Duration Functions"""
    def on_duration_edit_on_table(self, item):
        """Handles editing the 'Duration' column."""
        column = item.column()
        if column == 1:  # Only handle edits for the 'Duration' column
            try:
                new_duration = int(item.text())
                if new_duration < 1 or new_duration > 600:
                    raise ValueError("Duration out of range (1-600).")
                row = item.row()
                self.images[row]['duration'] = new_duration  # Update the image data
            except ValueError:
                # Revert to the previous value if the input is invalid
                row = item.row()
                item.setText(str(self.images[row]['duration']))

            print(f"Duration updated for {self.images[row]['path']}    ----   {self.images[row]['duration']} \n")  # Debugging output
    def set_all_images_duration(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            new_duration, ok = QInputDialog.getInt(self, "Set Duration", "Enter duration in seconds:", self.images[row]['duration'], 1, 600)
            if ok:
                for i in range(len(self.images)):
                    self.images[i]['duration'] = new_duration
                self.update_image_table()

    """01_02_Images Implementation Functions"""
    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)")
        if files:
            for file in files:
                self.images.append({'path': file, 'duration': 5})  # Default duration 5 seconds
            self.update_image_table()
            print("Images added:", self.images)  # Debugging output

    def update_image_table(self):
        self.image_table.setRowCount(len(self.images))
        for row, img in enumerate(self.images):
            # Create NEW QTableWidgetItem instances each time
            path_img = os.path.basename(img['path'])
            filename_item = QTableWidgetItem(path_img)  # Fresh item
            duration_item = QTableWidgetItem(str(img.get('duration', 5)))  # Fresh item
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)  # Make the item non-editable

            
            self.image_table.setItem(row, 0, filename_item)
            self.image_table.setItem(row, 1, duration_item)

    def move_image_up(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            if row > 0:
                self.images[row], self.images[row - 1] = self.images[row - 1], self.images[row]
                self.image_table.setCurrentItem(self.image_table.item(row - 1, 0))
                self.update_image_table()

    def move_image_down(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            if row < len(self.images) - 1:
                self.images[row], self.images[row + 1] = self.images[row + 1], self.images[row]
                self.image_table.setCurrentItem(self.image_table.item(row + 1, 0))
                self.update_image_table()

    def delete_image(self):
        selected_items = self.image_table.selectedItems()
        if selected_items:
            row = self.image_table.row(selected_items[0])
            del self.images[row]
            self.update_image_table()






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
            self.audio_table.setItem(row, 0, filename_item)

    def move_audio_up(self):
        selected_items = self.audio_table.selectedItems()
        if selected_items:
            row = self.audio_table.row(selected_items[0])
            if row > 0:
                self.audio_files[row], self.audio_files[row - 1] = self.audio_files[row - 1], self.audio_files[row]
                self.audio_table.setCurrentItem(self.audio_table.item(row - 1, 0))
                self.update_audio_table()

    def move_audio_down(self):
        selected_items = self.audio_table.selectedItems()
        if selected_items:
            row = self.audio_table.row(selected_items[0])
            if row < len(self.audio_files) - 1:
                self.audio_files[row], self.audio_files[row + 1] = self.audio_files[row + 1], self.audio_files[row]
                self.audio_table.setCurrentItem(self.audio_table.item(row + 1, 0))
                self.update_audio_table()

    def delete_audio(self):
        selected_items = self.audio_table.selectedItems()
        if selected_items:
            row = self.audio_table.row(selected_items[0])
            del self.audio_files[row]
            self.update_audio_table()




    """03_Export Functions"""

    def export_slideshow(self):
        print("Images:", self.images)  # Debugging output
        print("Audio Files:", self.audio_files)  # Debugging output
        if not self.images or not self.audio_files:
            print("Please add images and audio before exporting.")
            return
        
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Slideshow", "", "Video Files (*.mp4);;All Files (*)", options=options)
        if file_path:
            self.output_file = file_path
        else:
            print("Please select a location for the exported video.")
            return

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
        concat_inputs = []
        total_duration = 0

        common_width = 1920
        common_height = 1080

        # Assuming self.images is a list of dictionaries with 'path' as one of the keys
        image_path22 = self.images[0]['path']
        output_folder22 = os.path.join(os.path.dirname(image_path22), "A_Blur")  # Ensure correct folder structure

        # Ensure output folder exists
        if not os.path.exists(output_folder22):
            os.makedirs(output_folder22)

        for i in range(len(self.images)):
            img = self.images[i]['path']
            
            try:
                original_image = Image.open(img)
                if original_image.size != (common_width, common_height):
                    new_image_path = Image_resizer.process_image(img, output_folder22)
                    
                    if new_image_path:  # Only update path if the new image was created successfully
                        self.images[i]['path'] = new_image_path
                    else:
                        print(f"Failed to process image: {img}")
            except Exception as e:
                print(f"Error opening image {img}: {e}")

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

        # Add audio inputs
        audio_inputs = []
        audio_streams = []
        audio_index = len(self.images)

        for i, audio in enumerate(self.audio_files):
            inputs.append(f'-i "{audio["path"]}"')
            audio_streams.append(f"[{audio_index + i}:a]")

        # Concatenate audio files sequentially
        if len(audio_streams) > 1:
            filter_complex += f";{''.join(audio_streams)}concat=n={len(audio_streams)}:v=0:a=1[outa]"
            audio_map = "-map [outa]"
        else:
            audio_map = f"-map {audio_index}:a"

        # Construct final command
        command = f'ffmpeg -y {" ".join(inputs)} -filter_complex "{filter_complex}" -map "[outv]" {audio_map} -c:a aac -c:v libx264 -pix_fmt yuv420p -shortest "{self.output_file}"'

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



    """06_Project Functions"""
    def clear_project(self):
        self.images.clear()
        self.audio_files.clear()  # Clear audio files
        self.image_table.setRowCount(0)  # Proper way to clear the table
        self.audio_table.setRowCount(0)  # Clear audio table
        self.preview_label.clear()

    def save_project(self):
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
                    f.write(f"{img['path']},{img.get('duration', 5)}\n") 

    def load_project(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                count = int(lines[0].strip())
                self.audio_files = []  # Clear existing audio files
                if count == 1:
                    self.audio_file = lines[1].strip()  # Load the first audio file path
                    self.audio_files.append({'path': self.audio_file, 'duration': 0})  # Add to the list
                if count > 1:
                    for i in range(1, count + 1):
                        self.audio_file = lines[i].strip()  # Load the first audio file path
                        self.audio_files.append({'path': self.audio_file, 'duration': 0})  # Add to the list
                self.images = []
                for line in lines[i+1:]:
                    path, duration = line.strip().split(',')  # Split path and duration
                    self.images.append({'path': path, 'duration': int(duration)})  # Store as dict
                self.update_image_table()
                self.update_audio_table()  # Update the audio table to reflect loaded audio



    """07_Menu Functions"""
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