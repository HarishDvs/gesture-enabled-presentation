import sys
import os
import speech_recognition as sr
import comtypes.client
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QGroupBox, QGridLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from nltk.tokenize import word_tokenize

class PowerPointController:
    def __init__(self):
        self.powerpoint = None
        self.presentation = None
        self.slideshow = None

    def initialize_powerpoint(self):
        self.powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        self.powerpoint.Visible = 1

    def open_presentation(self, file_path):
        try:  # ✅ This line is now correctly indented!
            if not file_path.lower().endswith('.pptx'):
                raise ValueError("Only .pptx files are supported")

            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file at path {file_path} does not exist.")

            print(f"Trying to open: {os.path.abspath(file_path)}")  

            self.presentation = self.powerpoint.Presentations.Open(os.path.abspath(file_path))
            self.start_slideshow()

        except Exception as e:
            print(f"Error: {e}")


    def start_slideshow(self):
        if self.presentation:
            self.slideshow = self.presentation.SlideIndex = 1
            self.presentation.SlideShowSettings.Run()

    def next_slide(self):
        slide_show = self.presentation.SlideShowWindow.View
        slide_show.Next()

    def previous_slide(self):
        slide_show = self.presentation.SlideShowWindow.View
        slide_show.Previous()

    def goto_slide(self, slide_number):
        slide_show = self.presentation.SlideShowWindow.View
        if 1 <= slide_number <= self.presentation.Slides.Count:
            self.slideshow.GotoSlide(slide_number)

    def close(self):
        if self.slideshow:
            self.slideshow.SlideShowWindow.View.Exit()
        if self.presentation:
            self.presentation.Close()
        if self.powerpoint:
            self.powerpoint.Quit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Voice-Controlled PowerPoint")
        self.setGeometry(100, 100, 1000, 800)

        # Setting up UI components
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QHBoxLayout(self.central_widget)

        self.left_layout = QVBoxLayout()
        self.right_layout = QVBoxLayout()
        self.layout.addLayout(self.left_layout, 2)
        self.layout.addLayout(self.right_layout, 1)

        self.status_label = QLabel("Status: Idle")
        self.left_layout.addWidget(self.status_label)

        # Button to open PowerPoint
        self.controls_layout = QHBoxLayout()
        self.open_button = QPushButton("Open PowerPoint")
        self.open_button.clicked.connect(self.open_powerpoint)
        self.controls_layout.addWidget(self.open_button)

        # Button to activate voice control
        self.voice_control_button = QPushButton("Activate Voice Control")
        self.voice_control_button.clicked.connect(self.listen_for_command)
        self.controls_layout.addWidget(self.voice_control_button)

        # Button to stop listening
        self.stop_listening_button = QPushButton("Stop Listening")
        self.stop_listening_button.setEnabled(False)
        self.stop_listening_button.clicked.connect(self.stop_listening)
        self.controls_layout.addWidget(self.stop_listening_button)

        self.left_layout.addLayout(self.controls_layout)

        # PowerPoint Controller
        self.pp_controller = PowerPointController()

        # Adding tutorial section to explain how to use the app
        self.create_tutorial_section()

        # Command history section
        self.command_history_label = QLabel("Command History: None")
        self.left_layout.addWidget(self.command_history_label)

        # Dark Mode toggle
        self.dark_mode_button = QPushButton("Toggle Dark Mode")
        self.dark_mode_button.clicked.connect(self.toggle_dark_mode)
        self.left_layout.addWidget(self.dark_mode_button)

        # State variables
        self.is_listening = False
        self.command_history = []

        # Dark Mode State
        self.dark_mode = False

    def create_tutorial_section(self):
        tutorial_group = QGroupBox("How to Use")
        tutorial_layout = QGridLayout()
        tutorial_group.setLayout(tutorial_layout)

        tutorial_text = """
        1. To go to the next slide:
           - Say "Next slide" or "Next"
        2. To go to the previous slide:
           - Say "Previous slide" or "Previous"
        3. To go to a specific slide:
           - Say "Slide X", where X is the slide number (e.g., "Slide 5" will go to slide 5)
        """
        tutorial_label = QLabel(tutorial_text)
        tutorial_layout.addWidget(tutorial_label, 0, 0)

        self.right_layout.addWidget(tutorial_group)

    def open_powerpoint(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PowerPoint Presentation", "", "PowerPoint Files (*.pptx)")
        if file_path:
            try:
                self.pp_controller.initialize_powerpoint()
                self.pp_controller.open_presentation(file_path)
                self.status_label.setText(f"Status: Opened {os.path.basename(file_path)}")
            except Exception as e:
                self.status_label.setText(f"Error: {str(e)}")

    def listen_for_command(self):
        if self.is_listening:
            return
        self.is_listening = True
        self.stop_listening_button.setEnabled(True)
        self.status_label.setText("Listening for command...")
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise
            audio = recognizer.listen(source)
            try:
                command = recognizer.recognize_google(audio)
                print(f"Command received: {command}")
                self.process_command(command)
            except sr.UnknownValueError:
                self.status_label.setText("Error: Sorry, I didn't catch that.")
            except sr.RequestError as e:
                self.status_label.setText(f"Error: Could not request results; {e}")

    def stop_listening(self):
        self.is_listening = False
        self.stop_listening_button.setEnabled(False)
        self.status_label.setText("Voice control stopped.")

    def process_command(self, command):
        command_tokens = word_tokenize(command.lower())
        self.command_history.append(command)
        self.command_history_label.setText(f"Command History: {', '.join(self.command_history[-5:])}")

        if "next" in command_tokens:
            self.pp_controller.next_slide()
            self.status_label.setText("Status: Next slide")
        elif "previous" in command_tokens:
            self.pp_controller.previous_slide()
            self.status_label.setText("Status: Previous slide")
        elif "slide" in command_tokens:
            try:
                slide_index = int(command_tokens[command_tokens.index("slide") + 1])
                self.pp_controller.goto_slide(slide_index)
                self.status_label.setText(f"Status: Going to slide {slide_index}")
            except (ValueError, IndexError):
                self.status_label.setText("Error: Invalid slide number")
        else:
            self.status_label.setText("Error: Command not recognized")

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self.setStyleSheet("background-color: #333; color: white;")
        else:
            self.setStyleSheet("background-color: white; color: black;")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
