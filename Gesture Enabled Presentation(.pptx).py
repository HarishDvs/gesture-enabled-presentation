import sys
import cv2
import mediapipe as mp
import comtypes.client
import os
import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QGroupBox, QGridLayout
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import Qt, QTimer

class PowerPointController:
    def __init__(self):
        self.powerpoint = None
        self.presentation = None
        self.slideshow = None

    def initialize_powerpoint(self):
        self.powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        self.powerpoint.Visible = 1

    def open_presentation(self, file_path):
        if not file_path.lower().endswith('.pptx'):
            raise ValueError("Only .pptx files are supported")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"The file at path {file_path} does not exist.")
        self.presentation = self.powerpoint.Presentations.Open(file_path)
        self.start_slideshow()

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

class GestureDetector:
    def __init__(self):
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.7)
        self.mp_drawing = mp.solutions.drawing_utils
        self.last_gesture_time = datetime.datetime.now()
        self.cooldown = datetime.timedelta(seconds=1)  # 1 second cooldown

    def detect_gesture(self, frame):
        current_time = datetime.datetime.now()
        if current_time - self.last_gesture_time < self.cooldown:
            return None

        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = self.hands.process(frame_rgb)

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                self.mp_drawing.draw_landmarks(frame, hand_landmarks, self.mp_hands.HAND_CONNECTIONS)

                thumb_tip = hand_landmarks.landmark[self.mp_hands.HandLandmark.THUMB_TIP]
                index_tip = hand_landmarks.landmark[self.mp_hands.HandLandmark.INDEX_FINGER_TIP]
                middle_tip = hand_landmarks.landmark[self.mp_hands.HandLandmark.MIDDLE_FINGER_TIP]

                if index_tip.y < middle_tip.y and thumb_tip.y > index_tip.y:
                    self.last_gesture_time = current_time
                    return "next slide"
                if thumb_tip.y < index_tip.y and thumb_tip.y < middle_tip.y:
                    self.last_gesture_time = current_time
                    return "previous slide"
                if index_tip.y < thumb_tip.y and middle_tip.y < thumb_tip.y:
                    self.last_gesture_time = current_time
                    return "go to slide"

        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gesture-Controlled PowerPoint")
        self.setGeometry(100, 100, 1000, 800)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QHBoxLayout(self.central_widget)

        self.left_layout = QVBoxLayout()
        self.right_layout = QVBoxLayout()
        self.layout.addLayout(self.left_layout, 2)
        self.layout.addLayout(self.right_layout, 1)

        self.video_label = QLabel(self)
        self.left_layout.addWidget(self.video_label)

        self.controls_layout = QHBoxLayout()
        self.open_button = QPushButton("Open PowerPoint")
        self.open_button.clicked.connect(self.open_powerpoint)
        self.controls_layout.addWidget(self.open_button)

        self.start_button = QPushButton("Start Gesture Control")
        self.start_button.clicked.connect(self.toggle_gesture_control)
        self.controls_layout.addWidget(self.start_button)

        self.left_layout.addLayout(self.controls_layout)

        self.status_label = QLabel("Status: Idle")
        self.left_layout.addWidget(self.status_label)

        self.create_gesture_guide()

        self.pp_controller = PowerPointController()
        self.gesture_detector = GestureDetector()
        self.cap = None
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_frame)
        self.is_detecting = False

        self.log_file = None

    def create_gesture_guide(self):
        gesture_group = QGroupBox("Gesture Guide")
        gesture_layout = QGridLayout()
        gesture_group.setLayout(gesture_layout)

        gestures = [
            ("Next Slide", "Index finger extended"),
            ("Previous Slide", "Thumb extended"),
            ("Go to Slide", "Index and middle finger extended")
        ]

        for i, (title, description) in enumerate(gestures):
            image_label = QLabel()
            pixmap = QPixmap(f"/api/placeholder/150/150")  # Placeholder image
            image_label.setPixmap(pixmap)
            gesture_layout.addWidget(image_label, i, 0)

            text_label = QLabel(f"<b>{title}</b><br>{description}")
            gesture_layout.addWidget(text_label, i, 1)

        self.right_layout.addWidget(gesture_group)

    def open_powerpoint(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PowerPoint Presentation", "", "PowerPoint Files (*.pptx)")
        if file_path:
            try:
                self.pp_controller.initialize_powerpoint()
                self.pp_controller.open_presentation(file_path)
                self.status_label.setText(f"Status: Opened {os.path.basename(file_path)}")
            except Exception as e:
                self.status_label.setText(f"Error: {str(e)}")

    def toggle_gesture_control(self):
        if not self.is_detecting:
            self.start_gesture_control()
        else:
            self.stop_gesture_control()

    def start_gesture_control(self):
        if self.cap is None:
            self.cap = cv2.VideoCapture(0)
        self.is_detecting = True
        self.timer.start(30)
        self.start_button.setText("Stop Gesture Control")
        self.status_label.setText("Status: Gesture Control Active")
        self.log_file = open(f"{datetime.date.today()}_gestures.txt", 'w')

    def stop_gesture_control(self):
        self.is_detecting = False
        self.timer.stop()
        self.start_button.setText("Start Gesture Control")
        self.status_label.setText("Status: Gesture Control Stopped")
        if self.log_file:
            self.log_file.close()
            self.log_file = None

    def update_frame(self):
        ret, frame = self.cap.read()
        if ret:
            gesture = self.gesture_detector.detect_gesture(frame)
            if gesture:
                self.handle_gesture(gesture)
            
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            h, w, ch = frame.shape
            bytes_per_line = ch * w
            qt_image = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_image)
            self.video_label.setPixmap(pixmap.scaled(640, 480, Qt.KeepAspectRatio))

    def handle_gesture(self, gesture):
        if gesture == "next slide":
            self.pp_controller.next_slide()
        elif gesture == "previous slide":
            self.pp_controller.previous_slide()
        elif gesture == "go to slide":
            # For simplicity, we're just going to the next slide
            # You can implement voice recognition for slide number here if needed
            self.pp_controller.next_slide()
        
        self.status_label.setText(f"Status: Detected {gesture}")
        if self.log_file:
            self.log_file.write(f"{datetime.datetime.now()}: {gesture}\n")

    def closeEvent(self, event):
        self.stop_gesture_control()
        if self.cap:
            self.cap.release()
        self.pp_controller.close()
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
