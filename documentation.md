

---

# **Gesture & Voice-Controlled PowerPoint**  
**Version:** 2.0  
**Author:** Harish & Team  
**Date:** February 2025  

## **Table of Contents**  
1. [Introduction](#introduction)  
2. [Features](#features)  
3. [Installation](#installation)  
4. [Usage Guide](#usage-guide)  
   - [Starting the Application](#starting-the-application)  
   - [Loading a PowerPoint File](#loading-a-powerpoint-file)  
   - [Using Gesture Control](#using-gesture-control)  
   - [Using Voice Control](#using-voice-control)  
5. [Gesture & Voice Commands](#gesture--voice-commands)  
6. [Error Handling](#error-handling)  
7. [Project Structure](#project-structure)  
8. [Dependencies](#dependencies)  
9. [Future Enhancements](#future-enhancements)  
10. [Contributing](#contributing)  
11. [License](#license)  

---

## **1. Introduction**  
This project allows users to control PowerPoint presentations using **hand gestures** and **voice commands**. It enhances accessibility and improves user experience by eliminating the need for a traditional clicker.  

---

## **2. Features**  
✅ Supports `.pptx` PowerPoint files  
✅ **Hand Gestures** to navigate slides  
✅ **Voice Commands** for hands-free control  
✅ Real-time **gesture visualization**  
✅ **PyQt5 GUI** for easy interaction  
✅ Built-in **error handling** for a smoother experience  

---

## **3. Installation**  
To run the project, install the required dependencies:  

```bash
pip install opencv-python mediapipe pyqt5 comtypes speechrecognition pyaudio
```

Ensure you have **Python 3.x** installed.  

---

## **4. Usage Guide**  

### **Starting the Application**  
Run the main script:  

```bash
python main.py
```

### **Loading a PowerPoint File**  
1. Click **"Open Presentation"** in the GUI.  
2. Select a `.pptx` file.  
3. The PowerPoint will load in slideshow mode.  

### **Using Gesture Control**  
1. Click **"Activate Gesture Control"** in the GUI.  
2. Position your hand within the webcam’s view.  
3. Use predefined hand gestures to navigate slides.  

### **Using Voice Control**  
1. Click **"Activate Voice Control"** in the GUI.  
2. Speak predefined **commands** to navigate slides.  

---

## **5. Gesture & Voice Commands**  

### **Hand Gestures**  
| Gesture | Action |  
|---------|--------|  
| Swipe Right | Next Slide |  
| Swipe Left | Previous Slide |  
| Open Hand | Pause/Resume |  
| Show Five Fingers | Go to First Slide |  

### **Voice Commands**  
| Command | Action |  
|---------|--------|  
| "Next" | Move to the next slide |  
| "Previous" | Move to the previous slide |  
| "Slide X" | Jump to slide X (e.g., "Slide 5") |  

---

## **6. Error Handling**  
- If the **microphone is unavailable**, an error message is displayed.  
- If no **hand is detected**, the program will wait until a gesture is visible.  
- If **PowerPoint crashes**, the application will attempt to restart it.  

---

## **7. Project Structure**  
```
gesture_voice_ppt/
│── main.py  # Main entry point
│── gui.py  # PyQt5 GUI
│── voice_control.py  # Handles speech recognition
│── gesture_control.py  # Detects hand gestures
│── ppt_controller.py  # Controls PowerPoint operations
│── README.md  # Project overview
│── DOCUMENTATION.md  # Detailed documentation
└── assets/  # Icons, images, etc.
```

---

## **8. Dependencies**  
- **PyQt5** (Graphical User Interface)  
- **OpenCV** (Webcam access)  
- **MediaPipe** (Hand gesture detection)  
- **SpeechRecognition** (Voice commands)  
- **PyAudio** (Microphone input)  
- **comtypes** (PowerPoint automation)  

---

## **9. Future Enhancements**  
🚀 Add AI-powered gesture customization  
🚀 Support for **Google Slides & Keynote**  
🚀 Multiple language support for voice control  

---

## **10. Contributing**  
Contributions are welcome! Fork the repo, make changes, and submit a pull request.  

---

## **11. License**  
This project is open-source under the **MIT License**.  

---
