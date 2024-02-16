import win32com.client  # Library for controlling PowerPoint
from cvzone.HandTrackingModule import HandDetector  # Library for hand tracking
import cv2  # OpenCV library
import os
import numpy as np
import aspose.slides as slides
import aspose.pydrawing as drawing

# Connect to PowerPoint application
Application = win32com.client.Dispatch("PowerPoint.Application")
# Open the presentation file
Presentation = Application.Presentations.Open(r"C:\Users\Welcome\Desktop\Final Year Project\OpenCV editing.pptx")
print(Presentation.Name)
# Run the slideshow
Presentation.SlideShowSettings.Run()

# Parameters for camera and hand tracking
width, height = 900, 720  # Camera frame dimensions
gestureThreshold = 300  # Threshold for detecting hand gesture
maxZoomFactor = 2  # Maximum zoom factor
minZoomFactor = 1  # Minimum zoom factor
zoomFactor = 1  # Initial zoom factor

# Camera Setup
cap = cv2.VideoCapture(0)  # Initialize camera capture
cap.set(3, width)  # Set camera frame width
cap.set(4, height)  # Set camera frame height

# Hand Detector
detectorHand = HandDetector(detectionCon=int(0.8), maxHands=1)  # Initialize hand detector

# Variables for controlling presentation
imgList = []  # List to store images
delay = 30  # Delay for button press
buttonPressed = False  # Flag to indicate if a button is pressed
counter = 0  # Counter for delay
drawMode = False  # Flag for drawing mode
imgNumber = 20  # Image number
delayCounter = 0  # Counter for delay
annotations = [[]]  # List to store annotations
annotationNumber = -1  # Current annotation number
annotationStart = False  # Flag to indicate if annotation is started

# Main loop for processing frames
while True:
    # Get image frame from camera
    success, img = cap.read()
    imgCurrent = img.copy()  # Make a copy of the image for processing
    
    # Find hands in the image
    hands, img = detectorHand.findHands(img)  # Find hands and draw landmarks on the image
    
    # Process detected hands
    if hands and buttonPressed is False:  # If hand is detected and no button is pressed
        hand = hands[0]  # Get the first detected hand
        cx, cy = hand["center"]  # Get the center coordinates of the hand
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        
        # Check hand gesture
        if cy <= gestureThreshold:  # If hand is at the height of the face
            if fingers == [1, 1, 1, 1, 1]:  # If all fingers are up
                print("Next")
                buttonPressed = True  # Set button pressed flag
                if imgNumber > 0:  # Check if there are more slides
                    Presentation.SlideShowWindow.View.Next()  # Move to the next slide
                    annotations = [[]]  # Clear annotations
                    annotationNumber = -1  # Reset annotation number
                    annotationStart = False  # Reset annotation flag
            if fingers == [1, 0, 0, 0, 0]:  # If only thumb is up
                print("Previous")
                buttonPressed = True  # Set button pressed flag
                if imgNumber >0 :  # Check if there are previous slides
                    Presentation.SlideShowWindow.View.Previous()  # Move to the previous slide
                    imgNumber += 1  # Increment image number
                    annotations = [[]]  # Clear annotations
                    annotationNumber = -1  # Reset annotation number
                    annotationStart = False  # Reset annotation flag
    else:
        annotationStart = False  # Reset annotation flag if no hand is detected
    
    # Handle button press delay
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    # Draw annotations on the current frame
    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)  # Draw lines between annotation points

    # Display the processed image frame
    cv2.imshow("Image", imgCurrent)

    # Wait for key press
    key = cv2.waitKey(1)
    if key == ord('q'):  # If 'q' is pressed, exit the loop
        break