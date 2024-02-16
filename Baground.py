import cv2
import numpy as np

# Open a connection to the camera
cap = cv2.VideoCapture(0)

# Check if the camera is opened successfully
if not cap.isOpened():
    print("Error: Could not open camera.")
    exit()

# Create a background subtractor
bg_subtractor = cv2.createBackgroundSubtractorMOG2()

while True:
    # Capture frame-by-frame
    ret, frame = cap.read()

    # If the frame is read correctly, ret will be True
    if ret:
        # Apply background subtraction
        fg_mask = bg_subtractor.apply(frame)

        # Find contours in the foreground mask
        contours, hierarchy = cv2.findContours(fg_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # Filter out small contours
        min_contour_area = 1000
        large_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > min_contour_area]

        # Draw the gesture threshold line (adjust the coordinates as needed)
        cv2.line(frame, (0, 300), (640, 300), (0, 255, 0), 2)

        # Draw contours around the hand
        for cnt in large_contours:
            x, y, w, h = cv2.boundingRect(cnt)
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)

        # Display the frame in a window
        cv2.imshow('Hand Tracking', frame)

        # Break the loop when the 'q' key is pressed
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    else:
        print("Error: Couldn't read frame.")
        break

# Release the camera and close the window
cap.release()
cv2.destroyAllWindows()
