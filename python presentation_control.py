import os

import cv2
import numpy as np
import win32com.client

# Initialize PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

# Specify the path to your presentation file
presentation_path = os.path.abspath("zani.pptx")  # Replace with the correct path

try:
    # Open the PowerPoint presentation
    Presentation = Application.Presentations.Open(presentation_path)
except Exception as e:
    print(f"Failed to open presentation: {e}")
    Application.Quit()
    raise

Presentation.SlideShowSettings.StartingSlide = 1
Presentation.SlideShowSettings.EndingSlide = Presentation.Slides.Count
Presentation.SlideShowSettings.AdvanceMode = 1  # Manual advance
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720
gestureThreshold = 300
red_lower = np.array([0, 120, 70])
red_upper = np.array([10, 255, 255])
blue_lower = np.array([94, 80, 2])
blue_upper = np.array([126, 255, 255])
gesture_delay = 30

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Variables
buttonPressed = False
counter = 0

while True:
    # Get image frame
    success, img = cap.read()
    if not success:
        break

    # Convert image to HSV
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # Create masks for red and blue colors
    mask_red = cv2.inRange(hsv, red_lower, red_upper)
    mask_blue = cv2.inRange(hsv, blue_lower, blue_upper)

    # Find contours for the masks
    contours_red, _ = cv2.findContours(mask_red, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contours_blue, _ = cv2.findContours(mask_blue, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    # Draw contours and find the largest one for red
    if contours_red:
        largest_red = max(contours_red, key=cv2.contourArea)
        area_red = cv2.contourArea(largest_red)
        if area_red > 500:
            x, y, w, h = cv2.boundingRect(largest_red)
            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 0, 255), 3)
            if y <= gestureThreshold and buttonPressed is False:
                print("Next Slide")
                Presentation.SlideShowWindow.View.Next()
                buttonPressed = True

    # Draw contours and find the largest one for blue
    if contours_blue:
        largest_blue = max(contours_blue, key=cv2.contourArea)
        area_blue = cv2.contourArea(largest_blue)
        if area_blue > 500:
            x, y, w, h = cv2.boundingRect(largest_blue)
            cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 3)
            if y <= gestureThreshold and buttonPressed is False:
                print("Previous Slide")
                Presentation.SlideShowWindow.View.Previous()
                buttonPressed = True

    if buttonPressed:
        counter += 1
        if counter > gesture_delay:
            counter = 0
            buttonPressed = False

    # Display the image
    cv2.imshow("Image", img)

    # Break loop on 'q' key press
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
Application.Quit()
