import aspose.pydrawing as drawing
import aspose.slides as slides

# Create a presentation object
with slides.Presentation() as pres:
    # Title Slide
    slide = pres.slides[0]
    title = slide.shapes.title
    title.text = "Gesture-Based Control of PowerPoint Presentations Using Image Segmentation and Object Detection"
    subtitle = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    subtitle.text = "An Image Processing Project\n[Ruvindu Dulaksha]\n[COHDSE23.1F-006]"

    # Slide 2: Introduction
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Introduction"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Background:**\n- Importance of hands-free control in presentations.\n- Enhances user experience and engagement.\n\n**Objective:**\n- Develop a system to control PowerPoint presentations using hand gestures detected via webcam."

    # Slide 3: Methodology
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Methodology"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Overview:**\n- Capture real-time video from a webcam.\n- Use image segmentation to detect specific colors.\n- Recognize gestures based on detected colors.\n- Control PowerPoint slides based on gestures.\n\n**Components:**\n- Camera Setup\n- Image Segmentation\n- Gesture Detection\n- PowerPoint Control"

    # Slide 4: Technical Details
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Technical Details"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Image Segmentation:**\n- Convert captured frames to HSV color space.\n- Create masks for red and blue objects.\n\n**Gesture Detection:**\n- Find contours of segmented objects.\n- Identify largest contour for each color.\n- Determine gestures based on position and size of contours.\n\n**PowerPoint Control:**\n- Use win32com.client to automate PowerPoint.\n- Translate gestures to slide navigation actions."

    # Slide 5: Implementation
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Implementation"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Code Overview:**\n- Initialize PowerPoint and camera.\n- Process frames to detect colors.\n- Identify gestures and control slides.\n\n**Flow Diagram:**\n[Insert Flow Diagram Here]"

    # Slide 6: Results and Testing
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Results and Testing"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Demonstration:**\n- Screenshots or video of the system in action.\n\n**Performance:**\n- Accuracy of color detection.\n- Responsiveness of gesture recognition."

    # Slide 7: Conclusion
    slide = pres.slides.add_empty_slide(pres.layout_slides[0])
    title = slide.shapes.title
    title.text = "Conclusion"
    text_box = slide.shapes.add_textbox(slides.util.SlideUtil.get_anchored_rect(1, 50, 500, 100))
    text_box.text = "**Summary:**\n- Developed a gesture-based control system for PowerPoint.\n- Utilized image segmentation and object detection techniques.\n\n**Future Work:**\n- Improve gesture recognition accuracy.\n- Extend system to recognize more gestures."

    # Save the presentation
    pres.save("GestureControlPresentation.pptx", slides.export.SaveFormat.PPTX)

print("Presentation created successfully.")
