import os
import cv2
import tempfile
from fastapi import APIRouter, UploadFile, File
from fastapi.responses import FileResponse
from app.schemas.detection import DetectionResponse, DetectionResult, BoundingBox
from app.services.detector import detector
from app.services.enhancer import enhancer
from app.utils.image_utils import decode_image_from_bytes, encode_image_to_base64

router = APIRouter()

@router.post("/detect/image", response_model=DetectionResponse)
async def detect_image(file: UploadFile = File(...)):
    """
    Upload an image, enhance it (fix low light and blur), detect objects, and return the annotated image.
    """
    if not file.content_type.startswith("image/"):
        return DetectionResponse(
            success=False,
            message="Invalid file type. Please upload an image.",
            person_count=0,
            vehicle_count=0,
            detections=[]
        )

    # Read image bytes
    image_bytes = await file.read()
    image_np = decode_image_from_bytes(image_bytes)

    if image_np is None:
        return DetectionResponse(
            success=False,
            message="Could not parse the image.",
            person_count=0,
            vehicle_count=0,
            detections=[]
        )

    # Enhance the image
    enhanced_np = enhancer.enhance(image_np)

    # Run detection on enhanced image
    result = detector.detect(enhanced_np)

    detections = []
    person_count = 0
    vehicle_count = 0

    # Parse results
    for box in result.boxes:
        # Get bounding box coordinates
        x1, y1, x2, y2 = box.xyxy[0].tolist()
        confidence = float(box.conf[0])
        class_id = int(box.cls[0])
        class_name = result.names[class_id]

        if class_name == "person":
            person_count += 1
        else:
            vehicle_count += 1

        detections.append(DetectionResult(
            class_name=class_name,
            confidence=confidence,
            bbox=BoundingBox(x1=x1, y1=y1, x2=x2, y2=y2)
        ))

    # Get annotated image
    annotated_image = result.plot()
    base64_image = encode_image_to_base64(annotated_image)

    return DetectionResponse(
        success=True,
        message="Detection successful",
        person_count=person_count,
        vehicle_count=vehicle_count,
        detections=detections,
        annotated_image_base64=base64_image
    )

@router.post("/detect/video")
async def detect_video(file: UploadFile = File(...)):
    """
    Upload a video, process it frame-by-frame (enhance + detect), and return the processed video.
    """
    if not file.content_type.startswith("video/"):
        return {"error": "Invalid file type. Please upload a video."}

    # Save uploaded video to a temporary file
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
    
    try:
        content = await file.read()
        temp_input.write(content)
        temp_input.close()
        temp_output.close()

        # Open video capture
        cap = cv2.VideoCapture(temp_input.name)
        if not cap.isOpened():
            return {"error": "Could not open the video file."}

        # Get video properties
        orig_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        orig_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fps = int(cap.get(cv2.CAP_PROP_FPS))
        if fps == 0:
            fps = 30
            
        # CPU Optimization 1: Resize heavy videos to a max width of 640px
        max_width = 640
        if orig_width > max_width:
            scale = max_width / orig_width
            width = max_width
            height = int(orig_height * scale)
        else:
            width = orig_width
            height = orig_height
            
        # CPU Optimization 2: Frame skipping (process 1 out of every 5 frames)
        skip_frames = 5
            
        # Define the codec and create VideoWriter object
        fourcc = cv2.VideoWriter_fourcc(*'mp4v') 
        out = cv2.VideoWriter(temp_output.name, fourcc, fps, (width, height))

        frame_count = 0
        last_annotated_frame = None

        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break
                
            frame_count += 1
            
            # Resize frame first to drastically speed up bilateral filter and YOLO
            if orig_width > max_width:
                frame = cv2.resize(frame, (width, height))

            # Only run heavy AI and Filters every N frames
            if frame_count % skip_frames == 1 or last_annotated_frame is None:
                # Enhance frame
                enhanced_frame = enhancer.enhance(frame)
                
                # Run detection
                result = detector.detect(enhanced_frame)
                
                # Annotate frame
                last_annotated_frame = result.plot()
                
            # Write the frame (smooth playback by reusing last detection on skipped frames)
            out.write(last_annotated_frame)

        cap.release()
        out.release()

        # Return the processed video file
        return FileResponse(temp_output.name, media_type="video/mp4", filename="processed_video.mp4")
    except Exception as e:
        return {"error": str(e)}
