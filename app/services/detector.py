import cv2
from ultralytics import YOLO
from app.core.config import settings

class ObjectDetector:
    def __init__(self):
        # Load the YOLO model
        self.model = YOLO(settings.MODEL_PATH)
    
    def detect(self, image):
        """
        Run inference on an image and return results.
        """
        # Run inference
        results = self.model(image, conf=settings.CONFIDENCE_THRESHOLD, classes=settings.ALLOWED_CLASSES)
        return results[0]

detector = ObjectDetector()
