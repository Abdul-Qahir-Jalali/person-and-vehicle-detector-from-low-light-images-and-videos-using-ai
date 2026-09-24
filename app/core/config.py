from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    PROJECT_NAME: str = "Robust Object Detection System"
    API_V1_STR: str = "/api/v1"
    # Upgraded to YOLOv8 Medium (yolov8m.pt) for maximum accuracy on highly obscured objects
    MODEL_PATH: str = "yolov8m.pt"
    # Lowered confidence threshold further to catch heavily obscured vehicles in fog/glare
    CONFIDENCE_THRESHOLD: float = 0.15
    # Class IDs for YOLOv8 (COCO dataset): 0 is person, 2 is car, 3 is motorcycle, 5 is bus, 7 is truck
    ALLOWED_CLASSES: list[int] = [0, 2, 3, 5, 7]

settings = Settings()
