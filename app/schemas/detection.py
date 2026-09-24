from pydantic import BaseModel
from typing import List, Optional

class BoundingBox(BaseModel):
    x1: float
    y1: float
    x2: float
    y2: float

class DetectionResult(BaseModel):
    class_name: str
    confidence: float
    bbox: BoundingBox

class DetectionResponse(BaseModel):
    success: bool
    message: str
    person_count: int
    vehicle_count: int
    detections: List[DetectionResult]
    annotated_image_base64: Optional[str] = None
