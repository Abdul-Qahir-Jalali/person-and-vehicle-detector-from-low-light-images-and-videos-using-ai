import cv2
import base64
import numpy as np

def encode_image_to_base64(image_np: np.ndarray) -> str:
    """
    Encode a numpy image to base64 string.
    """
    _, buffer = cv2.imencode('.jpg', image_np)
    base64_str = base64.b64encode(buffer).decode('utf-8')
    return base64_str

def decode_image_from_bytes(image_bytes: bytes) -> np.ndarray:
    """
    Decode image bytes to a numpy array.
    """
    nparr = np.frombuffer(image_bytes, np.uint8)
    image_np = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    return image_np
