# Robust Object Detection System

This is a modular FastAPI application that uses YOLOv8 for detecting persons and vehicles in images.

## Project Structure

- `app/main.py`: Application entry point.
- `app/api/routes/detection.py`: API endpoint for object detection.
- `app/services/detector.py`: YOLOv8 object detection service.
- `app/core/config.py`: Configuration settings.
- `app/schemas/detection.py`: Pydantic models for validation and response.
- `app/utils/image_utils.py`: Utility functions for image processing.

## Setup

1. The virtual environment is located in the `venv` folder. Activate it:
   ```bash
   .\venv\Scripts\activate
   ```

2. Install dependencies (if not already installed):
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

To run the FastAPI application locally, use Uvicorn:

```bash
uvicorn app.main:app --reload
```

Then open your browser and navigate to the interactive API documentation at:
[http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)

You can use the `/api/v1/detect` endpoint to upload images and see the detection results (counts and bounding boxes of persons and vehicles).
