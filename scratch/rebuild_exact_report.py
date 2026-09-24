"""Rebuild the current project report in the retained older FYP DOCX shell."""
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree
import re

ROOT=Path(r'E:\hammad project')
DOC_DIR=ROOT/'project details'/'doccumentation'
TEMPLATE=ROOT/'scratch'/'legacy_template_repaired.docx'
OUTPUT=DOC_DIR/'Robust Object Detection System FYP Documentation.docx'
DIAGRAMS=ROOT/'scratch'/'current_diagrams'
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

TITLES={
'DeepFake Audio Detection for Urdu':'Robust Object Detection System Under Adverse Weather Conditions',
'1.1.1 Motivation for Solution':'1.1.1 Motivation for the Solution',
'2.1 Deepfake Audio and Speech Synthesis':'2.1 Object Detection and Computer Vision',
'2.1.1 Overview of Deepfake Technology':'2.1.1 Overview of Object Detection',
'2.1.2 Evolution of Speech Cloning and TTS Systems':'2.1.2 Evolution of Object Detection Models',
'2.1.3 Challenges in Detecting Deepfake Audio':'2.1.3 Challenges in Adverse Visibility',
'2.2 Machine Learning for Audio Deepfake Detection':'2.2 Deep Learning for Vision',
'2.2.1 Feature Extraction and Representation in Speech':'2.2.1 Feature Extraction and Representation in Images',
'2.2.2 Deep Learning Architectures for Speech Analysis':'2.2.2 Deep Learning Architectures for Image Analysis',
'2.2.3 Challenges in ML-Based Deepfake Detection':'2.2.3 Challenges in ML-Based Object Detection',
'2.3.1 Existing Deepfake Detection Systems':'2.3.1 Existing Object Detection Systems',
'2.4.1 Facebook Wav2Vec2-Large-XLSR-53':'2.4.1 YOLOv8 Object Detection',
'2.4.2 Flask Framework':'2.4.2 FastAPI Framework',
'2.4.3 ElevenLabs API':'2.4.3 OpenCV',
'3.2.1 Flask Server':'3.2.1 FastAPI Server',
'3.2.2 Elevenlabs API Integration':'3.2.2 OpenCV Integration',
'3.2.3 Audio Preprocessing and Feature Extraction':'3.2.3 Image Preprocessing and Enhancement',
'3.2.4 Wav2Vec2 Detection Model':'3.2.4 YOLOv8 Detection Model',
'3.3 Real-Time Processing and Communication':'3.3 Media Processing and HTTP Communication',
'3.3.1 Real-Time  WorkFlow':'3.3.1 Image and Video Processing Workflow',
'3.4.1 Data Flow for DeepFake Detection':'3.4.1 Image and Video Data Flow',
'3.4.3 Sequence Diagram':'3.4.3 Image Request Sequence',
'3.4.4 Activity Diagram':'3.4.4 Activity Diagram',
'3.4.5 System Overview Diagram':'3.4.5 System Overview Diagram',
'3.4.6 System Architecture Diagram':'3.4.6 System Architecture Diagram',
'3.4.7  Use Case  Diagram':'3.4.7 Use Case Diagram',
'4.1 Machine Learning Model':'4.1 Pretrained Detection Model',
'4.1.2 Dataset Selection':'4.1.2 Weights and Class Selection',
'4.1.3 Training and Evaluation':'4.1.3 Inference and Evaluation Plan',
'4.1.4 Model Optimization':'4.1.4 Confidence and Runtime Settings',
'4.2 Audio Preprocessing and Noise Removal':'4.2 Image Preprocessing and Enhancement',
'4.2.1 ElevenLabs API Integration':'4.2.1 OpenCV Enhancement Pipeline',
'4.2.2 Feature Preparation':'4.2.2 Input and Output Image Handling',
'4.3 Flask Integration and Model Serving':'4.3 FastAPI Integration and Model Serving',
'4.4.3 Integration Testing':'4.4.3 Verification and Error Handling',
'4.5 Deployment and Testing':'4.5 Local Deployment and Verification',
'4.5.2 System Testing':'4.5.2 System Verification Plan',
'5.6 Uploading an Audio File':'5.6 Uploading Image and Video Files',
'6.1 Evaluation Results':'6.1 Verification and Evaluation Status',
'6.1.1 Detection Performance':'6.1.1 Detection Performance Evaluation Plan',
'6.1.2 Noise-Removal Impact':'6.1.2 Enhancement Evaluation',
'6.1.3 Confidence-Score Evaluation':'6.1.3 Confidence Threshold Evaluation',
'6.1.4 Processing Time Analysis':'6.1.4 Processing Time Evaluation',
'6.1.5 Web App Testing':'6.1.5 Web Application Verification',
'6.2 Case Study Evaluation':'6.2 Evaluation Protocol',
'6.2.1 Real-World Audio Sources':'6.2.1 Representative Image and Video Sources',
'6.3 Key Findings':'6.3 Findings from the Implemented Prototype',
'6.4.1 Technical Challenges':'6.4.1 Technical Limitations',
'7.1 Conclusion':'7.1 Conclusion',
}

CONTENT={
'Chapter 1':"""This report documents the design and implementation of a web-based object detection prototype for images and videos in challenging visibility conditions. It describes the application structure, enhancement sequence, pretrained detector, request and response formats, browser workflow, and limits visible in the delivered source. The system combines OpenCV preprocessing with YOLOv8 inference behind a local FastAPI application. The project proposal identifies low light, rain, and fog as motivating conditions. The implementation supplies general image operations, but no labeled weather dataset or benchmark results. This report therefore distinguishes implemented functions from future evaluation instead of presenting an unverified accuracy claim.""",
'1.1 Problem Statement and Motivation':"""Object detectors depend on visual cues such as edges, texture, shape, and contrast. Low illumination, glare, fog, rain, blur, and compression can weaken or obscure these cues. A detector may miss an object, predict an incorrect class, or place a bounding box inaccurately. This project addresses the problem by applying image processing before pretrained object detection and exposing the workflow through a browser. It is a prototype for uploaded media, not a safety-certified monitoring system or an autonomous vehicle controller.""",
'1.1.1 Motivation for Solution':"""Computer vision is used in road monitoring, surveillance, and visual inspection, yet a model that works well on clear images may not behave equally on degraded inputs. A local upload-and-analysis tool lets a user examine images and videos without preparing a command-line inference pipeline. The enhancement stage makes the processing sequence explicit, so later experiments can compare it with inference on unmodified media. Whether the enhancement improves detection must be measured with labeled examples; visual changes alone do not establish better recognition, and no filter can reliably reconstruct details missing from the source.""",
'1.2 Scope':"""The application serves a static browser page and provides API routes under /api/v1. The image endpoint receives an uploaded image, decodes and enhances it, runs YOLO inference, then returns person and vehicle counts, individual labels, confidence scores, bounding-box coordinates, and an annotated image encoded in base64. The video endpoint accepts an uploaded video, processes frames with OpenCV, and returns an annotated MP4. The model is loaded from local weights and runs on the host machine.""",
'1.2.1 Inclusions and Exclusions':"""The project includes FastAPI request handling, static HTML, CSS and JavaScript, image and video uploads, OpenCV decoding and enhancement, YOLO inference, response formatting, and annotations. Its inference filter selects the COCO person class and four vehicle classes: car, motorcycle, bus, and truck. The repository does not include model training, a curated labeled dataset, weather-specific weights, recorded benchmark results, persistent user accounts, a database, or object identity tracking. Video is submitted as a file; the application does not stream from a camera.""",
'1.3 Aims and Objectives':"""The aim is to integrate image enhancement and pretrained object detection in a usable local web application. Objectives are to separate API routes, schemas, configuration, utilities, and model services; accept image and video files; apply a consistent processing sequence; identify selected people and vehicle categories; return interpretable detections for images; and generate annotated video output. A further objective is to document the configuration needed to reproduce inference, including the weight file, confidence threshold, allowed classes, maximum video-frame width, and sampling interval. The implementation should remain understandable enough for later testing and extension.""",
'1.4 Limitations':"""The model uses general pretrained weights and may not represent local roads or adverse weather. The code includes no fine-tuning process, labeled validation split, or measured accuracy results. Enhancement parameters are fixed and can help some scenes while degrading others. The image route counts detections in one image, not unique identities over time. Video inference runs on every fifth frame and reuses the latest annotated frame between detections, so boxes can be stale for moving objects. The application has no explicit upload-size cap, authentication, persistent storage, or production deployment controls. Runtime also depends on the host processor, model loading, input size, codec availability, and installed package versions.""",
'1.5 Intended Audience and Target Users':"""This report is intended for project evaluators, computer-science students, developers extending the code, and users who need to inspect uploaded images or videos. It explains what the prototype does and what remains unverified. A displayed box is a model prediction, not a verified fact. A researcher can use the described pipeline as a starting point for a controlled comparison, provided a labeled test set is assembled and the model, hardware, threshold, and evaluation settings are recorded. High-impact decisions should not rely on this prototype alone.""",
'Chapter 2':"""This chapter introduces the computer-vision concepts and software technologies relevant to the project. The system joins a pretrained detector, image enhancement operations, an HTTP application, and a browser interface. The technical challenge is to connect these components while making input, output, and assumptions explicit. General background is separated from project-specific implementation: research on detectors and enhancement explains the field, while statements about the implemented application are based on the repository. Performance claims belong to a controlled evaluation and are not inferred from the presence of a working interface.""",
'2.1 Object Detection and Computer Vision':"""Computer vision converts image pixels into descriptions of a scene. Image classification assigns one or more labels to an image; localization estimates object position; detection predicts both category and position for multiple objects. A detector commonly returns bounding boxes, class labels, and scores. This project passes selected predictions to the browser as counts and annotations. These are image-level predictions, not object identities or proof that an object exists. The distinction is important when a user interprets a tally or reviews a video, because repeated detections in consecutive frames can refer to the same physical object.""",
'2.1.1 Overview of Object Detection':"""A bounding box is generally expressed by two coordinate corners, here named x1, y1, x2, and y2. The selected class describes the model's category prediction, while confidence is the model score associated with that prediction. A threshold discards candidates below a configured value; the framework may also apply its own overlap filtering. In this application Ultralytics returns a result object containing boxes, scores, class IDs, and class names. The route converts those values to Pydantic models and uses the result's plotting function to produce the annotated image. Coordinates are tied to the image array passed into inference and are not ground-truth measurements.

Detection and classification answer different questions. A classification system may say that a whole image contains a car, while a detector also estimates where the car appears. Multiple instances of one category produce multiple boxes. A detector can miss an object, produce a false positive, or place a box too tightly or loosely. Evaluation therefore needs labeled boxes and a stated overlap rule, such as intersection over union. Counts are useful summaries but conceal localization errors and class confusions, so they should be considered alongside per-detection outputs.""",
'2.1.2 Evolution of Object Detection Models':"""Early detection systems combined manually designed image features with classifiers. Deep convolutional networks learned more effective representations from examples and enabled region-proposal methods such as R-CNN and its successors. These two-stage systems first identify candidate regions and then classify them. Single-stage methods such as SSD and YOLO predict locations and categories in a more direct pass, often reducing inference time. New model generations adjust their architectures, data pipelines, and training procedures, so family names alone do not determine a particular model's quality.

This project uses the Ultralytics YOLO interface and a local yolov8m.pt file. The choice reflects an existing pretrained implementation that can be integrated into a small Python service. It does not establish superiority over Faster R-CNN, SSD, or other YOLO variants. The repository also contains yolov8n.pt and yolov8s.pt weights, but the configuration selects the medium checkpoint. A fair model comparison would hold the dataset, image sizes, hardware, threshold-selection procedure, and evaluation metric constant.""",
'2.1.3 Challenges in Adverse Visibility':"""Low illumination reduces signal and may increase the relative effect of sensor noise. Bright headlights or reflections can saturate pixels and hide texture. Fog and rain scatter or obscure light, reducing contrast between objects and backgrounds; motion blur removes edge detail, and compression adds artifacts. Occlusion may hide part of an object even in clear weather. The effect varies with camera, distance, object size, exposure, weather intensity, and image processing in the capture device. One weather label therefore does not describe a uniform detection problem.

Image enhancement can alter local or global pixel values, smooth some noise, and sharpen some boundaries. It cannot reliably reconstruct an object hidden by an opaque obstruction or clipped by saturation. A transform that increases contrast in one scene may amplify noise in another. The implementation applies the same sequence to each accepted image and sampled video frame rather than choosing settings from a weather classifier. Robustness must be measured on representative labeled examples and compared with the same detector applied to original inputs.""",
'2.2 Machine Learning for Audio Deepfake Detection':'',
'2.2 Deep Learning for Vision':"""Deep-learning vision systems learn layered image representations from training data. Shallow feature maps can encode local intensity changes and edges, while later layers combine patterns over larger regions. Detection heads use those features to predict object categories and box positions at one or more scales. The project does not define these layers itself; Ultralytics loads pretrained weights and performs inference. Application code controls the image supplied to the model, the class filter, the confidence threshold, and the handling of results. The pretrained model's learned behavior is inherited from its training process and cannot be attributed to this project without additional evidence.""",
'2.2.1 Feature Extraction and Representation in Speech':'',
'2.2.1 Feature Extraction and Representation in Images':"""An image is stored as a grid of color values, commonly in three channels. Local edges and color transitions can distinguish a vehicle from a road or a person from the background. Larger shapes and context help resolve ambiguous patches. Small or distant objects occupy few pixels, so resizing and blur can remove useful detail. In the image route, OpenCV decodes to a color array and passes the enhanced array into YOLO. In the video route, frames wider than 640 pixels are resized proportionally before enhancement and inference. That limit lowers per-frame work but also changes the amount of visual detail available to the detector.

The enhancer converts the smoothed BGR frame to LAB, applies CLAHE to the luminance channel, then converts the image back to BGR before sharpening. This separates brightness processing from two color-opponent channels. CLAHE works on local tiles and limits contrast amplification according to its clip limit; its behavior depends on tile size and image content. The implementation fixes the tile grid at 8 by 8 and the clip limit at 2.0. These values are implementation settings, not a result of parameter optimization in a documented experiment.""",
'2.2.2 Deep Learning Architectures for Speech Analysis':'',
'2.2.2 Deep Learning Architectures for Image Analysis':"""A one-stage detector maps image features to candidate boxes and class scores in a single inference pipeline. The Ultralytics API hides the detailed model internals from the web route and supplies a result object. In code, the service calls the model with a confidence threshold and selected class IDs, then returns the first result. This abstraction keeps route logic focused on validation and response conversion, while the detector service handles model loading and prediction.

The repository contains pretrained YOLOv8 weights rather than a training project. It has no dataset manifest, annotation files, loss calculation, optimizer, epoch loop, or training configuration. The medium weights path is the default in configuration. Therefore the report describes transfer of an existing model into an application, not a new neural architecture or fine-tuned checkpoint. Any future training effort would require separate data provenance, class definitions, train/validation/test splits, reproducible settings, and review for overlap with evaluation data.""",
'2.2.3 Challenges in ML-Based Deepfake Detection':'',
'2.2.3 Challenges in ML-Based Object Detection':"""A detector's output can vary with domain shift, object scale, camera angle, occlusion, illumination, blur, and the frequency of classes in its training data. A confidence threshold changes how many candidate predictions are retained. Lowering it may recover weak true detections but can also admit more false positives. The configured value 0.15 is a fixed operating parameter; the repository does not show a validation-based threshold-selection study. A score should not be read as a calibrated probability unless calibration has been tested.

Robust evaluation requires more than counting visible boxes on a few sample images. Ground-truth annotations must cover each relevant class and operating condition. Localization quality is measured by matching predictions with labeled boxes under a declared overlap threshold. Precision and recall characterize false-positive and missed-detection behavior; average precision summarizes a precision-recall curve, and mean average precision averages over classes or specified overlap thresholds. Metrics are meaningful only with a declared dataset, split, counting rule, and model configuration. None of those project-specific metric results is supplied in the code repository.""",
'2.3 Existing Solutions and Research Gaps':'',
'2.3.1 Existing Deepfake Detection Systems':"""Object detection systems are available in research frameworks and general computer-vision libraries. Two-stage detectors emphasize region proposals and classification; single-stage detectors predict categories and positions more directly. YOLO implementations are often selected where users want a practical speed-accuracy balance. General pretrained detectors are commonly evaluated on broad labeled datasets such as COCO, whose taxonomy includes everyday categories such as people and vehicles. This project uses a small subset of those labels and a pretrained checkpoint, then exposes predictions through its own API and interface.

Image restoration and enhancement methods form a separate research area. Dehazing, denoising, contrast enhancement, and low-light adjustment attempt to change the appearance of an input before downstream recognition. A combined pipeline can be evaluated by comparing its detections with detections from unmodified images under identical model settings. The current repository implements an enhancement sequence but does not include a labeled weather benchmark or comparison table. Thus it is an integration prototype, not evidence that the selected pipeline closes a research gap or improves accuracy.""",
'2.3.2 Limitations of Current Approaches':"""A general detector may not cover rare local objects or difficult visual domains equally well. Performance can drop when deployment images differ from the data used to train the model. Enhancement adds compute and can create halos, suppress texture, change object edges, or amplify noise. Fixed parameters may be unsuitable for an entire dataset. Results are also sensitive to image resizing, threshold choice, confidence calibration, annotation quality, and software versions.

The project has no class-specific error analysis, per-weather scores, baseline comparison, or published latency measurements. Its five selected classes are inherited from the configured COCO IDs. Its video output omits frame-level structured detections and tracking identifiers. It should not be compared with another system unless that comparison uses a common dataset, fixed preprocessing, specified hardware, and an explicit evaluation protocol. Chapter 6 proposes such a protocol and identifies the missing evidence.""",
'2.4 Technologies Overview':"""The software uses Python services, a FastAPI web application, OpenCV media operations, Pydantic schemas, NumPy arrays, and browser-based HTML, CSS, and JavaScript. Ultralytics supplies the YOLO model API. Uvicorn runs the ASGI app during local development. Components are separated into route, service, schema, configuration, utility, and static-interface directories. This modular organization makes code paths visible but does not imply horizontal scaling or production readiness.""",
'2.4.1 YOLOv8 Object Detection':"""The detector service creates an Ultralytics YOLO model from the local path configured as yolov8m.pt. During prediction, it passes the image together with confidence 0.15 and allowed COCO class IDs [0, 2, 3, 5, 7]. Those IDs select person, car, motorcycle, bus, and truck. The first result is returned to the API route. That route extracts each box's coordinates, class ID, confidence value, and name, then asks the result object to draw annotations.

Pretrained inference is distinct from training. Loading weights initializes the model parameters; a prediction call uses the current parameter values to calculate candidate detections. No project-specific weight update is performed in the source. The application should therefore be described as using a pretrained YOLOv8 medium checkpoint. Its confidence scores and selected boxes reflect both that checkpoint and the configured threshold, and can change if weights or library versions change.""",
'2.4.2 FastAPI Framework':"""FastAPI defines the ASGI application, registers the versioned detection router, receives multipart uploads, and generates responses. The root route redirects to the static interface, and the static directory is mounted at /static. FastAPI's generated interactive API documentation is available at /docs while the app runs. The image endpoint declares a Pydantic response model; the video endpoint sends a video file response. FastAPI does not decode, enhance, or detect objects; those operations are delegated to utilities and services.

The API prefix is configured as /api/v1. The image operation is POST /api/v1/detect/image, and the video operation is POST /api/v1/detect/video. The repository's older test_client.py uses the path /api/v1/detect, which does not match the implemented image route and should be updated before it is used as a verification client. This difference illustrates why endpoint documentation should be generated or checked against actual route definitions.""",
'2.4.3 OpenCV':"""OpenCV is used to decode image bytes, process color channels, apply filters, capture video, resize frames, and encode the output file. The enhancer applies a lookup-table intensity transform, bilateral filtering, CLAHE in LAB luminance, and a sharpening kernel. Video processing reads frames from a temporary file and writes annotated frames through VideoWriter using the mp4v codec. The actual availability of video codecs depends on the OpenCV build and machine environment.

Media-processing functions operate on arrays and files, so input validation and resource cleanup matter. The code checks whether a video capture opens and returns an error object if it does not. The current implementation creates temporary input and output files but does not consistently remove them in every early-return and exception path. It also accepts or rejects based primarily on the declared upload content type rather than a full independent media inspection.""",
'2.4.4 Supporting Libraries and Tools':"""NumPy holds image data and constructs the lookup table and filter kernels. Pydantic defines bounding-box, detection-result, and response models. pydantic-settings provides configuration values such as model path, confidence threshold, and allowed classes. Python-multipart enables multipart upload parsing. HTML defines the interface, CSS controls presentation, and JavaScript switches tabs, handles drag-and-drop, posts files with FormData, and displays results. Requirements.txt lists the principal packages but does not pin versions. A reproducible evaluation should record exact package and model versions, operating system, codec support, and hardware.""",
'Chapter 3':"""This chapter describes the system as implemented in the supplied repository. The design is a local web application with a browser presentation layer, a FastAPI route layer, reusable media and model services, response schemas, and configuration. Images follow a synchronous request path and return JSON. Videos are written to temporary files, processed frame by frame, then returned as an MP4. The diagrams in this chapter use current module names and data formats; they are not a cloud deployment design.

3.1 Architecture Overview. The user interface is served from app/static. The browser submits a selected file to a versioned route. FastAPI validates the declared media type and delegates work to helper functions. Image bytes are decoded into a NumPy array, enhanced, and passed to the YOLO service. The API converts result objects into a stable response model. For video, OpenCV captures and writes frames, while the detector and enhancer are called at the configured sampling interval. The same service modules are reused by both routes.

3.1.1 High-Level Design. The presentation layer contains the HTML page, CSS stylesheet, and JavaScript behavior. The API layer routes image and video requests and handles upload responses. The application-services layer loads YOLO weights and applies the image-enhancement sequence. Utilities convert bytes to arrays and annotated arrays back to base64. Pydantic models describe image detections and bounding boxes. Configuration holds the model path, API prefix, confidence cutoff, and allowed classes. This separation gives each responsibility a visible location without introducing separate deployed services.

3.1.2 Architectural Layers. The browser is responsible for choosing and submitting a local file, indicating that processing is underway, and rendering returned results. It does not run the detector. FastAPI coordinates requests and formats outputs. The detector service owns the loaded Ultralytics model. The enhancer owns its CLAHE object and applies the configured sequence to arrays. Image utilities perform OpenCV decoding and encoding. Configuration is instantiated from pydantic-settings when the application starts. Model loading at startup reduces repeated initialization for each request, although it also means a missing or invalid model file can prevent a usable service from starting.

3.2 Component Details. The component boundaries follow source files: main.py creates the app; api/routes/detection.py defines endpoints; services/detector.py and enhancer.py perform inference and enhancement; schemas/detection.py defines response objects; core/config.py holds settings; utils/image_utils.py converts image data; and static files implement the browser page. The structure supports code-level change isolation. It does not currently include a database, queue, authentication layer, persistent job records, or tracking state.

3.2.1 FastAPI Server. Application startup creates a FastAPI instance, sets its name, description, and version, and registers the router under the configured prefix. It mounts app/static and redirects the root path to index.html. The image route returns a typed DetectionResponse. The video route returns a FileResponse for a successful output and a simple error object for invalid or unreadable media. Exceptions in video processing are caught and their message is returned; a production service should use structured logging and avoid exposing internal exception details to clients.

3.2.2 OpenCV Integration. OpenCV decodes the uploaded image bytes using imdecode and a NumPy buffer. The enhancer uses color conversion, a bilateral filter, CLAHE, and filter2D. The video route opens the temporary input with VideoCapture, reads width, height, and frame rate, resizes frames as required, and writes output through VideoWriter. Capture and writer resources are released after normal processing. Error branches and exceptions should be reviewed to ensure cleanup occurs for every path and that partially written output files are not left behind.

3.2.3 Image Preprocessing and Enhancement. The enhancer first maps intensity values with a gamma lookup table configured from gamma 1.5. The code computes each output entry using the reciprocal exponent, so the transformation lifts many mid-range values; the nearby source comment describes a glare-darkening intention that does not match this calculation. It next applies a bilateral filter with diameter 9 and sigma values 75. The smoothed BGR image is converted to LAB, CLAHE is applied to luminance with clip limit 2.0 and an 8 by 8 grid, and the image is converted back to BGR. A mild four-neighbor sharpening kernel is applied last. This fixed sequence changes image appearance but has no adaptive weather classifier.

3.2.4 YOLOv8 Detection Model. The detector service loads the configured yolov8m.pt file through Ultralytics. Its predict call uses confidence 0.15 and class IDs 0, 2, 3, 5, and 7. The image route reads xyxy coordinates, confidence, class ID, and class name from each box. It increments the person count for class name person and the vehicle count for the other allowed classes, then stores each detection in the response. The code does not maintain object identifiers between images or video frames.

3.2.5 User Interface. The page has Image Analysis and Video Analysis tabs. Each tab contains a drag-and-drop upload area and a file input. During processing, a loading indicator is displayed. Image results show the annotated image and separate person and vehicle counts. Video results show an HTML video player and a download link for the processed file. JavaScript submits multipart FormData to the corresponding endpoint. There is no live camera capture or progress percentage.

3.3 Media Processing and HTTP Communication. Each operation is initiated by an HTTP POST request carrying a multipart field named file. The client waits for the request to complete. Image analysis returns JSON; video analysis returns binary video content. The application is therefore request/response based rather than a real-time streaming detector. Processing time depends on image dimensions, video length, model speed, machine hardware, and codec support.

3.3.1 Image and Video Processing Workflow. For an image, the route checks the upload's declared content-type prefix, reads bytes, attempts decoding, enhances the array, runs inference, assembles detections and counts, plots boxes, and encodes the plotted image. For video, it validates the content type, writes the uploaded bytes to a temporary file, opens a capture, reads properties, optionally resizes, samples frames, reuses the most recent annotation between samples, writes output frames, releases resources, and returns a file response. The distinct paths produce different response types.

3.3.2 Data Exchange Patterns. The browser sends an image or video file through FormData. Image responses contain success, message, person_count, vehicle_count, detections, and optional annotated_image_base64. A detection includes class_name, confidence, and bbox coordinates. JavaScript turns the base64 image into a data URL. A successful video response is treated as a Blob, converted to an object URL, assigned to the video player, and used by the download link. No request identifier, frame-level video JSON, or persistent result record is returned.

3.4 Data Flow. The user begins with a local file and sees a result in the browser. The server holds uploaded bytes in memory for image processing and uses temporary disk files for video processing. Image arrays move from decoding to enhancement to inference and annotation. Video frames follow the analogous per-frame path only at the sampling interval. Results leave the server as JSON with an embedded image, or as an MP4 file. This flow defines the present application's boundaries and storage behavior.

3.4.1 Image and Video Data Flow. The upload reaches the relevant route. A declared media type that does not match the required family returns an error response. Image decoding failure returns an unsuccessful DetectionResponse with zero counts and an empty detection list. A decoded image goes through the enhancer and detector before the response is assembled. A video that cannot be opened returns an error object. Valid video frames are resized when wider than 640 pixels and inferred on every fifth frame; the initial frame is always processed because no previous annotation exists.

3.4.2 Detection Result Flow. The result object's box collection supplies one entry for each retained prediction. The route translates numeric class IDs into names through result.names and records coordinates and confidence. Person detections increment person_count. Since inference is restricted to one person class and four vehicle classes, remaining returned detections increment vehicle_count. The result plot is encoded as JPEG bytes and then base64 text for JSON transport. Video output instead embeds drawn boxes in frames and does not expose the individual detections as a separate response structure.

3.4.3 Image Request Sequence. The user selects an image; browser JavaScript creates FormData and posts it to the image endpoint. FastAPI invokes the route, which calls the image decoder, enhancer, and detector in sequence. The route creates the response model and serializes it as JSON. Browser code reads the result, updates both count fields, constructs a JPEG data URL, and displays the annotated image. The sequence does not include a database write, queue, external enhancement API, or authentication request.

3.4.4 Activity Diagram. The activity diagram represents the actual high-level path: upload, validate, decode or open, process, infer, annotate, and return. Invalid declared types and unreadable input lead to error handling. On the video path, frame sampling and reuse of the previous annotation are shown explicitly so that the diagram does not imply a fresh detection on every output frame.

3.4.5 System Overview Diagram. The overview distinguishes the browser from the local application and processing services. A user uploads media through the page. The backend coordinates OpenCV and YOLO operations. The browser receives an image response or an MP4 result and presents it for inspection. The drawing summarizes a single application instance and omits production components that are not in the repository.

3.4.6 System Architecture Diagram. The architecture figure maps the browser files, FastAPI app and routes, OpenCV utilities, detector service, Pydantic schemas, and annotated media responses. The image path uses JSON with base64 content; the video path returns a video file. The model and configuration are local resources. The project contains no external API integration for detection, no Flask server, and no cloud service in its implemented architecture.

3.4.7 Use Case Diagram. The primary actor selects image or video analysis, uploads a file, waits for processing, views the annotated result, and downloads processed video where applicable. An API client may submit a multipart request directly. The system responds with a structured image result or an MP4. Neither actor can configure accounts or retrieve stored job history because these functions are not implemented.""",
'Chapter 4':"""This chapter explains the current implementation in terms of model inference, enhancement, API integration, browser behavior, and local operation. The source is a modular prototype rather than a model-training package. Descriptions of settings correspond to values visible in the configuration and service files. Recommendations for additional checks are identified as plans and are not reported as completed experiments.

4.1 Pretrained Detection Model. The application loads a supplied YOLOv8 medium checkpoint and invokes inference when an accepted image or sampled video frame is received. The model is a reusable third-party component. It is not trained as part of the request and its weights are not updated by the application. This deployment arrangement simplifies integration but means performance depends on the checkpoint's original training distribution and on the quality of the submitted input.

4.1.1 Model Overview. ObjectDetector constructs YOLO(settings.MODEL_PATH) during module initialization and exposes a detect method. That method calls the model with the configured confidence threshold and allowed class IDs, then returns results[0]. The route uses the returned result for boxes, names, and plotting. Loading once avoids reconstructing the model for every individual request. It also means that model initialization is part of application startup and a missing weight file can prevent successful operation.

4.1.2 Weights and Class Selection. Configuration points to yolov8m.pt by a relative path. The repository includes YOLOv8 n, s, and m weight files, but the active default is the medium checkpoint. The class list is [0, 2, 3, 5, 7], corresponding to person, car, motorcycle, bus, and truck in COCO. No training or validation dataset is included. The proposal discussed public datasets, training, and standard metrics; those activities are not evidenced by the supplied application code and should not be represented as completed work.

4.1.3 Inference and Evaluation Plan. A request supplies an image array to the detector. The service applies the confidence and class filters, and the API converts retained boxes to a response. Evaluation should use a fixed labeled set not used to select settings. Images should be grouped by source or scene before splitting to reduce near-duplicate leakage. Each class should have enough examples to estimate errors. Predictions should be matched to ground-truth boxes using a stated intersection-over-union threshold, and precision, recall, and average precision should be reported per class. A baseline should use the same weights and threshold on unenhanced inputs. The repository contains no numeric results, so this section specifies a procedure rather than a measured outcome.

4.1.4 Confidence and Runtime Settings. The threshold is 0.15 and is passed directly to Ultralytics. A low threshold can retain weaker predictions and may increase false positives. The allowed class list limits returned categories but does not modify the model weights. Video frames wider than 640 pixels are reduced proportionally; other frames keep their dimensions. Video inference is sampled every five frames. A zero frame-rate property is replaced with 30 frames per second for writing. These values are defaults in code, not results of a documented optimization search.

4.2 Image Preprocessing and Enhancement. Enhancement runs before detector inference for both routes. The pipeline operates on decoded BGR arrays and applies a fixed series of transformations. This approach is easy to reproduce but does not choose parameters based on observed weather. The same transform should be retained when comparing model predictions unless the experiment explicitly studies a different setting.

4.2.1 OpenCV Enhancement Pipeline. First, the implementation creates an intensity lookup table with gamma 1.5 and uses cv2.LUT. The expression raises normalized input values to the reciprocal exponent; for mid-tone values this maps them upward, despite a code comment describing a darkening effect. Second, bilateral filtering smooths local variation while aiming to preserve edges. Third, the image is converted from BGR to LAB, CLAHE adjusts luminance, and the result returns to BGR. Finally, a mild sharpening kernel emphasizes local detail. These operations can change noise, edges, and contrast; their effect must be inspected and measured rather than assumed beneficial.

4.2.2 Input and Output Image Handling. The upload route reads raw bytes. image_utils converts them to a NumPy buffer and asks OpenCV to decode in color mode. A decode failure returns a structured failure result. After inference, the plotting function draws model predictions, and imencode converts the result to JPEG before base64 encoding. Browser code embeds the returned text in a data URL. The response may grow with image dimensions because the entire annotated image is carried inside JSON; the code does not set an explicit maximum upload size or compressed response size.

4.3 FastAPI Integration and Model Serving. FastAPI provides endpoint registration, upload objects, static file service, redirects, and automatic API documentation. Detection and enhancement are imported as services. Request processing is executed in the application process. No asynchronous task queue or job scheduler is present. The image route is declared async because it awaits file reading, but the OpenCV and model inference operations are performed in the same request flow.

4.3.1 Backend Implementation. The router defines POST /detect/image and POST /detect/video and is mounted under /api/v1. The image route checks the declared content type, reads the upload, decodes it, processes the array, and constructs DetectionResponse, DetectionResult, and BoundingBox models. It returns person and vehicle totals, a list of detections, and base64 annotation. The video route writes the incoming file to temporary paths, opens the input, constructs a video writer, loops over frames, and returns FileResponse. The current test_client.py targets /api/v1/detect, a path that is not implemented by this router.

4.3.2 Frontend Implementation. index.html defines the page title, upload zones, image counters, image preview, video player, and download link. style.css supplies a dark glass-style interface, tabs, upload areas, loading indicators, and responsive layout. app.js switches tabs, wires drag-and-drop to file inputs, creates FormData, and sends fetch requests. For image responses it populates the counts and base64 preview; for video it creates a browser object URL for playback and download. Errors are surfaced with a simple alert and the upload zone is restored.

4.4 System Implementation. The repository's module structure makes it possible to trace one upload from browser event through route, service, and response. The model and enhancer are module-level service instances. Settings are centralized rather than repeated in route code. The application currently processes each request within one Python process; multiple workers would each load their own model instance and memory requirements should be considered before increasing worker count.

4.4.1 Backend Development. main.py creates the application and mounts static assets. config.py defines the API prefix, model filename, threshold, and class IDs. detection.py implements media routes. detector.py wraps the Ultralytics model. enhancer.py defines CLAHE and filter operations. image_utils.py handles image encoding and decoding. detection.py under schemas defines typed response objects. This arrangement makes it clearer where to change routing, class filters, image behavior, or response fields. It does not yet include structured logs, centralized exception translation, upload limits, or temporary-file cleanup for all conditions.

4.4.2 Frontend Development. The page offers two tabs and separate file inputs whose accept attributes are image/* and video/*. Upload areas respond to clicks and drag-and-drop. While a request is active, loading text is shown and previous results are hidden. A successful image request updates two counters and an image element. A successful video request creates a Blob URL and sets the player and download link. The code does not show per-frame boxes in a separate table, allow parameter entry, or retain prior result history.

4.4.3 Verification and Error Handling. Code paths can be reviewed for valid image type, invalid image type, decode failure, successful image inference, invalid video type, unreadable video, and unexpected video-processing exceptions. The route has basic guards, but broader integration verification should check the response schema and file playback. The existing test_client.py is not aligned to the implemented route and cannot be treated as an acceptance test without updating its path. No claim of completed benchmark or full automated test coverage is made here.

4.5 Local Deployment and Verification. The repository README describes a local Uvicorn workflow. The model weights must be reachable from the current working directory or the configured model path must be adjusted. Dependencies must be installed in an environment compatible with the Python, OpenCV, PyTorch, and Ultralytics versions. Video support depends on the build's codec capabilities. A production deployment would require additional resource, privacy, and security decisions.

4.5.1 Deployment Environment. The application is started from the project directory with uvicorn app.main:app --reload during development. The root path redirects to the interface; /docs provides interactive API documentation. The reload flag is suited to local development, not a production server. Inference hardware is determined by the installed PyTorch/Ultralytics runtime and machine configuration. The repository does not specify a container image, pinned dependency lock, GPU requirements, or production process manager.

4.5.2 System Verification Plan. Verify the root redirect, static page, interactive docs, both route paths, multipart field name, error responses, image response fields, and processed video format. Exercise corrupt files and media with no target objects. For image responses, verify counts match the returned detections and the base64 value decodes to a viewable JPEG. For video, verify output dimensions, duration, frame rate, and readability. These are recommendations for repeatable verification; no unprovided test output is reported as a result.""",
 'Chapter 5':"""This chapter describes the screens and actions available to a user of the current browser interface. The page is a static front end served by the FastAPI application. It offers separate image and video workflows and presents the results differently because the API returns JSON for images and a file for videos. The instructions below describe the supplied implementation and avoid implying features such as accounts, history, live streaming, or object tracking that are not present.

5.1 Home Page Overview. When a user opens the application root, FastAPI redirects the browser to the static index page. The page heading is AI Object Detection and the summary describes detecting and counting persons and vehicles with enhancement. Two buttons switch between Image Analysis and Video Analysis. The visual layout places the content in a centered panel with a dark background, animated gradient, upload area, and loading and result states. A user may choose a file through the upload prompt or drag and drop a file into the corresponding zone.

5.2 Detection Result Interface. The image result view contains a person counter, a vehicle counter, and a preview element showing the annotated input. These totals come from the image response. The route returns additional per-detection class names, scores, and coordinates, although the page does not display them in a detailed table. The video result view contains an HTML video player and a download link for the processed MP4. The interface does not display a measured accuracy score or a confidence summary for the full video.

5.3 Navigation and Layout. Tabs are switched in JavaScript by toggling active classes on buttons and content panels. Each panel has a separate file input and upload zone. Once processing begins, the upload zone and prior result panel are hidden and a loading message is shown. At completion the loading element is hidden. On a failed image response the page alerts the user and restores the upload area. A network error also triggers an alert and restores the upload. The current code does not show upload progress, estimated wait time, or cancellation.

5.4 Interface Design Principles. The UI groups choices by media type, uses large clickable drop zones, and changes the visible state to indicate that a request is being processed. The dark background and translucent card are implemented in CSS. Result images are displayed within a preview container. The design is intended to make a local prototype approachable, but it should not obscure uncertainty about model predictions. A box is a prediction and a count can be wrong. Clear labeling and failure feedback are important because the browser otherwise has limited information about why a result is missing.

5.5 Accessing the System. From the project directory, install the packages in requirements.txt into a compatible Python environment and ensure yolov8m.pt is available at the configured relative path. Start the development server with uvicorn app.main:app --reload. Open the local address printed by the server; the root route redirects to the UI. The interactive API schema is at /docs. The reload option is for development. A public deployment needs a production server setup, input limits, resource control, and any authentication required by the deployment environment.

5.6 Uploading Image and Video Files. In Image Analysis, choose or drop an image file. The browser sends the file as the multipart field file to POST /api/v1/detect/image. The server checks the declared media type, decodes the image, enhances it, and runs the detector. In Video Analysis, choose or drop a video file; the browser posts to /api/v1/detect/video. The server saves the upload to a temporary file, reads frames, processes selected frames, and returns an MP4. Processing time varies with input duration, dimensions, codec, hardware, and model speed. The code does not specify an upload-size limit.

5.7 Interpreting Detection Results. Person and vehicle counters summarize returned image detections. The class name specifies the model's predicted category, the confidence field is the model score, and bbox contains x1, y1, x2, y2 coordinates. A score above the threshold is still not a calibrated probability or guarantee. Counts may include false positives and omit missed objects. The annotated video reuses each sampled frame's boxes for intervening frames, so a rectangle may not stay aligned with a moving object. There are no track IDs, unique-person counts, or frame-by-frame structured results.""",
 'Chapter 6':"""This chapter states what can be established from the repository and defines how performance should be assessed. The project includes a functioning code path for uploading media, applying fixed enhancement, running a pretrained detector, and returning annotations. The supplied materials do not include a versioned labeled evaluation set, reproducible metric results, or timing measurements. For that reason this chapter presents verification steps and an evaluation protocol rather than invented experimental findings.

6.1 Verification and Evaluation Status. Source inspection confirms the route paths, content-type checks, image decoder, enhancement sequence, YOLO configuration, image response model, video resize threshold, frame sampling interval, and frontend result handling. The configuration sets confidence to 0.15 and uses class IDs 0, 2, 3, 5, and 7. These are implementation facts, not measurements of model quality. A successful API response proves that a code path returned output for an input; it does not prove that the predicted class or box was correct.

6.1.1 Detection Performance Evaluation Plan. Build a labeled evaluation set containing images and video frames from clearly documented sources and conditions. Keep related frames or near-duplicate images in one split to avoid leakage. Annotate visible instances of the five supported classes and record occlusion and visibility where possible. Run the same weights on original inputs and enhanced inputs with a fixed threshold. Match predictions to labels using a declared intersection-over-union criterion, then calculate precision, recall, average precision, and mean average precision per class. Report sample counts, confidence threshold, image sizes, library versions, and hardware. The repository supplies no completed metric table.

6.1.2 Enhancement Evaluation. Compare three paths when data allows: unmodified input with YOLO, enhanced input with YOLO, and a manually inspected sample that identifies enhancement artifacts. Use identical weights and detector settings for the first two paths. Group results by low light, glare, fog, rain, blur, and clear conditions only when labels support those categories. Report examples where enhancement recovers a detection and where it removes or distorts useful evidence. Because the gamma lookup implementation brightens midtones, include exposure and glare analysis rather than relying on its source comment.

6.1.3 Confidence Threshold Evaluation. The configured value 0.15 controls which predictions are retained. Evaluate a threshold sweep on validation data, choose a value according to a stated operating objective, and then report final results once on a held-out test split. A lower setting may retain more true detections but can add false positives. Scores should be calibrated only through a separate calibration procedure. The prototype exposes no threshold control in the UI, so a new value currently requires a configuration change.

6.1.4 Processing Time Evaluation. Measure model startup separately from per-image and per-video request time. Record file dimensions, video duration, frame count, codec, CPU/GPU, package versions, and whether inference runs on CPU or GPU. For video, distinguish source frame rate from the actual inference frequency and the output-writing rate. Repeat measurements and report a median and range or other stated statistic. No numerical speed claim is supported by the supplied project files.

6.1.5 Web Application Verification. Exercise a valid image, an invalid declared type, unreadable image bytes, an image with no target objects, and examples containing each supported class. Confirm the response schema, count consistency, coordinate validity, and JPEG decoding. For video, test supported and unsupported media types, unreadable files, short clips, wide frames, and clips whose reported frame rate is zero. Confirm output playback and download. The existing test_client.py calls a route that differs from the implemented image endpoint, so it must be corrected before being used as a client check.

6.2 Evaluation Protocol. A reproducible study should freeze the checkpoint, package environment, class filter, enhancement values, and threshold before a test run. It should define a baseline, keep splits separate by scene or source, record every input and output, and explain how failed media are handled. Metric calculations should use the same matching rule for every condition. Qualitative examples should complement, not replace, numeric measures. If no suitable local adverse-weather dataset is available, the report should state that limitation rather than generalize from unrelated imagery.

6.2.1 Representative Image and Video Sources. Include daylight and nighttime scenes, clear and adverse visibility, different object sizes and distances, multiple classes, crowded scenes, and occlusion when a source set supports these categories. Record provenance and usage rights. Sample video frames consistently and prevent frames from the same clip appearing across training, validation, and test partitions if future training is performed. Public datasets may be considered after checking license, class taxonomy, weather labels, and suitability for the project's intended scene. The current repository contains no such curated dataset.

6.3 Findings from the Implemented Prototype. The source demonstrates an integrated browser-to-API path, fixed image enhancement, pretrained inference, image counts and boxes, and video file generation. Configuration and service boundaries are identifiable, which supports future parameter tests. No evidence in the repository establishes detection accuracy, a gain from enhancement, robust operation under fog or rain, or real-time processing speed. These outcomes should remain open questions until measured.

6.4 Limitations and Challenges. The system inherits limitations from its pretrained model, fixed preprocessing, API design, and host runtime. Image and video requests have different responses. The code does not enforce file-size or duration limits, and the video path has temporary-file and codec considerations. The prototype has no identity tracking, database, access control, or deployment monitoring. These boundaries should be considered before using it with private or operational footage.

6.4.1 Technical Limitations. Enhancement parameters are fixed. A video frame is resized only when wider than 640 pixels, and detection is run on every fifth frame; intermediate output frames reuse the most recent annotation. The gamma expression does not implement the darkening described by its comment. The confidence threshold is not selected from a documented validation study. The image response embeds a complete JPEG in JSON, which can be inefficient for large files. Exception handling and temporary-file cleanup should be strengthened, and test_client.py uses a stale endpoint path.

6.4.2 Operational Limitations. Model memory and inference throughput depend on the machine and library build. Video codecs may not be available in every environment. There is no authentication or rate limit, so a shared deployment would need controls for who can submit media and how much compute they can consume. Temporary media should be removed securely when no longer needed. Operators should set retention rules and protect files because uploaded footage may contain identifiable people or vehicles.

6.5 Future Implications. The implementation can support a controlled research study and can be extended with evaluation datasets, more transparent thresholds, error analysis, or tracked video. Each extension should be measured against a requirement. A user interface can make model predictions easier to inspect, but does not itself make them reliable. Responsible use depends on documented data handling, uncertainty communication, and appropriate human review.""",
 'Chapter 7':"""This chapter summarizes the prototype and identifies work that would improve its evidence, reliability, and maintainability. Conclusions are limited to the source code and project materials inspected. No experimental conclusion about weather robustness is drawn without a labeled test set and measured comparison.

7.1 Conclusion. The project provides a modular application that accepts uploaded images and videos, applies fixed OpenCV enhancement, and runs pretrained YOLOv8 inference for selected people and vehicle classes. The image route returns counts, bounding boxes, confidence values, and an annotated image; the video route returns an MP4 with sampled-frame annotations. A browser interface provides file upload and result display. The delivered code demonstrates integration of the components, but it does not include weather-specific training or a quantitative validation study. The project's objective is therefore met as a functioning prototype, while its accuracy and robustness remain to be evaluated.

7.2 Contributions. The code organizes route handling, configuration, schemas, image utilities, enhancement, detection, and static presentation into separate modules. It defines two media workflows and records concrete operating settings in configuration and route logic. It also provides a direct way to inspect predictions through an annotated image and to receive processed video. The documentation clarifies the distinction between pretrained inference and model training and gives a procedure for testing the claimed use case.

7.3 Lessons Learned. Integration is only one part of a computer-vision project. A working detector call does not quantify correctness, and an enhancement filter's name or visual effect does not prove better detections. The route and frontend must agree on endpoint names and response types; the supplied test client currently demonstrates the risk of an outdated route. Video processing needs explicit choices about frame size, sampling, codec, and temporal consistency. Documentation is strongest when it identifies which properties come from code and which require measurement.

7.4 Recommendations. Correct the stale client route, pin compatible dependencies, record the checkpoint checksum, and add functional checks for both media paths. Validate file content beyond its declared MIME type, constrain upload bytes and video duration, handle exceptions consistently, and guarantee temporary-file cleanup. Select thresholds on validation data rather than by intuition. Compare enhanced and unmodified inference on the same held-out samples. Preserve raw inputs and predictions only when justified by a documented retention policy. For shared deployment, add authentication, concurrency limits, resource monitoring, and clear user guidance about uncertain predictions.

7.5 Future Work. A first extension is a labeled dataset representing target scenes and visibility conditions, with documented source licenses and annotation rules. A second is a controlled comparison of enhancement settings and detector thresholds using precision, recall, and mAP. A third is a tracked video pipeline that associates detections between frames and reports when annotations were refreshed; tracking must not be confused with the current frame-reuse behavior. Other practical work includes file-size validation, cleanup, job progress, cancellation, version-pinned packaging, and a production deployment guide. Each change should be evaluated against documented requirements and reported with reproducible measurements.""",
}

SUPPLEMENT={
 'Chapter 1':"""Chapter 1 defines the motivation and limits. Chapter 2 provides the computer-vision background. Chapter 3 follows media through the system architecture. Chapter 4 records the implementation and settings. Chapter 5 explains the interface. Chapter 6 separates code verification from performance evaluation. Chapter 7 concludes with recommendations and future work. Together these chapters describe a prototype and make clear which results require further evidence.""",
 '1.1 Problem Statement and Motivation':"""The specific engineering problem is to connect input validation, image enhancement, detector inference, and clear result delivery in one reproducible application. The system also needs to distinguish still-image output from video-file output because their data and processing paths differ.""",
 '1.1.1 Motivation for the Solution':"""The web interface lowers the effort required to submit media and inspect an annotation. Separating services makes it possible to test enhancement and model inference independently later. The project can therefore serve both as a usable demonstration and as a basis for a measured experiment.""",
 '1.2 Scope':"""The API is designed for local use with one uploaded file per request. It does not maintain a session, store results in a database, or return an asynchronous job identifier. These boundaries shape the user workflow and expected response format.""",
 '1.2.1 Inclusions and Exclusions':"""A single image response contains structured detections and a rendered image. A video response is an MP4 artifact without a parallel table of frame-level predictions. This distinction should be retained in requirements and any future user guide.""",
 '1.3 Aims and Objectives':"""The implementation also aims to keep each main responsibility easy to locate in the repository, so that a developer can identify where routes, preprocessing, model settings, and browser behavior are defined.""",
 '1.4 Limitations':"""The video route stores temporary media on disk and depends on OpenCV's ability to read and write the selected format. Its error handling should be reviewed for partial files and cleanup. Image uploads are read into memory without an explicit byte limit. Concurrent requests can compete for model and CPU or GPU resources. These are practical prototype limits, not measured failure rates.""",
 '1.5 Intended Audience and Target Users':"""The report is useful to reviewers checking requirement coverage, maintainers tracing a route, and future students designing a dataset or experiment. Readers should use the references for general technical background and the source repository for the authoritative implementation details.""",
 'Chapter 2':"""The discussion covers detector terminology, learned visual features, adverse visibility, enhancement, and the selected software stack. It does not present a systematic literature review or a claim of novelty because the supplied project materials contain no complete search protocol.""",
 '2.2 Deep Learning for Vision':"""The detector receives an image representation and produces predictions after the framework's own preprocessing. The application then translates those predictions into domain-facing counts and boxes. This separates neural-network inference from the response contract consumed by the browser.""",
 '2.2.2 Deep Learning Architectures for Image Analysis':"""The model filename identifies a checkpoint variant, while the configuration and library version determine how it is loaded. A reproducible deployment should record both, plus a checksum, because a filename alone does not prove that two machines use identical weights.""",
 '2.4.4 Supporting Libraries and Tools':"""The exact environment matters: OpenCV codecs, PyTorch device availability, and Ultralytics versions can influence whether media opens and how quickly inference runs. A future lock file would improve reproducibility compared with the current unpinned requirements list.""",
 'Chapter 3':"""The architecture is intentionally small: requests enter one application process and the application returns a result directly. Each media path has a clear start and end, which makes it possible to trace input bytes, arrays, predictions, and returned output.""",
 '3.1.1 High-Level Design':"""The image route returns after inference and serialization have completed. The video route runs its frame loop before returning a file. Neither route provides a background job model, and browser interaction waits for the request rather than polling progress.""",
 '3.1.2 Architectural Layers':"""Configuration is a shared source for the API prefix, checkpoint location, threshold, and class filter. Pydantic response models make image output shape explicit. Static files are served by the same application, which keeps local setup simple but couples the interface to the backend process.""",
 '3.3.2 Data Exchange Patterns':"""The browser uses a multipart form field named file. Image bytes are embedded in the JSON response as base64, while video bytes are delivered as a file. Clients must therefore parse the image response as JSON and the video response as binary media.""",
 '3.4.7 Use Case Diagram':"""The diagram limits actors to a user and an API client.""",
 'Chapter 4':"""Implementation choices can be followed from settings to the route and then to the service. The following discussion records the code's current behavior, while recommended validation steps are explicitly described as future checks.""",
 '4.1.1 Model Overview':"""The class filter is passed to inference and limits the returned categories. No project-specific classifier is layered on top of YOLO output. The route treats the selected non-person classes as vehicles, which is valid under this fixed allow-list but would need revision if additional classes were enabled.""",
 '4.1.2 Weights and Class Selection':"""The proposal mentions public data and evaluation, but the inspected application tree contains only model weight files and inference code. Dataset selection, labels, licensing, training splits, and annotation policy remain unspecified. They should be established before any fine-tuning or quantitative project claim.""",
 '4.1.3 Inference and Evaluation Plan':"""A held-out split should be separated by original video or capture session rather than random adjacent frames. This avoids having visually near-identical frames in both tuning and evaluation. Report per-class scores and the number of ground-truth instances so that a large class does not hide poor performance on a smaller class.""",
 '4.5.1 Deployment Environment':"""A deployment note should identify the Python version, package versions, checkpoint checksum, operating system, compute device, and available video codecs. Without those details, another machine may behave differently even when the route code is unchanged.""",
 'Chapter 5':"""The user workflow begins in the browser and ends when the response is displayed or downloaded. The interface does not ask the user to choose a model or threshold, so the configured defaults determine the analysis. The following sections explain each visible state and the meaning of its output.""",
 '5.1 Home Page Overview':"""The image and video tabs are separate views but share the same upload-zone pattern. Each zone reacts to a click by opening its hidden file input and also accepts dropped files. The selected file is processed immediately; the page does not show a separate confirmation screen. While the request is pending, the matching loading element is displayed and the old result is hidden. On completion, the response handler reveals the result panel. This interaction keeps the number of steps small, but a user cannot cancel a long video request. The page currently has no progress percentage, queue position, or estimated completion time. A future version could add those states without changing the detector service, provided the backend also exposes job progress.""",
 '5.2 Detection Result Interface':"""The browser displays only the two totals and the annotated image, although the API returns a richer list. A future interface could show the class, score, and coordinates for each detection, with a way to inspect overlapping boxes. Such a table should make clear that confidence is a model score and not a guarantee of correctness. For video, the page offers playback and a download link but does not summarize object counts per frame.""",
 '5.3 Navigation and Layout':"""The state changes are implemented in the browser rather than through full-page navigation. Clicking a tab removes active classes from all tabs and panels, then activates the requested pair. The image and video upload handlers share a helper that controls loading and result visibility. The helper posts FormData to a route supplied at initialization. Image responses are parsed as JSON; video responses are read as a Blob. If the server returns a non-success HTTP status, the client raises an error and alerts the user. This design handles the basic local workflow, but the current alerts are generic and do not expose a recovery explanation. A more accessible interface could announce errors through a live region, keep focus in a predictable place, and provide a retry action.""",
 '5.6 Uploading Image and Video Files':"""The server determines acceptance from the declared content type prefix and then attempts to decode or open the bytes. The browser's accept attribute helps file selection but is not a security validation. Users should select files that their installed OpenCV build can decode. For video, output is generated as MP4 and may take substantially longer than image analysis because frames are read and written across the clip. The system does not display a per-frame object table or preserve upload history after the request.""",
 '5.7 Interpreting Detection Results':"""A blank or unexpected result may reflect a model limitation, a threshold choice, image degradation, or an unsupported class; it should not automatically be interpreted as proof that no person or vehicle was present. Users should keep the original media for comparison and review the annotated output in context.""",
 'Chapter 6':"""A useful evaluation must test both software behavior and recognition quality. These are related but distinct questions: the API can return a valid result while the model prediction is wrong. The verification plan therefore checks response contracts separately from labeled-data metrics.""",
 '6.1.1 Detection Performance Evaluation Plan':"""Use the same class taxonomy when labeling and scoring data. If a category is absent from an image, confirm that it is truly absent rather than merely difficult to see. Document annotation policy for partly visible vehicles, reflections, and truncated people. Evaluate the complete pipeline and the raw-image baseline on identical samples, then retain false-positive and missed-object examples for qualitative review.""",
 '6.2.1 Representative Image and Video Sources':"""When public datasets are considered, verify that their weather labels and class definitions match the experiment. Record any filtering, frame sampling, or remapping. Do not combine related frames across splits, and do not describe a collection as weather-specific unless its provenance supports that description.""",
 '6.3 Findings from the Implemented Prototype':"""The code also makes several limitations visible: the confidence threshold is fixed in configuration, image annotations are embedded in JSON, and video predictions are not exposed as structured frame records. These observations support concrete engineering work but are not estimates of accuracy or speed.""",
 '6.5 Future Implications':"""If future results support a useful operating range, report where those results apply and how often the system fails. Differences across conditions and object sizes should remain visible in the report rather than being hidden behind one aggregate score.""",
 'Chapter 7':"""The recommendations that follow prioritize measurable correctness, predictable media handling, and transparent results. They keep future model research separate from application hardening.""",
 '7.1 Conclusion':"""The application should be described as an integration prototype because that is what the source demonstrates. Its current outputs make predictions inspectable, but they do not certify road safety, surveillance suitability, or performance on a specific weather domain. Those conclusions require additional data and testing.""",
 '7.2 Contributions':"""The modular files also provide a useful teaching example of how a pretrained vision model can be wrapped in an HTTP API and a static interface. The integration makes the inference parameters and media-handling steps available for review. The report records the route mismatch in the sample client so that future maintenance can correct it.""",
 '7.3 Lessons Learned':"""The evaluation plan should be designed before collecting results so that class definitions, splits, and thresholds are not adjusted after seeing test performance. Stable input-output contracts also make it easier to compare model changes without rewriting the frontend.""",
 '7.4 Recommendations':"""Add file-size, pixel-count, frame-count, and duration limits before accepting untrusted uploads. Verify the actual decoded media rather than relying only on a browser-provided MIME type. Use temporary directories with deterministic cleanup and ensure the video writer is released on every exit path. Return structured errors that distinguish unsupported type, decode failure, and processing failure without leaking stack traces. Add endpoint and schema checks to an automated verification suite. For deployment, document memory and concurrency expectations because each application process may hold its own model. Record the threshold, allowed class IDs, weight checksum, and preprocessing settings alongside every reported metric. Keep source media private and define retention before storing any user files.""",
 '7.5 Future Work':"""Future changes should be prioritized from requirements and evaluated before being described as completed. A model update should not be merged into a performance report without a stable baseline and repeatable metrics.""",
}

# Fill the remaining gaps against the reference report's section lengths with
# project-specific detail, keeping complete sentences and avoiding repeated prose.
SUPPLEMENT['Chapter 1'] += ' This report separates implementation from future research.'
SUPPLEMENT['1.2 Scope'] += ' Requests stay on this host.'
SUPPLEMENT['1.2.1 Inclusions and Exclusions'] += ' No accounts exist.'
SUPPLEMENT['1.3 Aims and Objectives'] += ' Each objective can be checked through source review, repeatable requests, and a labeled evaluation.'
SUPPLEMENT['1.4 Limitations'] += ' Results depend on hardware, media quality, and codecs.'
SUPPLEMENT['2.1.1 Overview of Object Detection'] = 'Boxes pair categories with locations for visual review after model inference.'
SUPPLEMENT['2.3.2 Limitations of Current Approaches'] = 'Fixed filters may help scenes and harm others.'
SUPPLEMENT['2.4.1 YOLOv8 Object Detection'] = 'The selected checkpoint determines learned categories, while configuration controls the classes retained in the returned prediction set during each request.'
SUPPLEMENT['3.1 Architecture Overview'] = 'Both routes call shared image services.'
SUPPLEMENT['3.2.2 OpenCV Integration'] = 'Capture and writer objects must close after success or processing failure.'
SUPPLEMENT['3.4.4 Activity Diagram'] = 'Separate branches show image decoding and video capture, while error paths return when types are unsupported or media cannot open.'
SUPPLEMENT['5.3 Navigation and Layout'] += ' This keeps both interface views within the same page.'
SUPPLEMENT['6.3 Findings from the Implemented Prototype'] += ' No weather benchmark scores appear in the current project repository tree.'
SUPPLEMENT['Chapter 7'] += ' Testing remains necessary.'
SUPPLEMENT['7.1 Conclusion'] += ' An image result makes model output visible and supports inspection of counts and boxes. Video output provides a processed file, but sparse inference and frame reuse limit temporal interpretation. Before practical adoption, the application needs labeled evaluation, input safeguards, and documented operating conditions. The current evidence supports integration only at present.'
SUPPLEMENT['7.2 Contributions'] += ' These modules give maintainers clear code modification points.'
SUPPLEMENT['7.3 Lessons Learned'] += ' Claims should match available evidence, and every route should be tested with valid, malformed, empty, and unsupported media before local release.'
SUPPLEMENT['7.5 Future Work'] += ' Record each test environment.'

# Replace the front matter without changing its paragraph positions or styles.
FRONT={
 0:'Robust Object Detection System Under Adverse Weather Conditions',
 28:'APPROVAL CERTIFICATE',
 34:'It is certified that the project work presented in this report entitled “Robust Object Detection System Under Adverse Weather Conditions,” submitted by Abdul Qahir Jalali (Roll No. 53) and Raja Sharyar Muneer (Roll No. 83) of Session 2021–25 under the supervision of Dr. Maryam Bibi, is considered adequate in scope and quality for the Bachelor of Science in Computer Science, subject to approval by the competent university authorities.',
 60:'ABSTRACT',
 62:'Robust Object Detection System Under Adverse Weather Conditions is a web-based prototype for detecting selected people and vehicles in uploaded images and videos. It combines a fixed OpenCV enhancement sequence with pretrained YOLOv8 inference and a browser interface. The image route returns the predicted class, confidence score, bounding-box coordinates, person and vehicle counts, and an annotated image. The video route processes uploaded frames and returns an annotated MP4 file.',
 63:'The backend uses FastAPI and is divided into routes, configuration, response schemas, image utilities, and enhancement and detection services. Enhancement applies an intensity lookup table, bilateral filtering, CLAHE on the luminance channel, and mild sharpening. For video, wide frames are resized to a maximum width of 640 pixels and inference is performed on every fifth frame; the most recent annotation is reused between sampled frames. These are implementation settings rather than results of a parameter-optimization study.',
 64:'The detector loads local yolov8m.pt weights, uses a confidence threshold of 0.15, and filters predictions to the COCO person, car, motorcycle, bus, and truck classes. The code demonstrates pretrained inference; it does not contain a project-specific training pipeline, labeled weather dataset, benchmark results, or object tracking. The project therefore provides an integrated prototype while leaving detection accuracy and the effect of enhancement to a controlled evaluation.',
 65:'Keywords: object detection, YOLOv8, adverse visibility, OpenCV, FastAPI, image enhancement, video processing, prototype validation protocol.',
 70:'We certify that this report describes our project work. Technical sources used in the background and implementation discussion are acknowledged in the references. Statements about implemented behavior are based on the project files available for review; performance conclusions are not claimed without measured evaluation.',
 75:'(Student Signatures)',
 76:'Abdul Qahir Jalali         Raja Sharyar Muneer',
 77:'2021-UMDB-001017         2021-UMDB-001041',
 82:'All praise is due to Allah, the Lord of all the worlds. We are grateful for the strength and opportunity to complete this project and its documentation.',
 83:'We extend our sincere gratitude to our supervisor, Dr. Maryam Bibi, for her guidance, encouragement, and support throughout our academic work. We also thank the faculty members whose teaching helped us develop the knowledge needed to plan and implement this project.',
 84:'We are thankful to our parents, families, and friends for their patience and support during the project. Their encouragement helped us complete the work and prepare this report.',
}

FIXED={
 'Introduction':'Introduction',
 'Background and Technologies':'Background and Technologies',
 'System Design':'System Design',
 'Implementation':'Implementation',
 'User Interface and Features':'User Interface and Features',
 'Results and Discussion':'Verification and Evaluation',
 'Conclusions and Recommendations':'Conclusions and Recommendations',
 'Training parameters included:':'Inference configuration:',
 'Batch size: 4':'Weights: yolov8m.pt',
 'Learning rate: 1e-4':'Confidence: 0.15',
 'Epochs: 20':'Class IDs: 0, 2, 3, 5, 7',
 'Evaluation Metrics:':'Evaluation measures:',
 'The process includes:':'The enhancer applies:',
 'Resampled to 16,000 Hz':'Video width limit: 640 px',
 'Key features include:':'Interface features include:',
 'The environment setup included:':'The local environment requires:',
 'Screenshot Example: ':'Image interface state:',
 'Screenshot Examples:':'Result interface states:',
 'Figure 3.\u200e0.3: Activity Diagram':'Figure 3.4: Activity Diagram',
 'Figure 6.1: detection performance':'Figure 6.1: Validation and Error Paths',
}

ABBREVIATIONS=[
 'AI  Artificial Intelligence','API  Application Programming Interface',
 'CLAHE  Contrast Limited Adaptive Histogram Equalization',
 'COCO  Common Objects in Context','CPU  Central Processing Unit',
 'GPU  Graphics Processing Unit','HTTP  Hypertext Transfer Protocol',
 'IoU  Intersection over Union','JSON  JavaScript Object Notation',
 'MP4  MPEG-4 Part 14','YOLO  You Only Look Once',
]

REFERENCES=[
 '[1] J. Redmon, S. Divvala, R. Girshick, and A. Farhadi, “You Only Look Once: Unified, Real-Time Object Detection,” Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition, 2016. https://arxiv.org/abs/1506.02640',
 '[2] T.-Y. Lin et al., “Microsoft COCO: Common Objects in Context,” European Conference on Computer Vision, 2014. https://arxiv.org/abs/1405.0312',
 '[3] K. Zuiderveld, “Contrast Limited Adaptive Histogram Equalization,” in Graphics Gems IV, Academic Press, 1994, pp. 474–485.',
 '[4] C. Tomasi and R. Manduchi, “Bilateral Filtering for Gray and Color Images,” Proceedings of the Sixth International Conference on Computer Vision, 1998, pp. 839–846.',
 '[5] Ultralytics, “Predict Mode,” YOLO Documentation. https://docs.ultralytics.com/modes/predict (accessed September 2026).',
 '[6] FastAPI, “Tutorial – User Guide,” https://fastapi.tiangolo.com/tutorial/; OpenCV, “OpenCV Documentation,” https://docs.opencv.org/ (accessed September 2026).',
]

FIGURE_CAPTIONS={
 'Figure 3.\u200e0.1: Data Flow Diagram':'Figure 3.1: Image and Video Data Flow',
 'Figure 3.\u200e0.2 : Sequence Diagram':'Figure 3.2: Image Request Sequence',
 'Figure 3.\u200e0.3: Activity Diagram':'Figure 3.3: Media Processing Activity',
 'Figure 3.\u200e0.4: System Overview Diagram':'Figure 3.4: System Overview',
 'Figure 3.\u200e0.5: System Architecture Diagram':'Figure 3.5: Application Architecture',
 'Figure 3.\u200e0.6 : use case diagram':'Figure 3.6: User Use Cases',
 'Figure 4.\u200e0.1 : test results':'Figure 4.1: Proposed Evaluation Protocol',
 'Figure 5.\u200e0.1: home page overview':'Figure 5.1: Browser Interface',
 'Figure 5.\u200e0.2: detection result interface':'Figure 5.2: Image Result Interface',
 'Figure 5.\u200e0.3: interpreting detection result':'Figure 5.3: Video Result Interface',
 'Figure 6.1: detection performance':'Figure 6.1: Validation and Error Paths',
}

FIGURE_LIST=[
 'Figure 3.1: Image and Video Data Flow',
 'Figure 3.2: Image Request Sequence',
 'Figure 3.3: Media Processing Activity',
 'Figure 3.4: System Overview',
 'Figure 3.5: Application Architecture',
 'Figure 3.6: User Use Cases',
 'Figure 4.1: Proposed Evaluation Protocol',
 'Figure 5.1: Browser Interface',
 'Figure 5.2: Image Result Interface',
 'Figure 5.3: Video Result Interface',
 'Figure 6.1: Validation and Error Paths',
]

def p_style(p):
    return p.style.name

def has_picture(p):
    return bool(p._p.xpath('.//w:drawing'))

def clean_text(value):
    return re.sub(r'\s+',' ',value or '').strip()

def split_chapter_narrative(chapter_no, units):
    full=CONTENT.get(f'Chapter {chapter_no}','')
    if not full: return {}
    pieces={}
    section_units=[u for u in units if u['chapter']==chapter_no]
    markers=[]
    for u in section_units:
        if u['old']==f'Chapter {chapter_no}': continue
        marker=u['new']+'.'
        pos=full.find(marker)
        if pos>=0: markers.append((pos,marker,u))
    markers.sort(key=lambda x:x[0])
    # Intro text precedes the first numbered heading.
    first=markers[0][0] if markers else len(full)
    head=next((u for u in section_units if u['old']==f'Chapter {chapter_no}'),None)
    if head: pieces[id(head)]=full[:first].strip()
    for k,(pos,marker,u) in enumerate(markers):
        start=pos+len(marker)
        end=markers[k+1][0] if k+1<len(markers) else len(full)
        pieces[id(u)]=full[start:end].strip()
    return pieces

def sentence_words(text):
    return len((text or '').split())

def fit_section_text(raw,key,target):
    raw=clean_text(raw)
    extra=SUPPLEMENT.get(key,'')
    if len(raw.split())<target and extra:
        raw=clean_text(raw+' '+extra)
    # Keep complete sentences wherever possible so the source's many paragraph
    # slots do not split a thought into an awkward fragment.
    sentences=[s.strip() for s in re.split(r'(?<=[.!?])\s+',raw) if s.strip()]
    # A few legacy-length supplements repeat a sentence already included in
    # the section prose. Keep each complete sentence once in the final report.
    seen=set(); unique=[]
    for sentence in sentences:
        key=re.sub(r'\W+',' ',sentence.casefold()).strip()
        if key not in seen:
            seen.add(key); unique.append(sentence)
    sentences=unique
    if not sentences and target:
        sentences=[f'The {key.split(" ",1)[-1].lower()} is described from the current project implementation.']
    count=sum(len(s.split()) for s in sentences)
    while count>target and len(sentences)>1:
        last=sentences[-1]
        # Remove a trailing sentence only when that brings the paragraph group
        # closer to the legacy text length than retaining it.
        if abs((count-len(last.split()))-target) <= abs(count-target):
            sentences.pop(); count-=len(last.split())
        else: break
    if count>target and len(sentences)==1:
        parts=sentences[0].split()
        if len(parts)>target:
            sentences[0]=' '.join(parts[:target]).rstrip(' ,;:')+'.'
    return sentences

def split_words_for_slots(sentences,slots):
    if not slots: return []
    capacities=[max(1,len(p.text.split())) for p in slots]
    chunks=[]; cursor=0
    for i,cap in enumerate(capacities):
        if cursor>=len(sentences):
            chunks.append(''); continue
        left_slots=sum(1 for c in capacities[i+1:] if c>0)
        group=[]; group_words=0
        while cursor<len(sentences):
            if group and len(sentences)-cursor<=left_slots: break
            candidate=sentences[cursor]
            if group and abs(group_words-cap) <= abs(group_words+len(candidate.split())-cap): break
            group.append(candidate); group_words+=len(candidate.split()); cursor+=1
            if group_words>=cap: break
        if not group and cursor<len(sentences):
            group=[sentences[cursor]]; cursor+=1
        text=' '.join(group).strip()
        if text and not text[0].isupper(): text=text[0].upper()+text[1:]
        chunks.append(text)
    if cursor<len(sentences) and chunks:
        chunks[-1]=(chunks[-1]+' '+' '.join(sentences[cursor:])).strip()
    return chunks

def set_paragraph_text(p_xml,text):
    tnodes=p_xml.xpath('.//w:t',namespaces=NS)
    if tnodes:
        tnodes[0].text=text
        if text[:1].isspace() or text[-1:].isspace():
            tnodes[0].set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        for t in tnodes[1:]: t.text=''
    else:
        r=etree.SubElement(p_xml,W+'r')
        t=etree.SubElement(r,W+'t'); t.text=text

def update_prefix(p_xml,text):
    nodes=p_xml.xpath('.//w:t',namespaces=NS)
    if nodes:
        nodes[0].text=text
        # Keep the cached page-number text and the source's tab leaders.
        if len(nodes)>2:
            for n in nodes[1:-1]: n.text=''

def make_document():
    from docx import Document
    d=Document(TEMPLATE)
    paras=d.paragraphs
    assignments={}
    for i,text in FRONT.items(): assignments[i]=text

    # Set chapter and subsection headings and build the source-layout slots.
    units=[]; current=None; chapter_no=None
    for i,p in enumerate(paras):
        style=p_style(p); text=p.text.strip()
        if style.startswith('Heading') and text:
            if style=='Heading 1':
                if text.startswith('Chapter '):
                    try: chapter_no=int(text.split()[1])
                    except Exception: chapter_no=None
                else:
                    chapter_no=None
            new=TITLES.get(text,text)
            assignments[i]=new
            current={'index':i,'old':text,'new':new,'chapter':chapter_no,'slots':[]}
            units.append(current)
            continue
        if current is None or not text: continue
        if style.lower().startswith('toc') or style=='table of figures' or style=='Caption': continue
        # Front matter and reference lists are replaced separately below.
        if current['old'] in ('References','Abbreviations'): continue
        if text in FIXED:
            assignments[i]=FIXED[text]
            continue
        if text in FIGURE_CAPTIONS:
            assignments[i]=FIGURE_CAPTIONS[text]
            continue
        if has_picture(p):
            if 'activity diagram represents' in text.lower():
                assignments[i]='The activity diagram follows validation, preprocessing, inference, annotation, and response handling.'
            continue
        # Existing one-line subsection labels remain but are made current.
        if p_style(p)=='Body Text' and len(text.split())<=5:
            assignments[i]=text
            continue
        current['slots'].append((i,p))

    # Chapter 3–7 narrative is written in section order; split it at the
    # source headings, then fit each block to the legacy paragraph slots.
    chapter_pieces={}
    for no in range(3,8): chapter_pieces.update(split_chapter_narrative(no,units))
    missing=[]; used_words=0
    for u in units:
        if u['chapter'] not in range(1,8): continue
        slots=u['slots']
        if not slots: continue
        if u['old'] in ('References','Abbreviations'): continue
        if u['chapter'] in range(3,8):
            raw=chapter_pieces.get(id(u),'')
        else:
            raw=CONTENT.get(u['new'],CONTENT.get(u['old'],''))
        key=u['new']
        target=sum(sentence_words(p.text) for _,p in slots)
        if not raw:
            missing.append((u['old'],target,len(slots)))
            raw=f"The {key.split(' ',1)[-1].lower()} is described from the current project implementation."
        words=fit_section_text(raw,key,target)
        assigned=split_words_for_slots(words,[p for _,p in slots])
        for (idx,_),value in zip(slots,assigned): assignments[idx]=value
        used_words+=len(words)

    # Rebuild the static contents labels while preserving the original tabs
    # and cached page numbers. Heading locations and pagination remain source-derived.
    prefix_assignments={}
    toc_entries=[p for p in paras if p.style.name.lower().startswith('toc')]
    for i,p in enumerate(toc_entries):
        line=p.text
        label=line.split('\t',1)[0].strip()
        if label in TITLES: prefix_assignments[paras.index(p)]=TITLES[label]
    fig_entries=[p for p in paras if p.style.name=='table of figures']
    for i,p in enumerate(fig_entries[:len(FIGURE_LIST)]):
        prefix_assignments[paras.index(p)]=FIGURE_LIST[i]

    # References and abbreviations keep their source paragraph count and styling.
    ref_heading=next(i for i,p in enumerate(paras) if p.text.strip()=='References')
    abbr_heading=next(i for i,p in enumerate(paras) if p.text.strip()=='Abbreviations')
    ref_idx=[i for i in range(ref_heading+1,abbr_heading) if paras[i].text.strip() and not paras[i].style.name.lower().startswith('toc')]
    for i,text in zip(ref_idx,REFERENCES): assignments[i]=text
    abbr_idx=[i for i in range(abbr_heading+1,len(paras)) if paras[i].text.strip()]
    # Ten slots are available in the legacy layout; join CPU/GPU in one entry.
    abbr_values=[
      'AI  Artificial Intelligence','API  Application Programming Interface',
      'CLAHE  Contrast Limited Adaptive Histogram Equalization',
      'COCO  Common Objects in Context','CPU/GPU  Central/Graphics Processing Unit',
      'HTTP  Hypertext Transfer Protocol','IoU  Intersection over Union',
      'JSON  JavaScript Object Notation','MP4  MPEG-4 Part 14','YOLO  You Only Look Once',
    ]
    for i,text in zip(abbr_idx,abbr_values): assignments[i]=text

    # Patch only the main document XML; page furniture, styles, sections,
    # numbering, relationships, and all other source parts remain untouched.
    with ZipFile(TEMPLATE) as zin:
        root=etree.fromstring(zin.read('word/document.xml'))
        body_paras=root.xpath('//w:body/w:p',namespaces=NS)
        for i,text in assignments.items():
            if i<len(body_paras): set_paragraph_text(body_paras[i],text)
        for i,text in prefix_assignments.items():
            if i<len(body_paras): update_prefix(body_paras[i],text)
        # Describe replacement figures with accessible alternative text.
        rels=etree.fromstring(zin.read('word/_rels/document.xml.rels'))
        relmap={r.get('Id'):r.get('Target') for r in rels}
        alts={f'word/media/image{i}.png':name for i,name in enumerate([
          'Image workflow','System architecture','OpenCV enhancement sequence','Two upload paths',
          'Image response structure','Video frame processing','Evaluation protocol','Browser interface',
          'Image analysis result view','Video result view','Validation and error paths'],start=4)}
        for drawing in root.xpath('//w:drawing',namespaces=NS):
            embeds=drawing.xpath('.//a:blip/@r:embed',namespaces=NS)
            if not embeds: continue
            target='word/'+relmap.get(embeds[0],'').replace('\\','/')
            label=alts.get(target)
            if label:
                for node in drawing.xpath('.//*[local-name()="docPr"]'):
                    node.set('descr',label);node.set('title',label)
        document_xml=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
        with ZipFile(OUTPUT,'w',ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename=='word/document.xml': blob=document_xml
                elif re.fullmatch(r'word/media/image(?:[4-9]|1[0-4])\.png',item.filename):
                    number=int(re.search(r'image(\d+)',item.filename).group(1))
                    blob=(DIAGRAMS/f'image{number}.png').read_bytes()
                elif item.filename=='docProps/core.xml':
                    core=etree.fromstring(zin.read(item.filename))
                    title=core.find('.//{http://purl.org/dc/elements/1.1/}title')
                    subject=core.find('.//{http://purl.org/dc/elements/1.1/}subject')
                    if title is not None:title.text='Robust Object Detection System Under Adverse Weather Conditions'
                    if subject is not None:subject.text='Final Year Project Documentation'
                    blob=etree.tostring(core,xml_declaration=True,encoding='UTF-8',standalone=True)
                else: blob=zin.read(item.filename)
                zout.writestr(item,blob)
    return assignments,missing,used_words

if __name__=='__main__':
    assignments,missing,words=make_document()
    print(f'Created {OUTPUT}')
    print(f'Updated paragraphs: {len(assignments)}; chapter prose words: {words}')
    print('Unmapped sections:',missing)
