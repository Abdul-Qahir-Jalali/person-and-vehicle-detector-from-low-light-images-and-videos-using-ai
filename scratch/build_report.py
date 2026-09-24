from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path

ROOT=Path(r'E:\hammad project')
SRC=ROOT/'project details'/'doccumentation'/'Robust Object Detection System Documentation (Final Master).docx'
OUT=ROOT/'project details'/'doccumentation'/'Robust Object Detection System FYP Documentation.docx'
doc=Document(SRC)
body=doc._element.body
for child in list(body):
    if child.tag != qn('w:sectPr'):
        body.remove(child)

def para(text='', style=None, align=None, bold=False, italic=False):
    p=doc.add_paragraph(style=style)
    if align is not None: p.alignment=align
    r=p.add_run(text); r.bold=bold; r.italic=italic
    return p
def heading(text, level=1): return para(text, f'Heading {level}')
def pagebreak(): doc.add_page_break()
def section(title, paragraphs):
    heading(title,2)
    for t in paragraphs: para(t)
def chapter(n,title,intro,parts):
    pagebreak(); heading(f'Chapter {n}',1); para(title.upper(), 'Heading 2', WD_ALIGN_PARAGRAPH.CENTER, True); para(intro)
    for h,ps in parts: section(h,ps)

# Cover page follows the reference report's centered academic title page.
for _ in range(3): para()
para('Robust Object Detection System Under Adverse Weather Conditions', 'Title', WD_ALIGN_PARAGRAPH.CENTER, True)
para('Final Year Project Documentation', None, WD_ALIGN_PARAGRAPH.CENTER, True)
para('By', None, WD_ALIGN_PARAGRAPH.CENTER)
para('Abdul Qahir Jalali  |  2021-UMDB-001017', None, WD_ALIGN_PARAGRAPH.CENTER)
para('Raja Sharyar Muneer  |  2021-UMDB-001041', None, WD_ALIGN_PARAGRAPH.CENTER)
for _ in range(2): para()
para('Supervisor', None, WD_ALIGN_PARAGRAPH.CENTER, True)
para('Dr. Maryam Bibi', None, WD_ALIGN_PARAGRAPH.CENTER)
para('Associate Professor', None, WD_ALIGN_PARAGRAPH.CENTER)
for _ in range(2): para()
for t in ['DEPARTMENT OF COMPUTER SCIENCES & INFORMATION TECHNOLOGY','FACULTY OF SCIENCE','UNIVERSITY OF AZAD JAMMU & KASHMIR','MUZAFFARABAD','Session 2021–2025']:
    para(t, None, WD_ALIGN_PARAGRAPH.CENTER, t.startswith(('DEPARTMENT','FACULTY','UNIVERSITY')))

pagebreak(); heading('APPROVAL CERTIFICATE',1)
para('It is certified that the project work presented in this report entitled “Robust Object Detection System Under Adverse Weather Conditions,” submitted by Abdul Qahir Jalali (Roll No. 53) and Raja Sharyar Muneer (Roll No. 83), Session 2021–25, under the supervision of Dr. Maryam Bibi, is considered adequate in scope and quality for the Bachelor of Science in Computer Science (BSCS), subject to approval by the competent university authorities.')
para('\n\n____________________________                         ____________________________\nSupervisor                                                        External Examiner')
para('\n\n____________________________\nChairman')
heading('ABSTRACT',1)
para('This project implements a web-based object detection system intended to improve the visibility of people and vehicles in images and videos affected by difficult viewing conditions. The application accepts an image or video through a browser interface, applies a fixed OpenCV enhancement sequence to each processed image or frame, and runs a pretrained YOLOv8 model to identify selected COCO classes. The supported classes are person, car, motorcycle, bus, and truck. The image endpoint returns class names, confidence values, bounding boxes, person and vehicle counts, and an annotated image. The video endpoint returns an annotated MP4 file.')
para('The backend is built with FastAPI and separated into API routes, configuration, Pydantic response schemas, image enhancement and detection services, and image conversion utilities. Image enhancement combines gamma lookup-table adjustment, bilateral filtering, CLAHE on the luminance channel, and mild sharpening. Video processing resizes wide frames to a maximum width of 640 pixels and reuses detections between inference frames to reduce computation. The delivered code uses the YOLOv8 medium weights file, a confidence threshold of 0.15, and pretrained inference; no project-specific model training, curated dataset, quantitative benchmark, or tracking module is evidenced in the repository. Accordingly, the report describes the implemented system and identifies evaluation work needed to establish measured robustness.')
para('Keywords: object detection, YOLOv8, adverse visibility, image enhancement, OpenCV, FastAPI, video processing.')
heading('UNDERTAKING',1)
para('We certify that this report describes the project work completed by us. Sources used for technical background are acknowledged in the references. Implementation details are stated from the project files available at the time of writing.')
para('Abdul Qahir Jalali                                      Raja Sharyar Muneer')
heading('ACKNOWLEDGEMENTS',1)
para('All praise is due to Allah. We express our sincere gratitude to our supervisor, Dr. Maryam Bibi, for her guidance and support during our studies and this project. We also thank the faculty, our families, and friends for their encouragement and assistance.')
heading('TABLE OF CONTENTS',1)
for t in ['Abstract','Undertaking','Acknowledgements','Chapter 1  Introduction','Chapter 2  Background and Related Technologies','Chapter 3  System Analysis and Design','Chapter 4  Implementation','Chapter 5  User Interface and Operation','Chapter 6  Verification, Evaluation and Limitations','Chapter 7  Conclusion and Future Work','References','Abbreviations']:
    para(t)
heading('LIST OF FIGURES',1)
for t in ['Figure 3.1 Use-case diagram','Figure 3.2 System architecture','Figure 3.3 Activity flow','Figure 3.4 Image detection sequence']:
    para(t)

# Match the reference's two page geometries: front matter uses 1-inch side
# margins and the report body uses the source's wider inner/outer margins.
first=doc.sections[0]
first.top_margin=Inches(1); first.bottom_margin=Inches(1)
first.left_margin=Inches(1); first.right_margin=Inches(1)
report_section=doc.add_section(WD_SECTION.NEW_PAGE)
report_section.top_margin=Inches(1); report_section.bottom_margin=Inches(1)
report_section.left_margin=Inches(1.5); report_section.right_margin=Inches(1.25)

chapter(1,'Introduction','This chapter defines the problem, motivation, project scope, objectives, limitations, and intended users. The project is a working prototype for detecting selected people and vehicle classes from uploaded images and videos after image enhancement.',[
('1.1 Problem Statement and Motivation',['Object detectors can miss or localize objects poorly when illumination is uneven, glare is strong, contrast is low, or image detail is obscured. This project explores a practical processing pipeline that applies image enhancement before pretrained object detection and presents results through a simple browser interface.','The implementation provides an operational prototype. Its presence of enhancement stages does not by itself prove increased accuracy under every weather condition; that conclusion requires a controlled comparison against unprocessed inputs and labeled ground truth.']),
('1.1.1 Motivation for the Solution',['People and vehicles are common targets in road monitoring and visual inspection. A single application that accepts still images and video can make model inference accessible to users without requiring them to run command-line tools. Combining preprocessing with detection gives a reproducible workflow that can be evaluated and extended.']),
('1.2 Scope',['The system exposes a local web application and versioned HTTP API. The root page redirects to a static HTML interface. The image route validates the uploaded media type, decodes the image, enhances it, runs inference, and returns structured results with an annotated image. The video route processes uploaded video frame by frame and returns an MP4 output.']),
('1.2.1 Inclusions and Exclusions',['Included: browser-based image and video upload; YOLOv8 pretrained inference; selected COCO person and vehicle classes; image enhancement; person and vehicle totals for image results; bounding boxes and confidence values; annotated image and video output; and modular FastAPI services.','Excluded from the verified implementation: model training or fine-tuning; a project-specific dataset; object identity tracking; live camera streaming; persistent user accounts or databases; cloud deployment; and measured accuracy or throughput guarantees.']),
('1.3 Aims and Objectives',['The aim is to implement an accessible prototype for detecting people and vehicles in uploaded visual media under challenging visibility conditions. Objectives are to integrate a pretrained detector, apply an explicit enhancement pipeline, provide image and video endpoints, return understandable annotated results, and document limitations and a defensible evaluation plan.']),
('1.4 Limitations',['Detection quality depends on the pretrained model, the input image, threshold, and enhancement settings. Enhancement can amplify noise or suppress useful details. Video detection is performed on sampled frames and annotations are reused between inference frames, so boxes may not follow moving objects smoothly. The application currently has no stated upload-size limit, authentication, persistent storage, or production deployment configuration.']),
('1.5 Intended Audience and Target Users',['The report is intended for project evaluators, computer science students, developers extending the prototype, and users who need to submit images or videos for offline analysis. It is not a safety certification or evidence that the system is suitable as the sole basis for autonomous decisions.'])])

chapter(2,'Background and Related Technologies','This chapter introduces object detection and explains the technologies visible in the implementation. Technical background is kept separate from claims about project-specific performance.',[
('2.1 Object Detection and Computer Vision',['Object detection combines classification with localization: for each detected object, a model predicts a class and a bounding box. Confidence values represent model scores and should not be interpreted automatically as calibrated probabilities. Detection results can be affected by scale, occlusion, lighting, motion blur, compression, and domain shift.']),
('2.1.1 Overview of Object Detection',['An object detector processes image pixels and returns candidate boxes with class labels and scores. Post-processing filters predictions by confidence and may remove overlapping candidates. The project delegates this inference and result plotting to the Ultralytics YOLO interface.']),
('2.1.2 YOLO and Pretrained Inference',['YOLO is a family of one-stage detection models designed to predict object locations and classes from images. This project loads the local yolov8m.pt weights and calls inference with a confidence threshold of 0.15 and class IDs [0, 2, 3, 5, 7]. The weights are pretrained; the repository does not include training scripts, project labels, or evidence of fine-tuning.']),
('2.1.3 Challenges in Adverse Visibility',['Low illumination, glare, fog, rain, and blur can reduce contrast or hide boundaries. The current project applies general-purpose image operations before detection. These operations can change pixel appearance but do not reconstruct missing scene information. Their benefit must be measured with representative data rather than assumed.']),
('2.2 Image Enhancement Concepts',['Gamma transformation remaps pixel intensity nonlinearly. Bilateral filtering smooths regions while attempting to preserve edges. CLAHE enhances local contrast in image tiles and is applied to the L channel after conversion to LAB. A convolution kernel then performs mild sharpening. The sequence is fixed in code and is not selected adaptively for particular weather.']),
('2.3 Frameworks and Libraries',['FastAPI defines the HTTP application and routes. Pydantic models describe structured image responses. OpenCV handles decoding, color-space changes, filtering, video capture, resizing, and MP4 writing. NumPy constructs lookup tables and kernels. Ultralytics loads and runs YOLO. The frontend consists of static HTML, CSS, and JavaScript. Uvicorn is used to run the ASGI application.']),
('2.4 Rationale for the Technology Selection',['The selected stack supports a modular Python service and a browser-based interface. The detector, enhancement stage, request handling, and response schemas are separated into files, which makes replacement and testing more manageable. The prototype relies on local model weights and local compute; runtime depends on the host CPU/GPU, image dimensions, video length, and installed libraries.'])])

chapter(3,'System Analysis and Design','This chapter documents user interaction, system components, data movement, and the interfaces implemented in the repository.',[
('3.1 Requirements',['Functional requirements: accept image and video uploads; reject media whose declared content type has the wrong family; decode images; enhance inputs; run selected-class inference; return image counts, detections, and annotation; and generate annotated video output. Non-functional goals include understandable results, modular organization, and reasonable processing cost through video resizing and frame sampling. No numeric service-level target is specified.']),
('3.1.1 Actors and Use Cases',['The primary actor is a user operating the browser interface. The user selects image analysis or video analysis, chooses a local file, submits it, and views or downloads the result. An API client can call the endpoints directly. The server validates the declared content type and processes the uploaded bytes or temporary video file.']),
('3.2 Architecture Overview',['The browser sends an HTTP multipart upload to FastAPI. The API route validates and reads the upload. For images, utility functions decode bytes, the enhancer transforms the image, and the detector runs YOLO inference. The route formats results using Pydantic schemas and encodes the plotted image as base64. For video, OpenCV reads frames, the same enhancement and detection services process sampled frames, and VideoWriter creates the output MP4.']),
('3.2.1 Architectural Layers',['Presentation layer: static HTML, CSS, and JavaScript under app/static. API layer: FastAPI routes under app/api/routes. Service layer: image enhancement and object detection under app/services. Schema layer: response models in app/schemas. Utility layer: image byte decoding and base64 encoding. Configuration layer: model path, API prefix, confidence threshold, and allowed classes in app/core/config.py.']),
('3.3 Image Data Flow',['1. The user selects an image in the browser. 2. The frontend submits the file to POST /api/v1/detect/image. 3. The route checks the declared content type and decodes the image bytes. 4. The enhancer applies gamma adjustment, bilateral filtering, CLAHE, and sharpening. 5. YOLO runs with the configured threshold and class filter. 6. The route extracts labels, scores, boxes, and counts, plots annotations, and returns a JSON response.']),
('3.3.1 Video Data Flow',['The client submits a video to POST /api/v1/detect/video. The server writes the upload to a temporary MP4 file, opens it with OpenCV, reads dimensions and frame rate, and constrains wide frames to 640 pixels. Every fifth frame is selected for enhancement and inference, with an initial inference on the first frame. The latest annotated frame is written repeatedly between inference frames. The resulting MP4 is returned as a file response.']),
('3.4 API and Response Design',['Image route: POST /api/v1/detect/image with a multipart field named file. Successful responses include success, message, person_count, vehicle_count, detections, and optional annotated_image_base64. Each detection contains class_name, confidence, and bbox coordinates x1, y1, x2, y2. Invalid declared media types and undecodable images return a structured unsuccessful response. Video route: POST /api/v1/detect/video with the multipart field file; successful processing returns video/mp4, while invalid or unreadable input returns an error object.']),
('3.5 Design Diagrams',['The diagrams below summarize the implemented components and workflow. They describe the code-level prototype, not a deployed multi-user or cloud architecture.'])])
for label,file in [('Figure 3.1 Use-case diagram','uc.png'),('Figure 3.2 System architecture','arch.png'),('Figure 3.3 Activity flow','act.png'),('Figure 3.4 Sequence diagram','seq.png')]:
    p=para(label, None, WD_ALIGN_PARAGRAPH.CENTER, True)
    img=doc.add_paragraph(); img.alignment=WD_ALIGN_PARAGRAPH.CENTER
    img.add_run().add_picture(str(ROOT/file), width=Inches(5.4))

chapter(4,'Implementation','This chapter maps the design to the files and behavior in the delivered source tree.',[
('4.1 Application Startup',['app/main.py creates the FastAPI app using project settings, registers the detection router under /api/v1, mounts the static directory at /static, and redirects the root URL to /static/index.html. The app version is declared as 1.0.0.']),
('4.2 Object Detection Service',['app/services/detector.py loads the model at startup using the configured weights path. Each detect call supplies the image, confidence threshold 0.15, and allowed class IDs [0, 2, 3, 5, 7]. The first Ultralytics result is returned to the route. Class names are obtained from the result metadata.']),
('4.3 Image Enhancement Service',['app/services/enhancer.py applies a gamma lookup table with gamma set to 1.5, bilateral filtering with diameter 9 and sigma values 75, CLAHE with clip limit 2.0 and an 8 by 8 tile grid on LAB luminance, then a mild sharpening kernel. These fixed parameters are implementation settings, not evidence of optimality for every scene.']),
('4.4 Image Endpoint',['The endpoint uses UploadFile and a multipart file. It checks the content type prefix, reads bytes, decodes through the utility module, enhances the resulting array, and performs inference. Each predicted box is converted into a DetectionResult. Class person increments person_count; all remaining predictions increment vehicle_count, consistent with the configured vehicle-only set beyond person. The plotted output is encoded to base64.']),
('4.5 Video Endpoint',['The video route uses temporary files and OpenCV VideoCapture/VideoWriter. A zero-reported frame rate is replaced with 30. Frames wider than 640 pixels are resized proportionally. Detection is run once every five frames, and the most recent annotated frame is reused for the intervening output frames. The output codec is mp4v and the response filename is processed_video.mp4.']),
('4.6 Frontend',['The static page provides tabs for image and video analysis, file inputs restricted by browser accept attributes, upload zones, loading indicators, image-result counters, and a result preview. The JavaScript file controls interaction and communicates with API routes. The UI does not imply object tracking; the video route repeats sampled-frame annotations.']),
('4.7 Configuration and Dependencies',['Configuration defaults are declared in app/core/config.py. requirements.txt lists FastAPI, Uvicorn, python-multipart, Ultralytics, headless OpenCV, and pydantic-settings. Local model weight files are stored at the project root. A clean deployment must install compatible package versions and ensure the selected model file is present.'])])

chapter(5,'User Interface and Operation','This chapter describes the user-facing workflow and the response users can expect from the current prototype.',[
('5.1 Home Page',['Open the application root after starting the server. The root redirects to the static interface. The page presents two analysis tabs: Image Analysis and Video Analysis.']),
('5.2 Image Analysis',['Choose the image tab, select or drop an image, and submit it through the upload area. While the request runs, the page displays a processing indicator. A successful response presents the person and vehicle counts and an annotated image. The API response also contains individual class labels, confidence scores, and bounding-box coordinates.']),
('5.3 Video Analysis',['Choose the video tab and submit a supported video file. The server processes the upload and returns an MP4 with annotations. Processing time varies with video duration, frame size, hardware, and model inference speed. Because inference runs on sampled frames, the drawn boxes are held between detections and can appear temporally stale.']),
('5.4 Accessing the API',['Run the application with Uvicorn from the project directory using the documented command: uvicorn app.main:app --reload. The interactive API schema is available at /docs while the server is running. Image and video routes are under /api/v1/detect/image and /api/v1/detect/video respectively.']),
('5.5 Interpreting Results',['A bounding box indicates a predicted location; the accompanying class name is the predicted category. Confidence is a model score filtered by the configured threshold. Counts summarize image detections and do not represent unique identities. The current video response contains an annotated file but does not return per-frame metadata or track IDs.']),
('5.6 Input and Error Handling',['The endpoints check the declared media type and report an error for an invalid type. Image decode failure returns an unsuccessful image response. The video endpoint reports an error when the file cannot be opened. Users should provide valid media files and avoid relying on the tool for high-stakes decisions.'])])

chapter(6,'Verification, Evaluation and Limitations','The repository contains implementation files and a basic client script but no reproducible labeled test set or recorded quantitative results. This chapter distinguishes code-verifiable behavior from performance claims and proposes an evaluation method.',[
('6.1 Verification Status',['Source inspection confirms the route definitions, configured YOLO weights path, class IDs, threshold, enhancement operations, image response schema, video resizing, and frame sampling logic. This documentation does not claim that an end-to-end acceptance test, accuracy study, latency benchmark, or adverse-weather comparison has been completed.']),
('6.1.1 Functional Test Plan',['For image requests, test a valid image, a non-image media type, corrupt image bytes, an image with no target objects, and images containing each allowed class. Check response status, schema fields, finite coordinates, class labels, count consistency, and that the annotated image decodes. For video, test a valid short video, invalid media type, unreadable file, zero/unknown frame rate if reproducible, wide and narrow frames, and a clip shorter than five frames. Check that output is a playable MP4 with expected dimensions and duration.']),
('6.1.2 Quantitative Evaluation Plan',['Use a fixed, documented set of representative scenes with ground-truth boxes and class labels. Include clear conditions and adverse visibility categories such as low light, glare, fog, and rain only when examples are available and correctly labeled. Compare the raw-input pipeline with the enhancement-plus-detection pipeline using identical weights and thresholds. Report precision, recall, mAP at stated IoU thresholds, per-class results, sample counts, hardware, model version, and confidence threshold.']),
('6.1.3 Timing and Resource Evaluation',['Measure image latency and video processing time on specified hardware after separating model initialization from per-request inference. Record input dimensions, video duration, effective inference-frame rate, CPU/GPU utilization, and memory use. Repeat each measurement and report summary statistics. No numerical timing result is asserted here because none is supplied by the project evidence.']),
('6.2 Known Technical Limitations',['The model is pretrained and general-purpose; no domain-specific tuning is present. The fixed enhancer may improve contrast in some inputs and degrade others. The video route uses a single temporary-file suffix and codec choice, depends on OpenCV video support, and does not explicitly remove temporary files in all code paths. The route does not expose configurable sampling interval or maximum upload size. In the image route, any model class outside the configured person/vehicle set should be excluded by inference filtering; the current counting logic categorizes every returned non-person class as a vehicle.']),
('6.3 Security and Operational Considerations',['The prototype is intended for local development. Production use should add file-size limits, robust media probing, controlled temporary-file cleanup, request timeouts, concurrency/resource controls, authentication where needed, and explicit handling of processing exceptions. Uploaded media should be handled according to the operator’s privacy requirements.']),
('6.4 Evaluation Conclusion',['The implementation demonstrates an integrated upload, enhancement, detection, and annotation workflow. Whether the enhancement improves detection under adverse weather remains an empirical question. A labeled benchmark and controlled baseline comparison are required before making a reliability or accuracy claim.'])])

chapter(7,'Conclusion and Future Work','The project delivers a modular prototype that combines image enhancement and pretrained object detection behind a browser interface and HTTP API.',[
('7.1 Conclusion',['The implemented system accepts uploaded images and videos, enhances image content using fixed OpenCV operations, and applies YOLOv8 medium pretrained weights to selected person and vehicle classes. It returns structured image detections with bounding boxes and confidence scores, and creates annotated video output by sampling frames. The project meets its prototype objective of integrating these functions into a usable local application.']),
('7.2 Contributions',['The work provides a separated FastAPI application structure, reusable detection and enhancement services, an image response schema, image and video upload routes, and a static browser UI. It also makes preprocessing and video sampling parameters explicit so they can be evaluated and tuned.']),
('7.3 Recommendations',['First establish a versioned benchmark of labeled media, then compare enhanced and unenhanced inference using fixed model settings. Add robust input size and format validation, cleanup of temporary files, error handling, and tests for API response contracts. Document hardware and package versions for repeatability.']),
('7.4 Future Work',['Potential extensions include collecting and labeling representative local road scenes; evaluating alternate enhancement settings; fine-tuning only if suitable data and compute are available; adding object tracking for video; supporting progress and cancellation for long jobs; exposing threshold and sampling controls safely; and packaging a deployment configuration with resource limits. Each extension should be validated against measured requirements.'])])

heading('REFERENCES',1)
for t in [
'[1] Ultralytics, “Model Prediction with Ultralytics YOLO,” https://docs.ultralytics.com/modes/predict (accessed September 2026).',
'[2] FastAPI, “Tutorial – User Guide,” https://fastapi.tiangolo.com/tutorial/ (accessed September 2026).',
'[3] OpenCV, “The OpenCV Tutorials,” https://docs.opencv.org/ (accessed September 2026).',
'[4] Pydantic, “Pydantic Settings,” https://pydantic.dev/docs/validation/latest/api/pydantic_settings/ (accessed September 2026).',
'[5] Project source code, Robust Object Detection System, app/ and requirements.txt, supplied project repository (reviewed September 2026).']:
    para(t)
heading('ABBREVIATIONS',1)
for t in ['API  Application Programming Interface','CLAHE  Contrast Limited Adaptive Histogram Equalization','COCO  Common Objects in Context','CPU  Central Processing Unit','GPU  Graphics Processing Unit','HTTP  Hypertext Transfer Protocol','IoU  Intersection over Union','LAB  Lightness and color-opponent color space','MP4  MPEG-4 Part 14 video container','YOLO  You Only Look Once']:
    para(t)

# Keep the source report's page system, section breaks, header/footer and styles.
doc.core_properties.title='Robust Object Detection System Under Adverse Weather Conditions'
doc.core_properties.subject='Final Year Project Documentation'
doc.core_properties.author='Abdul Qahir Jalali; Raja Sharyar Muneer'
doc.save(OUT)
print(OUT)
