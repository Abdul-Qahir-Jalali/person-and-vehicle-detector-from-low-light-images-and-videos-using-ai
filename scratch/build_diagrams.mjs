import { mkdir } from 'node:fs/promises';
import { createRequire } from 'node:module';
const require=createRequire(import.meta.url);
const sharp=require('sharp');

const out='E:/hammad project/scratch/current_diagrams';
await mkdir(out,{recursive:true});
const specs={
4:[262,1354,'media workflow'],5:[1476,629,'system architecture'],6:[425,826,'enhancement steps'],7:[1175,1042,'image and video paths'],8:[621,508,'image API response'],9:[702,478,'video sampling'],10:[846,814,'evaluation protocol'],11:[1265,616,'browser interface'],12:[1348,665,'image result'],13:[1266,625,'video result'],14:[665,312,'validation and errors']
};
function esc(s){return s.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');}
function svg(w,h,title,body){
 return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="100%" height="100%" fill="white"/><style>text{font-family:Arial, sans-serif;fill:#1f2937}.title{font-size:24px;font-weight:bold}.label{font-size:18px}.small{font-size:15px}.box{fill:#f7f9fc;stroke:#334155;stroke-width:2}.line{stroke:#475569;stroke-width:2;fill:none;marker-end:url(#arrow)}.group{fill:#fff;stroke:#64748b;stroke-width:2}</style><defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="7" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#475569"/></marker></defs><text x="${w/2}" y="42" text-anchor="middle" class="title">${esc(title)}</text>${body}</svg>`;
}
function box(x,y,w,h,lines,cls='label',rx=8){const txt=lines.map((s,i)=>`<text x="${x+w/2}" y="${y+h/2+(i-(lines.length-1)/2)*22+7}" text-anchor="middle" class="${cls}">${esc(s)}</text>`).join('');return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}" class="box"/>${txt}`;}
function arrow(x1,y1,x2,y2){return `<path d="M${x1},${y1} L${x2},${y2}" class="line"/>`;}
function label(x,y,s,cls='small',anchor='middle'){return `<text x="${x}" y="${y}" text-anchor="${anchor}" class="${cls}">${esc(s)}</text>`;}

const drawings={
4:svg(262,1354,'Image workflow',[
 box(45,88,172,96,['Select image'],'small'),arrow(131,184,131,226),
 box(45,226,172,112,['Validate type and','decode bytes'],'small'),arrow(131,338,131,380),
 box(45,380,172,112,['Apply image','enhancement'],'small'),arrow(131,492,131,534),
 box(45,534,172,112,['Run YOLOv8','inference'],'small'),arrow(131,646,131,688),
 box(45,688,172,120,['Collect classes,','scores and boxes'],'small'),arrow(131,808,131,850),
 box(45,850,172,112,['Count people and','vehicles'],'small'),arrow(131,962,131,1004),
 box(45,1004,172,112,['Draw annotations','and encode image'],'small'),arrow(131,1116,131,1158),
 box(45,1158,172,112,['Return JSON to','browser'],'small')].join('')),
5:svg(1476,629,'System architecture',[
 box(42,202,236,150,['Browser interface','HTML · CSS · JavaScript']),
 box(342,202,250,150,['FastAPI application','/api/v1 routes']),
 box(666,112,260,132,['OpenCV utilities','decode · enhance · encode']),
 box(666,354,260,132,['YOLOv8 service','pretrained yolov8m.pt']),
 box(1000,202,250,150,['Pydantic response','boxes · scores · counts']),
 box(1290,202,150,150,['Annotated','media']),
 arrow(278,277,342,277),arrow(592,242,666,178),arrow(592,311,666,420),
 arrow(926,178,1000,242),arrow(926,420,1000,311),arrow(1250,277,1290,277),
 label(615,99,'Image path','small'),label(615,520,'Video frames use sampled inference','small')].join('')),
6:svg(425,826,'OpenCV enhancement sequence',[
 box(58,86,309,98,['Decoded BGR image']),arrow(212,184,212,226),
 box(58,226,309,108,['Gamma lookup table','(configured gamma 1.5)'],'small'),arrow(212,334,212,376),
 box(58,376,309,108,['Bilateral filter','edge-aware smoothing'],'small'),arrow(212,484,212,526),
 box(58,526,309,108,['LAB conversion and CLAHE','on luminance'],'small'),arrow(212,634,212,676),
 box(58,676,309,108,['Mild sharpening kernel','then detector input'],'small')].join('')),
7:svg(1175,1042,'Two upload paths',[
 box(434,78,306,92,['Browser upload'],'label'),arrow(587,170,587,216),
 box(434,216,306,102,['FastAPI validates media type'],'small'),
 `<path d="M587 318 L587 365 L280 365 L280 410 M587 365 L895 365 L895 410" class="line"/>`,
 box(78,410,405,142,['IMAGE','Decode → enhance → YOLOv8'],'label'),
 box(692,410,405,142,['VIDEO','OpenCV capture → sampled frames'],'label'),
 arrow(280,552,280,612),arrow(895,552,895,612),
 box(78,612,405,142,['JSON result','counts · detections · base64 image'],'small'),
 box(692,612,405,142,['MP4 result','annotated sampled-frame output'],'small'),
 arrow(280,754,280,814),arrow(895,754,895,814),
 box(240,814,695,118,['Browser displays image results or video playback'],'small')].join('')),
8:svg(621,508,'Image response structure',[
 box(56,86,509,70,['DetectionResponse'],'label'),
 box(56,186,246,240,['success · message','person_count','vehicle_count','detections[]','annotated_image_base64'],'small'),
 box(320,186,245,240,['DetectionResult','class_name','confidence','bbox { x1, y1, x2, y2 }'],'small'),
 arrow(302,306,320,306),label(310,462,'Coordinates are model output values; no identity tracking is returned.','small')].join('')),
9:svg(702,478,'Video frame processing',[
 box(28,124,150,100,['Read next','frame'],'small'),
 box(204,124,150,100,['Resize if','width > 640'],'small'),
 box(380,124,150,100,['Every fifth','frame: enhance','and detect'],'small'),
 box(556,124,122,100,['Write last','annotated','frame'],'small'),
 arrow(178,174,204,174),arrow(354,174,380,174),arrow(530,174,556,174),
 `<path d="M616 224 L616 314 L454 314 L454 224" class="line"/>`,
 label(536,344,'Intervening frames reuse the latest annotation','small'),
 box(115,374,472,66,['VideoWriter produces MP4 output'],'small')].join('')),
10:svg(846,814,'Evaluation protocol (proposed)',[
 box(222,88,402,88,['Assemble labeled image and video set'],'small'),arrow(423,176,423,222),
 box(222,222,402,88,['Record media conditions, classes and splits'],'small'),arrow(423,310,423,356),
 box(92,356,296,116,['Baseline: original image','same YOLO weights and threshold'],'small'),
 box(458,356,296,116,['Pipeline: enhancement + YOLO','same weights and threshold'],'small'),
 arrow(240,472,310,532),arrow(606,472,536,532),
 box(222,532,402,108,['Compare precision, recall and mAP','by class and visibility group'],'small'),arrow(423,640,423,684),
 box(222,684,402,80,['Report latency and hardware'],'small')].join('')),
11:svg(1265,616,'Browser interface',[
 `<rect x="180" y="85" width="905" height="458" rx="20" fill="#f8fafc" stroke="#334155" stroke-width="3"/>`,
 label(632,137,'AI Object Detection','title'),label(632,170,'Detect and count persons and vehicles','small'),
 box(400,210,205,64,['Image Analysis'],'small'),box(660,210,205,64,['Video Analysis'],'small'),
 box(340,318,585,112,['Drag, drop or browse for media'],'small'),
 label(632,478,'Image: person and vehicle totals · Video: playback and download','small')].join('')),
12:svg(1348,665,'Image analysis result view',[
 `<rect x="120" y="90" width="1108" height="485" rx="18" fill="#f8fafc" stroke="#334155" stroke-width="3"/>`,
 box(175,155,280,102,['Persons','count returned by API'],'small'),box(490,155,280,102,['Vehicles','count returned by API'],'small'),
 box(805,155,365,102,['Detection list','class · confidence · box'],'small'),
 `<rect x="195" y="302" width="950" height="216" fill="#e2e8f0" stroke="#64748b" stroke-width="2"/>`,
 label(670,420,'Annotated input image','label'),label(670,548,'Illustrative interface layout; no sample metrics are asserted.','small')].join('')),
13:svg(1266,625,'Video result view',[
 `<rect x="120" y="70" width="1026" height="480" rx="18" fill="#f8fafc" stroke="#334155" stroke-width="3"/>`,
 `<rect x="235" y="125" width="796" height="310" fill="#e2e8f0" stroke="#64748b" stroke-width="2"/>`,
 label(633,285,'Processed video preview','label'),
 `<circle cx="633" cy="485" r="24" fill="#fff" stroke="#334155" stroke-width="2"/><path d="M625 472 L645 485 L625 498 Z" fill="#334155"/>`,
 box(850,458,160,58,['Download MP4'],'small'),label(633,585,'Annotations are sampled and held between inference frames.','small')].join('')),
14:svg(665,312,'Validation and error paths',[
 box(30,112,144,78,['Upload'],'small'),box(218,112,190,78,['Check content type'],'small'),box(454,112,180,78,['Decode / open'],'small'),
 arrow(174,151,218,151),arrow(408,151,454,151),
 label(316,232,'Invalid type or unreadable media returns an error response.','small'),
 label(332,270,'The current code does not impose an upload-size limit.','small')].join(''))
};
for (const [id,[w,h]] of Object.entries(specs)) {
 const data=Buffer.from(drawings[id]);
 await sharp(data).png().toFile(`${out}/image${id}.png`);
}
console.log('Wrote 11 current-project diagrams.');
