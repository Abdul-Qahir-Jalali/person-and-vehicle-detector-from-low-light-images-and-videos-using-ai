document.addEventListener('DOMContentLoaded', () => {
    const tabs = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    // Tab Switching Logic
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            tab.classList.add('active');
            document.getElementById(tab.dataset.target).classList.add('active');
        });
    });

    // Image Upload Logic
    setupUploadZone('image-upload-zone', 'image-input', '/api/v1/detect/image', handleImageSuccess, 'image-loading', 'image-results');
    
    // Video Upload Logic
    setupUploadZone('video-upload-zone', 'video-input', '/api/v1/detect/video', handleVideoSuccess, 'video-loading', 'video-results');

    function setupUploadZone(zoneId, inputId, endpoint, successCallback, loadingId, resultsId) {
        const zone = document.getElementById(zoneId);
        const input = document.getElementById(inputId);
        const loading = document.getElementById(loadingId);
        const results = document.getElementById(resultsId);

        zone.addEventListener('click', () => input.click());

        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('dragover');
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                input.files = e.dataTransfer.files;
                processFile(input.files[0], endpoint, successCallback, zone, loading, results);
            }
        });

        input.addEventListener('change', () => {
            if (input.files.length) {
                processFile(input.files[0], endpoint, successCallback, zone, loading, results);
            }
        });
    }

    async function processFile(file, endpoint, successCallback, zone, loading, results) {
        const formData = new FormData();
        formData.append('file', file);

        zone.style.display = 'none';
        results.style.display = 'none';
        loading.style.display = 'block';

        try {
            const response = await fetch(endpoint, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error(`Server returned ${response.status}`);
            }

            // For video, it returns a file response
            if (endpoint.includes('video')) {
                const blob = await response.blob();
                successCallback(blob);
            } else {
                const data = await response.json();
                if (data.success) {
                    successCallback(data);
                } else {
                    alert(data.message || 'Error processing image');
                    zone.style.display = 'block';
                }
            }
        } catch (error) {
            console.error('Upload error:', error);
            alert('An error occurred during processing.');
            zone.style.display = 'block';
        } finally {
            loading.style.display = 'none';
        }
    }

    function handleImageSuccess(data) {
        const results = document.getElementById('image-results');
        const img = document.getElementById('result-image');
        const pCount = document.getElementById('person-count');
        const vCount = document.getElementById('vehicle-count');

        pCount.innerText = data.person_count;
        vCount.innerText = data.vehicle_count;
        img.src = `data:image/jpeg;base64,${data.annotated_image_base64}`;
        
        results.style.display = 'block';
    }

    function handleVideoSuccess(blob) {
        const results = document.getElementById('video-results');
        const video = document.getElementById('result-video');
        const downloadBtn = document.getElementById('download-video');

        const videoUrl = URL.createObjectURL(blob);
        video.src = videoUrl;
        downloadBtn.href = videoUrl;
        
        results.style.display = 'block';
    }
});
