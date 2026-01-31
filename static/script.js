// DOM Elements
const uploadBox = document.getElementById('uploadBox');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const fileInfo = document.getElementById('fileInfo');
const uploadSection = document.getElementById('uploadSection');
const progressSection = document.getElementById('progressSection');
const successSection = document.getElementById('successSection');
const errorSection = document.getElementById('errorSection');
const progressText = document.getElementById('progressText');
const progressFill = document.getElementById('progressFill');
const successMessage = document.getElementById('successMessage');
const errorMessage = document.getElementById('errorMessage');
const downloadBtn = document.getElementById('downloadBtn');
const newReportBtn = document.getElementById('newReportBtn');
const retryBtn = document.getElementById('retryBtn');

let generatedFilename = null;

// Browse button click
browseBtn.addEventListener('click', () => {
    fileInput.click();
});

// File input change
fileInput.addEventListener('change', (e) => {
    handleFileSelect(e.target.files[0]);
});

// Drag and drop handlers
uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.classList.add('drag-over');
});

uploadBox.addEventListener('dragleave', () => {
    uploadBox.classList.remove('drag-over');
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    handleFileSelect(file);
});

// Handle file selection
function handleFileSelect(file) {
    if (!file) return;

    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Invalid file type. Please upload an Excel file (.xlsx or .xls)');
        return;
    }

    // Validate file size (10MB)
    if (file.size > 10 * 1024 * 1024) {
        showError('File too large. Maximum size is 10MB');
        return;
    }

    // Show file info
    fileInfo.innerHTML = `
        <strong>📄 Selected File:</strong><br>
        ${file.name} (${(file.size / 1024).toFixed(2)} KB)
    `;
    fileInfo.classList.add('show');

    // Upload file
    uploadFile(file);
}

// Upload file to server
async function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    // Show progress
    showSection('progress');
    setProgress(10, 'Uploading file...');

    try {
        setProgress(30, 'Validating Excel format...');

        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        setProgress(60, 'Generating PowerPoint report...');

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || 'Upload failed');
        }

        setProgress(100, 'Finalizing report...');

        // Success
        setTimeout(() => {
            generatedFilename = data.filename;
            successMessage.textContent = data.message || 'Your EDR report is ready to download';
            showSection('success');
        }, 500);

    } catch (error) {
        showError(error.message);
    }
}

// Download file
downloadBtn.addEventListener('click', async () => {
    if (!generatedFilename) return;

    try {
        window.location.href = `/download/${generatedFilename}`;

        // Cleanup after 5 seconds
        setTimeout(async () => {
            await fetch(`/cleanup/${generatedFilename}`, { method: 'POST' });
        }, 5000);

    } catch (error) {
        showError('Error downloading file: ' + error.message);
    }
});

// New report button
newReportBtn.addEventListener('click', () => {
    reset();
});

// Retry button
retryBtn.addEventListener('click', () => {
    reset();
});

// Helper functions
function showSection(section) {
    uploadSection.classList.add('hidden');
    progressSection.classList.add('hidden');
    successSection.classList.add('hidden');
    errorSection.classList.add('hidden');

    switch (section) {
        case 'upload':
            uploadSection.classList.remove('hidden');
            break;
        case 'progress':
            progressSection.classList.remove('hidden');
            break;
        case 'success':
            successSection.classList.remove('hidden');
            break;
        case 'error':
            errorSection.classList.remove('hidden');
            break;
    }
}

function setProgress(percent, text) {
    progressFill.style.width = percent + '%';
    progressText.textContent = text;
}

function showError(message) {
    errorMessage.textContent = message;
    showSection('error');
}

function reset() {
    fileInput.value = '';
    fileInfo.innerHTML = '';
    fileInfo.classList.remove('show');
    generatedFilename = null;
    showSection('upload');
    setProgress(0, '');
}

// Initialize
showSection('upload');

// Avatar Eye Tracking Logic
document.addEventListener('mousemove', (e) => {
    const eyes = document.querySelectorAll('.eye');

    eyes.forEach(eye => {
        const pupil = eye.querySelector('.pupil');
        const eyeRect = eye.getBoundingClientRect();

        // Calculate eye center
        const eyeCenterX = eyeRect.left + eyeRect.width / 2;
        const eyeCenterY = eyeRect.top + eyeRect.height / 2;

        // Calculate angle between mouse and eye center
        const angle = Math.atan2(e.clientY - eyeCenterY, e.clientX - eyeCenterX);

        // Distance to move pupil (max 5px)
        const distance = Math.min(
            5,
            Math.hypot(e.clientX - eyeCenterX, e.clientY - eyeCenterY) / 10
        );

        // Calculate new pupil position
        const pupilX = Math.cos(angle) * distance;
        const pupilY = Math.sin(angle) * distance;

        // Apply transform centered
        pupil.style.transform = `translate(calc(-50% + ${pupilX}px), calc(-50% + ${pupilY}px))`;
    });
});
