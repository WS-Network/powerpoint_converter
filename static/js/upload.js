class FileUploader {
    constructor(chunkSize = 2 * 1024 * 1024) { // 2MB chunks for better memory usage
        this.chunkSize = chunkSize;
        this.currentChunk = 0;
        this.file = null;
        this.totalChunks = 0;
        this.maxRetries = 3;
        this.retryDelay = 1000; // 1 second
    }

    async uploadFile(file, onProgress) {
        this.file = file;
        this.currentChunk = 0;
        this.totalChunks = Math.ceil(file.size / this.chunkSize);

        try {
            for (let i = 0; i < this.totalChunks; i++) {
                let retries = 0;
                let success = false;

                while (!success && retries < this.maxRetries) {
                    try {
                        const chunk = this.getChunk(i);
                        await this.uploadChunk(chunk, i);
                        success = true;

                        if (onProgress) {
                            const progress = ((i + 1) / this.totalChunks) * 100;
                            onProgress(progress);
                        }

                        // Small delay between chunks to prevent memory buildup
                        await new Promise(resolve => setTimeout(resolve, 100));
                    } catch (error) {
                        retries++;
                        if (retries >= this.maxRetries) {
                            throw new Error(`Failed to upload chunk ${i} after ${this.maxRetries} attempts`);
                        }
                        console.warn(`Chunk ${i} upload failed, retrying... (${retries}/${this.maxRetries})`);
                        await new Promise(resolve => setTimeout(resolve, this.retryDelay));
                    }
                }
            }
            return true;
        } catch (error) {
            console.error('Upload failed:', error);
            throw error;
        }
    }

    getChunk(chunkNumber) {
        const start = chunkNumber * this.chunkSize;
        const end = Math.min(start + this.chunkSize, this.file.size);
        return this.file.slice(start, end);
    }

    async uploadChunk(chunk, chunkNumber) {
        const formData = new FormData();
        formData.append('file', chunk);
        formData.append('chunk', chunkNumber);
        formData.append('total', this.totalChunks);
        formData.append('filename', this.file.name);

        const response = await fetch('/upload-chunk', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        
        // Clear the chunk from memory
        formData.delete('file');
        return result;
    }
}

// Function to start the conversion process after upload
async function startConversion(filename, outputName, slideNumbers, direction) {
    const formData = new FormData();
    formData.append('filename', filename);
    formData.append('outputName', outputName);
    formData.append('slideNumbers', slideNumbers);
    formData.append('conversionDirection', direction);

    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        return result;
    } catch (error) {
        console.error('Conversion failed:', error);
        throw error;
    }
}

// Example usage with error handling and cleanup
document.addEventListener('DOMContentLoaded', () => {
    const uploader = new FileUploader();
    const form = document.getElementById('uploadForm');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const submitButton = form.querySelector('button[type="submit"]');

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const file = document.getElementById('file').files[0];
        if (!file) {
            progressText.textContent = 'Please select a file';
            return;
        }

        // Validate file size
        const maxSize = 100 * 1024 * 1024; // 100MB
        if (file.size > maxSize) {
            progressText.textContent = 'File size exceeds 100MB limit';
            return;
        }

        const outputName = document.getElementById('outputName').value;
        const slideNumbers = document.getElementById('slideNumbers').value;
        const direction = document.getElementById('conversionDirection').value;

        try {
            // Disable form during processing
            submitButton.disabled = true;
            progressBar.style.display = 'block';
            progressText.textContent = 'Uploading...';

            // Upload file in chunks
            await uploader.uploadFile(file, (progress) => {
                progressBar.value = progress;
                progressText.textContent = `Uploading: ${Math.round(progress)}%`;
            });

            // Start conversion
            progressText.textContent = 'Converting...';
            const result = await startConversion(file.name, outputName, slideNumbers, direction);

            if (result.download_url) {
                progressText.textContent = 'Download starting...';
                window.location.href = result.download_url;
            }
        } catch (error) {
            progressText.textContent = `Error: ${error.message}`;
            console.error('Error:', error);
        } finally {
            // Re-enable form
            submitButton.disabled = false;
        }
    });
}); 