// Export Tools Hub - Main JavaScript Script

document.addEventListener('DOMContentLoaded', function () {
    console.log('Export Tools Hub ready.');

    // 1. Auto update file input label when user selects single or multiple files
    const fileInputs = document.querySelectorAll('input[type="file"]');
    fileInputs.forEach(input => {
        input.addEventListener('change', function (e) {
            const files = e.target.files;
            let textLabel = 'Pilih file...';

            if (files && files.length > 1) {
                textLabel = `${files.length} file dipilih: ${files[0].name}, ${files[1].name}...`;
            } else if (files && files.length === 1) {
                textLabel = files[0].name;
            }

            const card = input.closest('.card-body');
            if (card) {
                const labelElement = card.querySelector('.file-label-text');
                if (labelElement) {
                    labelElement.textContent = textLabel;
                }
            }
        });
    });

    // 2. Client-side Multiple Upload Handler via Fetch API (FormData)
    const ciplForm = document.getElementById('ciplUploadForm');
    if (ciplForm) {
        ciplForm.addEventListener('submit', function (e) {
            const fileInput = document.getElementById('ciplFile');
            if (fileInput && fileInput.files.length > 0) {
                console.log(`Mengirim ${fileInput.files.length} file ke server CIPL...`);
            }
        });
    }
});

/**
 * Contoh Fungsi helper Fetch API (AJAX) untuk Upload Multiple Files
 */
async function uploadMultipleFiles(url, files) {
    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
        formData.append('files', files[i]);
    }

    try {
        const response = await fetch(url, {
            method: 'POST',
            body: formData
        });
        const result = await response.json();
        return result;
    } catch (error) {
        console.error('Multiple upload error:', error);
        throw error;
    }
}
