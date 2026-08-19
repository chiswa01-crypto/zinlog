/* ==========================================================================
   SCORPLOG LOGISTICS - JavaScript Interactive Engine & DOM Controller
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
    console.log('[SCORPLOG System] Interactive Engine Initialized.');
});

/**
 * Simulasi Login Portal Admin (#landingView -> #appView)
 */
function simulateLogin(event) {
    if (event) event.preventDefault();
    
    const landingView = document.getElementById('landingView');
    const appView = document.getElementById('appView');
    
    if (landingView && appView) {
        landingView.style.display = 'none';
        appView.style.display = 'flex';
        console.log('[SCORPLOG System] Login successful. Navigated to Workspace.');
    }
}

/**
 * Simulasi Logout Portal Admin (#appView -> #landingView)
 */
function simulateLogout() {
    const landingView = document.getElementById('landingView');
    const appView = document.getElementById('appView');
    
    if (landingView && appView) {
        appView.style.display = 'none';
        landingView.style.display = 'flex';
        console.log('[SCORPLOG System] Logged out successfully.');
    }
}

/**
 * Trigger Simulasi Extraction PDF via Gemini AI Overlay (#processingOverlay)
 */
function triggerExtraction() {
    const overlay = document.getElementById('processingOverlay');
    if (!overlay) return;

    overlay.style.display = 'flex';
    console.log('[SCORPLOG System] AI PDF Extraction Triggered...');

    setTimeout(() => {
        overlay.style.display = 'none';
        addMockQuarantineRow();
        alert('Simulasi AI Extraction Selesai!\nDokumen baru ditambahkan ke Dashboard Review Manual (Status: Karantina).');
    }, 2000);
}

/**
 * Menambahkan Baris Mock Data Baru ke Tabel Review Manual
 */
function addMockQuarantineRow() {
    const tableBody = document.getElementById('tableBody');
    if (!tableBody) return;

    const newRow = document.createElement('tr');
    newRow.innerHTML = `
        <td class="file-name-cell">
            <i class="bi bi-file-earmark-pdf-fill text-danger fs-5"></i>
            <span>NPE_SIMULASI_BARU_${Math.floor(Math.random() * 100)}.pdf</span>
        </td>
        <td>Just now</td>
        <td><span class="badge-status-quarantine"><i class="bi bi-exclamation-triangle-fill"></i> Karantina</span></td>
        <td class="error-text-highlight">Logika Netto: Value exceeds gross weight limit</td>
        <td>
            <button class="btn-review-fix" onclick="openReviewView('NPE_SIMULASI_BARU')">Review & Fix</button>
        </td>
    `;

    tableBody.prepend(newRow);
}

/**
 * Membuka Workspace Koreksi Data Split Screen (#dashboardView -> #reviewView)
 */
function openReviewView(fileName) {
    const dashboardView = document.getElementById('dashboardView');
    const reviewView = document.getElementById('reviewView');
    const currentFileNameElem = document.getElementById('currentFileName');

    if (dashboardView && reviewView) {
        dashboardView.style.display = 'none';
        reviewView.style.display = 'block';
        if (currentFileNameElem && fileName) {
            currentFileNameElem.textContent = fileName;
        }
        console.log(`[SCORPLOG System] Opened Review Workspace for ${fileName}`);
    }
}

/**
 * Kembali dari Workspace Koreksi Data ke Dashboard (#reviewView -> #dashboardView)
 */
function backToDashboard() {
    const dashboardView = document.getElementById('dashboardView');
    const reviewView = document.getElementById('reviewView');

    if (dashboardView && reviewView) {
        reviewView.style.display = 'none';
        dashboardView.style.display = 'block';
        console.log('[SCORPLOG System] Returned to Dashboard Review Manual.');
    }
}

/**
 * Menyimpan Hasil Koreksi & Mengubah Status Menjadi Approved
 */
function saveAndApprove(event) {
    if (event) event.preventDefault();
    alert('Koreksi data berhasil disimpan! Status dokumen telah diperbarui menjadi APPROVED.');
    backToDashboard();
}
