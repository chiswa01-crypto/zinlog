/* ==========================================================================
   SCORPLOG LOGISTICS - JavaScript Interactive Engine & DOM Controller
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
    console.log('[SCORPLOG System] Interactive Engine Initialized.');
});

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

function simulateLogout() {
    const landingView = document.getElementById('landingView');
    const appView = document.getElementById('appView');
    
    if (landingView && appView) {
        appView.style.display = 'none';
        landingView.style.display = 'flex';
        console.log('[SCORPLOG System] Logged out successfully.');
    }
}

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

function backToDashboard() {
    const dashboardView = document.getElementById('dashboardView');
    const reviewView = document.getElementById('reviewView');

    if (dashboardView && reviewView) {
        reviewView.style.display = 'none';
        dashboardView.style.display = 'block';
        console.log('[SCORPLOG System] Returned to Dashboard Review Manual.');
    }
}

function saveAndApprove(event) {
    if (event) event.preventDefault();
    alert('Koreksi data berhasil disimpan! Status dokumen telah diperbarui menjadi APPROVED.');
    backToDashboard();
}
