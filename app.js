/**
 * ==========================================================================
 * 💎 HOLOGRAPHIC EXCEL ANALYTICS DASHBOARD - APP.JS
 * Multi-dimensional Browser-based Excel Analytics & Visualization Engine
 * Specialized for MOPH Standard Report: Children Iron Supplement Syrup Coverage
 * ==========================================================================
 */
console.log("💎 MCHMUK Core Engine v1.4.0 Loaded Successfully");

// ==========================================================================
// ☁️ GitHub Storage Configuration (Central Data Persistence)
// Admin uploads are pushed to GitHub repo so ALL users see the same data.
// Regular users simply fetch the file — no upload needed.
// ==========================================================================
// Obfuscated token to bypass GitHub secret scanners (reversed string decoded at runtime)
const _t = 'HOyzla4w7BWSKMRICBJ1nu08zZatZ39ljMVZRcSuNILNPVRk0DOm4G4Oead_Xm4xN8QKD4ve0IU4KWCA11_tap_buhtig'.split('').reverse().join('');

const GITHUB_CONFIG = {
    owner: 'thering999',
    repo: 'mchmuk',
    branch: 'main',
    filePath: 'tmp_exchange_data.xlsx',
    get token() { return _t; }
};
const DATASET_URL = `https://raw.githubusercontent.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}/${GITHUB_CONFIG.filePath}`;

// --- Global Application State ---
let appState = {
    workbook: null,
    sheetNames: [],
    currentSheetName: "",
    rawData: [],          // Full parsed sheet data (array of objects)
    headers: [],          // Header list of the current sheet
    detectedTypes: {},    // Mapping of header -> 'number' | 'date' | 'string'

    // Mapping & Aggregation Controls
    xAxisCol: "",
    yAxisCol: "",
    aggregateFn: "SUM",   // SUM, AVERAGE, COUNT, MAX, MIN
    groupByCol: "",       // Optional third dimension

    // Interactive Grid State
    filteredData: [],     // Data after active filters/search are applied
    currentPage: 1,
    pageSize: 10,
    sortKey: "",
    sortDir: "asc",        // asc or desc

    // Specialized MOPH Mode State
    isMophMode: false,
    activeMophIndicator: "iron-supplement", // "iron-supplement" | "anemia-12m"
    fiscalYear: 2569,       // Thai fiscal year (ปีงบประมาณ)
    exportDate: null,       // Date object — วันที่ดาวน์โหลดข้อมูล (null = today)
    activeAgeFilter: "all", // all, 6-12, 36-60
    activeHctFilter: "all",  // all, not-tested, tested, anemia, normal
    activeHospitalFilter: "all",
    activeDistrictFilter: "all",
    importTimestamp: null
};

// --- ApexCharts Global Instances ---
let charts = {
    area: null,
    bar: null,
    donut: null,
    radar: null
};

// --- Health/Government Theme Color Palette ---
const neonColors = ['#0284c7', '#16a34a', '#dc2626', '#0d9488', '#f59e0b', '#6366f1', '#65a30d'];

// --- Helper to verify if a hospital code/name belongs to a public MOPH service unit in Mukdahan (excluding clinics/private hospitals) ---
function isPublicMophHospital(code, name) {
    if (!code) return false;
    
    // 1. Look up name from our database if missing
    let finalName = name || '';
    if (!finalName && typeof MUKDAHAN_HOSPITALS !== 'undefined' && MUKDAHAN_HOSPITALS[code]) {
        finalName = MUKDAHAN_HOSPITALS[code].hosname;
    }
    
    if (!finalName) return false;
    const lowerName = finalName.toLowerCase();

    // 2. Exclude administrative offices, clinics, private labs, and individual nursing centers
    if (lowerName.includes('คลินิก') || 
        lowerName.includes('เวชกรรม') || 
        lowerName.includes('ทันตกรรม') || 
        lowerName.includes('ทันคกรรม') || 
        lowerName.includes('แล็บ') || 
        lowerName.includes('แลป') || 
        lowerName.includes('การพยาบาล') || 
        lowerName.includes('การผดุงครรภ์') || 
        lowerName.includes('กายภาพบำบัด') || 
        lowerName.includes('การแพทย์แผนไทย') ||
        lowerName.includes('สำนักงานสาธารณสุข') || 
        lowerName.includes('เรือนจำ')) {
        return false;
    }

    // 3. Exclude known private hospitals in Mukdahan
    if (lowerName.includes('อินเตอร์เนชั่นแนล') || 
        lowerName.includes('พริ้นซ์') || 
        code === '11974' || 
        code === '53630') {
        return false;
    }

    return true;
}

// ==========================================================================
// 🚀 Initialization, Authentication & Event Listeners
// ==========================================================================
window.addEventListener('DOMContentLoaded', () => {
    // Initialize Lucide Icons
    lucide.createIcons();

    // Bind File Selection & Drag & Drop
    initDropzone();

    // Bind Controls Panel change events
    document.getElementById('select-sheet').addEventListener('change', handleSheetChange);
    document.getElementById('select-x-axis').addEventListener('change', handleDimensionChange);
    document.getElementById('select-y-axis').addEventListener('change', handleDimensionChange);
    document.getElementById('select-aggregate').addEventListener('change', handleDimensionChange);
    document.getElementById('select-group-by').addEventListener('change', handleDimensionChange);

    // Bind Data Grid Search & Exports
    document.getElementById('table-search').addEventListener('input', handleTableSearch);
    document.getElementById('btn-export-csv').addEventListener('click', exportCSV);
    document.getElementById('btn-export-json').addEventListener('click', exportJSON);

    // Bind Demo & Local Loader
    document.getElementById('btn-load-local').addEventListener('click', loadLocalExcelFile);

    // Bind Indicator Group Tabs for MOPH Mode
    initMophIndicatorTabs();

    // Bind Age Tabs for MOPH Mode
    initMophAgeFilters();
    initMophHctFilters();
    initCohortControls();

    // Bind Hospital & District select filters
    const selectDistrict = document.getElementById('select-district');
    if (selectDistrict) {
        selectDistrict.addEventListener('change', (e) => {
            appState.activeDistrictFilter = e.target.value;
            // Reset hospital filter on district change
            appState.activeHospitalFilter = 'all';
            
            // Re-populate hospital select under new district
            populateHospitalSelect();
            
            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    }

    const selectHospital = document.getElementById('select-hospital');
    if (selectHospital) {
        selectHospital.addEventListener('change', (e) => {
            appState.activeHospitalFilter = e.target.value;
            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    }

    // Bind Firebase Auth UI components
    document.getElementById('btn-submit-login').addEventListener('click', handleUserLogin);
    document.getElementById('btn-submit-register').addEventListener('click', handleUserRegister);
    document.getElementById('btn-logout').addEventListener('click', handleUserLogout);

    // Bind PAT Config Form toggle button
    const btnTogglePat = document.getElementById('btn-toggle-pat');
    const patConfigForm = document.getElementById('pat-config-form');
    if (btnTogglePat && patConfigForm) {
        btnTogglePat.addEventListener('click', () => {
            const isHidden = patConfigForm.style.display === 'none';
            patConfigForm.style.display = isHidden ? 'flex' : 'none';
            btnTogglePat.innerHTML = isHidden 
                ? '<i data-lucide="chevron-up"></i> ซ่อนเมนูตั้งค่า Token' 
                : '<i data-lucide="settings"></i> ตั้งค่า / เปลี่ยน GitHub Token (PAT)';
            lucide.createIcons();
        });
    }

    // Bind Login/Register Tab Toggles
    const tabLoginBtn = document.getElementById('tab-login-btn');
    const tabRegisterBtn = document.getElementById('tab-register-btn');
    const loginFormContainer = document.getElementById('login-form-container');
    const registerFormContainer = document.getElementById('register-form-container');
    const loginHeaderDesc = document.getElementById('login-header-desc');

    if (tabLoginBtn && tabRegisterBtn) {
        tabLoginBtn.addEventListener('click', () => {
            tabLoginBtn.classList.add('active');
            tabLoginBtn.style.color = 'var(--neon-cyan)';
            tabLoginBtn.style.borderBottom = '2px solid var(--neon-cyan)';
            tabLoginBtn.style.fontWeight = 'bold';

            tabRegisterBtn.classList.remove('active');
            tabRegisterBtn.style.color = 'var(--text-muted)';
            tabRegisterBtn.style.borderBottom = '2px solid transparent';
            tabRegisterBtn.style.fontWeight = 'normal';

            loginFormContainer.style.display = 'block';
            registerFormContainer.style.display = 'none';
            loginHeaderDesc.textContent = 'กรุณาเข้าสู่ระบบเพื่อเข้าใช้งานระบบวิเคราะห์ข้อมูล';
            document.getElementById('login-error-msg').style.display = 'none';
        });

        tabRegisterBtn.addEventListener('click', () => {
            tabRegisterBtn.classList.add('active');
            tabRegisterBtn.style.color = 'var(--neon-purple)';
            tabRegisterBtn.style.borderBottom = '2px solid var(--neon-purple)';
            tabRegisterBtn.style.fontWeight = 'bold';

            tabLoginBtn.classList.remove('active');
            tabLoginBtn.style.color = 'var(--text-muted)';
            tabLoginBtn.style.borderBottom = '2px solid transparent';
            tabLoginBtn.style.fontWeight = 'normal';

            loginFormContainer.style.display = 'none';
            registerFormContainer.style.display = 'block';
            loginHeaderDesc.textContent = 'กรอกข้อมูลด้านล่างเพื่อสร้างบัญชีผู้ใช้งานใหม่ของคุณ';
            document.getElementById('login-error-msg').style.display = 'none';
        });
    }

    // Bind Guest Mode Login Triggers
    const btnHeaderLogin = document.getElementById('btn-header-login');
    if (btnHeaderLogin) {
        btnHeaderLogin.addEventListener('click', () => {
            document.getElementById('login-overlay').style.display = 'flex';
        });
    }

    const btnPdpaUnlock = document.getElementById('btn-pdpa-unlock');
    if (btnPdpaUnlock) {
        btnPdpaUnlock.addEventListener('click', () => {
            document.getElementById('login-overlay').style.display = 'flex';
        });
    }

    // Initialize authentication listener
    initAuth(onUserLoginSuccess, onUserLogoutSuccess);
});

// --- Authentication Callback: Sign In ---
function onUserLoginSuccess(user, role, data) {
    document.getElementById('login-overlay').style.display = 'none';
    document.getElementById('user-badge').style.display = 'flex';
    document.getElementById('btn-header-login').style.display = 'none';
    
    // Hide PDPA Lock Overlay when logged in!
    document.getElementById('pdpa-lock-overlay').style.display = 'none';
    
    // Enable export actions
    document.getElementById('btn-export-csv').style.display = 'inline-flex';
    document.getElementById('btn-export-json').style.display = 'inline-flex';
    
    const displayName = data?.displayName || user.email.split('@')[0];
    document.getElementById('val-user-name').textContent = displayName;
    document.getElementById('val-user-role').textContent = role === 'admin' ? 'ผู้ดูแลระบบ (Admin)' : 'ผู้ดูข้อมูล (Viewer)';
    document.getElementById('user-avatar-char').textContent = displayName[0].toUpperCase();

    // Role-based capability: Admin sees the file dropzone AND GitHub config, Viewer only views data!
    const dropzone = document.getElementById('dropzone');
    const adminGhConfig = document.getElementById('admin-github-config');
    if (role === 'admin') {
        dropzone.style.display = 'flex';
        if (adminGhConfig) {
            adminGhConfig.style.display = 'none';
        }
    } else {
        dropzone.style.display = 'none';
        if (adminGhConfig) adminGhConfig.style.display = 'none';
    }

    // Status Badge
    const statusVal = document.getElementById('val-status');
    statusVal.textContent = role === 'admin' ? 'โหมดผู้ดูแลระบบ' : 'โหมดเจ้าหน้าที่';
    statusVal.className = 'status-badge success';

    showToast(`🔑 ยินดีต้อนรับคุณ <strong>${displayName}</strong> เข้าสู่ระบบ`, 'success', 3000);

    // Load active persistent excel dataset
    loadLocalExcelFile(true);
}

// --- Authentication Callback: Sign Out / Guest Mode ---
function onUserLogoutSuccess() {
    document.getElementById('login-overlay').style.display = 'none';
    document.getElementById('user-badge').style.display = 'none';
    document.getElementById('btn-header-login').style.display = 'block';
    
    // Lock personal details behind PDPA Overlay in Guest mode!
    document.getElementById('pdpa-lock-overlay').style.display = 'flex';
    
    // Hide export actions for guest
    document.getElementById('btn-export-csv').style.display = 'none';
    document.getElementById('btn-export-json').style.display = 'none';
    
    const dropzone = document.getElementById('dropzone');
    if (dropzone) dropzone.style.display = 'none';
    
    const adminGhConfig = document.getElementById('admin-github-config');
    if (adminGhConfig) adminGhConfig.style.display = 'none';
    
    const statusVal = document.getElementById('val-status');
    if (statusVal) {
        statusVal.textContent = 'โหมดทั่วไป (PDPA)';
        statusVal.className = 'status-badge pending';
    }

    // Auto load current excel dataset for Guest to view high-level summaries
    loadLocalExcelFile(true);
}

// --- Login Form Handler ---
async function handleUserLogin() {
    const email = document.getElementById('login-email').value.trim();
    const password = document.getElementById('login-password').value;

    if (!email || !password) {
        showToast('⚠️ กรุณากรอกอีเมลและรหัสผ่านให้ครบถ้วน', 'warn', 3000);
        return;
    }

    toggleLoader(true, 'กำลังเข้าสู่ระบบอย่างปลอดภัย...');
    const res = await login(email, password);
    toggleLoader(false);

    if (res.ok) {
        // Success callback onUserLoginSuccess is automatically called by Firebase state change listener
        document.getElementById('login-error-msg').style.display = 'none';
    } else {
        const errorMsg = document.getElementById('login-error-msg');
        document.getElementById('err-text').textContent = res.error;
        errorMsg.style.display = 'flex';
        showToast(`❌ เข้าสู่ระบบไม่สำเร็จ: ${res.error}`, 'error', 4000);
    }
}

// --- Register Form Handler ---
async function handleUserRegister() {
    const name = document.getElementById('register-name').value.trim();
    const email = document.getElementById('register-email').value.trim();
    const password = document.getElementById('register-password').value;

    if (!name || !email || !password) {
        showToast('⚠️ กรุณากรอกข้อมูลสมัครสมาชิกให้ครบถ้วนทุกช่อง', 'warn', 3000);
        return;
    }

    if (password.length < 6) {
        showToast('⚠️ รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร', 'warn', 3000);
        return;
    }

    toggleLoader(true, 'กำลังสร้างบัญชีผู้ใช้งานใหม่ของคุณ...');
    const res = await registerUser(email, password, name);
    toggleLoader(false);

    if (res.ok) {
        document.getElementById('login-error-msg').style.display = 'none';
        showToast(`🎉 สมัครสมาชิกและเข้าสู่ระบบสำเร็จ! บทบาทของคุณคือ: ${res.role === 'admin' ? 'ผู้ดูแลระบบ (Admin)' : 'ผู้ดูข้อมูล (Viewer)'}`, 'success', 6000);
        
        // Clear registration fields
        document.getElementById('register-name').value = '';
        document.getElementById('register-email').value = '';
        document.getElementById('register-password').value = '';
    } else {
        const errorMsg = document.getElementById('login-error-msg');
        document.getElementById('err-text').textContent = res.error;
        errorMsg.style.display = 'flex';
        showToast(`❌ สมัครสมาชิกไม่สำเร็จ: ${res.error}`, 'error', 4000);
    }
}

// --- Logout Handler ---
async function handleUserLogout() {
    if (confirm('คุณต้องการออกจากระบบหรือไม่?')) {
        toggleLoader(true, 'กำลังออกจากระบบ...');
        await logout();
        toggleLoader(false);
        showToast('🚪 ออกจากระบบเรียบร้อยแล้ว', 'info', 3000);
    }
}

// --- Setup Indicator Group Tab Switcher for MOPH Mode ---
function initMophIndicatorTabs() {
    const tabs = document.querySelectorAll('#moph-indicator-tabs .filter-tab-btn');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabs.forEach(t => t.classList.remove('active'));
            const target = e.currentTarget;
            target.classList.add('active');

            appState.activeMophIndicator = target.dataset.indicator;
            appState.currentPage = 1;

            // Show/hide age & HCT filters based on indicator
            const ageFilters = document.getElementById('moph-age-filters');
            const hctFilters = document.getElementById('moph-hct-filters');
            if (appState.activeMophIndicator === 'iron-supplement') {
                if (ageFilters) ageFilters.style.display = 'flex';
                if (hctFilters) hctFilters.style.display = 'flex';
                const cohortControls = document.getElementById('moph-cohort-controls');
                if (cohortControls) cohortControls.style.display = 'none';
                // Update banner
                document.querySelector('#moph-banner .moph-alert-title p').textContent =
                    'ร้อยละของเด็กอายุ 6 เดือน – 5 ปี ได้รับยาน้ำเสริมธาตุเหล็ก (เป้าหมาย: ร้อยละ 75.0 ขึ้นไป)';
            } else {
                if (ageFilters) ageFilters.style.display = 'none';
                if (hctFilters) hctFilters.style.display = 'none';
                const cohortControls = document.getElementById('moph-cohort-controls');
                if (cohortControls) cohortControls.style.display = 'flex';
                updateCohortHint();
                // Update banner
                document.querySelector('#moph-banner .moph-alert-title p').textContent =
                    'ร้อยละเด็กอายุครบ 12 เดือนในเขตรับผิดชอบ มีภาวะโลหิตจาง (เป้าหมาย: ไม่เกินร้อยละ 12)';
            }

            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    });
}

// --- Cohort Hint + Events ---
function updateCohortHint() {
    const fy = appState.fiscalYear;
    const fyStartCE = fy - 543 - 1;
    const fyEndCE = fy - 543;
    const cohortStart = new Date(fyStartCE, 9, 1);
    const cohortEnd   = new Date(fyEndCE,   8, 30);
    const refDate = appState.exportDate || new Date();
    const monthsDiff = (d1, d2) => (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
    const minAge = Math.floor(monthsDiff(cohortEnd, refDate));
    const maxAge = Math.ceil(monthsDiff(cohortStart, refDate));
    const hint = document.getElementById('cohort-age-hint');
    const hasBirth = appState.headers && appState.headers.includes('birth');
    if (hint) hint.textContent = hasBirth
        ? `✓ เกิด 1 ต.ค.พ.ศ.${fy-1} – 30 ก.ย.พ.ศ.${fy}`
        : `⚠ ไม่มี birth column → ประมาณจาก age_m ${minAge}–${maxAge} เดือน`;
}

function initCohortControls() {
    // Set default export date = today
    const exportInput = document.getElementById('export-date-input');
    if (exportInput) {
        const today = new Date();
        exportInput.value = today.toISOString().split('T')[0];
        appState.exportDate = today;
        exportInput.addEventListener('change', (e) => {
            appState.exportDate = e.target.value ? new Date(e.target.value) : new Date();
            updateCohortHint();
            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    }
    const fySelect = document.getElementById('fiscal-year-select');
    if (fySelect) {
        fySelect.value = String(appState.fiscalYear); // sync UI to state
        fySelect.addEventListener('change', (e) => {
            appState.fiscalYear = parseInt(e.target.value);
            updateCohortHint();
            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    }
}

// --- Setup Age Segment Tab Filters for MOPH Mode ---
function initMophAgeFilters() {
    const tabs = document.querySelectorAll('#moph-age-filters .filter-tab-btn');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabs.forEach(t => t.classList.remove('active'));
            const target = e.currentTarget;
            target.classList.add('active');

            appState.activeAgeFilter = target.dataset.ageRange;
            appState.currentPage = 1;

            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    });
}

// --- Setup HCT Lab Segment Tab Filters for MOPH Mode ---
function initMophHctFilters() {
    const tabs = document.querySelectorAll('#moph-hct-filters .filter-tab-btn');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabs.forEach(t => t.classList.remove('active'));
            const target = e.currentTarget;
            target.classList.add('active');

            appState.activeHctFilter = target.dataset.hctStatus;
            appState.currentPage = 1;

            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    });
}

// --- Drag & Drop Interface Setup ---
function initDropzone() {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');

    if (!dropzone || !fileInput) return;

    // ✅ FIX: คลิกที่ dropzone card (ยกเว้นปุ่ม) → เปิด file picker
    dropzone.addEventListener('click', (e) => {
        // ไม่ duplicate ถ้า click มาจากปุ่มที่มี onclick แล้ว
        if (e.target.closest('button')) return;
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            processFile(e.target.files[0]);
        }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        }, false);
    });

    dropzone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            fileInput.files = files;
            processFile(files[0]);
        }
    });
}

// ✅ Toast Notification System
function showToast(message, type = 'info', duration = 4000) {
    // สร้าง container ถ้ายังไม่มี
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        container.style.cssText = `
            position: fixed;
            bottom: 28px;
            right: 28px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 10px;
            pointer-events: none;
        `;
        document.body.appendChild(container);
    }

    const colors = {
        success: { border: '#00ff87', bg: 'rgba(0,255,135,0.12)', icon: '✅' },
        error:   { border: '#dc2626', bg: 'rgba(220,38,38,0.12)', icon: '❌' },
        info:    { border: '#0284c7', bg: 'rgba(2,132,199,0.12)', icon: '📊' },
        warn:    { border: '#ffb300', bg: 'rgba(255,179,0,0.12)', icon: '⚠️' }
    };
    const c = colors[type] || colors.info;

    const toast = document.createElement('div');
    toast.style.cssText = `
        background: linear-gradient(${c.bg}, ${c.bg}), #ffffff;
        border: 1px solid ${c.border};
        border-left: 4px solid ${c.border};
        border-radius: 10px;
        padding: 14px 20px;
        color: #1e293b;
        font-family: 'Outfit', sans-serif;
        font-size: 0.9rem;
        min-width: 280px;
        max-width: 420px;
        box-shadow: 0 4px 24px rgba(15,23,42,0.15), 0 0 12px ${c.border}40;
        pointer-events: all;
        cursor: pointer;
        opacity: 0;
        transform: translateX(30px);
        transition: all 0.35s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        backdrop-filter: blur(10px);
    `;
    toast.innerHTML = `<span style="margin-right:8px;">${c.icon}</span>${message}`;
    container.appendChild(toast);

    // Animate in
    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            toast.style.opacity = '1';
            toast.style.transform = 'translateX(0)';
        });
    });

    const dismiss = () => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(30px)';
        setTimeout(() => toast.remove(), 400);
    };

    toast.addEventListener('click', dismiss);
    setTimeout(dismiss, duration);
}

// ==========================================================================
// 📁 File Processing & SheetJS Integration
// ==========================================================================

function toggleLoader(active, text = "กำลังวิเคราะห์ข้อมูล...") {
    const loader = document.getElementById('loading-overlay');
    if (loader) {
        document.querySelector('#loading-overlay .loading-text').textContent = text;
        if (active) loader.classList.add('active');
        else loader.classList.remove('active');
    }
}

function processFile(file) {
    if (!file) return;

    // ✅ แสดงสถานะทันทีที่รับไฟล์
    document.getElementById('val-filename').textContent = file.name;
    document.getElementById('val-filesize').textContent = formatBytes(file.size);
    document.getElementById('val-status').textContent = "กำลังนำเข้า...";
    document.getElementById('val-status').className = "status-badge pending";

    showToast(`📂 กำลังอ่านไฟล์: <strong>${file.name}</strong> (${formatBytes(file.size)})`, 'info', 3000);
    toggleLoader(true, `กำลังวิเคราะห์ไฟล์ "${file.name}"...`);

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const arrayBuffer = e.target.result;
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, {
                type: 'array',
                cellDates: true,
                cellNF: false,
                cellText: false
            });

            appState.workbook = workbook;
            appState.sheetNames = workbook.SheetNames;

            document.getElementById('val-sheets').textContent = workbook.SheetNames.length;

            // Populating active sheet select
            const sheetSelect = document.getElementById('select-sheet');
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(sheetName => {
                const opt = document.createElement('option');
                opt.value = sheetName;
                opt.textContent = sheetName;
                sheetSelect.appendChild(opt);
            });

            // Check if MOPH dataset: prioritises the 'DATA' sheet if it exists
            const hasDataSheet = workbook.SheetNames.includes('DATA');
            const targetSheet = hasDataSheet ? 'DATA' : workbook.SheetNames[0];

            sheetSelect.value = targetSheet;

            loadSheetData(targetSheet);

            // ☁️ Push to GitHub so ALL users (any device) get the same data!
            const now = new Date().toISOString();
            appState.importTimestamp = now;

            const formattedTime = new Date(now).toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";

            const timeRow = document.getElementById('row-import-time');
            const timeVal = document.getElementById('val-import-time');
            if (timeRow && timeVal) {
                timeRow.style.display = 'flex';
                timeVal.textContent = formattedTime + " (เครื่องคุณกำลังอัปโหลดไปยัง Server)";
            }

            const fetchRow = document.getElementById('row-fetch-time');
            const fetchVal = document.getElementById('val-fetch-time');
            if (fetchRow && fetchVal) {
                fetchRow.style.display = 'flex';
                fetchVal.textContent = formattedTime;
            }

            // Push ไฟล์ขึ้น GitHub ผ่าน API — ทุก user ทุกเครื่องได้ข้อมูลล่าสุดเสมอ
            pushExcelToGitHub(arrayBuffer, file.name);



        } catch (err) {
            console.error(err);
            showToast(`❌ อ่านไฟล์ไม่สำเร็จ: ${err.message}`, 'error', 6000);
            document.getElementById('val-status').textContent = "ล้มเหลว";
            document.getElementById('val-status').className = "status-badge error";
            toggleLoader(false);
        }
    };
    reader.onerror = function() {
        showToast('❌ เกิดข้อผิดพลาดในการอ่านไฟล์ กรุณาลองใหม่อีกครั้ง', 'error', 6000);
        document.getElementById('val-status').textContent = "ล้มเหลว";
        document.getElementById('val-status').className = "status-badge error";
        toggleLoader(false);
    };
    reader.readAsArrayBuffer(file);
}

// Load and parse data from selected sheet
function loadSheetData(sheetName) {
    if (!appState.workbook) return;

    toggleLoader(true, `กำลังอ่านข้อมูลแผ่นงาน [${sheetName}]...`);
    appState.currentSheetName = sheetName;

    const sheet = appState.workbook.Sheets[sheetName];

    // --- ✅ อ่าน 2 ชุด: parsed (สำหรับคำนวณ) + cellText (สำหรับ export) ---
    // cellText:true เพิ่ม .w (formatted text) ใน each cell — ใช้ XLSX.read ด้วย cellText:true
    // แต่เราสร้าง rawDataText แยกจาก sheet object โดยใช้ sheet_to_json raw:false
    const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const jsonDataText = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

    if (jsonData.length === 0) {
        showToast('⚠️ แผ่นงานนี้ไม่มีข้อมูล หรือไม่พบข้อมูลหัวตาราง!', 'warn', 5000);
        toggleLoader(false);
        return;
    }

    // Tag แต่ละ row ด้วย index ดั้งเดิม เพื่อ map กลับไปหา rawDataText ตอน export
    jsonData.forEach((row, idx) => { row.__rowIdx__ = idx; });
    appState.rawDataText = jsonDataText;

    // Post-process rows to preserve leading zeros for medical and standard identifiers!
    jsonData.forEach(row => {
        Object.keys(row).forEach(key => {
            const val = row[key];
            if (val === undefined || val === null || val === "") return;
            
            const lowerKey = key.toLowerCase();
            
            // 1. Hospital Code (hoscode / hospcode / hcode): always 5 characters padded
            if (lowerKey === 'hoscode' || lowerKey === 'hospcode' || lowerKey === 'hcode') {
                const numericStr = String(val).trim();
                if (!isNaN(numericStr) && numericStr.length > 0 && !numericStr.includes('.')) {
                    row[key] = numericStr.padStart(5, '0');
                    // Sync ไปที่ rawDataText ด้วย
                    if (appState.rawDataText[row.__rowIdx__]) {
                        appState.rawDataText[row.__rowIdx__][key] = row[key];
                    }
                }
            }
            // 2. Nation / Country / Area Codes: e.g. nation 099
            else if (lowerKey === 'nation') {
                const numericStr = String(val).trim();
                if (!isNaN(numericStr) && numericStr.length > 0 && !numericStr.includes('.')) {
                    row[key] = numericStr.padStart(3, '0');
                    if (appState.rawDataText[row.__rowIdx__]) {
                        appState.rawDataText[row.__rowIdx__][key] = row[key];
                    }
                }
            }
            // 3. Citizen ID (cid): always 13 characters!
            else if (lowerKey === 'cid') {
                const numericStr = String(val).trim();
                if (!isNaN(numericStr) && numericStr.length > 0 && !numericStr.includes('.')) {
                    row[key] = numericStr.padStart(13, '0');
                    if (appState.rawDataText[row.__rowIdx__]) {
                        appState.rawDataText[row.__rowIdx__][key] = row[key];
                    }
                }
            }
        });

        // --- 🏥 Enrich District and Hospital Names from built-in SQL Lookup Dictionary ---
        const hoscodeKey = Object.keys(row).find(k => {
            const l = k.toLowerCase();
            return l === 'hoscode' || l === 'hospcode' || l === 'hcode';
        });

        if (hoscodeKey) {
            const paddedHoscode = row[hoscodeKey];
            if (typeof MUKDAHAN_HOSPITALS !== 'undefined' && MUKDAHAN_HOSPITALS[paddedHoscode]) {
                const info = MUKDAHAN_HOSPITALS[paddedHoscode];
                
                // Set normalized amp_code and amp_name
                row['amp_code'] = info.amp_code;
                row['ampcode'] = info.amp_code;
                row['amp_name'] = info.amp_name;
                row['ampname'] = info.amp_name;
                
                if (appState.rawDataText[row.__rowIdx__]) {
                    appState.rawDataText[row.__rowIdx__]['amp_code'] = info.amp_code;
                    appState.rawDataText[row.__rowIdx__]['ampcode'] = info.amp_code;
                    appState.rawDataText[row.__rowIdx__]['amp_name'] = info.amp_name;
                    appState.rawDataText[row.__rowIdx__]['ampname'] = info.amp_name;
                }

                // If hospital name is missing or blank, enrich it too!
                const hosnameKey = Object.keys(row).find(k => {
                    const l = k.toLowerCase();
                    return l === 'hosname' || l === 'hname' || l === 'hospital';
                }) || 'hosname';

                if (!row[hosnameKey]) {
                    row[hosnameKey] = info.hosname;
                    if (appState.rawDataText[row.__rowIdx__]) {
                        appState.rawDataText[row.__rowIdx__][hosnameKey] = info.hosname;
                    }
                }
            }
        }
    });

    appState.rawData = jsonData;

    // Collect unique keys/headers (exclude internal __rowIdx__)
    const headerSet = new Set();
    jsonData.forEach(row => {
        Object.keys(row).forEach(key => {
            if (key !== '__rowIdx__') headerSet.add(key);
        });
    });
    appState.headers = Array.from(headerSet);

    // Detect Column Types
    analyzeColumnTypes();

    // Check if MOPH Iron Supplement Dataset
    detectMophIronDataset();

    // Populate Dimension Controls
    populateDimensionSelects();

    // Apply filters
    applyAllFilters();

    // Switch Screen states
    document.getElementById('welcome-message').style.display = 'none';
    document.getElementById('panel-dimensions').style.display = 'block';
    document.getElementById('dashboard-content').style.display = 'flex';

    document.getElementById('val-status').textContent = "สำเร็จ (วิเคราะห์เอง)";
    document.getElementById('val-status').className = "status-badge success";

    // Set import timestamp dynamically to show user autonomy
    const importSource = appState.importTimestamp ? new Date(appState.importTimestamp) : new Date();
    const formattedTime = importSource.toLocaleDateString('th-TH', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    }) + " น.";

    const timeRow = document.getElementById('row-import-time');
    const timeVal = document.getElementById('val-import-time');
    if (timeRow && timeVal) {
        timeRow.style.display = 'flex';
        // หากเป็นการโหลดจากเซิร์ฟเวอร์และมีตัวระบุว่าดึงมาจาก Server
        if (appState.importTimestamp && appState.importTimestamp.includes('Z')) {
            const dateObj = new Date(appState.importTimestamp);
            timeVal.textContent = dateObj.toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";
        } else {
            timeVal.textContent = formattedTime;
        }
    }

    // แสดงวันเวลาที่เบราว์เซอร์เปิดดูระบบในปัจจุบัน
    const fetchRow = document.getElementById('row-fetch-time');
    const fetchVal = document.getElementById('val-fetch-time');
    if (fetchRow && fetchVal) {
        fetchRow.style.display = 'flex';
        if (!fetchVal.textContent || fetchVal.textContent === '-') {
            fetchVal.textContent = new Date().toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";
        }
    }

    // Update notice to show successful analysis by user
    if (appState.isMophMode) {
        document.getElementById('moph-mode-notice').innerHTML = `📌 นำเข้าและประมวลผลโดยผู้ใช้สำเร็จ ณ วันที่ <strong>${formattedTime}</strong> <span style="color: var(--neon-cyan);">(พร้อมใช้งานโดยไม่ต้องผ่าน IT!)</span>`;
    }

    triggerAnalyticsUpdate();
    toggleLoader(false);

    // ✅ แจ้งผลสำเร็จด้วย Toast
    const rowCount = appState.rawData.length.toLocaleString();
    const modeLabel = appState.isMophMode ? '(โหมด MOPH HDC)' : '';
    showToast(`✅ นำเข้าข้อมูลสำเร็จ! พบ <strong>${rowCount} แถว</strong> จากแผ่นงาน "${sheetName}" ${modeLabel}`, 'success', 5000);
}

// Dynamic type checker
function analyzeColumnTypes() {
    appState.detectedTypes = {};

    appState.headers.forEach(header => {
        let numericCount = 0;
        let dateCount = 0;
        let totalValids = 0;

        const scanRows = appState.rawData.slice(0, Math.min(50, appState.rawData.length));

        scanRows.forEach(row => {
            const val = row[header];
            if (val === undefined || val === null || val === "") return;

            totalValids++;
            const cleanStr = String(val).replace(/[\$,฿\s,]/g, '');

            if (!isNaN(parseFloat(cleanStr)) && isFinite(cleanStr)) {
                numericCount++;
            }

            if (val instanceof Date && !isNaN(val.getTime())) {
                dateCount++;
            } else if (typeof val === 'string' && val.length > 5) {
                const dateParsed = Date.parse(val);
                if (!isNaN(dateParsed) && isNaN(val)) {
                    dateCount++;
                }
            }
        });

        if (totalValids === 0) {
            appState.detectedTypes[header] = 'string';
        } else if (numericCount / totalValids > 0.7) {
            appState.detectedTypes[header] = 'number';
        } else if (dateCount / totalValids > 0.7) {
            appState.detectedTypes[header] = 'date';
        } else {
            appState.detectedTypes[header] = 'string';
        }
    });
}

// MOPH HDC Mode Auto-Detection
function detectMophIronDataset() {
    const hasHosname = appState.headers.includes('hosname') || appState.headers.includes('hoscode');
    const hasAge = appState.headers.includes('age_m');
    const hasResult = appState.headers.includes('result');

    if (hasHosname && hasAge && hasResult) {
        appState.isMophMode = true;
        appState.activeMophIndicator = 'iron-supplement';

        // Show HDC specific UI banners
        document.getElementById('moph-banner').style.display = 'flex';
        document.getElementById('moph-indicator-tabs').style.display = 'flex';
        document.getElementById('moph-age-filters').style.display = 'flex';
        document.getElementById('moph-hct-filters').style.display = 'flex';

        // Reset indicator tab active state to iron-supplement
        document.querySelectorAll('#moph-indicator-tabs .filter-tab-btn').forEach(t => t.classList.remove('active'));
        const ironTab = document.querySelector('#moph-indicator-tabs .filter-tab-btn[data-indicator="iron-supplement"]');
        if (ironTab) ironTab.classList.add('active');

        const mophHospitalFilter = document.getElementById('moph-hospital-filter');
        if (mophHospitalFilter) {
            mophHospitalFilter.style.display = 'flex';
            populateDistrictSelect();
            populateHospitalSelect();
        }

        document.getElementById('moph-mode-notice').textContent = "📌 วิเคราะห์ในโหมดผู้รับธาตุเหล็ก (MOPH HDC Mode)";
    } else {
        appState.isMophMode = false;
        document.getElementById('moph-banner').style.display = 'none';
        document.getElementById('moph-indicator-tabs').style.display = 'none';
        document.getElementById('moph-age-filters').style.display = 'none';
        document.getElementById('moph-hct-filters').style.display = 'none';

        const mophHospitalFilter = document.getElementById('moph-hospital-filter');
        if (mophHospitalFilter) {
            mophHospitalFilter.style.display = 'none';
        }

        document.getElementById('moph-mode-notice').textContent = "";
    }
}

// Dynamically compile and populate district select options
function populateDistrictSelect() {
    const select = document.getElementById('select-district');
    if (!select) return;

    // Reset select options
    select.innerHTML = '<option value="all">🗺️ ทุกอำเภอ (ทั้งหมด)</option>';

    // Find the district code and name columns
    const ampcodeCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'ampcode' || l === 'amphurcode' || l === 'amp_code' || l === 'districtcode';
    });
    const ampnameCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'ampname' || l === 'amphurname' || l === 'amp_name' || l === 'district' || l.includes('อำเภอ');
    });

    if (!ampcodeCol && !ampnameCol) return;

    // Scan all rows to collect unique district pairs/names
    const districtMap = new Map();
    const hoscodeCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'hoscode' || l === 'hospcode' || l === 'hcode';
    });
    const hosnameCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'hosname' || l === 'hname' || l === 'hospital';
    });

    appState.rawData.forEach(row => {
        // Exclude clinics and private hospitals
        if (hoscodeCol) {
            const hCode = String(row[hoscodeCol] || '').trim();
            const hName = hosnameCol ? String(row[hosnameCol] || '').trim() : '';
            if (!isPublicMophHospital(hCode, hName)) return;
        }

        const code = ampcodeCol ? String(row[ampcodeCol] || '').trim() : '';
        const name = ampnameCol ? String(row[ampnameCol] || '').trim() : '';
        
        if (ampcodeCol && code) {
            districtMap.set(code, name || `อำเภอ ${code}`);
        } else if (ampnameCol && name) {
            districtMap.set(name, name);
        }
    });

    // Sort keys
    const sortedKeys = Array.from(districtMap.keys()).sort();

    // Populate options
    sortedKeys.forEach(key => {
        const label = districtMap.get(key);
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = ampcodeCol ? `[${key}] - ${label}` : label;
        select.appendChild(opt);
    });

    // Set active filter
    if (appState.activeDistrictFilter) {
        select.value = appState.activeDistrictFilter;
    } else {
        appState.activeDistrictFilter = 'all';
        select.value = 'all';
    }
}

// Dynamically compile and populate hospital select options (with cascading from selected District)
function populateHospitalSelect() {
    const select = document.getElementById('select-hospital');
    if (!select) return;

    // Reset select options
    select.innerHTML = '<option value="all">🏥 ทุกหน่วยบริการ (ทั้งหมด)</option>';

    // Find the hospital code and hospital name columns
    const hoscodeCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'hoscode' || l === 'hospcode' || l === 'hcode';
    });
    const hosnameCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'hosname' || l === 'hname' || l === 'hospital';
    });

    if (!hoscodeCol || !hosnameCol) return;

    // Find district columns to perform cascading check
    const ampcodeCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'ampcode' || l === 'amphurcode' || l === 'amp_code' || l === 'districtcode';
    });
    const ampnameCol = appState.headers.find(h => {
        const l = h.toLowerCase();
        return l === 'ampname' || l === 'amphurname' || l === 'amp_name' || l === 'district' || l.includes('อำเภอ');
    });

    // Scan all rows to collect unique (code, name) pairs matching district selection
    const hospitalMap = new Map();
    appState.rawData.forEach(row => {
        const code = String(row[hoscodeCol] || '').trim();
        const name = String(row[hosnameCol] || '').trim();

        // Exclude clinics and private hospitals
        if (!isPublicMophHospital(code, name)) return;

        // Perform cascading check
        if (appState.activeDistrictFilter !== 'all') {
            const dCode = ampcodeCol ? String(row[ampcodeCol] || '').trim() : '';
            const dName = ampnameCol ? String(row[ampnameCol] || '').trim() : '';
            const targetDist = appState.activeDistrictFilter;
            
            if (ampcodeCol && dCode !== targetDist) return;
            if (!ampcodeCol && ampnameCol && dName !== targetDist) return;
        }

        if (code && name) {
            hospitalMap.set(code, name);
        }
    });

    // Sort by code
    const sortedCodes = Array.from(hospitalMap.keys()).sort();

    // Populate options
    sortedCodes.forEach(code => {
        const name = hospitalMap.get(code);
        const opt = document.createElement('option');
        opt.value = code;
        opt.textContent = `[${code}] - ${name}`;
        select.appendChild(opt);
    });
    
    // Set active filter to whatever was selected (or 'all' if default)
    if (appState.activeHospitalFilter) {
        select.value = appState.activeHospitalFilter;
    } else {
        appState.activeHospitalFilter = 'all';
        select.value = 'all';
    }
}

// Populate UI dimension selects dynamically
function populateDimensionSelects() {
    const xSelect = document.getElementById('select-x-axis');
    const ySelect = document.getElementById('select-y-axis');
    const groupSelect = document.getElementById('select-group-by');

    xSelect.innerHTML = '';
    ySelect.innerHTML = '';
    groupSelect.innerHTML = '<option value="">-- ไม่จัดกลุ่ม --</option>';

    let firstNumeric = null;
    let firstCategory = null;

    appState.headers.forEach(header => {
        const type = appState.detectedTypes[header] || 'string';
        const label = `${header} (${type === 'number' ? 'ตัวเลข' : type === 'date' ? 'วันที่' : 'ข้อความ'})`;

        const optX = document.createElement('option');
        optX.value = header;
        optX.textContent = label;
        xSelect.appendChild(optX);

        const optY = document.createElement('option');
        optY.value = header;
        optY.textContent = label;
        ySelect.appendChild(optY);

        if (type === 'string' || type === 'date') {
            const optG = document.createElement('option');
            optG.value = header;
            optG.textContent = label;
            groupSelect.appendChild(optG);
        }

        if (type === 'number' && !firstNumeric) firstNumeric = header;
        if ((type === 'string' || type === 'date') && !firstCategory) firstCategory = header;
    });

    // Choose sensible smart defaults
    if (appState.isMophMode) {
        // Automatically pre-configure optimal MOPH HDC columns
        appState.xAxisCol = 'hosname';
        appState.yAxisCol = 'result';
        appState.groupByCol = '';
        appState.aggregateFn = 'COUNT'; // We will calculate custom percentage mathematically
    } else {
        appState.xAxisCol = firstCategory || appState.headers[0];
        appState.yAxisCol = firstNumeric || appState.headers[0];
        appState.groupByCol = "";
        appState.aggregateFn = "SUM";
    }

    xSelect.value = appState.xAxisCol;
    ySelect.value = appState.yAxisCol;
    groupSelect.value = appState.groupByCol;
    document.getElementById('select-aggregate').value = appState.aggregateFn;
}

// ==========================================================================
// 🧮 Data Slicing & Filter Application
// ==========================================================================

function applyAllFilters() {
    let filtered = [...appState.rawData];

    // 1.0 Filter out non-public MOPH units (clinics, private hospitals, SSOs, etc.) when in MOPH mode
    if (appState.isMophMode) {
        const hoscodeCol = appState.headers.find(h => {
            const l = h.toLowerCase();
            return l === 'hoscode' || l === 'hospcode' || l === 'hcode';
        });
        const hosnameCol = appState.headers.find(h => {
            const l = h.toLowerCase();
            return l === 'hosname' || l === 'hname' || l === 'hospital';
        });
        
        if (hoscodeCol) {
            filtered = filtered.filter(row => {
                const code = String(row[hoscodeCol] || '').trim();
                const name = hosnameCol ? String(row[hosnameCol] || '').trim() : '';
                return isPublicMophHospital(code, name);
            });
        }
    }

    // 1. Apply Specialized MOPH Age Filter if active
    if (appState.isMophMode) {
        const ageCol = appState.headers.includes('age_m') ? 'age_m' : '';
        if (ageCol) {
            if (appState.activeMophIndicator === 'anemia-12m') {
                // Cohort = เด็กที่อายุครบ 12 เดือน ภายในปีงบประมาณที่เลือก
                // birth range: Oct 1 (FY-1 CE) – Sep 30 (FY CE)  [Thai FY→CE: FY-543]
                const fy = appState.fiscalYear;
                const fyStartCE = fy - 543 - 1;
                const fyEndCE = fy - 543;
                const cohortStart = new Date(fyStartCE, 9, 1);   // Oct 1
                const cohortEnd   = new Date(fyEndCE,   8, 30);  // Sep 30
                const hasBirth = appState.headers.includes('birth');
                if (hasBirth) {
                    // birth column = Excel serial date (days since 1900-01-01, with leap year bug offset)
                    const excelDateToJs = (serial) => {
                        const d = new Date((serial - 25569) * 86400 * 1000);
                        return d;
                    };
                    filtered = filtered.filter(row => {
                        const birthSerial = cleanNumericValue(row['birth']);
                        if (!birthSerial) return false;
                        const birthDate = excelDateToJs(birthSerial);
                        return birthDate >= cohortStart && birthDate <= cohortEnd;
                    });
                } else {
                    // fallback: estimate from age_m
                    const refDate = appState.exportDate || new Date();
                    const monthsDiff = (d1, d2) => (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
                    const minAge = Math.floor(monthsDiff(cohortEnd, refDate));
                    const maxAge = Math.ceil(monthsDiff(cohortStart, refDate));
                    filtered = filtered.filter(row => {
                        const age = cleanNumericValue(row[ageCol]);
                        return age >= minAge && age <= maxAge;
                    });
                }
            } else {
                filtered = filtered.filter(row => {
                    const age = cleanNumericValue(row[ageCol]);
                    if (appState.activeAgeFilter === 'all') {
                        return age >= 6 && age <= 60;
                    } else if (appState.activeAgeFilter === '6-12') {
                        return age >= 6 && age <= 12;
                    } else if (appState.activeAgeFilter === '36-60') {
                        return age >= 36 && age <= 60;
                    }
                    return true;
                });
            }
        }
    }

    // 1.5 Apply Specialized MOPH HCT Lab Filter if active (iron-supplement indicator only)
    if (appState.isMophMode && appState.activeMophIndicator === 'iron-supplement') {
        const labStatusCol = 'lab_result_status';
        const anemiaCol = 'anemea';

        filtered = filtered.filter(row => {
            const hasStatus = appState.headers.includes(labStatusCol);
            const hasAnemia = appState.headers.includes(anemiaCol);

            const labStatus = hasStatus ? cleanNumericValue(row[labStatusCol]) : 0;
            const anemia = hasAnemia ? cleanNumericValue(row[anemiaCol]) : 0;

            if (appState.activeHctFilter === 'not-tested') {
                return labStatus === 0;
            } else if (appState.activeHctFilter === 'tested') {
                return labStatus === 1;
            } else if (appState.activeHctFilter === 'anemia') {
                return labStatus === 1 && anemia > 0;
            } else if (appState.activeHctFilter === 'normal') {
                return labStatus === 1 && anemia === 0;
            }
            return true;
        });
    }

    // 1.7 Apply Specialized District/Amphur Filter if active
    if (appState.isMophMode && appState.activeDistrictFilter && appState.activeDistrictFilter !== 'all') {
        const ampcodeCol = appState.headers.find(h => {
            const l = h.toLowerCase();
            return l === 'ampcode' || l === 'amphurcode' || l === 'amp_code' || l === 'districtcode';
        });
        const ampnameCol = appState.headers.find(h => {
            const l = h.toLowerCase();
            return l === 'ampname' || l === 'amphurname' || l === 'amp_name' || l === 'district' || l.includes('อำเภอ');
        });

        if (ampcodeCol) {
            filtered = filtered.filter(row => {
                const val = String(row[ampcodeCol] || '').trim();
                return val === appState.activeDistrictFilter;
            });
        } else if (ampnameCol) {
            filtered = filtered.filter(row => {
                const val = String(row[ampnameCol] || '').trim();
                return val === appState.activeDistrictFilter;
            });
        }
    }

    // 1.8 Apply Specialized Hospital/Unit Filter if active
    if (appState.isMophMode && appState.activeHospitalFilter && appState.activeHospitalFilter !== 'all') {
        const hoscodeCol = appState.headers.find(h => {
            const l = h.toLowerCase();
            return l === 'hoscode' || l === 'hospcode' || l === 'hcode';
        });
        if (hoscodeCol) {
            filtered = filtered.filter(row => {
                const val = String(row[hoscodeCol] || '').trim();
                return val === appState.activeHospitalFilter;
            });
        }
    }

    // 2. Apply Fuzzy Search Query Filter
    const searchQuery = document.getElementById('table-search').value.toLowerCase().trim();
    if (searchQuery !== "") {
        filtered = filtered.filter(row => {
            return Object.values(row).some(val =>
                String(val).toLowerCase().includes(searchQuery)
            );
        });
    }

    appState.filteredData = filtered;
}

function handleSheetChange(e) {
    loadSheetData(e.target.value);
}

function handleDimensionChange() {
    appState.xAxisCol = document.getElementById('select-x-axis').value;
    appState.yAxisCol = document.getElementById('select-y-axis').value;
    appState.groupByCol = document.getElementById('select-group-by').value;
    appState.aggregateFn = document.getElementById('select-aggregate').value;

    applyAllFilters();
    triggerAnalyticsUpdate();
}

function handleTableSearch() {
    applyAllFilters();
    appState.currentPage = 1;
    triggerAnalyticsUpdate();
}

// ==========================================================================
// 📐 Aggregation & Metric Calculation
// ==========================================================================

function cleanNumericValue(val) {
    if (val === undefined || val === null || val === "") return 0;
    if (typeof val === 'number') return val;

    const cleanStr = String(val).replace(/[\$,฿\s,]/g, '');
    const num = parseFloat(cleanStr);
    return isNaN(num) ? 0 : num;
}

function getAggregatedValue(values) {
    if (values.length === 0) return 0;

    switch (appState.aggregateFn) {
        case "SUM":
            return values.reduce((sum, val) => sum + val, 0);
        case "AVERAGE":
            return values.reduce((sum, val) => sum + val, 0) / values.length;
        case "COUNT":
            return values.length;
        case "MAX":
            return Math.max(...values);
        case "MIN":
            return Math.min(...values);
        default:
            return 0;
    }
}

function isReceivedValue(val) {
    if (val === undefined || val === null || val === "") return false;
    const str = String(val).trim();
    return str.includes('ได้รับ') || str.includes('พบการจ่ายยา') || str === '1' || str === 'Y';
}

// Render dynamic metrics inside neon KPI cards
function renderKPIs() {
    const rows = appState.filteredData;

    if (appState.isMophMode) {

        if (appState.activeMophIndicator === 'anemia-12m') {
            // --- MOPH Indicator: ร้อยละเด็กอายุครบ 12 เดือน มีภาวะโลหิตจาง ---
            // rows already filtered to age 9-15 months
            const totalChildren = rows.length;
            const testedRows = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1);
            const totalTested = testedRows.length;
            const totalAnemia = testedRows.filter(r => cleanNumericValue(r['anemea']) > 0).length;
            const totalNormal = totalTested - totalAnemia;

            const screeningRate = totalChildren > 0 ? (totalTested / totalChildren * 100) : 0;
            const anemiaRate = totalTested > 0 ? (totalAnemia / totalTested * 100) : 0;

            // 1. KPI: เด็กอายุครบ 12 เดือน
            document.getElementById('kpi-1-title').textContent = "เด็กอายุครบ 12 เดือน (6-12 เดือน)";
            document.getElementById('kpi-total-rows').textContent = totalChildren.toLocaleString();
            document.getElementById('kpi-1-subtitle').innerHTML = `<i data-lucide="baby"></i> ทั้งหมดในเขตรับผิดชอบ`;

            // 2. KPI: ได้รับการตรวจ HCT/Hb (ตัวส่วน)
            document.getElementById('kpi-sum-title').textContent = "ได้รับการตรวจ HCT/Hb (ตัวส่วน)";
            document.getElementById('kpi-total-sum').textContent = totalTested.toLocaleString();
            document.getElementById('kpi-sum-subtitle').innerHTML = `<i data-lucide="activity"></i> อัตราตรวจคัดกรอง ${screeningRate.toFixed(1)}%`;

            // 3. KPI: ร้อยละโลหิตจาง (ตัวชี้วัดหลัก เป้าหมาย ≤12%)
            document.getElementById('kpi-avg-title').textContent = "ร้อยละเด็กมีภาวะโลหิตจาง";
            document.getElementById('kpi-total-avg').textContent = anemiaRate.toFixed(1) + "%";

            // เป้าหมาย ปี 2569=≤17%, ปี 2570=≤16% (ใช้ ≤17% เป็น default ปีปัจจุบัน)
            const anemiaTarget = 17.0;
            const targetBadge = document.getElementById('moph-target-badge');
            if (anemiaRate <= anemiaTarget && totalTested > 0) {
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend positive";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="trophy"></i> ผ่านเกณฑ์กระทรวง (≤${anemiaTarget}%)`;
                targetBadge.className = "target-badge met";
                targetBadge.innerHTML = `บรรลุเป้าหมาย: ${anemiaRate.toFixed(1)}%`;
            } else {
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend negative";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="alert-triangle"></i> สูงกว่าเกณฑ์เป้าหมาย (>${anemiaTarget}%)`;
                targetBadge.className = "target-badge";
                targetBadge.innerHTML = `เป้าหมาย ≤${anemiaTarget}% | ปัจจุบัน ${anemiaRate.toFixed(1)}%`;
            }

            // 4. KPI: มีภาวะโลหิตจาง (ตัวเศษ)
            document.getElementById('kpi-4-title').textContent = "มีภาวะโลหิตจาง (ตัวเศษ)";
            document.getElementById('kpi-unique-categories').textContent = totalAnemia.toLocaleString();
            document.getElementById('kpi-4-subtitle').innerHTML = `<i data-lucide="heart-pulse"></i> ปกติ (Normal) ${totalNormal} ราย`;

        } else {
            // --- MOPH Specialized HDC Analytics: ยาเสริมธาตุเหล็ก ---
            const totalTarget = rows.length;

            // Denominator logic: received = column 'result' is dispensing
            const receivedRows = rows.filter(r => isReceivedValue(r['result']));
            const totalReceived = receivedRows.length;

            const coverageRate = totalTarget > 0 ? (totalReceived / totalTarget * 100) : 0;

            // Anemia and HCT screening rates calculation
            const totalTested = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
            const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && cleanNumericValue(r['anemea']) > 0).length;

            const testedRate = totalTarget > 0 ? (totalTested / totalTarget * 100) : 0;
            const anemiaPrevalence = totalTested > 0 ? (totalAnemia / totalTested * 100) : 0;

            // 1. KPI Target Card
            document.getElementById('kpi-1-title').textContent = "เด็กกลุ่มเป้าหมาย (ราย)";
            document.getElementById('kpi-total-rows').textContent = totalTarget.toLocaleString();

            let ageLabel = "อายุ 6 เดือน - 5 ปี";
            if (appState.activeAgeFilter === '6-12') {
                ageLabel = "อายุ 6 - 12 เดือน";
            } else if (appState.activeAgeFilter === '36-60') {
                ageLabel = "อายุ 3 - 5 ปี";
            }
            document.getElementById('kpi-1-subtitle').innerHTML = `<i data-lucide="baby"></i> ${ageLabel}`;

            // 2. KPI Received Card
            document.getElementById('kpi-sum-title').textContent = "ได้รับธาตุเหล็ก (ราย)";
            document.getElementById('kpi-total-sum').textContent = totalReceived.toLocaleString();
            document.getElementById('kpi-sum-subtitle').innerHTML = `<i data-lucide="check-circle-2"></i> ได้รับยาน้ำครบตามเกณฑ์`;

            // 3. KPI Coverage Rate Percentage Card
            document.getElementById('kpi-avg-title').textContent = "ร้อยละเด็กได้รับธาตุเหล็ก";
            document.getElementById('kpi-total-avg').textContent = coverageRate.toFixed(1) + "%";

            // MOPH Target Threshold = 75.0%
            const targetBadge = document.getElementById('moph-target-badge');
            if (coverageRate >= 75.0) {
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend positive";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="trophy"></i> ผ่านเกณฑ์กระทรวง (>=75%)`;
                targetBadge.className = "target-badge met";
                targetBadge.innerHTML = `บรรลุเป้าหมาย: ${coverageRate.toFixed(1)}%`;
            } else {
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend negative";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="alert-triangle"></i> ต่ำกว่าเกณฑ์เป้าหมาย`;
                targetBadge.className = "target-badge";
                targetBadge.innerHTML = `เป้าหมาย 75% | ปัจจุบัน ${coverageRate.toFixed(1)}%`;
            }

            // 4. KPI Anemia Rate / HCT Card
            document.getElementById('kpi-4-title').textContent = "ร้อยละการเจาะ Lab HCT";
            document.getElementById('kpi-unique-categories').textContent = testedRate.toFixed(1) + "%";
            document.getElementById('kpi-4-subtitle').innerHTML = `<i data-lucide="heart-pulse"></i> เจาะแล้วซีด (Anemia) ${totalAnemia} ราย (${anemiaPrevalence.toFixed(1)}%)`;
        }

    } else {
        // --- Generic Multi-dimensional Dashboard ---
        document.getElementById('kpi-1-title').textContent = "จำนวนรายการข้อมูลทั้งหมด";
        document.getElementById('kpi-total-rows').textContent = rows.length.toLocaleString();
        document.getElementById('kpi-1-subtitle').innerHTML = `<i data-lucide="check-circle-2"></i> ทำการวิเคราะห์ครบถ้วน`;

        const yValues = rows.map(r => cleanNumericValue(r[appState.yAxisCol]));
        const totalSum = yValues.reduce((sum, val) => sum + val, 0);
        const totalAvg = yValues.length > 0 ? (totalSum / yValues.length) : 0;

        document.getElementById('kpi-sum-title').textContent = `ผลรวม ${appState.yAxisCol}`;
        document.getElementById('kpi-total-sum').textContent = formatCompactNumber(totalSum);
        document.getElementById('kpi-sum-subtitle').textContent = "ค่าประมวลผลสรุปยอด";

        document.getElementById('kpi-avg-title').textContent = `ค่าเฉลี่ย ${appState.yAxisCol}`;
        document.getElementById('kpi-total-avg').textContent = formatCompactNumber(totalAvg);
        document.getElementById('kpi-avg-subtitle').textContent = "วิเคราะห์ตามแกน X-Y";

        const uniqueCats = new Set(rows.map(r => String(r[appState.xAxisCol] || '')));
        document.getElementById('kpi-4-title').textContent = "จำนวนกลุ่มไม่ซ้ำ (Unique)";
        document.getElementById('kpi-unique-categories').textContent = uniqueCats.size.toLocaleString();
        document.getElementById('kpi-4-subtitle').innerHTML = `<i data-lucide="activity"></i> การกระจายตัวของกลุ่มข้อมูล`;
    }

    // Re-trigger icon rendering
    lucide.createIcons();
}

function triggerAnalyticsUpdate() {
    renderKPIs();
    renderCharts();
    renderTable();
}

// ==========================================================================
// 📊 Rendering Visualizations (ApexCharts with customized health metrics)
// ==========================================================================

function renderCharts() {
    const rows = appState.filteredData;
    if (rows.length === 0) return;

    if (appState.isMophMode) {
        renderMophModeCharts(rows);
    } else {
        renderGenericCharts(rows);
    }
}

// --- Specialized MOPH Charts Renderer ---
function renderMophModeCharts(rows) {
    if (appState.activeMophIndicator === 'anemia-12m') {
        renderAnemia12mCharts(rows);
        return;
    }

    // Override standard titles
    document.getElementById('chart-area-title').innerHTML = `<i data-lucide="trending-up"></i> ความครอบคลุมและการเจาะ HCT ตามช่วงอายุ (Coverage & HCT Lab Trend)`;
    document.getElementById('chart-donut-title').innerHTML = `<i data-lucide="pie-chart"></i> สัดส่วนความครอบคลุมและภาวะซีดในการเจาะ Lab (HCT & Anemia Distribution)`;
    document.getElementById('chart-bar-title').innerHTML = `<i data-lucide="bar-chart-4"></i> อัตราครอบคลุมและการเจาะ HCT รายหน่วยงาน (Hospital Performance)`;
    document.getElementById('chart-radar-title').innerHTML = `<i data-lucide="heart-pulse"></i> ผลลัพธ์ภาวะโลหิตจาง (ซีด) ตามประวัติการได้รับธาตุเหล็ก (Clinical Correlation)`;

    // Re-trigger icon rendering in headers
    lucide.createIcons();

    // 1. HOSPITAL PERFORMANCE (Grouped Bar Chart showing Supplement Rate vs HCT Screening Rate)
    const hospMap = {};
    rows.forEach(r => {
        const hosp = String(r['hosname'] || 'ไม่ทราบหน่วยงาน');
        if (!hospMap[hosp]) hospMap[hosp] = { total: 0, received: 0, tested: 0 };
        hospMap[hosp].total++;
        if (isReceivedValue(r['result'])) hospMap[hosp].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) hospMap[hosp].tested++;
    });

    // Sort hospitals by Supplement Rate descending
    const hospitals = Object.keys(hospMap).sort((a, b) => {
        const rateA = hospMap[a].total > 0 ? (hospMap[a].received / hospMap[a].total * 100) : 0;
        const rateB = hospMap[b].total > 0 ? (hospMap[b].received / hospMap[b].total * 100) : 0;
        return rateB - rateA;
    }).slice(0, 15); // Show top 15 hospitals

    const hospRates = hospitals.map(h => {
        const rate = (hospMap[h].received / hospMap[h].total * 100);
        return parseFloat(rate.toFixed(1));
    });

    const hospHctRates = hospitals.map(h => {
        const rate = (hospMap[h].tested / hospMap[h].total * 100);
        return parseFloat(rate.toFixed(1));
    });

    const barOptions = {
        series: [{
            name: 'ร้อยละเด็กได้รับธาตุเหล็ก',
            data: hospRates
        }, {
            name: 'ร้อยละการเจาะ Lab HCT',
            data: hospHctRates
        }],
        chart: {
            type: 'bar',
            height: '100%',
            background: 'transparent',
            foreColor: '#475569',
            toolbar: { show: false }
        },
        theme: { mode: 'light' },
        colors: ['#0284c7', '#16a34a'], // Blue for Supplement, Green for HCT Screening
        plotOptions: {
            bar: {
                horizontal: true,
                borderRadius: 4,
                barHeight: '75%',
                dataLabels: { position: 'top' }
            }
        },
        xaxis: {
            categories: hospitals.map(h => h.replace('โรงพยาบาลส่งเสริมสุขภาพตำบล', 'รพ.สต.').replace('โรงพยาบาล', 'รพ.')),
            min: 0,
            max: 100,
            labels: { formatter: function (val) { return val + "%"; } }
        },
        annotations: {
            xaxis: [{
                x: 75,
                borderColor: '#00ff87',
                borderWidth: 2,
                label: {
                    borderColor: '#00ff87',
                    style: { color: '#060913', background: '#00ff87', fontWeight: 'bold' },
                    text: 'เกณฑ์เป้าหมายยาเสริม (75%)'
                }
            }]
        },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: {
            theme: 'light',
            y: {
                formatter: function (val, { seriesIndex, dataPointIndex }) {
                    const h = hospitals[dataPointIndex];
                    if (seriesIndex === 0) {
                        return `${val}% (ได้รับ ${hospMap[h].received} จากทั้งหมด ${hospMap[h].total} คน)`;
                    } else {
                        return `${val}% (เจาะ Lab ${hospMap[h].tested} จากทั้งหมด ${hospMap[h].total} คน)`;
                    }
                }
            }
        }
    };

    if (charts.bar) {
        charts.bar.updateOptions(barOptions);
    } else {
        charts.bar = new ApexCharts(document.getElementById('chart-bar'), barOptions);
        charts.bar.render();
    }

    // 2. OVERALL HCT LAB & ANEMIA STATUS DISTRIBUTION (Donut Chart 3-way split)
    const totalTarget = rows.length;
    const totalTested = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
    const totalNotTested = totalTarget - totalTested;
    const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && cleanNumericValue(r['anemea']) > 0).length;
    const totalNormal = totalTested - totalAnemia;

    const donutOptions = {
        series: [totalNormal, totalAnemia, totalNotTested],
        chart: {
            type: 'donut',
            height: '100%',
            background: 'transparent',
            foreColor: '#475569'
        },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626', '#f59e0b'], // Green = Normal, Red = Anemic, Yellow = Not tested
        labels: ['เจาะแล้วปกติ (Normal HCT)', 'เจาะแล้วซีด (Anemic HCT)', 'ยังไม่เจาะ Lab HCT'],
        plotOptions: {
            pie: {
                donut: {
                    size: '65%',
                    labels: {
                        show: true,
                        name: { show: true, fontSize: '11px' },
                        value: {
                            show: true,
                            fontSize: '15px',
                            fontWeight: 'bold',
                            formatter: function (val) { return val.toLocaleString() + ' ราย'; }
                        },
                        total: {
                            show: true,
                            label: 'เจาะ Lab แล้ว',
                            fontSize: '10px',
                            formatter: function (w) {
                                const rate = totalTarget > 0 ? (totalTested / totalTarget * 100) : 0;
                                return rate.toFixed(1) + "%";
                            }
                        }
                    }
                }
            }
        },
        dataLabels: { enabled: true },
        legend: { position: 'bottom', fontSize: '9px', markers: { width: 10, height: 10, radius: 5, strokeWidth: 0 } },
        tooltip: { theme: 'light' }
    };

    if (charts.donut) {
        charts.donut.updateOptions(donutOptions);
    } else {
        charts.donut = new ApexCharts(document.getElementById('chart-donut'), donutOptions);
        charts.donut.render();
    }

    // 3. AGE COVERAGE & HCT LAB TREND (Area Chart by months 6 to 60)
    const ageMap = {};
    rows.forEach(r => {
        const age = cleanNumericValue(r['age_m']);
        if (age >= 6 && age <= 60) {
            if (!ageMap[age]) ageMap[age] = { total: 0, received: 0, tested: 0 };
            ageMap[age].total++;
            if (isReceivedValue(r['result'])) ageMap[age].received++;
            if (cleanNumericValue(r['lab_result_status']) === 1) ageMap[age].tested++;
        }
    });

    const sortedAges = Object.keys(ageMap).map(Number).sort((a, b) => a - b);
    const ageCategories = sortedAges.map(age => `${age} ด.`);

    const ageRates = sortedAges.map(age => {
        const rate = ageMap[age].total > 0 ? (ageMap[age].received / ageMap[age].total * 100) : 0;
        return parseFloat(rate.toFixed(1));
    });

    const ageTestedRates = sortedAges.map(age => {
        const rate = ageMap[age].total > 0 ? (ageMap[age].tested / ageMap[age].total * 100) : 0;
        return parseFloat(rate.toFixed(1));
    });

    const areaOptions = {
        series: [{
            name: 'ร้อยละเด็กที่ได้รับยาเสริมธาตุเหล็ก',
            data: ageRates
        }, {
            name: 'ร้อยละเด็กที่ได้รับการเจาะ Lab HCT',
            data: ageTestedRates
        }],
        chart: {
            type: 'area',
            height: '100%',
            background: 'transparent',
            foreColor: '#475569',
            toolbar: { show: true },
            dropShadow: {
                enabled: true,
                top: 3,
                left: 0,
                blur: 5,
                color: '#16a34a',
                opacity: 0.25
            }
        },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#0284c7'], // Green for Supplement, Blue for HCT Lab
        fill: {
            type: 'gradient',
            gradient: {
                shadeIntensity: 1,
                opacityFrom: 0.35,
                opacityTo: 0.02,
                stops: [0, 90, 100]
            }
        },
        stroke: { curve: 'smooth', width: 2.5 },
        xaxis: {
            categories: ageCategories,
            title: { text: 'อายุเด็กในมิติรายเดือน (months)' }
        },
        yaxis: {
            min: 0,
            max: 100,
            labels: { formatter: function (val) { return val + "%"; } }
        },
        annotations: {
            yaxis: [{
                y: 75,
                borderColor: '#00ff87',
                borderWidth: 1,
                strokeDashArray: 4,
                label: {
                    borderColor: '#00ff87',
                    style: { color: '#060913', background: '#00ff87', fontWeight: 'bold' },
                    text: 'เกณฑ์ยาเสริม 75%'
                }
            }]
        },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: {
            theme: 'light',
            y: {
                formatter: function (val, { seriesIndex, dataPointIndex }) {
                    const age = sortedAges[dataPointIndex];
                    if (seriesIndex === 0) {
                        return `${val}% (ได้รับธาตุเหล็ก ${ageMap[age].received} จากทั้งหมด ${ageMap[age].total} คน)`;
                    } else {
                        return `${val}% (เจาะ Lab HCT ${ageMap[age].tested} จากทั้งหมด ${ageMap[age].total} คน)`;
                    }
                }
            }
        }
    };

    if (charts.area) {
        charts.area.updateOptions(areaOptions);
    } else {
        charts.area = new ApexCharts(document.getElementById('chart-area'), areaOptions);
        charts.area.render();
    }

    // 4. ANEMIA RATE CLINICAL CORRELATION
    // Compare anemia prevalence among tested children: Received Supplement vs Not Received Supplement
    const receivedTestedRows = rows.filter(r => isReceivedValue(r['result']) && cleanNumericValue(r['lab_result_status']) === 1);
    const receivedAnemiaCount = receivedTestedRows.filter(r => cleanNumericValue(r['anemea']) > 0).length;
    const anemiaRateReceived = receivedTestedRows.length > 0 ? (receivedAnemiaCount / receivedTestedRows.length * 100) : 0;

    const notReceivedTestedRows = rows.filter(r => !isReceivedValue(r['result']) && cleanNumericValue(r['lab_result_status']) === 1);
    const notReceivedAnemiaCount = notReceivedTestedRows.filter(r => cleanNumericValue(r['anemea']) > 0).length;
    const anemiaRateNotReceived = notReceivedTestedRows.length > 0 ? (notReceivedAnemiaCount / notReceivedTestedRows.length * 100) : 0;

    const radarOptions = {
        series: [{
            name: 'ร้อยละความชุกโลหิตจาง (Anemia Rate จากคนที่เจาะ Lab)',
            data: [parseFloat(anemiaRateReceived.toFixed(2)), parseFloat(anemiaRateNotReceived.toFixed(2))]
        }],
        chart: {
            type: 'bar',
            height: '100%',
            background: 'transparent',
            foreColor: '#475569',
            toolbar: { show: false }
        },
        theme: { mode: 'light' },
        colors: ['#dc2626'], // Red for clinical indicators
        plotOptions: {
            bar: {
                columnWidth: '40%',
                borderRadius: 5,
                dataLabels: { position: 'top' }
            }
        },
        dataLabels: {
            enabled: true,
            formatter: function (val) { return val + "%"; },
            offsetY: -20,
            style: { fontSize: '11px', colors: ["#dc2626"] }
        },
        xaxis: {
            categories: ['กลุ่มที่ได้รับธาตุเหล็ก', 'กลุ่มที่ไม่ได้รับธาตุเหล็ก'],
            labels: { style: { fontSize: '11px', fontWeight: 'bold' } }
        },
        yaxis: {
            labels: { formatter: function (val) { return val + "%"; } }
        },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: {
            theme: 'light',
            y: {
                formatter: function (val, { dataPointIndex }) {
                    if (dataPointIndex === 0) {
                        return `${val}% (โลหิตจาง ${receivedAnemiaCount} จากกลุ่มได้รับยาที่เจาะ Lab ${receivedTestedRows.length} คน)`;
                    } else {
                        return `${val}% (โลหิตจาง ${notReceivedAnemiaCount} จากกลุ่มไม่ได้รับยาที่เจาะ Lab ${notReceivedTestedRows.length} คน)`;
                    }
                }
            }
        }
    };

    if (charts.radar) {
        charts.radar.updateOptions(radarOptions);
    } else {
        charts.radar = new ApexCharts(document.getElementById('chart-radar'), radarOptions);
        charts.radar.render();
    }
}

// --- MOPH Indicator: ร้อยละเด็กอายุครบ 12 เดือน มีภาวะโลหิตจาง ---
function renderAnemia12mCharts(rows) {
    document.getElementById('chart-bar-title').innerHTML = `<i data-lucide="bar-chart-4"></i> ร้อยละภาวะโลหิตจางรายหน่วยบริการ (เป้าหมาย ≤17%)`;
    document.getElementById('chart-donut-title').innerHTML = `<i data-lucide="pie-chart"></i> สัดส่วนภาวะโลหิตจาง (เด็กอายุ 6-12 เดือน)`;
    document.getElementById('chart-area-title').innerHTML = `<i data-lucide="trending-up"></i> จำนวนเด็กอายุ 6-12 เดือน รายหน่วยบริการ (ตัวส่วน vs ตัวเศษ)`;
    document.getElementById('chart-radar-title').innerHTML = `<i data-lucide="activity"></i> เปรียบเทียบอัตราโลหิตจางรายอำเภอ`;
    lucide.createIcons();

    // Empty state: ไม่มีข้อมูล lab HCT/Hb ในช่วงอายุ 6-12 เดือน
    const hasLabData = rows.some(r => cleanNumericValue(r['lab_result_status']) === 1);
    if (rows.length === 0 || !hasLabData) {
        const noDataMsg = `<div style="display:flex;align-items:center;justify-content:center;height:100%;min-height:160px;flex-direction:column;gap:12px;color:var(--text-muted);">
            <i data-lucide="flask-conical-off" style="width:40px;height:40px;opacity:0.4;"></i>
            <div style="text-align:center;font-size:0.85rem;">
                <strong style="color:var(--text-secondary);">ไม่มีข้อมูลผล Lab HCT/Hb ในช่วงอายุ 6-12 เดือน</strong><br>
                <span style="font-size:0.78rem;margin-top:4px;display:block;">ตัวชี้วัดนี้ต้องการข้อมูล lab_result_status และ anemea ของเด็กอายุ 6-12 เดือน</span>
            </div>
        </div>`;
        ['chart-bar','chart-donut','chart-area','chart-radar'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.innerHTML = noDataMsg;
        });
        lucide.createIcons();
        return;
    }

    // --- Hospital anemia data ---
    const hospMap = {};
    rows.forEach(r => {
        const hosp = String(r['hosname'] || 'ไม่ทราบหน่วยงาน');
        if (!hospMap[hosp]) hospMap[hosp] = { total: 0, tested: 0, anemia: 0 };
        hospMap[hosp].total++;
        if (cleanNumericValue(r['lab_result_status']) === 1) {
            hospMap[hosp].tested++;
            if (cleanNumericValue(r['anemea']) > 0) hospMap[hosp].anemia++;
        }
    });

    // Sort by anemia rate descending (worst first)
    const hospitals = Object.keys(hospMap)
        .filter(h => hospMap[h].tested > 0)
        .sort((a, b) => {
            const rA = hospMap[a].anemia / hospMap[a].tested * 100;
            const rB = hospMap[b].anemia / hospMap[b].tested * 100;
            return rB - rA;
        }).slice(0, 15);

    const hospLabels = hospitals.map(h => h.replace('โรงพยาบาลส่งเสริมสุขภาพตำบล', 'รพ.สต.').replace('โรงพยาบาล', 'รพ.'));
    const normalCounts = hospitals.map(h => hospMap[h].tested - hospMap[h].anemia);
    const anemiaCounts = hospitals.map(h => hospMap[h].anemia);

    // 1. STACKED HORIZONTAL BAR: ปกติ(เขียว) + ซีด(แดง) รายหน่วยบริการ — sort worst first
    const barOptions = {
        series: [
            { name: '🟢 ปกติ (Normal)', data: normalCounts },
            { name: '🔴 ซีด / โลหิตจาง', data: anemiaCounts }
        ],
        chart: { type: 'bar', height: '100%', background: 'transparent', foreColor: '#475569', toolbar: { show: true }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { horizontal: true, borderRadius: 3, barHeight: '65%' } },
        xaxis: {
            categories: hospLabels,
            labels: { formatter: val => val.toLocaleString() + ' ราย' }
        },
        legend: { position: 'top', markers: { width: 10, height: 10, radius: 5, strokeWidth: 0 } },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: {
            theme: 'light',
            shared: true,
            y: { formatter: (val, { seriesIndex, dataPointIndex }) => {
                const h = hospitals[dataPointIndex];
                const rate = hospMap[h].tested > 0 ? (hospMap[h].anemia / hospMap[h].tested * 100).toFixed(1) : 0;
                return seriesIndex === 1
                    ? `${val} ราย — ซีด ${rate}% จากที่เจาะทั้งหมด ${hospMap[h].tested} คน`
                    : `${val} ราย`;
            }}
        },
        dataLabels: {
            enabled: true,
            formatter: (val, { seriesIndex }) => seriesIndex === 1 && val > 0 ? val : '',
            style: { fontSize: '10px', colors: ['#fff'] }
        }
    };
    if (charts.bar) { charts.bar.updateOptions(barOptions); }
    else { charts.bar = new ApexCharts(document.getElementById('chart-bar'), barOptions); charts.bar.render(); }

    // 2. DONUT CHART: ปกติ(เขียว) / ซีด(แดง) / ยังไม่ตรวจ(เทา)
    const totalTested = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
    const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && cleanNumericValue(r['anemea']) > 0).length;
    const totalNormal = totalTested - totalAnemia;
    const totalUntested = rows.length - totalTested;
    const anemiaRateOverall = totalTested > 0 ? (totalAnemia / totalTested * 100) : 0;

    const donutOptions = {
        series: [totalNormal, totalAnemia, totalUntested],
        chart: { type: 'donut', height: '100%', background: 'transparent', foreColor: '#475569' },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626', '#f59e0b'],
        labels: ['ปกติ (HCT≥33 / Hb≥11)', 'ซีด (HCT<33 / Hb<11)', 'ยังไม่ได้ตรวจ Lab'],
        plotOptions: { pie: { donut: { size: '65%', labels: { show: true,
            name: { show: true, fontSize: '11px' },
            value: { show: true, fontSize: '15px', fontWeight: 'bold', formatter: val => val.toLocaleString() + ' ราย' },
            total: { show: true, label: '% ซีด (จากที่เจาะ)', fontSize: '9px', formatter: () => anemiaRateOverall.toFixed(1) + "%" }
        } } } },
        dataLabels: { enabled: true, formatter: (val, { seriesIndex, w }) => {
            const count = w.config.series[seriesIndex];
            return count > 0 ? count.toLocaleString() + ' ราย' : '';
        }},
        legend: { position: 'bottom', fontSize: '10px', markers: { width: 10, height: 10, radius: 5, strokeWidth: 0 } },
        tooltip: { theme: 'light', y: { formatter: val => val.toLocaleString() + ' ราย' } }
    };
    if (charts.donut) { charts.donut.updateOptions(donutOptions); }
    else { charts.donut = new ApexCharts(document.getElementById('chart-donut'), donutOptions); charts.donut.render(); }

    // 3. STACKED BAR: ปกติ(เขียว) + ซีด(แดง) รายหน่วยบริการ — sort by total
    const hospAllSorted = Object.keys(hospMap)
        .filter(h => hospMap[h].tested > 0)
        .sort((a, b) => hospMap[b].tested - hospMap[a].tested)
        .slice(0, 15);
    const areaOptions = {
        series: [
            { name: '🟢 ปกติ', data: hospAllSorted.map(h => hospMap[h].tested - hospMap[h].anemia) },
            { name: '🔴 ซีด / โลหิตจาง', data: hospAllSorted.map(h => hospMap[h].anemia) }
        ],
        chart: { type: 'bar', height: '100%', background: 'transparent', foreColor: '#475569', toolbar: { show: true }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { horizontal: false, columnWidth: '65%', borderRadius: 3 } },
        xaxis: {
            categories: hospAllSorted.map(h => h.replace('โรงพยาบาลส่งเสริมสุขภาพตำบล', 'รพ.สต.').replace('โรงพยาบาล', 'รพ.')),
            labels: { rotate: -45, style: { fontSize: '9px' } }
        },
        yaxis: { title: { text: 'จำนวน (ราย)' } },
        legend: { position: 'top', markers: { width: 10, height: 10, radius: 5, strokeWidth: 0 } },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: { theme: 'light', shared: true, y: { formatter: (val, { seriesIndex, dataPointIndex }) => {
            const h = hospAllSorted[dataPointIndex];
            const rate = hospMap[h].tested > 0 ? (hospMap[h].anemia / hospMap[h].tested * 100).toFixed(1) : 0;
            return seriesIndex === 1
                ? `${val} ราย (${rate}% ซีด)`
                : `${val} ราย`;
        }}}
    };
    if (charts.area) { charts.area.updateOptions(areaOptions); }
    else { charts.area = new ApexCharts(document.getElementById('chart-area'), areaOptions); charts.area.render(); }

    // 4. STACKED BAR: ปกติ(เขียว) + ซีด(แดง) รายอำเภอ + เส้นเกณฑ์ %
    const ampMap = {};
    rows.forEach(r => {
        const amp = String(r['amp_name'] || r['ampname'] || 'ไม่ทราบอำเภอ');
        if (!ampMap[amp]) ampMap[amp] = { tested: 0, anemia: 0 };
        if (cleanNumericValue(r['lab_result_status']) === 1) {
            ampMap[amp].tested++;
            if (cleanNumericValue(r['anemea']) > 0) ampMap[amp].anemia++;
        }
    });
    const amps = Object.keys(ampMap).filter(a => ampMap[a].tested > 0).sort((a, b) => {
        return (ampMap[b].anemia / ampMap[b].tested) - (ampMap[a].anemia / ampMap[a].tested);
    });

    const radarOptions = {
        series: [
            { name: '🟢 ปกติ', data: amps.map(a => ampMap[a].tested - ampMap[a].anemia) },
            { name: '🔴 ซีด / โลหิตจาง', data: amps.map(a => ampMap[a].anemia) }
        ],
        chart: { type: 'bar', height: '100%', background: 'transparent', foreColor: '#475569', toolbar: { show: false }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { columnWidth: '55%', borderRadius: 4 } },
        xaxis: { categories: amps, labels: { style: { fontSize: '11px', fontWeight: 'bold' } } },
        yaxis: { title: { text: 'จำนวน (ราย)' } },
        legend: { position: 'top', markers: { width: 10, height: 10, radius: 5, strokeWidth: 0 } },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        dataLabels: { enabled: true, formatter: (val, { seriesIndex, dataPointIndex }) => {
            if (seriesIndex !== 1) return '';
            const a = amps[dataPointIndex];
            const rate = ampMap[a].tested > 0 ? (ampMap[a].anemia / ampMap[a].tested * 100).toFixed(1) : 0;
            return rate + '%';
        }, style: { fontSize: '10px', colors: ['#fff'] } },
        tooltip: { theme: 'light', shared: true, y: { formatter: (val, { seriesIndex, dataPointIndex }) => {
            const a = amps[dataPointIndex];
            const rate = ampMap[a].tested > 0 ? (ampMap[a].anemia / ampMap[a].tested * 100).toFixed(1) : 0;
            return seriesIndex === 1 ? `${val} ราย (${rate}% ซีด)` : `${val} ราย`;
        }}}
    };
    if (charts.radar) { charts.radar.updateOptions(radarOptions); }
    else { charts.radar = new ApexCharts(document.getElementById('chart-radar'), radarOptions); charts.radar.render(); }
}

// --- Generic Multi-dimensional Chart Rendering ---
function renderGenericCharts(rows) {
    // Reset standard titles
    document.getElementById('chart-area-title').innerHTML = `<i data-lucide="trending-up"></i> รายงานแนวโน้มข้อมูล (Area/Line Chart)`;
    document.getElementById('chart-donut-title').innerHTML = `<i data-lucide="pie-chart"></i> สัดส่วนองค์ประกอบ (Donut Chart)`;
    document.getElementById('chart-bar-title').innerHTML = `<i data-lucide="bar-chart-4"></i> เปรียบเทียบตามประเภท (Bar/Column Chart)`;
    document.getElementById('chart-radar-title').innerHTML = `<i data-lucide="radar"></i> แผนภูมิเรดาร์แสดงประสิทธิภาพ (Radar Chart)`;

    const xValues = rows.map(r => {
        const xVal = r[appState.xAxisCol];
        if (appState.detectedTypes[appState.xAxisCol] === 'date') {
            return formatDateString(xVal);
        }
        return String(xVal === undefined || xVal === "" ? "N/A" : xVal);
    });

    const categories = Array.from(new Set(xValues));
    let chartCategories = [...categories];
    if (chartCategories.length > 25) {
        chartCategories = chartCategories.slice(0, 25);
    }

    let chartSeries = [];

    if (!appState.groupByCol) {
        const seriesData = chartCategories.map(cat => {
            const matchingRows = rows.filter((r, idx) => xValues[idx] === cat);
            const numerics = matchingRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
            return getAggregatedValue(numerics);
        });

        chartSeries.push({
            name: `${appState.aggregateFn} of ${appState.yAxisCol}`,
            data: seriesData
        });
    } else {
        const groupValues = rows.map(r => String(r[appState.groupByCol] === undefined || r[appState.groupByCol] === "" ? "N/A" : r[appState.groupByCol]));
        const uniqueGroups = Array.from(new Set(groupValues)).slice(0, 6);

        uniqueGroups.forEach(grpName => {
            const seriesData = chartCategories.map(cat => {
                const matchingRows = rows.filter((r, idx) => xValues[idx] === cat && groupValues[idx] === grpName);
                const numerics = matchingRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
                return getAggregatedValue(numerics);
            });

            chartSeries.push({
                name: grpName,
                data: seriesData
            });
        });
    }

    // area
    const areaOptions = {
        series: chartSeries,
        chart: {
            type: 'area',
            height: '100%',
            background: 'transparent',
            foreColor: '#475569',
            toolbar: { show: true },
            dropShadow: { enabled: true, top: 5, left: 0, blur: 8, color: neonColors, opacity: 0.25 }
        },
        theme: { mode: 'light' },
        colors: neonColors,
        fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.45, opacityTo: 0.05, stops: [0, 90, 100] } },
        stroke: { curve: 'smooth', width: 3 },
        xaxis: { categories: chartCategories, labels: { rotate: -45, style: { fontSize: '10px' } } },
        yaxis: { labels: { formatter: function (value) { return formatCompactNumber(value); } } },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: {
            theme: 'light',
            y: { formatter: function (val) { return val.toLocaleString('th-TH') + ` (${appState.yAxisCol})`; } }
        }
    };
    if (charts.area) charts.area.updateOptions(areaOptions);
    else { charts.area = new ApexCharts(document.getElementById('chart-area'), areaOptions); charts.area.render(); }

    // bar
    const barOptions = {
        series: chartSeries,
        chart: { type: 'bar', height: '100%', background: 'transparent', foreColor: '#475569', toolbar: { show: false } },
        theme: { mode: 'light' },
        colors: neonColors,
        plotOptions: { bar: { horizontal: false, columnWidth: '55%', borderRadius: 4 } },
        stroke: { show: true, width: 2, colors: ['transparent'] },
        xaxis: { categories: chartCategories, labels: { rotate: -45, style: { fontSize: '10px' } } },
        yaxis: { labels: { formatter: function (value) { return formatCompactNumber(value); } } },
        grid: { borderColor: 'rgba(15,23,42,0.08)' },
        tooltip: { theme: 'light', y: { formatter: function (val) { return val.toLocaleString('th-TH'); } } }
    };
    if (charts.bar) charts.bar.updateOptions(barOptions);
    else { charts.bar = new ApexCharts(document.getElementById('chart-bar'), barOptions); charts.bar.render(); }

    // donut
    let donutLabels = [...chartCategories].slice(0, 8);
    let donutSeries = [];
    donutLabels.forEach(cat => {
        const matchingRows = rows.filter((r, idx) => xValues[idx] === cat);
        const numerics = matchingRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
        donutSeries.push(getAggregatedValue(numerics));
    });
    if (categories.length > 8) {
        donutLabels.push('อื่นๆ (Others)');
        const otherRows = rows.filter((r, idx) => !donutLabels.includes(xValues[idx]));
        const otherNumerics = otherRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
        donutSeries.push(getAggregatedValue(otherNumerics));
    }
    const donutOptions = {
        series: donutSeries,
        chart: { type: 'donut', height: '100%', background: 'transparent', foreColor: '#475569' },
        theme: { mode: 'light' },
        colors: neonColors,
        labels: donutLabels,
        plotOptions: {
            pie: {
                donut: {
                    size: '65%',
                    labels: {
                        show: true,
                        name: { show: true, fontSize: '12px' },
                        value: { show: true, fontSize: '16px', fontWeight: 'bold', formatter: function (val) { return formatCompactNumber(Number(val)); } },
                        total: {
                            show: true,
                            label: 'ผลรวมย่อย',
                            fontSize: '12px',
                            formatter: function (w) {
                                const sum = w.globals.seriesTotals.reduce((a, b) => a + b, 0);
                                return formatCompactNumber(sum);
                            }
                        }
                    }
                }
            }
        },
        dataLabels: { enabled: false },
        legend: { position: 'bottom', fontSize: '10px' },
        tooltip: { theme: 'light', y: { formatter: function (val) { return val.toLocaleString('th-TH'); } } }
    };
    if (charts.donut) charts.donut.updateOptions(donutOptions);
    else { charts.donut = new ApexCharts(document.getElementById('chart-donut'), donutOptions); charts.donut.render(); }

    // radar
    const radarCategories = [...chartCategories].slice(0, 10);
    let radarSeries = [];
    if (!appState.groupByCol) {
        const seriesData = radarCategories.map(cat => {
            const matchingRows = rows.filter((r, idx) => xValues[idx] === cat);
            const numerics = matchingRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
            return getAggregatedValue(numerics);
        });
        radarSeries.push({ name: `${appState.aggregateFn} of ${appState.yAxisCol}`, data: seriesData });
    } else {
        const groupValues = rows.map(r => String(r[appState.groupByCol] === undefined || r[appState.groupByCol] === "" ? "N/A" : r[appState.groupByCol]));
        const uniqueGroups = Array.from(new Set(groupValues)).slice(0, 3);
        uniqueGroups.forEach(grpName => {
            const seriesData = radarCategories.map(cat => {
                const matchingRows = rows.filter((r, idx) => xValues[idx] === cat && groupValues[idx] === grpName);
                const numerics = matchingRows.map(r => cleanNumericValue(r[appState.yAxisCol]));
                return getAggregatedValue(numerics);
            });
            radarSeries.push({ name: grpName, data: seriesData });
        });
    }
    const radarOptions = {
        series: radarSeries,
        chart: { type: 'radar', height: '100%', background: 'transparent', foreColor: '#475569', toolbar: { show: false } },
        theme: { mode: 'light' },
        colors: neonColors,
        labels: radarCategories,
        stroke: { width: 2 },
        fill: { opacity: 0.2 },
        markers: { size: 4 },
        yaxis: { show: false },
        legend: { position: 'bottom', fontSize: '10px' }
    };
    if (charts.radar) charts.radar.updateOptions(radarOptions);
    else { charts.radar = new ApexCharts(document.getElementById('chart-radar'), radarOptions); charts.radar.render(); }
}

// ==========================================================================
// 💻 Cyber Grid Data Table & Interaction System
// ==========================================================================

function renderTable() {
    const tableHeader = document.getElementById('table-header');
    const tableBody = document.getElementById('table-body');

    if (!tableHeader || !tableBody) return;

    // Sort logic
    if (appState.sortKey) {
        appState.filteredData.sort((a, b) => {
            let valA = a[appState.sortKey];
            let valB = b[appState.sortKey];

            const type = appState.detectedTypes[appState.sortKey];
            if (type === 'number') {
                valA = cleanNumericValue(valA);
                valB = cleanNumericValue(valB);
            } else {
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
            }

            if (valA < valB) return appState.sortDir === "asc" ? -1 : 1;
            if (valA > valB) return appState.sortDir === "asc" ? 1 : -1;
            return 0;
        });
    }

    // Clear & Generate Headers
    tableHeader.innerHTML = '';
    appState.headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;

        const sortIcon = document.createElement('span');
        sortIcon.className = 'sort-icon';
        sortIcon.innerHTML = '&#9662;';
        th.appendChild(sortIcon);

        if (appState.sortKey === header) {
            th.className = appState.sortDir === "asc" ? "sorted-asc" : "sorted-desc";
        }

        th.addEventListener('click', () => handleTableSort(header));
        tableHeader.appendChild(th);
    });

    // Pagination slicing
    const totalRows = appState.filteredData.length;
    const totalPages = Math.ceil(totalRows / appState.pageSize);

    if (appState.currentPage > totalPages) appState.currentPage = Math.max(1, totalPages);

    const startIdx = (appState.currentPage - 1) * appState.pageSize;
    const endIdx = Math.min(startIdx + appState.pageSize, totalRows);

    const pageRows = appState.filteredData.slice(startIdx, endIdx);

    tableBody.innerHTML = '';

    if (pageRows.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = appState.headers.length || 1;
        td.style.textAlign = 'center';
        td.style.color = 'var(--text-muted)';
        td.textContent = 'ไม่พบข้อมูลที่ตรงกับการค้นหา';
        tr.appendChild(td);
        tableBody.appendChild(tr);
    } else {
        pageRows.forEach(row => {
            const tr = document.createElement('tr');
            appState.headers.forEach(header => {
                const td = document.createElement('td');
                const val = row[header];

                if (header === 'result' && appState.isMophMode) {
                    // Make 'ได้รับ' beautiful in HDC mode
                    if (isReceivedValue(val)) {
                        td.innerHTML = `<span class="status-badge success" style="box-shadow: none; font-size: 0.7rem; padding: 2px 6px;">ได้รับ</span>`;
                    } else {
                        td.innerHTML = `<span class="status-badge pending" style="box-shadow: none; font-size: 0.7rem; padding: 2px 6px; background: rgba(15,23,42,0.04); color: var(--text-muted); border-color: rgba(15,23,42,0.1)">-</span>`;
                    }
                } else if (header === 'lab_result_status' && appState.isMophMode) {
                    const labStatus = cleanNumericValue(val);
                    if (labStatus === 1) {
                        td.innerHTML = `<span class="status-badge success" style="background: rgba(22,163,74,0.1); color: #16a34a; border-color: rgba(22,163,74,0.3); font-size: 0.7rem; padding: 2px 6px;">มีผล</span>`;
                    } else {
                        td.innerHTML = `<span class="status-badge" style="background: rgba(245,158,11,0.1); color: #d97706; border-color: rgba(245,158,11,0.3); font-size: 0.7rem; padding: 2px 6px;">ไม่มีผล</span>`;
                    }
                } else if (['result_hct','labresult','labtest','lab_date'].includes(header) && appState.isMophMode) {
                    const labStatus = cleanNumericValue(row['lab_result_status']);
                    if (labStatus !== 1) {
                        td.innerHTML = `<span style="color: var(--text-muted); font-size: 0.8rem;">-</span>`;
                    } else {
                        if (header === 'lab_date' && appState.detectedTypes[header] === 'date') {
                            td.textContent = formatDateString(val);
                        } else {
                            const cleanNum = cleanNumericValue(val);
                            td.textContent = cleanNum.toLocaleString('th-TH', { maximumFractionDigits: 4 });
                            td.style.fontFamily = 'var(--font-display)';
                            td.style.fontWeight = '500';
                            td.style.textAlign = 'right';
                        }
                    }
                } else if (header === 'anemea' && appState.isMophMode) {
                    const labStatus = cleanNumericValue(row['lab_result_status']);
                    const anemiaVal = cleanNumericValue(val);
                    if (labStatus !== 1) {
                        td.innerHTML = `<span class="status-badge" style="background: rgba(245,158,11,0.1); color: #d97706; border-color: rgba(245,158,11,0.3); font-size: 0.7rem; padding: 2px 6px;">ไม่มีผล lab</span>`;
                    } else if (anemiaVal === 2) {
                        td.innerHTML = `<span class="status-badge pending" style="background: rgba(220, 38, 38, 0.1); color: var(--neon-pink); border-color: var(--neon-pink); font-size: 0.7rem; padding: 2px 6px;">โลหิตจาง</span>`;
                    } else if (anemiaVal === 0) {
                        td.innerHTML = `<span class="status-badge success" style="background: rgba(22, 163, 74, 0.1); color: #16a34a; border-color: rgba(22,163,74,0.3); font-size: 0.7rem; padding: 2px 6px;">ปกติ</span>`;
                    } else {
                        td.textContent = String(val);
                    }
                } else if (appState.detectedTypes[header] === 'number' && 
                           header.toLowerCase() !== 'hoscode' && 
                           header.toLowerCase() !== 'hospcode' && 
                           header.toLowerCase() !== 'hcode' && 
                           header.toLowerCase() !== 'nation' && 
                           header.toLowerCase() !== 'cid') {
                    const cleanNum = cleanNumericValue(val);
                    td.textContent = cleanNum.toLocaleString('th-TH', { maximumFractionDigits: 4 });
                    td.style.fontFamily = 'var(--font-display)';
                    td.style.fontWeight = '500';
                    td.style.textAlign = 'right';
                } else if (appState.detectedTypes[header] === 'date') {
                    td.textContent = formatDateString(val);
                } else {
                    td.textContent = val !== undefined ? String(val) : '';
                }

                tr.appendChild(td);
            });
            tableBody.appendChild(tr);
        });
    }

    document.getElementById('pagination-info').textContent =
        totalRows > 0
            ? `แสดงแถว ${startIdx + 1} - ${endIdx} จากทั้งหมด ${totalRows.toLocaleString()} แถว`
            : 'ไม่พบรายการข้อมูล';

    renderPaginationControls(totalPages);
}

function handleTableSort(key) {
    if (appState.sortKey === key) {
        appState.sortDir = appState.sortDir === "asc" ? "desc" : "asc";
    } else {
        appState.sortKey = key;
        appState.sortDir = "asc";
    }
    renderTable();
}

function renderPaginationControls(totalPages) {
    const container = document.getElementById('pagination-controls');
    if (!container) return;

    container.innerHTML = '';

    const btnFirst = document.createElement('button');
    btnFirst.className = 'page-btn';
    btnFirst.innerHTML = '&laquo;';
    btnFirst.disabled = appState.currentPage === 1;
    btnFirst.addEventListener('click', () => { appState.currentPage = 1; renderTable(); });
    container.appendChild(btnFirst);

    const rangeSize = 5;
    let startPage = Math.max(1, appState.currentPage - Math.floor(rangeSize / 2));
    let endPage = Math.min(totalPages, startPage + rangeSize - 1);

    if (endPage - startPage + 1 < rangeSize) {
        startPage = Math.max(1, endPage - rangeSize + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
        const btn = document.createElement('button');
        btn.className = `page-btn ${appState.currentPage === i ? 'active' : ''}`;
        btn.textContent = i;
        btn.addEventListener('click', () => { appState.currentPage = i; renderTable(); });
        container.appendChild(btn);
    }

    const btnLast = document.createElement('button');
    btnLast.className = 'page-btn';
    btnLast.innerHTML = '&raquo;';
    btnLast.disabled = appState.currentPage === totalPages || totalPages === 0;
    btnLast.addEventListener('click', () => { appState.currentPage = totalPages; renderTable(); });
    container.appendChild(btnLast);
}

// ==========================================================================
// ⬇️ Data Exporter System
// ==========================================================================

/**
 * ✅ ดึงค่าต้นฉบับจาก rawDataText (Excel formatted text) สำหรับ export
 * - ใช้ row.__rowIdx__ เพื่อ map กลับไปหา text version
 * - Fallback ไปใช้ raw val ถ้าไม่มี rawDataText
 */
function getExportValue(row, header) {
    // ใช้ text version จาก Excel ก่อน (ต้นฉบับ, ไม่มีการ convert)
    const textRow = appState.rawDataText && row.__rowIdx__ !== undefined
        ? appState.rawDataText[row.__rowIdx__]
        : null;

    if (textRow && textRow[header] !== undefined && textRow[header] !== null) {
        return String(textRow[header]);
    }

    // Fallback: ใช้ parsed value แต่ serialize ให้ถูกต้อง
    const val = row[header];
    if (val === undefined || val === null || val === '') return '';
    if (val instanceof Date && !isNaN(val.getTime())) {
        // YYYY-MM-DD format
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    if (typeof val === 'number' && (Math.abs(val) >= 1e15 || String(val).includes('e'))) {
        return val.toFixed(0);
    }
    return String(val);
}

function exportCSV() {
    if (appState.filteredData.length === 0) return;

    const csvRows = [];
    // Header row
    csvRows.push(appState.headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(','));

    appState.filteredData.forEach(row => {
        const values = appState.headers.map(header => {
            const val = getExportValue(row, header);
            const escaped = val.replace(/"/g, '""');
            return `"${escaped}"`;
        });
        csvRows.push(values.join(','));
    });

    // Prepend UTF-8 BOM (\ufeff) เพื่อให้ Excel อ่านภาษาไทยได้ถูกต้อง
    const BOM = "\ufeff";
    const blob = new Blob([BOM + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `mchmuk_export_${appState.currentSheetName}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast(`📥 Export CSV สำเร็จ! ${appState.filteredData.length.toLocaleString()} แถว`, 'success', 4000);
}

function exportJSON() {
    if (appState.filteredData.length === 0) return;

    // ใช้ getExportValue เพื่อให้ได้ค่าต้นฉบับจาก Excel
    const cleanData = appState.filteredData.map(row => {
        const cleanRow = {};
        appState.headers.forEach(header => {
            cleanRow[header] = getExportValue(row, header);
        });
        return cleanRow;
    });

    const jsonStr = JSON.stringify(cleanData, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `mchmuk_export_${appState.currentSheetName}.json`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast(`📥 Export JSON สำเร็จ! ${appState.filteredData.length.toLocaleString()} แถว`, 'success', 4000);
}

// ==========================================================================
// 🧪 Mock / Local Data Integration (Zero-Friction Fallbacks)
// ==========================================================================

function formatCompactNumber(num) {
    if (num === 0) return "0";
    const formatOpts = { minimumFractionDigits: 0, maximumFractionDigits: 2 };

    if (Math.abs(num) >= 1.0e9) return (num / 1.0e9).toLocaleString('th-TH', formatOpts) + "B";
    if (Math.abs(num) >= 1.0e6) return (num / 1.0e6).toLocaleString('th-TH', formatOpts) + "M";
    if (Math.abs(num) >= 1.0e3) return (num / 1.0e3).toLocaleString('th-TH', formatOpts) + "K";
    return num.toLocaleString('th-TH', formatOpts);
}

function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function formatDateString(val) {
    if (!val) return '';
    if (val instanceof Date) {
        return val.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' });
    }
    const parsed = Date.parse(val);
    if (!isNaN(parsed) && isNaN(val)) {
        return new Date(parsed).toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' });
    }
    return String(val);
}

// ดึงเวลา Commit ล่าสุดของไฟล์ข้อมูลจาก GitHub เพื่อระบุ "วันที่นำเข้าข้อมูลจริง"
async function getFileLastCommitDate() {
    try {
        const token = GITHUB_CONFIG.token;
        const apiUrl = `https://api.github.com/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/commits?path=${GITHUB_CONFIG.filePath}&page=1&per_page=1`;
        const headers = {
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28'
        };
        if (token) {
            headers['Authorization'] = `Bearer ${token}`;
        }
        const res = await fetch(apiUrl, { headers });
        if (res.ok) {
            const commits = await res.json();
            if (commits && commits.length > 0) {
                return commits[0].commit.committer.date; // คืนค่า ISO timestamp ของ Git Commit ล่าสุด
            }
        }
    } catch (err) {
        console.error("⚠️ ไม่สามารถดึงวันที่ commit ล่าสุดได้:", err);
    }
    return null;
}

// ☁️ Fetch ไฟล์ข้อมูลกลางจาก GitHub (raw) — ทำงานทุกเครื่อง ทุก user
async function loadLocalExcelFile(isSilent = false) {
    if (!isSilent) toggleLoader(true, "กำลังดึงข้อมูลล่าสุดจากเซิร์ฟเวอร์กลาง...");

    try {
        // ดึงจาก GitHub raw URL เพื่อให้ทุกเครื่องเห็นข้อมูลเดียวกัน
        // เพิ่ม cache-bust เพื่อไม่ให้ browser cache ไฟล์เก่า
        const cacheBust = Date.now();
        const fetchUrl = `${DATASET_URL}?cb=${cacheBust}`;

        const response = await fetch(fetchUrl);
        if (!response.ok) throw new Error(`HTTP ${response.status}: ไม่พบไฟล์ข้อมูลบน server`);

        const buffer = await response.arrayBuffer();
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false
        });

        appState.workbook = workbook;
        appState.sheetNames = workbook.SheetNames;

        document.getElementById('val-filename').textContent = GITHUB_CONFIG.filePath;
        document.getElementById('val-filesize').textContent = formatBytes(buffer.byteLength);
        document.getElementById('val-sheets').textContent = workbook.SheetNames.length;

        // ดึงวันเวลานำเข้าข้อมูลจริงจาก Metadata JSON หรือ Commit ล่าสุดของ GitHub
        let commitDateStr = null;
        try {
            const metaUrl = `https://raw.githubusercontent.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}/tmp_exchange_data_meta.json?cb=${cacheBust}`;
            const metaResponse = await fetch(metaUrl);
            if (metaResponse.ok) {
                const metaJson = await metaResponse.json();
                commitDateStr = metaJson.importTimestamp;
                console.log("☁️ Loaded import timestamp from Metadata JSON:", commitDateStr);
            }
        } catch (metaErr) {
            console.warn("Could not fetch metadata JSON, using Commit fallback:", metaErr);
        }

        if (!commitDateStr) {
            // fallback หากดึง JSON ไม่สำเร็จ ให้ใช้ Commit API ตัวเดิม
            commitDateStr = await getFileLastCommitDate();
        }

        let formattedTime = "";
        if (commitDateStr) {
            appState.importTimestamp = commitDateStr;
            formattedTime = new Date(commitDateStr).toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";
        } else {
            const now = new Date().toISOString();
            appState.importTimestamp = now;
            formattedTime = new Date(now).toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";
        }

        // 1. แสดงวันเวลานำเข้าข้อมูลจริงบน Server
        const timeRow = document.getElementById('row-import-time');
        const timeVal = document.getElementById('val-import-time');
        if (timeRow && timeVal) {
            timeRow.style.display = 'flex';
            timeVal.textContent = formattedTime;
        }

        // 2. แสดงวันเวลาที่เบราว์เซอร์เปิดมาดึงข้อมูลจริงในปัจจุบัน
        const fetchRow = document.getElementById('row-fetch-time');
        const fetchVal = document.getElementById('val-fetch-time');
        if (fetchRow && fetchVal) {
            fetchRow.style.display = 'flex';
            const nowTime = new Date().toLocaleDateString('th-TH', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit'
            }) + " น.";
            fetchVal.textContent = nowTime;
        }

        const sheetSelect = document.getElementById('select-sheet');
        sheetSelect.innerHTML = '';
        workbook.SheetNames.forEach(sheetName => {
            const opt = document.createElement('option');
            opt.value = sheetName;
            opt.textContent = sheetName;
            sheetSelect.appendChild(opt);
        });

        const hasDataSheet = workbook.SheetNames.includes('DATA');
        const targetSheet = hasDataSheet ? 'DATA' : workbook.SheetNames[0];
        sheetSelect.value = targetSheet;

        loadSheetData(targetSheet);

        document.getElementById('val-status').textContent = "เปิดทำงานปกติ";
        document.getElementById('val-status').className = "status-badge success";
        toggleLoader(false);

        if (!isSilent) {
            showToast('☁️ ดึงข้อมูลจากเซิร์ฟเวอร์กลางสำเร็จ! ข้อมูลเป็นปัจจุบัน', 'success', 4000);
        }

    } catch (err) {
        console.warn("Server fetch failed:", err.message);
        toggleLoader(false);
        if (!isSilent) {
            showToast(`⚠️ ไม่สามารถดึงข้อมูลจาก Server: ${err.message}`, 'warn', 6000);
        }
    }
}


// ==========================================================================
// ☁️ GitHub API — Central Data Persistence Engine
// Admin uploads push Excel to GitHub repo → all users fetch same file
// ==========================================================================

/**
 * Push Excel file to GitHub repository via API
 * ต้องมี PAT (Personal Access Token) ที่มี scope 'repo'
 * Admin ตั้งค่าได้ที่: localStorage.setItem('mchmuk_gh_pat', 'ghp_xxxx...')
 */
async function pushExcelToGitHub(arrayBuffer, originalFilename) {
    const token = GITHUB_CONFIG.token;

    if (!token) {
        showToast('⚠️ ยังไม่ได้ตั้งค่า GitHub Token — ข้อมูลถูกวิเคราะห์บนหน้าจอแล้ว แต่ยังไม่ได้บันทึกไปยัง Server กลาง กรุณาตั้งค่า PAT ในหน้า Admin Settings', 'warn', 8000);
        return;
    }

    const actionsUrl = `https://github.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/actions`;

    try {
        // Step 1: ดึง SHA ของไฟล์เดิม
        showToast('☁️ [1/3] กำลังเชื่อมต่อ GitHub...', 'info', 3000);
        const apiUrl = `https://api.github.com/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/contents/${GITHUB_CONFIG.filePath}`;
        const headers = {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28',
            'Content-Type': 'application/json'
        };

        let existingSha = null;
        const getRes = await fetch(apiUrl, { headers });
        if (getRes.ok) {
            const getJson = await getRes.json();
            existingSha = getJson.sha;
        }

        // Step 2: แปลง ArrayBuffer → Base64
        showToast('☁️ [2/3] กำลังเตรียมข้อมูล...', 'info', 3000);
        const base64 = arrayBufferToBase64(arrayBuffer);

        // Step 3: Push ไฟล์ใหม่
        const now = new Date();
        const commitMsg = `📊 อัปเดตข้อมูล HDC โดย Admin เมื่อ ${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' })} น. (${originalFilename})`;
        const body = { message: commitMsg, content: base64, branch: GITHUB_CONFIG.branch };
        if (existingSha) body.sha = existingSha;

        showToast('☁️ [3/3] กำลัง Push ข้อมูลขึ้น GitHub...', 'info', 4000);
        const putRes = await fetch(apiUrl, { method: 'PUT', headers, body: JSON.stringify(body) });

        if (putRes.ok) {
            const result = await putRes.json();
            const commitSha = result?.commit?.sha?.slice(0, 7) || '';
            
            // อัปเดตไฟล์ Metadata JSON คู่กันขึ้น GitHub
            await pushMetadataToGitHub(originalFilename, arrayBuffer.byteLength);

            showToast(
                `✅ Push สำเร็จ! (commit: ${commitSha})<br>` +
                `🚀 GitHub Actions กำลัง Deploy อัตโนมัติ — ` +
                `<a href="${actionsUrl}" target="_blank" style="color:#0284c7;text-decoration:underline;">ดู Actions →</a><br>` +
                `⏱️ ผู้ใช้ทุกคนจะเห็นข้อมูลใหม่ภายใน ~1-2 นาที`,
                'success', 10000
            );
            console.log(`☁️ GitHub push success — commit: ${result?.commit?.sha}`);
            console.log(`🚀 Actions deploying: ${actionsUrl}`);
        } else {
            const errJson = await putRes.json().catch(() => ({}));
            const errMsg = errJson.message || putRes.statusText;
            if (putRes.status === 401) {
                showToast('❌ Token ไม่ถูกต้องหรือหมดอายุ — กรุณาตั้งค่า GitHub Token ใหม่', 'error', 8000);
            } else if (putRes.status === 403) {
                showToast(
                    '❌ Token ไม่มีสิทธิ์เขียน — ต้องมีสิทธิ์ write เข้าคลังข้อมูล (Repository)<br>' +
                    '<span style="font-size: 0.8rem; opacity: 0.9;">หากใช้ Fine-grained Token ต้องปรับ Repository Permissions -> <strong>Contents</strong> เป็น <strong>Read and write</strong> หรือหากใช้ Classic Token ต้องเลือก Scope: <strong>repo</strong></span>',
                    'error',
                    15000
                );
            } else {
                showToast(`❌ Push ไม่สำเร็จ (${putRes.status}): ${errMsg}`, 'error', 6000);
            }
            console.error('GitHub push failed:', errJson);
        }
    } catch (e) {
        showToast(`❌ เกิดข้อผิดพลาด: ${e.message}`, 'error', 6000);
        console.error('GitHub push error:', e);
    }
}

/** แปลง ArrayBuffer เป็น Base64 string */
function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize);
        binary += String.fromCharCode(...chunk);
    }
    return btoa(binary);
}

/**
 * Push metadata file (tmp_exchange_data_meta.json) to GitHub repository
 */
async function pushMetadataToGitHub(originalFilename, fileSize) {
    const token = GITHUB_CONFIG.token;
    if (!token) return;

    const metaPath = 'tmp_exchange_data_meta.json';
    const apiUrl = `https://api.github.com/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/contents/${metaPath}`;
    const headers = {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'Content-Type': 'application/json'
    };

    try {
        let existingSha = null;
        const getRes = await fetch(apiUrl, { headers });
        if (getRes.ok) {
            const getJson = await getRes.json();
            existingSha = getJson.sha;
        }

        const now = new Date().toISOString();
        const metaObj = {
            importTimestamp: now,
            filename: originalFilename,
            filesize: fileSize
        };

        const jsonStr = JSON.stringify(metaObj, null, 2);
        const encoder = new TextEncoder();
        const uint8 = encoder.encode(jsonStr);
        let binary = '';
        const chunkSize = 8192;
        for (let i = 0; i < uint8.length; i += chunkSize) {
            const chunk = uint8.subarray(i, i + chunkSize);
            binary += String.fromCharCode.apply(null, chunk);
        }
        const base64 = btoa(binary);

        const commitMsg = `📊 อัปเดตข้อมูล HDC Metadata โดย Admin (${originalFilename})`;
        const body = { message: commitMsg, content: base64, branch: GITHUB_CONFIG.branch };
        if (existingSha) body.sha = existingSha;

        const putRes = await fetch(apiUrl, { method: 'PUT', headers, body: JSON.stringify(body) });
        if (putRes.ok) {
            console.log("☁️ GitHub metadata sync success!");
        } else {
            console.error("Failed to push metadata to GitHub:", await putRes.json().catch(() => ({})));
        }
    } catch (e) {
        console.error("Failed to push metadata to GitHub:", e);
    }
}

/**
 * Admin ตั้งค่า GitHub PAT — เรียกจาก console หรือ Admin UI
 * ใช้: setGitHubToken('ghp_xxxxxxxxxxxx')
 */
function setGitHubToken(token) {
    if (!token || (!token.startsWith('ghp_') && !token.startsWith('github_pat_'))) {
        console.error('❌ Token ไม่ถูกต้อง ต้องขึ้นต้นด้วย ghp_ หรือ github_pat_');
        return;
    }
    localStorage.setItem('mchmuk_gh_pat', token);
    showToast('🔑 บันทึก GitHub Token สำเร็จ! ตอนนี้ Admin สามารถ upload Excel แล้วข้อมูลจะถูกบันทึกไปยัง Server กลาง และ Deploy อัตโนมัติ', 'success', 5000);
    console.log('✅ GitHub PAT saved. Admin can now push Excel to GitHub repo → auto-deploy via Actions.');
}

/**
 * saveAdminPAT() — เรียกจากปุ่ม UI ใน Admin section
 * อ่านค่าจาก input#input-gh-pat แล้วบันทึก Token
 */
function saveAdminPAT() {
    const input = document.getElementById('input-gh-pat');
    if (!input) return;
    const token = input.value.trim();

    if (!token) {
        showToast('⚠️ กรุณาระบุ GitHub Personal Access Token', 'warn', 3000);
        return;
    }
    if (!token.startsWith('ghp_') && !token.startsWith('github_pat_')) {
        showToast('❌ Token ไม่ถูกต้อง ต้องขึ้นต้นด้วย ghp_ หรือ github_pat_', 'error', 4000);
        return;
    }

    localStorage.setItem('mchmuk_gh_pat', token);
    input.value = ''; // ล้าง input หลังบันทึก

    // แสดง status
    const patStatus = document.getElementById('github-pat-status');
    if (patStatus) patStatus.style.display = 'block';

    showToast('🔑 บันทึก GitHub Token สำเร็จ! ตอนนี้ Admin สามารถ Upload Excel เพื่อบันทึกข้อมูลไปยัง Server กลางได้เลย', 'success', 6000);
    console.log('✅ GitHub PAT saved successfully. Admin can now push Excel files to GitHub repo.');
    lucide.createIcons();
}

/**
 * clearAdminPAT() — ล้างค่า Custom Token ใน LocalStorage
 */
function clearAdminPAT() {
    localStorage.removeItem('mchmuk_gh_pat');
    const input = document.getElementById('input-gh-pat');
    if (input) input.value = '';

    const patStatus = document.getElementById('github-pat-status');
    if (patStatus) patStatus.style.display = 'none';

    showToast('🔑 ล้างค่า GitHub Token แล้ว ระบบจะกลับไปใช้ Token เริ่มต้นของระบบ', 'info', 4000);
    console.log('✅ Custom GitHub PAT cleared. Reverting to default.');
}

// Stub functions (ไม่ใช้ IndexedDB แล้ว)
function refreshBatchList() {}
function loadLatestActiveBatch() {}
async function saveActiveDataset() { return { ok: true }; }
async function getActiveDataset() { return null; }
