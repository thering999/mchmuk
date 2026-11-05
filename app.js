/**
 * ==========================================================================
 * 💎 HOLOGRAPHIC EXCEL ANALYTICS DASHBOARD - APP.JS
 * Multi-dimensional Browser-based Excel Analytics & Visualization Engine
 * Specialized for MOPH Standard Report: Children Iron Supplement Syrup Coverage
 * ==========================================================================
 */
console.log("💎 MCHMUK Core Engine v1.6.3 Loaded Successfully");

// ==========================================================================
// ☁️ GitHub Storage Configuration (Central Data Persistence)
// Admin uploads are pushed to GitHub repo so ALL users see the same data.
// Regular users simply fetch the file — no upload needed.
// ==========================================================================
// Upload ผ่าน Cloudflare Worker proxy — token เก็บใน Worker secret ไม่อยู่ใน repo

const GITHUB_CONFIG = {
    owner: 'thering999',
    repo: 'mchmuk',
    branch: 'main',
    filePath: 'tmp_exchange_data.xlsx',
    workerUrl: 'https://mchmuk-upload.habusaya.workers.dev',
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
    fiscalYear: 2570,       // Thai fiscal year (ปีงบประมาณ)
    exportDate: null,       // Date object — วันที่ดาวน์โหลดข้อมูล (null = today)
    activeAgeFilter: "all", // all, 6-12, 36-60
    activeQuarter: "all",   // all | 1-4 (fiscal quarter, iron-supplement mode)
    activeHctFilter: "all",  // all, not-tested, tested, anemia, normal
    activeHospitalFilter: "all",
    activeDistrictFilter: "all",
    activeTrendTab: "monthly",
    importTimestamp: null,
    // Advanced filter & risk state
    activeRiskFilter: 'all',
    advFilterAgeMin: null,
    advFilterAgeMax: null,
    advFilterLabMin: null,
    advFilterLabMax: null,
    advFilterDateFrom: null,
    advFilterDateTo: null,
    advFilterAnemia: 'all',
    _xlsxBuffer: null,
};

// --- ApexCharts Global Instances ---
let charts = {
    area: null,
    bar: null,
    donut: null,
    radar: null,
    trendMain: null,
    trendSide: null,
    scatter: null
};

// --- Health/Government Theme Color Palette ---
const neonColors = ['#0284c7', '#16a34a', '#dc2626', '#0d9488', '#f59e0b', '#6366f1', '#65a30d'];

const isNarrowScreen = () => window.innerWidth < 640;

// Web Worker wrapper — parse XLSX off main thread
function parseExcelWithWorker(buffer, sheetName) {
    return new Promise((resolve, reject) => {
        const worker = new Worker('excel-worker.js');
        worker.onmessage = (e) => { worker.terminate(); if (e.data.error) reject(new Error(e.data.error)); else resolve(e.data); };
        worker.onerror = (e) => { worker.terminate(); reject(e); };
        worker.postMessage({ buffer, sheetName }, [buffer]);
    });
}

// URL hash — encode active filters so links are shareable
function syncHashFromState() {
    if (!appState.isMophMode) return;
    const p = new URLSearchParams();
    p.set('fy', appState.fiscalYear);
    p.set('q', appState.activeQuarter);
    p.set('age', appState.activeAgeFilter);
    p.set('hct', appState.activeHctFilter);
    p.set('district', appState.activeDistrictFilter);
    p.set('hospital', appState.activeHospitalFilter);
    p.set('indicator', appState.activeMophIndicator);
    history.replaceState(null, '', '#' + p.toString());
}

function restoreStateFromHash() {
    const hash = location.hash.slice(1);
    if (!hash) return;
    const p = new URLSearchParams(hash);
    if (p.get('fy')) appState.fiscalYear = Number(p.get('fy'));
    if (p.get('q')) appState.activeQuarter = p.get('q');
    if (p.get('age')) appState.activeAgeFilter = p.get('age');
    if (p.get('hct')) appState.activeHctFilter = p.get('hct');
    if (p.get('district')) appState.activeDistrictFilter = p.get('district');
    if (p.get('hospital')) appState.activeHospitalFilter = p.get('hospital');
    const ind = p.get('indicator');
    if (ind && (ind !== 'wellchild' || hasWellchildData())) appState.activeMophIndicator = ind;
}

function syncUIControlsFromState() {
    const fyEl = document.getElementById('fiscal-year-select');
    if (fyEl) fyEl.value = String(appState.fiscalYear);
    const qEl = document.getElementById('quarter-select');
    if (qEl) qEl.value = appState.activeQuarter;
    document.querySelectorAll('[data-age-range]').forEach(b => b.classList.toggle('active', b.dataset.ageRange === appState.activeAgeFilter));
    document.querySelectorAll('[data-hct-status]').forEach(b => b.classList.toggle('active', b.dataset.hctStatus === appState.activeHctFilter));
    document.querySelectorAll('[data-indicator]').forEach(b => b.classList.toggle('active', b.dataset.indicator === appState.activeMophIndicator));
    const distEl = document.getElementById('select-district');
    if (distEl) distEl.value = appState.activeDistrictFilter;
    const hospEl = document.getElementById('select-hospital');
    if (hospEl) hospEl.value = appState.activeHospitalFilter;
}

// Must match .chart-body height in style.css; '100%' resolves against the whole card and overflows
const chartHeight = () => isNarrowScreen() ? 300 : 340;

function parseDateValue(val) {
    if (val === undefined || val === null || val === '') return null;
    if (val instanceof Date) return isNaN(val) ? null : val;
    const strVal = String(val).trim();
    // Excel serial: numeric value 10000–99999 (covers dates ~1927–2173)
    const num = parseFloat(strVal);
    if (!isNaN(num) && num > 10000 && num < 100000) {
        return new Date((num - 25569) * 86400 * 1000);
    }
    const thMatch = strVal.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (thMatch) {
        let [, d, m, y] = thMatch.map(Number);
        if (y >= 2500) y -= 543;
        return new Date(y, m - 1, d);
    }
    const isoMatch = strVal.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
        let [, y, m, d] = isoMatch.map(Number);
        if (y >= 2500) y -= 543;
        return new Date(y, m - 1, d);
    }
    return null;
}

// Thai fiscal year starts Oct 1: Oct 2026 belongs to FY2570
const fiscalYearOf = (date) => date.getFullYear() + 543 + (date.getMonth() >= 9 ? 1 : 0);
const fiscalQuarterOf = (date) => Math.floor(((date.getMonth() + 3) % 12) / 3) + 1;

function getServiceFiscalYears() {
    if (!appState.headers.includes('epi_date')) return [];
    const counts = {};
    appState.rawData.forEach(row => {
        const d = parseDateValue(row['epi_date']);
        if (d) { const fy = fiscalYearOf(d); counts[fy] = (counts[fy] || 0) + 1; }
    });
    return Object.keys(counts).map(Number).sort((a, b) => a - b).map(fy => ({ fy, count: counts[fy] }));
}

// Default the fiscal year to the latest year with substantial service data in the loaded file
function syncFiscalYearToData() {
    const years = getServiceFiscalYears();
    if (!years.length) return;
    const total = years.reduce((s, y) => s + y.count, 0);
    const main = years.filter(y => y.count >= total * 0.05);
    const latest = (main.length ? main : years)[(main.length ? main : years).length - 1].fy;
    const fySelect = document.getElementById('fiscal-year-select');
    if (fySelect && !fySelect.querySelector(`option[value="${latest}"]`)) {
        const opt = document.createElement('option');
        opt.value = String(latest);
        opt.textContent = String(latest);
        fySelect.appendChild(opt);
        [...fySelect.options].sort((a, b) => a.value - b.value).forEach(o => fySelect.appendChild(o));
    }
    appState.fiscalYear = latest;
    if (fySelect) fySelect.value = String(latest);
    syncAnemiaTabTarget();
}

function updatePeriodControls() {
    const box = document.getElementById('moph-cohort-controls');
    if (!box) return;
    box.style.display = appState.isMophMode ? 'flex' : 'none';
    const isIron = appState.activeMophIndicator === 'iron-supplement';
    document.getElementById('quarter-select-wrap').style.display = isIron ? 'flex' : 'none';
    document.getElementById('export-date-wrap').style.display = isIron ? 'none' : 'flex';
    document.getElementById('period-controls-title').textContent = isIron ? 'ช่วงเวลาให้บริการ' : 'Cohort filter';
    if (isIron) {
        const hint = document.getElementById('cohort-age-hint');
        const years = getServiceFiscalYears();
        if (hint) hint.textContent = years.length
            ? `ข้อมูลในไฟล์ (ตาม epi_date): ${years.map(y => `ปีงบ ${y.fy} ${y.count.toLocaleString()} ราย`).join(', ')}`
            : '⚠ ไม่มีคอลัมน์ epi_date จึงไม่กรองตามปีงบ';
    } else {
        updateCohortHint();
    }
}

const ANEMIA_TARGET_BY_FY = { 2568: 18.0, 2569: 17.0, 2570: 16.0, 2571: 15.0, 2572: 14.0 };
const getAnemiaTarget = (fy) => ANEMIA_TARGET_BY_FY[fy] ?? 16.0;
function syncAnemiaTabTarget() {
    const el = document.getElementById('anemia-tab-target');
    if (el) el.textContent = getAnemiaTarget(appState.fiscalYear);
}

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
    initTrendTabs();

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

    // Restore filter state from URL hash (shareable links)
    restoreStateFromHash();

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
    document.getElementById('btn-export-xlsx').style.display = 'inline-flex';
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
            adminGhConfig.style.display = GITHUB_CONFIG.token ? 'none' : 'block'; // ยังไม่มี token = ให้ใส่
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
    document.getElementById('btn-export-xlsx').style.display = 'none';
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
                updatePeriodControls();
                const disclaimer = document.getElementById('moph-anemia-disclaimer');
                if (disclaimer) disclaimer.style.display = 'none';
                // Update banner
                document.querySelector('#moph-banner .moph-alert-title p').textContent =
                    'ร้อยละของเด็กอายุ 6 เดือน – 5 ปี ได้รับยาน้ำเสริมธาตุเหล็ก (เป้าหมาย: ร้อยละ 75.0 ขึ้นไป)';
            } else if (appState.activeMophIndicator === 'wellchild') {
                if (ageFilters) ageFilters.style.display = 'none';
                if (hctFilters) hctFilters.style.display = 'none';
                updatePeriodControls();
                const disclaimer = document.getElementById('moph-anemia-disclaimer');
                if (disclaimer) {
                    disclaimer.style.display = 'block';
                    disclaimer.innerHTML = '<strong style="color:var(--text-secondary);">ℹ️ เด็กดี / WBC Screening:</strong> ใช้คอลัมน์ <code>wbc_result</code> หรือ <code>wellchild_visit</code> ในไฟล์ข้อมูล หากไม่มีคอลัมน์ดังกล่าวระบบจะแสดงจำนวนรายการทั้งหมดแทน';
                }
                document.querySelector('#moph-banner .moph-alert-title p').textContent =
                    'เด็กดี / WBC Screening — ต้องมีคอลัมน์ wbc_result หรือ wellchild_visit ในไฟล์';
            } else {
                if (ageFilters) ageFilters.style.display = 'none';
                if (hctFilters) hctFilters.style.display = 'none';
                updatePeriodControls();
                const disclaimer = document.getElementById('moph-anemia-disclaimer');
                if (disclaimer) disclaimer.style.display = 'block';
                // Update banner
                document.querySelector('#moph-banner .moph-alert-title p').textContent =
                    'ร้อยละเด็กอายุครบ 12 เดือนในเขตรับผิดชอบ มีภาวะโลหิตจาง (Coverage) — เป้าหมาย ≤17% (ปี 2569), ≤16% (ปี 2570), ≤15% (ปี 2571), ≤14% (ปี 2572)';
            }

            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    });
}

// --- Cohort Hint + Events ---
function updateCohortHint() {
    const fy = appState.fiscalYear;
    // เด็กที่ "อายุครบ 12 เดือน" ใน FY = เด็กที่เกิดในปีก่อนหน้า (1st birthday ตกใน FY)
    const fyStartCE = fy - 543 - 2;   // Oct 1 ปีก่อน FY
    const fyEndCE   = fy - 543 - 1;   // Sep 30 ก่อนเริ่ม FY
    const cohortStart = new Date(fyStartCE, 9, 1);
    const cohortEnd   = new Date(fyEndCE,   8, 30);
    const refDate = appState.exportDate || new Date();
    const monthsDiff = (d1, d2) => (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
    const minAge = Math.floor(monthsDiff(cohortEnd, refDate));
    const maxAge = Math.ceil(monthsDiff(cohortStart, refDate));
    const hint = document.getElementById('cohort-age-hint');
    const hasBirth = appState.headers && appState.headers.includes('birth');
    if (hint) hint.textContent = hasBirth
        ? `✓ เกิด 1 ต.ค.พ.ศ.${fy-2} – 30 ก.ย.พ.ศ.${fy-1}`
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
        syncAnemiaTabTarget();
        fySelect.addEventListener('change', (e) => {
            appState.fiscalYear = parseInt(e.target.value);
            appState.currentPage = 1;
            syncAnemiaTabTarget();
            updatePeriodControls();
            renderFyBadge();
            applyAllFilters();
            triggerAnalyticsUpdate();
        });
    }
    const quarterSelect = document.getElementById('quarter-select');
    if (quarterSelect) {
        quarterSelect.addEventListener('change', (e) => {
            appState.activeQuarter = e.target.value;
            appState.currentPage = 1;
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

function toggleLoader(active, text = "กำลังวิเคราะห์ข้อมูล...", skeleton = false) {
    const loader = document.getElementById('loading-overlay');
    if (loader) {
        document.querySelector('#loading-overlay .loading-text').textContent = text;
        if (active) {
            loader.classList.add('active');
            if (skeleton) loader.classList.add('skeleton-mode');
            else loader.classList.remove('skeleton-mode');
        } else {
            loader.classList.remove('active');
            loader.classList.remove('skeleton-mode');
        }
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
    reader.onload = async function (e) {
        try {
            const arrayBuffer = e.target.result;
            appState._xlsxBuffer = arrayBuffer;

            // Parse via Worker — non-blocking
            const result = await parseExcelWithWorker(arrayBuffer.slice(0), null);
            const sheetNames = result.sheetNames;
            const targetSheet = result.targetSheet;

            appState.sheetNames = sheetNames;

            document.getElementById('val-sheets').textContent = sheetNames.length;

            const sheetSelect = document.getElementById('select-sheet');
            sheetSelect.innerHTML = '';
            sheetNames.forEach(sn => {
                const opt = document.createElement('option');
                opt.value = sn; opt.textContent = sn;
                sheetSelect.appendChild(opt);
            });
            sheetSelect.value = targetSheet;

            await loadSheetData(targetSheet);

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
async function loadSheetData(sheetName) {
    toggleLoader(true, `กำลังอ่านข้อมูลแผ่นงาน [${sheetName}]...`, true);
    appState.currentSheetName = sheetName;

    let jsonData, jsonDataText;

    // Use Web Worker if buffer available (non-blocking), else fall back to main thread
    if (appState._xlsxBuffer) {
        try {
            const result = await parseExcelWithWorker(appState._xlsxBuffer.slice(0), sheetName);
            jsonData = result.rows;
            jsonDataText = result.rowsText;
        } catch (workerErr) {
            console.warn('Worker parse failed, falling back:', workerErr.message);
        }
    }

    if (!jsonData) {
        if (!appState.workbook) { toggleLoader(false); return; }
        const sheet = appState.workbook.Sheets[sheetName];
        jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        jsonDataText = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    }

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

    renderFyBadge();

    triggerAnalyticsUpdate();
    toggleLoader(false);

    // ✅ แจ้งผลสำเร็จด้วย Toast
    const rowCount = appState.rawData.length.toLocaleString();
    const modeLabel = appState.isMophMode ? '(โหมด MOPH HDC)' : '';
    showToast(`✅ นำเข้าข้อมูลสำเร็จ! พบ <strong>${rowCount} แถว</strong> จากแผ่นงาน "${sheetName}" ${modeLabel}`, 'success', 5000);
}

// ✅ แสดง fiscal year badge ใน file-info card
function renderFyBadge() {
    const fyRow = document.getElementById('row-fy-detected');
    const fyVal = document.getElementById('val-fy-detected');
    if (fyRow && fyVal && appState.headers.includes('epi_date')) {
        const fyList = getServiceFiscalYears();
        if (fyList.length > 0) {
            fyRow.style.display = 'flex';
            fyVal.innerHTML = fyList.map(({ fy, count }) => {
                const isCurrent = fy === appState.fiscalYear;
                const style = isCurrent
                    ? `background:var(--neon-cyan);color:#fff;box-shadow:0 0 10px var(--neon-cyan-glow);`
                    : `background:rgba(2,132,199,0.1);color:var(--neon-cyan);`;
                return `<span style="${style}padding:3px 10px;border-radius:20px;font-size:0.78rem;font-weight:700;letter-spacing:0.5px;margin-left:4px;">
                    ปีงบ ${fy} <span style="opacity:0.85;">(${count.toLocaleString()} ราย)</span>${isCurrent ? ' ✓' : ''}
                </span>`;
            }).join('');
        } else {
            fyRow.style.display = 'none';
        }
    }
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

const WELLCHILD_COL_RE = /wbc_result|wellchild_visit|well.child/i;
const hasWellchildData = () => appState.headers.some(h => WELLCHILD_COL_RE.test(h));

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
        syncFiscalYearToData();
        updatePeriodControls();

        const wellchildTab = document.querySelector('#moph-indicator-tabs .filter-tab-btn[data-indicator="wellchild"]');
        if (wellchildTab) wellchildTab.style.display = hasWellchildData() ? '' : 'none';

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
        updatePeriodControls();
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

    // 0. Hard filters: typearea ∈ {1,3}, discharge = 9 (กรองอัตโนมัติถ้ามีคอลัมน์)
    if (appState.headers.includes('typearea')) {
        filtered = filtered.filter(r => { const v = Number(r['typearea']); return v === 1 || v === 3; });
    }
    if (appState.headers.includes('discharge')) {
        filtered = filtered.filter(r => Number(r['discharge']) === 9);
    }

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
                // เด็กที่ 1st birthday ตกใน FY เกิดปีก่อน FY
                const fyStartCE = fy - 543 - 2;   // Oct 1 ปีก่อน FY
                const fyEndCE   = fy - 543 - 1;   // Sep 30 ก่อนเริ่ม FY
                const cohortStart = new Date(fyStartCE, 9, 1);   // Oct 1
                const cohortEnd   = new Date(fyEndCE,   8, 30);  // Sep 30
                const hasBirth = appState.headers.includes('birth');
                if (hasBirth) {
                    const parseBirth = parseDateValue;
                    // DEBUG: log sample to console
                    const sample = filtered.slice(0, 3);
                    console.log('[cohort debug] cohortStart:', cohortStart, 'cohortEnd:', cohortEnd);
                    console.log('[cohort debug] sample birth raw:', sample.map(r => r['birth']));
                    console.log('[cohort debug] sample birth parsed:', sample.map(r => parseBirth(r['birth'])));
                    // Step 1: cohort filter by birth date
                    filtered = filtered.filter(row => {
                        const birthDate = parseBirth(row['birth']);
                        if (!birthDate) return false;
                        return birthDate >= cohortStart && birthDate <= cohortEnd;
                    });
                    console.log('[cohort debug] after filter count:', filtered.length);

                    // Tag rows with lab window and anemia flags (lab_date within birth+6m to birth+13m-1d)
                    filtered.forEach(row => {
                        const birthDate = parseBirth(row['birth']);
                        row._hasLab = false;
                        row._isAnemic = false;
                        if (cleanNumericValue(row['lab_result_status']) === 1 && birthDate) {
                            const labDate = parseBirth(row['lab_date']);
                            if (labDate) {
                                const wStart = new Date(birthDate); wStart.setMonth(wStart.getMonth() + 6);
                                const wEnd   = new Date(birthDate); wEnd.setMonth(wEnd.getMonth() + 13); wEnd.setDate(wEnd.getDate() - 1);
                                if (labDate >= wStart && labDate <= wEnd) {
                                    row._hasLab = true;
                                    if (isAnemicRow(row)) row._isAnemic = true;
                                }
                            } else {
                                // No lab_date → accept lab at face value (export pre-filtered)
                                row._hasLab = true;
                                if (isAnemicRow(row)) row._isAnemic = true;
                            }
                        }
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
                if (appState.headers.includes('epi_date')) {
                    const fy = appState.fiscalYear;
                    const q = appState.activeQuarter;
                    filtered = filtered.filter(row => {
                        const d = parseDateValue(row['epi_date']);
                        if (!d || fiscalYearOf(d) !== fy) return false;
                        return q === 'all' || fiscalQuarterOf(d) === Number(q);
                    });
                }
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
    // Save pre-HCT-filter stats so KPI denominator/numerator stay correct regardless of active sub-filter
    appState.preHctFilterCount = filtered.length;
    appState.preHctTestedCount = filtered.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
    appState.preHctAnemiaCount = filtered.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;
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
                return labStatus === 1 && isAnemicRow(row);
            } else if (appState.activeHctFilter === 'normal') {
                return labStatus === 1 && !isAnemicRow(row);
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
        const searchCols = appState.isMophMode
            ? appState.headers.filter(h => ['hn','fname','lname','hosp_name','hcode','hoscode','hospcode'].includes(h.toLowerCase()))
            : null;
        filtered = filtered.filter(row => {
            const vals = searchCols && searchCols.length > 0
                ? searchCols.map(h => row[h])
                : Object.values(row);
            return vals.some(val => String(val).toLowerCase().includes(searchQuery));
        });
    }

    // 3. Advanced filter panel
    if (appState.advFilterAgeMin !== null || appState.advFilterAgeMax !== null) {
        const ageCol = appState.headers.find(h => h.toLowerCase() === 'age_m');
        if (ageCol) {
            filtered = filtered.filter(row => {
                const age = cleanNumericValue(row[ageCol]);
                if (appState.advFilterAgeMin !== null && age < appState.advFilterAgeMin) return false;
                if (appState.advFilterAgeMax !== null && age > appState.advFilterAgeMax) return false;
                return true;
            });
        }
    }
    if (appState.advFilterLabMin !== null || appState.advFilterLabMax !== null) {
        filtered = filtered.filter(row => {
            const lr = parseFloat(row['labresult']);
            if (isNaN(lr)) return appState.advFilterLabMin === null;
            if (appState.advFilterLabMin !== null && lr < appState.advFilterLabMin) return false;
            if (appState.advFilterLabMax !== null && lr > appState.advFilterLabMax) return false;
            return true;
        });
    }
    if (appState.advFilterDateFrom || appState.advFilterDateTo) {
        const dateFrom = appState.advFilterDateFrom ? new Date(appState.advFilterDateFrom) : null;
        const dateTo = appState.advFilterDateTo ? new Date(appState.advFilterDateTo) : null;
        if (dateTo) dateTo.setHours(23, 59, 59);
        filtered = filtered.filter(row => {
            const labDate = parseDateValue(row['lab_date']) || parseDateValue(row['epi_date']);
            if (!labDate) return !dateFrom && !dateTo;
            if (dateFrom && labDate < dateFrom) return false;
            if (dateTo && labDate > dateTo) return false;
            return true;
        });
    }
    if (appState.advFilterAnemia !== 'all') {
        filtered = filtered.filter(row => {
            const anemic = isAnemicRow(row);
            return appState.advFilterAnemia === 'yes' ? anemic : !anemic;
        });
    }
    // 4. Risk filter (MOPH mode only)
    if (appState.isMophMode && appState.activeRiskFilter !== 'all') {
        filtered = filtered.filter(row => getRiskLevel(row) === appState.activeRiskFilter);
    }

    appState.filteredData = filtered;
    syncHashFromState();
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

let _searchDebounceTimer = null;
function handleTableSearch() {
    clearTimeout(_searchDebounceTimer);
    _searchDebounceTimer = setTimeout(() => {
        applyAllFilters();
        appState.currentPage = 1;
        triggerAnalyticsUpdate();
        const countEl = document.getElementById('search-match-count');
        if (countEl) {
            const q = document.getElementById('table-search').value.trim();
            countEl.textContent = q ? `พบ ${appState.filteredData.length.toLocaleString()} รายการ` : '';
        }
    }, 300);
}

// ==========================================================================
// 📐 Aggregation & Metric Calculation
// ==========================================================================

// คำนวณภาวะโลหิตจางจาก labresult โดยตรง (anemea column ใน export มักไม่ถูก populate)
function isAnemicRow(r) {
    const lt = String(r['labtest'] || '').trim().replace(/,/g, '').replace(/^0+/, '');
    const lr = parseFloat(r['labresult']);
    if (!isNaN(lr) && lr > 0) {
        if (lt === '621201') return lr < 33;                     // HCT% < 33
        if (lt === '621401' || lt === '621402') return lr < 11;  // Hb g/dL < 11
    }
    // anemea: 2 = ซีด (โลหิตจาง), 0 = ปกติ — per table display logic
    const anemea = cleanNumericValue(r['anemea']);
    return anemea === 2 || anemea === 1;
}

function getRiskLevel(row) {
    const lt = String(row['labtest'] || '').trim().replace(/,/g, '').replace(/^0+/, '');
    const lr = parseFloat(row['labresult']);
    if (!isNaN(lr) && lr > 0) {
        if (lt === '621201') { if (lr < 28) return 'high'; if (lr < 33) return 'medium'; return 'low'; }
        if (lt === '621401' || lt === '621402') { if (lr < 9) return 'high'; if (lr < 11) return 'medium'; return 'low'; }
    }
    const anemea = cleanNumericValue(row['anemea']);
    return (anemea === 2 || anemea === 1) ? 'medium' : 'low';
}

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
            const testedRows = rows.filter(r => r._hasLab);
            const totalTested = testedRows.length;
            const totalAnemia = rows.filter(r => r._isAnemic).length;
            const totalNormal = totalTested - totalAnemia;

            const screeningRate = totalChildren > 0 ? (totalTested / totalChildren * 100) : 0;
            const anemiaRate = totalTested > 0 ? (totalAnemia / totalTested * 100) : 0;

            // 1. KPI: C = จำนวนเด็กอายุครบ 12 เดือน ในเขตรับผิดชอบทั้งหมด
            document.getElementById('kpi-1-title').textContent = "เด็กอายุครบ 12 เดือน ในเขตรับผิดชอบ [C]";
            document.getElementById('kpi-total-rows').textContent = totalChildren.toLocaleString();
            document.getElementById('kpi-1-subtitle').innerHTML = `<i data-lucide="baby"></i> ทั้งหมดในเขตรับผิดชอบ`;

            // 2. KPI: B = ได้รับการตรวจภาวะโลหิตจาง ช่วงอายุ 6-12 เดือน (ตัวส่วน)
            document.getElementById('kpi-sum-title').textContent = "ได้รับการตรวจภาวะโลหิตจาง [B]";
            document.getElementById('kpi-total-sum').textContent = totalTested.toLocaleString();
            document.getElementById('kpi-sum-subtitle').innerHTML = `<i data-lucide="activity"></i> อัตราตรวจคัดกรอง ${screeningRate.toFixed(1)}%`;

            // 3. KPI: A/B×100 = ร้อยละโลหิตจาง (ตัวชี้วัดหลัก)
            document.getElementById('kpi-avg-title').textContent = "ร้อยละโลหิตจาง [A/B×100]";
            document.getElementById('kpi-total-avg').textContent = totalTested > 0 ? anemiaRate.toFixed(1) + "%" : "–";

            const anemiaTarget = getAnemiaTarget(appState.fiscalYear);
            const targetBadge = document.getElementById('moph-target-badge');
            if (totalTested === 0) {
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="info"></i> ไม่มีข้อมูลการตรวจในปีงบ ${appState.fiscalYear}`;
                targetBadge.className = "target-badge";
                targetBadge.innerHTML = `เป้าหมาย ≤${anemiaTarget}% | ไม่มีข้อมูล`;
            } else if (anemiaRate <= anemiaTarget) {
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

            // 4. KPI: A = พบภาวะโลหิตจาง (ตัวเศษ)
            document.getElementById('kpi-4-title').textContent = "พบภาวะโลหิตจาง [A]";
            document.getElementById('kpi-unique-categories').textContent = totalAnemia.toLocaleString();
            document.getElementById('kpi-4-subtitle').innerHTML = `<i data-lucide="heart-pulse"></i> ปกติ (Normal) ${totalNormal} ราย`;

        } else if (appState.activeMophIndicator === 'wellchild') {
            // --- Well-child / WBC Screening indicator ---
            const wbcCol = appState.headers.find(h => WELLCHILD_COL_RE.test(h));
            const visitCol = appState.headers.find(h => /visit|attended/i.test(h));
            const total = rows.length;

            const attended = wbcCol
                ? rows.filter(r => { const v = String(r[wbcCol] || '').toLowerCase(); return v === '1' || v === 'yes' || v === 'ผ่าน' || v === 'done'; }).length
                : (visitCol ? rows.filter(r => cleanNumericValue(r[visitCol]) >= 1).length : 0);
            const attendRate = total > 0 ? (attended / total * 100).toFixed(1) : '–';

            document.getElementById('kpi-1-title').textContent = 'เด็กกลุ่มเป้าหมายทั้งหมด';
            document.getElementById('kpi-total-rows').textContent = total.toLocaleString();
            document.getElementById('kpi-1-subtitle').innerHTML = '<i data-lucide="baby"></i> เด็กดี / WBC Screening';
            document.getElementById('kpi-sum-title').textContent = 'มารับบริการ / ผ่านเกณฑ์';
            document.getElementById('kpi-total-sum').textContent = attended.toLocaleString();
            document.getElementById('kpi-sum-subtitle').innerHTML = wbcCol ? `<i data-lucide="check-circle-2"></i> จากคอลัมน์ ${wbcCol}` : '<i data-lucide="info"></i> ไม่พบคอลัมน์ wbc_result';
            document.getElementById('kpi-avg-title').textContent = 'ร้อยละครอบคลุม (Coverage%)';
            document.getElementById('kpi-total-avg').textContent = attendRate + (attendRate !== '–' ? '%' : '');
            document.getElementById('kpi-avg-subtitle').className = `kpi-trend ${wbcCol ? 'positive' : 'neutral'}`;
            document.getElementById('kpi-avg-subtitle').innerHTML = wbcCol ? '<i data-lucide="trending-up"></i> คำนวณจากข้อมูลจริง' : '<i data-lucide="alert-triangle"></i> ต้องการคอลัมน์ wbc_result';
            document.getElementById('kpi-4-title').textContent = 'จำนวนคอลัมน์ในข้อมูล';
            document.getElementById('kpi-unique-categories').textContent = appState.headers.length;
            document.getElementById('kpi-4-subtitle').innerHTML = `<i data-lucide="table-2"></i> ${appState.headers.slice(0,3).join(', ')}...`;
            const targetBadge = document.getElementById('moph-target-badge');
            if (targetBadge) {
                targetBadge.className = wbcCol ? 'target-badge met' : 'target-badge';
                targetBadge.textContent = wbcCol ? `คอลัมน์: ${wbcCol}` : 'ต้องการ wbc_result';
            }

        } else {
            // --- MOPH Specialized HDC Analytics: ยาเสริมธาตุเหล็ก ---
            const totalTarget = rows.length;

            // Denominator logic: received = column 'result' is dispensing
            const receivedRows = rows.filter(r => isReceivedValue(r['result']));
            const totalReceived = receivedRows.length;

            const coverageRate = totalTarget > 0 ? (totalReceived / totalTarget * 100) : 0;

            // Anemia and HCT screening rates — always use pre-HCT-filter counts so sub-filter doesn't distort rates
            const totalTested = appState.preHctTestedCount ?? rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
            const totalAnemia = appState.preHctAnemiaCount ?? rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;

            const hctBase = appState.preHctFilterCount > 0 ? appState.preHctFilterCount : totalTarget;
            const testedRate = hctBase > 0 ? (totalTested / hctBase * 100) : 0;
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
            document.getElementById('kpi-total-avg').textContent = totalTarget > 0 ? coverageRate.toFixed(1) + "%" : "–";

            // MOPH Target Threshold = 75.0%
            const targetBadge = document.getElementById('moph-target-badge');
            if (totalTarget === 0) {
                const quarterText = appState.activeQuarter === 'all' ? '' : ` ไตรมาส ${appState.activeQuarter}`;
                document.getElementById('kpi-avg-subtitle').className = "kpi-trend";
                document.getElementById('kpi-avg-subtitle').innerHTML = `<i data-lucide="info"></i> ไม่มีข้อมูลบริการในปีงบ ${appState.fiscalYear}${quarterText}`;
                targetBadge.className = "target-badge";
                targetBadge.innerHTML = `เป้าหมาย 75% | ไม่มีข้อมูล`;
            } else if (coverageRate >= 75.0) {
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

// ==========================================================================
// 🚨 Coverage Alert Banner — list hospitals below 75% target
// ==========================================================================
function renderCoverageAlerts() {
    const panel = document.getElementById('coverage-alert-panel');
    const chips = document.getElementById('coverage-alert-chips');
    if (!panel || !chips || !appState.isMophMode || appState.activeMophIndicator !== 'iron-supplement') {
        if (panel) panel.classList.remove('has-alerts');
        return;
    }

    const { hosnameCol } = trendGetCols();
    if (!hosnameCol) { panel.classList.remove('has-alerts'); return; }

    const rows = appState.filteredData;
    const hospMap = {};
    rows.forEach(r => {
        const h = String(r[hosnameCol] || '');
        if (!h) return;
        if (!hospMap[h]) hospMap[h] = { total: 0, received: 0 };
        hospMap[h].total++;
        if (isReceivedValue(r['result'])) hospMap[h].received++;
    });

    const belowTarget = Object.entries(hospMap)
        .filter(([name, b]) => b.total >= 5 && (b.received / b.total * 100) < getHospitalTarget(name))
        .sort((a, b) => (a[1].received / a[1].total) - (b[1].received / b[1].total));

    chips.innerHTML = '';
    if (belowTarget.length === 0) { panel.classList.remove('has-alerts'); return; }

    belowTarget.forEach(([name, b]) => {
        const pct = (b.received / b.total * 100).toFixed(1);
        const chip = document.createElement('span');
        chip.className = 'coverage-alert-chip';
        chip.textContent = `${name} (${pct}%${getHospitalTarget(name) !== 75 ? ` / เป้า ${getHospitalTarget(name)}%` : ''})`;
        chip.title = `คลิกเพื่อกรอง: ${name}`;
        chip.addEventListener('click', () => {
            appState.activeHospitalFilter = name;
            const hospEl = document.getElementById('select-hospital');
            if (hospEl) hospEl.value = name;
            applyAllFilters();
        });
        chips.appendChild(chip);
    });
    panel.classList.add('has-alerts');
    lucide.createIcons();
}

// ==========================================================================
// 🏆 Composite Performance Score
// ==========================================================================
function renderCompositeScore() {
    const card = document.getElementById('composite-score-card');
    const gradeEl = document.getElementById('composite-grade');
    const labelEl = document.getElementById('composite-label');
    const breakdownEl = document.getElementById('composite-breakdown');
    if (!card || !appState.isMophMode || appState.activeMophIndicator !== 'iron-supplement') {
        if (card) card.classList.remove('visible');
        return;
    }

    const rows = appState.filteredData;
    if (rows.length === 0) { card.classList.remove('visible'); return; }

    const totalReceived = rows.filter(r => isReceivedValue(r['result'])).length;
    const totalTested = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
    const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;

    const coverage = rows.length > 0 ? totalReceived / rows.length * 100 : 0;
    const hctRate = rows.length > 0 ? totalTested / rows.length * 100 : 0;
    const anemiaRate = totalTested > 0 ? totalAnemia / totalTested * 100 : 0;
    const anemiaTarget = ANEMIA_TARGET_BY_FY[appState.fiscalYear] || 16;

    // Score: coverage (weight 50%), HCT rate (weight 25%), anemia (weight 25% inverted)
    const covScore = Math.min(coverage / 75 * 50, 50);
    const hctScore = Math.min(hctRate / 50 * 25, 25);
    const anemiaScore = anemiaRate <= anemiaTarget ? 25 : Math.max(0, 25 - (anemiaRate - anemiaTarget) * 2);
    const total = Math.round(covScore + hctScore + anemiaScore);

    let grade, gradeClass;
    if (total >= 90) { grade = 'A'; gradeClass = 'grade-a'; }
    else if (total >= 75) { grade = 'B'; gradeClass = 'grade-b'; }
    else if (total >= 55) { grade = 'C'; gradeClass = 'grade-c'; }
    else { grade = 'D'; gradeClass = 'grade-d'; }

    gradeEl.textContent = grade;
    gradeEl.className = `composite-score-grade ${gradeClass}`;
    labelEl.textContent = `คะแนนภาพรวม ${total}/100`;
    breakdownEl.textContent = `Coverage ${coverage.toFixed(1)}% · HCT ${hctRate.toFixed(1)}% · โลหิตจาง ${anemiaRate.toFixed(1)}%`;
    card.classList.add('visible');
}

// ==========================================================================
// 📱 Mobile Bottom Sheet Filter
// ==========================================================================
function openFilterSheet() {
    document.getElementById('bottom-sheet-overlay').classList.add('open');
    document.getElementById('bottom-sheet').classList.add('open');
    syncBottomSheetSelects();
}

function closeFilterSheet() {
    document.getElementById('bottom-sheet-overlay').classList.remove('open');
    document.getElementById('bottom-sheet').classList.remove('open');
}

function syncBottomSheetSelects() {
    // Sync FY options
    const fyMain = document.getElementById('fiscal-year-select');
    const fyBS = document.getElementById('bs-fy-select');
    if (fyMain && fyBS) {
        fyBS.innerHTML = fyMain.innerHTML;
        fyBS.value = fyMain.value;
    }
    // Sync district
    const distMain = document.getElementById('select-district');
    const distBS = document.getElementById('bs-dist-select');
    if (distMain && distBS) {
        distBS.innerHTML = distMain.innerHTML;
        distBS.value = distMain.value;
    }
    // Sync hospital
    const hospMain = document.getElementById('select-hospital');
    const hospBS = document.getElementById('bs-hosp-select');
    if (hospMain && hospBS) {
        hospBS.innerHTML = hospMain.innerHTML;
        hospBS.value = hospMain.value;
    }
    // Sync quarter
    const qBS = document.getElementById('bs-q-select');
    const qMain = document.getElementById('quarter-select');
    if (qBS && qMain) qBS.value = qMain.value;
}

function triggerAnalyticsUpdate() {
    renderKPIs();
    renderCoverageAlerts();
    renderCompositeScore();
    renderCharts();
    renderTable();
    syncBottomSheetSelects();
}

// ==========================================================================
// 📊 Rendering Visualizations (ApexCharts with customized health metrics)
// ==========================================================================

function renderCharts() {
    const rows = appState.filteredData;
    if (rows.length === 0) {
        // Clear stale charts so the view matches the (empty) filter result
        Object.keys(charts).forEach(k => {
            if (charts[k]) { charts[k].destroy(); charts[k] = null; }
        });
        return;
    }

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
            height: chartHeight(),
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
                barHeight: '75%'
            }
        },
        dataLabels: { enabled: false },
        yaxis: { labels: { maxWidth: isNarrowScreen() ? 110 : 220, style: { fontSize: isNarrowScreen() ? '9px' : '11px' } } },
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
    const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;
    const totalNormal = totalTested - totalAnemia;

    const donutOptions = {
        series: [totalNormal, totalAnemia, totalNotTested],
        chart: {
            type: 'donut',
            height: chartHeight(),
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
        legend: { position: 'bottom', fontSize: '9px', markers: { size: 8, shape: 'circle', strokeWidth: 0, offsetY: 1 } },
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
            height: chartHeight(),
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
        dataLabels: { enabled: false },
        markers: { size: 3, strokeWidth: 0, hover: { size: 6 } },
        xaxis: {
            categories: ageCategories,
            tickAmount: isNarrowScreen() ? 6 : 12,
            labels: { rotate: 0, hideOverlappingLabels: true },
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
    const receivedAnemiaCount = receivedTestedRows.filter(r => isAnemicRow(r)).length;
    const anemiaRateReceived = receivedTestedRows.length > 0 ? (receivedAnemiaCount / receivedTestedRows.length * 100) : 0;

    const notReceivedTestedRows = rows.filter(r => !isReceivedValue(r['result']) && cleanNumericValue(r['lab_result_status']) === 1);
    const notReceivedAnemiaCount = notReceivedTestedRows.filter(r => isAnemicRow(r)).length;
    const anemiaRateNotReceived = notReceivedTestedRows.length > 0 ? (notReceivedAnemiaCount / notReceivedTestedRows.length * 100) : 0;

    const radarOptions = {
        series: [{
            name: 'ร้อยละความชุกโลหิตจาง (Anemia Rate จากคนที่เจาะ Lab)',
            data: [parseFloat(anemiaRateReceived.toFixed(2)), parseFloat(anemiaRateNotReceived.toFixed(2))]
        }],
        chart: {
            type: 'bar',
            height: chartHeight(),
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
            categories: isNarrowScreen() ? ['ได้รับยา', 'ไม่ได้รับยา'] : ['กลุ่มที่ได้รับธาตุเหล็ก', 'กลุ่มที่ไม่ได้รับธาตุเหล็ก'],
            labels: { style: { fontSize: '11px', fontWeight: 'bold' }, trim: false, rotate: 0 }
        },
        grid: { borderColor: 'rgba(15,23,42,0.08)', padding: { bottom: 16 } },
        yaxis: {
            labels: { formatter: function (val) { return val + "%"; } }
        },
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

    renderTrendSection(rows);
    renderScatterPlot(rows);
}

// --- MOPH Indicator: ร้อยละเด็กอายุครบ 12 เดือน มีภาวะโลหิตจาง ---
function renderAnemia12mCharts(rows) {
    document.getElementById('chart-bar-title').innerHTML = `<i data-lucide="bar-chart-4"></i> ร้อยละภาวะโลหิตจางรายหน่วยบริการ (เป้าหมาย ≤17%)`;
    document.getElementById('chart-donut-title').innerHTML = `<i data-lucide="pie-chart"></i> สัดส่วนภาวะโลหิตจาง (เด็กอายุ 6-12 เดือน)`;
    document.getElementById('chart-area-title').innerHTML = `<i data-lucide="trending-up"></i> จำนวนเด็กอายุ 6-12 เดือน รายหน่วยบริการ (ตัวส่วน vs ตัวเศษ)`;
    document.getElementById('chart-radar-title').innerHTML = `<i data-lucide="activity"></i> เปรียบเทียบอัตราโลหิตจางรายอำเภอ`;
    lucide.createIcons();

    // Empty state: ไม่มีข้อมูล lab HCT/Hb ในช่วงอายุ 6-12 เดือน
    const hasLabData = rows.some(r => r._hasLab);
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
        if (r._hasLab) {
            hospMap[hosp].tested++;
            if (r._isAnemic) hospMap[hosp].anemia++;
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
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#475569', toolbar: { show: true }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { horizontal: true, borderRadius: 3, barHeight: '65%' } },
        yaxis: { labels: { maxWidth: isNarrowScreen() ? 110 : 220, style: { fontSize: isNarrowScreen() ? '9px' : '11px' } } },
        xaxis: {
            categories: hospLabels,
            labels: { formatter: val => val.toLocaleString() + ' ราย' }
        },
        legend: { position: 'top', markers: { size: 8, shape: 'circle', strokeWidth: 0, offsetY: 1 } },
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
    const totalTested = rows.filter(r => r._hasLab).length;
    const totalAnemia = rows.filter(r => r._isAnemic).length;
    const totalNormal = totalTested - totalAnemia;
    const totalUntested = rows.length - totalTested;
    const anemiaRateOverall = totalTested > 0 ? (totalAnemia / totalTested * 100) : 0;

    const donutOptions = {
        series: [totalNormal, totalAnemia, totalUntested],
        chart: { type: 'donut', height: chartHeight(), background: 'transparent', foreColor: '#475569' },
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
        legend: { position: 'bottom', fontSize: '10px', markers: { size: 8, shape: 'circle', strokeWidth: 0, offsetY: 1 } },
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
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#475569', toolbar: { show: true }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { horizontal: false, columnWidth: '65%', borderRadius: 3 } },
        xaxis: {
            categories: hospAllSorted.map(h => h.replace('โรงพยาบาลส่งเสริมสุขภาพตำบล', 'รพ.สต.').replace('โรงพยาบาล', 'รพ.')),
            tickAmount: hospAllSorted.length,
            labels: { rotate: -45, style: { fontSize: '9px' } }
        },
        yaxis: { title: { text: 'จำนวน (ราย)' } },
        legend: { position: 'top', markers: { size: 8, shape: 'circle', strokeWidth: 0, offsetY: 1 } },
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
        if (r._hasLab) {
            ampMap[amp].tested++;
            if (r._isAnemic) ampMap[amp].anemia++;
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
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#475569', toolbar: { show: false }, stacked: true },
        theme: { mode: 'light' },
        colors: ['#16a34a', '#dc2626'],
        plotOptions: { bar: { columnWidth: '55%', borderRadius: 4 } },
        xaxis: { categories: amps, labels: { style: { fontSize: '11px', fontWeight: 'bold' } } },
        yaxis: { title: { text: 'จำนวน (ราย)' } },
        legend: { position: 'top', markers: { size: 8, shape: 'circle', strokeWidth: 0, offsetY: 1 } },
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

    renderTrendSection(rows);
    renderScatterPlot(rows);
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
            height: chartHeight(),
            background: 'transparent',
            foreColor: '#475569',
            toolbar: { show: true },
            dropShadow: { enabled: true, top: 5, left: 0, blur: 8, color: neonColors, opacity: 0.25 }
        },
        theme: { mode: 'light' },
        colors: neonColors,
        fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.45, opacityTo: 0.05, stops: [0, 90, 100] } },
        stroke: { curve: 'smooth', width: 3 },
        dataLabels: { enabled: false },
        markers: { size: 3, strokeWidth: 0, hover: { size: 6 } },
        xaxis: { categories: chartCategories, tickAmount: Math.min(chartCategories.length, 20), labels: { rotate: -45, hideOverlappingLabels: true, style: { fontSize: '10px' } } },
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
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#475569', toolbar: { show: false } },
        theme: { mode: 'light' },
        colors: neonColors,
        plotOptions: { bar: { horizontal: false, columnWidth: '55%', borderRadius: 4 } },
        stroke: { show: true, width: 2, colors: ['transparent'] },
        dataLabels: { enabled: false },
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
        chart: { type: 'donut', height: chartHeight(), background: 'transparent', foreColor: '#475569' },
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
        chart: { type: 'radar', height: chartHeight(), background: 'transparent', foreColor: '#475569', toolbar: { show: false } },
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

    // Hide trend section and scatter in generic mode
    const trendSec = document.getElementById('trend-section');
    if (trendSec) trendSec.style.display = 'none';
    const scatterCard = document.getElementById('chart-scatter-card');
    if (scatterCard) scatterCard.style.display = 'none';
    if (charts.scatter) { charts.scatter.destroy(); charts.scatter = null; }
}

// ==========================================================================
// 📈 Trend Analysis Section
// ==========================================================================

function initTrendTabs() {
    const tabs = document.querySelectorAll('#trend-section .filter-tab-btn[data-trend]');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabs.forEach(t => t.classList.remove('active'));
            e.currentTarget.classList.add('active');
            appState.activeTrendTab = e.currentTarget.dataset.trend;
            renderTrendSection(appState.filteredData);
        });
    });
}

function renderTrendSection(rows) {
    const sec = document.getElementById('trend-section');
    if (!sec) return;
    if (!appState.isMophMode) { sec.style.display = 'none'; return; }
    sec.style.display = 'flex';
    lucide.createIcons();

    const tab = appState.activeTrendTab || 'monthly';
    if (tab === 'monthly') renderTrendMonthly(rows);
    else if (tab === 'quarterly') renderTrendQuarterly(rows);
    else if (tab === 'district') renderTrendDistrict(rows);
    else if (tab === 'unit') renderTrendUnit(rows);
    else if (tab === 'fiscal') renderTrendFiscal(rows);
    else if (tab === 'heatmap') renderTrendHeatmap(rows);
    else if (tab === 'ranking') renderTrendRanking(rows);
    else if (tab === 'multify') renderTrendMultiFY(rows);
    else if (tab === 'scatter') renderTrendScatter(rows);
    else if (tab === 'map') renderTrendMap(rows);
}

const THAI_MONTHS_SHORT = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
// FY month order: Oct(9),Nov(10),Dec(11),Jan(0)..Sep(8)
const FY_MONTH_ORDER = [9,10,11,0,1,2,3,4,5,6,7,8];
const FY_MONTH_LABELS = ['ต.ค.','พ.ย.','ธ.ค.','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.'];
const TREND_COLORS = ['#0ea5e9','#8b5cf6','#f59e0b','#10b981','#f43f5e','#64748b'];

function trendChartTheme() {
    return {
        background: 'transparent',
        foreColor: '#64748b',
    };
}

function destroyTrendCharts() {
    if (charts.trendMain) { charts.trendMain.destroy(); charts.trendMain = null; }
    if (charts.trendSide) { charts.trendSide.destroy(); charts.trendSide = null; }
}

function renderTrendMain(el, options) {
    if (charts.trendMain) { charts.trendMain.destroy(); charts.trendMain = null; }
    // Enable download toolbar unless caller already set one
    if (options.chart && !options.chart.toolbar) {
        options.chart.toolbar = { show: true, tools: { download: true, selection: false, zoom: false, zoomin: false, zoomout: false, pan: false, reset: false } };
    }
    charts.trendMain = new ApexCharts(el, options);
    charts.trendMain.render();
}

function renderTrendSide(el, options) {
    if (charts.trendSide) { charts.trendSide.destroy(); charts.trendSide = null; }
    charts.trendSide = new ApexCharts(el, options);
    charts.trendSide.render();
}

function trendGetCols() {
    const ampnameCol = appState.headers.find(h => ['ampname','amphurname','amp_name','district'].includes(h.toLowerCase())) || null;
    const hosnameCol = appState.headers.find(h => ['hosname','hname','hospital'].includes(h.toLowerCase())) || null;
    return { ampnameCol, hosnameCol };
}

function renderTrendMonthly(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="line-chart"></i> แนวโน้มรายเดือน (ในปีงบที่เลือก)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-2"></i> สรุปรายไตรมาส';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    // Build month buckets for FY
    const buckets = {}; // key = monthIndex (0-11)
    FY_MONTH_ORDER.forEach(m => { buckets[m] = { total: 0, received: 0, tested: 0, anemic: 0 }; });

    rows.forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d) return;
        const m = d.getMonth();
        if (!buckets[m]) return;
        buckets[m].total++;
        if (isReceivedValue(r['result'])) buckets[m].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) {
            buckets[m].tested++;
            if (isAnemicRow(r)) buckets[m].anemic++;
        }
    });

    // Filter to months with data
    const labels = [], covSeries = [], screenSeries = [], anemiaSeries = [];
    FY_MONTH_ORDER.forEach((m, i) => {
        const b = buckets[m];
        if (b.total === 0) return;
        labels.push(FY_MONTH_LABELS[i]);
        covSeries.push(+(b.total > 0 ? b.received / b.total * 100 : 0).toFixed(1));
        screenSeries.push(+(b.total > 0 ? b.tested / b.total * 100 : 0).toFixed(1));
        anemiaSeries.push(+(b.tested > 0 ? b.anemic / b.tested * 100 : 0).toFixed(1));
    });

    // Prev fiscal year overlay — use all rawData (unfiltered by FY)
    const prevFY = appState.fiscalYear - 1;
    const prevBuckets = {};
    FY_MONTH_ORDER.forEach(m => { prevBuckets[m] = { total:0, received:0, tested:0, anemic:0 }; });
    (appState.rawData || []).forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d || fiscalYearOf(d) !== prevFY) return;
        const m = d.getMonth();
        if (!prevBuckets[m]) return;
        prevBuckets[m].total++;
        if (isReceivedValue(r['result'])) prevBuckets[m].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { prevBuckets[m].tested++; if (isAnemicRow(r)) prevBuckets[m].anemic++; }
    });
    const prevCovSeries = [], prevAnemiaSeries = [];
    FY_MONTH_ORDER.forEach((m, i) => {
        if (!labels.includes(FY_MONTH_LABELS[i])) return; // only months with current-year data
        const b = prevBuckets[m];
        prevCovSeries.push(+(b.total > 0 ? b.received / b.total * 100 : 0).toFixed(1));
        prevAnemiaSeries.push(+(b.tested > 0 ? b.anemic / b.tested * 100 : 0).toFixed(1));
    });
    const hasPrevData = isAnemia ? prevAnemiaSeries.some(v => v > 0) : prevCovSeries.some(v => v > 0);

    const series = isAnemia
        ? [{ name: 'ร้อยละโลหิตจาง', data: anemiaSeries }, ...(hasPrevData ? [{ name: `ปีงบ ${prevFY} (เปรียบเทียบ)`, data: prevAnemiaSeries }] : [])]
        : [{ name: 'ความครอบคลุม (Coverage%)', data: covSeries }, { name: 'อัตราเจาะ HCT%', data: screenSeries }, ...(hasPrevData ? [{ name: `ปีงบ ${prevFY} (Coverage%)`, data: prevCovSeries }] : [])];
    const colors = isAnemia ? ['#f43f5e', '#94a3b8'] : [TREND_COLORS[0], TREND_COLORS[2], '#94a3b8'];
    const widths = series.map((_, i) => i === series.length - 1 && hasPrevData ? 1.5 : 2);
    const dashes = series.map((_, i) => i === series.length - 1 && hasPrevData ? 6 : 0);

    // Linear forecast: project primary series 2 months past last data point
    const primary = isAnemia ? anemiaSeries : covSeries;
    const FORECAST_STEPS = 2;
    const lastLabelIdx = labels.length ? FY_MONTH_LABELS.indexOf(labels[labels.length - 1]) : -1;
    const steps = Math.min(FORECAST_STEPS, 11 - lastLabelIdx);
    if (primary.length >= 3 && steps > 0) {
        const fit = linearRegression(primary);
        const futureLabels = FY_MONTH_LABELS.slice(lastLabelIdx + 1, lastLabelIdx + 1 + steps);
        series.forEach(s => { s.data = [...s.data, ...futureLabels.map(() => null)]; });
        const forecast = [...primary.map(() => null), ...futureLabels.map(() => null)];
        forecast[primary.length - 1] = primary[primary.length - 1];
        futureLabels.forEach((_, k) => {
            const v = fit.slope * (primary.length + k) + fit.intercept;
            forecast[primary.length + k] = +Math.max(0, Math.min(100, v)).toFixed(1);
        });
        series.push({ name: `คาดการณ์ (Linear, ${fit.slope >= 0 ? '+' : ''}${fit.slope.toFixed(1)}%/เดือน)`, data: forecast });
        labels.push(...futureLabels.map(l => l + '*'));
        colors.splice(series.length - 1, 0, '#a855f7');
        widths.push(2);
        dashes.push(4);
    }

    const mainOpts = {
        chart: { type: 'line', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series,
        xaxis: { categories: labels, labels: { style: { fontSize: '11px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors,
        stroke: { curve: 'smooth', width: widths, dashArray: dashes },
        markers: { size: 4 },
        dataLabels: { enabled: true, formatter: v => v === null || v === undefined ? '' : v + '%', style: { fontSize: '10px' } },
        tooltip: { y: { formatter: v => v === null || v === undefined ? '-' : v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' },
        legend: { position: 'top' }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), mainOpts);

    // Quarterly summary side chart
    const qBuckets = { 1: { total:0,received:0,tested:0,anemic:0 }, 2:{total:0,received:0,tested:0,anemic:0}, 3:{total:0,received:0,tested:0,anemic:0}, 4:{total:0,received:0,tested:0,anemic:0} };
    rows.forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d) return;
        const q = fiscalQuarterOf(d);
        if (!qBuckets[q]) return;
        qBuckets[q].total++;
        if (isReceivedValue(r['result'])) qBuckets[q].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { qBuckets[q].tested++; if (isAnemicRow(r)) qBuckets[q].anemic++; }
    });
    const qLabels = ['Q1 (ต.ค.-ธ.ค.)','Q2 (ม.ค.-มี.ค.)','Q3 (เม.ย.-มิ.ย.)','Q4 (ก.ค.-ก.ย.)'];
    const qVals = [1,2,3,4].map(q => {
        const b = qBuckets[q];
        if (isAnemia) return +(b.tested > 0 ? b.anemic/b.tested*100 : 0).toFixed(1);
        return +(b.total > 0 ? b.received/b.total*100 : 0).toFixed(1);
    });
    const sideOpts = {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: isAnemia ? 'ร้อยละโลหิตจาง' : 'Coverage%', data: qVals }],
        xaxis: { categories: qLabels, labels: { style: { fontSize: '10px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: isAnemia ? ['#f43f5e'] : [TREND_COLORS[0]],
        dataLabels: { enabled: true, formatter: v => v + '%', style: { fontSize: '11px' } },
        plotOptions: { bar: { borderRadius: 4 } },
        tooltip: { y: { formatter: v => v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    };
    renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
}

function renderTrendQuarterly(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="bar-chart-4"></i> เปรียบเทียบรายไตรมาส (ปีงบปัจจุบัน vs ก่อนหน้า)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="pie-chart"></i> สัดส่วนเคส รายไตรมาส';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const qLabels = ['Q1 (ต.ค.-ธ.ค.)','Q2 (ม.ค.-มี.ค.)','Q3 (เม.ย.-มิ.ย.)','Q4 (ก.ค.-ก.ย.)'];

    // Build per-FY per-Q data from rawData (all years)
    const fyQMap = {}; // fyQMap[fy][q] = {total,received,tested,anemic}
    (appState.rawData || []).forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d) return;
        const fy = fiscalYearOf(d);
        const q = fiscalQuarterOf(d);
        if (!fyQMap[fy]) fyQMap[fy] = {};
        if (!fyQMap[fy][q]) fyQMap[fy][q] = { total:0,received:0,tested:0,anemic:0 };
        fyQMap[fy][q].total++;
        if (isReceivedValue(r['result'])) fyQMap[fy][q].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { fyQMap[fy][q].tested++; if (isAnemicRow(r)) fyQMap[fy][q].anemic++; }
    });

    const sortedFYs = Object.keys(fyQMap).map(Number).sort();
    const seriesList = sortedFYs.map(fy => ({
        name: `ปีงบ ${fy}`,
        data: [1,2,3,4].map(q => {
            const b = fyQMap[fy][q] || { total:0,received:0,tested:0,anemic:0 };
            if (isAnemia) return +(b.tested > 0 ? b.anemic/b.tested*100 : 0).toFixed(1);
            return +(b.total > 0 ? b.received/b.total*100 : 0).toFixed(1);
        })
    }));

    const mainOpts = {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: seriesList,
        xaxis: { categories: qLabels },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: TREND_COLORS,
        dataLabels: { enabled: seriesList.length <= 3, formatter: v => v + '%', style: { fontSize: '10px' } },
        plotOptions: { bar: { borderRadius: 3, columnWidth: seriesList.length > 2 ? '70%' : '55%' } },
        tooltip: { y: { formatter: v => v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' },
        legend: { position: 'top' }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), mainOpts);

    // Side: donut of Q record counts (current FY filtered rows)
    const qCounts = [1,2,3,4].map(q => rows.filter(r => { const d = parseDateValue(r['epi_date']); return d && fiscalQuarterOf(d) === q; }).length);
    const sideOpts = {
        chart: { type: 'donut', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: qCounts,
        labels: qLabels,
        colors: TREND_COLORS,
        dataLabels: { enabled: true, formatter: (v, { seriesIndex, w }) => `${w.globals.labels[seriesIndex]}: ${w.globals.series[seriesIndex].toLocaleString()}` },
        legend: { position: 'bottom', fontSize: '10px' },
        tooltip: { y: { formatter: v => v.toLocaleString() + ' ราย' } }
    };
    renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
}

function renderTrendDistrict(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="map-pin"></i> ร้อยละครอบคลุมรายอำเภอ';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="pie-chart"></i> สัดส่วนเคส รายอำเภอ';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const { ampnameCol } = trendGetCols();
    if (!ampnameCol) {
        destroyTrendCharts();
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์อำเภอในข้อมูล</p>';
        return;
    }

    const distMap = {};
    rows.forEach(r => {
        const dist = String(r[ampnameCol] || 'ไม่ทราบ');
        if (!distMap[dist]) distMap[dist] = { total:0,received:0,tested:0,anemic:0 };
        distMap[dist].total++;
        if (isReceivedValue(r['result'])) distMap[dist].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { distMap[dist].tested++; if (isAnemicRow(r)) distMap[dist].anemic++; }
    });

    const sorted = Object.entries(distMap).sort((a,b) => {
        const rateA = isAnemia ? (a[1].tested>0 ? a[1].anemic/a[1].tested : 0) : (a[1].total>0 ? a[1].received/a[1].total : 0);
        const rateB = isAnemia ? (b[1].tested>0 ? b[1].anemic/b[1].tested : 0) : (b[1].total>0 ? b[1].received/b[1].total : 0);
        return rateB - rateA;
    });

    const dLabels = sorted.map(([k]) => k);
    const dVals = sorted.map(([,b]) => {
        if (isAnemia) return +(b.tested > 0 ? b.anemic/b.tested*100 : 0).toFixed(1);
        return +(b.total > 0 ? b.received/b.total*100 : 0).toFixed(1);
    });

    const mainOpts = {
        chart: {
            type: 'bar', height: Math.max(300, dLabels.length * 28 + 60), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false },
            events: {
                dataPointSelection: (event, chartCtx, config) => {
                    const dist = dLabels[config.dataPointIndex];
                    if (!dist) return;
                    appState.activeDistrictFilter = dist;
                    const distEl = document.getElementById('select-district');
                    if (distEl) distEl.value = dist;
                    appState.activeHospitalFilter = 'all';
                    const hospEl = document.getElementById('select-hospital');
                    if (hospEl) hospEl.value = 'all';
                    applyAllFilters();
                    showToast(`📍 กรองข้อมูลอำเภอ: ${dist}`, 'info', 3000);
                }
            }
        },
        series: [{ name: isAnemia ? 'ร้อยละโลหิตจาง' : 'ความครอบคลุม%', data: dVals }],
        xaxis: { categories: dLabels, labels: { style: { fontSize: '11px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: [TREND_COLORS[0]],
        plotOptions: { bar: { horizontal: true, borderRadius: 3, dataLabels: { position: 'top' } } },
        dataLabels: { enabled: true, formatter: v => v + '%', offsetX: 5, style: { fontSize: '10px', colors: ['#64748b'] } },
        tooltip: { y: { formatter: v => v + '%' }, custom: ({ seriesIndex, dataPointIndex }) => `<div style="padding:6px 10px;font-size:12px;"><b>${dLabels[dataPointIndex]}</b><br/>${dVals[dataPointIndex]}% — คลิกเพื่อกรอง</div>` },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), mainOpts);

    const counts = sorted.map(([,b]) => b.total);
    const sideOpts = {
        chart: { type: 'donut', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: counts,
        labels: dLabels,
        colors: neonColors,
        legend: { position: 'bottom', fontSize: '10px' },
        tooltip: { y: { formatter: v => v.toLocaleString() + ' ราย' } }
    };
    renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
}

function renderTrendUnit(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="building-2"></i> ร้อยละครอบคลุมรายหน่วยบริการ (top 20)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="activity"></i> Coverage% vs เจาะ HCT% รายหน่วย';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const { hosnameCol } = trendGetCols();
    if (!hosnameCol) {
        destroyTrendCharts();
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์หน่วยบริการในข้อมูล</p>';
        return;
    }

    const hospMap = {};
    rows.forEach(r => {
        const h = String(r[hosnameCol] || 'ไม่ทราบ');
        if (!hospMap[h]) hospMap[h] = { total:0,received:0,tested:0,anemic:0 };
        hospMap[h].total++;
        if (isReceivedValue(r['result'])) hospMap[h].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { hospMap[h].tested++; if (isAnemicRow(r)) hospMap[h].anemic++; }
    });

    const sorted = Object.entries(hospMap).sort((a,b) => {
        const rA = isAnemia ? (a[1].tested>0?a[1].anemic/a[1].tested:0) : (a[1].total>0?a[1].received/a[1].total:0);
        const rB = isAnemia ? (b[1].tested>0?b[1].anemic/b[1].tested:0) : (b[1].total>0?b[1].received/b[1].total:0);
        return rB - rA;
    }).slice(0, 20);

    const hLabels = sorted.map(([k]) => k);
    const hVals = sorted.map(([,b]) => +(isAnemia ? (b.tested>0?b.anemic/b.tested*100:0) : (b.total>0?b.received/b.total*100:0)).toFixed(1));

    const mainOpts = {
        chart: { type: 'bar', height: Math.max(320, hLabels.length * 26 + 60), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: isAnemia ? 'ร้อยละโลหิตจาง' : 'ความครอบคลุม%', data: hVals }],
        xaxis: { categories: hLabels, labels: { style: { fontSize: '10px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: [TREND_COLORS[1]],
        plotOptions: { bar: { horizontal: true, borderRadius: 3 } },
        dataLabels: { enabled: true, formatter: v => v + '%', offsetX: 5, style: { fontSize: '10px', colors: ['#64748b'] } },
        tooltip: { y: { formatter: v => v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), mainOpts);

    // Side: grouped bar Coverage% vs HCT-tested% per unit
    const covVals = sorted.map(([,b]) => +(b.total>0?b.received/b.total*100:0).toFixed(1));
    const hctVals = sorted.map(([,b]) => +(b.total>0?b.tested/b.total*100:0).toFixed(1));
    const sideOpts = {
        chart: { type: 'bar', height: Math.max(320, hLabels.length * 26 + 60), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: 'Coverage%', data: covVals }, { name: 'เจาะ HCT%', data: hctVals }],
        xaxis: { categories: hLabels, labels: { style: { fontSize: '10px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: [TREND_COLORS[0], TREND_COLORS[2]],
        plotOptions: { bar: { horizontal: true, borderRadius: 2 } },
        dataLabels: { enabled: false },
        tooltip: { y: { formatter: v => v + '%' } },
        legend: { position: 'top' },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    };
    renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
}

function renderTrendFiscal(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="calendar-range"></i> เปรียบเทียบรายปีงบประมาณ (แนวโน้ม Q1–Q4)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-3"></i> ร้อยละครอบคลุมรายปีงบ';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    // Use all raw data across all FYs
    const fyQMap = {};
    (appState.rawData || []).forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d) return;
        const fy = fiscalYearOf(d);
        const q = fiscalQuarterOf(d);
        if (!fyQMap[fy]) fyQMap[fy] = {};
        if (!fyQMap[fy][q]) fyQMap[fy][q] = { total:0,received:0,tested:0,anemic:0 };
        fyQMap[fy][q].total++;
        if (isReceivedValue(r['result'])) fyQMap[fy][q].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { fyQMap[fy][q].tested++; if (isAnemicRow(r)) fyQMap[fy][q].anemic++; }
    });

    const sortedFYs = Object.keys(fyQMap).map(Number).sort();
    const qLabels = ['Q1 (ต.ค.-ธ.ค.)','Q2 (ม.ค.-มี.ค.)','Q3 (เม.ย.-มิ.ย.)','Q4 (ก.ค.-ก.ย.)'];

    const seriesList = sortedFYs.map(fy => ({
        name: `ปีงบ ${fy}`,
        data: [1,2,3,4].map(q => {
            const b = fyQMap[fy][q] || { total:0,received:0,tested:0,anemic:0 };
            if (isAnemia) return +(b.tested>0?b.anemic/b.tested*100:0).toFixed(1);
            return +(b.total>0?b.received/b.total*100:0).toFixed(1);
        })
    }));

    const mainOpts = {
        chart: { type: 'line', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: seriesList,
        xaxis: { categories: qLabels },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: TREND_COLORS,
        stroke: { curve: 'smooth', width: 2 },
        markers: { size: 5 },
        dataLabels: { enabled: true, formatter: v => v + '%', style: { fontSize: '10px' } },
        tooltip: { y: { formatter: v => v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' },
        legend: { position: 'top' }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), mainOpts);

    // Side: grouped bar — district comparison across last 2 fiscal years
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="map-pin"></i> เปรียบเทียบรายอำเภอ (2 ปีงบล่าสุด)';
    lucide.createIcons();
    const { ampnameCol, ampcodeCol } = trendGetCols();
    const prevFYs = sortedFYs.slice(-2);
    const distCol = ampcodeCol || ampnameCol;
    if (distCol && prevFYs.length >= 1) {
        const distFYMap = {};
        (appState.rawData || []).forEach(r => {
            const d = parseDateValue(r['epi_date']);
            if (!d) return;
            const fy = fiscalYearOf(d);
            if (!prevFYs.includes(fy)) return;
            const dist = String(r[distCol] || 'ไม่ทราบ').trim();
            if (!distFYMap[dist]) distFYMap[dist] = {};
            if (!distFYMap[dist][fy]) distFYMap[dist][fy] = { total:0,received:0,tested:0,anemic:0 };
            distFYMap[dist][fy].total++;
            if (isReceivedValue(r['result'])) distFYMap[dist][fy].received++;
            if (cleanNumericValue(r['lab_result_status']) === 1) {
                distFYMap[dist][fy].tested++;
                if (isAnemicRow(r)) distFYMap[dist][fy].anemic++;
            }
        });
        const dists = Object.keys(distFYMap).sort();
        const sideSeriesList = prevFYs.map(fy => ({
            name: `ปีงบ ${fy}`,
            data: dists.map(dist => {
                const b = distFYMap[dist][fy] || { total:0,received:0,tested:0,anemic:0 };
                if (isAnemia) return +(b.tested>0 ? b.anemic/b.tested*100 : 0).toFixed(1);
                return +(b.total>0 ? b.received/b.total*100 : 0).toFixed(1);
            })
        }));
        const sideOpts = {
            chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
            series: sideSeriesList,
            xaxis: { categories: dists, labels: { style: { fontSize: '10px' }, trim: true, maxHeight: 60 } },
            yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
            colors: TREND_COLORS,
            plotOptions: { bar: { borderRadius: 3, columnWidth: '70%' } },
            dataLabels: { enabled: false },
            tooltip: { y: { formatter: v => v + '%' } },
            legend: { position: 'top' },
            grid: { borderColor: 'rgba(100,116,139,0.15)' }
        };
        renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
    } else {
        // Fallback to annual overall
        const fyAnnual = sortedFYs.map(fy => {
            let total=0,received=0,tested=0,anemic=0;
            [1,2,3,4].forEach(q => { const b = fyQMap[fy][q] || {}; total += b.total||0; received += b.received||0; tested += b.tested||0; anemic += b.anemic||0; });
            return { fy, rate: isAnemia ? (tested>0?+(anemic/tested*100).toFixed(1):0) : (total>0?+(received/total*100).toFixed(1):0) };
        });
        const sideOpts = {
            chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
            series: [{ name: isAnemia ? 'ร้อยละโลหิตจาง' : 'ร้อยละครอบคลุม', data: fyAnnual.map(x => x.rate) }],
            xaxis: { categories: fyAnnual.map(x => `ปีงบ ${x.fy}`) },
            yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
            colors: [TREND_COLORS[3]],
            plotOptions: { bar: { borderRadius: 4, columnWidth: '55%' } },
            dataLabels: { enabled: true, formatter: v => v + '%', style: { fontSize: '12px' } },
            tooltip: { y: { formatter: v => v + '%' } },
            grid: { borderColor: 'rgba(100,116,139,0.15)' }
        };
        renderTrendSide(document.getElementById('trend-chart-side'), sideOpts);
    }
}

function renderScatterPlot(rows) {
    const card = document.getElementById('chart-scatter-card');
    if (!card) return;
    card.style.display = 'flex';

    const high = [], medium = [], low = [];
    rows.forEach(r => {
        const age = cleanNumericValue(r['age_m']);
        const lr = parseFloat(r['labresult']);
        if (isNaN(lr) || lr <= 0 || age <= 0) return;
        const level = getRiskLevel(r);
        const point = { x: age, y: lr, hn: r['hn'] || '', nm: ((r['fname'] || '') + ' ' + (r['lname'] || '')).trim() };
        if (level === 'high') high.push(point);
        else if (level === 'medium') medium.push(point);
        else low.push(point);
    });

    if (high.length + medium.length + low.length === 0) { card.style.display = 'none'; return; }

    const opts = {
        series: [
            { name: 'ความเสี่ยงสูง', data: high },
            { name: 'ความเสี่ยงปานกลาง', data: medium },
            { name: 'ปกติ', data: low },
        ],
        chart: { type: 'scatter', height: 320, background: 'transparent', foreColor: '#475569', toolbar: { show: true }, animations: { enabled: false } },
        colors: ['#dc2626', '#f59e0b', '#16a34a'],
        markers: { size: 4, opacity: 0.65 },
        xaxis: { title: { text: 'อายุ (เดือน)' }, min: 6, max: 60, tickAmount: 9 },
        yaxis: { title: { text: 'ค่า Lab Result' } },
        tooltip: {
            custom({ seriesIndex, dataPointIndex, w }) {
                const d = w.config.series[seriesIndex].data[dataPointIndex];
                const hn = d.hn ? `HN: ${d.hn}` : '';
                const nm = d.nm ? `<br>ชื่อ: ${d.nm}` : '';
                return `<div style="padding:8px;font-size:12px;line-height:1.6"><strong style="color:${['#dc2626','#f59e0b','#16a34a'][seriesIndex]}">${w.globals.seriesNames[seriesIndex]}</strong><br>อายุ: ${d.x} เดือน<br>Lab: ${d.y}<br>${hn}${nm}</div>`;
            }
        },
        legend: { position: 'top' },
        grid: { borderColor: 'rgba(15,23,42,0.08)' }
    };

    if (charts.scatter) { charts.scatter.destroy(); charts.scatter = null; }
    charts.scatter = new ApexCharts(document.getElementById('chart-scatter'), opts);
    charts.scatter.render();
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
    if (appState.isMophMode) {
        const thRisk = document.createElement('th');
        thRisk.textContent = 'ความเสี่ยง';
        thRisk.style.cssText = 'width:76px;min-width:70px;';
        tableHeader.appendChild(thRisk);
    }
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
        td.colSpan = (appState.headers.length || 1) + (appState.isMophMode ? 1 : 0);
        td.style.textAlign = 'center';
        td.style.color = 'var(--text-muted)';
        td.textContent = 'ไม่พบข้อมูลที่ตรงกับการค้นหา';
        tr.appendChild(td);
        tableBody.appendChild(tr);
    } else {
        pageRows.forEach(row => {
            const tr = document.createElement('tr');
            if (appState.isMophMode) {
                const tdRisk = document.createElement('td');
                const risk = getRiskLevel(row);
                if (risk === 'high') {
                    tdRisk.innerHTML = '<span style="background:rgba(220,38,38,0.12);color:#dc2626;border:1px solid rgba(220,38,38,0.4);border-radius:12px;padding:2px 7px;font-size:0.7rem;font-weight:700;">สูง</span>';
                    tr.style.background = 'rgba(220,38,38,0.03)';
                } else if (risk === 'medium') {
                    tdRisk.innerHTML = '<span style="background:rgba(245,158,11,0.12);color:#d97706;border:1px solid rgba(245,158,11,0.4);border-radius:12px;padding:2px 7px;font-size:0.7rem;font-weight:700;">กลาง</span>';
                } else {
                    tdRisk.innerHTML = '<span style="color:var(--text-muted);font-size:0.75rem;">-</span>';
                }
                tr.appendChild(tdRisk);
            }
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

// Print / save as PDF — browser's native print dialog
function printReport() {
    const printWin = window.open('', '_blank');
    if (!printWin) { showToast('⚠️ เบราว์เซอร์บล็อก popup — กรุณาอนุญาต popup แล้วลองใหม่', 'warn', 4000); return; }

    const fy = appState.fiscalYear || '';
    const rows = appState.filteredData || [];
    const received = appState.isMophMode ? rows.filter(r => isReceivedValue(r['result'])).length : null;
    const coverage = received !== null && rows.length > 0 ? (received / rows.length * 100).toFixed(1) : null;
    const totalTested = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
    const totalAnemia = rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;
    const anemiaRate = totalTested > 0 ? (totalAnemia / totalTested * 100).toFixed(1) : null;

    const now = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });

    const kpiRows = appState.isMophMode ? `
        <tr><th>ปีงบประมาณ</th><td>${fy}</td></tr>
        <tr><th>เด็กกลุ่มเป้าหมาย</th><td>${rows.length.toLocaleString()} ราย</td></tr>
        <tr><th>ได้รับยาเสริมธาตุเหล็ก</th><td>${received !== null ? received.toLocaleString() + ' ราย' : '-'}</td></tr>
        <tr><th>ร้อยละครอบคลุม</th><td>${coverage !== null ? coverage + '%' : '-'}</td></tr>
        <tr><th>เจาะ Lab HCT</th><td>${totalTested.toLocaleString()} ราย</td></tr>
        <tr><th>โลหิตจาง</th><td>${totalAnemia.toLocaleString()} ราย (${anemiaRate !== null ? anemiaRate + '%' : '-'})</td></tr>
    ` : `<tr><th>จำนวนแถว</th><td>${rows.length.toLocaleString()}</td></tr>`;

    printWin.document.write(`<!DOCTYPE html><html lang="th"><head>
        <meta charset="UTF-8">
        <title>รายงาน MCHMUK HDC มุกดาหาร ปีงบ ${fy}</title>
        <style>
            body { font-family: 'Sarabun', 'Tahoma', sans-serif; font-size: 13px; color: #0f172a; margin: 24px; }
            h1 { font-size: 16px; text-align: center; margin-bottom: 4px; }
            .sub { text-align: center; color: #64748b; font-size: 11px; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
            th, td { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: left; }
            th { background: #f1f5f9; font-weight: 600; }
            @media print { @page { margin: 15mm; } }
        </style>
    </head><body>
        <h1>รายงานสรุป HDC มุกดาหาร — ${appState.isMophMode ? 'ยาเสริมธาตุเหล็ก/โลหิตจาง' : appState.currentSheetName}</h1>
        <p class="sub">พิมพ์เมื่อ ${now} น.</p>
        <table><tbody>${kpiRows}</tbody></table>
    </body></html>`);
    printWin.document.close();
    printWin.focus();
    setTimeout(() => { printWin.print(); printWin.close(); }, 400);
}

function exportExcel() {
    if (!appState.filteredData || appState.filteredData.length === 0) {
        showToast('⚠️ ไม่มีข้อมูลสำหรับ Export', 'warn', 3000);
        return;
    }
    try {
        const wb = XLSX.utils.book_new();

        // Sheet 1: KPI Summary
        const rows = appState.filteredData;
        const kpiData = [['ตัวชี้วัด', 'ค่า']];
        if (appState.isMophMode) {
            const received = rows.filter(r => isReceivedValue(r['result'])).length;
            const totalTested = appState.preHctTestedCount != null ? appState.preHctTestedCount : rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1).length;
            const totalAnemia = appState.preHctAnemiaCount != null ? appState.preHctAnemiaCount : rows.filter(r => cleanNumericValue(r['lab_result_status']) === 1 && isAnemicRow(r)).length;
            const base = appState.preHctFilterCount > 0 ? appState.preHctFilterCount : rows.length;
            kpiData.push(['ปีงบประมาณ', appState.fiscalYear]);
            kpiData.push(['จำนวนเด็กกลุ่มเป้าหมาย (ราย)', rows.length]);
            kpiData.push(['ได้รับยาเสริมธาตุเหล็ก (ราย)', received]);
            kpiData.push(['ร้อยละครอบคลุม (%)', rows.length > 0 ? +(received / rows.length * 100).toFixed(2) : 0]);
            kpiData.push(['เจาะ Lab HCT (ราย)', totalTested]);
            kpiData.push(['ร้อยละการเจาะ Lab (%)', base > 0 ? +(totalTested / base * 100).toFixed(2) : 0]);
            kpiData.push(['โลหิตจาง (ราย)', totalAnemia]);
        } else {
            kpiData.push(['จำนวนแถวทั้งหมด', rows.length]);
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(kpiData), 'KPI Summary');

        // Sheet 2: Filtered Data
        const exportHeaders = appState.headers.filter(h => h !== '__rowIdx__');
        const dataArray = [exportHeaders, ...appState.filteredData.map(row => exportHeaders.map(h => getExportValue(row, h)))];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataArray), 'ข้อมูล');

        const dateStr = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `รายงาน_MCHMUK_${dateStr}.xlsx`);
        showToast(`📥 Export Excel สำเร็จ! ${appState.filteredData.length.toLocaleString()} แถว`, 'success', 4000);
    } catch(err) {
        showToast(`❌ Export Excel ล้มเหลว: ${err.message}`, 'error', 5000);
    }
}

function exportPDF() {
    window.print();
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
        appState._xlsxBuffer = buffer;
        // Persist to IDB for offline use (non-blocking, best-effort)
        saveActiveDataset(buffer).catch(() => {});

        // Parse via Worker — non-blocking
        const workerResult = await parseExcelWithWorker(buffer.slice(0), null);
        const sheetNames = workerResult.sheetNames;

        appState.sheetNames = sheetNames;

        document.getElementById('val-filename').textContent = GITHUB_CONFIG.filePath;
        document.getElementById('val-filesize').textContent = formatBytes(buffer.byteLength);
        document.getElementById('val-sheets').textContent = sheetNames.length;

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
        sheetNames.forEach(sn => {
            const opt = document.createElement('option');
            opt.value = sn; opt.textContent = sn;
            sheetSelect.appendChild(opt);
        });

        const targetSheet = workerResult.targetSheet;
        sheetSelect.value = targetSheet;

        // Restore hash filters before rendering if hash present
        if (location.hash) { restoreStateFromHash(); syncUIControlsFromState(); }

        await loadSheetData(targetSheet);

        document.getElementById('val-status').textContent = "เปิดทำงานปกติ";
        document.getElementById('val-status').className = "status-badge success";
        toggleLoader(false);

        if (!isSilent) {
            showToast('☁️ ดึงข้อมูลจากเซิร์ฟเวอร์กลางสำเร็จ! ข้อมูลเป็นปัจจุบัน', 'success', 4000);
        }

    } catch (err) {
        console.warn("Server fetch failed:", err.message);
        // Try IDB fallback for offline mode
        try {
            const cached = await getActiveDataset();
            if (cached && cached.buffer) {
                console.log('IDB: loading cached dataset (offline mode)');
                appState._xlsxBuffer = cached.buffer;
                const workerResult = await parseExcelWithWorker(cached.buffer.slice(0), null);
                const sheetNames = workerResult.sheetNames;
                appState.sheetNames = sheetNames;
                const sheetSelect = document.getElementById('select-sheet');
                sheetSelect.innerHTML = '';
                sheetNames.forEach(sn => { const o = document.createElement('option'); o.value = sn; o.textContent = sn; sheetSelect.appendChild(o); });
                sheetSelect.value = workerResult.targetSheet;
                if (location.hash) { restoreStateFromHash(); syncUIControlsFromState(); }
                await loadSheetData(workerResult.targetSheet);
                document.getElementById('val-status').textContent = "ออฟไลน์ (แคช)";
                document.getElementById('val-status').className = "status-badge pending";
                toggleLoader(false);
                showToast('📱 โหมดออฟไลน์ — แสดงข้อมูลที่แคชไว้ล่าสุด', 'warn', 5000);
                return;
            }
        } catch (idbErr) { console.warn('IDB fallback failed:', idbErr.message); }
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
 * Push Excel file to GitHub via Cloudflare Worker proxy
 * Token เก็บใน Worker secret — ไม่ต้องตั้งค่าที่ browser
 */
async function pushExcelToGitHub(arrayBuffer, originalFilename) {
    const workerUrl = GITHUB_CONFIG.workerUrl;
    const actionsUrl = `https://github.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/actions`;

    try {
        // Step 1: ดึง SHA ของไฟล์เดิมจาก Worker
        showToast('☁️ [1/3] กำลังเชื่อมต่อ...', 'info', 3000);
        const shaRes = await fetch(`${workerUrl}/sha?path=${GITHUB_CONFIG.filePath}`);
        const shaJson = await shaRes.json();
        const existingSha = shaJson.sha || null;

        // Step 2: แปลง ArrayBuffer → Base64
        showToast('☁️ [2/3] กำลังเตรียมข้อมูล...', 'info', 3000);
        const base64 = arrayBufferToBase64(arrayBuffer);

        // Step 3: ดึง meta SHA
        const metaShaRes = await fetch(`${workerUrl}/sha?path=tmp_exchange_data_meta.json`).catch(() => null);
        const metaShaJson = metaShaRes?.ok ? await metaShaRes.json() : {};
        const existingMetaSha = metaShaJson.sha || null;

        // Step 4: สร้าง meta content
        const now = new Date();
        const commitMsg = `📊 อัปเดตข้อมูล HDC โดย Admin เมื่อ ${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' })} น. (${originalFilename})`;
        const metaObj = {
            importTimestamp: now.toISOString(),
            filename: originalFilename,
            filesize: arrayBuffer.byteLength,
        };
        const metaContent = btoa(JSON.stringify(metaObj, null, 2));

        // Step 5: Push ผ่าน Worker
        showToast('☁️ [3/3] กำลัง Push ข้อมูลขึ้น GitHub...', 'info', 4000);
        const uploadRes = await fetch(`${workerUrl}/upload`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                content: base64,
                sha: existingSha,
                commitMsg,
                metaContent,
                metaSha: existingMetaSha,
            }),
        });

        const result = await uploadRes.json();

        if (result.ok) {
            const extraWarn = result.metaWarning ? `<br>⚠️ Metadata: ${result.metaWarning}` : '';
            showToast(
                `✅ Push สำเร็จ! (commit: ${result.commitSha})<br>` +
                `🚀 GitHub Actions กำลัง Deploy อัตโนมัติ — ` +
                `<a href="${actionsUrl}" target="_blank" style="color:#0284c7;text-decoration:underline;">ดู Actions →</a><br>` +
                `⏱️ ผู้ใช้ทุกคนจะเห็นข้อมูลใหม่ภายใน ~1-2 นาที${extraWarn}`,
                'success', 10000
            );
            if (result.metaWarning) console.warn('meta push warning:', result.metaWarning);
            console.log(`☁️ push success — commit: ${result.commitSha}`);
        } else {
            if (result.status === 401 || result.status === 403) {
                showToast('❌ Worker token หมดอายุ — กรุณาแจ้ง Admin อัปเดต secret', 'error', 8000);
            } else {
                showToast(`❌ Push ไม่สำเร็จ (${result.status}): ${result.message}`, 'error', 6000);
            }
            console.error('push failed:', result);
        }
    } catch (e) {
        showToast(`❌ เกิดข้อผิดพลาด: ${e.message}`, 'error', 6000);
        console.error('push error:', e);
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

function refreshBatchList() {}
function loadLatestActiveBatch() {}

// ==========================================================================
// 💾 IndexedDB — offline data persistence
//    Stores the raw XLSX ArrayBuffer so the dashboard works offline after
//    the first online load. Key: 'mchmuk-xlsx-buffer'
// ==========================================================================
const IDB_NAME = 'mchmuk-offline';
const IDB_STORE = 'datasets';
const IDB_KEY = 'mchmuk-xlsx-buffer';

function openIDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = e => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
        };
        req.onsuccess = e => resolve(e.target.result);
        req.onerror = e => reject(e.target.error);
    });
}

async function saveActiveDataset(buffer) {
    try {
        const db = await openIDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).put({ buffer, savedAt: Date.now() }, IDB_KEY);
            tx.oncomplete = () => { db.close(); resolve({ ok: true }); };
            tx.onerror = e => { db.close(); reject(e.target.error); };
        });
    } catch (err) {
        console.warn('IDB save failed:', err.message);
        return { ok: false };
    }
}

async function getActiveDataset() {
    try {
        const db = await openIDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readonly');
            const req = tx.objectStore(IDB_STORE).get(IDB_KEY);
            req.onsuccess = e => { db.close(); resolve(e.target.result || null); };
            req.onerror = e => { db.close(); reject(e.target.error); };
        });
    } catch (err) {
        console.warn('IDB load failed:', err.message);
        return null;
    }
}

// ==========================================================================
// 🌙 Dark Mode
// ==========================================================================
function toggleDarkMode() {
    const html = document.documentElement;
    const isDark = html.getAttribute('data-theme') === 'dark';
    html.setAttribute('data-theme', isDark ? 'light' : 'dark');
    const icon = document.getElementById('dark-mode-icon');
    if (icon) icon.setAttribute('data-lucide', isDark ? 'moon' : 'sun');
    lucide.createIcons();
    try { localStorage.setItem('mchmuk-theme', isDark ? 'light' : 'dark'); } catch (e) {}
}

(function applyStoredTheme() {
    try {
        const t = localStorage.getItem('mchmuk-theme');
        if (t) {
            document.documentElement.setAttribute('data-theme', t);
            // Update icon after DOM is ready
            window.addEventListener('DOMContentLoaded', () => {
                const icon = document.getElementById('dark-mode-icon');
                if (icon) icon.setAttribute('data-lucide', t === 'dark' ? 'sun' : 'moon');
            }, { once: true });
        }
    } catch (e) {}
})();

// ==========================================================================
// 🔄 Reset All Filters
// ==========================================================================
function resetAllFilters() {
    if (!appState.isMophMode) return;
    appState.activeAgeFilter = 'all';
    appState.activeHctFilter = 'all';
    appState.activeDistrictFilter = 'all';
    appState.activeHospitalFilter = 'all';
    appState.activeQuarter = 'all';
    appState.activeRiskFilter = 'all';
    appState.advFilterAgeMin = null;
    appState.advFilterAgeMax = null;
    appState.advFilterLabMin = null;
    appState.advFilterLabMax = null;
    appState.advFilterDateFrom = null;
    appState.advFilterDateTo = null;
    appState.advFilterAnemia = 'all';
    // reset UI controls
    const q = document.getElementById('quarter-select'); if (q) q.value = 'all';
    const d = document.getElementById('select-district'); if (d) d.value = 'all';
    const h = document.getElementById('select-hospital'); if (h) h.value = 'all';
    document.querySelectorAll('[data-age-range]').forEach(b => b.classList.toggle('active', b.dataset.ageRange === 'all'));
    document.querySelectorAll('[data-hct-status]').forEach(b => b.classList.toggle('active', b.dataset.hctStatus === 'all'));
    applyAllFilters();
    showToast('🔄 ล้าง Filter ทั้งหมดแล้ว', 'info', 2500);
}

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') return;
    if (e.key === 'Escape') { resetAllFilters(); closeFilterSheet(); closeDQModal(); closeAnnotationModal(); }
    if (e.key === 'F5' && !e.ctrlKey) { e.preventDefault(); loadLocalExcelFile(false); }
    if (e.key === 'd' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); toggleDarkMode(); }
});

// ==========================================================================
// 📊 Heatmap: District × Month Coverage
// ==========================================================================
function renderTrendHeatmap(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="grid"></i> Heatmap: ความครอบคลุมรายอำเภอ × เดือน (%)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-3"></i> ค่าเฉลี่ยรายอำเภอ';
    lucide.createIcons();

    const { ampnameCol } = trendGetCols();
    if (!ampnameCol) {
        destroyTrendCharts();
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์อำเภอ</p>';
        return;
    }

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    // Build dist × month map
    const dists = [...new Set(rows.map(r => String(r[ampnameCol] || 'ไม่ทราบ')))].sort();
    const mapData = {};
    dists.forEach(d => { mapData[d] = {}; FY_MONTH_ORDER.forEach(m => { mapData[d][m] = { total: 0, received: 0, tested: 0, anemic: 0 }; }); });

    rows.forEach(r => {
        const dist = String(r[ampnameCol] || 'ไม่ทราบ');
        const date = parseDateValue(r['epi_date']);
        if (!date || !mapData[dist]) return;
        const m = date.getMonth();
        if (mapData[dist][m] === undefined) return;
        mapData[dist][m].total++;
        if (isReceivedValue(r['result'])) mapData[dist][m].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { mapData[dist][m].tested++; if (isAnemicRow(r)) mapData[dist][m].anemic++; }
    });

    const series = dists.map(dist => ({
        name: dist,
        data: FY_MONTH_ORDER.map((m, i) => {
            const b = mapData[dist][m];
            const val = isAnemia
                ? (b.tested > 0 ? +(b.anemic / b.tested * 100).toFixed(1) : 0)
                : (b.total > 0 ? +(b.received / b.total * 100).toFixed(1) : 0);
            return { x: FY_MONTH_LABELS[i], y: val };
        })
    }));

    const opts = {
        chart: { type: 'heatmap', height: Math.max(240, dists.length * 32 + 60), background: 'transparent', foreColor: '#64748b', toolbar: { show: true, tools: { download: true } }, animations: { enabled: false } },
        series,
        dataLabels: { enabled: true, style: { fontSize: '10px', colors: ['#1e293b'] } },
        colors: isAnemia ? ['#ef4444'] : ['#0ea5e9'],
        plotOptions: { heatmap: { shadeIntensity: 0.5, colorScale: { ranges: isAnemia
            ? [{ from: 0, to: 10, color: '#bbf7d0', name: 'ต่ำ' }, { from: 10.1, to: 16, color: '#fde68a', name: 'กลาง' }, { from: 16.1, to: 100, color: '#fca5a5', name: 'สูง' }]
            : [{ from: 0, to: 50, color: '#fca5a5', name: 'ต่ำกว่าเป้า' }, { from: 50.1, to: 74.9, color: '#fde68a', name: 'ใกล้เป้า' }, { from: 75, to: 100, color: '#bbf7d0', name: 'ผ่านเป้า' }]
        } } },
        tooltip: { y: { formatter: v => v + '%' } },
        xaxis: { labels: { style: { fontSize: '10px' } } }
    };
    renderTrendMain(document.getElementById('trend-chart-main'), opts);

    // Side: average per district
    const avgVals = dists.map(dist => {
        const vals = FY_MONTH_ORDER.map(m => { const b = mapData[dist][m]; return isAnemia ? (b.tested > 0 ? b.anemic / b.tested * 100 : 0) : (b.total > 0 ? b.received / b.total * 100 : 0); }).filter(v => v > 0);
        return vals.length ? +(vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(1) : 0;
    });
    renderTrendSide(document.getElementById('trend-chart-side'), {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: 'ค่าเฉลี่ย%', data: avgVals }],
        xaxis: { categories: dists, labels: { style: { fontSize: '10px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: [TREND_COLORS[0]],
        plotOptions: { bar: { horizontal: true, borderRadius: 3 } },
        dataLabels: { enabled: true, formatter: v => v + '%', style: { fontSize: '10px', colors: ['#64748b'] } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });
}

// ==========================================================================
// 🏆 Ranking Table — hospitals with trend arrows
// ==========================================================================
function renderTrendRanking(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="trophy"></i> จัดอันดับหน่วยบริการ';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-2"></i> สรุป Top/Bottom 5';
    lucide.createIcons();
    destroyTrendCharts();

    const { hosnameCol, ampnameCol } = trendGetCols();
    if (!hosnameCol) {
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์หน่วยบริการ</p>';
        return;
    }

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const curFY = appState.fiscalYear;
    const prevFY = curFY - 1;

    // Build maps for current and previous FY
    const buildMap = (data, fy) => {
        const m = {};
        data.forEach(r => {
            const d = parseDateValue(r['epi_date']);
            if (!d || fiscalYearOf(d) !== fy) return;
            const h = String(r[hosnameCol] || 'ไม่ทราบ');
            if (!m[h]) m[h] = { total: 0, received: 0, tested: 0, anemic: 0, district: ampnameCol ? String(r[ampnameCol] || '') : '' };
            m[h].total++;
            if (isReceivedValue(r['result'])) m[h].received++;
            if (cleanNumericValue(r['lab_result_status']) === 1) { m[h].tested++; if (isAnemicRow(r)) m[h].anemic++; }
        });
        return m;
    };
    const curMap = buildMap(appState.rawData || rows, curFY);
    const prevMap = buildMap(appState.rawData || rows, prevFY);

    const getRate = (b) => isAnemia ? (b.tested > 0 ? b.anemic / b.tested * 100 : 0) : (b.total > 0 ? b.received / b.total * 100 : 0);
    const ranked = Object.entries(curMap)
        .map(([name, cur]) => {
            const curRate = getRate(cur);
            const prev = prevMap[name];
            const prevRate = prev ? getRate(prev) : null;
            const diff = prevRate !== null ? curRate - prevRate : null;
            return { name, district: cur.district, total: cur.total, curRate, prevRate, diff };
        })
        .filter(r => r.total >= 3)
        .sort((a, b) => isAnemia ? a.curRate - b.curRate : b.curRate - a.curRate);

    const arrow = (diff) => {
        if (diff === null) return '<span style="color:#94a3b8;">–</span>';
        const isGood = isAnemia ? diff < 0 : diff > 0;
        const isNeutral = Math.abs(diff) < 0.5;
        if (isNeutral) return `<span style="color:#94a3b8;">→ ทรง</span>`;
        return isGood
            ? `<span style="color:#15803d;">▲ +${Math.abs(diff).toFixed(1)}%</span>`
            : `<span style="color:#dc2626;">▼ -${Math.abs(diff).toFixed(1)}%</span>`;
    };

    const rows2 = ranked.map((r, i) => {
        const pct = r.curRate.toFixed(1);
        const target = isAnemia ? (r.curRate <= 16 ? 'met' : '') : (r.curRate >= getHospitalTarget(r.name) ? 'met' : '');
        const note = getAnnotation(r.name);
        const badgeCls = target === 'met' ? 'color:#15803d;background:rgba(21,128,61,0.1);border:1px solid rgba(21,128,61,0.3);' : 'color:#dc2626;background:rgba(220,38,38,0.08);border:1px solid rgba(220,38,38,0.2);';
        return `<tr>
            <td style="text-align:center;font-weight:700;color:var(--text-muted);">${i + 1}</td>
            <td style="font-weight:600;">${escapeHtml(r.name)} <button class="annot-btn${note ? ' has-note' : ''}" data-hosp="${escapeHtml(r.name)}" title="${note ? escapeHtml(note) : 'เพิ่มหมายเหตุ / ตั้งเป้า'}">${note ? '📝' : '✎'}</button></td>
            <td style="color:var(--text-muted);font-size:0.8rem;">${escapeHtml(r.district)}</td>
            <td style="text-align:center;">${r.total.toLocaleString()}</td>
            <td style="text-align:center;"><span style="padding:3px 10px;border-radius:12px;font-weight:700;${badgeCls}">${pct}%</span></td>
            <td style="text-align:center;">${arrow(r.diff)}</td>
        </tr>`;
    }).join('');

    document.getElementById('trend-chart-main').innerHTML = `
        <div style="overflow-x:auto;max-height:420px;overflow-y:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
            <thead style="position:sticky;top:0;background:var(--bg-card,white);">
                <tr style="border-bottom:2px solid var(--border-light);">
                    <th style="padding:8px 6px;text-align:center;color:var(--text-muted);">#</th>
                    <th style="padding:8px 6px;text-align:left;">หน่วยบริการ</th>
                    <th style="padding:8px 6px;text-align:left;">อำเภอ</th>
                    <th style="padding:8px 6px;text-align:center;">เด็ก</th>
                    <th style="padding:8px 6px;text-align:center;">${isAnemia ? 'โลหิตจาง%' : 'Coverage%'}</th>
                    <th style="padding:8px 6px;text-align:center;">เทียบปีก่อน</th>
                </tr>
            </thead>
            <tbody>${rows2 || '<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--text-muted);">ไม่มีข้อมูลเปรียบเทียบ</td></tr>'}</tbody>
        </table>
        </div>`;
    document.querySelectorAll('#trend-chart-main .annot-btn').forEach(btn => {
        btn.addEventListener('click', () => showAnnotationModal(btn.dataset.hosp));
    });

    // Side: top5 / bottom5 bar
    const top5 = ranked.slice(0, 5);
    const bot5 = ranked.slice(-5).reverse();
    const sideLabels = [...top5.map(r => r.name.slice(0, 12)), '---', ...bot5.map(r => r.name.slice(0, 12))];
    const sideVals = [...top5.map(r => r.curRate), null, ...bot5.map(r => r.curRate)];
    const sideColors = [...top5.map(r => isAnemia ? (r.curRate <= 16 ? '#10b981' : '#ef4444') : (r.curRate >= 75 ? '#10b981' : '#ef4444')), '#94a3b8', ...bot5.map(r => isAnemia ? (r.curRate <= 16 ? '#10b981' : '#ef4444') : (r.curRate >= 75 ? '#10b981' : '#ef4444'))];
    renderTrendSide(document.getElementById('trend-chart-side'), {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: isAnemia ? 'โลหิตจาง%' : 'Coverage%', data: sideVals.filter(v => v !== null) }],
        xaxis: { categories: sideLabels.filter((_, i) => sideVals[i] !== null), labels: { style: { fontSize: '9px' } } },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors: sideColors.filter((_, i) => sideVals[i] !== null),
        plotOptions: { bar: { distributed: true, horizontal: false, borderRadius: 3 } },
        legend: { show: false },
        dataLabels: { enabled: true, formatter: v => v.toFixed(1) + '%', style: { fontSize: '9px', colors: ['#64748b'] } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });
}

// ==========================================================================
// 🔍 Data Quality Report
// ==========================================================================
function showDataQualityReport() {
    const overlay = document.getElementById('dq-modal-overlay');
    const content = document.getElementById('dq-content');
    if (!overlay || !content) return;

    const rows = appState.rawData || appState.filteredData || [];
    const headers = appState.headers || [];

    if (rows.length === 0 || headers.length === 0) {
        content.innerHTML = '<p style="color:var(--text-muted);padding:20px;">ยังไม่มีข้อมูล</p>';
        overlay.style.display = 'flex';
        lucide.createIcons();
        return;
    }

    // Missing value analysis
    const colStats = headers.filter(h => h !== '__rowIdx__').map(h => {
        const vals = rows.map(r => r[h]);
        const missing = vals.filter(v => v === null || v === undefined || v === '').length;
        const missRate = (missing / rows.length * 100).toFixed(1);
        return { h, missing, missRate: parseFloat(missRate) };
    }).sort((a, b) => b.missRate - a.missRate);

    // Duplicate HN detection
    const hnCol = headers.find(h => /^hn$|^patient_hn$|^cid$/i.test(h));
    let dupHTML = '';
    if (hnCol) {
        const counts = {};
        rows.forEach(r => { const v = String(r[hnCol] || ''); if (v) counts[v] = (counts[v] || 0) + 1; });
        const dups = Object.entries(counts).filter(([, c]) => c > 1).sort((a, b) => b[1] - a[1]);
        dupHTML = `<div style="margin-top:16px;">
            <h4 style="font-size:0.85rem;font-weight:700;margin-bottom:8px;color:var(--text-secondary);">
                <i data-lucide="copy" style="width:14px;height:14px;vertical-align:middle;"></i>
                ตรวจสอบ ${hnCol} ซ้ำ — พบ ${dups.length.toLocaleString()} รายการซ้ำ จาก ${Object.keys(counts).length.toLocaleString()} HN ทั้งหมด
            </h4>
            ${dups.length > 0 ? `<div style="max-height:120px;overflow-y:auto;font-size:0.75rem;border:1px solid var(--border-light);border-radius:6px;padding:8px;">
                ${dups.slice(0, 50).map(([v, c]) => `<span style="display:inline-block;margin:2px 4px;padding:2px 8px;background:rgba(220,38,38,0.08);border-radius:10px;color:#dc2626;">${v} ×${c}</span>`).join('')}
                ${dups.length > 50 ? `<span style="color:var(--text-muted);">...และอีก ${dups.length - 50} รายการ</span>` : ''}
            </div>` : '<span style="color:#15803d;font-size:0.82rem;">✓ ไม่พบรายการซ้ำ</span>'}
        </div>`;
    }

    const missRows = colStats.map(s => {
        const bg = s.missRate > 20 ? 'rgba(220,38,38,0.08)' : s.missRate > 5 ? 'rgba(245,158,11,0.08)' : 'transparent';
        const col = s.missRate > 20 ? '#dc2626' : s.missRate > 5 ? '#d97706' : '#15803d';
        return `<tr style="background:${bg};">
            <td style="padding:4px 8px;font-size:0.8rem;font-weight:500;">${s.h}</td>
            <td style="padding:4px 8px;text-align:center;font-size:0.8rem;">${s.missing.toLocaleString()}</td>
            <td style="padding:4px 8px;text-align:center;"><span style="color:${col};font-weight:700;">${s.missRate}%</span></td>
        </tr>`;
    }).join('');

    content.innerHTML = `
        <div style="margin-bottom:12px;padding:10px 14px;background:rgba(2,132,199,0.06);border-radius:8px;font-size:0.82rem;color:var(--text-secondary);">
            <strong>ข้อมูลทั้งหมด:</strong> ${rows.length.toLocaleString()} แถว · ${headers.length} คอลัมน์
        </div>
        <h4 style="font-size:0.85rem;font-weight:700;margin-bottom:8px;color:var(--text-secondary);">
            <i data-lucide="alert-circle" style="width:14px;height:14px;vertical-align:middle;"></i>
            ข้อมูลว่าง (Missing Values) รายคอลัมน์
        </h4>
        <div style="max-height:240px;overflow-y:auto;border:1px solid var(--border-light);border-radius:8px;">
            <table style="width:100%;border-collapse:collapse;">
                <thead style="position:sticky;top:0;background:var(--bg-card,white);">
                    <tr style="border-bottom:1px solid var(--border-light);">
                        <th style="padding:6px 8px;text-align:left;font-size:0.78rem;color:var(--text-muted);">คอลัมน์</th>
                        <th style="padding:6px 8px;text-align:center;font-size:0.78rem;color:var(--text-muted);">ว่าง (แถว)</th>
                        <th style="padding:6px 8px;text-align:center;font-size:0.78rem;color:var(--text-muted);">%</th>
                    </tr>
                </thead>
                <tbody>${missRows}</tbody>
            </table>
        </div>
        ${dupHTML}`;

    overlay.style.display = 'flex';
    lucide.createIcons();
}

function closeDQModal() {
    const overlay = document.getElementById('dq-modal-overlay');
    if (overlay) overlay.style.display = 'none';
}

// ==========================================================================
// 📲 PWA Install Prompt
// ==========================================================================
let _deferredInstallPrompt = null;

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    _deferredInstallPrompt = e;
    const banner = document.getElementById('pwa-install-banner');
    if (banner) { banner.style.display = 'flex'; lucide.createIcons(); }
});

function pwaInstall() {
    if (!_deferredInstallPrompt) return;
    _deferredInstallPrompt.prompt();
    _deferredInstallPrompt.userChoice.then(() => {
        _deferredInstallPrompt = null;
        const banner = document.getElementById('pwa-install-banner');
        if (banner) banner.style.display = 'none';
    });
}

window.addEventListener('appinstalled', () => {
    const banner = document.getElementById('pwa-install-banner');
    if (banner) banner.style.display = 'none';
    showToast('✅ ติดตั้งแอป MCHMUK สำเร็จ!', 'success', 4000);
});

// ==========================================================================
// 🧰 Shared helpers
// ==========================================================================
function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function linearRegression(ys) {
    const n = ys.length;
    const xMean = (n - 1) / 2;
    const yMean = ys.reduce((s, v) => s + v, 0) / n;
    let num = 0, den = 0;
    ys.forEach((y, x) => { num += (x - xMean) * (y - yMean); den += (x - xMean) ** 2; });
    const slope = den === 0 ? 0 : num / den;
    return { slope, intercept: yMean - slope * xMean };
}

function readLocalJSON(key) {
    try { return JSON.parse(localStorage.getItem(key) || '{}') || {}; } catch (e) { return {}; }
}
function writeLocalJSON(key, obj) {
    try { localStorage.setItem(key, JSON.stringify(obj)); return true; } catch (e) { return false; }
}

// Rows across all FYs, restricted to the active district/hospital filter
function rawRowsForActiveUnit() {
    const { ampnameCol, hosnameCol } = trendGetCols();
    return (appState.rawData || []).filter(r =>
        (appState.activeDistrictFilter === 'all' || !ampnameCol || String(r[ampnameCol]) === appState.activeDistrictFilter) &&
        (appState.activeHospitalFilter === 'all' || !hosnameCol || String(r[hosnameCol]) === appState.activeHospitalFilter));
}

// ==========================================================================
// 🎯 Per-hospital target override + 📝 annotations (localStorage)
// ==========================================================================
const TARGETS_KEY = 'mchmuk-targets';
const ANNOT_KEY = 'mchmuk-annotations';
const DEFAULT_COVERAGE_TARGET = 75;

function getHospitalTarget(name) {
    const t = Number(readLocalJSON(TARGETS_KEY)[name]);
    return Number.isFinite(t) && t > 0 && t <= 100 ? t : DEFAULT_COVERAGE_TARGET;
}

function getAnnotation(name) {
    return readLocalJSON(ANNOT_KEY)[name] || '';
}

function showAnnotationModal(name) {
    if (!name || name === 'all') {
        showToast('เลือกหน่วยบริการก่อน (จากตัวกรองหรือตารางจัดอันดับ)', 'info', 3500);
        return;
    }
    const overlay = document.getElementById('annot-modal-overlay');
    if (!overlay) return;
    overlay.dataset.hosp = name;
    document.getElementById('annot-hosp-name').textContent = name;
    document.getElementById('annot-text').value = getAnnotation(name);
    const custom = readLocalJSON(TARGETS_KEY)[name];
    document.getElementById('annot-target').value = custom ?? '';
    document.getElementById('annot-target').placeholder = String(DEFAULT_COVERAGE_TARGET);
    overlay.style.display = 'block';
    document.getElementById('annot-text').focus();
}

function closeAnnotationModal() {
    const overlay = document.getElementById('annot-modal-overlay');
    if (overlay) overlay.style.display = 'none';
}

function saveAnnotationModal() {
    const overlay = document.getElementById('annot-modal-overlay');
    const name = overlay?.dataset.hosp;
    if (!name) return;
    const text = document.getElementById('annot-text').value.trim().slice(0, 1000);
    const targetRaw = document.getElementById('annot-target').value.trim();
    const target = targetRaw === '' ? null : Number(targetRaw);
    if (target !== null && (!Number.isFinite(target) || target <= 0 || target > 100)) {
        showToast('เป้าหมายต้องเป็นตัวเลข 1–100', 'error', 3500);
        return;
    }
    const notes = readLocalJSON(ANNOT_KEY);
    if (text) notes[name] = text; else delete notes[name];
    const targets = readLocalJSON(TARGETS_KEY);
    if (target !== null && target !== DEFAULT_COVERAGE_TARGET) targets[name] = target; else delete targets[name];
    if (!writeLocalJSON(ANNOT_KEY, notes) || !writeLocalJSON(TARGETS_KEY, targets)) {
        showToast('บันทึกไม่สำเร็จ (browser storage ถูกปิด)', 'error', 4000);
        return;
    }
    closeAnnotationModal();
    showToast(`💾 บันทึก ${name} แล้ว`, 'success', 2500);
    triggerAnalyticsUpdate();
}

// ==========================================================================
// 🔗 Share URL
// ==========================================================================
async function shareCurrentView() {
    syncHashFromState();
    const url = location.href;
    try {
        await navigator.clipboard.writeText(url);
        showToast('🔗 คัดลอกลิงก์มุมมองปัจจุบันแล้ว', 'success', 3000);
    } catch (e) {
        window.prompt('คัดลอกลิงก์นี้:', url);
    }
}

// ==========================================================================
// 📗 Export XLSX (filtered rows + per-hospital summary)
// ==========================================================================
function exportXLSX() {
    if (!appState.filteredData.length || typeof XLSX === 'undefined') return;
    const headers = appState.headers.filter(h => h !== '__rowIdx__');
    const wb = XLSX.utils.book_new();
    const dataSheet = XLSX.utils.json_to_sheet(appState.filteredData.map(r => {
        const o = {}; headers.forEach(h => { o[h] = r[h]; }); return o;
    }), { header: headers });
    XLSX.utils.book_append_sheet(wb, dataSheet, 'Data');

    if (appState.isMophMode) {
        const { hosnameCol, ampnameCol } = trendGetCols();
        if (hosnameCol) {
            const m = {};
            appState.filteredData.forEach(r => {
                const h = String(r[hosnameCol] || 'ไม่ทราบ');
                if (!m[h]) m[h] = { district: ampnameCol ? String(r[ampnameCol] || '') : '', total: 0, received: 0, tested: 0, anemic: 0 };
                m[h].total++;
                if (isReceivedValue(r['result'])) m[h].received++;
                if (cleanNumericValue(r['lab_result_status']) === 1) { m[h].tested++; if (isAnemicRow(r)) m[h].anemic++; }
            });
            const summary = Object.entries(m).map(([h, b]) => ({
                'หน่วยบริการ': h,
                'อำเภอ': b.district,
                'เด็กเป้าหมาย': b.total,
                'ได้รับยา': b.received,
                'Coverage%': +(b.total ? b.received / b.total * 100 : 0).toFixed(1),
                'เป้าหมาย%': getHospitalTarget(h),
                'เจาะ HCT': b.tested,
                'โลหิตจาง': b.anemic,
                'โลหิตจาง%': +(b.tested ? b.anemic / b.tested * 100 : 0).toFixed(1),
                'หมายเหตุ': getAnnotation(h),
            })).sort((a, b) => b['Coverage%'] - a['Coverage%']);
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Summary');
        }
    }
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `mchmuk_${appState.isMophMode ? 'FY' + appState.fiscalYear + '_' : ''}${stamp}.xlsx`);
    showToast(`📗 Export Excel สำเร็จ! ${appState.filteredData.length.toLocaleString()} แถว`, 'success', 4000);
}

// ==========================================================================
// 🔄 Auto-refresh — polls metadata timestamp, reloads only when data changed
// ==========================================================================
const AUTO_REFRESH_KEY = 'mchmuk-autorefresh';
const AUTO_REFRESH_MS = 10 * 60 * 1000;
let _autoRefreshTimer = null;

async function checkForNewData() {
    if (document.hidden) return;
    try {
        const metaUrl = `https://raw.githubusercontent.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}/tmp_exchange_data_meta.json?cb=${Date.now()}`;
        const res = await fetch(metaUrl);
        if (!res.ok) return;
        const meta = await res.json();
        if (meta.importTimestamp && meta.importTimestamp !== appState.importTimestamp) {
            await loadLocalExcelFile(true);
            showToast('🔄 พบข้อมูลใหม่ — อัปเดตแดชบอร์ดแล้ว', 'success', 4000);
        }
    } catch (e) {
        console.warn('Auto-refresh check failed:', e.message);
    }
}

function setAutoRefresh(on) {
    clearInterval(_autoRefreshTimer);
    _autoRefreshTimer = on ? setInterval(checkForNewData, AUTO_REFRESH_MS) : null;
    const btn = document.getElementById('btn-auto-refresh');
    if (btn) {
        btn.classList.toggle('auto-refresh-on', on);
        btn.title = on ? 'Auto-refresh: เปิด (ตรวจทุก 10 นาที) — คลิกเพื่อปิด' : 'Auto-refresh: ปิด — คลิกเพื่อเปิด';
    }
    try { localStorage.setItem(AUTO_REFRESH_KEY, on ? '1' : '0'); } catch (e) {}
}

function toggleAutoRefresh() {
    const on = !_autoRefreshTimer;
    setAutoRefresh(on);
    showToast(on ? '🔄 เปิด Auto-refresh (ทุก 10 นาที)' : '⏸ ปิด Auto-refresh', 'info', 2500);
}

document.addEventListener('DOMContentLoaded', () => {
    let on = false;
    try { on = localStorage.getItem(AUTO_REFRESH_KEY) === '1'; } catch (e) {}
    setAutoRefresh(on);
});

// ==========================================================================
// 📈 Multi-FY overlay — monthly line per fiscal year
// ==========================================================================
function renderTrendMultiFY() {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="layers"></i> เปรียบเทียบรายเดือน ทุกปีงบประมาณ (Overlay)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-3"></i> สะสม ณ เดือนเดียวกัน (YTD)';
    lucide.createIcons();

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const fyMap = {};
    rawRowsForActiveUnit().forEach(r => {
        const d = parseDateValue(r['epi_date']);
        if (!d) return;
        const fy = fiscalYearOf(d);
        const idx = FY_MONTH_ORDER.indexOf(d.getMonth());
        if (!fyMap[fy]) fyMap[fy] = FY_MONTH_ORDER.map(() => ({ total: 0, received: 0, tested: 0, anemic: 0 }));
        const b = fyMap[fy][idx];
        b.total++;
        if (isReceivedValue(r['result'])) b.received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { b.tested++; if (isAnemicRow(r)) b.anemic++; }
    });
    const fys = Object.keys(fyMap).map(Number).sort((a, b) => a - b);
    if (!fys.length) {
        destroyTrendCharts();
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่มีข้อมูลวันที่ (epi_date)</p>';
        return;
    }
    const rate = b => isAnemia ? (b.tested ? b.anemic / b.tested * 100 : null) : (b.total ? b.received / b.total * 100 : null);
    const series = fys.map(fy => ({ name: `ปีงบ ${fy}`, data: fyMap[fy].map(b => { const v = rate(b); return v === null ? null : +v.toFixed(1); }) }));
    const palette = ['#cbd5e1', '#94a3b8', '#f59e0b', '#8b5cf6', '#0ea5e9'];
    const colors = fys.map((fy, i) => fy === appState.fiscalYear ? '#0284c7' : palette[i % palette.length]);

    renderTrendMain(document.getElementById('trend-chart-main'), {
        chart: { type: 'line', height: chartHeight(), background: 'transparent', foreColor: '#64748b', animations: { enabled: false } },
        series,
        xaxis: { categories: FY_MONTH_LABELS },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors,
        stroke: { curve: 'smooth', width: fys.map(fy => fy === appState.fiscalYear ? 3.5 : 1.8) },
        markers: { size: fys.map(fy => fy === appState.fiscalYear ? 4 : 0) },
        annotations: isAnemia ? {} : { yaxis: [{ y: DEFAULT_COVERAGE_TARGET, borderColor: '#16a34a', strokeDashArray: 4, label: { text: `เป้า ${DEFAULT_COVERAGE_TARGET}%`, style: { background: '#16a34a', color: '#fff' } } }] },
        tooltip: { shared: true, y: { formatter: v => v === null || v === undefined ? '-' : v + '%' } },
        grid: { borderColor: 'rgba(100,116,139,0.15)' },
        legend: { position: 'top' }
    });

    // YTD up to the last month that has data in the selected FY, so years compare fairly
    const cur = fyMap[appState.fiscalYear];
    let lastIdx = 11;
    if (cur) { lastIdx = -1; cur.forEach((b, i) => { if (b.total) lastIdx = i; }); if (lastIdx < 0) lastIdx = 11; }
    const ytd = fys.map(fy => {
        const agg = fyMap[fy].slice(0, lastIdx + 1).reduce((a, b) => ({ total: a.total + b.total, received: a.received + b.received, tested: a.tested + b.tested, anemic: a.anemic + b.anemic }), { total: 0, received: 0, tested: 0, anemic: 0 });
        const v = rate(agg);
        return v === null ? 0 : +v.toFixed(1);
    });
    renderTrendSide(document.getElementById('trend-chart-side'), {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: `YTD ถึง ${FY_MONTH_LABELS[lastIdx]}`, data: ytd }],
        xaxis: { categories: fys.map(fy => 'ปีงบ ' + fy) },
        yaxis: { min: 0, max: 100, labels: { formatter: v => v + '%' } },
        colors,
        plotOptions: { bar: { distributed: true, borderRadius: 4 } },
        legend: { show: false },
        dataLabels: { enabled: true, formatter: v => v + '%' },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });
}

// ==========================================================================
// 🔵 Scatter — Coverage vs HCT screening rate per hospital
// ==========================================================================
function renderTrendScatter(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="scatter-chart"></i> Coverage vs อัตราเจาะ HCT รายหน่วยบริการ (คลิกจุดเพื่อกรอง)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="grid-2x2"></i> จำนวนหน่วยบริการรายกลุ่ม';
    lucide.createIcons();

    const { hosnameCol, ampnameCol } = trendGetCols();
    if (!hosnameCol) {
        destroyTrendCharts();
        document.getElementById('trend-chart-main').innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์หน่วยบริการ</p>';
        return;
    }
    const m = {};
    rows.forEach(r => {
        const h = String(r[hosnameCol] || 'ไม่ทราบ');
        if (!m[h]) m[h] = { district: ampnameCol ? String(r[ampnameCol] || 'ไม่ทราบ') : 'ทั้งหมด', total: 0, received: 0, tested: 0 };
        m[h].total++;
        if (isReceivedValue(r['result'])) m[h].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) m[h].tested++;
    });
    const HCT_REF = 50;
    const points = Object.entries(m).filter(([, b]) => b.total >= 3).map(([name, b]) => ({
        name, district: b.district, total: b.total,
        cov: +(b.received / b.total * 100).toFixed(1),
        hct: +(b.tested / b.total * 100).toFixed(1),
        target: getHospitalTarget(name),
    }));
    const byDist = {};
    points.forEach(p => { (byDist[p.district] = byDist[p.district] || []).push(p); });
    const distNames = Object.keys(byDist).sort();
    const series = distNames.map(d => ({ name: d, data: byDist[d].map(p => ({ x: p.cov, y: p.hct, meta: p })) }));

    renderTrendMain(document.getElementById('trend-chart-main'), {
        chart: {
            type: 'scatter', height: chartHeight(), background: 'transparent', foreColor: '#64748b', animations: { enabled: false },
            zoom: { enabled: true, type: 'xy' },
            events: {
                dataPointSelection: (_e, _ctx, cfg) => {
                    const p = series[cfg.seriesIndex]?.data[cfg.dataPointIndex]?.meta;
                    if (!p) return;
                    appState.activeHospitalFilter = p.name;
                    const el = document.getElementById('select-hospital');
                    if (el) el.value = p.name;
                    applyAllFilters();
                    showToast(`🏥 กรองข้อมูล: ${p.name}`, 'info', 2500);
                }
            }
        },
        series,
        colors: TREND_COLORS.concat(neonColors),
        xaxis: { type: 'numeric', min: 0, max: 100, tickAmount: 10, title: { text: 'Coverage % (ได้รับยา)' }, labels: { formatter: v => Math.round(v) + '%' } },
        yaxis: { min: 0, max: 100, tickAmount: 5, title: { text: 'อัตราเจาะ HCT %' }, labels: { formatter: v => Math.round(v) + '%' } },
        markers: { size: 7, strokeWidth: 1, strokeColors: '#fff' },
        annotations: {
            xaxis: [{ x: DEFAULT_COVERAGE_TARGET, borderColor: '#16a34a', strokeDashArray: 4, label: { text: `เป้า ${DEFAULT_COVERAGE_TARGET}%`, style: { background: '#16a34a', color: '#fff' } } }],
            yaxis: [{ y: HCT_REF, borderColor: '#d97706', strokeDashArray: 4, label: { text: `HCT ${HCT_REF}%`, style: { background: '#d97706', color: '#fff' } } }]
        },
        tooltip: {
            custom: ({ seriesIndex, dataPointIndex }) => {
                const p = series[seriesIndex].data[dataPointIndex].meta;
                return `<div style="padding:8px 10px;font-size:0.8rem;"><strong>${escapeHtml(p.name)}</strong><br>${escapeHtml(p.district)} · ${p.total.toLocaleString()} คน<br>Coverage ${p.cov}% (เป้า ${p.target}%)<br>เจาะ HCT ${p.hct}%</div>`;
            }
        },
        legend: { position: 'top' },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });

    const quad = { 'ผ่านทั้งคู่': 0, 'Coverage ผ่าน / HCT ต่ำ': 0, 'HCT ผ่าน / Coverage ต่ำ': 0, 'ต่ำทั้งคู่': 0 };
    points.forEach(p => {
        const c = p.cov >= p.target, h = p.hct >= HCT_REF;
        quad[c && h ? 'ผ่านทั้งคู่' : c ? 'Coverage ผ่าน / HCT ต่ำ' : h ? 'HCT ผ่าน / Coverage ต่ำ' : 'ต่ำทั้งคู่']++;
    });
    renderTrendSide(document.getElementById('trend-chart-side'), {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: 'หน่วยบริการ', data: Object.values(quad) }],
        xaxis: { categories: Object.keys(quad), labels: { style: { fontSize: '10px' } } },
        colors: ['#16a34a', '#f59e0b', '#0ea5e9', '#dc2626'],
        plotOptions: { bar: { distributed: true, horizontal: true, borderRadius: 4 } },
        legend: { show: false },
        dataLabels: { enabled: true },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });
}

// ==========================================================================
// 🗺️ District tile map (choropleth) — approximate positions, Mekong to the east
// ==========================================================================
const MUK_DISTRICT_TILES = [
    { key: 'หนองสูง', col: 0, row: 0, w: 1 },
    { key: 'หว้านใหญ่', col: 2, row: 0, w: 1 },
    { key: 'ดงหลวง', col: 0, row: 1, w: 1 },
    { key: 'เมือง', col: 1, row: 1, w: 2, label: 'เมืองมุกดาหาร' },
    { key: 'คำชะอี', col: 0, row: 2, w: 1 },
    { key: 'นิคมคำสร้อย', col: 1, row: 2, w: 1 },
    { key: 'ดอนตาล', col: 2, row: 2, w: 1 },
];

function renderTrendMap(rows) {
    document.getElementById('trend-chart-main-title').innerHTML = '<i data-lucide="map"></i> แผนที่อำเภอ (ตำแหน่งโดยประมาณ — คลิกเพื่อกรอง)';
    document.getElementById('trend-chart-side-title').innerHTML = '<i data-lucide="bar-chart-3"></i> เรียงลำดับรายอำเภอ';
    lucide.createIcons();
    destroyTrendCharts();

    const { ampnameCol } = trendGetCols();
    const mainEl = document.getElementById('trend-chart-main');
    if (!ampnameCol) { mainEl.innerHTML = '<p style="padding:20px;color:var(--text-muted);">ไม่พบคอลัมน์อำเภอ</p>'; return; }

    const isAnemia = appState.activeMophIndicator === 'anemia-12m';
    const m = {};
    rows.forEach(r => {
        const d = String(r[ampnameCol] || 'ไม่ทราบ');
        if (!m[d]) m[d] = { total: 0, received: 0, tested: 0, anemic: 0 };
        m[d].total++;
        if (isReceivedValue(r['result'])) m[d].received++;
        if (cleanNumericValue(r['lab_result_status']) === 1) { m[d].tested++; if (isAnemicRow(r)) m[d].anemic++; }
    });
    const rate = b => isAnemia ? (b.tested ? b.anemic / b.tested * 100 : null) : (b.total ? b.received / b.total * 100 : null);
    const anemiaTarget = getAnemiaTarget(appState.fiscalYear);
    const colorFor = v => {
        if (v === null) return '#e2e8f0';
        if (isAnemia) return v <= anemiaTarget * 0.6 ? '#86efac' : v <= anemiaTarget ? '#fde68a' : '#fca5a5';
        return v >= DEFAULT_COVERAGE_TARGET ? '#86efac' : v >= 50 ? '#fde68a' : '#fca5a5';
    };

    const used = new Set();
    const tiles = MUK_DISTRICT_TILES.map(t => {
        const name = Object.keys(m).find(d => !used.has(d) && d.includes(t.key));
        if (name) used.add(name);
        return { ...t, name: name || null, b: name ? m[name] : null };
    });
    let extraRow = 3;
    Object.keys(m).filter(d => !used.has(d)).forEach((d, i) => {
        tiles.push({ key: d, label: d, name: d, b: m[d], col: i % 3, row: extraRow + Math.floor(i / 3), w: 1 });
    });
    const rowsCount = Math.max(...tiles.map(t => t.row)) + 1;
    const TW = 140, TH = 86, GAP = 6;
    const W = 3 * TW + 2 * GAP + 30, H = rowsCount * (TH + GAP);
    const svgTiles = tiles.map(t => {
        const v = t.b ? rate(t.b) : null;
        const x = t.col * (TW + GAP), y = t.row * (TH + GAP), w = t.w * TW + (t.w - 1) * GAP;
        const selected = t.name && t.name === appState.activeDistrictFilter;
        const label = escapeHtml(t.label || t.name || t.key);
        return `<g class="map-tile${t.name ? '' : ' empty'}" ${t.name ? `data-dist="${escapeHtml(t.name)}" tabindex="0" role="button" aria-label="${label} ${v === null ? 'ไม่มีข้อมูล' : v.toFixed(1) + '%'}"` : ''}>
            <rect x="${x}" y="${y}" width="${w}" height="${TH}" rx="10" fill="${colorFor(v)}" stroke="${selected ? '#1e293b' : 'rgba(15,23,42,0.15)'}" stroke-width="${selected ? 3 : 1}"/>
            <text x="${x + w / 2}" y="${y + 32}" text-anchor="middle" font-size="13" font-weight="600" fill="#1e293b">${label}</text>
            <text x="${x + w / 2}" y="${y + 58}" text-anchor="middle" font-size="20" font-weight="800" fill="#0f172a">${v === null ? '–' : v.toFixed(1) + '%'}</text>
            ${t.b ? `<text x="${x + w / 2}" y="${y + 76}" text-anchor="middle" font-size="10" fill="#475569">${t.b.total.toLocaleString()} คน</text>` : ''}
        </g>`;
    }).join('');
    const river = `<rect x="${3 * TW + 2 * GAP + 8}" y="0" width="14" height="${3 * (TH + GAP) - GAP}" rx="7" fill="#7dd3fc" opacity="0.6"/>
        <text x="${3 * TW + 2 * GAP + 15}" y="${(3 * (TH + GAP)) / 2}" text-anchor="middle" font-size="10" fill="#0369a1" transform="rotate(90 ${3 * TW + 2 * GAP + 15} ${(3 * (TH + GAP)) / 2})">แม่น้ำโขง</text>`;
    const legend = isAnemia
        ? [['#86efac', `≤${(anemiaTarget * 0.6).toFixed(0)}%`], ['#fde68a', `≤${anemiaTarget}%`], ['#fca5a5', `>${anemiaTarget}%`]]
        : [['#86efac', `≥${DEFAULT_COVERAGE_TARGET}%`], ['#fde68a', '50–74.9%'], ['#fca5a5', '<50%']];

    mainEl.innerHTML = `<div class="district-map">
        <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" style="width:100%;max-height:${chartHeight() - 30}px;">${svgTiles}${river}</svg>
        <div class="district-map-legend">${legend.map(([c, l]) => `<span><i style="background:${c}"></i>${l}</span>`).join('')}<span><i style="background:#e2e8f0"></i>ไม่มีข้อมูล</span></div>
    </div>`;
    mainEl.querySelectorAll('.map-tile[data-dist]').forEach(g => {
        const pick = () => {
            const dist = g.dataset.dist;
            appState.activeDistrictFilter = appState.activeDistrictFilter === dist ? 'all' : dist;
            const el = document.getElementById('select-district');
            if (el) el.value = appState.activeDistrictFilter;
            appState.activeHospitalFilter = 'all';
            applyAllFilters();
        };
        g.addEventListener('click', pick);
        g.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });
    });

    const sorted = Object.entries(m).map(([d, b]) => [d, rate(b)]).filter(([, v]) => v !== null)
        .sort((a, b) => isAnemia ? a[1] - b[1] : b[1] - a[1]);
    renderTrendSide(document.getElementById('trend-chart-side'), {
        chart: { type: 'bar', height: chartHeight(), background: 'transparent', foreColor: '#64748b', toolbar: { show: false }, animations: { enabled: false } },
        series: [{ name: isAnemia ? 'โลหิตจาง%' : 'Coverage%', data: sorted.map(([, v]) => +v.toFixed(1)) }],
        xaxis: { categories: sorted.map(([d]) => d), labels: { style: { fontSize: '10px' } } },
        yaxis: { min: 0, max: 100 },
        colors: sorted.map(([, v]) => colorFor(v).replace('#86efac', '#16a34a').replace('#fde68a', '#d97706').replace('#fca5a5', '#dc2626')),
        plotOptions: { bar: { distributed: true, horizontal: true, borderRadius: 3 } },
        legend: { show: false },
        dataLabels: { enabled: true, formatter: v => v + '%' },
        grid: { borderColor: 'rgba(100,116,139,0.15)' }
    });
}
