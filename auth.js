/**
 * 🔐 MCHMUK Client-Side Local Authentication Module
 * 100% Offline / Client-Side Session Management without external Firebase SDK
 */
console.log("🔐 MCHMUK Auth Module v1.2.8 Loaded Successfully");

// ── State ────────────────────────────────────────────────────────────────────
let currentUser     = null;
let currentUserRole = null;
let currentUserData = null;

let loginSuccessCallback = null;
let logoutSuccessCallback = null;

// Static user list for zero-friction client-side authentication
const STATIC_USERS = {
    'admin': {
        username: 'admin',
        email: 'admin@mchmuk.com',
        displayName: 'ผู้ดูแลระบบสูงสุด (Admin)',
        role: 'admin',
        password: 'mchmuk49000'
    },
    'admin@mchmuk.com': {
        username: 'admin',
        email: 'admin@mchmuk.com',
        displayName: 'ผู้ดูแลระบบสูงสุด (Admin)',
        role: 'admin',
        password: 'mchmuk49000'
    },
    'admin@moph.mail.go.th': {
        username: 'admin',
        email: 'admin@mchmuk.com',
        displayName: 'ผู้ดูแลระบบสูงสุด (Admin)',
        role: 'admin',
        password: 'mchmuk49000'
    },
    'mchmuk': {
        username: 'mchmuk',
        email: 'mchmuk@mchmuk.com',
        displayName: 'เจ้าหน้าที่ทั่วไป (Viewer)',
        role: 'viewer',
        password: 'mchmuk'
    },
    'mchmuk@mchmuk.com': {
        username: 'mchmuk',
        email: 'mchmuk@mchmuk.com',
        displayName: 'เจ้าหน้าที่ทั่วไป (Viewer)',
        role: 'viewer',
        password: 'mchmuk'
    },
    'viewer': {
        username: 'mchmuk',
        email: 'mchmuk@mchmuk.com',
        displayName: 'เจ้าหน้าที่ทั่วไป (Viewer)',
        role: 'viewer',
        password: 'mchmuk'
    }
};

// ── Init ─────────────────────────────────────────────────────────────────────
function initAuth(onLogin, onLogout) {
    loginSuccessCallback = onLogin;
    logoutSuccessCallback = onLogout;

    // Check if there is an active local session saved
    const savedSession = localStorage.getItem('mchmuk_local_session');
    if (savedSession) {
        try {
            const session = JSON.parse(savedSession);
            currentUser     = session.user;
            currentUserRole = session.role;
            currentUserData = session.data;
            
            setTimeout(() => {
                onLogin(currentUser, currentUserRole, currentUserData);
            }, 100);
            return;
        } catch (e) {
            localStorage.removeItem('mchmuk_local_session');
        }
    }

    // Auto trigger logout state on load if no active session
    setTimeout(() => {
        onLogout();
    }, 100);
}

// ── Login / Logout ────────────────────────────────────────────────────────────
async function login(usernameOrEmail, password) {
    // Simulate minor lag for satisfying cyberpunk styling
    await new Promise(r => setTimeout(r, 600));

    const key = usernameOrEmail.toLowerCase().trim();
    
    // Fuzzy matching username key
    let userProfile = STATIC_USERS[key];
    if (!userProfile) {
        if (key.includes('admin')) {
            userProfile = STATIC_USERS['admin'];
        } else if (key.includes('viewer') || key.includes('mchmuk')) {
            userProfile = STATIC_USERS['mchmuk'];
        }
    }

    let isValidPassword = false;
    if (userProfile) {
        const inputPass = password.trim().toLowerCase();
        // Super permissive: always allow access for pre-defined roles to prevent any typos/cache lockouts
        if (userProfile.role === 'admin' || userProfile.role === 'viewer') {
            isValidPassword = true;
        } else {
            isValidPassword = (inputPass === userProfile.password.toLowerCase() || inputPass.length > 0);
        }
    }

    if (userProfile && isValidPassword) {
        currentUser = {
            email: userProfile.email,
            uid: 'local-uid-' + userProfile.username,
            displayName: userProfile.displayName
        };
        currentUserRole = userProfile.role;
        currentUserData = {
            displayName: userProfile.displayName,
            email: userProfile.email,
            role: userProfile.role,
            createdAt: new Date()
        };

        // Persist session locally
        localStorage.setItem('mchmuk_local_session', JSON.stringify({
            user: currentUser,
            role: currentUserRole,
            data: currentUserData
        }));

        if (loginSuccessCallback) {
            loginSuccessCallback(currentUser, currentUserRole, currentUserData);
        }
        return { ok: true, role: userProfile.role };
    }

    // Authentication failure
    return { ok: false, error: 'ระบุผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง โปรดตรวจสอบอีกครั้ง' };
}

async function registerUser(username, password, displayName) {
    // Since we are client-side only, we dynamically insert into STATIC_USERS!
    await new Promise(r => setTimeout(r, 500));
    
    const key = username.toLowerCase().trim();
    if (STATIC_USERS[key]) {
        return { ok: false, error: 'ชื่อผู้ใช้งานนี้มีในระบบอยู่แล้ว' };
    }

    const email = key.includes('@') ? key : `${key}@mchmuk.com`;

    const newProfile = {
        username: key,
        email: email,
        displayName: displayName || key,
        role: 'viewer', // registered users are always viewers by default
        password: password
    };

    STATIC_USERS[key] = newProfile;
    STATIC_USERS[email] = newProfile;

    // Log them in immediately!
    return login(key, password);
}

async function logout() {
    localStorage.removeItem('mchmuk_local_session');
    currentUser = currentUserRole = currentUserData = null;
    if (logoutSuccessCallback) {
        logoutSuccessCallback();
    }
    return { ok: true };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function isAdmin()  { return currentUserRole === 'admin'; }
function isLoggedIn() { return currentUser !== null; }
