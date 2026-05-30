/**
 * 🔐 MCHMUK Authentication Module
 * Multi-user Firebase Auth: login, logout, user management
 */

// ── State ────────────────────────────────────────────────────────────────────
let currentUser     = null;
let currentUserRole = null;
let currentUserData = null;

// ── Init ─────────────────────────────────────────────────────────────────────
function initAuth(onLogin, onLogout) {
    fbAuth.onAuthStateChanged(async (user) => {
        if (user) {
            currentUser     = user;
            currentUserData = await fetchUserProfile(user.uid);
            currentUserRole = currentUserData?.role || 'viewer';
            await fbDb.collection('users').doc(user.uid)
                .update({ lastLogin: firebase.firestore.FieldValue.serverTimestamp() })
                .catch(() => {});
            onLogin(user, currentUserRole, currentUserData);
        } else {
            currentUser = currentUserRole = currentUserData = null;
            onLogout();
        }
    });
}

async function fetchUserProfile(uid) {
    try {
        const doc = await fbDb.collection('users').doc(uid).get();
        return doc.exists ? doc.data() : null;
    } catch { return null; }
}

// ── Login / Logout ────────────────────────────────────────────────────────────
async function login(email, password) {
    try {
        await fbAuth.signInWithEmailAndPassword(email, password);
        return { ok: true };
    } catch (err) {
        const msg = {
            'auth/user-not-found':   'ไม่พบบัญชีผู้ใช้นี้',
            'auth/wrong-password':   'รหัสผ่านไม่ถูกต้อง',
            'auth/invalid-email':    'รูปแบบอีเมลไม่ถูกต้อง',
            'auth/invalid-credential': 'อีเมลหรือรหัสผ่านไม่ถูกต้อง',
            'auth/too-many-requests':'พยายามเข้าสู่ระบบหลายครั้งเกินไป กรุณารอสักครู่'
        };
        return { ok: false, error: msg[err.code] || err.message };
    }
}

async function logout() {
    await fbAuth.signOut().catch(() => {});
}

async function sendPasswordReset(email) {
    try {
        await fbAuth.sendPasswordResetEmail(email);
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
}

async function registerUser(email, password, displayName) {
    try {
        const result = await fbAuth.createUserWithEmailAndPassword(email, password);
        const newUid = result.user.uid;

        // Check if this is the first user in the system to automatically grant Admin role!
        const usersSnap = await fbDb.collection('users').limit(1).get().catch(() => ({ empty: true }));
        const role = usersSnap.empty ? 'admin' : 'viewer';

        await fbDb.collection('users').doc(newUid).set({
            email,
            displayName: displayName || email.split('@')[0],
            role: role,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        });

        // Set the current user's profile display name
        if (result.user) {
            await result.user.updateProfile({
                displayName: displayName || email.split('@')[0]
            }).catch(() => {});
        }

        return { ok: true, role };
    } catch (err) {
        const msg = {
            'auth/email-already-in-use': 'อีเมลนี้มีผู้ใช้งานแล้ว',
            'auth/weak-password': 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร',
            'auth/invalid-email': 'รูปแบบอีเมลไม่ถูกต้อง'
        };
        return { ok: false, error: msg[err.code] || err.message };
    }
}

// ── Admin: User Management ────────────────────────────────────────────────────
async function adminCreateUser(email, password, displayName, role, hosCode = '') {
    if (currentUserRole !== 'admin') return { ok: false, error: 'ไม่มีสิทธิ์' };

    // ใช้ secondary app เพื่อไม่ให้ sign out admin ออก
    const secondary = firebase.initializeApp(firebase.app().options, 'secondary_' + Date.now());
    try {
        const result = await secondary.auth().createUserWithEmailAndPassword(email, password);
        const newUid = result.user.uid;
        await secondary.auth().signOut();

        await fbDb.collection('users').doc(newUid).set({
            email, displayName, role, hosCode,
            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
            createdBy: currentUser.uid
        });
        return { ok: true, uid: newUid };
    } catch (err) {
        const msg = {
            'auth/email-already-in-use': 'อีเมลนี้มีผู้ใช้งานแล้ว',
            'auth/weak-password': 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร'
        };
        return { ok: false, error: msg[err.code] || err.message };
    } finally {
        secondary.delete().catch(() => {});
    }
}

async function adminListUsers() {
    if (currentUserRole !== 'admin') return [];
    try {
        const snap = await fbDb.collection('users').orderBy('createdAt', 'desc').get();
        return snap.docs.map(d => ({ uid: d.id, ...d.data() }));
    } catch { return []; }
}

async function adminUpdateUserRole(uid, role) {
    if (currentUserRole !== 'admin') return { ok: false, error: 'ไม่มีสิทธิ์' };
    try {
        await fbDb.collection('users').doc(uid).update({ role });
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
}

async function adminDeleteUser(uid) {
    if (currentUserRole !== 'admin') return { ok: false, error: 'ไม่มีสิทธิ์' };
    if (uid === currentUser.uid) return { ok: false, error: 'ไม่สามารถลบตัวเองได้' };
    try {
        await fbDb.collection('users').doc(uid).delete();
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function isAdmin()  { return currentUserRole === 'admin'; }
function isLoggedIn() { return currentUser !== null; }
