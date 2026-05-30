/**
 * 🗄️ MCHMUK Data Module
 * Upload/Download patient batches via Firebase Storage + Firestore metadata
 */

// ── Upload Batch (Admin) ──────────────────────────────────────────────────────
async function dbUploadBatch(records, filename, period = '') {
    if (!isAdmin()) return { ok: false, error: 'ไม่มีสิทธิ์อัปโหลด' };

    const batchId   = `batch_${Date.now()}`;
    const storePath = `batches/${batchId}.json`;

    try {
        showToast(`⬆️ กำลังอัปโหลด ${records.length.toLocaleString()} ราย...`, 'info', 8000);

        // 1. Upload JSON → Firebase Storage
        const blob = new Blob([JSON.stringify(records)], { type: 'application/json' });
        await fbStorage.ref(storePath).put(blob);

        // 2. Save metadata → Firestore
        await fbDb.collection('batches').doc(batchId).set({
            name:            period || filename,
            filename,
            period,
            recordCount:     records.length,
            fileSize:        blob.size,
            storagePath:     storePath,
            isActive:        true,
            uploadedBy:      currentUser.uid,
            uploadedByEmail: currentUser.email,
            uploadedAt:      firebase.firestore.FieldValue.serverTimestamp()
        });

        showToast(`✅ อัปโหลดสำเร็จ! ${records.length.toLocaleString()} รายการ`, 'success', 5000);
        return { ok: true, batchId };

    } catch (err) {
        showToast(`❌ อัปโหลดล้มเหลว: ${err.message}`, 'error', 6000);
        return { ok: false, error: err.message };
    }
}

// ── List Batches ──────────────────────────────────────────────────────────────
async function dbListBatches() {
    try {
        const snap = await fbDb.collection('batches')
            .where('isActive', '==', true)
            .orderBy('uploadedAt', 'desc')
            .get();
        return snap.docs.map(d => ({ id: d.id, ...d.data() }));
    } catch { return []; }
}

// ── Download Batch Data ───────────────────────────────────────────────────────
async function dbDownloadBatch(batchId) {
    try {
        const meta = await fbDb.collection('batches').doc(batchId).get();
        if (!meta.exists) return { ok: false, error: 'ไม่พบชุดข้อมูล' };

        const { storagePath, recordCount, filename } = meta.data();
        showToast(`⬇️ กำลังโหลดข้อมูล ${(recordCount || 0).toLocaleString()} ราย...`, 'info', 5000);

        const url  = await fbStorage.ref(storagePath).getDownloadURL();
        const resp = await fetch(url);
        const records = await resp.json();

        return { ok: true, records, metadata: meta.data() };
    } catch (err) {
        showToast(`❌ โหลดข้อมูลล้มเหลว: ${err.message}`, 'error', 6000);
        return { ok: false, error: err.message };
    }
}

// ── Get Latest Active Batch ───────────────────────────────────────────────────
async function dbGetActiveBatch() {
    try {
        const snap = await fbDb.collection('batches')
            .where('isActive', '==', true)
            .orderBy('uploadedAt', 'desc')
            .limit(1)
            .get();
        if (snap.empty) return null;
        const d = snap.docs[0];
        return { id: d.id, ...d.data() };
    } catch { return null; }
}

// ── Delete Batch (Admin) ──────────────────────────────────────────────────────
async function dbDeleteBatch(batchId) {
    if (!isAdmin()) return { ok: false, error: 'ไม่มีสิทธิ์' };
    try {
        const doc = await fbDb.collection('batches').doc(batchId).get();
        if (doc.exists) {
            await fbStorage.ref(doc.data().storagePath).delete().catch(() => {});
            await fbDb.collection('batches').doc(batchId).delete();
        }
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
}

// ── Set Active Batch ──────────────────────────────────────────────────────────
async function dbSetActiveBatch(batchId) {
    if (!isAdmin()) return { ok: false, error: 'ไม่มีสิทธิ์' };
    try {
        // Deactivate all batches
        const all = await fbDb.collection('batches').where('isActive','==',true).get();
        const batch = fbDb.batch();
        all.docs.forEach(d => batch.update(d.ref, { isActive: false }));
        batch.update(fbDb.collection('batches').doc(batchId), { isActive: true });
        await batch.commit();
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
}
