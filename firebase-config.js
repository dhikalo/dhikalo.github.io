/* ============================================
   FIREBASE CLOUD SYNC — Messstellen Manager
   Real-time sync + Prüfung workflow
   ============================================ */

'use strict';

// ─── FIREBASE CONFIGURATION ───
// IMPORTANT: Replace these with your own Firebase project config
// Go to: https://console.firebase.google.com → Create project → Web app → Copy config
const FIREBASE_CONFIG = {
    apiKey: "AIzaSyAkQMfyhiqrlQCZ6BC4DfaEC-Nt7NOXAH4",
    authDomain: "messtellen-manager.firebaseapp.com",
    projectId: "messtellen-manager",
    storageBucket: "messtellen-manager.firebasestorage.app",
    messagingSenderId: "842219144600",
    appId: "1:842219144600:web:23b7172e0142bc5854d7ed",
    measurementId: "G-NYS834RJF6"
};

// ─── CLOUD STATE ───
const CloudState = {
    initialized: false,
    db: null,
    auth: null,
    storage: null,
    role: localStorage.getItem('messstellen_role') || null, // 'messhelfer' or 'pruefer'
    userId: localStorage.getItem('messstellen_userId') || null,
    userName: localStorage.getItem('messstellen_userName') || '',
    syncStatus: 'offline', // 'synced', 'syncing', 'offline', 'error'
    unsubscribers: [], // Firestore listeners to clean up
    lastSyncTime: null,
    pendingWrites: 0,
    sessionId: Math.random().toString(36).substring(2) + Date.now().toString(36)
};

// ─── REVIEW STATUS CONSTANTS ───
const REVIEW_STATUS = {
    DRAFT: 'draft',          // Messhelfer still working
    SUBMITTED: 'submitted',  // Submitted for review
    APPROVED: 'approved',    // Prüfer approved
    REJECTED: 'rejected'     // Prüfer rejected with reason
};

// ─── INITIALIZE FIREBASE ───
function initFirebase() {
    try {
        // Check if config has been set
        if (FIREBASE_CONFIG.apiKey === "PASTE_YOUR_API_KEY_HERE") {
            console.warn('Firebase not configured — running in local-only mode');
            updateCloudStatus('offline');
            return false;
        }

        // Initialize Firebase app
        if (!firebase.apps || firebase.apps.length === 0) {
            firebase.initializeApp(FIREBASE_CONFIG);
        }

        // Initialize Firestore with offline persistence
        CloudState.db = firebase.firestore();
        CloudState.db.enablePersistence({ synchronizeTabs: true }).catch(err => {
            if (err.code === 'failed-precondition') {
                console.warn('Firestore persistence: multiple tabs open');
            } else if (err.code === 'unimplemented') {
                console.warn('Firestore persistence: browser not supported');
            }
        });

        // Initialize Auth
        CloudState.auth = firebase.auth();

        // Initialize Storage for photos
        CloudState.storage = firebase.storage();

        CloudState.initialized = true;
        updateCloudStatus('synced');
        console.log('Firebase initialized successfully');

        // Auto sign-in anonymously if no user
        if (!CloudState.auth.currentUser) {
            signInAnonymously();
        }

        return true;
    } catch (err) {
        console.error('Firebase init error:', err);
        updateCloudStatus('error');
        return false;
    }
}

// ─── AUTHENTICATION ───
async function signInAnonymously() {
    try {
        const result = await CloudState.auth.signInAnonymously();
        CloudState.userId = result.user.uid;
        localStorage.setItem('messstellen_userId', CloudState.userId);
        console.log('Signed in anonymously:', CloudState.userId);
    } catch (err) {
        console.error('Auth error:', err);
    }
}

function setUserRole(role, name) {
    CloudState.role = role;
    CloudState.userName = name || (role === 'pruefer' ? 'Prüfer' : 'Messhelfer');
    localStorage.setItem('messstellen_role', role);
    localStorage.setItem('messstellen_userName', CloudState.userName);
}

function getUserRole() {
    return CloudState.role;
}

// ─── CLOUD STATUS UI ───
function updateCloudStatus(status) {
    CloudState.syncStatus = status;
    const badge = document.getElementById('cloudStatusBadge');
    if (!badge) return;

    const icons = {
        synced: '<i class="fas fa-cloud" style="color:#10b981;"></i>',
        syncing: '<i class="fas fa-cloud-upload-alt fa-spin" style="color:#fbbf24;"></i>',
        offline: '<i class="fas fa-cloud-slash" style="color:#64748b;"></i>',
        error: '<i class="fas fa-exclamation-triangle" style="color:#ef4444;"></i>'
    };
    const labels = {
        synced: 'Cloud Sync',
        syncing: 'Syncing...',
        offline: 'Lokal',
        error: 'Sync-Fehler'
    };

    badge.innerHTML = `${icons[status] || icons.offline} <span class="cloud-label">${labels[status] || 'Offline'}</span>`;
    badge.className = 'cloud-status-badge cloud-' + status;
}

// ─── SAVE TO CLOUD ───
async function saveToCloud(projectName, projectData) {
    if (!CloudState.initialized || !CloudState.db) return;
    if (!projectName || projectName === 'Unbenanntes Projekt') return;

    updateCloudStatus('syncing');
    CloudState.pendingWrites++;

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));

        // Handle images: upload to Firebase Storage, replace base64 with URL
        const cloudData = await Promise.all(projectData.data.map(async (row, rowIdx) => {
            const cleanRow = {};
            for (const [key, val] of Object.entries(row)) {
                // Handle multi-photo arrays (e.g. MK-Bild) — each entry treated the same as a single-photo cell
                if (Array.isArray(val)) {
                    const outArr = [];
                    for (let i = 0; i < val.length; i++) {
                        const item = val[i];
                        if (typeof item === 'string' && item.startsWith('data:image') && item.length > 50000) {
                            try {
                                const url = await uploadPhotoToStorage(projectName, rowIdx, key + '#' + i, item);
                                outArr.push(url || '__IMAGE_REF__');
                            } catch (e) {
                                outArr.push('__IMAGE_REF__');
                            }
                        } else if (typeof item === 'string' && item.startsWith('https://firebasestorage')) {
                            outArr.push(item);
                        } else if (item === '__IDB_PHOTO__') {
                            outArr.push('__IMAGE_REF__');
                        } else if (typeof item === 'string' && item.startsWith('data:image')) {
                            // Small inline image — still upload to Storage to keep Firestore doc lean
                            try {
                                const url = await uploadPhotoToStorage(projectName, rowIdx, key + '#' + i, item);
                                outArr.push(url || '__IMAGE_REF__');
                            } catch (e) {
                                outArr.push('__IMAGE_REF__');
                            }
                        } else if (item) {
                            outArr.push(item);
                        }
                    }
                    cleanRow[key] = outArr;
                } else if (typeof val === 'string' && val.startsWith('data:image') && val.length > 50000) {
                    // Upload to Storage and store the download URL instead
                    try {
                        const url = await uploadPhotoToStorage(projectName, rowIdx, key, val);
                        cleanRow[key] = url || '__IMAGE_REF__';
                        // Store URL back in memory so subsequent saves don't re-upload
                        if (url) {
                            row[key] = url;
                        }
                    } catch (e) {
                        cleanRow[key] = '__IMAGE_REF__';
                    }
                } else if (typeof val === 'string' && val.startsWith('https://firebasestorage')) {
                    // Already a Storage URL — keep as-is
                    cleanRow[key] = val;
                } else if (val === '__IDB_PHOTO__') {
                    // Local IndexedDB placeholder — don't send to cloud
                    cleanRow[key] = '__IMAGE_REF__';
                } else if (typeof val === 'string' && val.startsWith('data:image')) {
                    // Small image (< 50KB) — skip sending inline to Firestore to avoid bloat
                    cleanRow[key] = '__IMAGE_REF__';
                } else {
                    cleanRow[key] = val;
                }
            }
            return cleanRow;
        }));

        // Split data into chunks if it's very large (Firestore 1MB limit per doc)
        const payload = {
            name: projectName,
            data: cloudData,
            mapData: projectData.mapData || null,
            starData: projectData.starData || [],
            hiddenMapColors: projectData.hiddenMapColors || [],
            newCols: projectData.newCols || [],
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: CloudState.userName || 'Unbekannt',
            updatedByRole: CloudState.role || 'unknown',
            userId: CloudState.userId,
            sessionId: CloudState.sessionId,
            reviewStatus: projectData.reviewStatus || REVIEW_STATUS.DRAFT,
            reviewComment: projectData.reviewComment || '',
            reviewedBy: projectData.reviewedBy || '',
            reviewedAt: projectData.reviewedAt || null,
            submittedAt: projectData.submittedAt || null,
            rowCount: (projectData.data || []).length
        };

        await docRef.set(payload, { merge: true });

        CloudState.lastSyncTime = Date.now();
        CloudState.pendingWrites--;
        if (CloudState.pendingWrites <= 0) {
            CloudState.pendingWrites = 0;
            updateCloudStatus('synced');
        }

    } catch (err) {
        console.error('Cloud save error:', err);
        CloudState.pendingWrites--;
        updateCloudStatus('error');

        if (typeof showToast === 'function') {
            // Firestore documents are capped at ~1MB. Photos are already offloaded
            // to Storage above, so this almost always means the project itself
            // (many rows / many custom columns of text) has grown too large for a
            // single document. Give a clear, actionable message instead of the
            // raw Firestore error, and keep working locally — saveToStorage()
            // already wrote everything to localStorage regardless of this failure.
            const isTooLarge = err.code === 'invalid-argument' ||
                (err.message && (err.message.includes('longer than') || err.message.includes('exceeds the maximum')));
            if (isTooLarge) {
                showToast('⚠️ Projekt zu groß für Cloud-Sync (Firestore-Limit 1MB). Daten sind lokal gespeichert, aber nicht in der Cloud. Bitte Projekt aufteilen oder Fotos prüfen.', 6000);
            } else {
                showToast('Cloud-Sync fehlgeschlagen: ' + err.message);
            }
        }
    }
}

// ─── LOAD FROM CLOUD ───
async function loadFromCloud(projectName) {
    if (!CloudState.initialized || !CloudState.db) return null;

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));
        const doc = await docRef.get();

        if (doc.exists) {
            return doc.data();
        }
        return null;
    } catch (err) {
        console.error('Cloud load error:', err);
        return null;
    }
}

// ─── REAL-TIME LISTENER ───
function listenForCloudUpdates(projectName, callback) {
    if (!CloudState.initialized || !CloudState.db) return null;

    const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));

    const unsubscribe = docRef.onSnapshot(doc => {
        if (doc.exists) {
            const data = doc.data();
            
            // 1. Ignore updates from our own session to prevent cursor jumping/reverts while typing
            if (data.sessionId === CloudState.sessionId) {
                return;
            }
            
            // 2. Ignore local writes (metadata changes) to prevent feedback loops
            if (doc.metadata && doc.metadata.hasPendingWrites) {
                return;
            }

            // Only update if the change came from someone else
            if (data.userId !== CloudState.userId ||
                (data.updatedByRole === 'pruefer' && CloudState.role === 'messhelfer') ||
                (data.updatedByRole === 'messhelfer' && CloudState.role === 'pruefer')) {
                if (callback) callback(data);
            }
        }
    }, err => {
        console.error('Snapshot error:', err);
    });

    CloudState.unsubscribers.push(unsubscribe);
    return unsubscribe;
}

// ─── LISTEN FOR ALL PROJECTS (PRÜFER VIEW) ───
function listenForAllProjects(callback) {
    if (!CloudState.initialized || !CloudState.db) return null;

    const query = CloudState.db.collection('projects')
        .orderBy('updatedAt', 'desc');

    const unsubscribe = query.onSnapshot(snapshot => {
        const projects = [];
        snapshot.forEach(doc => {
            projects.push({ id: doc.id, ...doc.data() });
        });
        if (callback) callback(projects);
    }, err => {
        console.error('Projects listener error:', err);
    });

    CloudState.unsubscribers.push(unsubscribe);
    return unsubscribe;
}

// ─── PRÜFUNG WORKFLOW ───
async function submitForReview(projectName) {
    if (!CloudState.initialized || !CloudState.db) {
        if (typeof showToast === 'function') showToast('Cloud nicht verfügbar — bitte Verbindung prüfen');
        return false;
    }

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));
        // Use set+merge so it works even if the document doesn't exist yet
        await docRef.set({
            reviewStatus: REVIEW_STATUS.SUBMITTED,
            submittedAt: firebase.firestore.FieldValue.serverTimestamp(),
            submittedBy: CloudState.userName,
            reviewComment: '',
            reviewedBy: '',
            name: projectName,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: CloudState.userName || 'Messhelfer',
            updatedByRole: CloudState.role || 'messhelfer',
            sessionId: CloudState.sessionId
        }, { merge: true });

        if (typeof showToast === 'function') showToast('✅ Zur Prüfung eingereicht!');
        return true;
    } catch (err) {
        console.error('Submit for review error:', err);
        if (typeof showToast === 'function') showToast('Einreichung fehlgeschlagen: ' + err.message);
        return false;
    }
}

async function approveProject(projectName) {
    if (!CloudState.initialized || !CloudState.db) return false;

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));
        await docRef.update({
            reviewStatus: REVIEW_STATUS.APPROVED,
            reviewedAt: firebase.firestore.FieldValue.serverTimestamp(),
            reviewedBy: CloudState.userName || 'Prüfer',
            reviewComment: '',
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: CloudState.userName || 'Prüfer',
            updatedByRole: CloudState.role || 'pruefer',
            sessionId: CloudState.sessionId
        });

        if (typeof showToast === 'function') showToast('✅ Projekt genehmigt!');
        return true;
    } catch (err) {
        console.error('Approve error:', err);
        return false;
    }
}

async function rejectProject(projectName, reason) {
    if (!CloudState.initialized || !CloudState.db) return false;

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));
        await docRef.update({
            reviewStatus: REVIEW_STATUS.REJECTED,
            reviewedAt: firebase.firestore.FieldValue.serverTimestamp(),
            reviewedBy: CloudState.userName || 'Prüfer',
            reviewComment: reason || 'Kein Grund angegeben',
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: CloudState.userName || 'Prüfer',
            updatedByRole: CloudState.role || 'pruefer',
            sessionId: CloudState.sessionId
        });

        if (typeof showToast === 'function') showToast('❌ Projekt abgelehnt');
        return true;
    } catch (err) {
        console.error('Reject error:', err);
        return false;
    }
}

async function reopenProject(projectName) {
    if (!CloudState.initialized || !CloudState.db) return false;

    try {
        const docRef = CloudState.db.collection('projects').doc(sanitizeDocId(projectName));
        await docRef.update({
            reviewStatus: REVIEW_STATUS.DRAFT,
            reviewComment: '',
            reviewedBy: '',
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: CloudState.userName || 'Messhelfer',
            updatedByRole: CloudState.role || 'messhelfer',
            sessionId: CloudState.sessionId
        });

        if (typeof showToast === 'function') showToast('📝 Projekt wieder geöffnet');
        return true;
    } catch (err) {
        console.error('Reopen error:', err);
        return false;
    }
}

// ─── GET ALL CLOUD PROJECTS ───
async function getAllCloudProjects() {
    if (!CloudState.initialized || !CloudState.db) return [];

    try {
        const snapshot = await CloudState.db.collection('projects')
            .orderBy('updatedAt', 'desc')
            .get();

        const projects = [];
        snapshot.forEach(doc => {
            projects.push({ id: doc.id, ...doc.data() });
        });
        return projects;
    } catch (err) {
        console.error('Get all projects error:', err);
        return [];
    }
}

// ─── UPLOAD PHOTO TO FIREBASE STORAGE ───
// Called when saving an image (base64) — stores it in Storage and returns the download URL
async function uploadPhotoToStorage(projectName, rowIndex, colName, base64DataUrl) {
    if (!CloudState.initialized || !CloudState.storage) return null;
    if (!base64DataUrl || !base64DataUrl.startsWith('data:image')) return null;

    try {
        // Convert base64 to Blob
        const response = await fetch(base64DataUrl);
        const blob = await response.blob();

        // Build a STABLE storage path so re-uploads overwrite the same file
        // instead of creating duplicates (old path had Date.now() → new file each save)
        const ext = blob.type.includes('png') ? 'png' : 'jpg';
        const safeName = sanitizeDocId(projectName);
        const safeCol = colName.replace(/[^a-zA-Z0-9_]/g, '_');
        const path = `projects/${safeName}/images/${rowIndex}_${safeCol}.${ext}`;

        const storageRef = CloudState.storage.ref(path);

        // Show upload progress indicator
        const progressEl = showUploadProgress();

        const uploadTask = storageRef.put(blob);
        return await new Promise((resolve, reject) => {
            uploadTask.on('state_changed',
                (snapshot) => {
                    const pct = Math.round((snapshot.bytesTransferred / snapshot.totalBytes) * 100);
                    updateUploadProgress(progressEl, pct);
                },
                (err) => {
                    hideUploadProgress(progressEl);
                    console.error('Photo upload error:', err);
                    reject(err);
                },
                async () => {
                    hideUploadProgress(progressEl);
                    const url = await uploadTask.snapshot.ref.getDownloadURL();
                    resolve(url);
                }
            );
        });
    } catch (err) {
        console.error('uploadPhotoToStorage error:', err);
        return null;
    }
}

// Upload progress UI helpers
function showUploadProgress() {
    const el = document.createElement('div');
    el.className = 'photo-upload-progress';
    el.innerHTML = `<i class="fas fa-cloud-upload-alt" style="color:var(--accent);"></i>
        <span>Foto hochladen...</span>
        <div class="photo-upload-bar"><div class="photo-upload-fill" style="width:0%"></div></div>`;
    document.body.appendChild(el);
    return el;
}
function updateUploadProgress(el, pct) {
    if (!el) return;
    const fill = el.querySelector('.photo-upload-fill');
    if (fill) fill.style.width = pct + '%';
}
function hideUploadProgress(el) {
    if (!el) return;
    el.style.opacity = '0';
    el.style.transition = 'opacity 0.4s';
    setTimeout(() => { if (el.parentNode) el.parentNode.removeChild(el); }, 400);
}

// ─── DELETE FROM CLOUD ───
async function deleteFromCloud(projectName) {
    if (!CloudState.initialized || !CloudState.db) return;

    try {
        await CloudState.db.collection('projects').doc(sanitizeDocId(projectName)).delete();
    } catch (err) {
        console.error('Cloud delete error:', err);
    }
}

// ─── UTILITIES ───
function sanitizeDocId(name) {
    // Firestore document IDs cannot contain: / . .. or be empty
    return name.replace(/[\/\.#\$\[\]]/g, '_').trim() || 'unnamed';
}

function cleanupListeners() {
    CloudState.unsubscribers.forEach(unsub => {
        if (typeof unsub === 'function') unsub();
    });
    CloudState.unsubscribers = [];
}

// ─── CHECK IF FIREBASE IS CONFIGURED ───
function isFirebaseConfigured() {
    return FIREBASE_CONFIG.apiKey !== "PASTE_YOUR_API_KEY_HERE" && 
           FIREBASE_CONFIG.apiKey !== "" && 
           CloudState.initialized;
}

// Export for use in app.js
window.CloudState = CloudState;
// ─── CLEANUP DUPLICATE PHOTOS IN FIREBASE STORAGE ───
// Old upload code used Date.now() in filenames, creating a new file on every save.
// This function finds all timestamped duplicates, keeps the newest per group, and deletes the rest.
async function cleanupDuplicatePhotos() {
    if (!CloudState.initialized || !CloudState.storage) {
        if (typeof showToast === 'function') showToast('❌ Firebase Storage nicht initialisiert');
        console.error('Firebase Storage not initialized');
        return 0;
    }

    if (typeof showToast === 'function') showToast('🔍 Suche nach doppelten Fotos…');
    console.log('Starting duplicate photo cleanup...');

    const rootRef = CloudState.storage.ref('projects');
    let projectFolders;
    try {
        projectFolders = await rootRef.listAll();
    } catch (e) {
        console.error('Cannot list project folders:', e);
        if (typeof showToast === 'function') showToast('❌ Fehler beim Lesen der Ordner');
        return 0;
    }

    let totalDeleted = 0;
    let totalKept = 0;

    // Regex: matches filenames like "0_Anhang_Global_1785151118046.jpg"
    // Captures: [1] = base name (e.g. "0_Anhang_Global"), [2] = timestamp, [3] = extension
    const timestampPattern = /^(.+)_(\d{10,14})\.(jpg|png)$/;

    for (const projectFolder of projectFolders.prefixes) {
        const imagesRef = projectFolder.child('images');
        let imageFiles;
        try {
            imageFiles = await imagesRef.listAll();
        } catch (e) {
            continue; // No images folder for this project
        }

        if (imageFiles.items.length === 0) continue;

        // Group timestamped files by their base name
        const groups = {};
        let stableFiles = 0;

        for (const fileRef of imageFiles.items) {
            const name = fileRef.name;
            const match = name.match(timestampPattern);

            if (match) {
                const baseName = match[1];
                const timestamp = parseInt(match[2]);
                const ext = match[3];
                const groupKey = baseName + '.' + ext;

                if (!groups[groupKey]) groups[groupKey] = [];
                groups[groupKey].push({ ref: fileRef, timestamp, name });
            } else {
                // Stable-format file (no timestamp) — always keep
                stableFiles++;
            }
        }

        console.log(`Project "${projectFolder.name}": ${imageFiles.items.length} files (${stableFiles} stable, ${Object.keys(groups).length} groups with timestamps)`);

        // For each group: keep newest, delete all older duplicates
        for (const [groupKey, files] of Object.entries(groups)) {
            if (files.length <= 1) {
                totalKept++;
                continue;
            }

            // Sort newest first
            files.sort((a, b) => b.timestamp - a.timestamp);

            // Keep the first (newest)
            totalKept++;
            console.log(`  Keeping: ${files[0].name} (newest of ${files.length})`);

            // Delete the rest
            for (let i = 1; i < files.length; i++) {
                try {
                    await files[i].ref.delete();
                    totalDeleted++;
                    console.log(`  Deleted: ${files[i].name}`);
                } catch (e) {
                    console.warn(`  Failed to delete ${files[i].name}:`, e.message || e);
                }
            }
        }
    }

    const msg = `✅ Aufräumen fertig: ${totalDeleted} Duplikate gelöscht, ${totalKept} Dateien behalten`;
    console.log(msg);
    if (typeof showToast === 'function') showToast(msg, 5000);
    return totalDeleted;
}

window.REVIEW_STATUS = REVIEW_STATUS;
window.initFirebase = initFirebase;
window.saveToCloud = saveToCloud;
window.loadFromCloud = loadFromCloud;
window.listenForCloudUpdates = listenForCloudUpdates;
window.listenForAllProjects = listenForAllProjects;
window.submitForReview = submitForReview;
window.approveProject = approveProject;
window.rejectProject = rejectProject;
window.reopenProject = reopenProject;
window.getAllCloudProjects = getAllCloudProjects;
window.deleteFromCloud = deleteFromCloud;
window.setUserRole = setUserRole;
window.getUserRole = getUserRole;
window.isFirebaseConfigured = isFirebaseConfigured;
window.updateCloudStatus = updateCloudStatus;
window.cleanupListeners = cleanupListeners;
window.uploadPhotoToStorage = uploadPhotoToStorage;
window.cleanupDuplicatePhotos = cleanupDuplicatePhotos;
