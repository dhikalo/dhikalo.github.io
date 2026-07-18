/* ============================================
   PHOTO STORE — IndexedDB-based image storage
   Overcomes the ~5 MB localStorage quota limit
   by keeping photos in IndexedDB (100s of MB).
   ============================================ */

'use strict';

const PhotoStore = (function () {
    const DB_NAME = 'messstellen_photos';
    const DB_VERSION = 1;
    const STORE_NAME = 'photos';
    // Placeholder value stored in AppState.data / localStorage
    // so the rest of the app knows "there IS a photo here, look it up in IDB".
    const IDB_REF = '__IDB_PHOTO__';

    let _db = null;

    // ── Open / create the database ──
    function openDB() {
        return new Promise(function (resolve, reject) {
            if (_db) { resolve(_db); return; }
            const req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = function (e) {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    // key = "projectName::rowIdx::colName"
                    db.createObjectStore(STORE_NAME);
                }
            };
            req.onsuccess = function (e) {
                _db = e.target.result;
                resolve(_db);
            };
            req.onerror = function (e) {
                console.error('PhotoStore: IndexedDB open failed', e);
                reject(e);
            };
        });
    }

    // ── Build a deterministic key for a photo ──
    function makeKey(projectName, rowIdx, colName) {
        return projectName + '::' + rowIdx + '::' + colName;
    }

    // ── Save a single photo (data-URL string) ──
    function savePhoto(projectName, rowIdx, colName, dataUrl) {
        return openDB().then(function (db) {
            return new Promise(function (resolve, reject) {
                var tx = db.transaction(STORE_NAME, 'readwrite');
                var store = tx.objectStore(STORE_NAME);
                store.put(dataUrl, makeKey(projectName, rowIdx, colName));
                tx.oncomplete = function () { resolve(); };
                tx.onerror = function (e) { reject(e); };
            });
        });
    }

    // ── Load a single photo → returns the data-URL or null ──
    function loadPhoto(projectName, rowIdx, colName) {
        return openDB().then(function (db) {
            return new Promise(function (resolve, reject) {
                var tx = db.transaction(STORE_NAME, 'readonly');
                var store = tx.objectStore(STORE_NAME);
                var req = store.get(makeKey(projectName, rowIdx, colName));
                req.onsuccess = function () { resolve(req.result || null); };
                req.onerror = function (e) { reject(e); };
            });
        });
    }

    // ── Delete a single photo ──
    function deletePhoto(projectName, rowIdx, colName) {
        return openDB().then(function (db) {
            return new Promise(function (resolve, reject) {
                var tx = db.transaction(STORE_NAME, 'readwrite');
                var store = tx.objectStore(STORE_NAME);
                store.delete(makeKey(projectName, rowIdx, colName));
                tx.oncomplete = function () { resolve(); };
                tx.onerror = function (e) { reject(e); };
            });
        });
    }

    // ── Delete ALL photos for a project ──
    function deleteProjectPhotos(projectName) {
        var prefix = projectName + '::';
        return openDB().then(function (db) {
            return new Promise(function (resolve, reject) {
                var tx = db.transaction(STORE_NAME, 'readwrite');
                var store = tx.objectStore(STORE_NAME);
                var cursorReq = store.openCursor();
                cursorReq.onsuccess = function (e) {
                    var cursor = e.target.result;
                    if (cursor) {
                        if (typeof cursor.key === 'string' && cursor.key.startsWith(prefix)) {
                            cursor.delete();
                        }
                        cursor.continue();
                    }
                };
                tx.oncomplete = function () { resolve(); };
                tx.onerror = function (e) { reject(e); };
            });
        });
    }

    // ── Save ALL photos from a project's data array into IDB ──
    // Returns a deep-copy of `dataRows` with data:image values replaced by IDB_REF.
    function extractAndSavePhotos(projectName, dataRows) {
        var strippedData = JSON.parse(JSON.stringify(dataRows));
        var promises = [];
        strippedData.forEach(function (row, rowIdx) {
            Object.keys(row).forEach(function (k) {
                if (typeof row[k] === 'string' && row[k].startsWith('data:image')) {
                    // Save the actual photo to IDB
                    promises.push(savePhoto(projectName, rowIdx, k, row[k]));
                    // Replace with lightweight placeholder
                    row[k] = IDB_REF;
                }
            });
        });
        return Promise.all(promises).then(function () {
            return strippedData;
        });
    }

    // ── Restore IDB_REF placeholders in a data array back to real data-URLs ──
    function restorePhotos(projectName, dataRows) {
        var promises = [];
        dataRows.forEach(function (row, rowIdx) {
            Object.keys(row).forEach(function (k) {
                if (row[k] === IDB_REF) {
                    promises.push(
                        loadPhoto(projectName, rowIdx, k).then(function (dataUrl) {
                            if (dataUrl) {
                                row[k] = dataUrl;
                            } else {
                                // Photo was in IDB placeholder but actual data missing
                                row[k] = '';
                            }
                        })
                    );
                }
            });
        });
        return Promise.all(promises).then(function () {
            return dataRows;
        });
    }

    // ── Check if IndexedDB is available ──
    function isAvailable() {
        return typeof indexedDB !== 'undefined';
    }

    // ── Expose ──
    return {
        IDB_REF: IDB_REF,
        isAvailable: isAvailable,
        openDB: openDB,
        savePhoto: savePhoto,
        loadPhoto: loadPhoto,
        deletePhoto: deletePhoto,
        deleteProjectPhotos: deleteProjectPhotos,
        extractAndSavePhotos: extractAndSavePhotos,
        restorePhotos: restorePhotos
    };
})();

// Make globally accessible
window.PhotoStore = PhotoStore;
