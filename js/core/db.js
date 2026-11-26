/**
 * MAIO Database Engine
 * Quản lý kho dữ liệu Local-first (IndexedDB) và liên kết đồng bộ Google Drive.
 */

import { gdrive } from './gdrive.js?v=5.3';

const DB_VERSION = 1;

class DatabaseEngine {
  constructor() {
    this.db = null;
    this.activeProjectId = 'CHC1';
    this.currentUserEmail = null;
    this.currentDbName = 'MAIO_GuestDB';
  }

  getDbNameForUser(email) {
    if (!email) return 'MAIO_GuestDB';
    const clean = email.toLowerCase().trim();
    if (clean === 'local_user@maio.local' || clean === 'local' || clean.startsWith('demo_user')) {
      return 'MAIO_LocalDB';
    }
    const safe = clean.replace(/[^a-z0-9]/g, '_');
    return `MAIO_UserDB_${safe}`;
  }

  async switchUser(userEmail) {
    if (this.db) {
      try {
        this.db.close();
      } catch (e) {}
      this.db = null;
    }
    this.currentUserEmail = userEmail ? userEmail.trim().toLowerCase() : null;
    this.currentDbName = this.getDbNameForUser(this.currentUserEmail);
    return await this.init();
  }

  close() {
    if (this.db) {
      try {
        this.db.close();
      } catch (e) {}
      this.db = null;
    }
    this.currentUserEmail = null;
    this.currentDbName = 'MAIO_GuestDB';
  }

  async init() {
    const targetDbName = this.currentDbName || 'MAIO_GuestDB';

    return new Promise((resolve, reject) => {
      const request = indexedDB.open(targetDbName, DB_VERSION);

      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        
        // Bảng dự án
        if (!db.objectStoreNames.contains('projects')) {
          db.createObjectStore('projects', { keyPath: 'id' });
        }
        // Bảng hợp đồng CPMS
        if (!db.objectStoreNames.contains('contracts')) {
          const store = db.createObjectStore('contracts', { keyPath: 'id' });
          store.createIndex('project_id', 'project_id', { unique: false });
          store.createIndex('group', 'group', { unique: false });
          store.createIndex('parent_id', 'parent_id', { unique: false });
        }
        // Bảng thanh toán CPMS
        if (!db.objectStoreNames.contains('payments')) {
          const store = db.createObjectStore('payments', { keyPath: 'id' });
          store.createIndex('project_id', 'project_id', { unique: false });
          store.createIndex('contract_id', 'contract_id', { unique: false });
        }
        // Bảng hồ sơ công văn CPMS
        if (!db.objectStoreNames.contains('documents')) {
          const store = db.createObjectStore('documents', { keyPath: 'id' });
          store.createIndex('project_id', 'project_id', { unique: false });
          store.createIndex('contract_id', 'contract_id', { unique: false });
        }
        // Bảng tài sản văn phòng
        if (!db.objectStoreNames.contains('office_assets')) {
          db.createObjectStore('office_assets', { keyPath: 'id' });
        }
        // Bảng thu chi sổ quỹ
        if (!db.objectStoreNames.contains('cash_transactions')) {
          db.createObjectStore('cash_transactions', { keyPath: 'id' });
        }
      };

      request.onsuccess = (event) => {
        this.db = event.target.result;
        resolve(this.db);
      };

      request.onerror = (event) => {
        console.error('IndexedDB Error:', event.target.error);
        reject(event.target.error);
      };
    });
  }

  /**
   * Thêm hoặc cập nhật bản ghi
   */
  async put(storeName, data) {
    if (!this.db) await this.init();
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction([storeName], 'readwrite');
      const store = tx.objectStore(storeName);
      const req = store.put(data);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  /**
   * Lấy toàn bộ bản ghi theo storeName
   */
  async getAll(storeName) {
    if (!this.db) await this.init();
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction([storeName], 'readonly');
      const store = tx.objectStore(storeName);
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  /**
   * Lấy một bản ghi theo khóa chính
   */
  async get(storeName, key) {
    if (!this.db) await this.init();
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction([storeName], 'readonly');
      const store = tx.objectStore(storeName);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  /**
   * Xóa một bản ghi
   */
  async delete(storeName, key) {
    if (!this.db) await this.init();
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction([storeName], 'readwrite');
      const store = tx.objectStore(storeName);
      const req = store.delete(key);
      req.onsuccess = () => resolve(true);
      req.onerror = () => reject(req.error);
    });
  }

  /**
   * Xóa toàn bộ dữ liệu của một dự án (dự án, hợp đồng, thanh toán, hồ sơ)
   */
  async deleteProjectCascade(projectId) {
    if (!this.db) await this.init();

    // 1. Xóa trong projects
    await this.delete('projects', projectId);

    // 2. Xóa trong contracts
    const allContracts = await this.getAll('contracts');
    for (const c of allContracts) {
      if (c.project_id === projectId) {
        await this.delete('contracts', c.id);
      }
    }

    // 3. Xóa trong payments
    const allPayments = await this.getAll('payments');
    for (const p of allPayments) {
      if (p.project_id === projectId) {
        await this.delete('payments', p.id);
      }
    }

    // 4. Xóa trong documents
    const allDocs = await this.getAll('documents');
    for (const d of allDocs) {
      if (d.project_id === projectId) {
        await this.delete('documents', d.id);
      }
    }

    return true;
  }

  /**
   * Xuất toàn bộ dữ liệu ra đối tượng JSON để lưu lên Google Drive
   */
  async exportSnapshot() {
    const stores = ['projects', 'contracts', 'payments', 'documents', 'office_assets', 'cash_transactions'];
    const snapshot = {
      version: '1.0.0',
      exported_at: new Date().toISOString(),
      active_project: this.activeProjectId,
      data: {},
    };

    for (const storeName of stores) {
      snapshot.data[storeName] = await this.getAll(storeName);
    }

    // Chuẩn hóa relative_path cho documents khi xuất snapshot lên Drive để đồng bộ đa máy
    if (Array.isArray(snapshot.data.documents)) {
      const rootCache = new Map();
      snapshot.data.documents = snapshot.data.documents.map((doc) => {
        if (!doc.local_link || String(doc.local_link).startsWith('http')) return doc;
        const pid = doc.project_id || this.activeProjectId || 'CHC1';
        if (!rootCache.has(pid)) {
          rootCache.set(pid, this.getProjectLocalRootPath(pid));
        }
        const rootPath = rootCache.get(pid);
        const rel = doc.relative_path || this.computeRelativePath(doc.local_link, rootPath);
        return rel ? { ...doc, relative_path: rel } : doc;
      });
    }

    return snapshot;
  }

  /**
   * Nạp dữ liệu snapshot từ Google Drive hoặc file JSON vào IndexedDB
   * Bảo toàn local_link theo thư mục gốc riêng của máy tính hiện tại
   */
  async importSnapshot(snapshot) {
    if (!snapshot || !snapshot.data) throw new Error('Dữ liệu không đúng định dạng snapshot MAIO.');
    if (!this.db) await this.init();

    // Đọc trước local_link hiện có trên máy này để không bị ghi đè đường dẫn của máy khác
    const existingDocs = await this.getAll('documents');
    const existingLocalByDocId = new Map();
    (existingDocs || []).forEach((d) => {
      if (d && d.id && d.local_link && !String(d.local_link).startsWith('http')) {
        existingLocalByDocId.set(d.id, d.local_link);
      }
    });

    for (const [storeName, records] of Object.entries(snapshot.data)) {
      if (Array.isArray(records)) {
        const tx = this.db.transaction([storeName], 'readwrite');
        const store = tx.objectStore(storeName);
        store.clear();
        for (const item of records) {
          if (storeName === 'documents' && item) {
            const pid = item.project_id || snapshot.active_project || this.activeProjectId;
            const localRoot = this.getProjectLocalRootPath(pid);
            const resolved =
              localRoot && item.relative_path
                ? this.resolveLocalLink(item.relative_path, localRoot)
                : existingLocalByDocId.get(item.id) || item.local_link || '';
            store.put({
              ...item,
              local_link: resolved,
            });
          } else {
            store.put(item);
          }
        }
      }
    }
    return true;
  }

  /**
   * Đồng bộ lên Google Drive theo cấu trúc thư mục chuẩn:
   * MAIO_Data / 01_QuanLy_DuAn / DuAn_<id> / { CPMS_Data, HoSo, HopDong_PL, ThanhToan }
   */
  async syncToGoogleDrive(options = {}) {
    if (!gdrive.isSignedIn()) {
      throw new Error('Vui lòng kết nối Google Drive trước khi đồng bộ.');
    }
    const curPrjId = options.projectId || this.activeProjectId || 'CHC1';
    this.activeProjectId = curPrjId;

    const snapshot = await this.exportSnapshot();
    snapshot.active_project = curPrjId;

    let cpmsFolderId = options.cpmsFolderId || null;
    let projectFolderId = options.targetProjectFolderId || null;

    try {
      if (!projectFolderId) {
        projectFolderId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`]);
      }

      if (!cpmsFolderId && projectFolderId) {
        cpmsFolderId = await gdrive.ensureFolderPath(['CPMS_Data'], projectFolderId);
      }
    } catch (e) {
      console.warn('Lỗi tạo cấu trúc thư mục phân cấp, lưu tại thư mục gốc MAIO_Data:', e);
    }

    const fileName = `MAIO_CPMS_Snapshot_${curPrjId}.json`;
    return await gdrive.saveJsonFile(fileName, snapshot, cpmsFolderId);
  }

  /**
   * Đồng bộ từ Google Drive về máy
   */
  async syncFromGoogleDrive(fileId) {
    if (!gdrive.isSignedIn()) {
      throw new Error('Vui lòng kết nối Google Drive trước.');
    }
    const snapshot = await gdrive.readJsonFile(fileId);
    await this.importSnapshot(snapshot);
    return snapshot;
  }

  /**
   * Hợp nhất dữ liệu từ Cục bộ (MAIO_LocalDB) sang tài khoản Google khi đăng nhập
   * Đảm bảo dự án đang mở ở Local không bao giờ bị biến mất khi chuyển sang Google Drive
   */
  async copyLocalDataIfEmpty(newUserEmail, preferredProjectId = null) {
    if (!newUserEmail) return;
    const clean = newUserEmail.toLowerCase().trim();
    if (clean === 'local_user@maio.local' || clean.startsWith('demo_user')) return;

    try {
      const localProjects = await this.readAllFromNamedDb('MAIO_LocalDB', 'projects');
      if (!localProjects || localProjects.length === 0) return;

      const curProjects = await this.getAll('projects');
      const existingPrjIds = new Set((curProjects || []).map((p) => p.id));

      // Xác định các dự án cần chuyển sang: toàn bộ nếu DB user đang rỗng, hoặc dự án đang mở (preferredProjectId) nếu chưa có trong DB user
      const projectsToCopy =
        curProjects.length === 0
          ? localProjects
          : localProjects.filter((p) => !existingPrjIds.has(p.id) || (preferredProjectId && p.id === preferredProjectId));

      if (projectsToCopy.length === 0) return;

      const copyPrjIds = new Set(projectsToCopy.map((p) => p.id));
      const localContracts = await this.readAllFromNamedDb('MAIO_LocalDB', 'contracts');
      const localPayments = await this.readAllFromNamedDb('MAIO_LocalDB', 'payments');
      const localDocs = await this.readAllFromNamedDb('MAIO_LocalDB', 'documents');

      for (const p of projectsToCopy) await this.save('projects', p);
      for (const c of localContracts) {
        if (!c.project_id || copyPrjIds.has(c.project_id)) await this.save('contracts', c);
      }
      for (const pay of localPayments) {
        if (!pay.project_id || copyPrjIds.has(pay.project_id)) await this.save('payments', pay);
      }
      for (const d of localDocs) {
        if (!d.project_id || copyPrjIds.has(d.project_id)) await this.save('documents', d);
      }

      console.log(`Đã đồng bộ ${projectsToCopy.length} dự án từ Cục bộ sang tài khoản Google: ${newUserEmail}`);
    } catch (e) {
      console.warn('Lỗi copy dữ liệu local sang user DB:', e);
    }
  }

  async readAllFromNamedDb(dbName, storeName) {
    return new Promise((resolve) => {
      const req = indexedDB.open(dbName);
      req.onerror = () => resolve([]);
      req.onsuccess = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(storeName)) {
          db.close();
          return resolve([]);
        }
        try {
          const tx = db.transaction(storeName, 'readonly');
          const store = tx.objectStore(storeName);
          const getAllReq = store.getAll();
          getAllReq.onsuccess = () => {
            db.close();
            resolve(getAllReq.result || []);
          };
          getAllReq.onerror = () => {
            db.close();
            resolve([]);
          };
        } catch (err) {
          db.close();
          resolve([]);
        }
      };
    });
  }

  /**
   * Lưu FileSystemDirectoryHandle theo dự án vào IndexedDB riêng (MAIO_HandlesDB)
   * để các lần đồng bộ sau không phải chọn lại thư mục
   */
  async saveProjectDirHandle(projectId, dirHandle) {
    if (!projectId || !dirHandle) return false;
    return new Promise((resolve) => {
      const req = indexedDB.open('MAIO_HandlesDB', 1);
      req.onupgradeneeded = (e) => {
        const hdb = e.target.result;
        if (!hdb.objectStoreNames.contains('dir_handles')) {
          hdb.createObjectStore('dir_handles', { keyPath: 'key' });
        }
      };
      req.onsuccess = (e) => {
        const hdb = e.target.result;
        try {
          const tx = hdb.transaction(['dir_handles'], 'readwrite');
          tx.objectStore('dir_handles').put({ key: `prj_${projectId}`, handle: dirHandle, updated_at: Date.now() });
          tx.oncomplete = () => {
            hdb.close();
            resolve(true);
          };
          tx.onerror = () => {
            hdb.close();
            resolve(false);
          };
        } catch (_) {
          hdb.close();
          resolve(false);
        }
      };
      req.onerror = () => resolve(false);
    });
  }

  /**
   * Lấy FileSystemDirectoryHandle đã lưu cho dự án
   */
  async getProjectDirHandle(projectId) {
    if (!projectId) return null;
    return new Promise((resolve) => {
      const req = indexedDB.open('MAIO_HandlesDB', 1);
      req.onupgradeneeded = (e) => {
        const hdb = e.target.result;
        if (!hdb.objectStoreNames.contains('dir_handles')) {
          hdb.createObjectStore('dir_handles', { keyPath: 'key' });
        }
      };
      req.onsuccess = (e) => {
        const hdb = e.target.result;
        try {
          const tx = hdb.transaction(['dir_handles'], 'readonly');
          const getReq = tx.objectStore('dir_handles').get(`prj_${projectId}`);
          getReq.onsuccess = () => {
            hdb.close();
            resolve(getReq.result?.handle || null);
          };
          getReq.onerror = () => {
            hdb.close();
            resolve(null);
          };
        } catch (_) {
          hdb.close();
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    });
  }

  /**
   * Lưu FileSystemFileHandle (file .cpms local) theo dự án vào MAIO_HandlesDB
   */
  async saveProjectFileHandle(projectId, fileHandle) {
    if (!projectId || !fileHandle) return false;
    return new Promise((resolve) => {
      const req = indexedDB.open('MAIO_HandlesDB', 1);
      req.onupgradeneeded = (e) => {
        const hdb = e.target.result;
        if (!hdb.objectStoreNames.contains('dir_handles')) {
          hdb.createObjectStore('dir_handles', { keyPath: 'key' });
        }
      };
      req.onsuccess = (e) => {
        const hdb = e.target.result;
        try {
          const tx = hdb.transaction(['dir_handles'], 'readwrite');
          tx.objectStore('dir_handles').put({ key: `cpms_file_${projectId}`, handle: fileHandle, updated_at: Date.now() });
          tx.oncomplete = () => {
            hdb.close();
            resolve(true);
          };
          tx.onerror = () => {
            hdb.close();
            resolve(false);
          };
        } catch (_) {
          hdb.close();
          resolve(false);
        }
      };
      req.onerror = () => resolve(false);
    });
  }

  /**
   * Lấy FileSystemFileHandle (file .cpms local) đã lưu cho dự án
   */
  async getProjectFileHandle(projectId) {
    if (!projectId) return null;
    return new Promise((resolve) => {
      const req = indexedDB.open('MAIO_HandlesDB', 1);
      req.onupgradeneeded = (e) => {
        const hdb = e.target.result;
        if (!hdb.objectStoreNames.contains('dir_handles')) {
          hdb.createObjectStore('dir_handles', { keyPath: 'key' });
        }
      };
      req.onsuccess = (e) => {
        const hdb = e.target.result;
        try {
          const tx = hdb.transaction(['dir_handles'], 'readonly');
          const getReq = tx.objectStore('dir_handles').get(`cpms_file_${projectId}`);
          getReq.onsuccess = () => {
            hdb.close();
            resolve(getReq.result?.handle || null);
          };
          getReq.onerror = () => {
            hdb.close();
            resolve(null);
          };
        } catch (_) {
          hdb.close();
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    });
  }

  /**
   * Xóa FileSystemFileHandle đã lưu khi xóa dự án
   */
  async deleteProjectFileHandle(projectId) {
    if (!projectId) return false;
    return new Promise((resolve) => {
      const req = indexedDB.open('MAIO_HandlesDB', 1);
      req.onupgradeneeded = (e) => {
        const hdb = e.target.result;
        if (!hdb.objectStoreNames.contains('dir_handles')) {
          hdb.createObjectStore('dir_handles', { keyPath: 'key' });
        }
      };
      req.onsuccess = (e) => {
        const hdb = e.target.result;
        try {
          const tx = hdb.transaction(['dir_handles'], 'readwrite');
          tx.objectStore('dir_handles').delete(`cpms_file_${projectId}`);
          tx.oncomplete = () => {
            hdb.close();
            resolve(true);
          };
          tx.onerror = () => {
            hdb.close();
            resolve(false);
          };
        } catch (_) {
          hdb.close();
          resolve(false);
        }
      };
      req.onerror = () => resolve(false);
    });
  }

  /**
   * Lưu thư mục gốc (local_root_path) riêng của máy tính hiện tại cho từng dự án
   */
  saveProjectLocalRootPath(projectId, rootPath) {
    if (!projectId || !rootPath) return;
    try {
      localStorage.setItem(`maio_local_root_${projectId}`, String(rootPath).trim());
    } catch (_) {}
  }

  getProjectLocalRootPath(projectId) {
    if (!projectId) return '';
    try {
      return localStorage.getItem(`maio_local_root_${projectId}`) || '';
    } catch (_) {
      return '';
    }
  }

  /**
   * Tách đường dẫn tương đối từ đường dẫn đầy đủ trên máy tính và thư mục gốc
   */
  computeRelativePath(fullPath, rootPath = '') {
    if (!fullPath) return '';
    const normFull = String(fullPath).replace(/\\/g, '/').trim();
    if (!normFull || normFull.startsWith('http')) return '';
    if (rootPath) {
      const normRoot = String(rootPath).replace(/\\/g, '/').replace(/\/+$/, '').trim();
      if (normRoot && normFull.toLowerCase().startsWith(normRoot.toLowerCase() + '/')) {
        return normFull.slice(normRoot.length + 1);
      }
      // Nếu rootPath chỉ là tên thư mục gốc (ví dụ "_Westlake Residential")
      const rootFolderName = normRoot.split('/').pop();
      if (rootFolderName) {
        const marker = '/' + rootFolderName.toLowerCase() + '/';
        const idx = normFull.toLowerCase().indexOf(marker);
        if (idx !== -1) {
          return normFull.slice(idx + marker.length);
        }
      }
    }
    return '';
  }

  /**
   * Ghép lại đường dẫn local đầy đủ trên máy hiện tại từ relative_path + local_root_path
   */
  resolveLocalLink(relativePath, localRootPath) {
    if (!relativePath) return '';
    const cleanRel = String(relativePath).replace(/^[/\\]+/, '').replace(/\//g, '\\');
    if (!localRootPath) return cleanRel;
    const cleanRoot = String(localRootPath).replace(/[/\\]+$/, '');
    return `${cleanRoot}\\${cleanRel}`;
  }
}

if (!window.__MAIO_DB_ENGINE__) {
  window.__MAIO_DB_ENGINE__ = new DatabaseEngine();
}
export const db = window.__MAIO_DB_ENGINE__;
