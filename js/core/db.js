/**
 * MAIO Database Engine
 * Quản lý kho dữ liệu Local-first (IndexedDB) và liên kết đồng bộ Google Drive.
 */

import { gdrive } from './gdrive.js?v=3.7';

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
    const safe = email.toLowerCase().trim().replace(/[^a-z0-9]/g, '_');
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
    return snapshot;
  }

  /**
   * Nạp dữ liệu snapshot từ Google Drive hoặc file JSON vào IndexedDB
   */
  async importSnapshot(snapshot) {
    if (!snapshot || !snapshot.data) throw new Error('Dữ liệu không đúng định dạng snapshot MAIO.');
    if (!this.db) await this.init();

    for (const [storeName, records] of Object.entries(snapshot.data)) {
      if (Array.isArray(records)) {
        const tx = this.db.transaction([storeName], 'readwrite');
        const store = tx.objectStore(storeName);
        store.clear();
        for (const item of records) {
          store.put(item);
        }
      }
    }
    return true;
  }

  /**
   * Đồng bộ lên Google Drive theo cấu trúc thư mục chuẩn
   */
  async syncToGoogleDrive() {
    if (!gdrive.isSignedIn()) {
      throw new Error('Vui lòng kết nối Google Drive trước khi đồng bộ.');
    }
    const snapshot = await this.exportSnapshot();
    const curPrjId = this.activeProjectId || 'CHC1';

    let cpmsFolderId = null;
    try {
      // 1. Đảm bảo cấu trúc thư mục dự án và data CPMS
      cpmsFolderId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`, 'CPMS_Data']);

      // 2. Tạo sẵn 3 thư mục phân loại hồ sơ scan theo chuẩn
      await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`, 'HopDong_PhuLuc']);
      await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`, 'Doc']);
      await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`, 'ThanhToan']);
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
}

if (!window.__MAIO_DB_ENGINE__) {
  window.__MAIO_DB_ENGINE__ = new DatabaseEngine();
}
export const db = window.__MAIO_DB_ENGINE__;
