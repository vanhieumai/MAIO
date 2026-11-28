/**
 * MAIO CPMS Contracts & Payments Module
 * Hỗ trợ phân cấp chuẩn xác:
 * - Hợp đồng gốc (Root, Level 0)
 *   - Phụ lục cấp 1 (Level 1, vd: APP-05, APP-06...)
 *     - Gói con cấp 2 (Level 2, vd: PKG-02, PKG-03 thuộc APP-06)
 *
 * QUY TẮC TÍNH TOÁN NGHIÊM NGẶT (CHỐNG DOUBLE COUNTING):
 * 1. Giá trị ký kết: Chỉ cộng dồn các nhánh trực tiếp (Direct Children), không cộng lặp các gói con cấp 2 đã nằm trong phụ lục cha.
 * 2. Giải ngân & Thanh toán: CHỈ cộng các dòng giao dịch thực tế (is_header !== true), tuyệt đối không cộng dòng Header tổng hợp.
 */

import { db } from '../../core/db.js?v=5.3';
import { maioOpen } from '../../core/maio-open.js?v=1.1';
import { CPMSImporterExporter } from '../../core/cpms-importer.js?v=5.7';

export class ContractsManager {
  constructor() {
    this.currentProjectId = 'CHC1';
    this.projects = [];
    this.currentGroup = 'A';
    this.searchQuery = '';
    this.categoryFilter = 'ALL';
    this.contracts = [];
    this.payments = [];
    this.documents = [];
    this.selectedContract = null;
    this.expandedContracts = new Set();
    this.activeFileHandle = null;
    this._toastTimer = null;
    this.disburseGroup = 'A';
    this.selectedDisburseContractId = null;
    this.disburseContractSearchTerm = '';
    this.disburseEventsBound = false;
    this.reminderShowCompleted = false;
    this.reminderDocType = 'ALL';
    this.reminderCategory = 'ALL';
    this.reminderSearchTerm = '';
    this.selectedReminderDocId = null;
    this.reminderEventsBound = false;
    this.overviewCompanyQuery = '';
    this.overviewSelectedCompanies = new Set();
    this.overviewSelectedContractIds = new Set();
    this.overviewCompanyEventsBound = false;
    this._driveAutoSaveTimer = null;
    this._driveFolderCache = new Map();
    this._lastKnownFileModified = 0;
    this._isSyncingLocalFile = false;
    this._isCheckingFileChange = false;
    this._localFileWatcherStarted = false;
    window.maioConfirm = (opts) => this.confirmAsync(opts);
    window.maioAlert = (opts) => this.alertAsync(opts);
    window.alert = (msg) => window.maioAlert(msg);
  }

  showLocalSaveStatus(message, isError = false) {
    let toast = document.getElementById('maio-local-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'maio-local-toast';
      toast.style.cssText = `
        position: fixed;
        bottom: 24px;
        right: 24px;
        z-index: 999999;
        padding: 10px 18px;
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 600;
        box-shadow: 0 10px 25px -5px rgba(0,0,0,0.2);
        display: flex;
        align-items: center;
        gap: 8px;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        opacity: 0;
        transform: translateY(10px);
      `;
      document.body.appendChild(toast);
    }
    toast.style.background = isError ? '#fef2f2' : '#f0fdf4';
    toast.style.color = isError ? '#dc2626' : '#15803d';
    toast.style.border = `1px solid ${isError ? '#fecaca' : '#bbf7d0'}`;
    toast.innerHTML = `<i class="fa-solid ${isError ? 'fa-triangle-exclamation' : 'fa-circle-check'}"></i> ${message}`;
    toast.style.opacity = '1';
    toast.style.transform = 'translateY(0)';

    clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
    }, 3500);
  }

  /**
   * Tạo binary buffer file .cpms/.xlsx từ dữ liệu dự án hiện tại trong IndexedDB
   */
  async buildProjectCPMSBuffer(projectId = this.currentProjectId) {
    return await CPMSImporterExporter.buildCPMSBinaryBuffer(projectId);
  }

  /**
   * Lên lịch tự động lưu file .cpms và snapshot JSON lên Google Drive mỗi khi có thay đổi
   */
  scheduleAutoSaveToDrive(immediate = false) {
    const gdriveSvc = window.__MAIO_GDRIVE_SERVICE__;
    if (!gdriveSvc || !gdriveSvc.isSignedIn()) return;

    clearTimeout(this._driveAutoSaveTimer);
    this._driveAutoSaveTimer = setTimeout(() => {
      this.saveProjectCPMSOnline().catch((err) => {
        console.warn('Lỗi tự động lưu file CPMS lên Google Drive:', err);
      });
    }, immediate ? 50 : 800);
  }

  /**
   * Tự động lưu file .cpms nhị phân và snapshot JSON vào CPMS_Data trên Google Drive
   */
  async saveProjectCPMSOnline() {
    const gdriveSvc = window.__MAIO_GDRIVE_SERVICE__;
    if (!gdriveSvc || !gdriveSvc.isSignedIn() || !this.currentProjectId) return false;

    const curPrjId = this.currentProjectId;
    const curPrj = (this.projects || []).find((p) => p.id === curPrjId);
    const safePrjId = String(curPrjId).replace(/[\/\\?%*:|"<>]/g, '_');
    const defaultFolderName = `DuAn_${safePrjId}`;

    let cached = this._driveFolderCache.get(curPrjId);
    if (!cached) {
      let projectFolderId = curPrj?.drive_folder_id || '';
      if (!projectFolderId) {
        const quanLyDuAnId = await gdriveSvc.ensureFolderPath(['01_QuanLy_DuAn']);
        const existingFolders = await gdriveSvc.findFoldersInParent(quanLyDuAnId);
        const prjNameLower = String(curPrj?.name || '').trim().toLowerCase();
        const nameTokens = prjNameLower
          .replace(/dự án|nhà ở|chung cư|cao tầng|khu đô thị/gi, '')
          .trim()
          .split(/\s+/)
          .filter((w) => w.length >= 3);

        let matched = existingFolders.find((f) => {
          const fname = String(f.name || '').toLowerCase();
          if (fname === defaultFolderName.toLowerCase() || fname.includes(safePrjId.toLowerCase())) return true;
          if (prjNameLower && fname.includes(prjNameLower)) return true;
          if (nameTokens.length >= 2 && nameTokens.every((t) => fname.includes(t))) return true;
          return false;
        });
        if (!matched && existingFolders.length === 1 && prjNameLower.includes('westlake') && existingFolders[0].name === 'DuAn_CHC1') {
          matched = existingFolders[0];
        }

        projectFolderId = matched
          ? matched.id
          : await gdriveSvc.createFolder(defaultFolderName, quanLyDuAnId);
      }
      const cpmsFolderId = await gdriveSvc.ensureFolderPath(['CPMS_Data'], projectFolderId);

      cached = { projectFolderId, cpmsFolderId };
      this._driveFolderCache.set(curPrjId, cached);
    }

    await db.syncToGoogleDrive({
      projectId: curPrjId,
      targetProjectFolderId: cached.projectFolderId,
      cpmsFolderId: cached.cpmsFolderId,
    });

    const cpmsBuffer = await this.buildProjectCPMSBuffer(curPrjId);
    const cpmsFileName = this.activeFileHandle?.name || curPrj?.source_filename || `MAIO_Data_${safePrjId}.cpms`;
    await gdriveSvc.uploadBinaryFile(cpmsFileName, cpmsBuffer, 'application/octet-stream', cached.cpmsFolderId);

    if (curPrj) {
      curPrj.drive_synced = true;
      curPrj.drive_folder_id = cached.projectFolderId;
      await db.put('projects', curPrj);
    }

    if (typeof window.__MAIO_APP__?.refreshOnlineDriveProjects === 'function') {
      await window.__MAIO_APP__.refreshOnlineDriveProjects();
      this.updateProjectSelectorUI();
    }

    this.showLocalSaveStatus(`Đã tự động lưu thay đổi lên Google Drive (${cpmsFileName})!`);
    return true;
  }

  /**
   * Tự động lưu mọi thay đổi trực tiếp vào file máy tính đã liên kết (và Google Drive nếu đang Online)
   */
  async syncToLocalFile(isManual = false) {
    const gdriveSvc = window.__MAIO_GDRIVE_SERVICE__;
    if (gdriveSvc && !gdriveSvc.isSignedIn() && gdriveSvc.currentUser && !gdriveSvc.currentUser.isOffline && typeof gdriveSvc.trySilentRefreshToken === 'function') {
      await gdriveSvc.trySilentRefreshToken();
    }
    const isDriveOnline = Boolean(gdriveSvc && gdriveSvc.isSignedIn());

    if (isDriveOnline) {
      this.scheduleAutoSaveToDrive(isManual);
    }

    if (!this.activeFileHandle && this.currentProjectId) {
      const savedHandle = await db.getProjectFileHandle(this.currentProjectId);
      if (savedHandle) {
        this.activeFileHandle = savedHandle;
      }
    }

    if (!this.activeFileHandle) {
      if (isManual && !isDriveOnline) {
        alert('Chưa có liên kết tệp trực tiếp trên máy tính. Hãy bấm "Nhập dữ liệu CPMS (Dẫn file Local)" để chọn tệp trên máy tính trước.');
      }
      return isDriveOnline;
    }

    try {
      if (typeof this.activeFileHandle.queryPermission === 'function') {
        let perm = await this.activeFileHandle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
          perm = await this.activeFileHandle.requestPermission({ mode: 'readwrite' });
        }
        if (perm !== 'granted') {
          this.showLocalSaveStatus(`Chưa được cấp quyền ghi vào file "${this.activeFileHandle.name}".`, true);
          return false;
        }
      }

      this._isSyncingLocalFile = true;
      const buffer = await CPMSImporterExporter.buildCPMSBinaryBuffer(this.currentProjectId);
      const writable = await this.activeFileHandle.createWritable();
      await writable.write(buffer);
      await writable.close();

      try {
        const updatedFile = await this.activeFileHandle.getFile();
        this._lastKnownFileModified = updatedFile.lastModified || Date.now();
      } catch (_) {
        this._lastKnownFileModified = Date.now();
      }

      const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
      if (curPrj) {
        curPrj.source_file_updated = new Date().toISOString();
        await db.put('projects', curPrj);
      }

      this.showLocalSaveStatus(`Đã tự động lưu thay đổi vào file "${this.activeFileHandle.name}"!`);
      return true;
    } catch (err) {
      console.warn('Lỗi lưu file local:', err);
      this.showLocalSaveStatus(`Không thể ghi file "${this.activeFileHandle.name}": ${err.message}`, true);
      return false;
    } finally {
      this._isSyncingLocalFile = false;
    }
  }

  /**
   * Theo dõi sự thay đổi của file .cpms trên máy tính (khi sửa từ trình duyệt khác hoặc Excel/CPMS Plus)
   */
  startLocalFileWatcher() {
    if (this._localFileWatcherStarted) return;
    this._localFileWatcherStarted = true;

    window.addEventListener('focus', () => {
      this.checkExternalLocalFileUpdate();
    });
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') {
        this.checkExternalLocalFileUpdate();
      }
    });
    setInterval(() => {
      this.checkExternalLocalFileUpdate();
    }, 4000);
  }

  async checkExternalLocalFileUpdate() {
    if (!this.activeFileHandle || this._isSyncingLocalFile || this._isCheckingFileChange) return;
    this._isCheckingFileChange = true;
    try {
      if (typeof this.activeFileHandle.queryPermission === 'function') {
        const perm = await this.activeFileHandle.queryPermission({ mode: 'read' });
        if (perm !== 'granted') return;
      }
      const file = await this.activeFileHandle.getFile();
      if (!this._lastKnownFileModified) {
        this._lastKnownFileModified = file.lastModified || 0;
        return;
      }
      if (file.lastModified && file.lastModified > this._lastKnownFileModified + 500) {
        this._lastKnownFileModified = file.lastModified;
        const result = await CPMSImporterExporter.importCPMSFile(file);
        this.currentProjectId = result.projectData.id;
        await this.loadData();
        this.render();
        if (typeof window.__MAIO_APP__?.updateDashboardStats === 'function') {
          await window.__MAIO_APP__.updateDashboardStats();
        }
        this.showLocalSaveStatus(`Phát hiện "${file.name}" vừa thay đổi từ bên ngoài — đã tự động cập nhật dữ liệu mới!`);
      }
    } catch (err) {
      console.warn('Lỗi kiểm tra thay đổi file local:', err);
    } finally {
      this._isCheckingFileChange = false;
    }
  }

  /**
   * Mở trực tiếp File Picker của trình duyệt để liên kết file máy tính
   */
  async pickAndLoadLocalCPMSFile() {
    if ('showOpenFilePicker' in window) {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [
            {
              description: 'Tệp dữ liệu CPMS (*.cpms, *.xlsx, *.xlsb)',
              accept: {
                'application/octet-stream': ['.cpms', '.xlsb'],
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
              }
            }
          ],
          multiple: false
        });

        if (handle) {
          this.activeFileHandle = handle;
          const file = await handle.getFile();
          this._lastKnownFileModified = file.lastModified || Date.now();
          this.startLocalFileWatcher();
          const result = await CPMSImporterExporter.importCPMSFile(file);
          const existingPrj = (this.projects || []).find((p) => p.id === result.projectData.id);
          result.projectData.data_origin = 'local';
          if (existingPrj?.drive_synced) result.projectData.drive_synced = true;
          if (existingPrj?.drive_folder_id) result.projectData.drive_folder_id = existingPrj.drive_folder_id;
          await db.put('projects', result.projectData);
          await db.saveProjectFileHandle(result.projectData.id, handle);
          this.currentProjectId = result.projectData.id;
          await this.loadData();
          this.render();
          this.showLocalSaveStatus(`Đã liên kết file "${handle.name}". Mọi chỉnh sửa sẽ tự động lưu vào file này!`);
          return;
        }
      } catch (err) {
        if (err.name === 'AbortError') return;
        console.warn('Không thể mở showOpenFilePicker, chuyển sang modal tải file:', err);
      }
    }
    // Fallback: Mở modal import file
    this.openProjectModal();
  }

  /**
   * Kiểm tra xem dự án hiện tại đã sẵn sàng (đã có dữ liệu thực hoặc đã dẫn link file local) hay chưa
   */
  isProjectReady() {
    if (!this.projects || this.projects.length === 0 || !this.currentProjectId) {
      return false;
    }
    const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
    if (!curPrj) return false;

    // 1. Nếu đang có File Handle trực tiếp trên máy tính
    if (this.activeFileHandle) return true;

    // 2. Nếu có tên file nguồn hợp lệ (không phải rỗng hay "Chưa liên kết")
    const hasSourceFile = Boolean(
      curPrj.source_filename &&
      curPrj.source_filename !== 'Chưa liên kết' &&
      curPrj.source_filename.trim()
    );

    // 3. Nếu đã có hợp đồng thực tế (không phải chỉ CPB mặc định hoặc giá trị > 0)
    const hasRealContracts = Array.isArray(this.contracts) && this.contracts.some(
      (c) => (!c.is_cpb && c.contract_type !== 'CPB') || (c.contract_val && c.contract_val > 0)
    );

    // Dự án rỗng placeholder (PRJ_DEFAULT hoặc tên có chứa "(Chưa có dữ liệu)")
    const isPlaceholderPrj = (
      curPrj.id === 'PRJ_DEFAULT' ||
      (curPrj.name && curPrj.name.includes('(Chưa có dữ liệu)'))
    );

    if (isPlaceholderPrj && !hasSourceFile && !hasRealContracts) {
      return false;
    }

    return hasSourceFile || hasRealContracts;
  }

  /**
   * Cập nhật trạng thái hiển thị giao diện:
   * Khi chưa có dự án và chưa dẫn link local: Ẩn cả 3 tab (Tổng quan, Hợp đồng, Giải ngân), hiển thị Empty State.
   */
  updateEmptyStateUI(currentTab = null) {
    const ready = this.isProjectReady();
    const emptyContainer = document.getElementById('cpms-empty-state-container');
    const overviewPane = document.getElementById('cpms-tab-overview');
    const contractsPane = document.getElementById('cpms-tab-contracts');
    const disbursementsPane = document.getElementById('cpms-tab-disbursements');
    const reminderPane = document.getElementById('cpms-tab-reminder');
    const projectInfoBar = document.getElementById('project-info-bar');

    if (!ready) {
      // Ẩn toàn bộ 4 tab nội dung
      overviewPane?.classList.add('hidden');
      contractsPane?.classList.add('hidden');
      disbursementsPane?.classList.add('hidden');
      reminderPane?.classList.add('hidden');
      if (projectInfoBar) projectInfoBar.style.display = 'none';

      // Hiển thị khung Empty State
      emptyContainer?.classList.remove('hidden');

      // Ẩn các nút thao tác dự án không cần thiết (mở thư mục, xuất file, xóa), chỉ giữ icon dẫn file local
      const btnOpenFolder = document.getElementById('btn-open-cpms-folder');
      const btnExport = document.getElementById('btn-export-cpms');
      const btnDelete = document.getElementById('btn-delete-project');
      if (btnOpenFolder) btnOpenFolder.style.display = 'none';
      if (btnExport) btnExport.style.display = 'none';
      if (btnDelete) btnDelete.style.display = 'none';
    } else {
      // Ẩn khung Empty State
      emptyContainer?.classList.add('hidden');
      if (projectInfoBar) projectInfoBar.style.display = '';

      // Hiển thị tab đang chọn
      const activeTab = currentTab || window.app?.currentCpmsTab || 'overview';
      if (overviewPane) {
        if (activeTab === 'overview') overviewPane.classList.remove('hidden');
        else overviewPane.classList.add('hidden');
      }
      if (contractsPane) {
        if (activeTab === 'contracts') contractsPane.classList.remove('hidden');
        else contractsPane.classList.add('hidden');
      }
      if (disbursementsPane) {
        if (activeTab === 'disbursements') disbursementsPane.classList.remove('hidden');
        else disbursementsPane.classList.add('hidden');
      }
      if (reminderPane) {
        if (activeTab === 'reminder') reminderPane.classList.remove('hidden');
        else reminderPane.classList.add('hidden');
      }

      // Hiện lại các nút thao tác dự án
      const btnOpenFolder = document.getElementById('btn-open-cpms-folder');
      const btnExport = document.getElementById('btn-export-cpms');
      const btnDelete = document.getElementById('btn-delete-project');
      if (btnOpenFolder) btnOpenFolder.style.display = '';
      if (btnExport) btnExport.style.display = '';
      if (btnDelete) btnDelete.style.display = '';
    }
  }

  async init() {
    await this.loadData();
    this.bindEvents();
    this.render();
  }

  async loadData() {
    this.projects = await db.getAll('projects');
    const allContracts = await db.getAll('contracts');

    // Tự động chuẩn hóa Tên và Địa điểm dự án nếu bị lưu nhầm tên file hoặc thiếu địa điểm từ các phiên bản trước
    let updatedAnyProject = false;
    for (const prj of this.projects) {
      const isPlaceholder = !prj.location || !String(prj.location).trim() ||
        prj.location === '—' ||
        prj.name === prj.source_filename?.replace(/\.[^/.]+$/, '') ||
        /^PRJ_\d+/.test(prj.name) ||
        /CHC1_\d+_\d+/.test(prj.name);

      if (isPlaceholder) {
        const prjContracts = allContracts.filter((c) => c.project_id === prj.id);
        const hasCHC1 = prjContracts.some((c) =>
          (c.contract_no && c.contract_no.includes('CHC1')) ||
          (c.id && c.id.includes('CHC1')) ||
          (c.scope_of_work && c.scope_of_work.toLowerCase().includes('tổng thầu'))
        );

        if (hasCHC1) {
          prj.code = prj.code && !prj.code.startsWith('PRJ_') ? prj.code : 'CHC1';
          prj.name = 'Dự án Nhà ở Chung cư Cao tầng Westlake Residential';
          prj.location = 'Ô đất CHC1, Khu D6, Phú Thượng, Tây Hồ, Hà Nội';
          if (!prj.client || prj.client === 'Chủ đầu tư') {
            prj.client = 'Công ty Cổ phần Xây dựng và Kỹ thuật Việt Nam';
          }
          if (!prj.contractor) {
            prj.contractor = 'Công ty Cổ phần Đầu tư Xây dựng và Kỹ thuật VNCN E&C';
          }
          await db.put('projects', prj);
          updatedAnyProject = true;
        }
      }
    }

    if (updatedAnyProject) {
      this.projects = await db.getAll('projects');
    }

    this.updateProjectSelectorUI();

    if (this.currentProjectId && !this.activeFileHandle) {
      const savedFileHandle = await db.getProjectFileHandle(this.currentProjectId);
      if (savedFileHandle) {
        this.activeFileHandle = savedFileHandle;
        if (!this._lastKnownFileModified && typeof savedFileHandle.queryPermission === 'function') {
          try {
            const p = await savedFileHandle.queryPermission({ mode: 'read' });
            if (p === 'granted') {
              const f = await savedFileHandle.getFile();
              this._lastKnownFileModified = f.lastModified || 0;
            }
          } catch (_) {}
        }
      }
    }
    if (this.activeFileHandle) {
      this.startLocalFileWatcher();
    }

    if (!this.currentProjectId) {
      this.contracts = [];
      this.payments = [];
      this.documents = [];
      return;
    }

    const allPayments = await db.getAll('payments');
    const allDocs = await db.getAll('documents');

    // Tự động nhận diện và nạp các mục Chi phí Ban (CPB) từ payments nếu chưa có trong contracts
    const curProject = this.projects.find((p) => p.id === this.currentProjectId) || {};
    const cpbPayments = allPayments.filter(
      (p) => p.project_id === this.currentProjectId && p.contract_id && /^B-CPB/i.test(p.contract_id)
    );
    const uniqueCpbIds = [...new Set(cpbPayments.map((p) => p.contract_id))];

    for (const cpbId of uniqueCpbIds) {
      const exists = allContracts.some(
        (c) => c.project_id === this.currentProjectId && c.id === cpbId
      );
      if (!exists) {
        const headerRow = allPayments.find(
          (p) => p.project_id === this.currentProjectId && p.contract_id === cpbId && p.is_header
        );
        const firstRow = allPayments.find(
          (p) => p.project_id === this.currentProjectId && p.contract_id === cpbId
        );
        const cpbName = (headerRow && headerRow.item_name) || (firstRow && firstRow.contract_no) || 'Chi Phí Ban';

        const newCpbContract = {
          id: cpbId,
          project_id: this.currentProjectId,
          group: 'B',
          parent_id: null,
          level: 0,
          contract_no: cpbId,
          contract_type: 'CPB',
          signing_date: null,
          end_date: null,
          status: 'Active',
          company_a: curProject.client || 'Chủ đầu tư',
          company_b: 'Ban Quản lý / Nội bộ',
          company_b2: '',
          rep_b2: '',
          scope_of_work: cpbName,
          exchange_rate: 1.0,
          estimated_val: 0,
          contract_val: 0,
          rep_a: '',
          rep_b: '',
          category: 'CPB',
          payee: 'Ban QLDA / Nội bộ',
          bank_account: '',
          bank_name: '',
          contact_person_b: '',
          created_at: new Date().toISOString().slice(0, 10),
          display_no: 'CPB',
          sort_order: 9999,
          is_cpb: true,
        };
        await db.put('contracts', newCpbContract);
        allContracts.push(newCpbContract);
      }
    }

    // Lọc theo currentProjectId
    this.contracts = allContracts.filter(
      (c) => c.project_id === this.currentProjectId
    );

    // Tự động nhận diện và cập nhật Mã Code (code_a, code_b, code_b2) cho các hợp đồng đang bị thiếu code từ dữ liệu CPMS local
    for (const c of this.contracts) {
      let changed = false;
      if (!c.code_a && c.company_a) {
        const foundCodeA = this.getCompanyCodeByName(c.company_a, 'A');
        if (foundCodeA) {
          c.code_a = foundCodeA;
          changed = true;
        }
      }
      if (!c.code_b && c.company_b) {
        const foundCodeB = this.getCompanyCodeByName(c.company_b, 'B');
        if (foundCodeB) {
          c.code_b = foundCodeB;
          changed = true;
        }
      }
      if (!c.code_b2 && c.company_b2) {
        const foundCodeB2 = this.getCompanyCodeByName(c.company_b2, 'B');
        if (foundCodeB2) {
          c.code_b2 = foundCodeB2;
          changed = true;
        }
      }
      if (changed) {
        await db.put('contracts', c);
      }
    }
    this.payments = allPayments.filter(
      (p) => p.project_id === this.currentProjectId
    );

    // Chuẩn hóa dữ liệu đợt tạm ứng: Khối lượng hoàn thành và KL chấp nhận thanh toán của đợt tạm ứng luôn bằng 0
    // Đồng thời tự động chuẩn hóa và đồng bộ nội dung thanh toán (item_name & payment_ref)
    for (const p of this.payments) {
      if (p.is_header) continue;
      let changed = false;
      const isAdvance = (p.item_name && /tạm\s*ứng/i.test(p.item_name)) ||
                        (p.payment_ref && /tạm\s*ứng/i.test(p.payment_ref));
      if (isAdvance && (p.actual_volume !== 0 || p.accepted_vol !== 0)) {
        p.actual_volume = 0;
        p.accepted_vol = 0;
        changed = true;
      }

      const rawName = (p.item_name !== undefined && p.item_name !== null) ? String(p.item_name).trim() : '';
      const rawRef = (p.payment_ref !== undefined && p.payment_ref !== null) ? String(p.payment_ref).trim() : '';
      const rawNote = (p.note !== undefined && p.note !== null) ? String(p.note).trim() : '';

      if (!rawName && rawRef) {
        p.item_name = rawRef;
        changed = true;
      } else if (!rawRef && rawName) {
        p.payment_ref = rawName;
        changed = true;
      } else if (!rawName && !rawRef && rawNote) {
        p.item_name = rawNote;
        p.payment_ref = rawNote;
        changed = true;
      }

      if (changed) {
        await db.put('payments', p);
      }
    }
    this.documents = allDocs.filter(
      (d) => d.project_id === this.currentProjectId
    );

    // Tự động kiểm tra và di chuyển/bổ sung dữ liệu supporting_docs cho các hồ sơ chưa có
    const docsNeedingMigration = this.documents.filter((d) => !d.supporting_docs);
    if (docsNeedingMigration.length > 0 && window.XLSX) {
      try {
        let wb = null;
        if (this.activeFileHandle) {
          const file = await this.activeFileHandle.getFile();
          const ab = await file.arrayBuffer();
          wb = window.XLSX.read(ab, { type: 'array' });
        } else {
          const resp = await fetch('./legacy/cpms-plus/excel/CHC1_20261118_103621.cpms');
          if (resp.ok) {
            const ab = await resp.arrayBuffer();
            wb = window.XLSX.read(ab, { type: 'array' });
          }
        }
        if (wb && wb.SheetNames.includes('DOCUMENT_TRACKER')) {
          const ws = wb.Sheets['DOCUMENT_TRACKER'];
          const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
          const docMap = new Map();
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (!r || !r[0]) continue;
            const dId = String(r[0]).trim();
            docMap.set(dId, {
              supporting_docs: String(r[12] || ''),
              lan_val: r[13] || 1,
              issuer: String(r[16] || ''),
              recipients: String(r[17] || ''),
              sub_type1: String(r[18] || ''),
              sub_type2: String(r[19] || ''),
            });
          }
          for (const d of docsNeedingMigration) {
            const info = docMap.get(d.id);
            if (info && info.supporting_docs) {
              d.supporting_docs = info.supporting_docs;
              if (!d.lan_val) d.lan_val = info.lan_val;
              if (!d.issuer) d.issuer = info.issuer;
              if (!d.recipients) d.recipients = info.recipients;
              if (!d.sub_type1) d.sub_type1 = info.sub_type1;
              if (!d.sub_type2) d.sub_type2 = info.sub_type2;
              await db.put('documents', d);
            }
          }
        }
      } catch (e) {
        console.warn('[MAIO Migration] Không thể tự động đọc dữ liệu bổ sung cho hồ sơ:', e);
      }
    }

    // Tự động gán ngày nhận mặc định là ngày ký cho các hồ sơ hợp đồng, phụ lục nếu chưa có ngày nhận
    for (const d of this.documents) {
      const cat = String(d.category || '').trim().toUpperCase();
      const isContractDoc = cat === 'CON' || cat === 'CONTRACT' || cat === 'HỢP ĐỒNG' || cat === 'PHỤ LỤC' || cat === 'PLHD' ||
        String(d.title || '').toLowerCase().includes('contract') ||
        String(d.title || '').toLowerCase().includes('hợp đồng') ||
        String(d.title || '').toLowerCase().includes('phụ lục');

      if (isContractDoc && (!d.received_date || String(d.received_date).trim() === '' || d.received_date === '—')) {
        const targetC = this.contracts.find((c) => c.id === d.contract_id);
        if (targetC && targetC.signing_date) {
          d.received_date = targetC.signing_date;
          await db.put('documents', d);
        } else if (d.start_date) {
          d.received_date = d.start_date;
          await db.put('documents', d);
        }
      } else if ((!d.received_date || String(d.received_date).trim() === '' || d.received_date === '—') && d.doc_no) {
        const m = String(d.doc_no).trim().match(/^(\d{2})(\d{2})(\d{2})\./);
        if (m) {
          const yr = Number(m[1]) > 50 ? `19${m[1]}` : `20${m[1]}`;
          d.received_date = `${yr}-${m[2]}-${m[3]}`;
          await db.put('documents', d);
        }
      }
    }

    // Tự động cập nhật số lượng hồ sơ nhắc việc trên thanh tab Reminder
    this.updateReminderBadge();
  }

  getBaseName(path) {
    if (!path) return '';
    const clean = String(path).replace(/[/\\]+$/, '');
    const parts = clean.split(/[/\\]/);
    return parts[parts.length - 1] || clean;
  }

  getFolderPath(filePath) {
    if (!filePath) return '';
    const clean = String(filePath).replace(/[/\\]+$/, '');
    const lastSegment = clean.split(/[/\\]/).pop() || '';
    // Chỉ cắt lấy thư mục cha nếu phần cuối là tên tệp tin có đuôi mở rộng (như .pdf, .xlsx, .docx, .zip...)
    if (/\.[a-zA-Z0-9]{2,5}$/.test(lastSegment)) {
      const idx = Math.max(clean.lastIndexOf('/'), clean.lastIndexOf('\\'));
      if (idx !== -1) {
        return clean.substring(0, idx);
      }
    }
    return clean;
  }

  /**
   * Tìm và mở file trực tiếp trong trình duyệt từ thư mục Local đã cấp quyền (File System Access API)
   */
  async getAuthorizedLocalDirHandle() {
    let dirHandle = window.__MAIO_APP__?.localDirectoryHandle || null;
    if (!dirHandle && this.currentProjectId && typeof db.getProjectDirHandle === 'function') {
      dirHandle = await db.getProjectDirHandle(this.currentProjectId);
      if (dirHandle && window.__MAIO_APP__) {
        window.__MAIO_APP__.localDirectoryHandle = dirHandle;
      }
    }
    if (!dirHandle || typeof dirHandle.queryPermission !== 'function') return null;

    let perm = await dirHandle.queryPermission({ mode: 'read' });
    if (perm === 'prompt' && typeof dirHandle.requestPermission === 'function') {
      perm = await dirHandle.requestPermission({ mode: 'read' });
    }
    if (perm !== 'granted') return null;
    return dirHandle;
  }

  isInlinePreviewableFile(fileName) {
    return /\.(pdf|png|jpe?g|gif|webp|svg|txt|json|xml|html?)$/i.test(String(fileName || ''));
  }

  isExcelPreviewableFile(fileName) {
    return /\.(xlsx|xlsm|xls|xlsb|csv)$/i.test(String(fileName || ''));
  }

  async showExcelPreviewModal(fileObj, doc = null) {
    if (typeof XLSX === 'undefined') {
      this.showToast('Thư viện đọc Excel (SheetJS) chưa sẵn sàng.', 'warning');
      return false;
    }

    const existing = document.getElementById('maio-excel-preview-modal');
    if (existing) existing.remove();

    const buffer = await fileObj.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheetNames = workbook.SheetNames || [];
    if (!sheetNames.length) {
      this.showToast(`Tệp "${fileObj.name}" không có trang tính nào.`, 'warning');
      return false;
    }

    const esc = (s) => String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

    const modal = document.createElement('div');
    modal.id = 'maio-excel-preview-modal';
    modal.className = 'modal-overlay';
    modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,0.6);backdrop-filter:blur(3px);z-index:10000;display:flex;align-items:center;justify-content:center;padding:16px;';

    modal.innerHTML = `
      <div style="background:var(--bg-card);color:var(--text-main);border:1px solid var(--border);border-radius:12px;width:96vw;max-width:1280px;height:88vh;display:flex;flex-direction:column;box-shadow:0 24px 60px rgba(0,0,0,0.3);overflow:hidden;">
        <div style="padding:12px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;">
          <div style="display:flex;align-items:center;gap:10px;min-width:0;">
            <i class="fa-solid fa-file-excel" style="color:#10b981;font-size:1.2rem;"></i>
            <span style="font-weight:700;font-size:0.98rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(fileObj.name)}</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
            ${doc ? `<button type="button" id="btn-excel-open-folder" class="button button-outline button-sm"><i class="fa-solid fa-folder-open" style="color:#f59e0b;"></i> Mở thư mục</button>` : ''}
            <button type="button" id="btn-excel-download-copy" class="button button-outline button-sm"><i class="fa-solid fa-download"></i> Tải về máy</button>
            <button type="button" id="btn-excel-preview-close" class="button button-outline button-sm"><i class="fa-solid fa-xmark"></i> Đóng</button>
          </div>
        </div>
        <div id="maio-excel-sheet-tabs" style="padding:8px 14px;border-bottom:1px solid var(--border);display:flex;gap:6px;overflow-x:auto;background:var(--bg-app);">
          ${sheetNames.map((name, idx) => `
            <button type="button" class="button button-sm ${idx === 0 ? 'button-primary' : 'button-outline'} maio-excel-tab-btn" data-sheet="${esc(name)}" style="white-space:nowrap;">
              ${esc(name)}
            </button>
          `).join('')}
        </div>
        <div id="maio-excel-sheet-body" style="flex:1;overflow:auto;padding:12px 16px;font-size:0.82rem;"></div>
      </div>
    `;

    const styleId = 'maio-excel-preview-style';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        #maio-excel-sheet-body table { border-collapse: collapse; width: max-content; min-width: 100%; }
        #maio-excel-sheet-body td, #maio-excel-sheet-body th {
          border: 1px solid var(--border);
          padding: 5px 8px;
          white-space: pre-wrap;
          vertical-align: middle;
          color: var(--text-main);
        }
        #maio-excel-sheet-body tr:first-child td, #maio-excel-sheet-body th {
          background: var(--bg-app);
          font-weight: 600;
        }
      `;
      document.head.appendChild(style);
    }

    document.body.appendChild(modal);

    const renderSheet = (sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const bodyEl = modal.querySelector('#maio-excel-sheet-body');
      if (!ws || !bodyEl) return;
      bodyEl.innerHTML = XLSX.utils.sheet_to_html(ws, { editable: false });
      modal.querySelectorAll('.maio-excel-tab-btn').forEach(btn => {
        const active = btn.dataset.sheet === sheetName;
        btn.classList.toggle('button-primary', active);
        btn.classList.toggle('button-outline', !active);
      });
    };

    renderSheet(sheetNames[0]);

    modal.querySelectorAll('.maio-excel-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => renderSheet(btn.dataset.sheet));
    });

    const closeModal = () => modal.remove();
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal();
    });
    modal.querySelector('#btn-excel-preview-close')?.addEventListener('click', closeModal);

    modal.querySelector('#btn-excel-open-folder')?.addEventListener('click', async () => {
      closeModal();
      await this.handleOpenDocFolder(doc);
    });

    modal.querySelector('#btn-excel-download-copy')?.addEventListener('click', () => {
      const blobUrl = URL.createObjectURL(fileObj);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = fileObj.name;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
    });

    return true;
  }

  async openFileObjectDirectly(fileObj, preOpenedWin = null, doc = null) {
    const blobUrl = URL.createObjectURL(fileObj);
    const canPreviewInline = this.isInlinePreviewableFile(fileObj.name);

    // Chỉ mở xem trực tiếp các định dạng trình duyệt hỗ trợ (PDF, ảnh, TXT, HTML...)
    if (canPreviewInline) {
      if (preOpenedWin && !preOpenedWin.closed) {
        preOpenedWin.location.href = blobUrl;
        setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
        this.showToast(`Đã mở tệp tin trực tiếp: "${fileObj.name}"`, 'success');
        return true;
      }

      const openedWin = window.open(blobUrl, '_blank');
      if (openedWin) {
        setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
        this.showToast(`Đã mở tệp tin trực tiếp: "${fileObj.name}"`, 'success');
        return true;
      }
    } else if (preOpenedWin && !preOpenedWin.closed) {
      preOpenedWin.close();
    }

    URL.revokeObjectURL(blobUrl);
    if (!canPreviewInline) {
      this.showToast(`Định dạng "${fileObj.name}" không xem trực tiếp trên web. Hãy chuột phải vào file trong cửa sổ thư mục và chọn Open để mở bằng ứng dụng máy tính.`, 'info');
      return false;
    }

    // Chỉ fallback tải xuống nếu là file xem inline nhưng Popup Blocker chặn mở tab mới
    const fallbackBlobUrl = URL.createObjectURL(fileObj);
    const a = document.createElement('a');
    a.href = fallbackBlobUrl;
    a.download = fileObj.name;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(fallbackBlobUrl), 60000);

    this.showToast(`Trình duyệt chặn mở tab mới, đã tải xuống: "${fileObj.name}"`, 'warning');
    return true;
  }

  async tryOpenFileFromLocalDirHandle(targetPath, preOpenedWin = null, doc = null) {
    try {
      const dirHandle = await this.getAuthorizedLocalDirHandle();
      if (!dirHandle) return false;

      const normTarget = String(targetPath || '').replace(/\\/g, '/').toLowerCase();
      const targetName = this.getBaseName(targetPath).toLowerCase();
      if (!targetName) return false;

      const searchInDir = async (handle, relPrefix = '', depth = 0) => {
        if (depth > 8) return null;
        for await (const [name, entry] of handle.entries()) {
          const entryRel = relPrefix ? `${relPrefix}/${name}` : name;
          if (entry.kind === 'file') {
            if (name.toLowerCase() === targetName) {
              return entry;
            }
          } else if (entry.kind === 'directory') {
            if (normTarget.includes(`/${name.toLowerCase()}/`) || depth < 4) {
              const found = await searchInDir(entry, entryRel, depth + 1);
              if (found) return found;
            }
          }
        }
        return null;
      };

      const fileHandle = await searchInDir(dirHandle);
      if (!fileHandle) return false;

      const fileObj = await fileHandle.getFile();
      return await this.openFileObjectDirectly(fileObj, preOpenedWin, doc);
    } catch (err) {
      console.warn('Không thể đọc file từ thư mục Local đã cấp quyền:', err);
      return false;
    }
  }

  async resolveSubDirHandleByPath(rootHandle, folderPath, targetFileName = '') {
    const segments = String(folderPath || '')
      .replace(/\\/g, '/')
      .split('/')
      .map(s => s.trim())
      .filter(Boolean);
    if (!segments.length) return rootHandle;

    const rootNameLower = String(rootHandle.name || '').toLowerCase();
    const rootIdx = segments.findIndex(s => s.toLowerCase() === rootNameLower);
    if (rootIdx >= 0) {
      let cur = rootHandle;
      let matchedAll = true;
      for (let i = rootIdx + 1; i < segments.length; i++) {
        const segLower = segments[i].toLowerCase();
        let nextDir = null;
        for await (const [name, entry] of cur.entries()) {
          if (entry.kind === 'directory' && name.toLowerCase() === segLower) {
            nextDir = entry;
            break;
          }
        }
        if (!nextDir) {
          matchedAll = false;
          break;
        }
        cur = nextDir;
      }
      if (matchedAll) return cur;
    }

    // Fallback 1: Nếu biết tên file, tìm chính xác thư mục cha chứa file đó trong dirHandle
    const fileLower = String(targetFileName || '').trim().toLowerCase();
    const normFolder = `/${String(folderPath || '').replace(/\\/g, '/').toLowerCase()}/`;
    if (fileLower) {
      const findParentOfFile = async (handle, depth = 0) => {
        if (depth > 8) return null;
        for await (const [name, entry] of handle.entries()) {
          if (entry.kind === 'file' && name.toLowerCase() === fileLower) {
            return handle;
          }
        }
        for await (const [name, entry] of handle.entries()) {
          if (entry.kind === 'directory') {
            if (normFolder.includes(`/${name.toLowerCase()}/`) || depth < 4) {
              const found = await findParentOfFile(entry, depth + 1);
              if (found) return found;
            }
          }
        }
        return null;
      };
      const parentByFile = await findParentOfFile(rootHandle);
      if (parentByFile) return parentByFile;
    }

    // Fallback 2: Tìm theo tên thư mục đích cuối cùng trong cây thư mục đã cấp quyền
    const targetFolderName = segments[segments.length - 1].toLowerCase();
    if (targetFolderName === rootNameLower) return rootHandle;

    const findDirByName = async (handle, depth = 0) => {
      if (depth > 8) return null;
      for await (const [name, entry] of handle.entries()) {
        if (entry.kind === 'directory') {
          if (name.toLowerCase() === targetFolderName) return entry;
          if (normFolder.includes(`/${name.toLowerCase()}/`) || depth < 4) {
            const found = await findDirByName(entry, depth + 1);
            if (found) return found;
          }
        }
      }
      return null;
    };

    return (await findDirByName(rootHandle)) || rootHandle;
  }

  async showLocalDirectoryBrowserModal(dirHandle, originalPath = '', parentStack = []) {
    const existing = document.getElementById('maio-local-folder-modal');
    if (existing) existing.remove();

    const dirs = [];
    const files = [];
    for await (const [name, entry] of dirHandle.entries()) {
      if (entry.kind === 'directory') {
        dirs.push({ name, handle: entry });
      } else if (entry.kind === 'file') {
        files.push({ name, handle: entry });
      }
    }
    dirs.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    files.sort((a, b) => a.name.localeCompare(b.name, 'vi'));

    const modal = document.createElement('div');
    modal.id = 'maio-local-folder-modal';
    modal.className = 'modal-overlay';
    modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,0.55);backdrop-filter:blur(3px);z-index:9999;display:flex;align-items:center;justify-content:center;padding:16px;';

    const esc = (s) => String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
    const itemsHtml = [
      ...dirs.map((d, idx) => `
        <div class="local-dir-row" data-kind="dir" data-idx="${idx}" style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--border);cursor:pointer;border-radius:6px;">
          <i class="fa-solid fa-folder" style="color:#f59e0b;font-size:1.05rem;"></i>
          <span style="flex:1;font-weight:600;color:var(--text-main);word-break:break-word;">${esc(d.name)}</span>
          <span style="font-size:0.75rem;color:var(--text-muted);">Thư mục</span>
        </div>
      `),
      ...files.map((f, idx) => `
        <div class="local-file-row" data-kind="file" data-idx="${idx}" style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--border);cursor:pointer;border-radius:6px;">
          <i class="fa-regular fa-file-lines" style="color:#3b82f6;font-size:1.05rem;"></i>
          <span style="flex:1;color:var(--text-main);word-break:break-word;">${esc(f.name)}</span>
          <button type="button" class="button button-outline button-sm" style="pointer-events:none;font-size:0.75rem;">
            <i class="fa-solid fa-arrow-up-right-from-square"></i> Mở
          </button>
        </div>
      `)
    ].join('');

    modal.innerHTML = `
      <div style="background:var(--bg-card);color:var(--text-main);border:1px solid var(--border);border-radius:12px;width:100%;max-width:680px;max-height:82vh;display:flex;flex-direction:column;box-shadow:0 20px 50px rgba(0,0,0,0.25);overflow:hidden;">
        <div style="padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;">
          <div style="min-width:0;">
            <div style="font-weight:700;font-size:1rem;display:flex;align-items:center;gap:8px;">
              <i class="fa-solid fa-folder-open" style="color:#f59e0b;"></i>
              <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(dirHandle.name)}</span>
            </div>
            ${originalPath ? `<div style="font-size:0.76rem;color:var(--text-muted);margin-top:2px;word-break:break-all;">${esc(originalPath)}</div>` : ''}
          </div>
          <div style="display:flex;align-items:center;gap:6px;flex-shrink:0;">
            ${parentStack.length > 0 ? `<button type="button" id="btn-local-dir-back" class="button button-outline button-sm"><i class="fa-solid fa-arrow-left"></i> Lên</button>` : ''}
            <button type="button" id="btn-local-dir-close" class="button button-outline button-sm"><i class="fa-solid fa-xmark"></i></button>
          </div>
        </div>
        <div style="padding:10px 14px;overflow-y:auto;flex:1;">
          ${itemsHtml || '<div style="padding:24px;text-align:center;color:var(--text-muted);">Thư mục này đang trống.</div>'}
        </div>
        <div style="padding:12px 18px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:10px;background:var(--bg-app);">
          <span style="font-size:0.78rem;color:var(--text-muted);">${dirs.length} thư mục • ${files.length} tệp tin</span>
          <div style="display:flex;gap:8px;">
            ${originalPath ? `<button type="button" id="btn-local-dir-copy" class="button button-outline button-sm"><i class="fa-regular fa-copy"></i> Copy đường dẫn</button>` : ''}
            <button type="button" id="btn-local-dir-done" class="button button-primary button-sm">Đóng</button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeModal = () => modal.remove();
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal();
    });
    modal.querySelector('#btn-local-dir-close')?.addEventListener('click', closeModal);
    modal.querySelector('#btn-local-dir-done')?.addEventListener('click', closeModal);

    modal.querySelector('#btn-local-dir-back')?.addEventListener('click', () => {
      const prev = parentStack[parentStack.length - 1];
      if (prev) {
        this.showLocalDirectoryBrowserModal(prev, originalPath, parentStack.slice(0, -1));
      }
    });

    modal.querySelector('#btn-local-dir-copy')?.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(originalPath);
        this.showToast('Đã sao chép đường dẫn thư mục vào Clipboard!', 'success');
      } catch (_) {}
    });

    modal.querySelectorAll('.local-dir-row').forEach(row => {
      row.addEventListener('click', () => {
        const idx = Number(row.dataset.idx);
        const target = dirs[idx];
        if (target) {
          this.showLocalDirectoryBrowserModal(target.handle, originalPath, [...parentStack, dirHandle]);
        }
      });
    });

    modal.querySelectorAll('.local-file-row').forEach(row => {
      row.addEventListener('click', async () => {
        const idx = Number(row.dataset.idx);
        const target = files[idx];
        if (!target) return;
        const preWin = this.isInlinePreviewableFile(target.name) ? window.open('', '_blank') : null;
        try {
          const fileObj = await target.handle.getFile();
          await this.openFileObjectDirectly(fileObj, preWin);
        } catch (err) {
          if (preWin && !preWin.closed) preWin.close();
          this.showToast('Không thể đọc tệp tin từ thư mục đã chọn.', 'error');
        }
      });
    });
  }

  async tryOpenFolderFromLocalDirHandle(folderPath, targetFileName = '') {
    try {
      const dirHandle = await this.getAuthorizedLocalDirHandle();
      if (!dirHandle) return false;

      const targetDir = await this.resolveSubDirHandleByPath(dirHandle, folderPath, targetFileName);
      if (!targetDir) return false;

      // Mở thẳng cửa sổ File Explorer chuẩn của Windows tại đúng thư mục đã tìm thấy (giống nút mở thư mục CPMS)
      if ('showOpenFilePicker' in window) {
        try {
          const [picked] = await window.showOpenFilePicker({
            startIn: targetDir,
            multiple: false
          });
          if (picked) {
            const fileObj = await picked.getFile();
            const preWin = this.isInlinePreviewableFile(fileObj.name) ? window.open('', '_blank') : null;
            await this.openFileObjectDirectly(fileObj, preWin);
          }
          return true;
        } catch (pickerErr) {
          if (pickerErr.name === 'AbortError') {
            return true;
          }
          console.warn('showOpenFilePicker fallback sang Modal duyệt thư mục:', pickerErr);
        }
      }

      await this.showLocalDirectoryBrowserModal(targetDir, folderPath, []);
      return true;
    } catch (err) {
      console.warn('Không thể mở thư mục từ quyền Local đã cấp:', err);
      return false;
    }
  }

  async openLocalPath(targetPath, isFolder = false, action = 'open', preOpenedWin = null, targetFileName = '', doc = null) {
    if (!targetPath || !String(targetPath).trim()) {
      if (preOpenedWin && !preOpenedWin.closed) preOpenedWin.close();
      this.showToast('Chưa có thông tin đường dẫn cục bộ trên máy tính!', 'warning');
      return false;
    }
    const cleanPath = String(targetPath).trim();
    const protocolAction = isFolder ? 'folder' : (action === 'select' ? 'select' : 'open');

    // Nếu là Windows Desktop và đang Mở file (!isFolder): kiểm tra MAIO-Open (đã cài -> mở thẳng; lần đầu -> nhắc cài)
    // Tùy chọn Mở thư mục (isFolder) giữ nguyên luồng cũ (showOpenFilePicker)
    if (!isFolder && maioOpen.isWindowsDesktop()) {
      const maioOpenStatus = maioOpen.getStatus();
      if (maioOpenStatus === 'installed') {
        if (preOpenedWin && !preOpenedWin.closed) preOpenedWin.close();
        maioOpen.openViaProtocol(protocolAction, cleanPath);
        this.showToast(`Đang mở trực tiếp trên máy tính qua MAIO-Open: "${this.getBaseName(cleanPath)}"`, 'success');
        return true;
      }

      if (!maioOpenStatus) {
        if (preOpenedWin && !preOpenedWin.closed) preOpenedWin.close();
        const choice = await maioOpen.promptInstallModal({
          actionLabel: 'Mở file',
        });
        if (choice === 'installed') {
          maioOpen.openViaProtocol(protocolAction, cleanPath);
          this.showToast(`Đang mở trực tiếp trên máy tính qua MAIO-Open: "${this.getBaseName(cleanPath)}"`, 'success');
          return true;
        }
        if (choice === 'cancel') {
          return false;
        }
        if (choice === 'browser_only') {
          if (!this.isInlinePreviewableFile(cleanPath)) {
            const folderPath = this.getFolderPath(cleanPath) || cleanPath;
            const openedFolder = await this.tryOpenFolderFromLocalDirHandle(folderPath, this.getBaseName(cleanPath));
            if (openedFolder) return true;
          }
        }
      }
    }

    // Sử dụng trực tiếp quyền thư mục Local đã cấp trên trình duyệt (chế độ trình duyệt)
    if (!isFolder) {
      const openedDirect = await this.tryOpenFileFromLocalDirHandle(cleanPath, preOpenedWin, doc);
      if (openedDirect) return true;
    } else {
      const openedFolderDirect = await this.tryOpenFolderFromLocalDirHandle(cleanPath, targetFileName);
      if (openedFolderDirect) return true;
    }

    if (preOpenedWin && !preOpenedWin.closed) {
      preOpenedWin.close();
    }

    // Fallback nếu chưa cấp quyền thư mục Local hoặc trình duyệt không hỗ trợ File System Access API
    let copied = false;
    try {
      await navigator.clipboard.writeText(cleanPath);
      copied = true;
    } catch (_) {}

    await this.alertAsync({
      title: copied
        ? `Đã sao chép đường dẫn ${isFolder ? 'thư mục' : 'tệp tin'}`
        : `Đường dẫn ${isFolder ? 'thư mục' : 'tệp tin'} trên máy tính`,
      type: copied ? 'success' : 'info',
      message: copied
        ? `Đường dẫn đã được sao chép vào Clipboard:\n\n${cleanPath}\n\nHãy vào Cài đặt -> Chọn thư mục máy tính gốc để cấp quyền mở trực tiếp trên trình duyệt, hoặc nhấn [Windows + R] / dán (Ctrl+V) vào File Explorer.`
        : `Vui lòng sao chép đường dẫn dưới đây để mở trên File Explorer:\n\n${cleanPath}`,
    });
    return false;
  }

  async handleOpenDocFile(doc) {
    if (!doc) return;
    const link = doc.local_link || doc.file_link || doc.drive_link;
    if (!link || !String(link).trim()) {
      await this.alertAsync({
        title: 'Chưa có đường dẫn tệp tin',
        type: 'warning',
        message: 'Hồ sơ này chưa có đường dẫn tệp tin (File Link)!\n\nVui lòng bấm nút Chỉnh sửa (hình cây bút) để cập nhật đường dẫn tệp tin lưu trữ.',
      });
      return;
    }
    const cleanLink = String(link).trim();
    if (cleanLink.startsWith('http://') || cleanLink.startsWith('https://')) {
      window.open(cleanLink, '_blank');
      return;
    }

    const canInline = this.isInlinePreviewableFile(cleanLink);
    const isWinDesktop = maioOpen.isWindowsDesktop();
    const maioStatus = maioOpen.getStatus();

    // Nếu là Windows Desktop và đã cài (hoặc chưa chọn bỏ qua) MAIO-Open: ưu tiên mở thẳng file gốc hoặc nhắc cài 1 lần
    if (isWinDesktop && maioStatus !== 'browser_only') {
      await this.openLocalPath(cleanLink, false, 'open', null, '', doc);
      return;
    }

    // Với các file không xem trực tiếp được trên trình duyệt (Excel, Word, zip, dwg...), chuyển thẳng sang mở thư mục chứa file đó để chuột phải -> Open làm việc trực tiếp
    if (!canInline) {
      await this.handleOpenDocFolder(doc);
      return;
    }

    // Mở trước tab trống đồng bộ nếu là file xem trực tiếp (PDF/ảnh/txt) để tránh bị Popup Blocker chặn
    const hasDirHandle = Boolean(window.__MAIO_APP__?.localDirectoryHandle);
    const preOpenedWin = hasDirHandle ? window.open('', '_blank') : null;

    await this.openLocalPath(cleanLink, false, 'open', preOpenedWin, '', doc);
  }

  async handleOpenDocFolder(doc, contract = null) {
    if (!doc) return;
    const rawLink = doc.local_link || doc.file_link || '';
    let folderPath = '';
    if (rawLink && !rawLink.startsWith('http')) {
      folderPath = this.getFolderPath(rawLink);
    }

    if (!folderPath && doc.drive_link && doc.drive_link.startsWith('http')) {
      window.open(doc.drive_link, '_blank');
      return;
    }

    if (!folderPath && contract) {
      folderPath = this.getDriveDocPathForContract(contract);
    }

    if (!folderPath) {
      await this.alertAsync({
        title: 'Chưa có thư mục lưu trữ',
        type: 'warning',
        message: 'Chưa có thông tin thư mục lưu trữ cho hồ sơ này.',
      });
      return;
    }

    // Gửi lệnh mở File Explorer hoặc hiển thị cửa sổ duyệt thư mục đã cấp quyền
    const targetFileName = (rawLink && !rawLink.startsWith('http')) ? this.getBaseName(rawLink) : '';
    await this.openLocalPath(folderPath, true, 'open', null, targetFileName);
  }

  /**
   * Lấy ngày nhận của hồ sơ (với hợp đồng, phụ lục mặc định để ngày nhận là ngày ký)
   */
  getDocReceivedDate(doc, fallbackContract = null) {
    if (!doc) return null;
    let recDate = doc.received_date;
    if (!recDate || String(recDate).trim() === '' || recDate === '—') {
      const cat = String(doc.category || '').trim().toUpperCase();
      const isContractDoc = cat === 'CON' || cat === 'CONTRACT' || cat === 'HỢP ĐỒNG' || cat === 'PHỤ LỤC' || cat === 'PLHD' ||
        String(doc.title || '').toLowerCase().includes('contract') ||
        String(doc.title || '').toLowerCase().includes('hợp đồng') ||
        String(doc.title || '').toLowerCase().includes('phụ lục');

      if (isContractDoc) {
        const targetC = (this.contracts || []).find((c) => c.id === doc.contract_id) || fallbackContract;
        if (targetC && targetC.signing_date) {
          recDate = targetC.signing_date;
        } else if (doc.start_date) {
          recDate = doc.start_date;
        }
      } else if (doc.doc_no && /^(\d{2})(\d{2})(\d{2})\./.test(String(doc.doc_no).trim())) {
        const m = String(doc.doc_no).trim().match(/^(\d{2})(\d{2})(\d{2})\./);
        if (m) {
          const yr = Number(m[1]) > 50 ? `19${m[1]}` : `20${m[1]}`;
          recDate = `${yr}-${m[2]}-${m[3]}`;
        }
      }
    }
    return recDate;
  }

  removeAccents(str) {
    if (!str || typeof str !== 'string') return '';
    return str
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .toLowerCase();
  }

  getCompanyCodeByName(name, group = 'B') {
    if (!name || typeof name !== 'string') return '';
    const n = name.trim();
    if (!n) return '';

    const upper = n.toUpperCase();
    const knownShortCodes = [
      'CLI', 'VEC', 'VNCN', 'CTC', 'KD', 'NSH', 'TID', 'HHG', 'HD', 'EC',
      'EC-KD', 'EC-TG', 'EC-PHC', 'EC-NSH', 'HTC', 'HTCONS', 'VTIC', 'VNCC',
      'ACIT', 'DPA', 'DP', 'ICE', 'TG', 'HDAT', 'HADU', 'CVL', 'PADO', 'MKD',
      'PCV', 'MN', 'SHVN', 'MS', 'TTT', 'KTM', 'KKVL', 'GLX', 'LKH', 'HIWIN',
      'ATVN', 'HISUNG', 'TP', 'GOOD', 'TPHAT', 'AUSTDOOR', 'PL', 'TT', 'PENTA',
      'GTS', 'HH', 'ABUSS', 'TA', 'TM', 'HA', 'VIETA', 'DK', 'DKE', 'HTTL'
    ];
    if (knownShortCodes.includes(upper)) return upper;

    const lower = n.toLowerCase();

    // Bên A (Chủ đầu tư / Tổng thầu)
    if (group === 'A' || lower.includes('chủ đầu tư') || lower.includes('vietnam investment') || lower.includes('cli')) {
      if (lower.includes('việt nam') || lower.includes('vina') || lower.includes('kỹ thuật việt nam') || lower.includes('vec')) return 'VEC';
      return 'CLI';
    }

    // Mapping từ điển chuẩn CPMS Plus cho các nhà thầu / đối tác Bên B
    if (lower.includes('vncn') || lower.includes('xây dựng và kỹ thuật vncn')) return 'VNCN';
    if (lower.includes('coteccons') || lower.includes('ctc')) return 'CTC';
    if (lower.includes('khang đạt') || lower.includes('kd')) return 'KD';
    if (lower.includes('nam sông hồng') || lower.includes('nsh')) return 'NSH';
    if (lower.includes('tường kính tid') || lower.includes('tid')) return 'TID';
    if (lower.includes('hhg') || lower.includes('hồng hà')) return 'HHG';
    if (lower.includes('htcons') || lower.includes('h&t') || lower.includes('ht cons')) return 'HTCONS';
    if (lower.includes('vtic')) return 'VTIC';
    if (lower.includes('vncc')) return 'VNCC';
    if (lower.includes('acit') || lower.includes('á châu')) return 'ACIT';
    if (lower.includes('dp architect') || lower.includes('kiến trúc dp')) return 'DP';
    if (lower.includes('indochine') || lower.includes('ice')) return 'ICE';
    if (lower.includes('thời gian') || lower.includes('ec-tg')) return 'TG';
    if (lower.includes('huy dương') || lower.includes('hà nội')) return 'HD';
    if (lower.includes('hoàng đạt')) return 'HDAT';
    if (lower.includes('hà dung')) return 'HADU';
    if (lower.includes('cát vạn lợi')) return 'CVL';
    if (lower.includes('việt á')) return 'VIETA';
    if (lower.includes('eusteel')) return 'EUS';
    if (lower.includes('pado')) return 'PADO';
    if (lower.includes('mikado')) return 'MKD';
    if (lower.includes('poong cheon')) return 'PCV';
    if (lower.includes('minh nghĩa')) return 'MN';
    if (lower.includes('sh vn') || lower.includes('shvn')) return 'SHVN';
    if (lower.includes('ms plus')) return 'MS';
    if (lower.includes('ttt')) return 'TTT';
    if (lower.includes('ktm')) return 'KTM';
    if (lower.includes('kim khí vật liệu')) return 'KKVL';
    if (lower.includes('galaxy')) return 'GLX';
    if (lower.includes('lim kim hai')) return 'LKH';
    if (lower.includes('hiwin')) return 'HIWIN';
    if (lower.includes('an toàn việt nam') || lower.includes('atvn')) return 'ATVN';
    if (lower.includes('hisung')) return 'HISUNG';
    if (lower.includes('thuận phát')) return 'TP';
    if (lower.includes('good')) return 'GOOD';
    if (lower.includes('trường phát')) return 'TPHAT';
    if (lower.includes('austdoor')) return 'AUSTDOOR';
    if (lower.includes('phương linh') || lower.includes('phc')) return 'PL';
    if (lower.includes('thanh tâm')) return 'TT';
    if (lower.includes('penta')) return 'PENTA';
    if (lower.includes('thủy sanh')) return 'GTS';
    if (lower.includes('huế hậu')) return 'HH';
    if (lower.includes('abuss')) return 'ABUSS';
    if (lower.includes('tuấn anh')) return 'TA';
    if (lower.includes('tân mỹ')) return 'TM';
    if (lower.includes('hà an')) return 'HA';
    if (lower.includes('dk engineering')) return 'DK';
    if (lower.includes('hoàng thành thăng long')) return 'HTTL';

    // Tra cứu trong các hợp đồng đã có cùng tên công ty trong dự án
    if (this.contracts) {
      const existing = this.contracts.find((c) =>
        (c.company_b && c.company_b.trim().toLowerCase() === lower && c.code_b) ||
        (c.company_a && c.company_a.trim().toLowerCase() === lower && c.code_a)
      );
      if (existing) {
        if (existing.company_b && existing.company_b.trim().toLowerCase() === lower && existing.code_b) return existing.code_b;
        if (existing.company_a && existing.company_a.trim().toLowerCase() === lower && existing.code_a) return existing.code_a;
      }
    }

    return '';
  }

  /**
   * Xác định nguồn của dự án ('local' hoặc 'online') để hiển thị rõ trong danh sách & cạnh tên dự án
   */
  getProjectOriginMode(project) {
    if (!project) return 'local';
    if (project.id === this.currentProjectId && this.activeFileHandle) {
      return 'local';
    }
    if (project.data_origin === 'local') return 'local';
    if (project.data_origin === 'online' || project.source === 'drive_sync') return 'online';

    const gdriveSvc = window.__MAIO_GDRIVE_SERVICE__;
    if (gdriveSvc && typeof gdriveSvc.isSignedIn === 'function' && gdriveSvc.isSignedIn()) {
      return 'online';
    }
    return 'local';
  }

  /**
   * Format tên hiển thị của dự án trong dropdown danh sách dự án với nhãn [Local] / [Online] ở phía trước
   */
  formatProjectOptionLabel(project) {
    if (!project) return '';
    const mode = this.getProjectOriginMode(project);
    const onlineList = window.__MAIO_APP__?.onlineDriveProjects || [];
    const hasOnline = Boolean(
      project.drive_synced ||
        mode === 'online' ||
        onlineList.some((o) => o.projectId === project.id)
    );
    const linkedLocal =
      (project.id === this.currentProjectId ? this.activeFileHandle?.name : '') ||
      project.source_filename ||
      '';

    if (hasOnline && linkedLocal) {
      return `☁️💻 [Online ↔ Local: ${linkedLocal}] ${project.name || project.id}`;
    }
    if (hasOnline) {
      return `☁️ [Online] ${project.name || project.id}`;
    }
    return `💻 [Local] ${project.name || project.id}`;
  }

  updateProjectSelectorUI() {
    const selector = document.getElementById('project-selector');
    if (!selector) return;

    const onlineList = window.__MAIO_APP__?.onlineDriveProjects || [];

    if (!this.projects || this.projects.length === 0) {
      this.currentProjectId = null;
      this.activeFileHandle = null;
      this._lastKnownFileModified = 0;

      let emptyOptionsHtml = '';
      if (onlineList.length > 0) {
        emptyOptionsHtml += `<option value="" disabled selected>-- Chọn dự án Online trên Google Drive (${onlineList.length}) --</option>`;
        onlineList.forEach((item) => {
          emptyOptionsHtml += `<option value="DRIVE_ONLINE:${item.fileId}" data-drive-prj-id="${item.projectId}">☁️ [Online] Dự án ${item.projectId}</option>`;
        });
      } else {
        emptyOptionsHtml += `<option value="" disabled selected>-- Chưa có dự án nào --</option>`;
      }
      emptyOptionsHtml += `<option value="CREATE_NEW">+ Tạo dự án mới...</option>`;
      emptyOptionsHtml += `<option value="IMPORT_NEW">+ Dẫn link file CPMS từ máy tính (.cpms)...</option>`;
      selector.innerHTML = emptyOptionsHtml;

      const activePrjName = document.getElementById('active-project-name');
      if (activePrjName) activePrjName.textContent = 'Chưa có dự án nào';
      const topProjName = document.getElementById('top-project-name');
      if (topProjName) topProjName.textContent = onlineList.length > 0 ? 'Chọn dự án Online từ Google Drive' : 'Chưa có dự án nào';
      const topProjLoc = document.getElementById('top-project-location');
      if (topProjLoc) topProjLoc.textContent = '—';
      if (typeof window.__MAIO_APP__?.updateProjectSourceBadge === 'function') {
        window.__MAIO_APP__.updateProjectSourceBadge();
      }
      return;
    }

    let optionsHtml = '';
    // Đảm bảo currentProjectId hợp lệ
    if (!this.projects.some((p) => p.id === this.currentProjectId)) {
      this.currentProjectId = this.projects[0].id;
    }

    this.projects.forEach((p) => {
      const isSel = p.id === this.currentProjectId ? 'selected' : '';
      optionsHtml += `<option value="${p.id}" ${isSel}>${this.formatProjectOptionLabel(p)}</option>`;
    });

    // Bổ sung các dự án Online trên Google Drive chưa có bản local trên máy này
    const localIds = new Set(this.projects.map((p) => p.id));
    const localFileNames = new Set(
      this.projects
        .flatMap((p) => [
          p.source_filename,
          p.id === this.currentProjectId ? this.activeFileHandle?.name : '',
        ])
        .filter(Boolean)
        .map((n) => String(n).trim().toLowerCase())
    );
    onlineList.forEach((item) => {
      const itemFileLower = String(item.fileName || '').trim().toLowerCase();
      const itemIdAsCpmsLower = `${String(item.projectId || '').trim().toLowerCase()}.cpms`;
      if (
        localIds.has(item.projectId) ||
        (itemFileLower && localFileNames.has(itemFileLower)) ||
        localFileNames.has(itemIdAsCpmsLower)
      ) {
        return;
      }
      optionsHtml += `<option value="DRIVE_ONLINE:${item.fileId}" data-drive-prj-id="${item.projectId}">☁️ [Online] Dự án ${item.projectId}</option>`;
    });

    optionsHtml += `<option value="CREATE_NEW">+ Tạo dự án mới...</option>`;
    optionsHtml += `<option value="IMPORT_NEW">+ Dẫn link file CPMS từ máy tính (.cpms)...</option>`;
    selector.innerHTML = optionsHtml;
    selector.value = this.currentProjectId;

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
    const topProjLoc = document.getElementById('top-project-location');
    if (topProjLoc) {
      topProjLoc.textContent = (curPrj && curPrj.location && String(curPrj.location).trim()) ? String(curPrj.location).trim() : '—';
    }
    const topProjName = document.getElementById('top-project-name');
    if (topProjName) {
      topProjName.textContent = curPrj ? (curPrj.name || curPrj.id) : 'Chưa có dự án nào';
    }

    // Cập nhật tên dự án trên Dashboard nếu có
    const activePrjName = document.getElementById('active-project-name');
    if (activePrjName) {
      if (curPrj) activePrjName.textContent = curPrj.name || curPrj.id;
    }

    if (typeof window.__MAIO_APP__?.updateProjectSourceBadge === 'function') {
      window.__MAIO_APP__.updateProjectSourceBadge();
    }
  }

  bindEvents() {
    if (this.eventsBound) return;
    this.eventsBound = true;

    const tabBtns = document.querySelectorAll('.group-tab-btn');
    tabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        tabBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        this.currentGroup = btn.getAttribute('data-group');
        this.categoryFilter = 'ALL';
        this.render();
      });
    });

    const searchInput = document.getElementById('contract-search-input');
    const clearContractSearchBtn = document.getElementById('btn-clear-contract-search');
    searchInput?.addEventListener('input', (e) => {
      this.searchQuery = e.target.value.trim();
      if (clearContractSearchBtn) {
        clearContractSearchBtn.classList.toggle('hidden', !e.target.value);
      }
      this.render();
    });

    clearContractSearchBtn?.addEventListener('click', () => {
      if (searchInput) {
        searchInput.value = '';
        searchInput.focus();
      }
      this.searchQuery = '';
      clearContractSearchBtn.classList.add('hidden');
      this.render();
    });

    const catFilter = document.getElementById('contract-category-filter');
    catFilter?.addEventListener('change', (e) => {
      this.categoryFilter = e.target.value;
      this.render();
    });

    document.getElementById('btn-add-contract')?.addEventListener('click', () => {
      this.openContractFormModal();
    });

    document.getElementById('btn-add-cpb')?.addEventListener('click', () => {
      this.openCreateCPBModal();
    });

    // Lắng nghe chọn dự án
    const projectSelector = document.getElementById('project-selector');
    projectSelector?.addEventListener('change', async (e) => {
      const val = e.target.value;
      if (val === 'CREATE_NEW') {
        projectSelector.value = this.currentProjectId || '';
        this.openCreateProjectModal();
      } else if (val === 'IMPORT_NEW') {
        // Trả lại giá trị trước đó trên UI
        projectSelector.value = this.currentProjectId;
        this.pickAndLoadLocalCPMSFile();
      } else if (String(val).startsWith('DRIVE_ONLINE:')) {
        const fileId = val.slice('DRIVE_ONLINE:'.length);
        const selOpt = projectSelector.selectedOptions?.[0];
        const drivePrjId = selOpt?.getAttribute('data-drive-prj-id') || '';
        if (typeof window.__MAIO_APP__?.openOnlineDriveProject === 'function') {
          await window.__MAIO_APP__.openOnlineDriveProject(fileId, drivePrjId);
        }
      } else {
        this.activeFileHandle = null;
        this._lastKnownFileModified = 0;
        this.currentProjectId = val;
        await this.loadData();
        this.render();
        if (typeof window.__MAIO_APP__?.checkDriveProjectUpdates === 'function') {
          window.__MAIO_APP__.checkDriveProjectUpdates(val);
        }
      }
    });

    // Nút Tạo dự án mới
    document.getElementById('btn-create-project')?.addEventListener('click', () => {
      this.openCreateProjectModal();
    });

    // Nút Chỉnh sửa thông tin dự án trên thanh dự án (luôn hiện ở cả 4 tab)
    document.getElementById('btn-top-edit-project')?.addEventListener('click', () => {
      const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
      if (!curPrj) {
        this.showToast('Vui lòng chọn hoặc tạo dự án trước khi chỉnh sửa!', 'warning');
        return;
      }
      this.openEditProjectModal(curPrj);
    });

    // Nút Import .cpms (Dẫn file Local trực tiếp)
    document.getElementById('btn-import-project')?.addEventListener('click', () => {
      this.pickAndLoadLocalCPMSFile();
    });

    // Lắng nghe sự kiện chuyển nhanh đến Hợp đồng liên kết
    document.addEventListener('click', (e) => {
      const jumpBtn = e.target.closest('.btn-jump-linked-contract');
      if (jumpBtn) {
        e.preventDefault();
        e.stopPropagation();
        const targetId = jumpBtn.getAttribute('data-jump-id');
        this.jumpToContract(targetId);
      }
    });

    // Nút Mở thư mục chứa file CPMS đang liên kết
    document.getElementById('btn-open-cpms-folder')?.addEventListener('click', async () => {
      if (!this.activeFileHandle && this.currentProjectId) {
        const saved = await db.getProjectFileHandle(this.currentProjectId);
        if (saved) this.activeFileHandle = saved;
      }
      const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
      const fullPath = curPrj?.source_full_path || '';

      if (fullPath) {
        await this.openLocalPath(fullPath, false, 'select');
        return;
      }

      if (this.activeFileHandle && 'showOpenFilePicker' in window) {
        try {
          const [picked] = await window.showOpenFilePicker({
            startIn: this.activeFileHandle,
            types: [
              {
                description: 'Tệp dữ liệu CPMS (*.cpms, *.xlsx, *.xlsb)',
                accept: {
                  'application/octet-stream': ['.cpms', '.xlsb'],
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                },
              },
            ],
            multiple: false,
          });
          if (picked && picked.name !== this.activeFileHandle.name) {
            this.activeFileHandle = picked;
            const file = await picked.getFile();
            this._lastKnownFileModified = file.lastModified || Date.now();
            const result = await CPMSImporterExporter.importCPMSFile(file);
            result.projectData.data_origin = 'local';
            await db.put('projects', result.projectData);
            await db.saveProjectFileHandle(result.projectData.id, picked);
            this.currentProjectId = result.projectData.id;
            await this.loadData();
            this.render();
          }
          return;
        } catch (e) {
          if (e.name === 'AbortError') return;
        }
      }

      alert('Chưa có file CPMS Local nào đang được liên kết cho dự án này. Vui lòng bấm "Dẫn file Local từ máy tính" trước.');
    });

    // Nút Xuất CPMS (Lưu file .cpms vào thư mục đang được gắn link)
    document.getElementById('btn-export-cpms')?.addEventListener('click', async () => {
      const btn = document.getElementById('btn-export-cpms');
      const origText = btn ? btn.innerHTML : '';
      try {
        if (btn) {
          btn.disabled = true;
          btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
        }

        if (!this.activeFileHandle && this.currentProjectId) {
          const saved = await db.getProjectFileHandle(this.currentProjectId);
          if (saved) this.activeFileHandle = saved;
        }

        const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
        const now = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const ts = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        const prjCode = String(curPrj?.code || this.currentProjectId || 'MAIO').replace(/[\/\\?%*:|"<>]/g, '_');
        const suggestedName = `${prjCode}_${ts}.cpms`;

        if ('showSaveFilePicker' in window) {
          try {
            const saveHandle = await window.showSaveFilePicker({
              suggestedName,
              startIn: this.activeFileHandle || 'documents',
              types: [
                {
                  description: 'Tệp dữ liệu CPMS (*.cpms)',
                  accept: { 'application/octet-stream': ['.cpms'] },
                },
              ],
            });
            const buffer = await CPMSImporterExporter.buildCPMSBinaryBuffer(this.currentProjectId);
            const writable = await saveHandle.createWritable();
            await writable.write(buffer);
            await writable.close();
            this.showLocalSaveStatus(`Đã xuất file "${saveHandle.name}" thành công!`);
            return;
          } catch (pickerErr) {
            if (pickerErr.name === 'AbortError') return;
            console.warn('Fallback sang tải xuống trình duyệt:', pickerErr);
          }
        }

        const exportedFile = await CPMSImporterExporter.exportToCPMSFormat(this.currentProjectId, 'cpms');
        this.showLocalSaveStatus(`Đã xuất file "${exportedFile}" thành công!`);
      } catch (err) {
        alert('Lỗi xuất file CPMS: ' + err.message);
      } finally {
        if (btn) {
          btn.disabled = false;
          btn.innerHTML = origText;
        }
      }
    });

    // Nút Xóa Dự án
    document.getElementById('btn-delete-project')?.addEventListener('click', () => {
      this.deleteCurrentProject();
    });

    // Nút Khởi tạo dự án & Dẫn file trong khung Empty State
    document.getElementById('btn-empty-create-proj')?.addEventListener('click', () => {
      this.openCreateProjectModal();
    });
    document.getElementById('btn-empty-pick-file')?.addEventListener('click', () => {
      this.pickAndLoadLocalCPMSFile();
    });
  }

  async deleteCurrentProject() {
    if (this.isDeletingProject) return;
    this.isDeletingProject = true;

    try {
      const curProject = this.projects.find((p) => p.id === this.currentProjectId) || { id: this.currentProjectId, name: this.currentProjectId };
      const prjName = curProject.name || curProject.id;

      const confirmed = await this.confirmAsync({
        title: 'Cảnh báo xóa Dự án',
        message: `Bạn có chắc chắn muốn xóa vĩnh viễn dự án "${prjName}" (Mã: ${this.currentProjectId}) không?`,
        subMessage:
          'Hành động này sẽ xóa vĩnh viễn toàn bộ Hợp đồng & Phụ lục, Sổ thanh toán (IPC) và Hồ sơ liên quan. Không thể phục hồi!',
        confirmText: 'Xóa dự án',
        cancelText: 'Hủy bỏ',
        confirmType: 'danger',
      });

      if (!confirmed) return;

      const deletedPrjId = this.currentProjectId;
      if (typeof db.deleteProjectCascade === 'function') {
        await db.deleteProjectCascade(deletedPrjId);
      } else {
        // Fallback an toàn nếu trình duyệt chưa reload module db.js
        await db.delete('projects', deletedPrjId);
        const allContracts = await db.getAll('contracts');
        for (const c of allContracts) {
          if (c.project_id === deletedPrjId) await db.delete('contracts', c.id);
        }
        const allPayments = await db.getAll('payments');
        for (const p of allPayments) {
          if (p.project_id === deletedPrjId) await db.delete('payments', p.id);
        }
        const allDocs = await db.getAll('documents');
        for (const d of allDocs) {
          if (d.project_id === deletedPrjId) await db.delete('documents', d.id);
        }
      }

      if (typeof db.deleteProjectFileHandle === 'function') {
        await db.deleteProjectFileHandle(deletedPrjId);
      }
      this.activeFileHandle = null;
      this._lastKnownFileModified = 0;

      // Cập nhật lại danh sách dự án
      this.projects = await db.getAll('projects');
      if (this.projects.length > 0) {
        this.currentProjectId = this.projects[0].id;
      } else {
        // Nếu không còn dự án nào, đặt rỗng hoàn toàn
        this.currentProjectId = null;
      }

      await this.loadData();
      this.render();
      this.updateEmptyStateUI();
      if (typeof window.__MAIO_APP__?.updateProjectSourceBadge === 'function') {
        window.__MAIO_APP__.updateProjectSourceBadge();
      }
      if (window.app) {
        await window.app.updateDashboardStats();
      }

      // Cập nhật lại số liệu Dashboard nếu đang mở
      const statContractsCount = document.getElementById('stat-contracts-count');
      const statPaymentsCount = document.getElementById('stat-payments-count');
      const statDocsCount = document.getElementById('stat-docs-count');
      const statContractsVal = document.getElementById('stat-contracts-val');
      if (statContractsCount) statContractsCount.textContent = this.contracts.length;
      if (statPaymentsCount) statPaymentsCount.textContent = this.payments.filter((p) => !p.is_header).length;
      if (statDocsCount) statDocsCount.textContent = this.documents.length;
      if (statContractsVal) statContractsVal.textContent = this.formatCurrency(0);

      alert(`Đã xóa thành công toàn bộ dữ liệu của dự án "${prjName}"!`);
    } catch (err) {
      alert('Lỗi khi xóa dự án: ' + err.message);
    } finally {
      this.isDeletingProject = false;
    }
  }

  /**
   * Cập nhật lại đường dẫn local cho toàn bộ hồ sơ của dự án khi đồng bộ từ Drive về thư mục máy tính mới
   */
  async updateDocumentLocalPaths(newBaseDir, projectId = null) {
    const prjId = projectId || this.currentProjectId;
    if (!prjId || !newBaseDir) return 0;

    const allDocs = await db.getAll('documents');
    const prjDocs = allDocs.filter((d) => d.project_id === prjId);
    let count = 0;

    const cleanBase = newBaseDir.replace(/\\/g, '/').replace(/\/+$/, '');

    for (const d of prjDocs) {
      const origLink = d.local_link || d.file_link || '';
      let fileName = '';
      if (origLink && !origLink.startsWith('http')) {
        const parts = origLink.replace(/\\/g, '/').split('/');
        fileName = parts[parts.length - 1];
      }
      if (!fileName) {
        fileName = `${d.doc_no || d.id || 'HoSo'}.pdf`;
      }

      // Phân nhóm thư mục con chuẩn CPMS
      let subFolder = 'Doc';
      const cat = (d.category || '').toLowerCase();
      if (cat.includes('hợp đồng') || cat.includes('phụ lục') || cat.includes('contract')) {
        subFolder = 'HopDong_PhuLuc';
      } else if (cat.includes('thanh toán') || cat.includes('nghiệm thu') || cat.includes('payment')) {
        subFolder = 'ThanhToan';
      }

      const newLocalPath = `${cleanBase}/01_QuanLy_DuAn/DuAn_${prjId}/${subFolder}/${fileName}`;
      d.local_link = newLocalPath;
      d.file_link = newLocalPath; // Cập nhật lại link chính trong file CPMS
      await db.put('documents', d);
      count++;
    }

    // Tự động ghi đè cập nhật vào file local nếu đang mở file handle
    if (this.activeFileHandle) {
      try {
        await this.saveDirectlyToLocalFile();
      } catch (e) {
        console.warn('Lỗi ghi đè file local khi cập nhật đường dẫn hồ sơ:', e);
      }
    }

    await this.loadData();
    this.render();
    return count;
  }

  formatCurrency(amount) {
    if (!amount && amount !== 0) return '0 ₫';
    return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(amount);
  }

  formatDate(val) {
    if (!val) return '—';
    let d = null;
    if (val instanceof Date) {
      d = val;
    } else if (typeof val === 'number') {
      const utcD = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(utcD.getTime())) {
        d = new Date(utcD.getUTCFullYear(), utcD.getUTCMonth(), utcD.getUTCDate());
      }
    } else if (typeof val === 'string') {
      const trimmed = val.trim();
      // Nếu là định dạng DD/MM/YYYY hoặc D/M/YYYY
      const dmY = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmY) {
        return `${String(dmY[1]).padStart(2, '0')}/${String(dmY[2]).padStart(2, '0')}/${dmY[3]}`;
      }
      // Nếu là định dạng YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
        const parts = trimmed.slice(0, 10).split('-');
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
      }
      // Thử parse qua Date (như "Mon Aug 01 2022 00:00:00 GMT+0700 (Indochina Time)")
      const parsed = new Date(trimmed);
      if (!isNaN(parsed.getTime())) {
        d = parsed;
      }
    }

    if (d && !isNaN(d.getTime())) {
      const day = String(d.getDate()).padStart(2, '0');
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const year = d.getFullYear();
      return `${day}/${month}/${year}`;
    }

    return String(val).split(' ')[0] || String(val);
  }

  formatDateForInput(val) {
    if (!val) return '';
    let d = null;
    if (val instanceof Date) {
      d = val;
    } else if (typeof val === 'number') {
      const utcD = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(utcD.getTime())) {
        d = new Date(utcD.getUTCFullYear(), utcD.getUTCMonth(), utcD.getUTCDate());
      }
    } else if (typeof val === 'string') {
      const trimmed = val.trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
        return trimmed.slice(0, 10);
      }
      const dmY = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmY) {
        return `${dmY[3]}-${String(dmY[2]).padStart(2, '0')}-${String(dmY[1]).padStart(2, '0')}`;
      }
      const parsed = new Date(trimmed);
      if (!isNaN(parsed.getTime())) {
        d = parsed;
      }
    }

    if (d && !isNaN(d.getTime())) {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    return '';
  }

  formatDnNo(no) {
    if (no === undefined || no === null || String(no).trim() === '') return '—';
    const s = String(no).trim();
    if (/^IPC[-_\s]*/i.test(s)) {
      return s.replace(/^IPC[-_\s]*/i, 'ĐN-');
    }
    if (/^(ĐN|DN)[-_\s]*/i.test(s)) {
      return s;
    }
    return `ĐN-${s}`;
  }

  isContractCanceled(c) {
    const s = String(c?.status || '').trim().toLowerCase();
    return s === 'cancel' || s === 'cancelled' || s === 'canceled' || s === 'hủy' || s === 'terminated';
  }

  isContractCompleted(c) {
    const s = String(c?.status || '').trim().toLowerCase();
    return s === 'completed' || s === 'closed' || s === 'done' || s === 'hoàn thành' || s === 'đã quyết toán';
  }

  isContractOrAncestorCanceled(c) {
    if (!c) return false;
    if (this.isContractCanceled(c)) return true;
    let cur = c;
    let guard = 0;
    while (cur && cur.parent_id && guard < 10) {
      const parent = this.contracts.find((item) => item.id === cur.parent_id);
      if (!parent) break;
      if (this.isContractCanceled(parent)) return true;
      cur = parent;
      guard++;
    }
    return false;
  }

  getContractStatusBadgeHtml(status, compact = false) {
    const raw = String(status || 'Active').trim();
    const s = raw.toLowerCase();
    const style = compact ? ' style="font-size: 0.66rem; padding: 1px 5px; margin-left: 5px;"' : '';
    if (s === 'cancel' || s === 'cancelled' || s === 'canceled' || s === 'hủy' || s === 'terminated') {
      return `<span class="badge badge-status-cancel"${style}><i class="fa-solid fa-ban"></i> ${raw}</span>`;
    }
    if (s === 'completed' || s === 'closed' || s === 'done' || s === 'hoàn thành' || s === 'đã quyết toán') {
      return `<span class="badge badge-status-completed"${style}><i class="fa-solid fa-circle-check"></i> ${raw}</span>`;
    }
    if (s === 'suspended' || s === 'pending' || s === 'tạm ngưng' || s === 'tạm dừng') {
      return `<span class="badge badge-warning"${style}>${raw}</span>`;
    }
    return `<span class="badge badge-success"${style}>${raw}</span>`;
  }

  getDocumentStatusBadgeHtml(status) {
    const raw = String(status || 'Completed').trim();
    const s = raw.toLowerCase();
    if (s === 'completed' || s === 'hoàn thành' || s === 'closed' || s === 'done') {
      return `<span class="badge-doc-status badge-doc-completed" title="Hồ sơ đã hoàn thành"><i class="fa-solid fa-circle-check"></i> ${raw}</span>`;
    }
    if (s === 'accepted' || s === 'approved' || s === 'phê duyệt' || s === 'chấp thuận') {
      return `<span class="badge-doc-status badge-doc-approved" title="Hồ sơ đã phê duyệt"><i class="fa-solid fa-check-double"></i> ${raw}</span>`;
    }
    if (s.startsWith('cancel') || s === 'hủy' || s === 'đã hủy' || s === 'rejected' || s === 'từ chối') {
      return `<span class="badge-doc-status badge-doc-canceled" title="Hồ sơ bị hủy / Canceled"><i class="fa-solid fa-ban"></i> ${raw}</span>`;
    }
    if (s.startsWith('process') || s === 'đang xử lý' || s === 'in progress' || s === 'đang thực hiện') {
      return `<span class="badge-doc-status badge-doc-processing" title="Hồ sơ đang xử lý / Processing"><i class="fa-solid fa-rotate"></i> ${raw}</span>`;
    }
    if (s === 'returned' || s === 'trả lại' || s === 'yêu cầu sửa') {
      return `<span class="badge-doc-status badge-doc-returned" title="Hồ sơ bị trả lại"><i class="fa-solid fa-arrow-rotate-left"></i> ${raw}</span>`;
    }
    if (s === 'checking' || s === 'underreview' || s === 'đang kiểm tra' || s === 'kiểm tra' || s === 'rà soát') {
      return `<span class="badge-doc-status badge-doc-checking" title="Đang kiểm tra / Rà soát"><i class="fa-solid fa-magnifying-glass"></i> ${raw}</span>`;
    }
    if (s === 'pending' || s === 'chờ xử lý' || s === 'submitted' || s === 'đã nộp') {
      return `<span class="badge-doc-status badge-doc-pending" title="Chờ xử lý"><i class="fa-regular fa-clock"></i> ${raw}</span>`;
    }
    if (s === 'draft' || s === 'dự thảo') {
      return `<span class="badge-doc-status badge-doc-draft" title="Hồ sơ dự thảo"><i class="fa-regular fa-file-lines"></i> ${raw}</span>`;
    }
    return `<span class="badge-doc-status badge-doc-default"><i class="fa-solid fa-file"></i> ${raw}</span>`;
  }

  /**
   * Lấy badge phân loại hồ sơ có màu sắc riêng biệt cho từng loại (IPC, VO, CON, APP, TT, RFA...)
   */
  getDocumentCategoryBadgeHtml(category) {
    const raw = String(category || 'IPC').trim();
    const cat = raw.toUpperCase();

    // 1. IPC - Interim Payment Certificate
    if (cat.includes('IPC')) {
      return `<span class="badge-doc-type badge-type-ipc" title="Interim Payment Certificate - Hồ sơ thanh toán đợt">IPC</span>`;
    }
    // 2. VO - Variation Orders (Phát sinh / Thay đổi)
    if (cat.includes('VO') || cat.includes('PHÁT SINH') || cat.includes('VARIATION')) {
      return `<span class="badge-doc-type badge-type-vo" title="Variation Orders - Hồ sơ phát sinh / Bổ sung">VO</span>`;
    }
    // 3. CON / CONTRACT / HỢP ĐỒNG GỐC
    if (cat === 'CON' || cat.includes('CONTRACT') || cat.includes('HỢP ĐỒNG') || cat === 'HD') {
      return `<span class="badge-doc-type badge-type-con" title="Contract - Hợp đồng chính">${raw}</span>`;
    }
    // 4. APP / PHỤ LỤC / APPENDIX
    if (cat === 'APP' || cat.includes('PHỤ LỤC') || cat.includes('APPENDIX') || cat === 'PLHD' || cat === 'PL') {
      return `<span class="badge-doc-type badge-type-app" title="Appendix - Phụ lục hợp đồng">${raw}</span>`;
    }
    // 5. TT / THANH TOÁN KHỐI LƯỢNG / TẠM ỨNG
    if (cat === 'TT' || cat.includes('THANH TOÁN') || cat.includes('PAYMENT')) {
      return `<span class="badge-doc-type badge-type-tt" title="Thanh toán / Đề nghị thanh toán">${raw}</span>`;
    }
    // 6. RFA / RFI / ĐỆ TRÌNH DUYỆT
    if (cat.includes('RFA') || cat.includes('RFI') || cat.includes('TRÌNH') || cat.includes('SUBMIT')) {
      return `<span class="badge-doc-type badge-type-rfa" title="Request for Approval - Đệ trình phê duyệt">${raw}</span>`;
    }
    // 7. BB / BBNT / BIÊN BẢN NGHIỆM THU
    if (cat.includes('BB') || cat.includes('NGHIỆM THU') || cat.includes('ACCEPTANCE')) {
      return `<span class="badge-doc-type badge-type-bb" title="Biên bản nghiệm thu / Bàn giao">${raw}</span>`;
    }
    // 8. TK / BẢN VẼ / THIẾT KẾ
    if (cat.includes('TK') || cat.includes('DESIGN') || cat.includes('BẢN VẼ')) {
      return `<span class="badge-doc-type badge-type-tk" title="Hồ sơ thiết kế / Bản vẽ kỹ thuật">${raw}</span>`;
    }
    // 9. CV / CÔNG VĂN / THƯ TỪ
    if (cat === 'CV' || cat.includes('CÔNG VĂN') || cat.includes('LETTER') || cat.includes('OFFICIAL')) {
      return `<span class="badge-doc-type badge-type-cv" title="Công văn / Thư từ">${raw}</span>`;
    }
    // 10. NOTE / GHI CHÚ / MEMO
    if (cat === 'NOTE' || cat.includes('GHI CHÚ') || cat.includes('MEMO')) {
      return `<span class="badge-doc-type badge-type-note" title="Ghi chú / Thông báo nội bộ">${raw}</span>`;
    }
    // 11. CMT / Ý KIẾN / BÌNH LUẬN
    if (cat === 'CMT' || cat.includes('COMMENT') || cat.includes('Ý KIẾN')) {
      return `<span class="badge-doc-type badge-type-cmt" title="Ý kiến góp ý / Review comment">${raw}</span>`;
    }
    // 12. TRN / TRANSMITTAL / PHIẾU CHUYỂN
    if (cat === 'TRN' || cat.includes('TRANSMITTAL') || cat.includes('CHUYỂN')) {
      return `<span class="badge-doc-type badge-type-trn" title="Phiếu chuyển tài liệu / Transmittal">${raw}</span>`;
    }
    // 13. OTHER / Khác
    return `<span class="badge-doc-type badge-type-other" title="${raw}">${raw}</span>`;
  }

  getPaymentRate(p) {
    const r = parseFloat(p?.exchange_rate);
    return !isNaN(r) && r > 0 ? r : 1.0;
  }

  getPaymentVnd(p, field, fallbackField = null) {
    const rate = this.getPaymentRate(p);
    let raw = Number(p?.[field]);
    if (isNaN(raw)) raw = 0;
    if (raw === 0 && fallbackField) {
      const fb = Number(p?.[fallbackField]);
      if (!isNaN(fb)) raw = fb;
    }
    return Math.round(raw * rate);
  }

  formatMoneyPlain(val) {
    const num = Math.round(Number(val) || 0);
    return new Intl.NumberFormat('vi-VN').format(num);
  }

  parsePaymentDateObj(val) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) return val;
    if (typeof val === 'number' && val > 20000) {
      // Excel serial date -> chuyển sang Local Date đúng ngày
      const ms = Math.round((val - 25569) * 86400 * 1000);
      const utcD = new Date(ms);
      if (isNaN(utcD.getTime())) return null;
      return new Date(utcD.getUTCFullYear(), utcD.getUTCMonth(), utcD.getUTCDate());
    }
    const s = String(val).trim();
    if (!s) return null;
    // DD/MM/YYYY
    const mDmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (mDmy) {
      const d = new Date(Number(mDmy[3]), Number(mDmy[2]) - 1, Number(mDmy[1]));
      return isNaN(d.getTime()) ? null : d;
    }
    // YYYY-MM-DD
    const mYmd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (mYmd) {
      const d = new Date(Number(mYmd[1]), Number(mYmd[2]) - 1, Number(mYmd[3]));
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  matchesDnSearch(p, searchTerm) {
    if (!searchTerm) return true;
    const rawNo = String(p?.payment_no || '').trim().toLowerCase();
    const formattedNo = String(this.formatDnNo(p?.payment_no || '')).trim().toLowerCase();
    const normTerm = searchTerm.replace(/^đn[-\s]?|^dn[-\s]?/i, '').trim();
    if (!normTerm) return true;
    if (rawNo.includes(searchTerm) || formattedNo.includes(searchTerm)) return true;
    if (rawNo.includes(normTerm) || formattedNo.includes(normTerm)) return true;
    // Khớp theo giá trị số nếu cả 2 là số (vd gõ '5' khớp '05')
    if (/^\d+$/.test(normTerm) && /^\d+$/.test(rawNo)) {
      return Number(rawNo) === Number(normTerm);
    }
    return false;
  }

  formatNumberWithDots(val) {
    if (val === undefined || val === null || val === '') return '0';
    if (typeof val === 'number') {
      if (isNaN(val)) return '0';
      return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(val);
    }
    const str = String(val).trim();
    if (str === '-') return '-';
    const isNegative = str.startsWith('-');
    const unsigned = isNegative ? str.slice(1) : str;

    // Tách phần nguyên và phần thập phân (theo dấu phẩy ',' chuẩn vi-VN)
    const parts = unsigned.split(',');
    const intDigits = parts[0].replace(/\D/g, '');
    const decDigits = parts.length > 1 ? parts[1].replace(/\D/g, '').slice(0, 2) : null;

    if (!intDigits && !decDigits) return isNegative ? '-' : '0';

    let formattedInt = '0';
    if (intDigits) {
      try {
        formattedInt = new Intl.NumberFormat('vi-VN').format(BigInt(intDigits));
      } catch (e) {
        formattedInt = intDigits;
      }
    }
    const sign = isNegative && (formattedInt !== '0' || (decDigits && decDigits.length > 0)) ? '-' : '';
    return decDigits !== null ? `${sign}${formattedInt},${decDigits}` : `${sign}${formattedInt}`;
  }

  parseNumberFromDots(val) {
    if (!val && val !== 0) return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    const str = String(val).trim();
    if (!str || str === '-') return 0;
    const isNegative = str.startsWith('-');
    const unsigned = isNegative ? str.slice(1) : str;
    // Bỏ dấu chấm phân cách nghìn, đổi dấu phẩy thập phân sang dấu chấm JS
    const cleanStr = unsigned.replace(/\./g, '').replace(',', '.').replace(/[^\d.]/g, '');
    const num = parseFloat(cleanStr);
    if (isNaN(num)) return 0;
    return isNegative ? -num : num;
  }

  setupCurrencyInputListeners(container = document) {
    container.querySelectorAll('.currency-input').forEach((input) => {
      input.type = 'text';
      input.setAttribute('inputmode', 'decimal');

      if (input.value !== undefined && input.value !== '') {
        input.value = this.formatNumberWithDots(input.value);
      }

      input.addEventListener('input', () => {
        const cursorPosition = input.selectionStart || 0;
        const originalLength = input.value.length;
        const rawVal = input.value;
        const trimmed = rawVal.trim();

        if (!trimmed) {
          input.value = '';
          return;
        }
        if (trimmed === '-') {
          input.value = '-';
          return;
        }

        const isNegative = trimmed.startsWith('-');
        const unsigned = isNegative ? trimmed.slice(1) : trimmed;
        const hasComma = unsigned.includes(',');
        const parts = unsigned.split(',');
        const intDigits = parts[0].replace(/\D/g, '');
        const decDigits = hasComma ? parts[1].replace(/\D/g, '').slice(0, 2) : null;

        if (!intDigits && !hasComma) {
          input.value = isNegative ? '-' : '';
          return;
        }

        let formattedInt = '0';
        if (intDigits) {
          try {
            formattedInt = new Intl.NumberFormat('vi-VN').format(BigInt(intDigits));
          } catch (e) {
            formattedInt = intDigits;
          }
        }
        const sign = isNegative ? '-' : '';
        const formatted = hasComma ? `${sign}${formattedInt},${decDigits}` : `${sign}${formattedInt}`;
        input.value = formatted;

        const newLength = formatted.length;
        const newCursorPosition = Math.max(0, cursorPosition + (newLength - originalLength));
        try {
          input.setSelectionRange(newCursorPosition, newCursorPosition);
        } catch (err) {}
      });

      input.addEventListener('blur', () => {
        const v = input.value.trim();
        if (!v || v === '-' || v === '-,' || v === ',') {
          input.value = '0';
        } else {
          input.value = this.formatNumberWithDots(this.parseNumberFromDots(v));
        }
      });

      input.addEventListener('focus', () => {
        input.select();
      });
    });
  }

  getDrivePathForContract(contract) {
    const curPrjId = this.currentProjectId || 'CHC1';
    const isMain = !contract.parent_id;
    const parent = !isMain ? this.contracts.find((c) => c.id === contract.parent_id) : null;
    const cleanNo = (no) => String(no || '').replace(/[\/\\?%*:|"<>]/g, '_');

    const groupFolder = contract.group === 'A' ? '01_Nhom_A_ChuDauTu' : '02_Nhom_B_ThauPhu_NCC';

    if (contract.group === 'A') {
      if (isMain) {
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cleanNo(contract.contract_no)} / HopDong_PhuLuc / HopDong_Goc /`;
      } else {
        const parentNo = parent ? cleanNo(parent.contract_no) : cleanNo(contract.parent_id);
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${parentNo} / HopDong_PhuLuc / PhuLuc_${cleanNo(contract.contract_no)} /`;
      }
    } else {
      // Nhóm B: Phân theo 8 nhóm (Construction, Procurement, Consultancy, Design, Equipment, Service, Maintenance, Rental)
      const validCategories = [
        'Construction', 'Procurement', 'Consultancy', 'Design',
        'Equipment', 'Service', 'Maintenance', 'Rental'
      ];
      let cat = contract.category || 'Construction';
      const matched = validCategories.find((c) => c.toLowerCase() === cat.toLowerCase());
      cat = matched || 'Construction';

      if (isMain) {
        const partnerName = cleanNo(contract.company_b || 'NCC');
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${cleanNo(contract.contract_no)}_${partnerName} / HopDong_PhuLuc / HopDong_Goc /`;
      } else {
        const rootMain = parent && parent.parent_id ? this.contracts.find((c) => c.id === parent.parent_id) || parent : parent;
        const parentNo = rootMain ? cleanNo(rootMain.contract_no) : cleanNo(contract.parent_id);
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${parentNo} / HopDong_PhuLuc / PhuLuc_${cleanNo(contract.contract_no)} /`;
      }
    }
  }

  /**
   * Xác định thư mục con chuẩn trên Google Drive cho một hồ sơ:
   * - HopDong_PL: Hợp đồng, Phụ lục (CON, APP, PL...)
   * - ThanhToan: Hồ sơ thanh toán, tạm ứng (IPC, TT, ADV, Dong tien...)
   * - HoSo: Các hồ sơ còn lại (VO, RFA, BB, TK, CV, NOTE, CMT, TRN, OTHER...)
   */
  getDriveSubfolderKeyForDoc(doc) {
    if (!doc) return 'HoSo';
    const cat = String(doc.category || '').trim().toUpperCase();
    const title = String(doc.title || doc.name || '').toLowerCase();
    const docNo = String(doc.doc_no || '').toUpperCase();
    const filePath = String(doc.local_link || doc.file_link || '').toLowerCase();

    const isIpcByText =
      /(^|[.\-_/\s])ipc([.\-_/\s]|$)/i.test(docNo) ||
      /(^|[.\-_/\s])ipc([.\-_/\s]|$)/i.test(filePath) ||
      title.includes('interim payment') ||
      title.includes('thanh toán') ||
      title.includes('tạm ứng') ||
      filePath.includes('thanh toan') ||
      filePath.includes('dong tien');

    if (['IPC', 'TT', 'ADV', 'PAY'].includes(cat) || isIpcByText) {
      return 'ThanhToan';
    }

    if (
      ['CON', 'APP', 'PL', 'HD', 'HĐ'].includes(cat) ||
      title.includes('hợp đồng') ||
      title.includes('phụ lục') ||
      title.includes('contract') ||
      title.includes('appendix')
    ) {
      return 'HopDong_PL';
    }

    return 'HoSo';
  }

  /**
   * Trả về mảng các thư mục con (tương đối bên trong DuAn_<MaDA>) theo cây 3 trụ cột chuẩn của MAIO:
   * - Nhóm A: 01_Nhom_A_ChuDauTu / {SoHieuHD} / {HopDong_PhuLuc/... | ThanhToan/[PhuLuc_{SoPL}/]{ThuMucIPC} | HoSo}
   * - Nhóm B: 02_Nhom_B_ThauPhu_NCC / {8 nhom CPMS} / {SoHieuHD}_{TenNhaThau} / {HopDong_PhuLuc/... | ThanhToan/[PhuLuc_{SoPL}/]{ThuMucIPC} | HoSo}
   * - Hồ sơ chưa gắn HĐ: 00_Chua_Gan_HopDong
   */
  getDrivePathSegmentsForDoc(doc) {
    if (!doc) return ['00_Chua_Gan_HopDong'];
    const rawCid = String(doc.contract_id || '').trim();
    if (!rawCid || rawCid.toUpperCase() === 'UNASSIGNED') {
      return ['00_Chua_Gan_HopDong'];
    }
    const cidLower = rawCid.toLowerCase();
    const allContracts = this.contracts || [];
    const contract =
      allContracts.find((c) => String(c.id || '').trim().toLowerCase() === cidLower) ||
      allContracts.find((c) => String(c.contract_no || '').trim().toLowerCase() === cidLower) ||
      allContracts.find((c) => String(c.display_no || '').trim().toLowerCase() === cidLower);
    if (!contract) return ['00_Chua_Gan_HopDong'];

    const cleanName = (str, maxLen = 70) => {
      const cleaned = String(str || '')
        .replace(/[\/\\?%*:|"<>]/g, '_')
        .replace(/\s+/g, ' ')
        .trim();
      return (cleaned.length > maxLen ? cleaned.slice(0, maxLen).trim() : cleaned) || 'Khong_Ten';
    };

    // Tìm Hợp đồng gốc (Root Level 0)
    let rootMain = contract;
    if (contract.parent_id) {
      const parent = allContracts.find((c) => c.id === contract.parent_id);
      if (parent) {
        if (parent.parent_id) {
          rootMain = allContracts.find((c) => c.id === parent.parent_id) || parent;
        } else {
          rootMain = parent;
        }
      }
    }

    const group = contract.group || rootMain.group || 'B';
    let baseSegments = [];

    if (group === 'A') {
      const mainNo = cleanName(rootMain.contract_no || rootMain.id, 60);
      baseSegments = ['01_Nhom_A_ChuDauTu', mainNo];
    } else {
      const validCategories = [
        'Construction', 'Procurement', 'Consultancy', 'Design',
        'Equipment', 'Service', 'Maintenance', 'Rental'
      ];
      const rawCat = contract.category || rootMain.category || 'Construction';
      const matchedCat = validCategories.find((c) => c.toLowerCase() === String(rawCat).toLowerCase()) || 'Construction';
      const mainNo = cleanName(rootMain.contract_no || rootMain.id, 35);
      const partner = cleanName(rootMain.company_b || 'NCC', 45);
      baseSegments = ['02_Nhom_B_ThauPhu_NCC', matchedCat, `${mainNo}_${partner}`];
    }

    const subKey = this.getDriveSubfolderKeyForDoc(doc);
    if (subKey === 'HopDong_PL') {
      if (!contract.parent_id) {
        return [...baseSegments, 'HopDong_PhuLuc', 'HopDong_Goc'];
      }
      const plNo = cleanName(contract.contract_no || contract.id, 45);
      return [...baseSegments, 'HopDong_PhuLuc', `PhuLuc_${plNo}`];
    }

    if (subKey === 'ThanhToan') {
      let plFolder = '';
      if (contract.parent_id) {
        plFolder = `PhuLuc_${cleanName(contract.contract_no || contract.id, 45)}`;
      } else {
        // Nếu gắn ở HĐ gốc nhưng tên/mã hồ sơ có chỉ rõ phụ lục con (vd: IPC-APP10, IPC-PKG03...)
        const haystack = `${doc.doc_no || ''} ${doc.local_link || ''} ${doc.title || ''}`
          .toUpperCase()
          .replace(/[^A-Z0-9]/g, '');
        const subContracts = allContracts.filter(
          (c) => c.parent_id && (c.parent_id === rootMain.id || allContracts.some((p) => p.id === c.parent_id && p.parent_id === rootMain.id))
        );
        const matchedSub = subContracts.find((sub) => {
          const token = String(sub.contract_no || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
          return token.length >= 3 && haystack.includes(token);
        });
        if (matchedSub) {
          plFolder = `PhuLuc_${cleanName(matchedSub.contract_no || matchedSub.id, 45)}`;
        }
      }

      const rawIpc =
        String(doc.doc_no || '').trim() ||
        String(doc.title || doc.name || doc.id || 'IPC').trim();
      const cleanedIpc = cleanName(rawIpc, 55);
      const ipcFolder = /ipc/i.test(cleanedIpc) ? cleanedIpc : `IPC_${cleanedIpc}`;

      return plFolder
        ? [...baseSegments, 'ThanhToan', plFolder, ipcFolder]
        : [...baseSegments, 'ThanhToan', ipcFolder];
    }

    return [...baseSegments, 'HoSo'];
  }

  getDriveDocPathForContract(contract) {
    const curPrjId = this.currentProjectId || 'CHC1';
    const safePrjId = String(curPrjId).replace(/[\/\\?%*:|"<>]/g, '_');
    if (!contract) {
      return `MAIO_Data / 01_QuanLy_DuAn / DuAn_${safePrjId} /`;
    }
    const segs = this.getDrivePathSegmentsForDoc({ contract_id: contract.id, category: 'DOC' });
    return `MAIO_Data / 01_QuanLy_DuAn / DuAn_${safePrjId} / ${segs.join(' / ')} /`;
  }

  renderProjectInfo() {
    const infoContainer = document.getElementById('project-info-bar');
    if (!infoContainer) return;

    if (!this.projects || this.projects.length === 0) {
      const topProjName = document.getElementById('top-project-name');
      const topProjLoc = document.getElementById('top-project-location');
      if (topProjName) topProjName.textContent = 'Chưa có dự án nào';
      if (topProjLoc) topProjLoc.textContent = '—';

      infoContainer.innerHTML = '';
      infoContainer.style.display = 'none';
      return;
    }

    infoContainer.style.display = '';

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || this.projects[0];
    if (!curPrj) return;

    // Địa điểm đọc chính xác từ dữ liệu dự án trong file CPMS (thuộc tính location)
    const location = curPrj.location && String(curPrj.location).trim() ? String(curPrj.location).trim() : '—';
    const sampleContract = this.contracts[0] || {};
    const clientName = curPrj.client || sampleContract.company_a || 'Chủ đầu tư';

    // 1. Cập nhật thông tin cơ bản trên Top Banner (Vùng hiển thị cơ bản phía trên)
    const projectSelector = document.getElementById('project-selector');
    if (projectSelector && projectSelector.value !== curPrj.id) {
      projectSelector.value = curPrj.id;
    }
    const topProjName = document.getElementById('top-project-name');
    const topProjLoc = document.getElementById('top-project-location');
    if (topProjName) topProjName.textContent = curPrj.name || curPrj.id;
    if (topProjLoc) topProjLoc.textContent = location;

    if (typeof window.__MAIO_APP__?.updateProjectSourceBadge === 'function') {
      window.__MAIO_APP__.updateProjectSourceBadge();
    }

    // Cập nhật text trong dropdown dự án nếu cần
    const selOpt = document.querySelector(`#project-selector option[value="${curPrj.id}"]`);
    if (selOpt) {
      selOpt.textContent = this.formatProjectOptionLabel(curPrj);
    }

    // 2. Thống kê chi tiết Hợp đồng & Phụ lục nhóm A (loại trừ hợp đồng Cancel)
    const mainA = this.contracts.filter((c) => !c.parent_id && c.group === 'A' && !this.isContractCanceled(c));
    const appA = this.contracts.filter((c) => c.parent_id && c.group === 'A' && !this.isContractOrAncestorCanceled(c));
    let totalSignedA = 0;
    let totalPaidA = 0;
    mainA.forEach((c) => {
      const st = this.calculateTotals(c);
      totalSignedA += st.totalSigned;
      totalPaidA += st.totalPaid;
    });

    // 3. Thống kê chi tiết Hợp đồng & Phụ lục nhóm B (loại trừ hợp đồng Cancel)
    const mainB = this.contracts.filter((c) => !c.parent_id && c.group === 'B' && !this.isContractCanceled(c));
    const appB = this.contracts.filter((c) => c.parent_id && c.group === 'B' && !this.isContractOrAncestorCanceled(c));
    let totalSignedB = 0;
    let totalPaidB = 0;
    mainB.forEach((c) => {
      const st = this.calculateTotals(c);
      totalSignedB += st.totalSigned;
      totalPaidB += st.totalPaid;
    });

    const activeFileName = this.activeFileHandle?.name || curPrj.source_filename || null;

    // 4. Render thanh Project Info Bar chi tiết bên dưới
    infoContainer.innerHTML = `
      <div class="project-info-meta">
        <span class="proj-meta-item">
          <i class="fa-solid fa-hashtag" style="color: var(--primary);"></i> Mã DA: <strong>${curPrj.code || curPrj.id}</strong>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-file-contract" style="color: #2563eb;"></i> HĐ với A: <strong>${mainA.length} Hợp đồng${appA.length > 0 ? ` (${appA.length} PL)` : ''}</strong> — <span style="color: var(--primary); font-weight: 700;">${this.formatCurrency(totalSignedA)}</span>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-handshake" style="color: #6366f1;"></i> HĐ với B: <strong>${mainB.length} Hợp đồng${appB.length > 0 ? ` (${appB.length} PL)` : ''}</strong> — <span style="color: #4f46e5; font-weight: 700;">${this.formatCurrency(totalSignedB)}</span>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-circle-down" style="color: #10b981;"></i> Bên A giải ngân: <strong style="color: var(--success);">${this.formatCurrency(totalPaidA)}</strong>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-circle-up" style="color: #06b6d4;"></i> Giải ngân cho bên B: <strong style="color: #0891b2;">${this.formatCurrency(totalPaidB)}</strong>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item" title="${this.activeFileHandle ? `File máy tính: ${this.activeFileHandle.name} (Tự động lưu khi sửa đổi)` : (activeFileName ? `File data local: ${activeFileName}` : 'Chưa liên kết file data CPMS local')}">
          <i class="fa-solid fa-laptop-file" style="color: #6366f1;"></i> File Data Local: <strong style="color: ${this.activeFileHandle ? '#10b981' : (activeFileName ? 'var(--primary)' : 'var(--text-muted)')};">${activeFileName || 'Chưa liên kết'}</strong>
          ${this.activeFileHandle ? '<span class="badge badge-success" style="font-size: 0.68rem; margin-left: 4px; padding: 1px 5px;"><i class="fa-solid fa-bolt"></i> Tự động lưu</span>' : ''}
        </span>
      </div>
      <div style="flex-shrink: 0; margin-left: auto; display: flex; gap: 6px;">
        ${this.activeFileHandle ? `
          <button id="btn-manual-sync-file" class="button button-outline" style="font-size: 0.75rem; padding: 4px 10px; height: 28px; color: #0284c7; border-color: #bae6fd;" title="Ghi đè thủ công ngay vào file máy tính">
            <i class="fa-solid fa-floppy-disk"></i> Lưu file
          </button>
        ` : ''}
        <button id="btn-edit-project-info" class="button button-outline" style="font-size: 0.75rem; padding: 4px 10px; height: 28px;" title="Chỉnh sửa thông tin dự án">
          <i class="fa-solid fa-pen-to-square"></i> Sửa thông tin
        </button>
      </div>
    `;

    document.getElementById('btn-manual-sync-file')?.addEventListener('click', () => {
      this.syncToLocalFile(true);
    });

    document.getElementById('btn-edit-project-info')?.addEventListener('click', () => {
      this.openEditProjectModal(curPrj);
    });
  }

  openEditProjectModal(project) {
    let modal = document.getElementById('project-edit-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'project-edit-modal';
      modal.className = 'modal-overlay';
      document.body.appendChild(modal);
    }

    modal.innerHTML = `
      <div class="modal-container" style="max-width: 520px;">
        <div class="modal-header">
          <h2 style="font-size: 1.15rem; font-weight: 700;">
            <i class="fa-solid fa-pen-to-square" style="color: var(--primary);"></i> Chỉnh sửa thông tin dự án
          </h2>
          <button id="btn-close-edit-prj" class="button button-outline" style="padding: 4px 10px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body" style="padding: 16px 20px;">
          <form id="form-edit-project">
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Mã dự án (Project Code / ID) <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="project_id" value="${project.code || project.id}" required style="font-weight: 600; text-transform: uppercase;">
              <div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 3px;">Trùng khớp với trường <code>PROJ</code> / <code>ProjectID</code> trong sheet <code>PROJECT_SETTINGS</code> của file <code>.cpms</code> local.</div>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Tên dự án <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="project_name" value="${project.name || project.id}" required>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Chủ đầu tư (Bên A)</label>
              <input type="text" class="form-control" name="client" value="${project.client || ''}">
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Địa điểm xây dựng / Địa chỉ</label>
              <input type="text" class="form-control" name="location" value="${project.location || ''}">
            </div>
            <div class="form-group" style="margin-bottom: 16px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Tổng thầu / Đơn vị thực hiện (Bên B)</label>
              <input type="text" class="form-control" name="contractor" value="${project.contractor || ''}">
            </div>
            <div style="display: flex; justify-content: flex-end; gap: 8px;">
              <button type="button" id="btn-cancel-edit-prj" class="button button-outline">Hủy</button>
              <button type="submit" class="button button-primary"><i class="fa-solid fa-check"></i> Lưu thay đổi</button>
            </div>
          </form>
        </div>
      </div>
    `;

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-edit-prj').onclick = closeModal;
    modal.querySelector('#btn-cancel-edit-prj').onclick = closeModal;

    modal.querySelector('#form-edit-project').onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const oldId = project.id;
      const rawNewId = String(fd.get('project_id') || oldId).trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '_');
      const newId = rawNewId || oldId;

      const wasOnlineLinked = Boolean(
        project.drive_synced ||
        project.data_origin === 'online' ||
        (window.__MAIO_APP__?.onlineDriveProjects || []).some((o) => o.projectId === oldId)
      );

      const updatedPrj = {
        ...project,
        id: newId,
        code: newId,
        name: fd.get('project_name').trim(),
        client: fd.get('client').trim(),
        location: fd.get('location').trim(),
        contractor: fd.get('contractor').trim(),
        ...(wasOnlineLinked ? { drive_synced: true } : {}),
        updated_at: new Date().toISOString(),
      };

      if (newId !== oldId) {
        const allPrjs = await db.getAll('projects');
        if (allPrjs.some((p) => p.id === newId)) {
          await this.alertAsync({
            title: 'Trùng mã dự án',
            type: 'warning',
            message: `Mã dự án "${newId}" đã tồn tại. Vui lòng chọn mã khác!`,
          });
          return;
        }
        const [allC, allP, allD] = await Promise.all([
          db.getAll('contracts'),
          db.getAll('payments'),
          db.getAll('documents'),
        ]);
        for (const c of allC.filter((x) => x.project_id === oldId)) {
          c.project_id = newId;
          await db.put('contracts', c);
        }
        for (const p of allP.filter((x) => x.project_id === oldId)) {
          p.project_id = newId;
          await db.put('payments', p);
        }
        for (const d of allD.filter((x) => x.project_id === oldId)) {
          d.project_id = newId;
          await db.put('documents', d);
        }
        const oldFileHandle = this.activeFileHandle || (await db.getProjectFileHandle?.(oldId));
        if (oldFileHandle && typeof db.saveProjectFileHandle === 'function') {
          this.activeFileHandle = oldFileHandle;
          await db.saveProjectFileHandle(newId, oldFileHandle);
        }
        const oldDirHandle = await db.getProjectDirHandle?.(oldId);
        if (oldDirHandle && typeof db.saveProjectDirHandle === 'function') {
          await db.saveProjectDirHandle(newId, oldDirHandle);
        }
        await db.delete('projects', oldId);
        this.currentProjectId = newId;
      }

      await db.put('projects', updatedPrj);
      if (newId !== oldId && this._driveFolderCache?.has(oldId)) {
        this._driveFolderCache.set(newId, this._driveFolderCache.get(oldId));
        this._driveFolderCache.delete(oldId);
      }
      await this.loadData();
      closeModal();
      this.render();
      if (typeof window.__MAIO_APP__?.updateDashboardStats === 'function') {
        await window.__MAIO_APP__.updateDashboardStats();
      }
      await this.syncToLocalFile(true);
      this.showToast('Đã cập nhật thông tin dự án và đồng bộ!', 'success');
    };
  }

  openCloneProjectModal() {
    const curProject = this.projects.find((p) => p.id === this.currentProjectId) || {
      id: this.currentProjectId,
      name: this.currentProjectId,
    };

    let modal = document.getElementById('project-clone-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'project-clone-modal';
      modal.className = 'modal-overlay';
      document.body.appendChild(modal);
    }

    const defaultNewId = `${curProject.id}_COPY`;
    const defaultNewName = `${curProject.name || curProject.id} (Bản sao)`;

    const totalContracts = this.contracts.length;
    const totalPayments = this.payments.length;
    const totalDocs = this.documents.length;

    modal.innerHTML = `
      <div class="modal-container" style="max-width: 540px;">
        <div class="modal-header">
          <h2 style="font-size: 1.15rem; font-weight: 700;">
            <i class="fa-solid fa-copy" style="color: #6366f1;"></i> Sao chép (Clone) Dự án
          </h2>
          <button id="btn-close-clone-prj" class="button button-outline" style="padding: 4px 10px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body" style="padding: 16px 20px;">
          <div style="background: #f8fafc; border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 10px 14px; margin-bottom: 14px; font-size: 0.8rem;">
            <div><strong>Dự án nguồn:</strong> ${curProject.name || curProject.id} (Mã: <code>${curProject.id}</code>)</div>
            <div style="color: var(--text-muted); margin-top: 4px;">
              Quy mô dữ liệu: <strong>${totalContracts}</strong> Hợp đồng/Phụ lục | <strong>${totalPayments}</strong> Đợt thanh toán | <strong>${totalDocs}</strong> Hồ sơ
            </div>
          </div>

          <form id="form-clone-project">
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Mã dự án mới <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="new_project_id" value="${defaultNewId}" required placeholder="Ví dụ: PRJ_2026_01">
              <small style="color: var(--text-muted); font-size: 0.75rem;">Mã dự án viết liền, không dấu, không trùng lặp.</small>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Tên dự án mới <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="new_project_name" value="${defaultNewName}" required>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Chủ đầu tư</label>
              <input type="text" class="form-control" name="client" value="${curProject.client || ''}">
            </div>
            <div class="form-group" style="margin-bottom: 14px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Địa điểm</label>
              <input type="text" class="form-control" name="location" value="${curProject.location || ''}">
            </div>

            <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: var(--radius-sm); padding: 10px 12px; margin-bottom: 16px;">
              <div style="font-weight: 600; font-size: 0.8rem; margin-bottom: 6px; color: #1e40af;">Tùy chọn sao chép dữ liệu:</div>
              <label style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; margin-bottom: 4px; cursor: pointer;">
                <input type="checkbox" name="copy_contracts" checked disabled>
                <span>Sao chép toàn bộ danh mục Hợp đồng & Phụ lục con (${totalContracts} bản ghi)</span>
              </label>
              <label style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; margin-bottom: 4px; cursor: pointer;">
                <input type="checkbox" name="copy_payments" checked>
                <span>Sao chép toàn bộ Sổ theo dõi thanh toán IPC (${totalPayments} đợt)</span>
              </label>
              <label style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; cursor: pointer;">
                <input type="checkbox" name="copy_docs" checked>
                <span>Sao chép toàn bộ Hồ sơ & Công văn (${totalDocs} tài liệu)</span>
              </label>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 8px;">
              <button type="button" id="btn-cancel-clone-prj" class="button button-outline">Hủy</button>
              <button type="submit" class="button button-primary" style="background: #6366f1; border-color: #6366f1;">
                <i class="fa-solid fa-copy"></i> Tiến hành Sao chép
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-clone-prj').onclick = closeModal;
    modal.querySelector('#btn-cancel-clone-prj').onclick = closeModal;

    modal.querySelector('#form-clone-project').onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const newId = fd.get('new_project_id').trim();
      const newName = fd.get('new_project_name').trim();
      const client = fd.get('client').trim();
      const location = fd.get('location').trim();
      const copyPayments = fd.get('copy_payments') === 'on';
      const copyDocs = fd.get('copy_docs') === 'on';

      if (!newId || !newName) {
        alert('Vui lòng nhập đầy đủ Mã dự án và Tên dự án mới!');
        return;
      }

      if (this.projects.some((p) => p.id === newId)) {
        alert(`Mã dự án "${newId}" đã tồn tại. Vui lòng chọn một mã khác!`);
        return;
      }

      try {
        const newProject = {
          id: newId,
          code: newId,
          name: newName,
          client: client || curProject.client || '',
          location: location || curProject.location || '',
          contractor: curProject.contractor || '',
          created_at: new Date().toISOString(),
          cloned_from: curProject.id,
        };
        await db.put('projects', newProject);

        const contractIdMap = {};
        for (const c of this.contracts) {
          const newCId = `${newId}__${c.id}`;
          contractIdMap[c.id] = newCId;
        }

        for (const c of this.contracts) {
          const clonedC = {
            ...c,
            id: contractIdMap[c.id],
            project_id: newId,
            parent_id: c.parent_id ? (contractIdMap[c.parent_id] || c.parent_id) : null,
          };
          await db.put('contracts', clonedC);
        }

        if (copyPayments) {
          for (const p of this.payments) {
            const clonedP = {
              ...p,
              id: `${newId}__${p.id}`,
              project_id: newId,
              contract_id: contractIdMap[p.contract_id] || p.contract_id,
            };
            await db.put('payments', clonedP);
          }
        }

        if (copyDocs) {
          for (const d of this.documents) {
            const clonedD = {
              ...d,
              id: `${newId}__${d.id}`,
              project_id: newId,
              contract_id: contractIdMap[d.contract_id] || d.contract_id,
            };
            await db.put('documents', clonedD);
          }
        }

        closeModal();
        this.currentProjectId = newId;
        await this.loadData();
        this.render();

        alert(`🎉 Sao chép dự án thành công!\n\nĐã tạo dự án mới: "${newName}" (Mã: ${newId}).`);
      } catch (err) {
        alert('Lỗi sao chép dự án: ' + err.message);
      }
    };
  }

  openCloneContractModal(sourceContractId) {
    const sourceContract = this.contracts.find((c) => c.id === sourceContractId);
    if (!sourceContract) return;

    const isMain = !sourceContract.parent_id;
    const directChildren = this.getDirectChildren(sourceContract.id);
    const allDescendantIds = this.getAllDescendantIds(sourceContract.id).filter((id) => id !== sourceContract.id);

    let modal = document.getElementById('contract-clone-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'contract-clone-modal';
      modal.className = 'modal-overlay';
      document.body.appendChild(modal);
    }

    const nextContractId = isMain
      ? CPMSImporterExporter.generateNextContractID(null, this.contracts, sourceContract.group)
      : CPMSImporterExporter.generateNextContractID(
          this.contracts.find((c) => c.id === sourceContract.parent_id) || sourceContract,
          this.contracts,
          sourceContract.group
        );

    const defaultContractNo = `${sourceContract.contract_no || sourceContract.id} (Bản sao)`;
    const defaultScope = `${sourceContract.scope_of_work || ''} (Sao chép)`;

    modal.innerHTML = `
      <div class="modal-container" style="max-width: 540px;">
        <div class="modal-header">
          <h2 style="font-size: 1.15rem; font-weight: 700;">
            <i class="fa-solid fa-copy" style="color: #6366f1;"></i> Sao chép (Clone) ${isMain ? 'Hợp đồng' : 'Phụ lục'}
          </h2>
          <button id="btn-close-clone-ctr" class="button button-outline" style="padding: 4px 10px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body" style="padding: 16px 20px;">
          <div style="background: #f8fafc; border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 10px 14px; margin-bottom: 14px; font-size: 0.8rem;">
            <div><strong>Hợp đồng nguồn:</strong> ${sourceContract.contract_no || sourceContract.id} (Nhóm ${sourceContract.group})</div>
            <div style="color: var(--text-muted); margin-top: 2px;">Nội dung: ${sourceContract.scope_of_work || '—'}</div>
            <div style="color: var(--text-muted); margin-top: 2px;">Giá trị: <strong>${this.formatCurrency(sourceContract.contract_val)}</strong></div>
            ${directChildren.length > 0 ? `<div style="color: var(--primary); font-weight: 600; margin-top: 4px;">Có ${directChildren.length} phụ lục con trực tiếp (${allDescendantIds.length} mục con cháu)</div>` : ''}
          </div>

          <form id="form-clone-contract">
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Mã định danh hệ thống (ID) <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="new_contract_id" value="${nextContractId}" required>
              <small style="color: var(--text-muted); font-size: 0.75rem;">Mã phân cấp hệ thống chuẩn CPMS (vd: A-2026-02).</small>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Số hiệu Hợp đồng / Phụ lục mới <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="new_contract_no" value="${defaultContractNo}" required>
            </div>
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Nội dung / Phạm vi công việc <span style="color: var(--danger);">*</span></label>
              <input type="text" class="form-control" name="new_scope" value="${defaultScope}" required>
            </div>
            <div class="form-row" style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 12px;">
              <div class="form-group">
                <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Giá trị hợp đồng (VNĐ)</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" name="contract_val" value="${this.formatNumberWithDots(sourceContract.contract_val || 0)}">
              </div>
              <div class="form-group">
                <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Ngày ký</label>
                <input type="date" class="form-control" name="signing_date" value="${this.formatDateForInput(new Date())}">
              </div>
            </div>

            ${
              directChildren.length > 0
                ? `
              <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: var(--radius-sm); padding: 10px 12px; margin-bottom: 16px;">
                <label style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; cursor: pointer;">
                  <input type="checkbox" name="copy_appendix" checked>
                  <span><strong>Sao chép kèm theo ${directChildren.length} Phụ lục con</strong> (sẽ tự động tạo các mã con tương ứng dưới HĐ mới)</span>
                </label>
              </div>
            `
                : ''
            }

            <div style="display: flex; justify-content: flex-end; gap: 8px;">
              <button type="button" id="btn-cancel-clone-ctr" class="button button-outline">Hủy</button>
              <button type="submit" class="button button-primary" style="background: #6366f1; border-color: #6366f1;">
                <i class="fa-solid fa-copy"></i> Tiến hành Sao chép
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    this.setupCurrencyInputListeners(modal);

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-clone-ctr').onclick = closeModal;
    modal.querySelector('#btn-cancel-clone-ctr').onclick = closeModal;

    modal.querySelector('#form-clone-contract').onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const newId = fd.get('new_contract_id').trim();
      const newNo = fd.get('new_contract_no').trim();
      const newScope = fd.get('new_scope').trim();
      const val = this.parseNumberFromDots(fd.get('contract_val'));
      const sDate = fd.get('signing_date');
      const copyAppendix = fd.get('copy_appendix') === 'on';

      if (!newId || !newNo) {
        alert('Vui lòng nhập đầy đủ Mã định danh và Số hiệu HĐ!');
        return;
      }

      if (this.contracts.some((c) => c.id === newId)) {
        alert(`Mã hợp đồng "${newId}" đã tồn tại. Vui lòng chọn mã khác!`);
        return;
      }

      try {
        const newContract = {
          ...sourceContract,
          id: newId,
          contract_no: newNo,
          scope_of_work: newScope,
          contract_val: val,
          signing_date: sDate,
          project_id: this.currentProjectId,
          parent_id: isMain ? null : sourceContract.parent_id,
          created_at: new Date().toISOString(),
        };
        await db.put('contracts', newContract);

        if (copyAppendix && directChildren.length > 0) {
          const cloneSubTree = async (oldParentId, newParentObj) => {
            const subs = this.getDirectChildren(oldParentId);
            for (let i = 0; i < subs.length; i++) {
              const sub = subs[i];
              const subId = `${newParentObj.id}-${String(i + 1).padStart(2, '0')}`;
              const subContract = {
                ...sub,
                id: subId,
                parent_id: newParentObj.id,
                project_id: this.currentProjectId,
                created_at: new Date().toISOString(),
              };
              await db.put('contracts', subContract);
              await cloneSubTree(sub.id, subContract);
            }
          };

          await cloneSubTree(sourceContract.id, newContract);
        }

        closeModal();
        await this.loadData();
        this.render();
        await this.syncToLocalFile();

        alert(`🎉 Sao chép Hợp đồng thành công!\n\nĐã tạo "${newNo}" (Mã: ${newId}).`);
        this.openContractDetail(newContract.id);
      } catch (err) {
        alert('Lỗi sao chép Hợp đồng: ' + err.message);
      }
    };
  }

  removeVietnameseTones(str) {
    if (!str && str !== 0) return '';
    let s = String(str);
    s = s.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, 'a');
    s = s.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, 'e');
    s = s.replace(/ì|í|ị|ỉ|ĩ/g, 'i');
    s = s.replace(/ò|ó|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, 'o');
    s = s.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, 'u');
    s = s.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, 'y');
    s = s.replace(/đ/g, 'd');
    s = s.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, 'a');
    s = s.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, 'e');
    s = s.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, 'i');
    s = s.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, 'o');
    s = s.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, 'u');
    s = s.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, 'y');
    s = s.replace(/Đ/g, 'd');
    try {
      s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    } catch (e) {}
    return s.toLowerCase().trim();
  }

  isContractMatched(contract, query) {
    if (!query) return true;
    const qRaw = query.toLowerCase().trim();
    const qNoTone = this.removeVietnameseTones(query);

    const checkMatch = (val) => {
      if (!val && val !== 0) return false;
      const s = String(val).toLowerCase();
      if (s.includes(qRaw)) return true;
      const sNoTone = this.removeVietnameseTones(s);
      return sNoTone.includes(qNoTone);
    };

    // 1. Kiểm tra trực tiếp trên các thông tin của Hợp đồng chính
    if (
      checkMatch(contract.contract_no) ||
      checkMatch(contract.scope_of_work) ||
      checkMatch(contract.company_a) ||
      checkMatch(contract.company_b) ||
      checkMatch(contract.id) ||
      checkMatch(contract.rep_a) ||
      checkMatch(contract.rep_b) ||
      checkMatch(contract.category) ||
      checkMatch(contract.payee) ||
      checkMatch(contract.bank_account) ||
      checkMatch(contract.bank_name)
    ) {
      return true;
    }

    // Lấy toàn bộ ID con cháu (HĐ và tất cả phụ lục con)
    const allFamilyIds = this.getAllDescendantIds(contract.id);

    // 2. Kiểm tra trên tất cả Phụ lục con cháu
    const descendants = this.contracts.filter((c) => c.id !== contract.id && allFamilyIds.includes(c.id));
    for (const app of descendants) {
      if (
        checkMatch(app.contract_no) ||
        checkMatch(app.scope_of_work) ||
        checkMatch(app.company_b) ||
        checkMatch(app.id) ||
        checkMatch(app.category)
      ) {
        return true;
      }
    }

    // 3. Kiểm tra trong tất cả các đợt thanh toán / giải ngân (ĐẶC BIỆT LÀ GHI CHÚ)
    const payments = this.payments.filter((p) => allFamilyIds.includes(p.contract_id) || allFamilyIds.includes(p.parent_id));
    for (const p of payments) {
      if (
        checkMatch(p.note) ||            // Ghi chú đợt thanh toán
        checkMatch(p.item_name) ||       // Nội dung diễn giải thanh toán
        checkMatch(p.payment_no) ||      // Số ĐN
        checkMatch(p.invoice_ref) ||     // Số hóa đơn VAT
        checkMatch(p.payment_ref)        // Tham chiếu thanh toán
      ) {
        return true;
      }
    }

    // 4. Kiểm tra trong tất cả các hồ sơ / tài liệu (Document Tracker)
    const docs = this.documents.filter((d) => allFamilyIds.includes(d.contract_id));
    for (const d of docs) {
      if (
        checkMatch(d.notes) ||           // Ghi chú hồ sơ
        checkMatch(d.title) ||           // Tiêu đề tài liệu
        checkMatch(d.doc_no) ||          // Số công văn / hồ sơ
        checkMatch(d.id) ||
        checkMatch(d.category)
      ) {
        return true;
      }
    }

    return false;
  }

  updateCategoryFilterUI() {
    const select = document.getElementById('contract-category-filter');
    if (!select) return;

    // Lấy tất cả HĐ gốc của nhóm hiện tại
    const groupContracts = this.contracts.filter(
      (c) => c.group === this.currentGroup && !c.parent_id
    );

    // Danh sách 8 nhóm tiêu chuẩn
    const standardCategories = [
      { key: 'Construction', name: 'Thi công / Xây lắp' },
      { key: 'Procurement', name: 'Mua sắm vật tư' },
      { key: 'Consultancy', name: 'Tư vấn giám sát / Thẩm tra' },
      { key: 'Design', name: 'Thiết kế / Khảo sát' },
      { key: 'Equipment', name: 'Thiết bị MEP / Cơ điện' },
      { key: 'Service', name: 'Dịch vụ / Vệ sinh / Bảo vệ' },
      { key: 'Maintenance', name: 'Bảo trì / Vận hành' },
      { key: 'Rental', name: 'Thuê thiết bị / Nhân công' },
      { key: 'Design & Build', name: 'Tổng thầu EPC / D&B' },
      { key: 'CPB', name: 'Chi phí Ban / Quản lý' },
    ];

    const catCounts = {};
    let otherCount = 0;

    groupContracts.forEach((c) => {
      const cCat = (c.category || '').trim();
      if (!cCat) {
        otherCount++;
        return;
      }

      const matched = standardCategories.find(
        (sc) => sc.key.toLowerCase() === cCat.toLowerCase()
      );
      if (matched) {
        catCounts[matched.key] = (catCounts[matched.key] || 0) + 1;
      } else {
        catCounts[cCat] = (catCounts[cCat] || 0) + 1;
      }
    });

    let optionsHtml = `<option value="ALL">📁 Tất cả loại hợp đồng (${groupContracts.length})</option>`;

    // Render các standard categories có trong dữ liệu trước
    standardCategories.forEach((sc) => {
      const count = catCounts[sc.key] || 0;
      if (count > 0) {
        optionsHtml += `<option value="${sc.key}">🏷️ ${sc.key} - ${sc.name} (${count})</option>`;
      }
    });

    // Các category đặc thù khác (nếu có ngoài standard)
    Object.keys(catCounts).forEach((catKey) => {
      const isStd = standardCategories.some(
        (sc) => sc.key.toLowerCase() === catKey.toLowerCase()
      );
      if (!isStd) {
        optionsHtml += `<option value="${catKey}">🏷️ ${catKey} (${catCounts[catKey]})</option>`;
      }
    });

    // Các standard categories còn lại nếu có count = 0 (để người dùng biết đầy đủ các nhóm nghiệp vụ)
    standardCategories.forEach((sc) => {
      const count = catCounts[sc.key] || 0;
      if (count === 0 && this.currentGroup === 'B') {
        optionsHtml += `<option value="${sc.key}">🏷️ ${sc.key} - ${sc.name} (0)</option>`;
      }
    });

    if (otherCount > 0) {
      optionsHtml += `<option value="OTHER">📁 Khác / Chưa phân loại (${otherCount})</option>`;
    }

    select.innerHTML = optionsHtml;

    // Giữ lựa chọn hiện tại nếu có, ngược lại về ALL
    const exists = Array.from(select.options).some((opt) => opt.value === this.categoryFilter);
    if (exists) {
      select.value = this.categoryFilter;
    } else {
      this.categoryFilter = 'ALL';
      select.value = 'ALL';
    }
  }

  getEffectiveCategory(contract) {
    if (!contract) return '';
    if (contract.is_cpb || contract.contract_type === 'CPB' || (contract.id && contract.id.startsWith('B-CPB-'))) {
      return 'CPB';
    }
    let cat = (contract.category || '').trim();
    if (cat) return cat;
    // Kế thừa từ hợp đồng cha nếu là phụ lục con
    let cur = contract;
    let guard = 0;
    while (cur && cur.parent_id && guard < 10) {
      const parent = this.contracts.find((item) => item.id === cur.parent_id);
      if (!parent) break;
      if (parent.is_cpb || parent.contract_type === 'CPB') return 'CPB';
      if ((parent.category || '').trim()) return parent.category.trim();
      cur = parent;
      guard++;
    }
    return '';
  }

  updateDisburseCategoryFilterUI() {
    const select = document.getElementById('disburse-category-filter');
    if (!select) return;

    const groupContracts = this.contracts.filter((c) => c.group === this.disburseGroup);
    const standardCategories = [
      { key: 'Construction', name: 'Thi công / Xây lắp' },
      { key: 'Procurement', name: 'Mua sắm vật tư' },
      { key: 'Consultancy', name: 'Tư vấn giám sát / Thẩm tra' },
      { key: 'Design', name: 'Thiết kế / Khảo sát' },
      { key: 'Equipment', name: 'Thiết bị MEP / Cơ điện' },
      { key: 'Service', name: 'Dịch vụ / Vệ sinh / Bảo vệ' },
      { key: 'Maintenance', name: 'Bảo trì / Vận hành' },
      { key: 'Rental', name: 'Thuê thiết bị / Nhân công' },
      { key: 'Design & Build', name: 'Tổng thầu EPC / D&B' },
      { key: 'CPB', name: 'Chi phí Ban / Quản lý' },
    ];

    const catCounts = {};
    let otherCount = 0;

    groupContracts.forEach((c) => {
      const cCat = this.getEffectiveCategory(c);
      if (!cCat) {
        otherCount++;
        return;
      }
      const matched = standardCategories.find(
        (sc) => sc.key.toLowerCase() === cCat.toLowerCase()
      );
      if (matched) {
        catCounts[matched.key] = (catCounts[matched.key] || 0) + 1;
      } else {
        catCounts[cCat] = (catCounts[cCat] || 0) + 1;
      }
    });

    let optionsHtml = `<option value="ALL">📁 Tất cả loại hợp đồng (${groupContracts.length})</option>`;

    standardCategories.forEach((sc) => {
      const count = catCounts[sc.key] || 0;
      if (count > 0) {
        optionsHtml += `<option value="${sc.key}">🏷️ ${sc.key} - ${sc.name} (${count})</option>`;
      }
    });

    Object.keys(catCounts).forEach((catKey) => {
      const isStd = standardCategories.some(
        (sc) => sc.key.toLowerCase() === catKey.toLowerCase()
      );
      if (!isStd) {
        optionsHtml += `<option value="${catKey}">🏷️ ${catKey} (${catCounts[catKey]})</option>`;
      }
    });

    if (otherCount > 0) {
      optionsHtml += `<option value="OTHER">📁 Khác / Chưa phân loại (${otherCount})</option>`;
    }

    select.innerHTML = optionsHtml;

    const curFilter = this.disburseCategoryFilter || 'ALL';
    const exists = Array.from(select.options).some((opt) => opt.value === curFilter);
    if (exists) {
      select.value = curFilter;
    } else {
      this.disburseCategoryFilter = 'ALL';
      select.value = 'ALL';
    }
  }

  getMainContracts() {
    const list = this.contracts.filter((c) => {
      const matchGroup = c.group === this.currentGroup;
      const isMain = !c.parent_id;
      if (!matchGroup || !isMain) return false;

      const matchSearch = this.isContractMatched(c, this.searchQuery);
      if (!matchSearch) return false;

      if (!this.categoryFilter || this.categoryFilter === 'ALL') {
        return true;
      }

      if (this.categoryFilter === 'OTHER') {
        const standardCats = ['construction', 'procurement', 'consultancy', 'design', 'equipment', 'service', 'maintenance', 'rental', 'design & build'];
        const cat = (c.category || '').toLowerCase().trim();
        return !cat || !standardCats.includes(cat);
      }

      const target = this.categoryFilter.toLowerCase().trim();
      const cCat = (c.category || '').toLowerCase().trim();
      if (cCat === target) return true;

      // Kiểm tra xem có phụ lục / con cháu nào có category trùng không
      const familyIds = this.getAllDescendantIds(c.id);
      return this.contracts.some((ch) => familyIds.includes(ch.id) && (ch.category || '').toLowerCase().trim() === target);
    });

    // Sắp xếp danh sách hợp đồng theo ngày ký hợp đồng từ sớm nhất đến muộn nhất (CPB xếp cuối)
    list.sort((a, b) => {
      const isCpbA = Boolean(a.is_cpb || a.contract_type === 'CPB' || (a.id && String(a.id).startsWith('B-CPB-')));
      const isCpbB = Boolean(b.is_cpb || b.contract_type === 'CPB' || (b.id && String(b.id).startsWith('B-CPB-')));
      if (isCpbA !== isCpbB) return isCpbA ? 1 : -1;

      const dateA = this.parsePaymentDateObj(a.signing_date || a.created_at);
      const dateB = this.parsePaymentDateObj(b.signing_date || b.created_at);
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;

      if (timeA && timeB) {
        if (timeA !== timeB) return timeA - timeB;
      } else if (timeA && !timeB) {
        return -1;
      } else if (!timeA && timeB) {
        return 1;
      }

      const codeA = String(a.display_no || a.id || '').trim();
      const codeB = String(b.display_no || b.id || '').trim();
      return codeA.localeCompare(codeB, 'vi', { numeric: true, sensitivity: 'base' });
    });

    return list;
  }

  // Lấy các con trực tiếp (Level 1 dưới HĐ gốc, hoặc Level 2 dưới Phụ lục)
  getDirectChildren(contractId) {
    return this.contracts.filter((c) => c.parent_id === contractId);
  }

  // Lấy toàn bộ cây con cháu (để tra cứu thanh toán / hồ sơ)
  getAllDescendantIds(contractId) {
    const ids = [contractId];
    const queue = [contractId];
    while (queue.length > 0) {
      const curr = queue.shift();
      const children = this.contracts.filter((c) => c.parent_id === curr);
      for (const ch of children) {
        ids.push(ch.id);
        queue.push(ch.id);
      }
    }
    return ids;
  }

  // Lấy các thanh toán thực tế (loại bỏ hoàn toàn dòng header tổng hợp)
  getRealPayments(contractId) {
    const familyIds = this.getAllDescendantIds(contractId);
    return this.payments.filter((p) => {
      const matchFamily = familyIds.includes(p.contract_id);
      const isRealTransaction = !p.is_header; // Chỉ lấy giao dịch thực tế
      return matchFamily && isRealTransaction;
    });
  }

  // Lấy hồ sơ tài liệu
  getDocuments(contractId) {
    const familyIds = this.getAllDescendantIds(contractId);
    return this.documents.filter((d) => familyIds.includes(d.contract_id));
  }

  /**
   * TÍNH TOÁN TÀI CHÍNH CHUẨN XÁC:
   * - Đối với Giá trị HĐ:
   *   Nếu Hợp đồng gốc đã có giá trị tạm tính (vd: 1.200 tỷ ₫), giá trị hợp đồng là 1.200 tỷ ₫.
   *   TUYỆT ĐỐI KHÔNG CỘNG LẪN giá trị các phụ lục vào HĐ gốc vì phụ lục là các hạng mục chi tiết thuộc phạm vi HĐ này.
   *   (Tổng phụ lục được tính riêng để theo dõi, không cộng dồn làm đội giá trị HĐ).
   * - Đối với Giải ngân:
   *   Chỉ cộng các dòng thanh toán thực tế (is_header === false), không bao giờ cộng dòng Header.
   */
  calculateTotals(contract, depth = 0) {
    const directChildren = depth < 10 ? this.getDirectChildren(contract.id) : [];

    const exchangeRate = Number(contract.exchange_rate) || 1.0;
    const baseValRaw = contract.contract_val || 0;
    const baseVal = Math.round(baseValRaw * exchangeRate);

    // Tính tổng phụ lục con đệ quy (hỗ trợ cả phụ lục có phụ lục con tạm tính, loại trừ Cancel)
    let appendixSum = 0;
    let additiveAppendixSum = 0;

    directChildren.forEach((ch) => {
      if (this.isContractCanceled(ch)) return;
      const chStats = this.calculateTotals(ch, depth + 1);
      const chVal = chStats.totalSigned;
      appendixSum += chVal;

      const isAdditive = ch.is_additive === true || ch.is_additive === 1 || ch.is_additive === '1';
      if (isAdditive) {
        additiveAppendixSum += chVal;
      }
    });

    // Giá trị hợp đồng / phụ lục chuẩn xác (VNĐ):
    // 1. Nếu HĐ chủ hoặc Phụ lục được đánh dấu là tạm tính (is_provisional_val = true):
    //    KHÔNG cộng giá trị baseVal của chính nó, mà tổng hợp bằng tổng giá trị các phụ lục con trực tiếp (không bị Cancel).
    // 2. Nếu có giá trị gốc (baseVal > 0): Tổng = baseVal + các phụ lục con ĐƯỢC TÍCH CHỌN CỘNG DỒN.
    // 3. Nếu giá trị gốc = 0: Tổng = appendixSum (tổng các gói/phụ lục con).
    const isProvisional = contract.is_provisional_val === true || contract.is_provisional_val === 1 || contract.is_provisional_val === '1';
    let totalSigned = 0;
    if (isProvisional) {
      totalSigned = appendixSum;
    } else if (baseVal > 0) {
      totalSigned = baseVal + additiveAppendixSum;
    } else {
      totalSigned = appendixSum;
    }

    // Lấy thanh toán thực tế (nhân tỷ giá thời điểm thanh toán của từng đợt)
    const realPayments = this.getRealPayments(contract.id);
    const totalAccepted = realPayments.reduce((sum, p) => sum + this.getPaymentVnd(p, 'accepted_vol'), 0);
    const totalRequested = realPayments.reduce((sum, p) => sum + this.getPaymentVnd(p, 'request_amount'), 0);
    const totalPaid = realPayments.reduce((sum, p) => sum + this.getPaymentVnd(p, 'paid_amount'), 0);
    const remaining = totalRequested - totalPaid;
    const paymentRate = totalSigned > 0 ? (totalPaid / totalSigned) * 100 : 0;

    return {
      baseVal,
      baseValRaw,
      exchangeRate,
      appendixSum,
      additiveAppendixSum,
      isProvisional,
      totalSigned,
      totalAccepted,
      totalRequested,
      totalPaid,
      remaining,
      paymentRate: Math.min(100, Math.round(paymentRate * 10) / 10),
      directChildrenCount: directChildren.length,
      allDescendantsCount: this.getAllDescendantIds(contract.id).length - 1,
      ipcCount: realPayments.length,
    };
  }

  render() {
    this.updateEmptyStateUI('contracts');
    if (!this.isProjectReady()) return;

    this.renderProjectInfo();
    this.updateCategoryFilterUI();
    this.renderOverviewCompanyReport();

    // Ẩn / hiện nút Thêm Chi phí Ban (CPB) theo nhóm đang chọn (Chỉ hiện ở Nhóm B)
    const btnAddCpb = document.getElementById('btn-add-cpb');
    if (btnAddCpb) {
      if (this.currentGroup === 'B') {
        btnAddCpb.classList.remove('hidden');
      } else {
        btnAddCpb.classList.add('hidden');
      }
    }

    const container = document.getElementById('contracts-list-container');
    if (!container) return;

    // Nếu tài khoản mới chưa có dự án nào
    if (!this.projects || this.projects.length === 0) {
      container.innerHTML = `
        <div style="text-align: center; padding: 48px 24px; background: #fff; border-radius: var(--radius-md); border: 1px dashed var(--border-color); max-width: 650px; margin: 20px auto;">
          <div style="width: 64px; height: 64px; border-radius: 50%; background: #eff6ff; color: var(--primary); display: flex; align-items: center; justify-content: center; font-size: 1.8rem; margin: 0 auto 16px;">
            <i class="fa-solid fa-folder-plus"></i>
          </div>
          <h2 style="font-size: 1.25rem; font-weight: 700; color: #0f172a; margin-bottom: 8px;">Tài khoản của bạn chưa có dự án nào</h2>
          <p style="color: var(--text-muted); font-size: 0.9rem; margin-bottom: 24px; line-height: 1.5;">
            Hệ thống MAIO quản lý dữ liệu độc lập theo từng tài khoản Google cá nhân. Hãy bắt đầu bằng cách thêm dự án mới hoặc tải file dữ liệu dự án của bạn lên.
          </p>
          <div style="display: flex; justify-content: center; gap: 12px; flex-wrap: wrap;">
            <button id="btn-empty-add-proj" class="button button-primary">
              <i class="fa-solid fa-plus"></i> Thêm Dự án mới
            </button>
            <button id="btn-empty-import" class="button button-outline">
              <i class="fa-solid fa-laptop-file" style="color: var(--primary);"></i> Dẫn file CPMS Local
            </button>
          </div>
        </div>
      `;
      document.getElementById('btn-empty-add-proj')?.addEventListener('click', () => {
        this.openCreateProjectModal();
      });
      document.getElementById('btn-empty-import')?.addEventListener('click', () => {
        this.pickAndLoadLocalCPMSFile();
      });
      return;
    }

    const mainContracts = this.getMainContracts();

    if (mainContracts.length === 0) {
      const isFiltered = this.categoryFilter && this.categoryFilter !== 'ALL';
      container.innerHTML = `
        <div style="text-align: center; padding: 48px; background: #fff; border-radius: var(--radius-md); border: 1px dashed var(--border-color);">
          <i class="fa-solid fa-folder-open" style="font-size: 2.5rem; color: var(--text-light); margin-bottom: 12px;"></i>
          <p style="color: var(--text-muted); font-weight: 500;">
            ${isFiltered 
              ? `Không tìm thấy hợp đồng nào thuộc loại "<strong>${this.categoryFilter}</strong>" trong Nhóm ${this.currentGroup}.`
              : `Không tìm thấy hợp đồng nào phù hợp trong Nhóm ${this.currentGroup}.`}
          </p>
          ${isFiltered ? `
            <button id="btn-reset-category-filter" class="button button-outline" style="margin-top: 12px; font-size: 0.825rem;">
              <i class="fa-solid fa-rotate-left"></i> Xem tất cả loại hợp đồng
            </button>
          ` : ''}
        </div>
      `;
      document.getElementById('btn-reset-category-filter')?.addEventListener('click', () => {
        this.categoryFilter = 'ALL';
        this.render();
      });
      return;
    }

    container.innerHTML = mainContracts
      .map((c, index) => {
        const stats = this.calculateTotals(c);
        const directChildren = this.getDirectChildren(c.id);
        const realPayments = this.getRealPayments(c.id);
        const isExpanded = this.expandedContracts.has(c.id) || Boolean(this.searchQuery);
        const isCpb = c.is_cpb || c.contract_type === 'CPB' || (c.id && c.id.startsWith('B-CPB-'));

        // THẺ ĐẶC THÙ DÀNH CHO CHI PHÍ BAN (CPB - CHI PHÍ QUẢN LÝ DỰ ÁN KHÔNG CÓ HỢP ĐỒNG)
        if (isCpb) {
          return `
            <div class="contract-group-wrapper cpb-group-wrapper" data-contract-id="${c.id}" style="border: 1.5px solid #f59e0b; border-radius: var(--radius-md); background: #fffcf5; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(245, 158, 11, 0.08);">
              <!-- Card CPB -->
              <div class="contract-card-main cpb-card-main" data-open-detail="${c.id}" style="cursor: pointer; padding: 16px 20px;">
                <div class="contract-header" style="display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 12px;">
                  <div class="contract-header-left" style="flex: 1;">
                    <div class="contract-no" style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                      <span class="badge badge-contract-stt" style="background: #fef3c7; color: #b45309; border: 1px solid #fde68a; font-weight: 700; font-size: 0.8rem; padding: 2px 7px; border-radius: 4px;" title="Số thứ tự hợp đồng">#${index + 1}</span>
                      <i class="fa-solid fa-receipt" style="color: #d97706; font-size: 1.15rem;"></i>
                      <span class="contract-no-text" style="font-size: 1.05rem; font-weight: 800; color: #92400e;">${c.contract_no || c.id}</span>
                      <span class="badge" style="background: #fef3c7; color: #b45309; border: 1px solid #fde68a; font-weight: 700;">
                        <i class="fa-solid fa-calculator"></i> Chi phí Quản lý (CPB)
                      </span>
                      <span class="badge badge-success">${c.status || 'Active'}</span>
                    </div>
                    <div class="contract-title" style="font-size: 1.05rem; font-weight: 700; color: #1e293b; margin-top: 5px;">
                      ${c.scope_of_work || 'Chi Phí Ban (Chi phí quản lý dự án)'}
                    </div>
                  </div>

                  <div class="contract-header-right" style="display: flex; align-items: center; gap: 8px;">
                    <button class="button button-primary button-sm btn-add-cpb-payment" data-cpb-id="${c.id}" style="background: #d97706; border-color: #b45309; font-weight: 600; font-size: 0.8rem; padding: 5px 12px; height: 32px;" title="Thêm đợt chi / phiếu chi cho khoản này">
                      <i class="fa-solid fa-plus"></i> Đợt chi
                    </button>
                    <button class="button button-outline btn-edit-cpb-action" data-cpb-id="${c.id}" style="padding: 5px 10px; height: 32px; color: #b45309; border-color: #fde68a;" title="Chỉnh sửa thông tin Chi phí Ban">
                      <i class="fa-solid fa-pen"></i>
                    </button>
                    <button class="button button-outline btn-delete-cpb-action" data-cpb-id="${c.id}" style="padding: 5px 10px; height: 32px; color: var(--danger); border-color: #fca5a5;" title="Xóa khoản Chi phí Ban này">
                      <i class="fa-solid fa-trash-can"></i>
                    </button>
                    <div class="contract-val-block" style="text-align: right; min-width: 150px; margin-left: 6px;">
                      <span class="contract-val-label" style="color: #b45309; font-weight: 600;">Tổng tiền thực chi</span>
                      <div class="contract-val-amount" style="color: #b45309; font-size: 1.25rem; font-weight: 800;">
                        ${this.formatCurrency(stats.totalPaid)}
                      </div>
                      <div class="contract-val-sub" style="font-size: 0.75rem; color: #78716c;">
                        ${c.contract_val > 0 ? `Dự toán: ${this.formatCurrency(c.contract_val)}` : '(Chi phí quản lý không có HĐ)'}
                      </div>
                    </div>
                  </div>
                </div>

                <div class="contract-meta" style="display: flex; gap: 16px; flex-wrap: wrap; font-size: 0.825rem; color: #475569; border-top: 1px dashed #fde68a; padding-top: 10px; margin-top: 8px;">
                  <span><i class="fa-solid fa-building-user" style="color: #d97706;"></i> Đơn vị phụ trách: <strong>${c.company_b || 'Ban Quản lý / Nội bộ'}</strong></span>
                  ${c.payee ? `<span><i class="fa-solid fa-user-check" style="color: #d97706;"></i> Đối tượng nhận: <strong>${c.payee}</strong></span>` : ''}
                  ${c.signing_date || c.created_at ? `<span><i class="fa-solid fa-calendar-day" style="color: #d97706;"></i> Ngày tạo: ${this.formatDate(c.signing_date || c.created_at)}</span>` : ''}
                  <span><i class="fa-solid fa-receipt" style="color: #d97706;"></i> Đã giải ngân: <strong>${realPayments.length} đợt / phiếu chi</strong></span>
                </div>
              </div>

              <!-- Accordion Xem danh sách các đợt chi của CPB -->
              <div class="accordion-toggle-bar cpb-accordion-toggle" data-toggle-appendix="${c.id}" style="background: #fef3c7; border-top: 1px solid #fde68a; color: #92400e; padding: 10px 16px; font-weight: 600; cursor: pointer; display: flex; justify-content: space-between; align-items: center;">
                <span>
                  <i class="fa-solid ${isExpanded ? 'fa-folder-open' : 'fa-folder'}"></i>
                  <span>${isExpanded ? 'Thu gọn danh sách đợt chi' : 'Xem danh sách'} <strong>${realPayments.length} đợt chi</strong> (Chi tiết phiếu chi, số tiền, ngày chi)</span>
                </span>
                <i class="fa-solid ${isExpanded ? 'fa-chevron-up' : 'fa-chevron-down'}"></i>
              </div>

              ${isExpanded ? `
                <div class="appendix-table-container" style="background: #fff; padding: 12px 16px; border-bottom-left-radius: var(--radius-md); border-bottom-right-radius: var(--radius-md);">
                  ${realPayments.length === 0 ? `
                    <div style="text-align: center; padding: 24px; color: #94a3b8; font-size: 0.875rem;">
                      <i class="fa-solid fa-receipt" style="font-size: 1.5rem; margin-bottom: 6px; display: block; color: #d97706;"></i>
                      Chưa có đợt chi nào. Bấm nút <strong>"+ Đợt chi"</strong> ở trên để thêm đợt chi cho khoản này.
                    </div>
                  ` : `
                    <table class="appendix-grid-table" style="width: 100%; border-collapse: collapse;">
                      <thead>
                        <tr style="background: #fffbeb; border-bottom: 1px solid #fde68a;">
                          <th style="width: 90px; padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: left;">Đợt / Phiếu</th>
                          <th style="width: 110px; padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: left;">Ngày chi</th>
                          <th style="padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: left;">Nội dung chi</th>
                          <th style="width: 150px; padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: right;">Số tiền thực chi</th>
                          <th style="padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: left;">Ghi chú / Chứng từ</th>
                          <th style="width: 100px; padding: 8px 12px; font-size: 0.8rem; color: #92400e; text-align: center;">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody>
                        ${realPayments.map((p, pIdx) => `
                          <tr style="border-bottom: 1px solid #f1f5f9;">
                            <td style="padding: 8px 12px; font-weight: 700; color: #b45309;">
                              ${this.formatDnNo(p.payment_no || String(pIdx + 1))}
                            </td>
                            <td style="padding: 8px 12px; font-size: 0.825rem; color: #475569;">
                              ${this.formatDate(p.payment_date)}
                            </td>
                            <td style="padding: 8px 12px; font-weight: 600; color: #1e293b;">
                              ${p.item_name || 'Chi phí quản lý'}
                            </td>
                            <td style="padding: 8px 12px; text-align: right; font-weight: 700; color: #d97706;">
                              ${this.formatCurrency(p.paid_amount || p.request_amount)}
                            </td>
                            <td style="padding: 8px 12px; font-size: 0.8rem; color: #64748b;">
                              ${p.note || (p.invoice_ref ? `Chứng từ: ${p.invoice_ref}` : '—')}
                            </td>
                            <td style="padding: 8px 12px; text-align: center;">
                              <div style="display: flex; gap: 4px; justify-content: center;">
                                <button class="button button-outline btn-edit-cpb-payment" data-pay-id="${p.id}" data-contract-id="${c.id}" style="padding: 2px 7px; height: 26px; font-size: 0.75rem; color: #d97706; border-color: #fde68a;" title="Sửa đợt chi">
                                  <i class="fa-solid fa-pen"></i>
                                </button>
                                <button class="button button-outline btn-delete-cpb-payment" data-pay-id="${p.id}" data-contract-id="${c.id}" style="padding: 2px 7px; height: 26px; font-size: 0.75rem; color: var(--danger); border-color: #fca5a5;" title="Xóa đợt chi">
                                  <i class="fa-solid fa-trash-can"></i>
                                </button>
                              </div>
                            </td>
                          </tr>
                        `).join('')}
                      </tbody>
                    </table>
                  `}
                </div>
              ` : ''}
            </div>
          `;
        }

        const cardStatusClass = this.isContractCanceled(c)
          ? ' card-status-cancel'
          : this.isContractCompleted(c)
          ? ' card-status-completed'
          : '';

        return `
          <div class="contract-group-wrapper${cardStatusClass}" data-contract-id="${c.id}">
            <!-- Card Hợp đồng -->
            <div class="contract-card-main" data-open-detail="${c.id}">
              <div class="contract-header">
                <div class="contract-header-left">
                  <div class="contract-no">
                    <span class="badge badge-contract-stt" style="background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; font-weight: 700; font-size: 0.8rem; padding: 2px 7px; border-radius: 4px;" title="Số thứ tự hợp đồng">#${index + 1}</span>
                    <i class="fa-solid fa-file-signature"></i>
                    <span class="contract-no-text">${c.contract_no || c.id}</span>
                    <span class="badge ${c.group === 'A' ? 'badge-primary' : 'badge-warning'}">Nhóm ${c.group}</span>
                    ${c.category ? `<span class="badge" style="background: #eef2ff; color: #4338ca; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${c.category}</span>` : ''}
                    ${this.getContractStatusBadgeHtml(c.status)}
                  </div>
                  <div class="contract-title">${c.scope_of_work || 'Hợp đồng thực hiện dự án'}</div>
                </div>

                <div class="contract-header-right">
                  <button class="button button-outline btn-clone-contract-action" data-clone-cid="${c.id}" title="Sao chép Hợp đồng này">
                    <i class="fa-solid fa-copy" style="color: #6366f1;"></i>
                  </button>
                  <button class="button button-outline btn-delete-contract-action" data-del-cid="${c.id}" title="Xóa Hợp đồng này">
                    <i class="fa-solid fa-trash-can"></i>
                  </button>
                  <div class="contract-val-block">
                    <span class="contract-val-label">Giá trị Hợp đồng</span>
                    <div class="contract-val-amount">
                      ${this.formatCurrency(stats.totalSigned)}
                    </div>
                    ${directChildren.length > 0 ? `<div class="contract-val-sub">(Đã giao ${directChildren.length} PL: ${this.formatCurrency(stats.appendixSum)})</div>` : ''}
                    ${stats.additiveAppendixSum > 0 ? `<div class="contract-val-sub" style="color: #059669; font-weight: 600;">(+${this.formatCurrency(stats.additiveAppendixSum)} PL phát sinh)</div>` : ''}
                    ${c.exchange_rate && c.exchange_rate !== 1 ? `<div class="contract-val-sub" style="color: #64748b;">${this.formatNumberWithDots(c.contract_val)} nguyên tệ (TG: ${c.exchange_rate})</div>` : ''}
                  </div>
                </div>
              </div>

              <div class="contract-meta">
                <span>
                  <i class="fa-solid fa-handshake"></i> Bên A: <strong>${c.company_a || 'Chủ đầu tư'}</strong>
                  ${c.code_a ? ` <span class="badge" style="font-size: 0.68rem; background: #e0f2fe; color: #0369a1; padding: 1px 5px; font-family: monospace;">${c.code_a}</span>` : ''}
                  ${c.tax_code_a ? ` <span class="badge" style="font-size: 0.72rem; background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; padding: 1px 5px;"><i class="fa-solid fa-barcode"></i> MST: ${c.tax_code_a}</span>` : ''}
                </span>
                <span>
                  <i class="fa-solid fa-building"></i> Bên B: <strong>${c.company_b || 'Đơn vị thực hiện'}</strong>
                  ${c.code_b || c.company_code ? ` <span class="badge" style="font-size: 0.68rem; background: #f1f5f9; color: #475569; padding: 1px 5px; font-family: monospace;">${c.code_b || c.company_code}</span>` : ''}
                  ${c.tax_code ? `<span class="badge" style="font-size: 0.72rem; background: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; margin-left: 4px; padding: 1px 5px;"><i class="fa-solid fa-barcode"></i> MST: ${c.tax_code}</span>` : ''}
                  ${c.company_b2 ? ` & <strong>${c.company_b2}</strong>${c.code_b2 ? ` <span class="badge" style="font-size: 0.68rem; background: #f1f5f9; color: #475569; padding: 1px 5px; font-family: monospace;">${c.code_b2}</span>` : ''}${c.tax_code_b2 ? ` <span class="badge" style="font-size: 0.72rem; background: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; margin-left: 4px; padding: 1px 5px;"><i class="fa-solid fa-barcode"></i> MST: ${c.tax_code_b2}</span>` : ''} <span class="badge" style="font-size: 0.68rem; background: #e0e7ff; color: #4338ca; padding: 1px 5px; vertical-align: middle;">Liên danh</span>` : ''}
                </span>
                <span><i class="fa-solid fa-calendar-check"></i> Ngày ký: <strong>${this.formatDate(c.signing_date)}</strong></span>
                ${
                  c.linked_contract_id
                    ? `<button type="button" class="btn-jump-linked-contract" data-jump-id="${c.linked_contract_id}" title="Nhấp để chuyển nhanh đến hợp đồng liên kết" style="border: 1px solid #c7d2fe; background: #eef2ff; color: #4338ca; border-radius: 4px; padding: 2px 8px; font-size: 0.75rem; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 5px; margin-left: 4px;">
                        <i class="fa-solid fa-link"></i> HĐ liên kết: <strong>${(this.contracts.find((tc) => tc.id === c.linked_contract_id)?.contract_no) || c.linked_contract_id}</strong> <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.68rem;"></i>
                      </button>`
                    : ''
                }
              </div>

              <div class="contract-stats">
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Phụ lục cấp 1</span>
                  <span class="contract-stat-val" style="color: var(--primary);">${stats.directChildrenCount} phụ lục</span>
                </div>
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Đợt giải ngân (ĐN)</span>
                  <span class="contract-stat-val">${stats.ipcCount} đợt</span>
                </div>
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Đã giải ngân (${stats.paymentRate}%)</span>
                  <span class="contract-stat-val" style="color: var(--success);">${this.formatCurrency(stats.totalPaid)}</span>
                </div>
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Chênh lệch / Công nợ</span>
                  <span class="contract-stat-val" style="color: ${stats.remaining > 0 ? 'var(--danger)' : 'var(--text-main)'};">
                    ${this.formatCurrency(stats.remaining)}
                  </span>
                </div>
              </div>
            </div>

            <!-- Thanh Accordion Cây Phụ lục -->
            ${
              directChildren.length > 0
                ? `
              <div class="accordion-toggle-bar" data-toggle-appendix="${c.id}">
                <span>
                  <i class="fa-solid ${isExpanded ? 'fa-folder-open' : 'fa-folder'}"></i>
                  <span>${isExpanded ? 'Thu gọn' : 'Xem danh sách'} <strong>${directChildren.length} phụ lục con</strong></span>
                </span>
                <i class="fa-solid ${isExpanded ? 'fa-chevron-up' : 'fa-chevron-down'}"></i>
              </div>

              ${
                isExpanded
                  ? `
                <div class="appendix-table-container">
                  <table class="appendix-grid-table">
                    <thead>
                      <tr>
                        <th class="col-app-no">Số hiệu PL</th>
                        <th class="col-app-scope">Nội dung / Phạm vi công việc</th>
                        <th class="col-app-status">Tình trạng</th>
                        <th class="col-app-val">Giá trị ký kết</th>
                        <th class="col-app-paid">Đã giải ngân</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${this.renderAppendixRowsHierarchical(c.id)}
                    </tbody>
                  </table>
                </div>
              `
                  : ''
              }
            `
                : ''
            }
          </div>
        `;
      })
      .join('');

    // Bắt sự kiện Toggle Cây phụ lục / Danh sách đợt chi CPB
    container.querySelectorAll('[data-toggle-appendix]').forEach((toggleBar) => {
      toggleBar.addEventListener('click', (e) => {
        e.stopPropagation();
        const cid = toggleBar.getAttribute('data-toggle-appendix');
        if (this.expandedContracts.has(cid)) {
          this.expandedContracts.delete(cid);
        } else {
          this.expandedContracts.add(cid);
        }
        this.render();
      });
    });

    // Bắt sự kiện Click vào nút Sao chép HĐ
    container.querySelectorAll('.btn-clone-contract-action').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const cid = btn.getAttribute('data-clone-cid');
        this.openCloneContractModal(cid);
      });
    });

    // Bắt sự kiện Click vào nút Xóa Hợp đồng
    container.querySelectorAll('.btn-delete-contract-action').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const cid = btn.getAttribute('data-del-cid');
        if (cid) await this.deleteContract(cid);
      });
    });

    // Bắt sự kiện Thêm Đợt chi cho CPB
    container.querySelectorAll('.btn-add-cpb-payment').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const cid = btn.getAttribute('data-cpb-id');
        const cpbContract = this.contracts.find((c) => c.id === cid);
        if (cpbContract) this.openPaymentEntryModal(cpbContract);
      });
    });

    // Bắt sự kiện Sửa thông tin Chi phí Ban (CPB)
    container.querySelectorAll('.btn-edit-cpb-action').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const cid = btn.getAttribute('data-cpb-id');
        const cpbContract = this.contracts.find((c) => c.id === cid);
        if (cpbContract) this.openCreateCPBModal(cpbContract);
      });
    });

    // Bắt sự kiện Xóa Chi phí Ban (CPB)
    container.querySelectorAll('.btn-delete-cpb-action').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const cid = btn.getAttribute('data-cpb-id');
        await this.deleteCPB(cid);
      });
    });

    // Bắt sự kiện Sửa Đợt chi của CPB
    container.querySelectorAll('.btn-edit-cpb-payment').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const payId = btn.getAttribute('data-pay-id');
        const cid = btn.getAttribute('data-contract-id');
        const payment = this.payments.find((p) => p.id === payId);
        const cpbContract = this.contracts.find((c) => c.id === cid);
        if (payment && cpbContract) this.openEditPaymentModal(payment, cpbContract);
      });
    });

    // Bắt sự kiện Xóa Đợt chi của CPB
    container.querySelectorAll('.btn-delete-cpb-payment').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const payId = btn.getAttribute('data-pay-id');
        const payment = this.payments.find((p) => p.id === payId);
        if (!payment) return;

        this.showConfirmDialog({
          title: 'Xóa đợt chi?',
          message: `Bạn có chắc chắn muốn xóa đợt chi <strong>"${payment.item_name || payment.payment_no}"</strong> (${this.formatCurrency(payment.paid_amount || payment.request_amount)}) không?`,
          confirmText: 'Xóa vĩnh viễn',
          confirmType: 'danger',
          onConfirm: async () => {
            await db.delete('payments', payId);
            await this.loadData();
            this.render();
            await this.syncToLocalFile();
            this.showToast('Đã xóa đợt chi thành công!', 'success');
          },
        });
      });
    });

    // Bắt sự kiện Click vào HĐ gốc hoặc Phụ lục
    container.querySelectorAll('[data-open-detail]').forEach((el) => {
      el.addEventListener('click', (e) => {
        if (
          e.target.closest('[data-toggle-appendix]') ||
          e.target.closest('.btn-clone-contract-action') ||
          e.target.closest('.btn-delete-contract-action') ||
          e.target.closest('.btn-jump-linked-contract') ||
          e.target.closest('.btn-add-cpb-payment') ||
          e.target.closest('.btn-edit-cpb-action') ||
          e.target.closest('.btn-delete-cpb-action') ||
          e.target.closest('.btn-edit-cpb-payment') ||
          e.target.closest('.btn-delete-cpb-payment')
        ) return;
        const targetId = el.getAttribute('data-open-detail');
        this.openContractDetail(targetId);
      });
    });
  }

  /**
   * Render các dòng phụ lục theo cấp bậc phân tầng rõ ràng:
   * Level 1: APP-05, APP-06...
   * Level 2 (Con của APP-06): PKG-02, PKG-03 thụt lề có icon rõ ràng
   */
  renderAppendixRowsHierarchical(parentId, indentLevel = 0) {
    const children = this.getDirectChildren(parentId);
    if (children.length === 0) return '';

    // Sắp xếp phụ lục con theo ngày ký từ sớm đến muộn
    children.sort((a, b) => {
      const dateA = this.parsePaymentDateObj(a.signing_date || a.created_at);
      const dateB = this.parsePaymentDateObj(b.signing_date || b.created_at);
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;

      if (timeA && timeB) {
        if (timeA !== timeB) return timeA - timeB;
      } else if (timeA && !timeB) {
        return -1;
      } else if (!timeA && timeB) {
        return 1;
      }

      const codeA = String(a.display_no || a.id || '').trim();
      const codeB = String(b.display_no || b.id || '').trim();
      return codeA.localeCompare(codeB, 'vi', { numeric: true, sensitivity: 'base' });
    });

    return children
      .map((child) => {
        const childStats = this.calculateTotals(child);
        const subChildren = this.getDirectChildren(child.id);
        const hasSub = subChildren.length > 0;
        const indentPx = indentLevel * 20;

        let rowHtml = `
          <tr class="appendix-row-clickable" data-open-detail="${child.id}" title="Bấm để vào quản lý riêng ${child.contract_no}">
            <td class="col-app-no" style="padding-left: ${14 + indentPx}px;">
              ${indentLevel > 0 ? '<span style="color: var(--primary); margin-right: 4px;">↳</span>' : '<span style="color: var(--text-muted); margin-right: 4px;">↳</span>'}
              <strong>${child.contract_no}</strong>
              ${hasSub ? '<small style="color: var(--primary); display:block; font-size:0.7rem;">(Gói cha tổng)</small>' : ''}
            </td>
            <td class="col-app-scope">
              ${hasSub ? '<strong>' + child.scope_of_work + '</strong>' : child.scope_of_work || 'Phụ lục hợp đồng'}
            </td>
            <td class="col-app-status">
              ${this.getContractStatusBadgeHtml(child.status)}
            </td>
            <td class="col-app-val">
              ${this.formatCurrency(this.calculateTotals(child).totalSigned)}
              ${child.is_additive ? '<span class="badge badge-success" style="font-size: 0.65rem; margin-left: 4px;" title="Cộng dồn vào HĐ cha">+ Cộng dồn</span>' : '<span class="badge" style="font-size: 0.65rem; margin-left: 4px; background: #f1f5f9; color: #64748b;" title="Không cộng dồn vào HĐ cha">Không cộng dồn</span>'}
            </td>
            <td class="col-app-paid">${this.formatCurrency(childStats.totalPaid)}</td>
          </tr>
        `;

        // Nếu mục này có các gói con (như APP-06 có PKG-02, PKG-03) thì đệ quy render ngay bên dưới
        if (hasSub) {
          rowHtml += this.renderAppendixRowsHierarchical(child.id, indentLevel + 1);
        }

        return rowHtml;
      })
      .join('');
  }

  /**
   * Lấy nội dung diễn giải thanh toán thông minh từ dữ liệu CPMS Plus:
   * Ưu tiên cột 15 (PaymentRef) chuẩn CPMS, kế tiếp là cột 6 (ItemName),
   * tiếp đến là Ghi chú (Note) và số đợt nếu cả hai đều rỗng.
   */
  getPaymentContent(p) {
    if (!p) return '';
    const ref = (p.payment_ref !== undefined && p.payment_ref !== null) ? String(p.payment_ref).trim() : '';
    const name = (p.item_name !== undefined && p.item_name !== null) ? String(p.item_name).trim() : '';
    const note = (p.note !== undefined && p.note !== null) ? String(p.note).trim() : '';

    if (ref && !/^đợt\s*thanh\s*toán$/i.test(ref)) return ref;
    if (name && !/^đợt\s*thanh\s*toán$/i.test(name)) return name;
    if (ref) return ref;
    if (name) return name;
    if (note) return note;

    const noStr = String(p.payment_no || p.item_index || '').trim().replace(/^ĐN[-_\s]*/i, '').replace(/^DN[-_\s]*/i, '');
    return noStr ? `Đợt thanh toán số ${noStr}` : 'Đợt thanh toán';
  }

  // ==========================================
  // CHI TIẾT HỢP ĐỒNG HOẶC PHỤ LỤC (4 PHÂN KHU)
  // ==========================================
  openContractDetail(contractId) {
    const contract = this.contracts.find((c) => c.id === contractId);
    if (!contract) return;

    this.selectedContract = contract;
    const isMain = !contract.parent_id;
    const parentContract = !isMain ? this.contracts.find((c) => c.id === contract.parent_id) : null;

    const stats = this.calculateTotals(contract);
    const directChildren = this.getDirectChildren(contract.id);
    const realPayments = this.getRealPayments(contract.id);
    const docs = this.getDocuments(contract.id);

    // Chuẩn hóa ngày nhận: Các hợp đồng, phụ lục mặc định để ngày nhận là ngày ký
    docs.forEach((d) => {
      d.received_date = this.getDocReceivedDate(d, contract);
    });

    // Sắp xếp thứ tự hồ sơ theo ngày nhận (mặc định tăng dần từ ngày cũ nhất đến mới nhất theo tiến trình thực hiện)
    docs.sort((a, b) => {
      const timeA = this.parsePaymentDateObj(a.received_date)?.getTime() || 0;
      const timeB = this.parsePaymentDateObj(b.received_date)?.getTime() || 0;
      if (timeA && timeB) {
        if (timeA !== timeB) return timeA - timeB;
      } else if (timeA && !timeB) {
        return -1; // Hồ sơ có ngày nhận xếp trước
      } else if (!timeA && timeB) {
        return 1;
      }
      return String(a.id || '').localeCompare(String(b.id || ''));
    });

    // Sắp xếp các đợt giải ngân theo ngày thanh toán hoặc item_index
    realPayments.sort((a, b) => {
      const dateA = this.parsePaymentDateObj(a.payment_date)?.getTime() || 0;
      const dateB = this.parsePaymentDateObj(b.payment_date)?.getTime() || 0;
      return dateA - dateB;
    });

    // 1. Tính tổng các cột cho Bảng thu chi (Tab 3)
    let dnTotalActual = 0;
    let dnTotalAccepted = 0;
    let dnTotalRequested = 0;
    let dnTotalPaid = 0;
    let totalAdvance = 0;
    let advanceCount = 0;

    realPayments.forEach((p) => {
      dnTotalActual += this.getPaymentVnd(p, 'actual_volume');
      dnTotalAccepted += this.getPaymentVnd(p, 'accepted_vol');
      dnTotalRequested += this.getPaymentVnd(p, 'request_amount');
      dnTotalPaid += this.getPaymentVnd(p, 'paid_amount');

      const isAdv = (p.item_name && /tạm\s*ứng/i.test(p.item_name)) ||
                    (p.payment_ref && /tạm\s*ứng/i.test(p.payment_ref)) ||
                    (p.note && /tạm\s*ứng/i.test(p.note));
      if (isAdv) {
        totalAdvance += this.getPaymentVnd(p, 'paid_amount');
        advanceCount++;
      }
    });
    const dnTotalRemaining = dnTotalRequested - dnTotalPaid;

    // 2. Căn cứ vẽ Dashboard (Tab 4): Với những HĐ không có giá trị (totalSigned = 0), lấy Giá trị nghiệm thu để làm căn cứ
    const hasSignedVal = stats.totalSigned > 0;
    const baseTarget = hasSignedVal 
      ? stats.totalSigned 
      : (stats.totalAccepted > 0 ? stats.totalAccepted : (stats.totalPaid > 0 ? stats.totalPaid : 0));
    const isZeroContract = !hasSignedVal && stats.totalAccepted > 0;

    const paidVolume = Math.max(0, stats.totalPaid - totalAdvance);
    const paidRate = baseTarget > 0 ? Math.min(100, Math.round((stats.totalPaid / baseTarget) * 1000) / 10) : 0;
    const advanceRate = baseTarget > 0 ? Math.min(100, Math.round((totalAdvance / baseTarget) * 1000) / 10) : 0;
    const payVolumeRate = baseTarget > 0 ? Math.min(100, Math.round((paidVolume / baseTarget) * 1000) / 10) : 0;
    const acceptedRate = baseTarget > 0 ? Math.min(100, Math.round((stats.totalAccepted / baseTarget) * 1000) / 10) : 0;
    const remainingRate = baseTarget > 0 ? Math.round((Math.max(0, stats.remaining) / baseTarget) * 1000) / 10 : 0;

    // Danh sách phân loại hồ sơ duy nhất để lọc
    const docCategories = [...new Set(docs.map((d) => (d.category || 'Hồ sơ').trim()).filter(Boolean))].sort();
    const docCatCount = {};
    docs.forEach((d) => {
      const cat = (d.category || 'Hồ sơ').trim();
      docCatCount[cat] = (docCatCount[cat] || 0) + 1;
    });

    let modal = document.getElementById('contract-detail-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'contract-detail-modal';
      modal.className = 'modal-overlay';
      document.body.appendChild(modal);
    }

    modal.innerHTML = `
      <div class="modal-container contract-detail-dialog" style="height: 88vh; max-height: 88vh; display: flex; flex-direction: column;">
        <!-- Modal Header kèm Breadcrumb -->
        <div class="modal-header" style="flex-shrink: 0; display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: nowrap; gap: 16px; padding: 14px 20px;">
          <div style="flex: 1 1 auto; min-width: 0;">
            ${
              !isMain && parentContract
                ? `
              <div class="modal-breadcrumb" style="margin-bottom: 4px;">
                <span class="breadcrumb-link" id="btn-back-to-parent" data-parent-id="${parentContract.id}">
                  <i class="fa-solid fa-arrow-left"></i> ${parentContract.contract_no || parentContract.id} (${parentContract.parent_id ? 'Phụ lục cha' : 'Hợp đồng cha'})
                </span>
                <span>/</span>
                <span><strong>${contract.contract_no}</strong></span>
              </div>
            `
                : ''
            }
            <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
              <h2 style="font-size: 1.25rem; font-weight: 800; margin: 0; color: #0f172a;">
                ${isMain ? '' : '<span style="color: var(--primary);">[' + (directChildren.length > 0 ? 'Phụ lục cha' : 'Phụ lục / Gói con') + ']</span> '}
                ${contract.contract_no || contract.id}
              </h2>
              <span class="badge ${contract.group === 'A' ? 'badge-primary' : 'badge-warning'}">Nhóm ${contract.group}</span>
              ${contract.category ? `<span class="badge" style="background: #eef2ff; color: #4f46e5; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${contract.category}</span>` : ''}
              ${this.getContractStatusBadgeHtml(contract.status)}
            </div>
            ${contract.scope_of_work ? `<p style="font-size: 0.83rem; color: var(--text-muted); margin: 4px 0 0 0; line-height: 1.45; word-break: break-word;">${contract.scope_of_work}</p>` : ''}
          </div>
          <div style="flex: 0 0 auto; flex-shrink: 0; display: flex; align-items: center; gap: 8px; margin-top: 2px;">
            <button id="btn-clone-current-contract" class="button button-outline" style="font-size: 0.85rem; padding: 6px 10px;" title="Sao chép hợp đồng này">
              <i class="fa-solid fa-copy" style="color: #6366f1;"></i>
            </button>
            <button id="btn-edit-contract-info" class="button button-outline" style="font-size: 0.8rem; padding: 6px 12px; display: inline-flex; align-items: center; gap: 5px;">
              <i class="fa-solid fa-pen-to-square"></i> Sửa thông tin
            </button>
            <button id="btn-delete-current-contract" class="button button-outline btn-delete-contract-action" title="Xóa Hợp đồng này" style="padding: 6px 10px;">
              <i class="fa-solid fa-trash-can"></i>
            </button>
            <button id="btn-close-modal" class="button button-outline" style="padding: 6px 12px;" title="Đóng popup"><i class="fa-solid fa-xmark"></i></button>
          </div>
        </div>

        <!-- 4 TABS NGHIỆP VỤ & BỘ LỌC PHÂN LOẠI HỒ SƠ -->
        <div class="modal-nav-wrap" style="flex-shrink: 0; padding: 0 20px; background: #fff; border-bottom: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center; gap: 12px;">
          <div class="tab-nav" style="margin-bottom: 0;">
            <button class="tab-btn active modal-tab-btn" data-target="m-tab-info">
              <i class="fa-solid fa-circle-info"></i> 1. Thông tin ${isMain ? 'Hợp đồng' : 'Phụ lục'}
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-docs">
              <i class="fa-solid fa-folder-closed"></i> 2. Hồ sơ <span class="modal-tab-badge" id="modal-tab-docs-count">${docs.length}</span>
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-ipc">
              <i class="fa-solid fa-money-bill-wave"></i> 3. Giải ngân <span class="modal-tab-badge">${realPayments.length} đợt</span>
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-dash">
              <i class="fa-solid fa-chart-pie"></i> 4. Dashboard & Báo cáo
            </button>
          </div>
          <!-- Bộ lọc phân loại hồ sơ (Hiển thị ngay vị trí user yêu cầu khi vào tab Hồ sơ) -->
          <div id="docs-category-filter-wrap" style="display: none; align-items: center; gap: 6px; flex-shrink: 0;">
            <label for="select-doc-category-filter" style="font-size: 0.78rem; font-weight: 600; color: #64748b; margin: 0; white-space: nowrap; display: flex; align-items: center; gap: 4px;">
              <i class="fa-solid fa-filter" style="color: var(--primary);"></i> Phân loại:
            </label>
            <select id="select-doc-category-filter" class="form-control" style="font-size: 0.78rem; padding: 2px 8px; height: 30px; border-radius: 6px; min-width: 170px; background: #fff; cursor: pointer; border-color: #cbd5e1;">
              <option value="">Tất cả phân loại (${docs.length})</option>
              ${docCategories.map((cat) => `<option value="${cat}">${cat} (${docCatCount[cat] || 0})</option>`).join('')}
            </select>
          </div>
        </div>

        <div class="modal-body" style="flex: 1 1 0; min-height: 0; overflow-y: auto; padding: 20px;">
          <!-- PHÂN KHU 1: THÔNG TIN -->
          <div id="m-tab-info" class="modal-tab-pane">
            <div class="card" style="margin-bottom: 20px;">
              <h3 class="card-title"><i class="fa-solid fa-file-contract"></i> Thông tin ${isMain ? 'Hợp đồng' : 'Phụ lục'}</h3>
              <div class="form-grid-2">
                <div><strong>Mã nội bộ (Contract ID):</strong> <p style="color: var(--text-muted); font-family: monospace; font-weight: 600;">${contract.id}</p></div>
                <div><strong>Số hiệu chính thức:</strong> <p style="font-weight: 700; color: var(--primary);">${contract.contract_no}</p></div>
                <div><strong>Loại hợp đồng:</strong> <p><span class="badge" style="background: #eef2ff; color: #4f46e5; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${contract.category || 'Chưa phân loại'}</span></p></div>
                <div><strong>Tình trạng:</strong> <p>${this.getContractStatusBadgeHtml(contract.status)}</p></div>
                <div><strong>Ngày ký kết:</strong> <p style="color: var(--text-muted);">${this.formatDate(contract.signing_date)}</p></div>
                <div><strong>Thời hạn kết thúc:</strong> <p style="color: var(--text-muted);">${contract.end_date ? this.formatDate(contract.end_date) : '—'}</p></div>
                <div>
                  <strong>Giá trị ký kết Hợp đồng:</strong>
                  <p style="font-weight: 700; color: var(--primary);">
                    ${this.formatCurrency(stats.baseVal)}
                    ${contract.exchange_rate && contract.exchange_rate !== 1 ? `<small style="font-weight: normal; color: var(--text-muted); display: block;">(${this.formatNumberWithDots(contract.contract_val)} nguyên tệ × Tỷ giá ${contract.exchange_rate})</small>` : ''}
                    ${!isMain ? (contract.is_additive ? `<span class="badge badge-success" style="font-size: 0.72rem; margin-top: 3px; display: inline-block;"><i class="fa-solid fa-plus-circle"></i> Có cộng dồn vào HĐ cha</span>` : `<span class="badge" style="font-size: 0.72rem; margin-top: 3px; display: inline-block; background: #f1f5f9; color: #64748b;"><i class="fa-solid fa-circle-info"></i> Không cộng dồn vào HĐ cha</span>`) : ''}
                  </p>
                </div>
                ${
                  isMain && directChildren.length > 0
                    ? `<div><strong>Tổng giá trị các Phụ lục đã giao:</strong> <p style="font-weight: 700; color: #0284c7;">${this.formatCurrency(stats.appendixSum)}</p></div>`
                    : (!isMain ? `<div><strong>Thuộc HĐ / Phụ lục cha:</strong> <p style="color: var(--primary); font-weight: 600;">${parentContract?.contract_no || contract.parent_id}</p></div>` : '')
                }
                <div><strong>Bên A (Chủ đầu tư / Tổng thầu):</strong> <p style="color: var(--text-muted);">${contract.company_a || '—'} ${contract.code_a ? `[${contract.code_a}]` : ''} ${contract.tax_code_a ? `— MST: ${contract.tax_code_a}` : ''} ${contract.rep_a ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_a})</span>` : ''}</p></div>
                <div><strong>Bên B1 (Nhà thầu chính):</strong> <p style="color: var(--text-muted);">${contract.company_b || '—'} ${contract.code_b || contract.company_code ? `[${contract.code_b || contract.company_code}]` : ''} ${contract.tax_code ? `— MST: ${contract.tax_code}` : ''} ${contract.rep_b ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_b})</span>` : ''}</p></div>
                ${
                  contract.company_b2
                    ? `<div><strong>Bên B2 (Nhà thầu Liên danh):</strong> <p style="color: var(--text-muted);">${contract.company_b2} ${contract.code_b2 ? `[${contract.code_b2}]` : ''} ${contract.tax_code_b2 ? `— MST: ${contract.tax_code_b2}` : ''} ${contract.rep_b2 ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_b2})</span>` : ''}</p></div>`
                    : ''
                }
                ${
                  contract.contact_person_b || contract.contact_person_a
                    ? `<div><strong>Người liên hệ:</strong> <p style="color: var(--text-muted);">${[contract.contact_person_a ? `Bên A: ${contract.contact_person_a}` : '', contract.contact_person_b ? `Bên B: ${contract.contact_person_b}` : ''].filter(Boolean).join(' | ')}</p></div>`
                    : ''
                }
                ${
                  contract.linked_contract_id
                    ? `<div><strong>Hợp đồng liên kết tương tự:</strong> <p>
                        <button type="button" class="btn-jump-linked-contract" data-jump-id="${contract.linked_contract_id}" style="border: 1px solid #c7d2fe; background: #eef2ff; color: #4338ca; border-radius: 4px; padding: 3px 8px; font-size: 0.8rem; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 5px;">
                          <i class="fa-solid fa-link"></i> Đi đến HĐ: <strong>${(this.contracts.find((tc) => tc.id === contract.linked_contract_id)?.contract_no) || contract.linked_contract_id}</strong> <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.68rem;"></i>
                        </button>
                       </p></div>`
                    : ''
                }
                <div><strong>Đơn vị thụ hưởng:</strong> <p style="color: var(--text-muted);">${contract.payee || contract.company_b || '—'}</p></div>
                <div><strong>Tài khoản ngân hàng:</strong> <p style="color: var(--text-muted);">${contract.bank_account || '—'} ${contract.bank_name ? `(${contract.bank_name})` : ''}</p></div>
              </div>
            </div>

            <!-- Danh sách phụ lục / gói con trực tiếp -->
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px;">
              <h3 style="font-size: 0.95rem; font-weight: 700;">
                <i class="fa-solid fa-layer-group"></i> ${isMain ? 'Danh sách Phụ lục trực tiếp' : 'Các Gói con trực thuộc'} (${directChildren.length})
              </h3>
              <button id="btn-add-appendix" class="button button-primary button-sm">
                <i class="fa-solid fa-plus"></i> ${isMain ? 'Thêm Phụ lục mới' : 'Thêm Gói con mới'}
              </button>
            </div>

            ${
              directChildren.length > 0
                ? `
              <div class="data-table-wrap" style="border: 1px solid #e2e8f0; border-radius: 6px; overflow-x: auto; background: #ffffff;">
                <table class="data-table" style="width: 100%; margin: 0;">
                  <thead>
                    <tr>
                      <th style="width: 120px; text-align: center;">Số hiệu</th>
                      <th style="text-align: center;">Nội dung / Phạm vi công việc</th>
                      <th style="width: 110px; text-align: center;">Tình trạng</th>
                      <th style="width: 160px; text-align: center;">Giá trị ký kết</th>
                      <th style="width: 120px; text-align: center;">Ngày ký</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${directChildren
                      .map(
                        (a) => `
                        <tr class="clickable-appendix-row" data-sub-id="${a.id}" style="cursor: pointer;" title="Bấm để vào quản lý riêng ${a.contract_no}">
                          <td><strong style="color: var(--primary);">${a.contract_no}</strong></td>
                          <td>${a.scope_of_work || '—'}</td>
                          <td style="text-align: center;">${this.getContractStatusBadgeHtml(a.status)}</td>
                          <td class="num-col" style="color: var(--primary); font-weight: 700;">
                            ${this.formatCurrency(this.calculateTotals(a).totalSigned)}
                            ${a.is_additive ? '<span class="badge badge-success" style="font-size: 0.65rem; margin-left: 4px;">+ Cộng dồn</span>' : '<span class="badge" style="font-size: 0.65rem; margin-left: 4px; background: #f1f5f9; color: #64748b;">Không cộng dồn</span>'}
                          </td>
                          <td style="text-align: center;">${this.formatDate(a.signing_date)}</td>
                        </tr>
                      `
                      )
                      .join('')}
                  </tbody>
                </table>
              </div>
            `
                : `
              <div style="text-align: center; padding: 22px; background: #f8fafc; border: 1px dashed var(--border-color); border-radius: var(--radius-md); color: var(--text-muted); font-size: 0.85rem;">
                <i class="fa-solid fa-file-circle-plus" style="font-size: 1.5rem; color: #94a3b8; margin-bottom: 6px; display: block;"></i>
                Chưa có ${isMain ? 'phụ lục' : 'gói con'} nào cho hợp đồng này.
                <div style="margin-top: 8px;">
                  <button id="btn-add-appendix-empty" class="button button-outline" style="font-size: 0.8rem; padding: 4px 12px;">
                    <i class="fa-solid fa-plus"></i> ${isMain ? 'Thêm Phụ lục ngay' : 'Thêm Gói con ngay'}
                  </button>
                </div>
              </div>
            `
            }
          </div>

          <!-- PHÂN KHU 2: HỒ SƠ RIÊNG -->
          <div id="m-tab-docs" class="modal-tab-pane hidden">
            <!-- Toolbar gọn gàng, bỏ vùng cây thư mục lớn -->
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; gap: 10px; flex-wrap: wrap;">
              <h3 style="font-size: 0.95rem; font-weight: 700; margin: 0; display: flex; align-items: center; gap: 8px;">
                <i class="fa-solid fa-list-check" style="color: var(--primary);"></i> Danh mục Hồ sơ & Chứng từ (<span id="doc-list-heading-count">${docs.length}</span>)
              </h3>
              <div style="display: flex; gap: 8px; align-items: center;">
                <button id="btn-open-drive-folder" class="button button-outline" style="font-size: 0.8rem; padding: 5px 12px; color: #0284c7; border-color: #bae6fd; background: #f0f9ff; display: inline-flex; align-items: center; gap: 6px;" title="Mở thư mục Google Drive của hợp đồng / phụ lục này">
                  <i class="fa-brands fa-google-drive" style="color: #0284c7;"></i> Mở Thư mục Drive
                </button>
                <button id="btn-add-document" class="button button-primary" style="font-size: 0.8rem; padding: 5px 12px; display: inline-flex; align-items: center; gap: 6px;">
                  <i class="fa-solid fa-plus"></i> Thêm Hồ sơ mới
                </button>
              </div>
            </div>

            <div class="data-table-wrap modal-sticky-table-wrap" style="border: 1px solid #e2e8f0; border-radius: 6px; overflow-x: auto; overflow-y: auto; max-height: calc(88vh - 175px); position: relative; background: #ffffff;">
              <table class="data-table" style="width: 100%; margin: 0;">
                <thead>
                  <tr>
                    <th style="width: 110px; white-space: nowrap; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Mã hồ sơ</th>
                    <th style="width: 140px; white-space: nowrap; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Số công văn / Hồ sơ</th>
                    <th style="width: 70px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Phân loại</th>
                    <th style="text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Tiêu đề tài liệu</th>
                    <th id="th-sort-doc-date" style="width: 95px; text-align: center; white-space: nowrap; cursor: pointer; user-select: none; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);" title="Bấm để đảo chiều sắp xếp theo ngày nhận">Ngày nhận <i class="fa-solid fa-arrow-down-short-wide" style="font-size: 0.7rem; margin-left: 3px; color: var(--primary);"></i></th>
                    <th style="width: 95px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Tình trạng</th>
                    <th style="width: 110px; white-space: nowrap; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Tệp tin</th>
                    <th style="width: 105px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  ${
                    docs.length === 0
                      ? '<tr><td colspan="8" style="text-align:center; padding: 20px;">Chưa có hồ sơ nào. Bấm "+ Thêm Hồ sơ mới" để tạo.</td></tr>'
                      : docs
                          .map(
                            (d) => `
                        <tr class="doc-item-row" data-doc-id="${d.id}" data-doc-category="${(d.category || 'Hồ sơ').trim()}" data-received-time="${this.parsePaymentDateObj(d.received_date)?.getTime() || 0}" style="cursor: pointer;" title="Nhấp đúp để mở popup sửa hồ sơ">
                          <td style="white-space: nowrap; font-family: monospace; font-weight: 600; color: #334155; text-align: center;">${d.id}</td>
                          <td style="white-space: nowrap;">${d.doc_no || '—'}</td>
                          <td style="text-align: center; white-space: nowrap;">${this.getDocumentCategoryBadgeHtml(d.category)}</td>
                          <td style="max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${d.title || ''}">${d.title || '—'}</td>
                          <td style="text-align: center; white-space: nowrap;">${this.formatDate(d.received_date)}</td>
                          <td style="text-align: center; white-space: nowrap;">${this.getDocumentStatusBadgeHtml(d.status)}</td>
                          <td style="width: 110px; max-width: 120px; padding: 4px 6px;">
                            <div style="display: flex; flex-direction: column; gap: 3px;">
                              <!-- Link Local thu gọn -->
                              ${(d.local_link || (!String(d.file_link || '').startsWith('http') && d.file_link)) ? `
                                <div style="display: flex; align-items: center; gap: 4px; font-size: 0.72rem;" title="${d.local_link || d.file_link}">
                                  <i class="fa-solid fa-file" style="color: #64748b; font-size: 0.7rem; flex-shrink: 0;"></i>
                                  <span style="max-width: 75px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-family: monospace; font-size: 0.72rem; color: #1e293b;">
                                    ${this.getBaseName(d.local_link || d.file_link)}
                                  </span>
                                  <button class="button button-outline btn-copy-doc-local" data-path="${encodeURIComponent(d.local_link || d.file_link)}" style="padding: 0 4px; font-size: 0.65rem; border-color: #cbd5e1; height: 18px; line-height: 1; flex-shrink: 0;" title="Sao chép đường dẫn file">
                                    <i class="fa-regular fa-copy"></i>
                                  </button>
                                </div>
                              ` : ''}
                              <!-- Link Drive Online thu gọn -->
                              ${(d.drive_link || (String(d.file_link || '').startsWith('http') && d.file_link)) ? `
                                <div>
                                  <a href="${d.drive_link || d.file_link}" target="_blank" style="padding: 1px 6px; font-size: 0.68rem; color: #0284c7; background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 3px; display: inline-flex; align-items: center; gap: 3px; text-decoration: none; height: 18px; line-height: 1;">
                                    <i class="fa-brands fa-google-drive" style="font-size: 0.68rem;"></i> Drive
                                  </a>
                                </div>
                              ` : ''}
                              ${(!d.local_link && !d.drive_link && !d.file_link) ? '<span style="color: #94a3b8; font-size: 0.72rem;">—</span>' : ''}
                            </div>
                          </td>
                          <td style="white-space: nowrap; width: 105px; text-align: center; padding: 4px 6px;">
                            <div style="display: flex; gap: 3px; align-items: center; justify-content: center;">
                              <!-- Mở tệp tin -->
                              <button class="button button-outline btn-open-doc-file" data-doc-id="${d.id}" style="padding: 2px 6px; font-size: 0.72rem; color: #0284c7; border-color: #bae6fd; background: #f0f9ff;" title="Mở file tài liệu này">
                                <i class="fa-solid fa-file-arrow-up"></i>
                              </button>
                              <!-- Mở thư mục -->
                              <button class="button button-outline btn-open-doc-folder" data-doc-id="${d.id}" style="padding: 2px 6px; font-size: 0.72rem; color: #d97706; border-color: #fde68a; background: #fffbeb;" title="Mở thư mục chứa file">
                                <i class="fa-solid fa-folder-open"></i>
                              </button>
                              <!-- Chỉnh sửa thông tin hồ sơ & lưu trữ -->
                              <button class="button button-outline btn-edit-doc" data-doc-id="${d.id}" style="padding: 2px 6px; font-size: 0.72rem; color: #4f46e5; border-color: #c7d2fe; background: #eef2ff;" title="Chỉnh sửa thông tin lưu trữ & hồ sơ">
                                <i class="fa-solid fa-pen-to-square"></i>
                              </button>
                              <!-- Xóa hồ sơ -->
                              <button class="button button-outline btn-delete-doc" data-doc-id="${d.id}" style="padding: 2px 6px; font-size: 0.72rem; color: var(--danger); border-color: #fecdd3; background: #fff1f2;" title="Xóa hồ sơ">
                                <i class="fa-solid fa-trash"></i>
                              </button>
                            </div>
                          </td>
                        </tr>
                      `
                          )
                          .join('')
                  }
                </tbody>
              </table>
            </div>
          </div>

          <!-- PHÂN KHU 3: THANH TOÁN (PAYMENT LEDGER / IPC) ĐỒNG BỘ BẢNG THU CHI -->
          <div id="m-tab-ipc" class="modal-tab-pane hidden">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 10px;">
              <div style="font-size: 0.82rem; color: var(--text-muted); display: flex; gap: 14px; flex-wrap: wrap; align-items: center;">
                <span><i class="fa-solid fa-building"></i> ${contract.group === 'A' ? 'Bên A' : 'Bên B'}: <strong>${contract.group === 'A' ? (contract.company_a || 'Chủ đầu tư') : (contract.company_b || 'Ban QLDA / Thầu phụ')}</strong></span>
                <span><i class="fa-solid fa-file-contract"></i> Giá trị HĐ: <strong>${this.formatMoneyPlain(stats.totalSigned)}</strong></span>
                <span><i class="fa-solid fa-circle-check" style="color: var(--success);"></i> Đã ${contract.group === 'A' ? 'thu' : 'thanh toán'}: <strong style="color: ${contract.group === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatMoneyPlain(dnTotalPaid)}</strong></span>
                <span><i class="fa-solid fa-clock-rotate-left"></i> Còn lại: <strong style="color: ${dnTotalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(dnTotalRemaining)}</strong></span>
              </div>
              <button id="btn-add-payment-entry" class="button button-primary" style="font-size: 0.8rem; padding: 5px 12px;">
                <i class="fa-solid fa-plus"></i> Thêm đợt TT
              </button>
            </div>

            <div class="data-table-wrap modal-sticky-table-wrap" style="border: 1px solid #e2e8f0; border-radius: 6px; overflow-x: auto; overflow-y: auto; max-height: calc(88vh - 175px); position: relative; background: #ffffff;">
              <table class="data-table disburse-detail-table" style="min-width: 1060px; width: 100%; margin: 0;">
                <thead>
                  <tr>
                    <th style="width: 55px; min-width: 55px; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Đợt TT</th>
                    <th style="width: 85px; min-width: 85px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Ngày TT</th>
                    <th style="width: 170px; max-width: 200px; min-width: 130px; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Nội dung thanh toán</th>
                    <th style="width: 140px; min-width: 135px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">KL hoàn thành</th>
                    <th style="width: 140px; min-width: 135px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">KL chấp nhận TT</th>
                    <th style="width: 140px; min-width: 135px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Đề nghị TT</th>
                    <th style="width: 140px; min-width: 135px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Đã ${contract.group === 'A' ? 'thu' : 'thanh toán'}</th>
                    <th style="width: 130px; min-width: 125px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Còn lại</th>
                    <th style="width: 70px; min-width: 70px; text-align: center; position: sticky; top: 0; z-index: 10; background: #f8fafc; box-shadow: 0 1px 2px rgba(0,0,0,0.06);">Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  ${
                    realPayments.length === 0
                      ? '<tr><td colspan="9" style="text-align:center; padding: 25px; color: #94a3b8;"><i class="fa-regular fa-folder-open" style="font-size: 1.5rem; display: block; margin-bottom: 6px;"></i>Chưa có đợt thanh toán nào. Bấm "+ Thêm đợt TT" để tạo mới.</td></tr>'
                      : realPayments
                          .map((p, idx) => {
                            const pRate = this.getPaymentRate(p);
                            const pActRaw = Number(p.actual_volume) || 0;
                            const pAccRaw = Number(p.accepted_vol) || 0;
                            const pReqRaw = Number(p.request_amount) || 0;
                            const pPaidRaw = Number(p.paid_amount) || 0;
                            const pActVnd = Math.round(pActRaw * pRate);
                            const pAccVnd = Math.round(pAccRaw * pRate);
                            const pReqVnd = Math.round(pReqRaw * pRate);
                            const pPaidVnd = Math.round(pPaidRaw * pRate);
                            const pRemVnd = pReqVnd - pPaidVnd;
                            const rawNo = String(p.payment_no || (idx + 1)).trim();
                            const cleanNo = rawNo.replace(/^ĐN[-_\s]*/i, '').replace(/^DN[-_\s]*/i, '') || String(idx + 1);
                            const safeContent = this.getPaymentContent(p);

                            return `
                        <tr>
                          <td style="text-align: center; font-weight: 700;">
                            <span style="background: #e0f2fe; color: #0369a1; padding: 2px 7px; border-radius: 4px; font-size: 0.74rem;">
                              ${cleanNo}
                            </span>
                          </td>
                          <td style="text-align: center; color: #475569; white-space: nowrap;">
                            ${this.formatDate(p.payment_date)}
                          </td>
                          <td style="max-width: 200px; word-break: break-word; line-height: 1.35; padding: 5px 8px;">
                            <span style="font-weight: 600; color: #1e293b;">${safeContent}</span>
                            ${p.invoice_ref ? `<span class="badge badge-primary" style="font-size: 0.68rem; margin-left: 4px; padding: 1px 5px;">HĐ: ${p.invoice_ref}</span>` : ''}
                            ${p.note ? `<div style="font-size: 0.72rem; color: #64748b; margin-top: 1px;"><i class="fa-regular fa-note-sticky"></i> ${p.note}</div>` : ''}
                            ${pRate !== 1 ? `<div style="font-size: 0.68rem; color: #64748b;">Tỷ giá TT: ${pRate}</div>` : ''}
                          </td>
                          <td style="text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-size: 0.77rem; padding: 5px 8px;">
                            ${this.formatMoneyPlain(pActVnd)}
                          </td>
                          <td style="text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-size: 0.77rem; padding: 5px 8px;">
                            ${this.formatMoneyPlain(pAccVnd)}
                          </td>
                          <td style="text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 600; font-size: 0.77rem; padding: 5px 8px;">
                            ${this.formatMoneyPlain(pReqVnd)}
                          </td>
                          <td style="text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 700; font-size: 0.77rem; padding: 5px 8px; color: ${pPaidVnd < 0 ? 'var(--danger)' : (contract.group === 'A' ? 'var(--success)' : '#0891b2')};">
                            ${this.formatMoneyPlain(pPaidVnd)}
                          </td>
                          <td style="text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-size: 0.77rem; padding: 5px 8px; color: ${pRemVnd > 0 ? 'var(--danger)' : '#64748b'};">
                            ${this.formatMoneyPlain(pRemVnd)}
                          </td>
                          <td style="text-align: center; white-space: nowrap;">
                            <button class="button button-outline btn-edit-payment" data-pay-id="${p.id}" style="padding: 2px 6px; font-size: 0.73rem; color: var(--primary); margin-right: 3px;" title="Chỉnh sửa đợt thanh toán">
                              <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button class="button button-outline btn-delete-payment" data-pay-id="${p.id}" style="padding: 2px 5px; font-size: 0.73rem; color: var(--danger);" title="Xóa đợt thanh toán">
                              <i class="fa-solid fa-trash"></i>
                            </button>
                          </td>
                        </tr>
                      `;
                          })
                          .join('')
                  }
                </tbody>
                <tfoot>
                  <tr style="background: #f8fafc; font-weight: 700; border-top: 2px solid #cbd5e1;">
                    <td colspan="3" style="padding: 6px 8px; color: #1e293b; text-align: left; font-size: 0.77rem;">
                      TỔNG CỘNG (${realPayments.length} đợt ${contract.group === 'A' ? 'thu' : 'chi'})
                    </td>
                    <td style="text-align: right; padding: 6px 8px; font-size: 0.77rem; white-space: nowrap; font-variant-numeric: tabular-nums;">${this.formatMoneyPlain(dnTotalActual)}</td>
                    <td style="text-align: right; padding: 6px 8px; font-size: 0.77rem; white-space: nowrap; font-variant-numeric: tabular-nums;">${this.formatMoneyPlain(dnTotalAccepted)}</td>
                    <td style="text-align: right; padding: 6px 8px; font-size: 0.77rem; white-space: nowrap; font-variant-numeric: tabular-nums;">${this.formatMoneyPlain(dnTotalRequested)}</td>
                    <td style="text-align: right; padding: 6px 8px; font-size: 0.77rem; white-space: nowrap; font-variant-numeric: tabular-nums; color: ${contract.group === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatMoneyPlain(dnTotalPaid)}</td>
                    <td style="text-align: right; padding: 6px 8px; font-size: 0.77rem; white-space: nowrap; font-variant-numeric: tabular-nums; color: ${dnTotalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(dnTotalRemaining)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>

          <!-- PHÂN KHU 4: DASHBOARD TIẾN ĐỘ & XUẤT BÁO CÁO -->
          <div id="m-tab-dash" class="modal-tab-pane hidden">
            <div class="card" style="margin-bottom: 20px;">
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; flex-wrap: wrap; gap: 8px;">
                <h3 class="card-title" style="margin-bottom: 0;">
                  <i class="fa-solid fa-chart-line" style="color: var(--primary);"></i> Tiến độ Thực hiện & Giải ngân
                </h3>
                <div style="font-size: 0.85rem; display: flex; gap: 14px; align-items: center; flex-wrap: wrap;">
                  <span>Tạm ứng: <strong style="color: #d97706; font-size: 1.05rem;">${advanceRate}%</strong></span>
                  <span>Thanh toán KL: <strong style="color: #059669; font-size: 1.05rem;">${payVolumeRate}%</strong></span>
                  <span>Tổng Giải ngân: <strong style="color: #2563eb; font-size: 1.15rem;">${paidRate}%</strong></span>
                  <span>Nghiệm thu: <strong style="color: #ea580c; font-size: 1.05rem;">${acceptedRate}%</strong></span>
                </div>
              </div>

              <!-- Thanh tiến độ có phân biệt màu riêng cho Tạm ứng và Thanh toán khối lượng -->
              <div class="dashboard-progress-container">
                <div class="dashboard-progress-track">
                  <!-- Phân đoạn 1: Tạm ứng (Màu vàng hổ phách) -->
                  ${advanceRate > 0 ? `
                    <div class="dashboard-progress-advance" style="width: ${advanceRate}%;" title="Đã tạm ứng: ${this.formatCurrency(totalAdvance)} (${advanceRate}%)">
                      ${advanceRate >= 8 ? `<span class="progress-bar-label">Tạm ứng ${advanceRate}%</span>` : ''}
                    </div>
                  ` : ''}

                  <!-- Phân đoạn 2: Thanh toán khối lượng (Màu xanh lục) -->
                  ${payVolumeRate > 0 ? `
                    <div class="dashboard-progress-volume" style="width: ${payVolumeRate}%; ${advanceRate === 0 ? 'border-radius: 9999px 0 0 9999px;' : ''}" title="Thanh toán khối lượng: ${this.formatCurrency(paidVolume)} (${payVolumeRate}%)">
                      ${payVolumeRate >= 8 ? `<span class="progress-bar-label">TT KL ${payVolumeRate}%</span>` : ''}
                    </div>
                  ` : ''}

                  <!-- Vạch đánh dấu giá trị nghiệm thu (Màu cam) -->
                  <div class="dashboard-progress-marker-accepted" style="left: ${Math.min(100, acceptedRate)}%;" title="Mốc nghiệm thu: ${this.formatCurrency(stats.totalAccepted)} (${acceptedRate}%)">
                    <div class="dashboard-marker-pin">
                      <i class="fa-solid fa-location-pin"></i> Nghiệm thu: ${acceptedRate}%
                    </div>
                  </div>
                </div>
              </div>

              <!-- Lưới 6 Thẻ Thống kê chỉ số phân định chuẩn xác 2 hàng x 3 cột -->
              <div style="margin-top: 24px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px;">
                <!-- HÀNG 1: Căn cứ HĐ, Tạm ứng & Thanh toán khối lượng -->
                <div class="stat-box" style="padding: 12px 14px; background: #f8fafc; border: 1px solid #e2e8f0;">
                  <div class="stat-label"><i class="fa-solid fa-file-contract"></i> Giá trị HĐ / Tạm tính</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: #1e293b;">${this.formatCurrency(baseTarget)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: #64748b;">100%</div>
                </div>
                <div class="stat-box" style="padding: 12px 14px; background: #fffbeb; border: 1px solid #fde68a;">
                  <div class="stat-label" style="color: #b45309;"><i class="fa-solid fa-coins"></i> Giá trị Tạm ứng</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: #d97706;">${this.formatCurrency(totalAdvance)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: #b45309;">${advanceRate}%</div>
                </div>
                <div class="stat-box" style="padding: 12px 14px; background: #f0fdf4; border: 1px solid #bbf7d0;">
                  <div class="stat-label" style="color: #166534;"><i class="fa-solid fa-file-invoice-dollar"></i> Thanh toán Khối lượng</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: #15803d;">${this.formatCurrency(paidVolume)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: #166534;">${payVolumeRate}%</div>
                </div>

                <!-- HÀNG 2: Tổng nghiệm thu, Tổng đã giải ngân & Công nợ tồn -->
                <div class="stat-box" style="padding: 12px 14px; background: #fff7ed; border: 1px solid #fed7aa;">
                  <div class="stat-label" style="color: #c2410c;"><i class="fa-solid fa-clipboard-check"></i> Tổng Nghiệm thu</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: #ea580c;">${this.formatCurrency(stats.totalAccepted)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: #c2410c;">${acceptedRate}%</div>
                </div>
                <div class="stat-box" style="padding: 12px 14px; background: #eff6ff; border: 1px solid #bfdbfe;">
                  <div class="stat-label" style="color: #1e40af;"><i class="fa-solid fa-money-bill-transfer"></i> Tổng Đã giải ngân</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: #2563eb;">${this.formatCurrency(stats.totalPaid)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: #1e40af;">${paidRate}%</div>
                </div>
                <div class="stat-box" style="padding: 12px 14px; background: ${stats.remaining > 0 ? '#fef2f2' : '#f8fafc'}; border: 1px solid ${stats.remaining > 0 ? '#fecaca' : '#e2e8f0'};">
                  <div class="stat-label" style="color: ${stats.remaining > 0 ? '#b91c1c' : '#64748b'};"><i class="fa-solid fa-clock-rotate-left"></i> Công nợ tồn / Còn lại</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: ${stats.remaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatCurrency(stats.remaining)}</div>
                  <div class="stat-note" style="font-size: 0.82rem; font-weight: 700; color: ${stats.remaining > 0 ? '#b91c1c' : '#64748b'};">${remainingRate}%</div>
                </div>
              </div>
            </div>

            <!-- Xuất Báo cáo theo mẫu CPMS -->
            <div class="card">
              <h3 class="card-title"><i class="fa-solid fa-file-excel" style="color: #10b981;"></i> Xuất Báo cáo Dữ liệu theo Mẫu</h3>
              <p class="card-desc">Tùy chọn các nội dung muốn xuất ra bảng tính Excel để in hoặc lưu trữ báo cáo:</p>
              
              <div style="display: flex; flex-direction: column; gap: 10px; margin-bottom: 20px;">
                <label style="display: flex; align-items: center; gap: 8px; font-size: 0.9rem; cursor: pointer;">
                  <input type="checkbox" id="chk-exp-info" checked>
                  <strong>1. Thông tin chi tiết ${isMain ? 'Hợp đồng & Phụ lục' : 'Phụ lục này'}</strong>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; font-size: 0.9rem; cursor: pointer;">
                  <input type="checkbox" id="chk-exp-pay" checked>
                  <strong>2. Bảng kê Dòng tiền & Các đợt Thanh toán (Payment Ledger)</strong>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; font-size: 0.9rem; cursor: pointer;">
                  <input type="checkbox" id="chk-exp-docs" checked>
                  <strong>3. Danh mục Hồ sơ, Công văn theo dõi (Document Tracker)</strong>
                </label>
              </div>

              <div style="display: flex; gap: 12px; align-items: center;">
                <button id="btn-export-excel-report" class="button button-primary">
                  <i class="fa-solid fa-download"></i> Tải về File Báo cáo (.CSV / Excel)
                </button>
                <span id="export-msg" style="font-size: 0.85rem; color: var(--success);"></span>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    const handleKeydown = (e) => {
      if (e.key === 'Escape' && !modal.classList.contains('hidden')) {
        closeModal();
      }
    };

    const closeModal = () => {
      modal.classList.add('hidden');
      document.removeEventListener('keydown', handleKeydown);
    };

    document.getElementById('btn-close-modal')?.addEventListener('click', closeModal);
    this.bindModalOutsideClick(modal, closeModal);
    document.addEventListener('keydown', handleKeydown);
    modal.classList.remove('hidden');

    // Sự kiện Breadcrumb
    document.getElementById('btn-back-to-parent')?.addEventListener('click', (e) => {
      const pid = e.currentTarget.getAttribute('data-parent-id');
      this.openContractDetail(pid);
    });

    // Sự kiện Click vào từng phụ lục trong Tab 1
    modal.querySelectorAll('.clickable-appendix-row').forEach((el) => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const subId = el.getAttribute('data-sub-id');
        this.openContractDetail(subId);
      });
    });

    // Chuyển tab trong modal
    const modalTabBtns = modal.querySelectorAll('.modal-tab-btn');
    const docFilterWrap = modal.querySelector('#docs-category-filter-wrap');
    modalTabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        modalTabBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        const targetId = btn.getAttribute('data-target');
        modal.querySelectorAll('.modal-tab-pane').forEach((pane) => {
          if (pane.id === targetId) pane.classList.remove('hidden');
          else pane.classList.add('hidden');
        });
        if (docFilterWrap) {
          docFilterWrap.style.display = (targetId === 'm-tab-docs') ? 'flex' : 'none';
        }
      });
    });

    // Lắng nghe bộ lọc phân loại hồ sơ (Tab 2)
    const categorySelect = modal.querySelector('#select-doc-category-filter');
    if (categorySelect) {
      categorySelect.addEventListener('change', () => {
        const selectedCat = (categorySelect.value || '').trim().toLowerCase();
        const tbody = modal.querySelector('#m-tab-docs tbody');
        if (!tbody) return;
        const rows = tbody.querySelectorAll('tr.doc-item-row');
        let visibleCount = 0;
        rows.forEach((row) => {
          const rowCat = (row.getAttribute('data-doc-category') || '').trim().toLowerCase();
          if (!selectedCat || rowCat === selectedCat) {
            row.style.display = '';
            visibleCount++;
          } else {
            row.style.display = 'none';
          }
        });

        let emptyRow = tbody.querySelector('.doc-empty-filter-row');
        if (visibleCount === 0 && rows.length > 0) {
          if (!emptyRow) {
            emptyRow = document.createElement('tr');
            emptyRow.className = 'doc-empty-filter-row';
            tbody.appendChild(emptyRow);
          }
          emptyRow.style.display = '';
          emptyRow.innerHTML = `<td colspan="8" style="text-align: center; padding: 25px; color: #94a3b8;"><i class="fa-regular fa-folder-open" style="font-size: 1.5rem; display: block; margin-bottom: 6px;"></i>Không có hồ sơ nào thuộc phân loại "${categorySelect.value}".</td>`;
        } else if (emptyRow) {
          emptyRow.style.display = 'none';
        }

        const docCountHeader = modal.querySelector('#doc-list-heading-count');
        if (docCountHeader) {
          docCountHeader.textContent = selectedCat ? `${visibleCount}/${docs.length}` : `${docs.length}`;
        }
      });
    }

    // Đảo chiều sắp xếp thứ tự hồ sơ theo ngày nhận khi click Header cột Ngày nhận
    let docSortAsc = true;
    const thSortDocDate = modal.querySelector('#th-sort-doc-date');
    if (thSortDocDate) {
      thSortDocDate.addEventListener('click', () => {
        const tbody = modal.querySelector('#m-tab-docs tbody');
        if (!tbody) return;
        docSortAsc = !docSortAsc;
        const rows = Array.from(tbody.querySelectorAll('tr.doc-item-row'));
        rows.sort((rowA, rowB) => {
          const timeA = Number(rowA.getAttribute('data-received-time')) || 0;
          const timeB = Number(rowB.getAttribute('data-received-time')) || 0;
          if (timeA && timeB) {
            return docSortAsc ? (timeA - timeB) : (timeB - timeA);
          } else if (timeA && !timeB) {
            return docSortAsc ? -1 : 1;
          } else if (!timeA && timeB) {
            return docSortAsc ? 1 : -1;
          }
          return 0;
        });
        rows.forEach((r) => tbody.appendChild(r));
        const icon = thSortDocDate.querySelector('i');
        if (icon) {
          icon.className = docSortAsc ? 'fa-solid fa-arrow-down-short-wide' : 'fa-solid fa-arrow-up-wide-short';
        }
      });
    }

    // Mở Drive Folder (Tự động mở trực tiếp thư mục lưu trữ trên máy tính)
    document.getElementById('btn-open-drive-folder')?.addEventListener('click', async () => {
      const docPath = this.getDriveDocPathForContract(contract);
      const drivePath = this.getDrivePathForContract(contract);
      const targetPath = docPath || drivePath;
      if (targetPath) {
        await this.openLocalPath(targetPath, true);
      } else {
        alert(`Đường dẫn thư mục Google Drive:\n\n1. Hợp đồng & Phụ lục:\n${drivePath}\n\n2. Thư mục Doc (Tài liệu khác):\n${docPath}\n\n(Tự động đồng bộ với Google Drive Cloud & Máy tính)`);
      }
    });

    // Sao chép Hợp đồng
    document.getElementById('btn-clone-current-contract')?.addEventListener('click', () => {
      this.openCloneContractModal(contract.id);
    });

    // Sửa thông tin
    document.getElementById('btn-edit-contract-info')?.addEventListener('click', () => {
      this.openContractFormModal(contract);
    });

    // Xóa hợp đồng hiện tại
    document.getElementById('btn-delete-current-contract')?.addEventListener('click', async () => {
      await this.deleteContract(contract.id);
    });

    // Thêm phụ lục mới
    document.getElementById('btn-add-appendix')?.addEventListener('click', () => {
      this.openAppendixFormModal(contract);
    });
    document.getElementById('btn-add-appendix-empty')?.addEventListener('click', () => {
      this.openAppendixFormModal(contract);
    });

    // Thêm hồ sơ mới
    document.getElementById('btn-add-document')?.addEventListener('click', () => {
      this.openDocumentFormModal(contract);
    });

    // Nhập đợt thanh toán
    document.getElementById('btn-add-payment-entry')?.addEventListener('click', () => {
      this.openPaymentEntryModal(contract);
    });

    // Thao tác với hồ sơ: Mở file, Mở thư mục, Chỉnh sửa, Xóa
    modal.querySelectorAll('.btn-open-doc-file').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const docId = btn.getAttribute('data-doc-id');
        const doc = (this.documents || []).find((d) => d.id === docId);
        if (doc) this.handleOpenDocFile(doc);
      });
    });

    modal.querySelectorAll('.btn-open-doc-folder').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const docId = btn.getAttribute('data-doc-id');
        const doc = (this.documents || []).find((d) => d.id === docId);
        if (doc) this.handleOpenDocFolder(doc, contract);
      });
    });

    modal.querySelectorAll('.btn-edit-doc').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const docId = btn.getAttribute('data-doc-id');
        const doc = (this.documents || []).find((d) => d.id === docId);
        if (doc) this.openDocumentFormModal(contract, doc);
      });
    });

    // Nhấp đúp dòng hồ sơ để mở popup sửa hồ sơ
    modal.querySelectorAll('tr.doc-item-row').forEach((row) => {
      row.addEventListener('dblclick', (e) => {
        if (e.target.closest('button, a')) return;
        const docId = row.getAttribute('data-doc-id');
        const doc = (this.documents || []).find((d) => d.id === docId);
        if (doc) this.openDocumentFormModal(contract, doc);
      });
    });

    // Xóa hồ sơ
    modal.querySelectorAll('.btn-delete-doc').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        const docId = btn.getAttribute('data-doc-id');
        const ok = await this.confirmAsync({
          title: 'Xác nhận xóa hồ sơ',
          message: `Bạn có chắc muốn xóa hồ sơ "${docId}" không?`,
          confirmText: 'Xóa hồ sơ',
          cancelText: 'Hủy bỏ',
          confirmType: 'danger',
        });
        if (ok) {
          await db.delete('documents', docId);
          await this.loadData();
          await this.syncToLocalFile();
          this.openContractDetail(contract.id);
        }
      });
    });

    // Sao chép đường dẫn file local của hồ sơ
    modal.querySelectorAll('.btn-copy-doc-local').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const p = decodeURIComponent(btn.getAttribute('data-path') || '');
        if (p) {
          navigator.clipboard.writeText(p).then(() => {
            const orig = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-check" style="color: var(--success);"></i>';
            setTimeout(() => { btn.innerHTML = orig; }, 1500);
          }).catch(() => {
            prompt('Đường dẫn file trên máy tính (sao chép Ctrl+C):', p);
          });
        }
      });
    });

    // Sửa đợt giải ngân (ĐN)
    modal.querySelectorAll('.btn-edit-payment').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        const payId = btn.getAttribute('data-pay-id');
        const payment = this.payments.find((p) => p.id === payId);
        if (payment) {
          this.openEditPaymentModal(payment, contract);
        }
      });
    });

    // Xóa thanh toán
    modal.querySelectorAll('.btn-delete-payment').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        const payId = btn.getAttribute('data-pay-id');
        const payment = this.payments.find((p) => p.id === payId);
        const label = payment ? `"${payment.item_name || payment.payment_no}"` : 'này';
        this.showConfirmDialog({
          title: 'Xác nhận xóa đợt thanh toán?',
          message: `Bạn có chắc muốn xóa đợt thanh toán ${label} không?`,
          confirmText: 'Xóa vĩnh viễn',
          confirmType: 'danger',
          onConfirm: async () => {
            await db.delete('payments', payId);
            await this.loadData();
            await this.syncToLocalFile();
            this.openContractDetail(contract.id);
            const tabIpcBtn = document.querySelector('#contract-detail-modal [data-target="m-tab-ipc"]');
            tabIpcBtn?.click();
            this.showToast('Đã xóa đợt thanh toán thành công!', 'success');
          },
        });
      });
    });

    // Xuất báo cáo theo mẫu
    document.getElementById('btn-export-excel-report')?.addEventListener('click', () => {
      this.exportContractReport(contract, isMain);
    });
  }

  /**
   * Chuyển nhanh (Jump Navigation) đến Hợp đồng liên kết
   */
  jumpToContract(targetId) {
    if (!targetId) return;
    const target = (this.contracts || []).find((c) => c.id === targetId);
    if (!target) {
      alert(`Không tìm thấy Hợp đồng liên kết "${targetId}" trong dự án hiện tại.`);
      return;
    }

    // Đóng bất kỳ modal nào đang mở
    document.querySelectorAll('.modal-overlay').forEach((m) => m.remove());

    // Nếu tab Quản lý Hợp đồng chưa được chọn, chuyển sang tab Hợp đồng
    const subtabHD = document.getElementById('subtab-contracts');
    if (subtabHD && !subtabHD.classList.contains('active')) {
      subtabHD.click();
    }

    // Nếu khác nhóm A/B hiện tại, tự động chuyển tab nhóm
    const targetGroup = target.group || (target.id.startsWith('B') ? 'B' : 'A');
    if (this.currentGroup !== targetGroup) {
      this.currentGroup = targetGroup;
      const tabBtns = document.querySelectorAll('.group-tab-btn');
      tabBtns.forEach((b) => {
        if (b.getAttribute('data-group') === targetGroup) {
          b.classList.add('active');
        } else {
          b.classList.remove('active');
        }
      });
      this.render();
    }

    // Cuộn tới phần tử và kích hoạt animation phát sáng
    setTimeout(() => {
      const el =
        document.getElementById(`contract-card-${target.id}`) ||
        document.querySelector(`[data-cid="${target.id}"]`) ||
        document.querySelector(`[data-clone-cid="${target.id}"]`)?.closest('.contract-item');
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.classList.add('pulse-highlight');
        setTimeout(() => el.classList.remove('pulse-highlight'), 3000);
      }
      // Mở modal chi tiết của HĐ đích
      this.openContractDetail(target.id);
    }, 150);
  }

  // ==========================================
  // MODAL KHỞI TẠO DỰ ÁN MỚI 2 BƯỚC (QUY TRÌNH CHUẨN MAIO & CPMS PLUS)
  // BƯỚC 1: Chọn nơi lưu file CPMS trên máy tính
  // BƯỚC 2: Nhập thông tin cơ bản của dự án
  // ==========================================
  openCreateProjectModal() {
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 620px; padding: 0; overflow: hidden; border-radius: var(--radius-lg); box-shadow: 0 20px 40px rgba(0,0,0,0.2);">
        <!-- Modal Header -->
        <div class="modal-header" style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #fff; padding: 20px 24px; border-bottom: none;">
          <div>
            <h3 style="font-size: 1.2rem; font-weight: 700; color: #fff; margin: 0 0 4px 0; display: flex; align-items: center; gap: 10px;">
              <i class="fa-solid fa-folder-plus" style="color: #60a5fa;"></i> Khởi tạo Dự án Mới (2 Bước)
            </h3>
            <p style="font-size: 0.825rem; color: #94a3b8; margin: 0;">Khởi tạo tệp CPMS cá nhân lưu trên máy tính và thiết lập thông tin ban đầu</p>
          </div>
          <button class="button button-outline btn-close-wizard" style="color: #cbd5e1; border-color: #475569; padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <!-- Stepper Indicator -->
        <div style="background: #f8fafc; padding: 14px 24px; border-bottom: 1px solid var(--border-color); display: flex; align-items: center; justify-content: space-between;">
          <div id="step-indicator-1" class="wizard-step" style="display: flex; align-items: center; gap: 8px; font-size: 0.85rem; font-weight: 700; color: var(--primary);">
            <span id="step-badge-1" style="width: 24px; height: 24px; border-radius: 50%; background: var(--primary); color: #fff; display: inline-flex; align-items: center; justify-content: center; font-size: 0.75rem;">1</span>
            <span>Chọn nơi lưu file .cpms</span>
          </div>
          <div style="flex: 1; height: 2px; background: #e2e8f0; margin: 0 12px;"></div>
          <div id="step-indicator-2" class="wizard-step" style="display: flex; align-items: center; gap: 8px; font-size: 0.85rem; font-weight: 600; color: #94a3b8;">
            <span id="step-badge-2" style="width: 24px; height: 24px; border-radius: 50%; background: #e2e8f0; color: #64748b; display: inline-flex; align-items: center; justify-content: center; font-size: 0.75rem;">2</span>
            <span>Nhập thông tin dự án</span>
          </div>
        </div>

        <div class="modal-body" style="padding: 24px;">
          <!-- BƯỚC 1: CHỌN NƠI LƯU TỆP CPMS MÁY TÍNH -->
          <div id="wizard-step-1" class="wizard-step-pane">
            <div style="text-align: center; padding: 10px 0 18px;">
              <div style="width: 56px; height: 56px; border-radius: 50%; background: #eff6ff; color: var(--primary); display: flex; align-items: center; justify-content: center; font-size: 1.6rem; margin: 0 auto 12px;">
                <i class="fa-solid fa-hard-drive"></i>
              </div>
              <h4 style="font-size: 1.05rem; font-weight: 700; color: #0f172a; margin-bottom: 6px;">Bước 1: Chọn nơi lưu tệp dữ liệu trên máy tính</h4>
              <p style="font-size: 0.85rem; color: var(--text-muted); line-height: 1.5; max-width: 480px; margin: 0 auto;">
                Hệ thống MAIO quản lý dữ liệu an toàn dựa trên tệp cá nhân (<code>.cpms</code>) lưu trực tiếp trên máy tính của bạn. Dữ liệu này sẽ tự động khởi tạo theo chuẩn tương thích CPMS Plus.
              </p>
            </div>

            <!-- Khung chọn vị trí lưu file -->
            <div style="background: #f8fafc; border: 1.5px dashed #cbd5e1; border-radius: var(--radius-md); padding: 24px; text-align: center;">
              <button type="button" id="btn-wizard-pick-file" class="button button-primary button-lg" style="font-weight: 700; padding: 12px 24px; margin-bottom: 8px;">
                <i class="fa-solid fa-folder-arrow-up"></i> Chọn nơi lưu tệp (.cpms) trên máy tính...
              </button>
              <p style="font-size: 0.8rem; color: #64748b; margin: 0;">Trình duyệt sẽ mở hộp thoại Lưu tệp để bạn chọn thư mục và đặt tên</p>
              <div style="margin-top: 12px;">
                <button type="button" id="btn-wizard-manual-fallback" class="button button-outline button-sm" style="font-size: 0.8rem; border-color: transparent; color: var(--primary); text-decoration: underline; background: transparent;">
                  Hoặc tự đặt tên tệp để tải về máy sau khi tạo
                </button>
              </div>
            </div>

            <!-- Vùng hiển thị tệp đã chọn -->
            <div id="wizard-file-selected-box" class="hidden" style="margin-top: 16px; background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: var(--radius-md); padding: 14px 18px; display: flex; align-items: center; gap: 12px;">
              <i class="fa-solid fa-circle-check" style="color: #10b981; font-size: 1.4rem;"></i>
              <div style="flex: 1;">
                <div style="font-size: 0.75rem; text-transform: uppercase; font-weight: 700; color: #047857;">Tệp khởi tạo sẽ lưu tại:</div>
                <strong id="wizard-file-name-label" style="font-size: 0.95rem; color: #065f46;">chua_chon.cpms</strong>
              </div>
              <button type="button" id="btn-wizard-repick-file" class="button button-outline button-sm" style="font-size: 0.8rem; background: #fff;">Đổi tệp</button>
            </div>

            <!-- Fallback nhập tên file nếu trình duyệt không hỗ trợ File System Access API -->
            <div id="wizard-fallback-filename-box" class="hidden" style="margin-top: 16px;">
              <label class="form-label" style="font-weight: 600; font-size: 0.85rem;">Tên tệp dữ liệu khởi tạo (.cpms):</label>
              <input type="text" id="wizard-fallback-filename" class="form-control" placeholder="MAIO_Project_New.cpms">
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 24px; border-top: 1px solid #f1f5f9; padding-top: 16px;">
              <button type="button" class="button button-outline btn-close-wizard">Hủy</button>
              <button type="button" id="btn-wizard-next-step" class="button button-primary" disabled>
                Tiếp tục: Nhập thông tin dự án <i class="fa-solid fa-arrow-right" style="margin-left: 6px;"></i>
              </button>
            </div>
          </div>

          <!-- BƯỚC 2: NHẬP THÔNG TIN CƠ BẢN CỦA DỰ ÁN -->
          <div id="wizard-step-2" class="wizard-step-pane hidden">
            <form id="form-wizard-create-project">
              <div style="margin-bottom: 16px; padding: 10px 14px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: var(--radius-sm); font-size: 0.85rem; color: #166534; display: flex; align-items: center; gap: 8px;">
                <i class="fa-solid fa-file-contract"></i>
                <span>Tệp lưu liên kết: <strong id="step2-file-preview">...</strong></span>
              </div>

              <div class="form-group" style="margin-bottom: 14px;">
                <label class="form-label" style="font-weight: 700; font-size: 0.875rem;">Mã dự án (Project Code / ID) <span style="color: var(--danger);">*</span></label>
                <input type="text" class="form-control" name="project_id" id="wizard-input-pid" placeholder="Vd: CHC1, MT2, DA_01..." required style="font-weight: 600;">
                <small style="color: var(--text-muted); font-size: 0.775rem;">Mã viết liền không dấu, dùng để định danh dữ liệu và các hợp đồng</small>
              </div>

              <div class="form-group" style="margin-bottom: 14px;">
                <label class="form-label" style="font-weight: 700; font-size: 0.875rem;">Tên dự án đầy đủ <span style="color: var(--danger);">*</span></label>
                <input type="text" class="form-control" name="project_name" placeholder="Vd: Dự án Xây dựng Cầu Mỹ Thuận 2..." required>
              </div>

              <div class="form-group" style="margin-bottom: 14px;">
                <label class="form-label" style="font-weight: 700; font-size: 0.875rem;">Địa điểm xây dựng</label>
                <input type="text" class="form-control" name="location" placeholder="Vd: Quận Tây Hồ, TP. Hà Nội">
              </div>

              <div class="form-grid-2" style="gap: 14px; margin-bottom: 18px;">
                <div class="form-group">
                  <label class="form-label" style="font-weight: 700; font-size: 0.875rem;">Chủ đầu tư (Bên A)</label>
                  <input type="text" class="form-control" name="client" placeholder="Vd: Ban Quản lý Dự án..." value="Chủ đầu tư Bên A">
                </div>
                <div class="form-group">
                  <label class="form-label" style="font-weight: 700; font-size: 0.875rem;">Đơn vị thi công / Nhà thầu (Bên B)</label>
                  <input type="text" class="form-control" name="contractor" placeholder="Vd: Công ty Cổ phần Xây dựng..." value="Nhà thầu Bên B">
                </div>
              </div>

              <div style="display: flex; justify-content: space-between; gap: 10px; margin-top: 24px; border-top: 1px solid #f1f5f9; padding-top: 16px;">
                <button type="button" id="btn-wizard-prev-step" class="button button-outline">
                  <i class="fa-solid fa-arrow-left"></i> Quay lại Bước 1
                </button>
                <div style="display: flex; gap: 10px;">
                  <button type="button" class="button button-outline btn-close-wizard">Hủy</button>
                  <button type="submit" id="btn-wizard-final-submit" class="button button-primary" style="font-weight: 700; padding: 10px 20px;">
                    <i class="fa-solid fa-check"></i> Hoàn tất Khởi tạo Dự án
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeWizard = () => modal.remove();
    modal.querySelectorAll('.btn-close-wizard').forEach((btn) => btn.addEventListener('click', closeWizard));

    const stepPane1 = modal.querySelector('#wizard-step-1');
    const stepPane2 = modal.querySelector('#wizard-step-2');
    const stepIndicator1 = modal.querySelector('#step-indicator-1');
    const stepIndicator2 = modal.querySelector('#step-indicator-2');
    const stepBadge1 = modal.querySelector('#step-badge-1');
    const stepBadge2 = modal.querySelector('#step-badge-2');
    const btnNextStep = modal.querySelector('#btn-wizard-next-step');
    const btnPrevStep = modal.querySelector('#btn-wizard-prev-step');
    const fileSelectedBox = modal.querySelector('#wizard-file-selected-box');
    const fileNameLabel = modal.querySelector('#wizard-file-name-label');
    const step2FilePreview = modal.querySelector('#step2-file-preview');
    const inputPid = modal.querySelector('#wizard-input-pid');
    const fallbackBox = modal.querySelector('#wizard-fallback-filename-box');
    const fallbackInput = modal.querySelector('#wizard-fallback-filename');

    let chosenFileHandle = null;
    let chosenFileName = '';

    const goToStep = (stepNumber) => {
      if (stepNumber === 1) {
        stepPane1.classList.remove('hidden');
        stepPane2.classList.add('hidden');
        stepIndicator1.style.color = 'var(--primary)';
        stepIndicator1.style.fontWeight = '700';
        stepBadge1.style.background = 'var(--primary)';
        stepBadge1.style.color = '#fff';
        stepIndicator2.style.color = '#94a3b8';
        stepIndicator2.style.fontWeight = '600';
        stepBadge2.style.background = '#e2e8f0';
        stepBadge2.style.color = '#64748b';
      } else if (stepNumber === 2) {
        stepPane1.classList.add('hidden');
        stepPane2.classList.remove('hidden');
        stepIndicator1.style.color = '#10b981';
        stepIndicator1.style.fontWeight = '600';
        stepBadge1.style.background = '#10b981';
        stepBadge1.style.color = '#fff';
        stepIndicator2.style.color = 'var(--primary)';
        stepIndicator2.style.fontWeight = '700';
        stepBadge2.style.background = 'var(--primary)';
        stepBadge2.style.color = '#fff';

        step2FilePreview.textContent = chosenFileName || 'MAIO_Project_New.cpms';
      }
    };

    const pickSaveFile = async () => {
      if ('showSaveFilePicker' in window) {
        try {
          const nowStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
          const handle = await window.showSaveFilePicker({
            suggestedName: `MAIO_Project_${nowStr}.cpms`,
            types: [
              {
                description: 'Tệp dữ liệu CPMS (*.cpms)',
                accept: { 'application/octet-stream': ['.cpms'] }
              }
            ]
          });

          if (handle) {
            chosenFileHandle = handle;
            chosenFileName = handle.name;
            fileNameLabel.textContent = chosenFileName;
            fileSelectedBox.classList.remove('hidden');
            fallbackBox.classList.add('hidden');
            btnNextStep.disabled = false;

            // Gợi ý mã dự án tự động từ tên tệp nếu người dùng chưa nhập
            const rawCode = chosenFileName.replace(/\.[^/.]+$/, '').replace(/[^a-zA-Z0-9_]/g, '_').toUpperCase();
            if (inputPid && !inputPid.value) {
              inputPid.value = rawCode;
            }

            // Tự động chuyển tiếp sang bước 2 để trải nghiệm mượt mà
            goToStep(2);
          }
        } catch (err) {
          if (err.name === 'AbortError') return;
          console.warn('Không thể mở showSaveFilePicker, chuyển sang nhập tên file thủ công:', err);
          setupFallbackFile();
        }
      } else {
        setupFallbackFile();
      }
    };

    const setupFallbackFile = () => {
      fallbackBox.classList.remove('hidden');
      const nowStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      fallbackInput.value = `MAIO_Project_${nowStr}.cpms`;
      chosenFileName = fallbackInput.value;
      btnNextStep.disabled = false;
      fallbackInput.addEventListener('input', () => {
        chosenFileName = fallbackInput.value.trim() || `MAIO_Project_${nowStr}.cpms`;
        btnNextStep.disabled = !chosenFileName;
      });
    };

    modal.querySelector('#btn-wizard-pick-file')?.addEventListener('click', pickSaveFile);
    modal.querySelector('#btn-wizard-repick-file')?.addEventListener('click', pickSaveFile);
    modal.querySelector('#btn-wizard-manual-fallback')?.addEventListener('click', () => {
      setupFallbackFile();
      goToStep(2);
    });

    btnNextStep.addEventListener('click', () => {
      if (!chosenFileName) {
        alert('Vui lòng chọn nơi lưu tệp .cpms trên máy tính trước.');
        return;
      }
      goToStep(2);
    });

    btnPrevStep.addEventListener('click', () => {
      goToStep(1);
    });

    // SUBMIT BƯỚC 2: KHỞI TẠO TOÀN BỘ DỮ LIỆU DỰ ÁN & GHI TỆP LOCAL
    modal.querySelector('#form-wizard-create-project').addEventListener('submit', async (e) => {
      e.preventDefault();
      const submitBtn = modal.querySelector('#btn-wizard-final-submit');
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang khởi tạo...';

      try {
        const fd = new FormData(e.target);
        const newId = fd.get('project_id').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '_');
        const newName = fd.get('project_name').trim();
        const client = fd.get('client')?.trim() || 'Chủ đầu tư Bên A';
        const contractor = fd.get('contractor')?.trim() || 'Nhà thầu Bên B';
        const location = fd.get('location')?.trim() || '';

        const finalFileName = chosenFileName || `${newId}.cpms`;

        // 1. Lưu project vào IndexedDB
        const newProj = {
          id: newId,
          code: newId,
          name: newName,
          client: client,
          contractor: contractor,
          location: location,
          source_filename: finalFileName,
          created_at: new Date().toISOString(),
          source_file_updated: new Date().toISOString()
        };
        await db.put('projects', newProj);

        // 2. Khởi tạo sẵn mục Chi phí Ban (CPB) mặc định cho dự án mới
        const defaultCpbId = 'B-CPB-001';
        const defaultCpbContract = {
          id: defaultCpbId,
          project_id: newId,
          group: 'B',
          parent_id: null,
          level: 0,
          contract_no: defaultCpbId,
          contract_type: 'CPB',
          signing_date: new Date().toISOString().slice(0, 10),
          end_date: null,
          status: 'Active',
          company_a: newProj.client,
          company_b: 'Ban Quản lý / Nội bộ',
          company_b2: '',
          rep_b2: '',
          scope_of_work: 'Chi Phí Ban (Chi phí quản lý dự án)',
          exchange_rate: 1.0,
          estimated_val: 0,
          contract_val: 0,
          rep_a: '',
          rep_b: '',
          category: 'CPB',
          payee: 'Ban QLDA / Nội bộ',
          bank_account: '',
          bank_name: '',
          contact_person_b: '',
          created_at: new Date().toISOString().slice(0, 10),
          display_no: 'CPB',
          sort_order: 9999,
          is_cpb: true,
        };
        await db.put('contracts', defaultCpbContract);

        const defaultCpbHeader = {
          id: `HDR-${defaultCpbId}`,
          is_header: true,
          project_id: newId,
          parent_id: 'GROUP-B',
          contract_id: defaultCpbId,
          contract_no: defaultCpbId,
          item_name: 'Chi Phí Ban (Chi phí quản lý dự án)',
          exchange_rate: 1.0,
          contract_val: 0,
          actual_volume: 0,
          accepted_vol: 0,
          request_amount: 0,
          paid_amount: 0,
          remaining: 0,
        };
        await db.put('payments', defaultCpbHeader);

        // 3. Nếu có File Handle đã chọn từ showSaveFilePicker: Ghi dữ liệu ban đầu trực tiếp vào tệp máy tính
        if (chosenFileHandle) {
          this.activeFileHandle = chosenFileHandle;
          const buffer = await CPMSImporterExporter.buildCPMSBinaryBuffer(newId);
          const writable = await chosenFileHandle.createWritable();
          await writable.write(buffer);
          await writable.close();
        } else {
          // Tự động tải file CPMS khởi tạo về máy tính cho người dùng nếu không dùng File System API
          try {
            await CPMSImporterExporter.exportToCPMSFormat(newId, 'cpms');
          } catch (e) {
            console.warn('Lỗi xuất file tải về:', e);
          }
        }

        // 4. Đổi sang dự án vừa tạo & cập nhật toàn bộ hệ thống
        this.currentProjectId = newId;
        await this.loadData();
        this.render();
        if (window.app) {
          await window.app.updateDashboardStats();
        }
        this.updateEmptyStateUI();

        if (this.activeFileHandle) {
          this.showLocalSaveStatus(`Đã liên kết tệp "${finalFileName}". Mọi chỉnh sửa sẽ tự động lưu vào tệp này!`);
        }

        closeWizard();
        alert(`Đã khởi tạo thành công dự án "${newName}"!\nDữ liệu ban đầu đã được thiết lập vào tệp "${finalFileName}".`);
      } catch (err) {
        console.error('Lỗi khởi tạo dự án:', err);
        alert('Lỗi khởi tạo dự án: ' + err.message);
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<i class="fa-solid fa-check"></i> Hoàn tất Khởi tạo Dự án';
      }
    });
  }

  // ==========================================
  // MODAL QUẢN LÝ DỰ ÁN (TẠO MỚI HOẶC IMPORT .CPMS)
  // ==========================================
  openProjectModal() {
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 650px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700;">
            <i class="fa-solid fa-laptop-file" style="color: var(--primary);"></i> Dẫn link & Nhập file Data CPMS từ máy tính
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <div style="padding: 0 20px; background: #fff; border-bottom: 1px solid var(--border-color);">
          <div class="tab-nav" style="margin-bottom: 0;">
            <button class="tab-btn active prj-tab-btn" data-target="tab-import-cpms">
              <i class="fa-solid fa-file-import"></i> 1. Dẫn file CPMS Local (.cpms / .xlsx)
            </button>
            <button class="tab-btn prj-tab-btn" data-target="tab-create-manual">
              <i class="fa-solid fa-pen-nib"></i> 2. Tạo dự án mới thủ công
            </button>
          </div>
        </div>

        <div class="modal-body">
          <!-- TAB 1: IMPORT FILE .CPMS -->
          <div id="tab-import-cpms" class="prj-tab-pane">
            <p style="font-size: 0.9rem; color: var(--text-muted); margin-bottom: 16px;">
              Chọn file dữ liệu dự án trên máy tính của bạn (định dạng <code>.cpms</code>, <code>.xlsb</code>, hoặc <code>.xlsx</code> được xuất từ CPMS Plus). Dữ liệu chỉ hiển thị trong tài khoản của bạn, an toàn và bảo mật.
            </p>

            <div class="upload-dropzone" id="dropzone-cpms">
              <i class="fa-solid fa-cloud-arrow-up upload-icon"></i>
              <h4 style="font-size: 1.05rem; font-weight: 700; margin-bottom: 6px;">Kéo thả file .cpms vào đây</h4>
              <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 14px;">hoặc bấm vào khung để chọn file từ máy tính của bạn</p>
              <input type="file" id="file-input-cpms" accept=".cpms,.xlsb,.xlsx" style="display: none;">
              <button type="button" class="button button-outline" id="btn-browse-cpms">
                <i class="fa-solid fa-folder-open"></i> Chọn file từ máy tính
              </button>
            </div>

            <!-- Vùng hiển thị file đã chọn -->
            <div id="selected-file-info" class="hidden" style="margin-top: 16px; background: var(--bg-card-subtle); padding: 14px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
              <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                  <strong id="file-name-label" style="color: var(--primary);">tên_file.cpms</strong>
                  <div id="file-size-label" style="font-size: 0.8rem; color: var(--text-muted);">0 KB</div>
                </div>
                <button type="button" id="btn-clear-file" class="button button-outline" style="padding: 4px 8px; font-size: 0.8rem;">
                  <i class="fa-solid fa-xmark"></i>
                </button>
              </div>
            </div>

            <div id="import-status-msg" style="margin-top: 14px; font-size: 0.9rem; display: none;"></div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="button" class="button button-primary" id="btn-start-import" disabled>
                <i class="fa-solid fa-upload"></i> Bắt đầu Nạp dữ liệu
              </button>
            </div>
          </div>

          <!-- TAB 2: TẠO THỦ CÔNG -->
          <div id="tab-create-manual" class="prj-tab-pane hidden">
            <form id="form-create-project">
              <div class="form-group">
                <label class="form-label">Mã dự án (Project Code / ID)</label>
                <input type="text" class="form-control" name="project_id" placeholder="Vd: CHC1, MT2, VEC..." required>
              </div>
              <div class="form-group">
                <label class="form-label">Tên dự án đầy đủ</label>
                <input type="text" class="form-control" name="project_name" placeholder="Vd: Dự án Xây dựng Cầu Mỹ Thuận 2..." required>
              </div>
              <div class="form-grid-2">
                <div class="form-group">
                  <label class="form-label">Chủ đầu tư / Khách hàng</label>
                  <input type="text" class="form-control" name="client" placeholder="Vd: Ban QLDA 7 - Bộ GTVT">
                </div>
                <div class="form-group">
                  <label class="form-label">Địa điểm xây dựng</label>
                  <input type="text" class="form-control" name="location" placeholder="Vd: Tiền Giang - Vĩnh Long">
                </div>
              </div>
              <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
                <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
                <button type="submit" class="button button-primary">Tạo Dự án</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Chuyển tab trong modal
    const tabBtns = modal.querySelectorAll('.prj-tab-btn');
    tabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        tabBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        const targetId = btn.getAttribute('data-target');
        modal.querySelectorAll('.prj-tab-pane').forEach((pane) => {
          if (pane.id === targetId) pane.classList.remove('hidden');
          else pane.classList.add('hidden');
        });
      });
    });

    // Xử lý kéo thả / chọn file .cpms
    const dropzone = modal.querySelector('#dropzone-cpms');
    const fileInput = modal.querySelector('#file-input-cpms');
    const btnBrowse = modal.querySelector('#btn-browse-cpms');
    const fileInfoBox = modal.querySelector('#selected-file-info');
    const fileNameLabel = modal.querySelector('#file-name-label');
    const fileSizeLabel = modal.querySelector('#file-size-label');
    const btnClear = modal.querySelector('#btn-clear-file');
    const btnStartImport = modal.querySelector('#btn-start-import');
    const statusMsg = modal.querySelector('#import-status-msg');

    let selectedFile = null;

    const handleFile = (file) => {
      if (!file) return;
      selectedFile = file;
      fileNameLabel.textContent = file.name;
      fileSizeLabel.textContent = `${(file.size / 1024).toFixed(1)} KB — Định dạng CPMS`;
      fileInfoBox.classList.remove('hidden');
      btnStartImport.disabled = false;
      statusMsg.style.display = 'none';
    };

    btnBrowse.addEventListener('click', async () => {
      if ('showOpenFilePicker' in window) {
        try {
          const [handle] = await window.showOpenFilePicker({
            types: [
              {
                description: 'File dữ liệu CPMS (*.cpms, *.xlsx, *.xlsb)',
                accept: {
                  'application/octet-stream': ['.cpms', '.xlsb'],
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                }
              }
            ],
            multiple: false
          });
          if (handle) {
            this.activeFileHandle = handle;
            const file = await handle.getFile();
            handleFile(file);
            return;
          }
        } catch (err) {
          if (err.name === 'AbortError') return;
        }
      }
      fileInput.click();
    });
    dropzone.addEventListener('click', (e) => {
      if (e.target !== btnBrowse && !btnBrowse.contains(e.target)) {
        fileInput.click();
      }
    });

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFile(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFile(e.target.files[0]);
      }
    });

    btnClear.addEventListener('click', () => {
      selectedFile = null;
      fileInput.value = '';
      fileInfoBox.classList.add('hidden');
      btnStartImport.disabled = true;
    });

    // Thực hiện Import file .cpms
    btnStartImport.addEventListener('click', async () => {
      if (!selectedFile) return;

      btnStartImport.disabled = true;
      btnStartImport.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đọc file CPMS...';
      statusMsg.style.display = 'block';
      statusMsg.className = 'badge badge-warning';
      statusMsg.innerHTML = 'Đang phân tích cấu trúc Hợp đồng, Payment Ledger và Hồ sơ...';

      try {
        const result = await CPMSImporterExporter.importCPMSFile(selectedFile);
        statusMsg.className = 'badge badge-success';
        statusMsg.innerHTML = `Đã nạp thành công: ${result.contractsCount} HĐ/PL, ${result.paymentsCount} Đợt thanh toán, ${result.docsCount} Hồ sơ!`;

        // Đổi sang dự án vừa import
        this.currentProjectId = result.projectData.id;
        await this.loadData();
        this.render();

        if (this.activeFileHandle) {
          this.showLocalSaveStatus(`Đã liên kết tệp "${this.activeFileHandle.name}". Mọi chỉnh sửa sẽ tự động lưu vào tệp này!`);
        }

        setTimeout(() => {
          closeSubModal();
          alert(`Đã nạp thành công dự án "${result.projectData.name}"!\nBao gồm:\n- ${result.contractsCount} Hợp đồng & Phụ lục\n- ${result.paymentsCount} Đợt thanh toán (IPC)\n- ${result.docsCount} Hồ sơ theo dõi.`);
        }, 500);
      } catch (err) {
        console.error('Lỗi import:', err);
        statusMsg.className = 'badge badge-danger';
        statusMsg.innerHTML = 'Lỗi nạp file: ' + err.message;
        btnStartImport.disabled = false;
        btnStartImport.innerHTML = '<i class="fa-solid fa-upload"></i> Thử lại';
      }
    });

    // Thực hiện Tạo thủ công
    modal.querySelector('#form-create-project').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const newId = fd.get('project_id').trim();
      const newName = fd.get('project_name').trim();

      const newProj = {
        id: newId,
        code: newId,
        name: newName,
        client: fd.get('client')?.trim() || 'Chủ đầu tư',
        location: fd.get('location')?.trim() || '',
        created_at: new Date().toISOString(),
      };

      await db.put('projects', newProj);

      // Tự động khởi tạo sẵn mục Chi phí Ban (CPB) mặc định cho dự án mới
      const defaultCpbId = 'B-CPB-001';
      const defaultCpbContract = {
        id: defaultCpbId,
        project_id: newId,
        group: 'B',
        parent_id: null,
        level: 0,
        contract_no: defaultCpbId,
        contract_type: 'CPB',
        signing_date: new Date().toISOString().slice(0, 10),
        end_date: null,
        status: 'Active',
        company_a: newProj.client || 'Chủ đầu tư',
        company_b: 'Ban Quản lý / Nội bộ',
        company_b2: '',
        rep_b2: '',
        scope_of_work: 'Chi Phí Ban (Chi phí quản lý dự án)',
        exchange_rate: 1.0,
        estimated_val: 0,
        contract_val: 0,
        rep_a: '',
        rep_b: '',
        category: 'CPB',
        payee: 'Ban QLDA / Nội bộ',
        bank_account: '',
        bank_name: '',
        contact_person_b: '',
        created_at: new Date().toISOString().slice(0, 10),
        display_no: 'CPB',
        sort_order: 9999,
        is_cpb: true,
      };
      await db.put('contracts', defaultCpbContract);

      const defaultCpbHeader = {
        id: `HDR-${defaultCpbId}`,
        is_header: true,
        project_id: newId,
        parent_id: 'GROUP-B',
        contract_id: defaultCpbId,
        contract_no: defaultCpbId,
        item_name: 'Chi Phí Ban (Chi phí quản lý dự án)',
        exchange_rate: 1.0,
        contract_val: 0,
        actual_volume: 0,
        accepted_vol: 0,
        request_amount: 0,
        paid_amount: 0,
        remaining: 0,
      };
      await db.put('payments', defaultCpbHeader);

      this.currentProjectId = newId;
      await this.loadData();
      this.render();
      this.updateEmptyStateUI();
      if (window.app) {
        await window.app.updateDashboardStats();
      }
      closeSubModal();
      alert(`Đã tạo thành công dự án "${newName}" kèm mục Chi phí Ban (CPB) quản lý nội bộ!`);
    });
  }

  // ==========================================
  // MODAL THÊM / CHỈNH SỬA CHI PHÍ BAN (CPB)
  // ==========================================
  openCreateCPBModal(editCpb = null) {
    const isEdit = !!editCpb;
    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || {};

    // Sinh mã tự động B-CPB-001, B-CPB-002... nếu tạo mới
    let formId = '';
    if (isEdit) {
      formId = editCpb.id;
    } else {
      const cpbContracts = this.contracts.filter(
        (c) => c.is_cpb || c.contract_type === 'CPB' || (c.id && c.id.startsWith('B-CPB-'))
      );
      let maxNum = 0;
      cpbContracts.forEach((c) => {
        const m = (c.id || '').match(/B-CPB-(\d+)/i);
        if (m) {
          const num = parseInt(m[1], 10);
          if (num > maxNum) maxNum = num;
        }
      });
      formId = `B-CPB-${String(maxNum + 1).padStart(3, '0')}`;
    }

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 620px;">
        <div class="modal-header" style="border-bottom: 1px solid var(--border-color); padding: 14px 20px;">
          <div>
            <div style="font-size: 0.8rem; color: #b45309; font-weight: 700; text-transform: uppercase;">
              <i class="fa-solid fa-calculator"></i> Chi phí Quản lý Dự án (Nhóm B)
            </div>
            <h3 style="font-size: 1.15rem; font-weight: 800; margin-top: 2px; color: #1e293b;">
              ${isEdit ? `Chỉnh sửa Chi phí Ban: ${editCpb.contract_no || editCpb.id}` : '+ Thêm mới Chi phí Ban (CPB)'}
            </h3>
          </div>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <div class="modal-body" style="padding: 20px;">
          <form id="form-cpb">
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label" style="display: flex; justify-content: space-between;">
                  <span>Mã khoản chi phí (CPB ID) *</span>
                  <span style="font-weight: normal; color: var(--danger); font-size: 0.75rem;"><i class="fa-solid fa-lock"></i> Tự động</span>
                </label>
                <input type="text" class="form-control" name="id" value="${formId}" readonly style="background: #f1f5f9; color: #334155; font-weight: 700; font-family: monospace; cursor: not-allowed;" required>
              </div>

              <div class="form-group">
                <label class="form-label">Tình trạng *</label>
                <select class="form-control" name="status">
                  <option value="Active" ${!isEdit || editCpb?.status === 'Active' ? 'selected' : ''}>Active (Đang thực hiện)</option>
                  <option value="Completed" ${editCpb?.status === 'Completed' ? 'selected' : ''}>Completed (Đã quyết toán)</option>
                  <option value="Suspended" ${editCpb?.status === 'Suspended' ? 'selected' : ''}>Suspended (Tạm dừng)</option>
                </select>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Tên khoản chi phí quản lý (Nội dung) *</label>
              <input type="text" class="form-control" name="scope_of_work" value="${editCpb?.scope_of_work || 'Chi Phí Ban (Chi phí quản lý dự án)'}" placeholder="Vd: Chi Phí Ban, Chi phí văn phòng điều hành công trường..." required>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Đơn vị / Ban phụ trách</label>
                <input type="text" class="form-control" name="company_b" value="${editCpb?.company_b || 'Ban Quản lý / Nội bộ'}" placeholder="Ban Điều hành / Quản lý dự án">
              </div>

              <div class="form-group">
                <label class="form-label">Đối tượng thụ hưởng / Người nhận</label>
                <input type="text" class="form-control" name="payee" value="${editCpb?.payee || 'Ban QLDA / Nội bộ'}" placeholder="Cán bộ phụ trách / Nội bộ">
              </div>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Dự toán / Ngân sách dự kiến (VNĐ)</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" name="contract_val" value="${this.formatNumberWithDots(editCpb?.contract_val || 0)}" placeholder="0 (Nếu không khống chế ngân sách)">
                <small style="color: var(--text-muted); font-size: 0.72rem;">Để 0 nếu là khoản chi linh hoạt theo thực tế phát sinh</small>
              </div>

              <div class="form-group">
                <label class="form-label">Ngày khởi tạo</label>
                <input type="date" class="form-control" name="signing_date" value="${this.formatDateForInput(editCpb?.signing_date || editCpb?.created_at) || new Date().toISOString().slice(0, 10)}">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú diễn giải</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú về mục đích sử dụng quỹ chi phí ban...">${editCpb?.note || ''}</textarea>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px; border-top: 1px solid var(--border-color); padding-top: 14px;">
              <button type="submit" class="button button-primary" style="background: #d97706; border-color: #b45309; font-weight: 600; padding: 6px 16px;">
                <i class="fa-solid fa-floppy-disk"></i> ${isEdit ? 'Lưu thay đổi' : 'Tạo khoản Chi phí Ban'}
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    this.setupCurrencyInputListeners(modal);

    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Bấm ra ngoài vùng modal-container thì đóng
    this.bindModalOutsideClick(modal, closeSubModal);

    modal.querySelector('#form-cpb').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const cpbId = fd.get('id').trim();
      const scopeWork = fd.get('scope_of_work').trim();
      const contractVal = this.parseNumberFromDots(fd.get('contract_val'));

      const cpbData = {
        id: cpbId,
        project_id: this.currentProjectId,
        group: 'B',
        parent_id: null,
        level: 0,
        contract_no: cpbId,
        contract_type: 'CPB',
        signing_date: fd.get('signing_date') || new Date().toISOString().slice(0, 10),
        end_date: null,
        status: fd.get('status') || 'Active',
        company_a: curPrj.client || 'Chủ đầu tư',
        company_b: fd.get('company_b')?.trim() || 'Ban Quản lý / Nội bộ',
        company_b2: '',
        rep_b2: '',
        scope_of_work: scopeWork,
        exchange_rate: 1.0,
        estimated_val: contractVal,
        contract_val: contractVal,
        rep_a: '',
        rep_b: '',
        category: 'CPB',
        payee: fd.get('payee')?.trim() || 'Ban QLDA / Nội bộ',
        bank_account: '',
        bank_name: '',
        contact_person_b: '',
        created_at: editCpb?.created_at || new Date().toISOString().slice(0, 10),
        display_no: 'CPB',
        sort_order: 9999,
        is_cpb: true,
        note: fd.get('note')?.trim() || '',
      };

      await db.put('contracts', cpbData);

      // Đảm bảo dòng Header trong bảng payments luôn tồn tại
      const headerId = `HDR-${cpbId}`;
      const existingHeader = this.payments.find((p) => p.id === headerId || (p.contract_id === cpbId && p.is_header));
      const headerRow = {
        id: headerId,
        is_header: true,
        project_id: this.currentProjectId,
        parent_id: 'GROUP-B',
        contract_id: cpbId,
        contract_no: cpbId,
        item_name: scopeWork,
        exchange_rate: 1.0,
        contract_val: contractVal,
        actual_volume: existingHeader?.actual_volume || 0,
        accepted_vol: existingHeader?.accepted_vol || 0,
        request_amount: existingHeader?.request_amount || 0,
        paid_amount: existingHeader?.paid_amount || 0,
        remaining: existingHeader?.remaining || 0,
      };
      await db.put('payments', headerRow);

      await this.loadData();
      closeSubModal();
      this.render();
      await this.syncToLocalFile();
      this.showToast(`Đã lưu thành công khoản Chi phí Ban "${scopeWork}" (${cpbId})!`, 'success');
    });
  }

  // ==========================================
  // XÓA KHOẢN CHI PHÍ BAN (CPB)
  // ==========================================
  async deleteCPB(cpbId) {
    const cpb = this.contracts.find((c) => c.id === cpbId);
    if (!cpb) return;

    const realPayments = this.getRealPayments(cpbId);
    const totalPaid = realPayments.reduce((s, p) => s + (p.paid_amount || p.request_amount || 0), 0);

    const title = 'Xác nhận xóa Chi phí Ban?';
    const message = `Bạn có chắc chắn muốn xóa khoản Chi phí Ban <strong>"${cpb.scope_of_work || cpb.id}"</strong> (${cpbId}) không?`;
    const subMessage = realPayments.length > 0
      ? `⚠️ Khoản này đang có ${realPayments.length} đợt chi (${this.formatCurrency(totalPaid)}). Khi xóa, toàn bộ ${realPayments.length} đợt chi này sẽ bị xóa khỏi hệ thống!`
      : '';

    this.showConfirmDialog({
      title,
      message,
      subMessage,
      confirmText: 'Xóa Chi phí Ban',
      cancelText: 'Hủy bỏ',
      confirmType: 'danger',
      onConfirm: async () => {
        // Xóa hợp đồng CPB
        await db.delete('contracts', cpbId);

        // Xóa Header và toàn bộ đợt thanh toán con của CPB này
        const allPayments = await db.getAll('payments');
        for (const p of allPayments) {
          if (p.contract_id === cpbId || p.id === `HDR-${cpbId}`) {
            await db.delete('payments', p.id);
          }
        }

        // Xóa các tài liệu liên quan nếu có
        const allDocs = await db.getAll('documents');
        for (const d of allDocs) {
          if (d.contract_id === cpbId) {
            await db.delete('documents', d.id);
          }
        }

        await this.loadData();
        this.render();
        await this.syncToLocalFile();
        this.showToast(`Đã xóa thành công khoản Chi phí Ban "${cpb.scope_of_work || cpbId}"!`, 'success');
      },
    });
  }

  // ==========================================
  // XÓA HỢP ĐỒNG (KÈM PHỤ LỤC, ĐỢT THANH TOÁN & HỒ SƠ LIÊN QUAN)
  // ==========================================
  getAllDescendantContractIds(parentId) {
    const result = [];
    const children = (this.contracts || []).filter((c) => c.parent_id === parentId);
    for (const child of children) {
      result.push(child.id);
      result.push(...this.getAllDescendantContractIds(child.id));
    }
    return result;
  }

  async deleteContract(contractId) {
    const contract = (this.contracts || []).find((c) => c.id === contractId);
    if (!contract) return;

    const descendantIds = this.getAllDescendantContractIds(contractId);
    const allIds = [contractId, ...descendantIds];

    const delPayments = (this.payments || []).filter((p) => allIds.includes(p.contract_id));
    const delDocs = (this.documents || []).filter((d) => allIds.includes(d.contract_id));

    const okDelete = await this.confirmAsync({
      title: 'Xác nhận xóa Hợp đồng',
      message: `Bạn có chắc chắn muốn xóa Hợp đồng "${contract.contract_no || contract.id}" không?`,
      subMessage:
        descendantIds.length > 0 || delPayments.length > 0 || delDocs.length > 0
          ? `Sẽ xóa theo: ${descendantIds.length} Phụ lục/Gói con, ${delPayments.length} đợt thanh toán và ${delDocs.length} hồ sơ liên quan. Hành động này không thể hoàn tác!`
          : 'Dữ liệu sẽ được tự động đồng bộ sang file CPMS. Hành động này không thể hoàn tác!',
      confirmText: 'Xóa hợp đồng',
      cancelText: 'Hủy bỏ',
      confirmType: 'danger',
    });
    if (!okDelete) return;

    try {
      // Gỡ liên kết ở các hợp đồng khác đang trỏ tới hợp đồng bị xóa
      for (const other of (this.contracts || [])) {
        if (!allIds.includes(other.id) && allIds.includes(other.linked_contract_id)) {
          other.linked_contract_id = '';
          await db.put('contracts', other);
        }
      }

      for (const cid of allIds) {
        await db.delete('contracts', cid);
      }
      for (const p of delPayments) {
        await db.delete('payments', p.id);
      }
      for (const d of delDocs) {
        await db.delete('documents', d.id);
      }

      await this.loadData();
      await this.syncToLocalFile();

      // Đóng modal chi tiết nếu đang mở
      document.getElementById('contract-detail-modal')?.remove();
      document.querySelectorAll('.modal-overlay').forEach((m) => m.remove());

      const activeCpmsTab = window.app?.currentCpmsTab || 'contracts';
      this.render();
      if (window.app) {
        await window.app.updateDashboardStats();
      }

      // Giữ nguyên giao diện ở tab hiện tại (không bị nhảy về Tổng quan)
      if (window.app && window.app.switchCpmsTab) {
        window.app.switchCpmsTab(activeCpmsTab);
      } else {
        this.updateEmptyStateUI(activeCpmsTab);
      }

      alert(`Đã xóa thành công Hợp đồng "${contract.contract_no || contract.id}"!`);
    } catch (err) {
      console.error('Lỗi khi xóa hợp đồng:', err);
      alert('Không thể xóa hợp đồng: ' + (err.message || err));
    }
  }

  // ==========================================
  // FORM NHẬP / SỬA HỢP ĐỒNG HOẶC PHỤ LỤC (CHUẨN CPMS PLUS)
  // ==========================================
  openContractFormModal(editContract = null, parentContract = null) {
    const isEdit = !!editContract;
    const isAddingAppendix = !isEdit && !!parentContract;
    const targetParent = isEdit ? (editContract.parent_id ? this.contracts.find((c) => c.id === editContract.parent_id) : null) : parentContract;
    const isAppendix = (isEdit && !!editContract.parent_id) || isAddingAppendix;
    const isRootLevel = !isAddingAppendix && (!editContract || !editContract.parent_id || editContract.level === 0);

    let currentSelectGroup = editContract
      ? editContract.group
      : parentContract
      ? parentContract.group
      : this.currentGroup;

    // Sinh mã tự động chuẩn CPMS phân cấp đa tầng (disabled không cho sửa tay)
    let formId = '';
    if (isEdit) {
      formId = editContract.id;
    } else if (isAddingAppendix) {
      formId = CPMSImporterExporter.generateNextContractID(parentContract, this.contracts, parentContract.group);
    } else {
      formId = CPMSImporterExporter.generateNextContractID(null, this.contracts, currentSelectGroup);
    }

    // Giá trị ban đầu (nếu thêm phụ lục thì kế thừa từ hợp đồng cha)
    const baseSource = editContract || parentContract || {};
    const defaultContractType = isAddingAppendix
      ? (parentContract.parent_id ? 'Sub-Package' : 'Appendix')
      : isEdit
      ? (editContract.contract_type || (isAppendix ? 'Appendix' : 'MAIN'))
      : 'MAIN';

    const isAdditive = editContract
      ? (editContract.is_additive === true || editContract.is_additive === 1 || editContract.is_additive === '1')
      : false; // Mặc định phụ lục mới là Không cộng dồn

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || {};

    // Sửa giá trị mặc định Bên A: Không để mặc định là CLI
    let defaultCompanyA = editContract?.company_a || baseSource.company_a || '';
    if (!defaultCompanyA) {
      if (curPrj.client && curPrj.client.trim().toUpperCase() !== 'CLI' && curPrj.client.trim() !== 'Chủ đầu tư') {
        defaultCompanyA = curPrj.client.trim();
      } else {
        const foundA = (this.contracts || []).find((c) => c.group === 'A' && c.company_a && c.company_a.trim().toUpperCase() !== 'CLI');
        if (foundA) defaultCompanyA = foundA.company_a.trim();
      }
    }

    let defaultCompanyB = editContract?.company_b || baseSource.company_b || '';
    let defaultCompanyB2 = editContract?.company_b2 || baseSource.company_b2 || '';

    // Tự động suy luận Mã Code nếu chưa có
    let defaultCodeA = editContract?.code_a || baseSource.code_a || '';
    if (!defaultCodeA && defaultCompanyA) {
      defaultCodeA = this.getCompanyCodeByName(defaultCompanyA, 'A');
    }

    let defaultCodeB = editContract?.code_b || editContract?.company_code || baseSource.code_b || baseSource.company_code || '';
    if (!defaultCodeB && defaultCompanyB) {
      defaultCodeB = this.getCompanyCodeByName(defaultCompanyB, 'B');
    }

    let defaultCodeB2 = editContract?.code_b2 || baseSource.code_b2 || '';
    if (!defaultCodeB2 && defaultCompanyB2) {
      defaultCodeB2 = this.getCompanyCodeByName(defaultCompanyB2, 'B');
    }

    const currentLinkedId = editContract?.linked_contract_id || baseSource.linked_contract_id || '';

    // Thu thập danh sách gợi ý dữ liệu đã có trong dự án
    const getUnique = (prop) => {
      const set = new Set();
      (this.contracts || []).forEach((c) => {
        const val = c[prop];
        if (val && String(val).trim()) set.add(String(val).trim());
      });
      return Array.from(set);
    };

    const companiesA = getUnique('company_a');
    const codesA = Array.from(new Set([...getUnique('code_a'), 'CLI', 'VEC', 'CHUDT', 'CDT', 'PMU'].filter(Boolean)));
    const taxCodesA = getUnique('tax_code_a');
    const repsA = getUnique('rep_a');
    const contactsA = getUnique('contact_person_a');

    const companiesB = getUnique('company_b');
    const codesB = Array.from(new Set([...getUnique('code_b'), ...getUnique('company_code'), 'CTC', 'VNCN', 'NSH', 'KD', 'TID', 'HHG', 'HD', 'EC-KD', 'EC-PHC', 'EC-TG', 'EC-NSH'].filter(Boolean)));
    const repsB = getUnique('rep_b');
    const taxCodesB1 = getUnique('tax_code');
    const contactsB = getUnique('contact_person_b');

    const companiesB2 = getUnique('company_b2');
    const codesB2 = Array.from(new Set([...getUnique('code_b2'), 'CTC', 'VNCN', 'NSH', 'KD', 'TID', 'HHG', 'EC-KD'].filter(Boolean)));
    const repsB2 = getUnique('rep_b2');
    const taxCodesB2 = getUnique('tax_code_b2');
    const contactsB2 = getUnique('contact_person_b2');

    const payees = getUnique('payee');
    const bankAccounts = getUnique('bank_account');
    const bankNames = getUnique('bank_name');

    // Danh sách hợp đồng mẫu & HĐ liên kết
    const sampleContracts = (this.contracts || []).filter((c) => !c.is_cpb && c.contract_type !== 'CPB');
    const contractsA = (this.contracts || []).filter((c) => c.group === 'A' && !c.is_cpb);
    const contractsB = (this.contracts || []).filter((c) => c.group === 'B' && !c.is_cpb);

    let formTitle = 'Thêm Hợp đồng mới';
    if (isEdit) {
      formTitle = isAppendix ? `Chỉnh sửa Phụ lục: ${editContract.contract_no || editContract.id}` : `Chỉnh sửa Hợp đồng: ${editContract.contract_no || editContract.id}`;
    } else if (isAddingAppendix) {
      formTitle = `Thêm Phụ lục cho Hợp đồng: ${parentContract.contract_no || parentContract.id}`;
    }

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 860px; max-height: 92vh; display: flex; flex-direction: column;">
        <div class="modal-header" style="flex-shrink: 0; border-bottom: 1px solid var(--border-color); padding: 14px 20px;">
          <div>
            <div style="font-size: 0.8rem; color: var(--primary); font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">
              <i class="fa-solid fa-folder-open"></i> ${curPrj.name || 'Dự án CPMS'}
            </div>
            <h3 style="font-size: 1.15rem; font-weight: 800; margin-top: 2px;">
              ${formTitle}
            </h3>
          </div>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <div class="modal-body" style="overflow-y: auto; overflow-x: hidden; padding: 20px; box-sizing: border-box; width: 100%;">
          <form id="form-contract" style="width: 100%; box-sizing: border-box;">
            ${
              !isEdit && sampleContracts.length > 0
                ? `
                <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: var(--radius-sm); padding: 10px 14px; margin-bottom: 16px;">
                  <label class="form-label" style="font-size: 0.82rem; font-weight: 700; color: #1e40af; margin-bottom: 4px; display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-copy"></i> Sao chép từ Hợp đồng mẫu đã có (Tự động điền nhanh):
                  </label>
                  <select id="contract-template-picker" class="form-control" style="font-size: 0.85rem; background: #fff;">
                    <option value="">-- Chọn một hợp đồng trong dự án để sao chép thông tin --</option>
                    ${sampleContracts.map((sc) => `<option value="${sc.id}">[Nhóm ${sc.group}] ${sc.id} — ${sc.contract_no} (${sc.company_b || sc.company_a || 'Chưa rõ đối tác'})</option>`).join('')}
                  </select>
                  <small style="color: #64748b; font-size: 0.73rem;">Khi chọn, hệ thống tự động điền: Bên A, Bên B, Mã Code, Mã số thuế, Người đại diện, Tài khoản ngân hàng... Bạn chỉ cần nhập số HĐ mới, ngày ký và giá trị.</small>
                </div>
              `
                : ''
            }

            <!-- 1. THÔNG TIN HỢP ĐỒNG & PHÂN CẤP -->
            <div style="font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-file-invoice"></i> Thông tin Hợp đồng & Phân cấp
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label" style="display: flex; justify-content: space-between;">
                  <span>Mã hợp đồng *</span>
                  <span style="font-weight: normal; color: var(--danger); font-size: 0.75rem;"><i class="fa-solid fa-lock"></i> Tự động nhảy</span>
                </label>
                <input type="text" class="form-control" id="contract-input-id" name="id" value="${formId}" readonly style="background: #f1f5f9; color: #334155; font-weight: 700; font-family: monospace; cursor: not-allowed;" required>
                <small style="color: var(--text-muted); font-size: 0.72rem;">Chuẩn CPMS: Root (A-2026-01) ➔ Cấp 1 (A-2026-01-08) ➔ Cấp 2 (A-2026-01-08-01)</small>
              </div>

              <div class="form-group">
                <label class="form-label">Phân nhóm *</label>
                <select class="form-control" id="contract-select-group" name="group" ${isAppendix || isEdit ? 'disabled' : ''} style="${isAppendix || isEdit ? 'background: #f1f5f9; cursor: not-allowed;' : ''}">
                  <option value="A" ${currentSelectGroup === 'A' ? 'selected' : ''}>Nhóm A: Chủ đầu tư / Tổng thầu (Dòng tiền thu)</option>
                  <option value="B" ${currentSelectGroup === 'B' ? 'selected' : ''}>Nhóm B: Nhà thầu phụ / Nhà cung cấp (Dòng tiền chi)</option>
                </select>
                ${isAppendix || isEdit ? `<input type="hidden" name="group" value="${currentSelectGroup}">` : ''}
              </div>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Số hiệu hợp đồng / Phụ lục *</label>
                <input type="text" class="form-control" name="contract_no" id="input-contract-no" value="${editContract?.contract_no || (isAddingAppendix ? `APP-${String((this.getDirectChildren(parentContract.id).length || 0) + 1).padStart(2, '0')}` : '')}" placeholder="Vd: 114/2023/HĐXD/VINAE&C-CTC hoặc APP-08" required>
              </div>

              <div class="form-group">
                <label class="form-label">Loại hợp đồng *</label>
                <select class="form-control" name="contract_type" id="select-contract-type">
                  ${
                    isRootLevel
                      ? `<option value="MAIN" selected>Hợp đồng chính (MAIN)</option>`
                      : `
                        <option value="Appendix" ${defaultContractType === 'Appendix' ? 'selected' : ''}>Phụ lục hợp đồng (Appendix)</option>
                        <option value="Sub-Package" ${defaultContractType === 'Sub-Package' ? 'selected' : ''}>Gói thầu con (Sub-Package)</option>
                      `
                  }
                </select>
              </div>
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">Ngày ký *</label>
                <input type="date" class="form-control" name="signing_date" id="input-signing-date" value="${this.formatDateForInput(editContract?.signing_date) || new Date().toISOString().slice(0, 10)}" required>
              </div>

              <div class="form-group">
                <label class="form-label">Ngày kết thúc</label>
                <input type="date" class="form-control" name="end_date" id="input-end-date" value="${this.formatDateForInput(editContract?.end_date)}">
              </div>

              <div class="form-group">
                <label class="form-label">Tình trạng</label>
                <select class="form-control" name="status" id="select-status">
                  <option value="Active" ${editContract?.status === 'Active' || !editContract ? 'selected' : ''}>Đang thực hiện (Active)</option>
                  <option value="Completed" ${this.isContractCompleted(editContract) ? 'selected' : ''}>Đã quyết toán (Completed)</option>
                  <option value="Suspended" ${editContract?.status === 'Suspended' ? 'selected' : ''}>Tạm ngưng (Suspended)</option>
                  <option value="Cancel" ${this.isContractCanceled(editContract) ? 'selected' : ''}>Đã hủy (Cancel)</option>
                </select>
              </div>
            </div>

            <!-- HỢP ĐỒNG LIÊN KẾT TƯƠNG TỰ (TÌM KIẾM & ĐỒNG BỘ 2 CHIỀU) -->
            <div class="form-group" style="background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: var(--radius-sm); padding: 12px; margin-top: 4px; position: relative;">
              <input type="hidden" name="linked_contract_id" id="input-linked-contract-id" value="${currentLinkedId}">
              <label class="form-label" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <span style="font-weight: 700; color: #4338ca;"><i class="fa-solid fa-link" style="color: #6366f1;"></i> Hợp đồng liên kết tương tự</span>
                <small style="color: var(--text-muted); font-size: 0.75rem;">Liên kết 2 chiều giữa HĐ Bên A (Chủ đầu tư) và HĐ Bên B (Thầu phụ) để chuyển đổi nhanh</small>
              </label>

              <!-- Box hiển thị trạng thái hợp đồng đã liên kết -->
              <div id="linked-contract-display-container"></div>

              <!-- Dropdown Tìm kiếm Hợp đồng Liên kết -->
              <div id="linked-contract-search-panel" class="hidden" style="margin-top: 8px; background: #ffffff; border: 1px solid #c7d2fe; border-radius: 6px; box-shadow: 0 4px 14px rgba(0,0,0,0.08); padding: 10px;">
                <div style="display: flex; gap: 8px; margin-bottom: 8px;">
                  <div style="position: relative; flex: 1;">
                    <i class="fa-solid fa-magnifying-glass" style="position: absolute; left: 10px; top: 10px; color: #94a3b8; font-size: 0.85rem;"></i>
                    <input type="text" id="input-search-linked-term" class="form-control" style="padding-left: 32px; font-size: 0.85rem;" placeholder="Tìm theo số hiệu HĐ, đối tác, nội dung gói thầu (gõ không dấu)...">
                  </div>
                  <button type="button" class="button button-outline" id="btn-close-linked-search" style="padding: 4px 10px; font-size: 0.8rem;">Đóng</button>
                </div>
                <div id="linked-contracts-list" style="max-height: 220px; overflow-y: auto; display: flex; flex-direction: column; gap: 4px;">
                  <!-- Render động các lựa chọn hợp đồng -->
                </div>
              </div>
            </div>

            <!-- 2. ĐỐI TÁC & THÔNG TIN GÓI THẦU -->
            <div style="font-weight: 700; font-size: 0.95rem; color: #4f46e5; margin: 18px 0 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-handshake"></i> Đối tác & Thông tin gói thầu
            </div>

            <!-- KHỐI BÊN A (CHỦ ĐẦU TƯ / TỔNG THẦU) -->
            <div class="partner-card-block">
              <div style="font-weight: 700; font-size: 0.88rem; color: #1e40af; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
                <i class="fa-solid fa-building" style="color: #2563eb;"></i> Bên A (Chủ đầu tư / Tổng thầu)
              </div>

              <!-- Hàng 1: Tên Bên A & Mã Code A -->
              <div class="form-grid-pair">
                <div class="form-group">
                  <label class="form-label">Tên công ty Bên A</label>
                  <div class="custom-autocomplete-wrap">
                    <input type="text" class="form-control" name="company_a" id="input-company-a" autocomplete="off" value="${defaultCompanyA}" placeholder="Vd: Công ty Cổ phần Xây dựng và Kỹ thuật Việt Nam">
                    <div class="autocomplete-dropdown hidden"></div>
                  </div>
                </div>

                <div class="form-group">
                  <label class="form-label">Mã Code A</label>
                  <div class="custom-autocomplete-wrap">
                    <input type="text" class="form-control" name="code_a" id="input-code-a" autocomplete="off" value="${defaultCodeA}" placeholder="Vd: CLI, VEC" style="text-transform: uppercase; font-family: monospace; font-weight: 600;">
                    <div class="autocomplete-dropdown hidden"></div>
                  </div>
                </div>
              </div>

              <!-- Hàng 2: Mã số thuế Bên A, Người đại diện Bên A, Người liên hệ Bên A -->
              <div class="form-grid-3">
                <div class="form-group">
                  <label class="form-label"><i class="fa-solid fa-barcode" style="color: #059669;"></i> Mã số thuế Bên A</label>
                  <div class="custom-autocomplete-wrap">
                    <input type="text" class="form-control" name="tax_code_a" id="input-tax-code-a" autocomplete="off" value="${editContract?.tax_code_a || baseSource.tax_code_a || ''}" placeholder="Vd: 0101234567">
                    <div class="autocomplete-dropdown hidden"></div>
                  </div>
                </div>

                <div class="form-group">
                  <label class="form-label">Người đại diện Bên A</label>
                  <div class="custom-autocomplete-wrap">
                    <input type="text" class="form-control" name="rep_a" id="input-rep-a" autocomplete="off" value="${editContract?.rep_a || baseSource.rep_a || ''}" placeholder="Họ tên người đại diện pháp luật">
                    <div class="autocomplete-dropdown hidden"></div>
                  </div>
                </div>

                <div class="form-group">
                  <label class="form-label">Người liên hệ Bên A</label>
                  <div class="custom-autocomplete-wrap">
                    <input type="text" class="form-control" name="contact_person_a" id="input-contact-a" autocomplete="off" value="${editContract?.contact_person_a || baseSource.contact_person_a || ''}" placeholder="Họ tên & số điện thoại liên hệ">
                    <div class="autocomplete-dropdown hidden"></div>
                  </div>
                </div>
              </div>
            </div>

            <!-- KHỐI BÊN B (NHÀ THẦU PHỤ / NHÀ CUNG CẤP) -->
            <div class="partner-card-block">
              <div style="font-weight: 700; font-size: 0.88rem; color: #334155; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
                <i class="fa-solid fa-users-gear" style="color: #6366f1;"></i> Liên danh Nhà thầu Bên B (Nhà thầu phụ / Nhà cung cấp)
              </div>
              
              <!-- NHÀ THẦU B1 (CHÍNH) -->
              <div style="border-bottom: 1px dashed #cbd5e1; padding-bottom: 12px; margin-bottom: 12px;">
                <div class="form-grid-pair">
                  <div class="form-group">
                    <label class="form-label">Nhà thầu B1 (Chính) *</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="company_b" id="input-company-b" autocomplete="off" value="${defaultCompanyB}" placeholder="Tên công ty đứng đầu liên danh / Nhà thầu chính" required>
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Mã Code B1 *</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="code_b" id="input-code-b" autocomplete="off" value="${defaultCodeB}" placeholder="Vd: CTC, VNCN" style="text-transform: uppercase; font-family: monospace; font-weight: 600;">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>
                </div>

                <div class="form-grid-3">
                  <div class="form-group">
                    <label class="form-label"><i class="fa-solid fa-barcode" style="color: #059669;"></i> Mã số thuế B1</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="tax_code" id="input-tax-code-b" autocomplete="off" value="${editContract?.tax_code || baseSource.tax_code || ''}" placeholder="Vd: 0101234567">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Người đại diện B1</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="rep_b" id="input-rep-b" autocomplete="off" value="${editContract?.rep_b || baseSource.rep_b || ''}" placeholder="Họ tên đại diện Bên B1">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Người liên hệ B1</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="contact_person_b" id="input-contact-b" autocomplete="off" value="${editContract?.contact_person_b || baseSource.contact_person_b || ''}" placeholder="Họ tên & số ĐT liên hệ">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>
                </div>
              </div>

              <!-- NHÀ THẦU B2 (LIÊN DANH) -->
              <div>
                <div class="form-grid-pair">
                  <div class="form-group">
                    <label class="form-label">Nhà thầu B2 (Liên danh nếu có)</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="company_b2" id="input-company-b2" autocomplete="off" value="${defaultCompanyB2}" placeholder="Tên công ty liên danh B2 (để trống nếu không có)">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Mã Code B2</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="code_b2" id="input-code-b2" autocomplete="off" value="${defaultCodeB2}" placeholder="Vd: EC-KD" style="text-transform: uppercase; font-family: monospace; font-weight: 600;">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>
                </div>

                <div class="form-grid-3">
                  <div class="form-group">
                    <label class="form-label"><i class="fa-solid fa-barcode" style="color: #059669;"></i> Mã số thuế B2</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="tax_code_b2" id="input-tax-code-b2" autocomplete="off" value="${editContract?.tax_code_b2 || baseSource.tax_code_b2 || ''}" placeholder="Vd: 0108765432">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Người đại diện B2</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="rep_b2" id="input-rep-b2" autocomplete="off" value="${editContract?.rep_b2 || baseSource.rep_b2 || ''}" placeholder="Họ tên đại diện Bên B2">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>

                  <div class="form-group">
                    <label class="form-label">Người liên hệ B2</label>
                    <div class="custom-autocomplete-wrap">
                      <input type="text" class="form-control" name="contact_person_b2" id="input-contact-b2" autocomplete="off" value="${editContract?.contact_person_b2 || baseSource.contact_person_b2 || ''}" placeholder="Họ tên & số ĐT liên hệ">
                      <div class="autocomplete-dropdown hidden"></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <!-- NỘI DUNG / PHẠM VI CÔNG VIỆC -->
            <div class="form-group">
              <label class="form-label">Nội dung / Phạm vi công việc</label>
              <textarea class="form-control" name="scope_of_work" id="input-scope-of-work" rows="2" placeholder="Tên hạng mục thi công, mua sắm thiết bị, tư vấn...">${editContract?.scope_of_work || (isAddingAppendix ? `Phụ lục bổ sung khối lượng cho ${parentContract.contract_no}` : '')}</textarea>
            </div>

            <!-- KHỐI GIÁ TRỊ HỢP ĐỒNG, TỶ GIÁ & PHÂN LOẠI -->
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: var(--radius-sm); padding: 14px; margin-bottom: 14px;">
              <div style="font-weight: 700; font-size: 0.88rem; color: #1e40af; margin-bottom: 10px; display: flex; align-items: center; justify-content: space-between;">
                <span><i class="fa-solid fa-coins" style="color: #eab308;"></i> Giá trị hợp đồng & Tỷ giá quy đổi</span>
                <span id="converted-vnd-badge" style="font-size: 0.78rem; font-weight: 600; color: #15803d; background: #dcfce7; padding: 2px 10px; border-radius: 12px; border: 1px solid #bbf7d0;">
                  Quy đổi: <strong id="converted-vnd-text">0 đ</strong>
                </span>
              </div>

              <div class="form-grid-3">
                <div class="form-group">
                  <label class="form-label">Giá trị ký kết (Nguyên tệ) *</label>
                  <input type="text" inputmode="numeric" class="form-control currency-input" name="contract_val" id="input-contract-val" value="${this.formatNumberWithDots(editContract?.contract_val || 0)}" required style="font-weight: 700; color: var(--primary);">
                  <small style="color: #64748b; font-size: 0.72rem;">Số tiền nguyên tệ trong HĐ (VND, USD...)</small>
                </div>

                <div class="form-group">
                  <label class="form-label">Tỷ giá (Exchange Rate)</label>
                  <input type="number" step="any" class="form-control" name="exchange_rate" id="input-exchange-rate" value="${editContract?.exchange_rate || baseSource.exchange_rate || 1}" style="font-weight: 600;">
                  <small style="color: #64748b; font-size: 0.72rem;">Mặc định là 1 nếu ký tiền VNĐ</small>
                </div>

                <div class="form-group">
                  <label class="form-label"><i class="fa-solid fa-money-bill-wave" style="color: #16a34a;"></i> Giá trị quy đổi (VNĐ)</label>
                  <input type="text" class="form-control" id="input-converted-vnd" readonly style="background: #f0fdf4; border-color: #86efac; color: #15803d; font-weight: 800; font-family: monospace; cursor: not-allowed;" value="0 đ">
                  <small style="color: #16a34a; font-size: 0.72rem;">= Giá trị ký kết × Tỷ giá</small>
                </div>
              </div>

              <div class="form-group" style="margin-top: 10px; margin-bottom: 0;">
                <label class="form-label">Phân loại hợp đồng</label>
                <select class="form-control" name="category" id="select-category">
                  <option value="Construction" ${(editContract?.category || baseSource.category) === 'Construction' ? 'selected' : ''}>Thi công / Xây lắp (Construction)</option>
                  <option value="Procurement" ${(editContract?.category || baseSource.category) === 'Procurement' ? 'selected' : ''}>Mua sắm vật tư (Procurement)</option>
                  <option value="Consultancy" ${(editContract?.category || baseSource.category) === 'Consultancy' ? 'selected' : ''}>Tư vấn giám sát / Thẩm tra (Consultancy)</option>
                  <option value="Design" ${(editContract?.category || baseSource.category) === 'Design' ? 'selected' : ''}>Thiết kế / Khảo sát (Design)</option>
                  <option value="Equipment" ${(editContract?.category || baseSource.category) === 'Equipment' ? 'selected' : ''}>Thiết bị MEP / Cơ điện (Equipment)</option>
                  <option value="Service" ${(editContract?.category || baseSource.category) === 'Service' ? 'selected' : ''}>Dịch vụ thí nghiệm / Vệ sinh (Service)</option>
                  <option value="Maintenance" ${(editContract?.category || baseSource.category) === 'Maintenance' ? 'selected' : ''}>Bảo trì / Vận hành (Maintenance)</option>
                  <option value="Rental" ${(editContract?.category || baseSource.category) === 'Rental' ? 'selected' : ''}>Thuê thiết bị / Nhân công (Rental)</option>
                  <option value="Design & Build" ${(editContract?.category || baseSource.category) === 'Design & Build' ? 'selected' : ''}>Tổng thầu EPC / D&B (Design & Build)</option>
                  <option value="Khác" ${(editContract?.category || baseSource.category) === 'Khác' || !(editContract?.category || baseSource.category) ? 'selected' : ''}>Khác / Chưa phân loại</option>
                </select>
              </div>

              ${
                isAppendix
                  ? `
                <!-- TÙY CHỌN TÍNH TOÁN GIÁ TRỊ PHỤ LỤC -->
                <div style="background: #ffffff; border: 1px dashed #cbd5e1; border-radius: var(--radius-sm); padding: 12px 14px; margin-top: 14px;">
                  <label class="form-label" style="font-weight: 700; color: #1e40af; margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-calculator" style="color: #3b82f6;"></i> Tùy chọn cộng dồn vào Giá trị Hợp đồng chính:
                  </label>
                  <div style="display: flex; flex-direction: column; gap: 8px; font-size: 0.85rem;">
                    <label style="display: flex; align-items: flex-start; gap: 8px; cursor: pointer; color: #334155;">
                      <input type="radio" name="is_additive" value="0" ${!isAdditive ? 'checked' : ''} style="margin-top: 3px;">
                      <div>
                        <strong>Không cộng dồn vào Hợp đồng chính</strong> <span class="badge" style="background: #f1f5f9; color: #475569; font-size: 0.7rem;">Mặc định chuẩn CPMS</span>
                        <div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">
                          Giá trị dùng để theo dõi giải ngân riêng của phụ lục này; <u>KHÔNG</u> làm tăng giá trị HĐ chính & Tổng nhóm giải ngân (tránh cộng dồn trùng lặp).
                        </div>
                      </div>
                    </label>
                    <label style="display: flex; align-items: flex-start; gap: 8px; cursor: pointer; color: #334155;">
                      <input type="radio" name="is_additive" value="1" ${isAdditive ? 'checked' : ''} style="margin-top: 3px;">
                      <div>
                        <strong style="color: #059669;"><i class="fa-solid fa-plus-circle"></i> Cộng thêm vào Giá trị Hợp đồng chính (Phụ lục phát sinh / bổ sung)</strong>
                        <div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">
                          Giá trị quy đổi VNĐ của phụ lục này sẽ được cộng thêm vào tổng giá trị của Hợp đồng chính và Tổng nhóm trong Báo cáo Giải ngân.
                        </div>
                      </div>
                    </label>
                  </div>
                </div>
              `
                  : ''
              }

              <!-- TÙY CHỌN GIÁ TRỊ TẠM TÍNH (ÁP DỤNG CHO CẢ HỢP ĐỒNG CHỦ VÀ PHỤ LỤC) -->
              <div style="background: #fffbeb; border: 1px dashed #fcd34d; border-radius: var(--radius-sm); padding: 10px 14px; margin-top: 12px;">
                <label style="display: flex; align-items: flex-start; gap: 8px; cursor: pointer; color: #92400e; font-size: 0.84rem; margin: 0;">
                  <input type="checkbox" name="is_provisional_val" id="chk-provisional-val" value="1" ${editContract?.is_provisional_val === true || editContract?.is_provisional_val === 1 || editContract?.is_provisional_val === '1' ? 'checked' : ''} style="margin-top: 3px; width: 15px; height: 15px; cursor: pointer; accent-color: #d97706;">
                  <div>
                    <strong>${isAppendix ? 'Giá trị phụ lục này là tạm tính' : 'Giá trị hợp đồng là tạm tính'}</strong>
                    <div style="font-size: 0.75rem; color: #b45309; margin-top: 2px;">
                      Khi tích chọn, giá trị ${isAppendix ? 'phụ lục này' : 'hợp đồng'} sẽ <strong>không tính</strong> vào tổng hợp; hệ thống sẽ tự động lấy <strong>tổng giá trị các phụ lục con</strong> trực thuộc ${isAppendix ? 'phụ lục này' : 'hợp đồng này'} để tổng hợp.
                    </div>
                  </div>
                </label>
              </div>
            </div>

            <!-- 3. THÔNG TIN THỤ HƯỞNG & THANH TOÁN -->
            <div style="font-weight: 700; font-size: 0.95rem; color: #0891b2; margin: 18px 0 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-building-columns"></i> Thông tin thụ hưởng & Thanh toán
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">Đơn vị thụ hưởng</label>
                <div class="custom-autocomplete-wrap">
                  <input type="text" class="form-control" name="payee" id="input-payee" autocomplete="off" value="${editContract?.payee || baseSource.payee || editContract?.company_b || baseSource.company_b || ''}" placeholder="Tên đơn vị thụ hưởng thanh toán">
                  <div class="autocomplete-dropdown hidden"></div>
                </div>
              </div>

              <div class="form-group">
                <label class="form-label">Số tài khoản ngân hàng</label>
                <div class="custom-autocomplete-wrap">
                  <input type="text" class="form-control" name="bank_account" id="input-bank-account" autocomplete="off" value="${editContract?.bank_account || baseSource.bank_account || ''}" placeholder="Vd: 1221 0000 7136 52">
                  <div class="autocomplete-dropdown hidden"></div>
                </div>
              </div>

              <div class="form-group">
                <label class="form-label">Ngân hàng & Chi nhánh</label>
                <div class="custom-autocomplete-wrap">
                  <input type="text" class="form-control" name="bank_name" id="input-bank-name" autocomplete="off" value="${editContract?.bank_name || baseSource.bank_name || ''}" placeholder="Vd: BIDV - Chi nhánh Hà Thành">
                  <div class="autocomplete-dropdown hidden"></div>
                </div>
              </div>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; padding-top: 14px; border-top: 1px solid var(--border-color);">
              <button type="submit" class="button button-primary" style="font-weight: 600; padding: 6px 16px;">
                <i class="fa-solid fa-floppy-disk"></i> ${isEdit ? 'Lưu thay đổi Hợp đồng' : (isAddingAppendix ? 'Lưu Phụ lục' : 'Lưu Hợp đồng mới')}
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    this.setupCurrencyInputListeners(modal);

    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Bấm ra ngoài vùng modal-container thì đóng
    this.bindModalOutsideClick(modal, closeSubModal);

    // Khởi tạo Custom Autocomplete Dropdown ngay bên dưới ô nhập
    const setupCustomAutocomplete = (inputEl, items, onSelect = null) => {
      if (!inputEl) return;
      const wrap = inputEl.closest('.custom-autocomplete-wrap');
      if (!wrap) return;
      const dropdown = wrap.querySelector('.autocomplete-dropdown');
      if (!dropdown) return;

      const renderItems = (query = '') => {
        const q = String(query).trim().toLowerCase();
        const filtered = items.filter((it) => {
          if (!it) return false;
          if (!q) return true;
          return String(it).toLowerCase().includes(q);
        });

        if (filtered.length === 0) {
          dropdown.innerHTML = '<div style="padding: 8px 12px; font-size: 0.8rem; color: #94a3b8; text-align: center;">Không có gợi ý phù hợp</div>';
          dropdown.classList.remove('hidden');
          return;
        }

        dropdown.innerHTML = filtered
          .slice(0, 15)
          .map((it) => `<div class="autocomplete-item" data-val="${String(it).replace(/"/g, '&quot;')}">${it}</div>`)
          .join('');

        dropdown.querySelectorAll('.autocomplete-item').forEach((itemEl) => {
          itemEl.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const val = itemEl.getAttribute('data-val');
            inputEl.value = val;
            dropdown.classList.add('hidden');
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            if (onSelect) onSelect(val);
          });
        });

        dropdown.classList.remove('hidden');
      };

      inputEl.addEventListener('focus', () => {
        renderItems(inputEl.value);
      });

      inputEl.addEventListener('input', () => {
        renderItems(inputEl.value);
      });

      inputEl.addEventListener('blur', () => {
        setTimeout(() => {
          dropdown.classList.add('hidden');
        }, 220);
      });
    };

    // Thiết lập autocomplete cho các trường Bên A
    setupCustomAutocomplete(modal.querySelector('#input-company-a'), companiesA, (selectedName) => {
      const matched = (this.contracts || []).find((c) => c.company_a && c.company_a.trim().toLowerCase() === selectedName.trim().toLowerCase());
      const fillIfEmpty = (sel, newVal) => {
        const el = modal.querySelector(sel);
        if (el && (!el.value || el.value.trim() === '') && newVal) el.value = newVal;
      };
      if (matched) {
        fillIfEmpty('#input-code-a', matched.code_a || this.getCompanyCodeByName(selectedName, 'A'));
        fillIfEmpty('#input-tax-code-a', matched.tax_code_a);
        fillIfEmpty('#input-rep-a', matched.rep_a);
        fillIfEmpty('#input-contact-a', matched.contact_person_a);
      } else {
        fillIfEmpty('#input-code-a', this.getCompanyCodeByName(selectedName, 'A'));
      }
    });
    setupCustomAutocomplete(modal.querySelector('#input-code-a'), codesA);
    setupCustomAutocomplete(modal.querySelector('#input-tax-code-a'), taxCodesA);
    setupCustomAutocomplete(modal.querySelector('#input-rep-a'), repsA);
    setupCustomAutocomplete(modal.querySelector('#input-contact-a'), contactsA);

    // Tự động gợi ý mã code khi người dùng gõ tên công ty Bên A
    const inputCompA = modal.querySelector('#input-company-a');
    const inputCodeA = modal.querySelector('#input-code-a');
    inputCompA?.addEventListener('input', (e) => {
      const autoCode = this.getCompanyCodeByName(e.target.value, 'A');
      if (autoCode && inputCodeA && (!inputCodeA.value || inputCodeA.value.trim() === '')) {
        inputCodeA.value = autoCode;
      }
    });

    // Thiết lập autocomplete cho các trường Bên B1
    setupCustomAutocomplete(modal.querySelector('#input-company-b'), companiesB, (selectedName) => {
      const matched = (this.contracts || []).find((c) => c.company_b && c.company_b.trim().toLowerCase() === selectedName.trim().toLowerCase());
      const fillIfEmpty = (sel, newVal) => {
        const el = modal.querySelector(sel);
        if (el && (!el.value || el.value.trim() === '') && newVal) el.value = newVal;
      };
      if (matched) {
        fillIfEmpty('#input-code-b', matched.code_b || matched.company_code || this.getCompanyCodeByName(selectedName, 'B'));
        fillIfEmpty('#input-tax-code-b', matched.tax_code);
        fillIfEmpty('#input-rep-b', matched.rep_b);
        fillIfEmpty('#input-contact-b', matched.contact_person_b);
        fillIfEmpty('#input-payee', matched.payee || matched.company_b);
        fillIfEmpty('#input-bank-account', matched.bank_account);
        fillIfEmpty('#input-bank-name', matched.bank_name);
      } else {
        fillIfEmpty('#input-code-b', this.getCompanyCodeByName(selectedName, 'B'));
      }
    });
    setupCustomAutocomplete(modal.querySelector('#input-code-b'), codesB);
    setupCustomAutocomplete(modal.querySelector('#input-tax-code-b'), taxCodesB1);
    setupCustomAutocomplete(modal.querySelector('#input-rep-b'), repsB);
    setupCustomAutocomplete(modal.querySelector('#input-contact-b'), contactsB);

    // Tự động gợi ý mã code khi người dùng gõ tên công ty Bên B1
    const inputCompB = modal.querySelector('#input-company-b');
    const inputCodeB = modal.querySelector('#input-code-b');
    inputCompB?.addEventListener('input', (e) => {
      const autoCode = this.getCompanyCodeByName(e.target.value, 'B');
      if (autoCode && inputCodeB && (!inputCodeB.value || inputCodeB.value.trim() === '')) {
        inputCodeB.value = autoCode;
      }
    });

    // Thiết lập autocomplete cho các trường Bên B2
    setupCustomAutocomplete(modal.querySelector('#input-company-b2'), companiesB2, (selectedName) => {
      const matched = (this.contracts || []).find((c) => c.company_b2 && c.company_b2.trim().toLowerCase() === selectedName.trim().toLowerCase());
      const fillIfEmpty = (sel, newVal) => {
        const el = modal.querySelector(sel);
        if (el && (!el.value || el.value.trim() === '') && newVal) el.value = newVal;
      };
      if (matched) {
        fillIfEmpty('#input-code-b2', matched.code_b2 || this.getCompanyCodeByName(selectedName, 'B'));
        fillIfEmpty('#input-tax-code-b2', matched.tax_code_b2);
        fillIfEmpty('#input-rep-b2', matched.rep_b2);
        fillIfEmpty('#input-contact-b2', matched.contact_person_b2);
      } else {
        fillIfEmpty('#input-code-b2', this.getCompanyCodeByName(selectedName, 'B'));
      }
    });
    setupCustomAutocomplete(modal.querySelector('#input-code-b2'), codesB2);
    setupCustomAutocomplete(modal.querySelector('#input-tax-code-b2'), taxCodesB2);
    setupCustomAutocomplete(modal.querySelector('#input-rep-b2'), repsB2);
    setupCustomAutocomplete(modal.querySelector('#input-contact-b2'), contactsB2);

    // Tự động gợi ý mã code khi người dùng gõ tên công ty Bên B2
    const inputCompB2 = modal.querySelector('#input-company-b2');
    const inputCodeB2 = modal.querySelector('#input-code-b2');
    inputCompB2?.addEventListener('input', (e) => {
      const autoCode = this.getCompanyCodeByName(e.target.value, 'B');
      if (autoCode && inputCodeB2 && (!inputCodeB2.value || inputCodeB2.value.trim() === '')) {
        inputCodeB2.value = autoCode;
      }
    });

    // Autocomplete thụ hưởng & Ngân hàng
    setupCustomAutocomplete(modal.querySelector('#input-payee'), payees);
    setupCustomAutocomplete(modal.querySelector('#input-bank-account'), bankAccounts);
    setupCustomAutocomplete(modal.querySelector('#input-bank-name'), bankNames);

    // Component Tìm kiếm & Lựa chọn Hợp đồng liên kết tương tự
    const linkedInput = modal.querySelector('#input-linked-contract-id');
    const linkedDisplay = modal.querySelector('#linked-contract-display-container');
    const linkedSearchPanel = modal.querySelector('#linked-contract-search-panel');
    const linkedSearchTerm = modal.querySelector('#input-search-linked-term');
    const linkedList = modal.querySelector('#linked-contracts-list');
    const btnCloseLinked = modal.querySelector('#btn-close-linked-search');

    const candidateContracts = (this.contracts || []).filter(
      (c) => !c.is_cpb && c.contract_type !== 'CPB' && c.id !== (editContract?.id || '')
    );

    const renderLinkedContractDisplay = (selectedId) => {
      if (!linkedDisplay) return;
      if (!selectedId) {
        linkedDisplay.innerHTML = `
          <button type="button" class="button button-outline" id="btn-open-linked-search" style="width: 100%; border-style: dashed; border-color: #cbd5e1; justify-content: center; gap: 8px; padding: 10px; font-size: 0.85rem; color: #4338ca; background: #fff;">
            <i class="fa-solid fa-magnifying-glass"></i> Chọn hợp đồng liên kết tương tự (Tìm kiếm theo số hiệu, đối tác...)
          </button>
        `;
        linkedDisplay.querySelector('#btn-open-linked-search')?.addEventListener('click', () => {
          openLinkedSearch();
        });
        return;
      }

      const found = (this.contracts || []).find((c) => c.id === selectedId);
      if (found) {
        const isGrpA = found.group === 'A';
        linkedDisplay.innerHTML = `
          <div style="display: flex; align-items: center; justify-content: space-between; background: #ffffff; border: 1px solid #818cf8; border-radius: 6px; padding: 8px 12px; gap: 10px;">
            <div style="display: flex; align-items: center; gap: 10px; overflow: hidden;">
              <span class="badge" style="background: ${isGrpA ? '#dbeafe' : '#fef3c7'}; color: ${isGrpA ? '#1e40af' : '#92400e'}; font-weight: 700; flex-shrink: 0;">
                Nhóm ${found.group}
              </span>
              <div style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                <div style="font-weight: 700; color: #1e1b4b; font-size: 0.88rem;">
                  ${found.contract_no || found.id} <span style="font-size: 0.75rem; color: #64748b; font-weight: normal;">(${found.id})</span>
                </div>
                <div style="font-size: 0.75rem; color: #475569; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                  ${found.company_b || found.company_a || 'Chưa rõ đối tác'} ${found.scope_of_work ? `• ${found.scope_of_work}` : ''}
                </div>
              </div>
            </div>
            <div style="display: flex; gap: 6px; flex-shrink: 0;">
              <button type="button" class="button button-outline" id="btn-change-linked" style="padding: 4px 8px; font-size: 0.75rem;">
                <i class="fa-solid fa-magnifying-glass"></i> Đổi
              </button>
              <button type="button" class="button button-outline" id="btn-clear-linked" style="padding: 4px 8px; font-size: 0.75rem; color: #dc2626; border-color: #fca5a5;">
                <i class="fa-solid fa-xmark"></i> Gỡ
              </button>
            </div>
          </div>
        `;
        linkedDisplay.querySelector('#btn-change-linked')?.addEventListener('click', () => {
          openLinkedSearch();
        });
        linkedDisplay.querySelector('#btn-clear-linked')?.addEventListener('click', () => {
          if (linkedInput) linkedInput.value = '';
          renderLinkedContractDisplay('');
        });
      } else {
        linkedDisplay.innerHTML = `
          <div style="display: flex; align-items: center; justify-content: space-between; background: #fff1f2; border: 1px solid #fecdd3; border-radius: 6px; padding: 8px 12px;">
            <div style="font-size: 0.85rem; color: #be123c;">HĐ liên kết: <strong>${selectedId}</strong> (Không tồn tại trong dự án)</div>
            <button type="button" class="button button-outline" id="btn-clear-linked" style="padding: 4px 8px; font-size: 0.75rem; color: #dc2626;">Gỡ</button>
          </div>
        `;
        linkedDisplay.querySelector('#btn-clear-linked')?.addEventListener('click', () => {
          if (linkedInput) linkedInput.value = '';
          renderLinkedContractDisplay('');
        });
      }
    };

    const renderLinkedOptions = (filterTerm = '') => {
      if (!linkedList) return;
      const cleanTerm = this.removeAccents(filterTerm).toLowerCase();

      const matchedContracts = candidateContracts.filter((c) => {
        if (!cleanTerm) return true;
        const searchTarget = this.removeAccents(
          `${c.id} ${c.contract_no || ''} ${c.company_a || ''} ${c.company_b || ''} ${c.scope_of_work || ''} ${c.code_a || ''} ${c.code_b || ''}`
        ).toLowerCase();
        return searchTarget.includes(cleanTerm);
      });

      if (matchedContracts.length === 0) {
        linkedList.innerHTML = `<div style="padding: 12px; font-size: 0.82rem; color: #94a3b8; text-align: center;">Không tìm thấy hợp đồng nào phù hợp</div>`;
        return;
      }

      const candA = matchedContracts.filter((c) => c.group === 'A');
      const candB = matchedContracts.filter((c) => c.group === 'B');

      let html = '';
      if (candA.length > 0) {
        html += `<div style="font-size: 0.75rem; font-weight: 700; color: #1e40af; background: #eff6ff; padding: 4px 8px; border-radius: 4px; margin-top: 4px;">NHÓM A: CHỦ ĐẦU TƯ / TỔNG THẦU (${candA.length})</div>`;
        html += candA.map((c) => `
          <div class="linked-candidate-item" data-id="${c.id}" style="padding: 8px 10px; border-bottom: 1px solid #f1f5f9; cursor: pointer; border-radius: 4px; transition: background 0.15s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
            <div style="display: flex; justify-content: space-between; align-items: center;">
              <span style="font-weight: 700; color: #1e293b; font-size: 0.85rem;">${c.contract_no || c.id}</span>
              <span style="font-size: 0.72rem; color: #64748b; font-family: monospace;">${c.id}</span>
            </div>
            <div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">
              ${c.company_a || c.company_b || ''} ${c.scope_of_work ? `• ${c.scope_of_work}` : ''}
            </div>
          </div>
        `).join('');
      }

      if (candB.length > 0) {
        html += `<div style="font-size: 0.75rem; font-weight: 700; color: #92400e; background: #fef3c7; padding: 4px 8px; border-radius: 4px; margin-top: 6px;">NHÓM B: NHÀ THẦU PHỤ / NHÀ CUNG CẤP (${candB.length})</div>`;
        html += candB.map((c) => `
          <div class="linked-candidate-item" data-id="${c.id}" style="padding: 8px 10px; border-bottom: 1px solid #f1f5f9; cursor: pointer; border-radius: 4px; transition: background 0.15s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
            <div style="display: flex; justify-content: space-between; align-items: center;">
              <span style="font-weight: 700; color: #1e293b; font-size: 0.85rem;">${c.contract_no || c.id}</span>
              <span style="font-size: 0.72rem; color: #64748b; font-family: monospace;">${c.id}</span>
            </div>
            <div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">
              ${c.company_b || ''} ${c.scope_of_work ? `• ${c.scope_of_work}` : ''}
            </div>
          </div>
        `).join('');
      }

      linkedList.innerHTML = html;

      linkedList.querySelectorAll('.linked-candidate-item').forEach((item) => {
        item.addEventListener('click', () => {
          const cid = item.getAttribute('data-id');
          if (linkedInput) linkedInput.value = cid;
          renderLinkedContractDisplay(cid);
          if (linkedSearchPanel) linkedSearchPanel.classList.add('hidden');
        });
      });
    };

    const openLinkedSearch = () => {
      if (!linkedSearchPanel) return;
      linkedSearchPanel.classList.remove('hidden');
      if (linkedSearchTerm) {
        linkedSearchTerm.value = '';
        linkedSearchTerm.focus();
      }
      renderLinkedOptions('');
    };

    btnCloseLinked?.addEventListener('click', () => {
      if (linkedSearchPanel) linkedSearchPanel.classList.add('hidden');
    });

    linkedSearchTerm?.addEventListener('input', (e) => {
      renderLinkedOptions(e.target.value);
    });

    // Khởi tạo hiển thị ban đầu cho Hợp đồng liên kết
    renderLinkedContractDisplay(currentLinkedId);

    // Live update Thành tiền quy đổi VNĐ tự động theo thời gian thực
    const inputVal = modal.querySelector('#input-contract-val');
    const inputRate = modal.querySelector('#input-exchange-rate');
    const inputVnd = modal.querySelector('#input-converted-vnd');
    const badgeVndText = modal.querySelector('#converted-vnd-text');

    const updateConvertedVnd = () => {
      const rawVal = this.parseNumberFromDots(inputVal?.value || '0');
      const rate = parseFloat(inputRate?.value) || 1.0;
      const vndVal = Math.round(rawVal * rate);
      const formattedVnd = this.formatCurrency(vndVal);
      if (inputVnd) inputVnd.value = formattedVnd;
      if (badgeVndText) badgeVndText.textContent = formattedVnd;
    };

    if (inputVal && inputRate) {
      inputVal.addEventListener('input', updateConvertedVnd);
      inputRate.addEventListener('input', updateConvertedVnd);
      updateConvertedVnd();
    }

    // Nếu tạo mới Hợp đồng cấp 0 và thay đổi nhóm A/B thì tự sinh lại ContractID tương ứng
    if (!isEdit && !isAddingAppendix) {
      const grpSelect = modal.querySelector('#contract-select-group');
      const idInput = modal.querySelector('#contract-input-id');
      grpSelect?.addEventListener('change', (e) => {
        const nextId = CPMSImporterExporter.generateNextContractID(null, this.contracts, e.target.value);
        idInput.value = nextId;
      });
    }

    // Tự động điền dữ liệu khi chọn Hợp đồng mẫu
    const tplPicker = modal.querySelector('#contract-template-picker');
    tplPicker?.addEventListener('change', (e) => {
      const selectedId = e.target.value;
      if (!selectedId) return;
      const tpl = (this.contracts || []).find((c) => c.id === selectedId);
      if (!tpl) return;

      const setVal = (sel, val) => {
        const el = modal.querySelector(sel);
        if (el && val !== undefined && val !== null) el.value = val;
      };

      setVal('#input-company-a', tpl.company_a || '');
      setVal('#input-code-a', tpl.code_a || this.getCompanyCodeByName(tpl.company_a, 'A'));
      setVal('#input-tax-code-a', tpl.tax_code_a || '');
      setVal('#input-rep-a', tpl.rep_a || '');
      setVal('#input-contact-a', tpl.contact_person_a || '');
      setVal('#input-company-b', tpl.company_b || '');
      setVal('#input-code-b', tpl.code_b || tpl.company_code || this.getCompanyCodeByName(tpl.company_b, 'B'));
      setVal('#input-tax-code-b', tpl.tax_code || '');
      setVal('#input-rep-b', tpl.rep_b || '');
      setVal('#input-contact-b', tpl.contact_person_b || '');
      setVal('#input-company-b2', tpl.company_b2 || '');
      setVal('#input-code-b2', tpl.code_b2 || this.getCompanyCodeByName(tpl.company_b2, 'B'));
      setVal('#input-tax-code-b2', tpl.tax_code_b2 || '');
      setVal('#input-rep-b2', tpl.rep_b2 || '');
      setVal('#input-contact-b2', tpl.contact_person_b2 || '');
      setVal('#input-scope-of-work', tpl.scope_of_work || '');
      setVal('#select-category', tpl.category || 'Construction');
      setVal('#input-payee', tpl.payee || tpl.company_b || '');
      setVal('#input-bank-account', tpl.bank_account || '');
      setVal('#input-bank-name', tpl.bank_name || '');
      setVal('#input-contract-val', this.formatNumberWithDots(tpl.contract_val || 0));
      setVal('#input-exchange-rate', tpl.exchange_rate || 1);
      if (tpl.linked_contract_id) {
        if (linkedInput) linkedInput.value = tpl.linked_contract_id;
        renderLinkedContractDisplay(tpl.linked_contract_id);
      }
      updateConvertedVnd();
    });

    modal.querySelector('#form-contract').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);

      const targetGroupId = isAddingAppendix
        ? parentContract.group
        : isEdit
        ? editContract.group
        : formData.get('group') || currentSelectGroup;

      const targetParentId = isAddingAppendix
        ? parentContract.id
        : isEdit
        ? editContract.parent_id || null
        : null;

      const targetLevel = isAddingAppendix
        ? (parentContract.level || 0) + 1
        : isEdit
        ? editContract.level || 0
        : 0;

      const compA = (formData.get('company_a') || '').trim();
      const compB = (formData.get('company_b') || '').trim();
      const compB2 = (formData.get('company_b2') || '').trim();

      const codeA = (formData.get('code_a') || '').trim().toUpperCase() || this.getCompanyCodeByName(compA, 'A');
      const codeB = (formData.get('code_b') || '').trim().toUpperCase() || this.getCompanyCodeByName(compB, 'B');
      const codeB2 = (formData.get('code_b2') || '').trim().toUpperCase() || this.getCompanyCodeByName(compB2, 'B');

      const oldLinkedId = editContract ? (editContract.linked_contract_id || '') : '';
      const newLinkedId = (formData.get('linked_contract_id') || '').trim();

      const updatedData = {
        ...(editContract || {}),
        id: formData.get('id'),
        project_id: (editContract && editContract.project_id) || (parentContract && parentContract.project_id) || this.currentProjectId,
        group: targetGroupId,
        level: targetLevel,
        parent_id: targetParentId,
        contract_no: formData.get('contract_no'),
        contract_type: formData.get('contract_type') || (targetLevel === 0 ? 'MAIN' : 'Appendix'),
        signing_date: formData.get('signing_date') || null,
        end_date: formData.get('end_date') || null,
        status: formData.get('status') || 'Active',
        company_a: compA,
        code_a: codeA,
        tax_code_a: (formData.get('tax_code_a') || '').trim(),
        rep_a: (formData.get('rep_a') || '').trim(),
        contact_person_a: (formData.get('contact_person_a') || '').trim(),
        company_b: compB,
        code_b: codeB,
        tax_code: (formData.get('tax_code') || '').trim(),
        rep_b: (formData.get('rep_b') || '').trim(),
        contact_person_b: (formData.get('contact_person_b') || '').trim(),
        company_b2: compB2,
        code_b2: codeB2,
        tax_code_b2: (formData.get('tax_code_b2') || '').trim(),
        rep_b2: (formData.get('rep_b2') || '').trim(),
        contact_person_b2: (formData.get('contact_person_b2') || '').trim(),
        company_code: codeB || codeA || '',
        scope_of_work: (formData.get('scope_of_work') || '').trim(),
        contract_val: this.parseNumberFromDots(formData.get('contract_val')),
        exchange_rate: parseFloat(formData.get('exchange_rate')) || 1.0,
        is_additive: isAppendix ? (formData.get('is_additive') === '1') : true,
        is_provisional_val: formData.get('is_provisional_val') === '1',
        category: formData.get('category') || 'Construction',
        payee: (formData.get('payee') || compB || '').trim(),
        bank_account: (formData.get('bank_account') || '').trim(),
        bank_name: (formData.get('bank_name') || '').trim(),
        linked_contract_id: newLinkedId,
      };

      // 1. Lưu hợp đồng hiện tại
      await db.put('contracts', updatedData);

      // 2. Đồng bộ 2 chiều (Bidirectional link sync)
      if (oldLinkedId !== newLinkedId) {
        // Gỡ liên kết ở hợp đồng cũ (nếu có và hợp đồng cũ đang trỏ về HĐ này)
        if (oldLinkedId) {
          const oldLinkedContract = (this.contracts || []).find((c) => c.id === oldLinkedId);
          if (oldLinkedContract && oldLinkedContract.linked_contract_id === updatedData.id) {
            oldLinkedContract.linked_contract_id = '';
            await db.put('contracts', oldLinkedContract);
          }
        }
        // Thiết lập liên kết ở hợp đồng mới
        if (newLinkedId) {
          const newLinkedContract = (this.contracts || []).find((c) => c.id === newLinkedId);
          if (newLinkedContract) {
            // Nếu newLinkedContract đang trỏ tới một hợp đồng khác (khác updatedData.id), gỡ liên kết ở HĐ đó
            if (newLinkedContract.linked_contract_id && newLinkedContract.linked_contract_id !== updatedData.id) {
              const prevTarget = (this.contracts || []).find((c) => c.id === newLinkedContract.linked_contract_id);
              if (prevTarget && prevTarget.linked_contract_id === newLinkedId) {
                prevTarget.linked_contract_id = '';
                await db.put('contracts', prevTarget);
              }
            }
            newLinkedContract.linked_contract_id = updatedData.id;
            await db.put('contracts', newLinkedContract);
          }
        }
      } else if (newLinkedId) {
        // Đảm bảo tính nhất quán nếu đối tác chưa trỏ ngược lại
        const targetContract = (this.contracts || []).find((c) => c.id === newLinkedId);
        if (targetContract && targetContract.linked_contract_id !== updatedData.id) {
          targetContract.linked_contract_id = updatedData.id;
          await db.put('contracts', targetContract);
        }
      }

      await db.put('contracts', updatedData);
      await this.loadData();
      this.render();
      this.updateEmptyStateUI();
      if (window.app) {
        await window.app.updateDashboardStats();
      }
      closeSubModal();
      await this.syncToLocalFile();
      this.showToast('Đã lưu dữ liệu Hợp đồng thành công!', 'success');
      this.openContractDetail(updatedData.id);
    });
  }

  // ==========================================
  // THÊM PHỤ LỤC MỚI (TỰ ĐỘNG CHUYỂN SANG FORM CHUẨN CPMS)
  // ==========================================
  openAppendixFormModal(parentContract) {
    this.openContractFormModal(null, parentContract);
  }

  // ==========================================
  // FORM THÊM / CHỈNH SỬA HỒ SƠ (CHUẨN frmDocumentEdit CPMS PLUS)
  // ==========================================
  openDocumentFormModal(targetContract, editDoc = null) {
    const isEdit = !!editDoc;
    const docs = this.getDocuments(targetContract.id);
    const docId = isEdit
      ? editDoc.id
      : `DOC-${new Date().toISOString().slice(0, 10).replace(/-/g, '')}-${String(docs.length + 1).padStart(3, '0')}`;

    // Thu thập danh sách mã đối tác trong dự án để gợi ý Issuer / Recipient
    const companyCodes = Array.from(new Set([
      targetContract.code_a,
      targetContract.code_b,
      targetContract.code_b2,
      'CLI', 'VEC', 'VNCN', 'CTC', 'KD', 'NSH', 'TID', 'HHG', 'HD', 'EC-KD', 'PMU'
    ].filter(Boolean)));

    // Parse supporting docs nếu là chế độ chỉnh sửa (dạng GTTT:10|THKL:10|NTVL:1|HSK:1|NOTE:Bộ CO CQ gốc x1)
    const supportingObj = {};
    let supportingNote = '';
    if (editDoc && editDoc.supporting_docs) {
      String(editDoc.supporting_docs).split('|').forEach((part) => {
        if (!part) return;
        const colonIdx = part.indexOf(':');
        if (colonIdx > 0) {
          const code = part.slice(0, colonIdx).trim().toUpperCase();
          const val = part.slice(colonIdx + 1).trim();
          if (code === 'NOTE') {
            supportingNote = val;
          } else {
            let normalizedCode = code;
            if (code === 'ASB' || code === 'HC') normalizedCode = 'HC';
            if (code === 'NTTVL' || code === 'NTVL') normalizedCode = 'NTVL';
            if (code === 'HSKHAC' || code === 'OTHER' || code === 'HSK') normalizedCode = 'HSK';
            supportingObj[normalizedCode] = parseInt(val) || 1;
          }
        }
      });
    }

    const renderQtySelect = (code, currentVal) => {
      const val = parseInt(currentVal, 10) || 0;
      let opts = '';
      for (let i = 0; i <= 10; i++) {
        opts += `<option value="${i}" ${val === i ? 'selected' : ''}>${i}</option>`;
      }
      if (val > 10) {
        opts += `<option value="${val}" selected>${val}</option>`;
      }
      const isActive = val > 0;
      return `
        <select class="select-supporting-qty" data-code="${code}" style="width: 52px !important; min-width: 52px !important; max-width: 52px !important; flex: 0 0 52px !important; height: 26px; padding: 2px 4px; font-size: 0.82rem; font-weight: 700; text-align: center; border-radius: 4px; cursor: pointer; outline: none; box-sizing: border-box; background-color: ${isActive ? '#eff6ff' : '#f8fafc'}; border: 1px solid ${isActive ? '#93c5fd' : '#cbd5e1'}; color: ${isActive ? '#1d4ed8' : '#64748b'};">
          ${opts}
        </select>
      `;
    };

    const defaultCategory = editDoc ? (editDoc.category || 'Other') : 'Other';
    const initialProcessedDate = this.formatDateForInput(editDoc?.start_date || editDoc?.received_date) || new Date().toISOString().slice(0, 10);
    const initialCompletedDate = this.formatDateForInput(editDoc?.due_date || editDoc?.return_date) || '';

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 950px; max-height: 94vh; display: flex; flex-direction: column;">
        <!-- Header -->
        <div class="modal-header" style="flex-shrink: 0; border-bottom: 1px solid var(--border-color); padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; background: #faf5ff; flex-wrap: nowrap !important; gap: 14px;">
          <div style="flex: 1; min-width: 0;">
            <div style="display: flex; gap: 8px; align-items: center; flex-wrap: nowrap; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
              <span style="font-size: 0.88rem; font-weight: 700; color: #7e22ce; display: inline-flex; align-items: center; gap: 6px; flex-shrink: 0;">
                <i class="fa-solid fa-file-invoice"></i> ${targetContract.contract_no || targetContract.id}
              </span>
              <span style="color: #cbd5e1; flex-shrink: 0;">|</span>
              <span style="color: #475569; font-weight: 600; font-size: 0.82rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 260px;" title="${targetContract.company_b || targetContract.company_a || '—'}">
                ${targetContract.company_b || targetContract.company_a || '—'}
              </span>
              <span style="color: #cbd5e1; flex-shrink: 0;">|</span>
              <span style="color: #15803d; font-weight: 700; font-size: 0.85rem; flex-shrink: 0;">
                ${this.formatCurrency(targetContract.contract_val || 0)}
              </span>
              ${targetContract.signing_date ? `<span style="color: #cbd5e1; flex-shrink: 0;">|</span><span style="color: #64748b; font-size: 0.78rem; flex-shrink: 0;"><i class="fa-regular fa-calendar-check" style="margin-right: 3px;"></i>${this.formatDate(targetContract.signing_date)}</span>` : ''}
            </div>
            <div style="font-size: 0.78rem; color: #64748b; margin-top: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${targetContract.scope_of_work || ''}">
              ${targetContract.scope_of_work || 'Hồ sơ tài liệu chứng từ theo dõi (Document Tracker)'}
            </div>
          </div>

          <div style="display: flex; gap: 8px; align-items: center; flex-shrink: 0;">
            <button type="button" class="button button-outline" id="btn-doc-edit-open-file" style="padding: 5px 12px; font-size: 0.8rem; color: #0284c7; border-color: #38bdf8; background: #fff; font-weight: 600; display: inline-flex; align-items: center; gap: 6px;">
              <i class="fa-solid fa-file-arrow-up"></i> Mở file
            </button>
            <button type="button" class="button button-outline" id="btn-doc-edit-open-folder" style="padding: 5px 12px; font-size: 0.8rem; color: #d97706; border-color: #fde68a; background: #fff; font-weight: 600; display: inline-flex; align-items: center; gap: 6px;">
              <i class="fa-solid fa-folder-open"></i> Mở thư mục
            </button>
            <button class="button button-outline btn-close-submodal" style="padding: 4px 8px; margin-left: 4px;" title="Đóng"><i class="fa-solid fa-xmark"></i></button>
          </div>
        </div>

        <!-- Body -->
        <div class="modal-body" style="overflow-y: auto; padding: 20px; box-sizing: border-box;">
          <form id="form-doc-edit">
            <input type="hidden" name="id" value="${docId}">
            <input type="hidden" name="contract_id" value="${targetContract.id}">
            <input type="file" id="doc-file-hidden-picker" style="display: none;">

            <div style="display: grid; grid-template-columns: 1.25fr 1fr; gap: 18px; align-items: start;">
              <!-- CỘT TRÁI: THÔNG TIN HỒ SƠ VÀ LƯU TRỮ -->
              <div>
                <!-- 1. Document Type (Phân loại hồ sơ) -->
                <div class="form-group" style="margin-bottom: 14px;">
                  <label class="form-label" style="font-weight: 700; color: #1e1b4b; display: flex; justify-content: space-between;">
                    <span><i class="fa-solid fa-folder"></i> Document (Phân loại hồ sơ)</span>
                    <span style="font-size: 0.75rem; color: #64748b; font-family: monospace;">Mã: ${docId}</span>
                  </label>
                  <div style="display: flex; gap: 16px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: var(--radius-sm); padding: 8px 12px; font-size: 0.85rem;">
                    <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
                      <input type="radio" name="category" value="Contract" ${defaultCategory.includes('Contract') || defaultCategory.includes('Hợp đồng') || defaultCategory === 'CON' ? 'checked' : ''}>
                      <strong>Contract</strong> <span style="font-size: 0.72rem; color: #64748b;">(HĐ / Phụ lục)</span>
                    </label>
                    <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
                      <input type="radio" name="category" value="IPC" ${defaultCategory.includes('IPC') || defaultCategory.includes('thanh toán') ? 'checked' : ''}>
                      <strong>IPC</strong> <span style="font-size: 0.72rem; color: #64748b;">(Thanh toán)</span>
                    </label>
                    <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
                      <input type="radio" name="category" value="Other" ${!defaultCategory.includes('Contract') && !defaultCategory.includes('Hợp đồng') && !defaultCategory.includes('CON') && !defaultCategory.includes('IPC') && !defaultCategory.includes('thanh toán') ? 'checked' : ''}>
                      <strong>Other</strong> <span style="font-size: 0.72rem; color: #64748b;">(Công văn / Khác)</span>
                    </label>
                  </div>

                  <!-- Khối DocType, SubType 1, SubType 2 dành riêng cho Option Other (Chuẩn frmDocumentEdit) -->
                  <div id="other-doctype-container" style="${defaultCategory === 'Other' || (!defaultCategory.includes('Contract') && !defaultCategory.includes('Hợp đồng') && !defaultCategory.includes('CON') && !defaultCategory.includes('IPC') && !defaultCategory.includes('thanh toán')) ? 'display: block;' : 'display: none;'} margin-top: 10px; padding: 10px 12px; background: #f1f5f9; border-radius: var(--radius-sm); border: 1px dashed #cbd5e1;">
                    <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px;">
                      <div>
                        <label style="font-size: 0.75rem; font-weight: 700; color: #334155; margin-bottom: 3px; display: block; white-space: nowrap;">Mã hồ sơ</label>
                        <input type="text" class="form-control" name="other_doctype" id="cboDocType" list="list-other-doctypes" value="${editDoc?.category && editDoc.category !== 'Contract' && editDoc.category !== 'CON' && editDoc.category !== 'IPC' ? editDoc.category : (editDoc?.sub_type1 || 'RFA')}" placeholder="VD: RFA, LTR..." style="text-transform: uppercase; font-family: monospace; font-size: 0.82rem; font-weight: 700; height: 32px; padding: 4px 8px;">
                        <datalist id="list-other-doctypes">
                          <option value="RFA">RFA (Request For Approval)</option>
                          <option value="LTR">LTR (Official Letter / Công văn)</option>
                          <option value="MOM">MOM (Minutes of Meeting)</option>
                          <option value="VO">VO (Variation Order)</option>
                          <option value="APP">APP (Appendix / Phụ lục)</option>
                          <option value="ITP">ITP (Inspection & Test Plan)</option>
                          <option value="RFC">RFC (Request For Change)</option>
                          <option value="RFI">RFI (Request For Information)</option>
                          <option value="SI">SI (Site Instruction)</option>
                          <option value="NCR">NCR (Non-Conformance Report)</option>
                          <option value="CAR">CAR (Corrective Action Report)</option>
                          <option value="TRN">TRN (Transmittal)</option>
                          <option value="CMT">CMT (Comment)</option>
                          <option value="SR">SR (Site Report)</option>
                          <option value="AD">AD (Architectural Drawing)</option>
                          <option value="SD">SD (Shop Drawing)</option>
                          <option value="ABD">ABD (As-Built Drawing)</option>
                          <option value="MOH">MOH (Method of Handover)</option>
                          <option value="SPT">SPT (Submittal)</option>
                        </datalist>
                      </div>
                      <div>
                        <label style="font-size: 0.75rem; font-weight: 700; color: #334155; margin-bottom: 3px; display: block; white-space: nowrap;">Phân nhánh 1</label>
                        <select class="form-control" name="sub_type1" id="cboSubType1" style="font-size: 0.82rem; font-weight: 600; height: 32px; padding: 4px 8px; white-space: nowrap;">
                          <option value="">-- Nhánh 1 --</option>
                        </select>
                      </div>
                      <div>
                        <label style="font-size: 0.75rem; font-weight: 700; color: #334155; margin-bottom: 3px; display: block; white-space: nowrap;">Phân nhánh 2</label>
                        <select class="form-control" name="sub_type2" id="cboSubType2" style="font-size: 0.82rem; font-weight: 600; height: 32px; padding: 4px 8px; white-space: nowrap;">
                          <option value="">-- Nhánh 2 --</option>
                        </select>
                      </div>
                    </div>
                  </div>
                </div>

                <!-- 2. Name (Tiêu đề tài liệu) -->
                <div class="form-group" style="margin-bottom: 14px;">
                  <label class="form-label">Name (Tiêu đề hồ sơ / Tên chứng từ) <span style="color: var(--danger);">*</span></label>
                  <input type="text" class="form-control" name="title" id="input-doc-title" value="${editDoc?.title || ''}" placeholder="Vd: Interim Payment Certificate - Thi công kết cấu... hoặc Bản vẽ nghiệm thu" required style="font-weight: 600;">
                </div>

                <!-- 3. Issuer & Recipient -->
                <div style="display: grid; grid-template-columns: 1fr auto 1fr; gap: 8px; align-items: flex-end; margin-bottom: 14px;">
                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label">Issuer (Bên gửi / Ban hành)</label>
                    <input type="text" class="form-control" name="issuer" id="input-doc-issuer" list="list-doc-issuers" value="${editDoc?.issuer || targetContract.code_b || ''}" placeholder="Vd: VNCN, CTC..." style="text-transform: uppercase; font-family: monospace; font-weight: 600;">
                    <datalist id="list-doc-issuers">
                      ${companyCodes.map(c => `<option value="${c}"></option>`).join('')}
                    </datalist>
                  </div>

                  <div style="padding-bottom: 2px;">
                    <button type="button" class="button button-outline" id="btn-swap-doc-parties" style="padding: 7px 10px;" title="Hoán đổi Bên gửi <=> Bên nhận">
                      <i class="fa-solid fa-right-left"></i>
                    </button>
                  </div>

                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label">Recipient (Bên nhận)</label>
                    <input type="text" class="form-control" name="recipients" id="input-doc-recipients" list="list-doc-recipients" value="${editDoc?.recipients || targetContract.code_a || ''}" placeholder="Vd: VEC, CLI..." style="text-transform: uppercase; font-family: monospace; font-weight: 600;">
                    <datalist id="list-doc-recipients">
                      ${companyCodes.map(c => `<option value="${c}"></option>`).join('')}
                    </datalist>
                  </div>
                </div>

                <!-- 4. Number (Số hiệu hồ sơ / Số công văn) -->
                <div class="form-group" style="margin-bottom: 14px;">
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <label class="form-label" style="margin-bottom: 0;">Number (Số công văn / Số hiệu hồ sơ) <span style="color: var(--danger);">*</span></label>
                    <div style="display: flex; gap: 10px; align-items: center;">
                      <label style="display: inline-flex; align-items: center; gap: 5px; font-size: 0.75rem; color: #475569; cursor: pointer; user-select: none; white-space: nowrap;" title="Tick chọn để tạo số mới tăng dần (+1). Bỏ tick để sinh số bản sửa đổi / Revision (.A, .B...)">
                        <input type="checkbox" id="chkNewDoc" checked>
                        <span>Tạo số mới</span>
                      </label>
                      <button type="button" id="btn-auto-gen-docno" class="button button-outline" style="padding: 2px 8px; font-size: 0.72rem; color: #4338ca; border-color: #c7d2fe; background: #eef2ff;">
                        <i class="fa-solid fa-wand-magic-sparkles"></i> Sinh số chuẩn
                      </button>
                    </div>
                  </div>
                  <input type="text" class="form-control" name="doc_no" id="input-doc-no" value="${editDoc?.doc_no || ''}" placeholder="Vd: 260715.CHC1.NSH.IPC.EC.001 hoặc 125/CV-CHC1" required style="font-family: monospace; font-weight: 700; font-size: 0.88rem;">
                </div>

                <!-- 5. Link (Đường dẫn tệp tin trên máy tính / OneDrive / Mạng) -->
                <div class="form-group" style="margin-bottom: 14px;">
                  <label class="form-label" style="display: flex; justify-content: space-between; align-items: center;">
                    <span><i class="fa-solid fa-laptop"></i> Link (Đường dẫn tệp tin trên máy tính / Local / OneDrive)</span>
                    <button type="button" id="btn-browse-doc-file" class="button button-outline" style="padding: 2px 8px; font-size: 0.73rem; color: #0284c7; border-color: #bae6fd;">
                      <i class="fa-solid fa-folder-open"></i> Duyệt file...
                    </button>
                  </label>
                  <div style="display: flex; gap: 6px;">
                    <input type="text" class="form-control" name="local_link" id="input-doc-local-link" value="${editDoc?.local_link || (!String(editDoc?.file_link || '').startsWith('http') ? editDoc?.file_link : '') || ''}" placeholder="Vd: D:\\OneDrive\\DuAn_CHC1\\ThanhToan\\IPC03.xlsx" style="font-size: 0.85rem; font-family: monospace;">
                    <button type="button" id="btn-copy-doc-input-link" class="button button-outline" style="padding: 6px 10px;" title="Sao chép đường dẫn">
                      <i class="fa-regular fa-copy"></i>
                    </button>
                  </div>
                  <small style="color: #64748b; font-size: 0.72rem;">Đường dẫn file thực tế trên ổ đĩa máy tính hoặc thư mục dùng chung của dự án.</small>
                </div>

                <!-- 6. Link Drive Online -->
                <div class="form-group" style="margin-bottom: 14px;">
                  <label class="form-label" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; gap: 6px; flex-wrap: wrap;">
                    <span><i class="fa-brands fa-google-drive" style="color: #0284c7;"></i> Google Drive Link (Xem trực tuyến Online URL)</span>
                    <span style="display: inline-flex; gap: 6px; align-items: center;">
                      <button type="button" id="btn-doc-edit-open-drive-file" class="button button-outline" style="padding: 2px 8px; font-size: 0.73rem; color: #0284c7; border-color: #bae6fd; background: #f0f9ff;" title="Mở file trên Google Drive">
                        <i class="fa-solid fa-up-right-from-square"></i> Mở file
                      </button>
                      <button type="button" id="btn-doc-edit-open-drive-folder" class="button button-outline" style="padding: 2px 8px; font-size: 0.73rem; color: #0369a1; border-color: #bae6fd; background: #f0f9ff;" title="Mở thư mục chứa hồ sơ trên Google Drive">
                        <i class="fa-brands fa-google-drive"></i> Mở thư mục Drive
                      </button>
                    </span>
                  </label>
                  <input type="url" class="form-control" name="drive_link" id="input-doc-drive-link" value="${editDoc?.drive_link || (String(editDoc?.file_link || '').startsWith('http') ? editDoc?.file_link : '') || ''}" placeholder="https://drive.google.com/file/d/..." style="font-size: 0.85rem;">
                </div>

                <!-- 7. Ngày xử lý & Ngày hoàn thành -->
                <div class="form-grid-2" style="margin-bottom: 14px;">
                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; gap: 6px;">
                      <span style="font-size: 0.82rem; font-weight: 600; color: #334155; white-space: nowrap;">Ngày xử lý</span>
                      <button type="button" class="btn-set-today" data-target="input-doc-start-date" style="border: 1px solid #bfdbfe; background: #eff6ff; color: #1d4ed8; font-size: 0.72rem; font-weight: 600; padding: 2px 7px; border-radius: 4px; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; flex-shrink: 0;">
                        <i class="fa-regular fa-calendar-days" style="font-size: 0.68rem;"></i> Hôm nay
                      </button>
                    </label>
                    <input type="date" class="form-control" name="start_date" id="input-doc-start-date" value="${initialProcessedDate}">
                  </div>

                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; gap: 6px;">
                      <span style="font-size: 0.82rem; font-weight: 600; color: #334155; white-space: nowrap;">Ngày hoàn thành</span>
                      <button type="button" class="btn-set-today" data-target="input-doc-due-date" style="border: 1px solid #bfdbfe; background: #eff6ff; color: #1d4ed8; font-size: 0.72rem; font-weight: 600; padding: 2px 7px; border-radius: 4px; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; flex-shrink: 0;">
                        <i class="fa-regular fa-calendar-days" style="font-size: 0.68rem;"></i> Hôm nay
                      </button>
                    </label>
                    <input type="date" class="form-control" name="due_date" id="input-doc-due-date" value="${initialCompletedDate}">
                  </div>
                </div>

                <!-- 8. Send To & Status -->
                <div class="form-grid-2" style="margin-bottom: 14px;">
                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label">Người nhận lại / Phụ trách</label>
                    <input type="text" class="form-control" name="recipient" id="input-doc-recipient" value="${editDoc?.recipient || ''}" placeholder="Vd: Đạt QS KD hoặc Ban QLDA">
                  </div>

                  <div class="form-group" style="margin-bottom: 0;">
                    <label class="form-label">Tình trạng hồ sơ</label>
                    <select class="form-control" name="status" id="select-doc-status">
                      <option value="Completed" ${(editDoc?.status || 'Completed') === 'Completed' ? 'selected' : ''}>Đã hoàn thành (Completed)</option>
                      <option value="Accepted" ${editDoc?.status === 'Accepted' ? 'selected' : ''}>Đã phê duyệt (Accepted)</option>
                      <option value="Approved" ${editDoc?.status === 'Approved' ? 'selected' : ''}>Đã chấp thuận (Approved)</option>
                      <option value="Checking" ${editDoc?.status === 'Checking' ? 'selected' : ''}>Đang kiểm tra / Rà soát</option>
                      <option value="Pending" ${editDoc?.status === 'Pending' ? 'selected' : ''}>Chờ xử lý (Pending)</option>
                      <option value="UnderReview" ${editDoc?.status === 'UnderReview' ? 'selected' : ''}>Đang xem xét (UnderReview)</option>
                      <option value="Draft" ${editDoc?.status === 'Draft' ? 'selected' : ''}>Dự thảo (Draft)</option>
                    </select>
                  </div>
                </div>

                <!-- 9. Note (Ghi chú) -->
                <div class="form-group" style="margin-bottom: 0;">
                  <label class="form-label">Ghi chú diễn biến hồ sơ</label>
                  <textarea class="form-control" name="notes" id="input-doc-notes" rows="2" placeholder="Nhập ghi chú chi tiết về tình trạng bàn giao, xử lý hồ sơ...">${editDoc?.notes || ''}</textarea>
                </div>
              </div>

              <!-- CỘT PHẢI: SUPPORTING DOCUMENTS (HỒ SƠ KÈM THEO CHUẨN CPMS PLUS) -->
              <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: var(--radius-sm); padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <div style="font-weight: 700; font-size: 0.9rem; color: #1e293b; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #e2e8f0; padding: 0 4px 8px 4px;">
                  <span><i class="fa-solid fa-layer-group" style="color: #6366f1;"></i> Supporting Documents</span>
                  <span style="font-size: 0.75rem; color: #64748b; font-weight: 600; white-space: nowrap;">Số lượng</span>
                </div>

                <div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.82rem;">
                  <!-- 1. GTTT -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="GTTT" ${supportingObj['GTTT'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Payment Certificate (GTTT)</span>
                    </label>
                    ${renderQtySelect('GTTT', supportingObj['GTTT'] || 0)}
                  </div>

                  <!-- 2. THKL -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="THKL" ${supportingObj['THKL'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Quantity Summary (THKL)</span>
                    </label>
                    ${renderQtySelect('THKL', supportingObj['THKL'] || 0)}
                  </div>

                  <!-- 3. DGKL -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="DGKL" ${supportingObj['DGKL'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Quantity Breakdown (DGKL)</span>
                    </label>
                    ${renderQtySelect('DGKL', supportingObj['DGKL'] || 0)}
                  </div>

                  <!-- 4. NTCL -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="NTCL" ${supportingObj['NTCL'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Quality Acceptance (NTCL)</span>
                    </label>
                    ${renderQtySelect('NTCL', supportingObj['NTCL'] || 0)}
                  </div>

                  <!-- 5. NTVL -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="NTVL" ${supportingObj['NTVL'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Material Acceptance (NTVL)</span>
                    </label>
                    ${renderQtySelect('NTVL', supportingObj['NTVL'] || 0)}
                  </div>

                  <!-- 6. HC (As-Built / Hoàn công) -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="HC" ${supportingObj['HC'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">As-Built Documents (HC)</span>
                    </label>
                    ${renderQtySelect('HC', supportingObj['HC'] || 0)}
                  </div>

                  <!-- 7. HSK (Hồ sơ khác) -->
                  <div style="display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; border-radius: 4px; background: #f8fafc; gap: 8px;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 1; margin: 0; user-select: none; min-width: 0; white-space: nowrap;">
                      <input type="checkbox" class="chk-supporting" data-code="HSK" ${supportingObj['HSK'] ? 'checked' : ''}>
                      <span style="font-size: 0.82rem; font-weight: 500; color: #334155;">Other Documents (HSK)</span>
                    </label>
                    ${renderQtySelect('HSK', supportingObj['HSK'] || 0)}
                  </div>

                  <!-- Ghi chú hồ sơ kèm theo (txtHoSoKemNote chuẩn VBA) -->
                  <div style="margin-top: 6px;">
                    <label class="form-label" style="font-size: 0.78rem; font-weight: 600; color: #475569; margin-bottom: 4px;">
                      Ghi chú hồ sơ kèm theo (HoSoKem Note)
                    </label>
                    <input type="text" class="form-control" id="input-supporting-note" value="${supportingNote}" placeholder="Vd: Bộ CO CQ gốc x1 hoặc Báo cáo tiến độ x5" style="font-size: 0.8rem;">
                  </div>
                </div>

                <div style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed #e2e8f0; font-size: 0.72rem; color: #64748b;">
                  Mã hóa tự động chuẩn CPMS Plus: Lưu chuỗi cấu trúc hồ sơ đính kèm vào sheet <code>DOCUMENT_TRACKER</code>.
                </div>
              </div>
            </div>

            <!-- Footer Buttons -->
            <div style="display: flex; justify-content: ${isEdit ? 'space-between' : 'flex-end'}; align-items: center; gap: 10px; margin-top: 20px; padding-top: 14px; border-top: 1px solid var(--border-color); flex-wrap: wrap;">
              ${
                isEdit
                  ? `<button type="button" id="btn-doc-edit-delete" class="button button-outline" style="color: var(--danger); border-color: #fca5a5; font-weight: 600;">
                      <i class="fa-solid fa-trash-can"></i> Xóa hồ sơ
                    </button>`
                  : ''
              }
              <div style="display: flex; gap: 10px; align-items: center; margin-left: auto;">
                <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
                <button type="submit" class="button button-primary" style="padding: 8px 20px; font-weight: 700;">
                  <i class="fa-solid fa-floppy-disk"></i> ${isEdit ? 'Lưu thay đổi Hồ sơ' : 'Lưu Hồ sơ mới'}
                </button>
              </div>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Nút Open File trên Header của form
    modal.querySelector('#btn-doc-edit-open-file')?.addEventListener('click', () => {
      const locLink = modal.querySelector('#input-doc-local-link')?.value || '';
      const drvLink = modal.querySelector('#input-doc-drive-link')?.value || '';
      this.handleOpenDocFile({ local_link: locLink, drive_link: drvLink, file_link: drvLink || locLink });
    });

    // Nút Open Folder trên Header của form
    modal.querySelector('#btn-doc-edit-open-folder')?.addEventListener('click', () => {
      const locLink = modal.querySelector('#input-doc-local-link')?.value || '';
      const drvLink = modal.querySelector('#input-doc-drive-link')?.value || '';
      this.handleOpenDocFolder({ local_link: locLink, drive_link: drvLink, file_link: drvLink || locLink }, targetContract);
    });

    // Nút Mở file Drive ở dòng Google Drive Link
    modal.querySelector('#btn-doc-edit-open-drive-file')?.addEventListener('click', () => {
      const drvLink = (modal.querySelector('#input-doc-drive-link')?.value || '').trim();
      if (drvLink && drvLink.startsWith('http')) {
        window.open(drvLink, '_blank');
      } else {
        this.showAppAlert?.('Hồ sơ này chưa có Google Drive Link.');
      }
    });

    // Nút Mở thư mục Drive ở dòng Google Drive Link
    modal.querySelector('#btn-doc-edit-open-drive-folder')?.addEventListener('click', async () => {
      const gdriveSvc = window.__MAIO_GDRIVE_SERVICE__;
      if (gdriveSvc && gdriveSvc.isSignedIn()) {
        try {
          const curPrjId = this.currentProjectId || 'CHC1';
          const safePrjId = String(curPrjId).replace(/[\/\\?%*:|"<>]/g, '_');
          const prjFolderId = await gdriveSvc.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${safePrjId}`]);
          const docNoVal = modal.querySelector('#txtSoCV')?.value || editDoc?.doc_no || '';
          const titleVal = modal.querySelector('input[name="title"]')?.value || editDoc?.title || '';
          const locVal = modal.querySelector('#input-doc-local-link')?.value || editDoc?.local_link || '';
          const catRadio = modal.querySelector('input[name="category"]:checked')?.value || editDoc?.category || 'OTHER';
          const resolvedCat = catRadio === 'Other' ? (modal.querySelector('#cboDocType')?.value || 'OTHER') : (catRadio === 'Contract' ? 'CON' : 'IPC');
          const segs = this.getDrivePathSegmentsForDoc({
            contract_id: targetContract.id,
            category: resolvedCat,
            doc_no: docNoVal,
            title: titleVal,
            local_link: locVal,
          });
          const targetFolderId = await gdriveSvc.ensureFolderPath(segs, prjFolderId);
          window.open(`https://drive.google.com/drive/folders/${targetFolderId}`, '_blank');
          return;
        } catch (err) {
          console.warn('Không mở được thư mục Drive trực tiếp:', err);
        }
      }
      window.open('https://drive.google.com/', '_blank');
    });

    // Nút Xóa hồ sơ ngay trong popup sửa hồ sơ
    modal.querySelector('#btn-doc-edit-delete')?.addEventListener('click', async () => {
      if (!editDoc) return;
      const ok = await this.confirmAsync({
        title: 'Xác nhận xóa hồ sơ',
        message: `Bạn có chắc muốn xóa hồ sơ "${editDoc.doc_no || editDoc.id}" không?`,
        confirmText: 'Xóa hồ sơ',
        cancelText: 'Hủy bỏ',
        confirmType: 'danger',
      });
      if (!ok) return;
      await db.delete('documents', editDoc.id);
      await this.loadData();
      await this.syncToLocalFile();
      this.updateReminderBadge();
      closeSubModal();
      this.openContractDetail(targetContract.id);
      setTimeout(() => {
        document.querySelector('.modal-tab-btn[data-target="m-tab-docs"]')?.click();
      }, 60);
      this.showLocalSaveStatus(`Đã xóa hồ sơ "${editDoc.doc_no || editDoc.id}"!`);
    });

    // Nút Hoán đổi Bên gửi <=> Bên nhận
    modal.querySelector('#btn-swap-doc-parties')?.addEventListener('click', () => {
      const inputIssuer = modal.querySelector('#input-doc-issuer');
      const inputRecip = modal.querySelector('#input-doc-recipients');
      if (inputIssuer && inputRecip) {
        const tmp = inputIssuer.value;
        inputIssuer.value = inputRecip.value;
        inputRecip.value = tmp;
      }
    });

    // Nút Thiết lập ngày hôm nay (today)
    modal.querySelectorAll('.btn-set-today').forEach((btn) => {
      btn.addEventListener('click', () => {
        const targetInputId = btn.getAttribute('data-target');
        const inputEl = modal.querySelector(`#${targetInputId}`);
        if (inputEl) {
          inputEl.value = new Date().toISOString().slice(0, 10);
        }
      });
    });

    // Nút Duyệt file... (Browse file)
    const hiddenFileInput = modal.querySelector('#doc-file-hidden-picker');
    const localLinkInput = modal.querySelector('#input-doc-local-link');
    const titleInput = modal.querySelector('#input-doc-title');

    modal.querySelector('#btn-browse-doc-file')?.addEventListener('click', () => {
      hiddenFileInput?.click();
    });

    hiddenFileInput?.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) {
        // Lấy tên file hoặc đường dẫn gợi ý
        const defaultDocFolder = `D:\\OneDrive\\DuAn_${this.currentProjectId || 'CHC1'}\\`;
        const suggestedPath = defaultDocFolder + file.name;
        if (localLinkInput) {
          localLinkInput.value = suggestedPath;
        }
        if (titleInput && (!titleInput.value || titleInput.value.trim() === '')) {
          const nameWithoutExt = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
          titleInput.value = nameWithoutExt;
        }
      }
    });

    // Nút sao chép đường dẫn ô Link
    modal.querySelector('#btn-copy-doc-input-link')?.addEventListener('click', () => {
      const val = localLinkInput?.value || '';
      if (val) {
        navigator.clipboard.writeText(val).then(() => {
          alert(`Đã sao chép đường dẫn: ${val}`);
        });
      }
    });

    // Xử lý chọn số lượng Supporting Documents qua dropdown combobox (0..10)
    modal.querySelectorAll('.select-supporting-qty').forEach((sel) => {
      sel.addEventListener('change', () => {
        const code = sel.getAttribute('data-code');
        const chk = modal.querySelector(`.chk-supporting[data-code="${code}"]`);
        const val = parseInt(sel.value, 10) || 0;
        if (chk) {
          chk.checked = val > 0;
        }
        if (val > 0) {
          sel.style.backgroundColor = '#eff6ff';
          sel.style.borderColor = '#93c5fd';
          sel.style.color = '#1d4ed8';
        } else {
          sel.style.backgroundColor = '#f8fafc';
          sel.style.borderColor = '#cbd5e1';
          sel.style.color = '#64748b';
        }
      });
    });

    // Khi người dùng tick checkbox: nếu đang là 0 thì chuyển thành 1. Khi bỏ tick: chuyển về 0
    modal.querySelectorAll('.chk-supporting').forEach((chk) => {
      chk.addEventListener('change', () => {
        const code = chk.getAttribute('data-code');
        const sel = modal.querySelector(`.select-supporting-qty[data-code="${code}"]`);
        if (sel) {
          if (chk.checked) {
            if (sel.value === '0') sel.value = '1';
            sel.style.backgroundColor = '#eff6ff';
            sel.style.borderColor = '#93c5fd';
            sel.style.color = '#1d4ed8';
          } else {
            sel.value = '0';
            sel.style.backgroundColor = '#f8fafc';
            sel.style.borderColor = '#cbd5e1';
            sel.style.color = '#64748b';
          }
        }
      });
    });

    // Cấu trúc phân nhánh SubType 1 và SubType 2 theo mã nguồn VBA frmDocumentEdit
    const subType1Map = {
      RFA: ['ITP', 'MSA', 'HSE', 'SCH', 'GEN', 'MS', 'SD'],
      MOM: ['HSE', 'MEP', 'KOM'],
    };

    const subType2Map = {
      GEN: ['FS', 'E'],
      MSA: ['HVAC', 'E', 'FS', 'PL', 'FF'],
      ITP: ['HVAC', 'E', 'PL', 'FF'],
      SCH: ['E'],
      MS: ['HVAC', 'E', 'PL', 'FF'],
      SD: ['HVAC', 'E', 'ELV', 'FA', 'PL', 'FF'],
    };

    const updateSubType1 = (dt) => {
      const cbo1 = modal.querySelector('#cboSubType1');
      if (!cbo1) return;
      const upperDt = String(dt || '').trim().toUpperCase();
      const list = subType1Map[upperDt] || [];
      const curVal = cbo1.value;
      cbo1.innerHTML = '<option value="">-- Nhánh 1 --</option>' + list.map((s) => `<option value="${s}" ${curVal === s ? 'selected' : ''}>${s}</option>`).join('');
      cbo1.disabled = list.length === 0;
      if (list.length === 0) {
        cbo1.value = '';
        updateSubType2('');
      } else {
        updateSubType2(cbo1.value);
      }
    };

    const updateSubType2 = (st1) => {
      const cbo2 = modal.querySelector('#cboSubType2');
      if (!cbo2) return;
      const upperSt1 = String(st1 || '').trim().toUpperCase();
      const list = subType2Map[upperSt1] || [];
      const curVal = cbo2.value;
      cbo2.innerHTML = '<option value="">-- Nhánh 2 --</option>' + list.map((s) => `<option value="${s}" ${curVal === s ? 'selected' : ''}>${s}</option>`).join('');
      cbo2.disabled = list.length === 0;
      if (list.length === 0) {
        cbo2.value = '';
      }
    };

    // Hàm sinh số chuẩn CPMS Plus (TryGenerateDocNumber chuẩn 100% VBA)
    const tryGenerateDocNumber = () => {
      const catRadio = modal.querySelector('input[name="category"]:checked')?.value || 'Other';
      const docNoInput = modal.querySelector('#input-doc-no');
      const chkNew = modal.querySelector('#chkNewDoc');
      if (!docNoInput) return;

      if (catRadio === 'Contract') {
        docNoInput.value = targetContract.contract_no || targetContract.id;
        if (chkNew) {
          chkNew.checked = true;
          chkNew.disabled = true;
        }
        return;
      }

      const issuerInput = modal.querySelector('#input-doc-issuer');
      const recipInput = modal.querySelector('#input-doc-recipients');
      const issuer = (issuerInput?.value || '').trim().toUpperCase();
      const recip = (recipInput?.value || '').trim().toUpperCase();
      if (!issuer || !recip) return;

      let docTypeCode = '';
      if (catRadio === 'IPC') {
        docTypeCode = 'IPC';
      } else {
        const cboDocType = modal.querySelector('#cboDocType');
        docTypeCode = (cboDocType?.value || '').trim().toUpperCase() || 'RFA';
      }

      // Ghép APP/PKG suffix nếu là Phụ lục / Gói con (Chuẩn CPMS Plus)
      let fullDocType = docTypeCode;
      const cno = String(targetContract.contract_no || '').trim().toUpperCase();
      if (cno.startsWith('APP-') || cno.startsWith('PKG-')) {
        fullDocType = fullDocType + '-' + cno.replace(/-/g, '');
      }

      // Ghép SubType1 và SubType2 nếu có (dành cho RFA, MOM...)
      if (catRadio === 'Other') {
        const st1 = modal.querySelector('#cboSubType1')?.value?.trim().toUpperCase();
        const st2 = modal.querySelector('#cboSubType2')?.value?.trim().toUpperCase();
        if (st1) fullDocType += '-' + st1;
        if (st2) fullDocType += '-' + st2;
      }

      // Ngày YYMMDD
      const dateVal = modal.querySelector('#input-doc-start-date')?.value || new Date().toISOString().slice(0, 10);
      const yymmdd = dateVal.slice(2, 4) + dateVal.slice(5, 7) + dateVal.slice(8, 10);
      const projCode = this.currentProjectId || 'CHC1';

      // Tính serial cuối cùng theo đúng phân loại hồ sơ (targetContract.id, docTypeCode, issuer, recip)
      const allDocs = this.documents || [];
      let maxSerial = 0;
      allDocs.forEach((d) => {
        if (d.contract_id !== targetContract.id) return;
        const dIssuer = String(d.issuer || '').split(';')[0].trim().toUpperCase();
        if (dIssuer !== issuer) return;
        const dRecip = String(d.recipients || '').split(';')[0].trim().toUpperCase();
        if (dRecip !== recip) return;

        let dCat = String(d.category || '').trim().toUpperCase();
        if (dCat === 'CONTRACT') dCat = 'CON';
        if (dCat !== docTypeCode && dCat !== 'OTHER') return;
        if (dCat === 'OTHER') {
          const dOtherType = String(d.other_doctype || d.sub_type1 || '').trim().toUpperCase();
          if (dOtherType && dOtherType !== docTypeCode) return;
        }

        const dNo = String(d.doc_no || '').trim();
        const parts = dNo.split('.');
        for (let p of parts) {
          if (/^\d{3}$/.test(p)) {
            const sn = parseInt(p, 10);
            if (sn > maxSerial) maxSerial = sn;
            break;
          }
        }
      });

      if (maxSerial === 0) {
        if (chkNew) {
          chkNew.checked = true;
          chkNew.disabled = true;
        }
        docNoInput.value = `${yymmdd}.${projCode}.${issuer}.${fullDocType}.${recip}.001`;
      } else {
        if (chkNew) chkNew.disabled = false;
        const isNew = chkNew ? chkNew.checked : true;
        if (isNew) {
          const nextSerial = String(maxSerial + 1).padStart(3, '0');
          docNoInput.value = `${yymmdd}.${projCode}.${issuer}.${fullDocType}.${recip}.${nextSerial}`;
        } else {
          // Tìm version suffix tiếp theo (A, B, C...)
          const usedSuffixes = new Set();
          allDocs.forEach((d) => {
            if (d.contract_id !== targetContract.id) return;
            const dIssuer = String(d.issuer || '').split(';')[0].trim().toUpperCase();
            if (dIssuer !== issuer) return;
            const dRecip = String(d.recipients || '').split(';')[0].trim().toUpperCase();
            if (dRecip !== recip) return;

            const parts = String(d.doc_no || '').split('.');
            const sIdx = parts.findIndex((p) => p === String(maxSerial).padStart(3, '0'));
            if (sIdx !== -1 && sIdx + 1 < parts.length) {
              const suf = parts[sIdx + 1].trim().toUpperCase();
              if (/^[A-Z]$/.test(suf)) usedSuffixes.add(suf);
            }
          });

          let verSuffix = 'A';
          for (let i = 0; i < 26; i++) {
            const ch = String.fromCharCode(65 + i);
            if (!usedSuffixes.has(ch)) {
              verSuffix = ch;
              break;
            }
          }
          docNoInput.value = `${yymmdd}.${projCode}.${issuer}.${fullDocType}.${recip}.${String(maxSerial).padStart(3, '0')}.${verSuffix}`;
        }
      }
    };

    // Sự kiện chuyển đổi Document Type (Contract / IPC / Other)
    modal.querySelectorAll('input[name="category"]').forEach((r) => {
      r.addEventListener('change', () => {
        const val = r.value;
        const otherContainer = modal.querySelector('#other-doctype-container');
        if (otherContainer) {
          otherContainer.style.display = val === 'Other' ? 'block' : 'none';
        }
        if (val === 'Other') {
          const dtVal = modal.querySelector('#cboDocType')?.value || 'RFA';
          updateSubType1(dtVal);
        }
        tryGenerateDocNumber();
      });
    });

    // Sự kiện khi thay đổi cboDocType, cboSubType1, cboSubType2
    modal.querySelector('#cboDocType')?.addEventListener('input', (e) => {
      updateSubType1(e.target.value);
      tryGenerateDocNumber();
    });
    modal.querySelector('#cboDocType')?.addEventListener('change', (e) => {
      updateSubType1(e.target.value);
      tryGenerateDocNumber();
    });
    modal.querySelector('#cboSubType1')?.addEventListener('change', (e) => {
      updateSubType2(e.target.value);
      tryGenerateDocNumber();
    });
    modal.querySelector('#cboSubType2')?.addEventListener('change', () => {
      tryGenerateDocNumber();
    });

    // Sự kiện tick / bỏ tick chkNewDoc
    modal.querySelector('#chkNewDoc')?.addEventListener('change', () => {
      tryGenerateDocNumber();
    });

    // Nút Sinh số chuẩn
    modal.querySelector('#btn-auto-gen-docno')?.addEventListener('click', () => {
      tryGenerateDocNumber();
    });

    // Khi thay đổi ngày nhận hoặc đối tác -> cập nhật số nếu cần
    modal.querySelector('#input-doc-start-date')?.addEventListener('change', () => {
      tryGenerateDocNumber();
    });
    modal.querySelector('#input-doc-issuer')?.addEventListener('change', () => {
      tryGenerateDocNumber();
    });
    modal.querySelector('#input-doc-recipients')?.addEventListener('change', () => {
      tryGenerateDocNumber();
    });

    // Khởi tạo trạng thái ban đầu của cboSubType1 và cboSubType2
    const initialDt = modal.querySelector('#cboDocType')?.value || 'RFA';
    updateSubType1(initialDt);
    if (editDoc?.sub_type1) {
      const c1 = modal.querySelector('#cboSubType1');
      if (c1) {
        c1.value = editDoc.sub_type1;
        updateSubType2(editDoc.sub_type1);
        if (editDoc.sub_type2) {
          const c2 = modal.querySelector('#cboSubType2');
          if (c2) c2.value = editDoc.sub_type2;
        }
      }
    }

    // Xử lý lưu hồ sơ
    modal.querySelector('#form-doc-edit').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);

      // Mã hóa supporting docs chuẩn CPMS Plus VBA: GTTT:10|THKL:10|NTVL:1|HSK:1|NOTE:...
      const supportingArr = [];
      ['GTTT', 'THKL', 'DGKL', 'NTCL', 'NTVL', 'HC', 'HSK'].forEach((code) => {
        const chk = modal.querySelector(`.chk-supporting[data-code="${code}"]`);
        const sel = modal.querySelector(`.select-supporting-qty[data-code="${code}"]`);
        const qty = parseInt(sel?.value, 10) || 0;
        if (chk && chk.checked && qty > 0) {
          supportingArr.push(`${code}:${qty}`);
        }
      });
      const noteVal = modal.querySelector('#input-supporting-note')?.value?.trim();
      if (noteVal) {
        supportingArr.push(`NOTE:${noteVal}`);
      }
      const supportingDocsStr = supportingArr.join('|');

      const rawCat = formData.get('category') || 'Other';
      let finalCategory = rawCat;
      const otherDocType = (formData.get('other_doctype') || '').trim().toUpperCase();
      if (rawCat === 'Other' && otherDocType) {
        finalCategory = otherDocType;
      }

      const localLink = formData.get('local_link')?.trim() || '';
      const driveLink = formData.get('drive_link')?.trim() || '';

      const docData = {
        ...(editDoc || {}),
        id: formData.get('id'),
        project_id: targetContract.project_id || this.currentProjectId,
        contract_id: targetContract.id,
        category: finalCategory,
        other_doctype: otherDocType,
        sub_type1: (formData.get('sub_type1') || '').trim().toUpperCase(),
        sub_type2: (formData.get('sub_type2') || '').trim().toUpperCase(),
        title: (formData.get('title') || '').trim(),
        doc_no: (formData.get('doc_no') || '').trim(),
        issuer: (formData.get('issuer') || '').trim().toUpperCase(),
        recipients: (formData.get('recipients') || '').trim().toUpperCase(),
        local_link: localLink,
        drive_link: driveLink,
        file_link: driveLink || localLink || '',
        start_date: formData.get('start_date') || null,
        due_date: formData.get('due_date') || null,
        received_date: formData.get('start_date') || null,
        return_date: formData.get('due_date') || null,
        recipient: (formData.get('recipient') || '').trim(),
        status: formData.get('status') || 'Completed',
        notes: (formData.get('notes') || '').trim(),
        supporting_docs: supportingDocsStr,
        quantity: 1,
        lan_val: editDoc?.lan_val || 1,
      };

      await db.put('documents', docData);
      await this.loadData();
      closeSubModal();
      await this.syncToLocalFile();
      this.openContractDetail(targetContract.id);
      alert(`Đã lưu thành công thông tin Hồ sơ "${docData.doc_no || docData.id}"!`);
    });
  }

  // ==========================================
  // FORM NHẬP THANH TOÁN (PAYMENT ENTRY) & ĐỢT CHI CPB
  // ==========================================
  openPaymentEntryModal(targetContract = null) {
    // Nếu chưa truyền targetContract, ưu tiên chọn HĐ đang chọn trong tab giải ngân hoặc HĐ đầu tiên của nhóm
    if (!targetContract) {
      if (this.selectedDisburseContractId) {
        targetContract = this.contracts.find((c) => c.id === this.selectedDisburseContractId);
      }
      if (!targetContract) {
        const groupContracts = this.contracts.filter((c) => c.group === (this.disburseGroup || 'A'));
        targetContract = groupContracts[0] || this.contracts[0];
      }
    }

    if (!targetContract) {
      alert('Chưa có hợp đồng nào trong dự án để thêm đợt thanh toán. Vui lòng tạo hợp đồng trước!');
      return;
    }

    const isCpb = targetContract.is_cpb || targetContract.contract_type === 'CPB' || (targetContract.id && targetContract.id.startsWith('B-CPB-'));
    const realPayments = this.getRealPayments(targetContract.id);
    const nextIdx = realPayments.length + 1;
    const ledgerId = `LG-${Date.now()}`;

    // Lọc danh sách hợp đồng khả dụng để người dùng có thể đổi HĐ ngay trong form
    const availableContracts = this.contracts.filter(c => c.project_id === (targetContract.project_id || this.currentProjectId));
    const defaultPayRate = Number(targetContract.exchange_rate) || 1;

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 700px;">
        <div class="modal-header">
          <h3 id="payment-modal-title" style="font-size: 1.05rem; font-weight: 700; color: ${isCpb ? '#92400e' : 'inherit'};">
            ${isCpb ? `<i class="fa-solid fa-receipt" style="color: #d97706;"></i> Nhập đợt Chi phí cho ${targetContract.scope_of_work || targetContract.contract_no}` : `Nhập đợt Giải ngân cho ${targetContract.contract_no}`}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-payment">
            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label">Chọn Hợp đồng / Khoản chi phí cần giải ngân <span style="color: var(--danger);">*</span></label>
              <select class="form-control" name="contract_id" id="payment-select-contract" required>
                ${availableContracts.map((c) => {
                  const isC = c.id === targetContract.id;
                  const cType = c.is_cpb || c.contract_type === 'CPB' ? 'CPB' : `Nhóm ${c.group}`;
                  return `<option value="${c.id}" ${isC ? 'selected' : ''}>[${cType}] ${c.contract_no || c.id} — ${c.scope_of_work || (c.is_cpb ? 'Chi phí Ban' : 'Hợp đồng')}</option>`;
                }).join('')}
              </select>
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label" id="lbl-payment-no">${isCpb ? 'Số Phiếu chi / Đợt chi' : 'Số ĐN thanh toán'}</label>
                <input type="text" class="form-control" name="payment_no" value="${String(nextIdx).padStart(2, '0')}" placeholder="Vd: 01 hoặc PC-01" required>
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-payment-date">${isCpb ? 'Ngày chi tiền' : 'Ngày thanh toán'}</label>
                <input type="date" class="form-control" name="payment_date" value="${new Date().toISOString().slice(0, 10)}" required>
              </div>
              <div class="form-group">
                <label class="form-label">Tỷ giá thời điểm TT</label>
                <input type="number" step="any" min="0" class="form-control" id="pay-exchange-rate" name="exchange_rate" value="${defaultPayRate}" placeholder="Vd: 1 hoặc 25450.5" style="font-weight: 600;">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label" id="lbl-item-name">${isCpb ? 'Nội dung khoản chi' : 'Diễn giải nội dung đợt thanh toán'}</label>
              <input type="text" class="form-control" name="item_name" placeholder="${isCpb ? 'Vd: Chi tiếp khách bàn giao hồ sơ, Chi tiền nước văn phòng...' : 'Vd: Tạm ứng đợt 1 / Nghiệm thu giai đoạn 2'}" required>
            </div>

            <div class="form-grid-2">
              <div class="form-group" id="grp-actual-vol">
                <label class="form-label">Khối lượng hoàn thành</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="pay-actual" name="actual_volume" value="0">
              </div>
              <div class="form-group" id="grp-accepted-vol">
                <label class="form-label">Khối lượng chấp nhận TT</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="pay-accepted" name="accepted_vol" value="0">
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-request-amount">${isCpb ? 'Số tiền đề nghị chi' : 'Đề nghị thanh toán'}</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="pay-requested" name="request_amount" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-paid-amount">${isCpb ? 'Số tiền thực tế đã chi' : 'Đã thanh toán'}</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="pay-paid" name="paid_amount" value="0" required>
              </div>
            </div>
            <div class="form-group">
              <label class="form-label" id="lbl-invoice-ref">${isCpb ? 'Số Chứng từ / Phiếu chi / Hóa đơn' : 'Số Hóa đơn VAT / Tham chiếu'}</label>
              <input type="text" class="form-control" name="invoice_ref" placeholder="Vd: PC-001 hoặc HD-1234">
            </div>

            <div id="pay-vnd-preview" style="background: #f0f9ff; border: 1px solid #bae6fd; border-radius: var(--radius-sm); padding: 8px 12px; margin-bottom: 12px; font-size: 0.8rem; color: #0369a1; display: ${defaultPayRate !== 1 ? 'flex' : 'none'}; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
              <span><i class="fa-solid fa-coins"></i> Quy đổi VNĐ (× Tỷ giá <strong id="preview-rate-val">${defaultPayRate}</strong>):</span>
              <span>Thực thanh toán: <strong id="preview-paid-vnd" style="font-size: 0.88rem; color: #0f172a;">0 ₫</strong></span>
            </div>

            <div style="font-size: 0.75rem; color: #64748b; margin-bottom: 10px;">
              <i class="fa-solid fa-circle-info"></i> Cho phép nhập <strong>số âm (-)</strong> để ghi nhận dòng tiền trả về / điều chỉnh giảm, và tỷ giá <strong>số thập phân</strong> (vd: 25450.5).
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú diễn giải</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú người nhận tiền, số UNC hoặc mục đích chi..."></textarea>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 14px;">
              <button type="submit" id="btn-submit-payment" class="button button-primary button-sm" style="${isCpb ? 'background: #d97706; border-color: #b45309;' : ''}; font-weight: 600; padding: 6px 16px;">
                ${isCpb ? '<i class="fa-solid fa-receipt"></i> Ghi sổ Chi phí' : '<i class="fa-solid fa-floppy-disk"></i> Ghi sổ Thanh toán'}
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    this.setupCurrencyInputListeners(modal);

    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Bấm ra ngoài vùng modal-container thì đóng
    this.bindModalOutsideClick(modal, closeSubModal);

    // Cập nhật realtime preview quy đổi VNĐ
    const rateInput = modal.querySelector('#pay-exchange-rate');
    const paidInput = modal.querySelector('#pay-paid');
    const vndPreviewBox = modal.querySelector('#pay-vnd-preview');
    const previewRateVal = modal.querySelector('#preview-rate-val');
    const previewPaidVnd = modal.querySelector('#preview-paid-vnd');

    const updatePayVndPreview = () => {
      const rate = parseFloat(rateInput?.value) || 1;
      const paid = this.parseNumberFromDots(paidInput?.value || '0');
      if (vndPreviewBox) {
        vndPreviewBox.style.display = rate !== 1 ? 'flex' : 'none';
      }
      if (previewRateVal) previewRateVal.textContent = rate;
      if (previewPaidVnd) {
        const vnd = Math.round(paid * rate);
        previewPaidVnd.textContent = this.formatCurrency(vnd);
        previewPaidVnd.style.color = vnd < 0 ? 'var(--danger)' : '#0f172a';
      }
    };
    rateInput?.addEventListener('input', updatePayVndPreview);
    paidInput?.addEventListener('input', updatePayVndPreview);

    // Lắng nghe thay đổi dropdown Hợp đồng
    const contractSelect = modal.querySelector('#payment-select-contract');
    contractSelect?.addEventListener('change', (e) => {
      const selectedId = e.target.value;
      const foundContract = this.contracts.find((c) => c.id === selectedId);
      if (foundContract) {
        const isCurrentCpb = foundContract.is_cpb || foundContract.contract_type === 'CPB' || (foundContract.id && foundContract.id.startsWith('B-CPB-'));
        const titleEl = modal.querySelector('#payment-modal-title');
        const submitBtn = modal.querySelector('#btn-submit-payment');
        const actualGrp = modal.querySelector('#grp-actual-vol');
        const acceptedGrp = modal.querySelector('#grp-accepted-vol');
        
        if (titleEl) {
          titleEl.style.color = isCurrentCpb ? '#92400e' : 'inherit';
          titleEl.innerHTML = isCurrentCpb
            ? `<i class="fa-solid fa-receipt" style="color: #d97706;"></i> Nhập đợt Chi phí cho ${foundContract.scope_of_work || foundContract.contract_no}`
            : `Nhập đợt Giải ngân cho ${foundContract.contract_no}`;
        }
        if (submitBtn) {
          submitBtn.style.background = isCurrentCpb ? '#d97706' : '';
          submitBtn.style.borderColor = isCurrentCpb ? '#b45309' : '';
          submitBtn.innerHTML = isCurrentCpb ? '<i class="fa-solid fa-receipt"></i> Ghi sổ Chi phí' : '<i class="fa-solid fa-floppy-disk"></i> Ghi sổ Thanh toán';
        }
        if (rateInput) {
          rateInput.value = Number(foundContract.exchange_rate) || 1;
          updatePayVndPreview();
        }
      }
    });

    modal.querySelector('#form-payment').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const chosenContractId = formData.get('contract_id') || targetContract.id;
      const actualContract = this.contracts.find((c) => c.id === chosenContractId) || targetContract;

      const exRate = parseFloat(formData.get('exchange_rate'));
      const validRate = !isNaN(exRate) && exRate > 0 ? exRate : 1.0;
      const reqAmt = this.parseNumberFromDots(formData.get('request_amount'));
      const paidAmt = this.parseNumberFromDots(formData.get('paid_amount'));

      const paymentData = {
        id: ledgerId,
        is_header: false, // Dòng giao dịch thực tế
        project_id: actualContract.project_id || this.currentProjectId,
        contract_id: actualContract.id,
        parent_id: actualContract.parent_id || actualContract.id,
        payment_no: formData.get('payment_no'),
        item_index: nextIdx,
        item_name: formData.get('item_name'),
        payment_ref: formData.get('item_name'),
        exchange_rate: validRate,
        payment_date: formData.get('payment_date'),
        actual_volume: this.parseNumberFromDots(formData.get('actual_volume')),
        accepted_vol: this.parseNumberFromDots(formData.get('accepted_vol')),
        request_amount: reqAmt,
        paid_amount: paidAmt,
        remaining: reqAmt - paidAmt,
        invoice_ref: formData.get('invoice_ref'),
        note: formData.get('note'),
      };

      await db.put('payments', paymentData);
      await this.loadData();
      closeSubModal();
      this.selectedDisburseContractId = actualContract.id;
      this.render();
      this.renderDisbursementsTab();
      await this.syncToLocalFile();
      if (document.getElementById('contract-detail-modal') && !document.getElementById('contract-detail-modal').classList.contains('hidden')) {
        this.openContractDetail(actualContract.id);
        const tabIpcBtn = document.querySelector('#contract-detail-modal [data-target="m-tab-ipc"]');
        tabIpcBtn?.click();
      }
      this.showToast(isActualCpb ? 'Đã ghi sổ thành công đợt Chi phí Ban!' : 'Đã ghi sổ thành công đợt Giải ngân!', 'success');
    });
  }

  // ==========================================
  // FORM CHỈNH SỬA ĐỢT GIẢI NGÂN (SỬA ĐN) & ĐỢT CHI CPB
  // ==========================================
  openEditPaymentModal(payment, targetContract) {
    const isCpb = targetContract.is_cpb || targetContract.contract_type === 'CPB' || (targetContract.id && targetContract.id.startsWith('B-CPB-'));
    const currentPayRate = this.getPaymentRate(payment);
    const initialPaidVnd = Math.round((Number(payment.paid_amount) || 0) * currentPayRate);
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 700px;">
        <div class="modal-header">
          <h3 style="font-size: 1.05rem; font-weight: 700; color: ${isCpb ? '#92400e' : 'inherit'};">
            <i class="fa-solid fa-pen-to-square" style="color: ${isCpb ? '#d97706' : 'var(--primary)'};"></i>
            ${isCpb ? `Chỉnh sửa Đợt Chi phí: ${this.formatDnNo(payment.payment_no)}` : `Chỉnh sửa Đợt Giải ngân: ${this.formatDnNo(payment.payment_no)}`}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-edit-payment">
            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số Phiếu chi / Đợt chi' : 'Số ĐN thanh toán'} <span style="color: var(--danger);">*</span></label>
                <input type="text" class="form-control" name="payment_no" value="${payment.payment_no || ''}" placeholder="Vd: 01 hoặc PC-01" required>
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Ngày chi tiền' : 'Ngày thanh toán'}</label>
                <input type="date" class="form-control" name="payment_date" value="${this.formatDateForInput(payment.payment_date)}" required>
              </div>
              <div class="form-group">
                <label class="form-label">Tỷ giá thời điểm TT</label>
                <input type="number" step="any" min="0" class="form-control" id="edit-pay-exchange-rate" name="exchange_rate" value="${currentPayRate}" placeholder="Vd: 1 hoặc 25450.5" style="font-weight: 600;">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">${isCpb ? 'Nội dung khoản chi' : 'Diễn giải nội dung đợt thanh toán'}</label>
              <input type="text" class="form-control" name="item_name" value="${payment.item_name || ''}" placeholder="${isCpb ? 'Vd: Chi tiếp khách, chi văn phòng...' : 'Vd: Tạm ứng đợt 1 / Nghiệm thu giai đoạn 2'}" required>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Khối lượng hoàn thành</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="edit-pay-actual" name="actual_volume" value="${this.formatNumberWithDots(payment.actual_volume || 0)}">
              </div>
              <div class="form-group">
                <label class="form-label">Khối lượng chấp nhận TT</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="edit-pay-accepted" name="accepted_vol" value="${this.formatNumberWithDots(payment.accepted_vol || 0)}">
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số tiền đề nghị chi' : 'Đề nghị thanh toán'}</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="edit-pay-requested" name="request_amount" value="${this.formatNumberWithDots(payment.request_amount || 0)}" required>
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số tiền thực tế đã chi' : 'Đã thanh toán'}</label>
                <input type="text" inputmode="decimal" class="form-control currency-input" id="edit-pay-paid" name="paid_amount" value="${this.formatNumberWithDots(payment.paid_amount || 0)}" required>
              </div>
            </div>
            <div class="form-group">
              <label class="form-label">${isCpb ? 'Số Chứng từ / Phiếu chi / Hóa đơn' : 'Số Hóa đơn VAT / Tham chiếu'}</label>
              <input type="text" class="form-control" name="invoice_ref" value="${payment.invoice_ref || ''}" placeholder="Vd: PC-001 hoặc HD-1234">
            </div>

            <div id="edit-pay-vnd-preview" style="background: #f0f9ff; border: 1px solid #bae6fd; border-radius: var(--radius-sm); padding: 8px 12px; margin-bottom: 12px; font-size: 0.8rem; color: #0369a1; display: ${currentPayRate !== 1 ? 'flex' : 'none'}; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
              <span><i class="fa-solid fa-coins"></i> Quy đổi VNĐ (× Tỷ giá <strong id="edit-preview-rate-val">${currentPayRate}</strong>):</span>
              <span>Thực thanh toán: <strong id="edit-preview-paid-vnd" style="font-size: 0.88rem; color: ${initialPaidVnd < 0 ? 'var(--danger)' : '#0f172a'};">${this.formatCurrency(initialPaidVnd)}</strong></span>
            </div>

            <div style="font-size: 0.75rem; color: #64748b; margin-bottom: 10px;">
              <i class="fa-solid fa-circle-info"></i> Cho phép nhập <strong>số âm (-)</strong> để ghi nhận dòng tiền trả về / điều chỉnh giảm, và tỷ giá <strong>số thập phân</strong> (vd: 25450.5).
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú diễn giải</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú người nhận tiền, số UNC...">${payment.note || ''}</textarea>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 16px; border-top: 1px solid var(--border-color); padding-top: 14px;">
              <button type="button" class="button button-sm btn-delete-payment-in-modal" style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; font-weight: 600; padding: 6px 12px;" title="Xóa đợt thanh toán này khỏi hệ thống">
                <i class="fa-solid fa-trash-can"></i> Xóa đợt này
              </button>
              <button type="submit" class="button button-primary button-sm" style="${isCpb ? 'background: #d97706; border-color: #b45309;' : ''}; font-weight: 600; padding: 6px 16px;">
                <i class="fa-solid fa-floppy-disk"></i> Lưu thay đổi
              </button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    this.setupCurrencyInputListeners(modal);

    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Bấm ra ngoài vùng modal-container thì đóng
    this.bindModalOutsideClick(modal, closeSubModal);

    // Nút xóa đợt thanh toán ngay trong modal chỉnh sửa
    modal.querySelector('.btn-delete-payment-in-modal')?.addEventListener('click', () => {
      const dnLabel = this.formatDnNo(payment.payment_no || '');
      const amtLabel = this.formatCurrency(payment.paid_amount || payment.request_amount);
      this.showConfirmDialog({
        title: 'Xác nhận xóa đợt thanh toán?',
        message: `Bạn có chắc chắn muốn xóa đợt thanh toán <strong>"${payment.item_name || dnLabel}"</strong> (${amtLabel}) không?`,
        subMessage: 'Sau khi xóa, dữ liệu sẽ được loại bỏ khỏi sổ theo dõi và tự động đồng bộ ngay lập tức.',
        confirmText: 'Xóa vĩnh viễn',
        cancelText: 'Hủy bỏ',
        confirmType: 'danger',
        onConfirm: async () => {
          await db.delete('payments', payment.id);
          await this.loadData();
          closeSubModal();
          this.render();
          this.renderDisbursementsTab();
          await this.syncToLocalFile();
          if (document.getElementById('contract-detail-modal') && !document.getElementById('contract-detail-modal').classList.contains('hidden')) {
            this.openContractDetail(targetContract.id);
            const tabIpcBtn = document.querySelector('#contract-detail-modal [data-target="m-tab-ipc"]');
            tabIpcBtn?.click();
          }
          this.showToast(`Đã xóa thành công đợt thanh toán ${dnLabel}!`, 'success');
        },
      });
    });

    const editRateInput = modal.querySelector('#edit-pay-exchange-rate');
    const editPaidInput = modal.querySelector('#edit-pay-paid');
    const editVndBox = modal.querySelector('#edit-pay-vnd-preview');
    const editPreviewRate = modal.querySelector('#edit-preview-rate-val');
    const editPreviewPaidVnd = modal.querySelector('#edit-preview-paid-vnd');

    const updateEditVndPreview = () => {
      const rate = parseFloat(editRateInput?.value) || 1;
      const paid = this.parseNumberFromDots(editPaidInput?.value || '0');
      if (editVndBox) editVndBox.style.display = rate !== 1 ? 'flex' : 'none';
      if (editPreviewRate) editPreviewRate.textContent = rate;
      if (editPreviewPaidVnd) {
        const vnd = Math.round(paid * rate);
        editPreviewPaidVnd.textContent = this.formatCurrency(vnd);
        editPreviewPaidVnd.style.color = vnd < 0 ? 'var(--danger)' : '#0f172a';
      }
    };
    editRateInput?.addEventListener('input', updateEditVndPreview);
    editPaidInput?.addEventListener('input', updateEditVndPreview);

    modal.querySelector('#form-edit-payment').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const exRate = parseFloat(formData.get('exchange_rate'));
      const validRate = !isNaN(exRate) && exRate > 0 ? exRate : 1.0;
      const reqAmt = this.parseNumberFromDots(formData.get('request_amount'));
      const paidAmt = this.parseNumberFromDots(formData.get('paid_amount'));

      payment.payment_no = formData.get('payment_no');
      payment.payment_date = formData.get('payment_date');
      payment.item_name = formData.get('item_name');
      payment.payment_ref = formData.get('item_name');
      payment.exchange_rate = validRate;
      payment.actual_volume = this.parseNumberFromDots(formData.get('actual_volume'));
      payment.accepted_vol = this.parseNumberFromDots(formData.get('accepted_vol'));
      payment.request_amount = reqAmt;
      payment.paid_amount = paidAmt;
      payment.remaining = reqAmt - paidAmt;
      payment.invoice_ref = formData.get('invoice_ref');
      payment.note = formData.get('note');

      await db.put('payments', payment);
      await this.loadData();
      closeSubModal();
      this.render();
      this.renderDisbursementsTab();
      await this.syncToLocalFile();
      if (document.getElementById('contract-detail-modal') && !document.getElementById('contract-detail-modal').classList.contains('hidden')) {
        this.openContractDetail(targetContract.id);
        const tabIpcBtn = document.querySelector('#contract-detail-modal [data-target="m-tab-ipc"]');
        tabIpcBtn?.click();
      }
      this.showToast(`Đã cập nhật thành công đợt Giải ngân ${this.formatDnNo(payment.payment_no)}!`, 'success');
    });
  }

  // ==========================================
  // XUẤT BÁO CÁO THEO MẪU CPMS
  // ==========================================
  exportContractReport(contract, isMain = true) {
    const stats = this.calculateTotals(contract);
    const directChildren = this.getDirectChildren(contract.id);
    const realPayments = this.getRealPayments(contract.id);
    const docs = this.getDocuments(contract.id);

    const prj = this.projects.find((p) => p.id === (contract.project_id || this.currentProjectId));
    const prjName = prj ? (prj.name || prj.id) : (contract.project_id || this.currentProjectId);

    let csvContent = '\uFEFF';
    csvContent += `BÁO CÁO QUẢN LÝ ${isMain ? 'HỢP ĐỒNG' : 'PHỤ LỤC'} & GIẢI NGÂN (CPMS)\r\n`;
    csvContent += `Dự án: ${prjName}\r\n`;
    csvContent += `Số hiệu: ${contract.contract_no} - ${contract.scope_of_work}\r\n`;
    csvContent += `Giá trị ký kết: ${stats.totalSigned}\r\n`;
    csvContent += `Tổng thực thanh toán: ${stats.totalPaid} | Công nợ còn lại: ${stats.remaining}\r\n\r\n`;

    if (directChildren.length > 0) {
      csvContent += `--- DANH SÁCH PHỤ LỤC TRỰC THUỘC ---\r\n`;
      csvContent += `Mã PL,Số hiệu PL,Nội dung công việc,Giá trị ký kết,Ngày ký,Trạng thái\r\n`;
      directChildren.forEach((a) => {
        csvContent += `"${a.id}","${a.contract_no}","${(a.scope_of_work || '').replace(/"/g, '""')}",${a.contract_val},"${this.formatDate(a.signing_date)}","${a.status}"\r\n`;
      });
      csvContent += `\r\n`;
    }

    csvContent += `--- BẢNG KÊ CÁC ĐỢT GIẢI NGÂN (ĐỀ NGHỊ THANH TOÁN - ĐN) ---\r\n`;
    csvContent += `ĐN,Nội dung thanh toán,Ngày thanh toán,Giá trị nghiệm thu,Đề nghị thanh toán,Thực trả,Chênh lệch,Hóa đơn VAT,Ghi chú\r\n`;
    realPayments.forEach((p) => {
      const pAcc = this.getPaymentVnd(p, 'accepted_vol');
      const pReq = this.getPaymentVnd(p, 'request_amount');
      const pPaid = this.getPaymentVnd(p, 'paid_amount');
      const pRem = pReq - pPaid;
      csvContent += `"${this.formatDnNo(p.payment_no || p.item_index)}","${(this.getPaymentContent(p) || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${pAcc},${pReq},${pPaid},${pRem},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
    });
    csvContent += `\r\n`;

    csvContent += `--- DANH MỤC HỒ SƠ TÀI LIỆU (DOCUMENT TRACKER) ---\r\n`;
    csvContent += `Mã hồ sơ,Số công văn,Phân loại,Tiêu đề,Ngày nhận,Tình trạng\r\n`;
    docs.forEach((d) => {
      csvContent += `"${d.id}","${d.doc_no || ''}","${d.category || ''}","${(d.title || '').replace(/"/g, '""')}","${this.formatDate(d.received_date)}","${d.status || ''}"\r\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `BaoCao_CPMS_${contract.id}_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);

    const msg = document.getElementById('export-msg');
    if (msg) {
      msg.textContent = 'Đã xuất file báo cáo thành công!';
      setTimeout(() => (msg.textContent = ''), 3000);
    }
  }

  // ==========================================
  // HỆ THỐNG DIALOG & THÔNG BÁO TÙY CHỈNH (CUSTOM CONFIRM & TOAST)
  // ==========================================
  bindModalOutsideClick(modal, closeFn) {
    if (!modal || typeof closeFn !== 'function') return;
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        closeFn();
      }
    });
  }

  showConfirmDialog({
    title = 'Xác nhận hành động',
    message = 'Bạn có chắc chắn muốn thực hiện hành động này không?',
    subMessage = '',
    confirmText = 'Xác nhận',
    cancelText = 'Hủy bỏ',
    confirmType = 'danger',
    onConfirm = null,
  }) {
    const isDanger = confirmType === 'danger';
    const modal = document.createElement('div');
    modal.className = 'modal-overlay custom-confirm-overlay';
    modal.style.zIndex = '999999';
    modal.innerHTML = `
      <div class="modal-container custom-confirm-container" style="max-width: 440px; border-radius: 14px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2), 0 10px 10px -5px rgba(0,0,0,0.1); overflow: hidden; background: #ffffff;">
        <div style="padding: 24px 24px 18px; text-align: center;">
          <div style="width: 52px; height: 52px; border-radius: 50%; background: ${isDanger ? '#fee2e2' : '#e0f2fe'}; color: ${isDanger ? '#dc2626' : '#0284c7'}; display: inline-flex; align-items: center; justify-content: center; font-size: 1.5rem; margin-bottom: 14px;">
            <i class="fa-solid ${isDanger ? 'fa-triangle-exclamation' : 'fa-circle-question'}"></i>
          </div>
          <h4 style="margin: 0 0 8px; font-size: 1.1rem; font-weight: 800; color: #1e293b;">${title}</h4>
          <p style="margin: 0; font-size: 0.88rem; color: #475569; line-height: 1.5;">${message}</p>
          ${subMessage ? `<div style="margin-top: 10px; font-size: 0.78rem; color: #94a3b8; line-height: 1.4;">${subMessage}</div>` : ''}
        </div>
        <div style="padding: 14px 24px; background: #f8fafc; border-top: 1px solid var(--border-color); display: flex; justify-content: flex-end; gap: 10px;">
          <button type="button" class="button button-outline button-sm btn-confirm-cancel" style="padding: 6px 14px;">${cancelText}</button>
          <button type="button" class="button button-sm btn-confirm-ok" style="${isDanger ? 'background: #dc2626; color: #fff; border-color: #b91c1c;' : 'background: var(--primary); color: #fff;'}; font-weight: 600; padding: 6px 16px;">
            ${confirmText}
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeModal = () => modal.remove();

    // Click ra ngoài modal-container thì đóng
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal();
    });

    modal.querySelector('.btn-confirm-cancel')?.addEventListener('click', closeModal);

    modal.querySelector('.btn-confirm-ok')?.addEventListener('click', async () => {
      closeModal();
      if (typeof onConfirm === 'function') {
        await onConfirm();
      }
    });
  }

  /**
   * Hiển thị hộp thoại xác nhận đồng bộ giao diện MAIO và trả về Promise<boolean>
   */
  confirmAsync(optionsOrMessage) {
    const opts =
      typeof optionsOrMessage === 'string'
        ? { message: optionsOrMessage, confirmType: 'primary' }
        : optionsOrMessage || {};
    const {
      title = 'Xác nhận thao tác',
      message = 'Bạn có chắc chắn muốn thực hiện thao tác này không?',
      subMessage = '',
      confirmText = 'Xác nhận',
      cancelText = 'Hủy bỏ',
      confirmType = 'primary',
    } = opts;

    return new Promise((resolve) => {
      const isDanger = confirmType === 'danger';
      const formattedMsg = String(message).replace(/\n/g, '<br>');
      const modal = document.createElement('div');
      modal.className = 'modal-overlay custom-confirm-overlay';
      modal.style.zIndex = '999999';
      modal.innerHTML = `
        <div class="modal-container custom-confirm-container" style="max-width: 440px; border-radius: 14px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2), 0 10px 10px -5px rgba(0,0,0,0.1); overflow: hidden; background: #ffffff;">
          <div style="padding: 24px 24px 18px; text-align: center;">
            <div style="width: 52px; height: 52px; border-radius: 50%; background: ${isDanger ? '#fee2e2' : '#e0f2fe'}; color: ${isDanger ? '#dc2626' : '#0284c7'}; display: inline-flex; align-items: center; justify-content: center; font-size: 1.5rem; margin-bottom: 14px;">
              <i class="fa-solid ${isDanger ? 'fa-triangle-exclamation' : 'fa-circle-question'}"></i>
            </div>
            <h4 style="margin: 0 0 8px; font-size: 1.1rem; font-weight: 800; color: #1e293b;">${title}</h4>
            <p style="margin: 0; font-size: 0.88rem; color: #475569; line-height: 1.5;">${formattedMsg}</p>
            ${subMessage ? `<div style="margin-top: 10px; font-size: 0.78rem; color: #94a3b8; line-height: 1.4;">${subMessage}</div>` : ''}
          </div>
          <div style="padding: 14px 24px; background: #f8fafc; border-top: 1px solid var(--border-color); display: flex; justify-content: flex-end; gap: 10px;">
            <button type="button" class="button button-outline button-sm btn-confirm-cancel" style="padding: 6px 14px;">${cancelText}</button>
            <button type="button" class="button button-sm btn-confirm-ok" style="${isDanger ? 'background: #dc2626; color: #fff; border-color: #b91c1c;' : 'background: var(--primary); color: #fff;'}; font-weight: 600; padding: 6px 16px;">
              ${confirmText}
            </button>
          </div>
        </div>
      `;

      document.body.appendChild(modal);
      let settled = false;
      const finish = (val) => {
        if (settled) return;
        settled = true;
        document.removeEventListener('keydown', onKey);
        modal.remove();
        resolve(val);
      };
      const onKey = (e) => {
        if (e.key === 'Escape') finish(false);
      };
      document.addEventListener('keydown', onKey);
      modal.addEventListener('click', (e) => {
        if (e.target === modal) finish(false);
      });
      modal.querySelector('.btn-confirm-cancel')?.addEventListener('click', () => finish(false));
      modal.querySelector('.btn-confirm-ok')?.addEventListener('click', () => finish(true));
    });
  }

  /**
   * Hiển thị hộp thoại thông báo đồng bộ giao diện MAIO (thay thế window.alert mặc định)
   */
  alertAsync(optionsOrMessage) {
    const opts =
      typeof optionsOrMessage === 'string'
        ? { message: optionsOrMessage }
        : optionsOrMessage || {};

    const rawMsg = String(opts.message ?? '');
    const lower = rawMsg.toLowerCase();

    let autoType = 'info';
    if (lower.includes('lỗi') || lower.includes('không thể') || lower.includes('chưa thể') || lower.includes('cảnh báo')) {
      autoType = 'danger';
    } else if (lower.includes('thành công') || lower.includes('đã sao chép') || lower.includes('đã xóa') || lower.includes('đã lưu') || lower.includes('đã tạo') || lower.includes('đã cập nhật') || lower.includes('đã nạp') || lower.includes('đã xuất') || lower.includes('đã đồng bộ')) {
      autoType = 'success';
    } else if (lower.includes('chưa ') || lower.includes('vui lòng') || lower.includes('lưu ý')) {
      autoType = 'warning';
    }

    const type = opts.type || autoType;
    const defaultTitle =
      type === 'danger' ? 'Thông báo lỗi' : type === 'success' ? 'Thành công' : type === 'warning' ? 'Lưu ý' : 'Thông báo';
    const title = opts.title || defaultTitle;
    const okText = opts.okText || 'Đã hiểu';

    const iconMap = {
      danger: { bg: '#fee2e2', color: '#dc2626', icon: 'fa-circle-exclamation' },
      success: { bg: '#dcfce7', color: '#16a34a', icon: 'fa-circle-check' },
      warning: { bg: '#fef3c7', color: '#d97706', icon: 'fa-triangle-exclamation' },
      info: { bg: '#e0f2fe', color: '#0284c7', icon: 'fa-circle-info' },
    };
    const palette = iconMap[type] || iconMap.info;

    return new Promise((resolve) => {
      const formattedMsg = rawMsg
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');

      const modal = document.createElement('div');
      modal.className = 'modal-overlay custom-confirm-overlay';
      modal.style.zIndex = '999999';
      modal.innerHTML = `
        <div class="modal-container custom-confirm-container" style="max-width: 440px; border-radius: 14px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2), 0 10px 10px -5px rgba(0,0,0,0.1); overflow: hidden; background: #ffffff;">
          <div style="padding: 24px 24px 18px; text-align: center;">
            <div style="width: 52px; height: 52px; border-radius: 50%; background: ${palette.bg}; color: ${palette.color}; display: inline-flex; align-items: center; justify-content: center; font-size: 1.5rem; margin-bottom: 14px;">
              <i class="fa-solid ${palette.icon}"></i>
            </div>
            <h4 style="margin: 0 0 8px; font-size: 1.1rem; font-weight: 800; color: #1e293b;">${title}</h4>
            <p style="margin: 0; font-size: 0.88rem; color: #475569; line-height: 1.5; word-break: break-word;">${formattedMsg}</p>
          </div>
          <div style="padding: 12px 24px; background: #f8fafc; border-top: 1px solid var(--border-color); display: flex; justify-content: center;">
            <button type="button" class="button button-primary button-sm btn-alert-ok" style="min-width: 110px; font-weight: 600; padding: 7px 20px;">
              ${okText}
            </button>
          </div>
        </div>
      `;

      document.body.appendChild(modal);
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        document.removeEventListener('keydown', onKey);
        modal.remove();
        resolve(true);
      };
      const onKey = (e) => {
        if (e.key === 'Escape' || e.key === 'Enter') finish();
      };
      document.addEventListener('keydown', onKey);
      modal.addEventListener('click', (e) => {
        if (e.target === modal) finish();
      });
      modal.querySelector('.btn-alert-ok')?.addEventListener('click', finish);
      modal.querySelector('.btn-alert-ok')?.focus();
    });
  }

  showToast(message, type = 'info') {
    let container = document.getElementById('maio-toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'maio-toast-container';
      container.style.cssText = 'position: fixed; bottom: 24px; right: 24px; z-index: 999999; display: flex; flex-direction: column; gap: 8px; pointer-events: none;';
      document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `maio-toast maio-toast-${type}`;
    const bg = type === 'success' ? '#059669' : (type === 'danger' ? '#dc2626' : '#2563eb');
    const icon = type === 'success' ? 'fa-circle-check' : (type === 'danger' ? 'fa-circle-exclamation' : 'fa-circle-info');
    toast.style.cssText = `background: ${bg}; color: #ffffff; padding: 10px 16px; border-radius: 8px; font-size: 0.85rem; font-weight: 600; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.15); display: flex; align-items: center; gap: 8px; pointer-events: auto; animation: toastSlideIn 0.2s ease-out;`;
    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'all 0.25s ease';
      setTimeout(() => toast.remove(), 260);
    }, 3000);
  }

  // Lọc các đợt thanh toán thực tế của Tab Giải ngân theo các bộ lọc đang kích hoạt
  getFilteredDisbursePayments() {
    const realPayments = this.payments.filter((p) => !p.is_header);
    const searchTerm = (this.disburseContractSearchTerm || '').trim().toLowerCase();
    const dnOnly = Boolean(this.disburseSearchDnOnly);
    const yearFilter = this.disburseFilterYear || '';
    const monthFilter = this.disburseFilterMonth || '';
    const dateFrom = this.disburseFilterDateFrom || '';
    const dateTo = this.disburseFilterDateTo || '';

    const hasDateFilter = Boolean(yearFilter || monthFilter || dateFrom || dateTo);
    const hasDnFilter = Boolean(dnOnly && searchTerm);

    const fromDateObj = dateFrom ? this.parsePaymentDateObj(dateFrom) : null;
    const toDateObj = dateTo ? this.parsePaymentDateObj(dateTo) : null;
    if (fromDateObj) fromDateObj.setHours(0, 0, 0, 0);
    if (toDateObj) toDateObj.setHours(23, 59, 59, 999);

    return realPayments.filter((p) => {
      if (hasDateFilter) {
        const pd = this.parsePaymentDateObj(p.payment_date);
        if (!pd) return false;
        if (yearFilter && pd.getFullYear() !== Number(yearFilter)) return false;
        if (monthFilter && pd.getMonth() + 1 !== Number(monthFilter)) return false;
        if (fromDateObj && pd.getTime() < fromDateObj.getTime()) return false;
        if (toDateObj && pd.getTime() > toDateObj.getTime()) return false;
      }
      if (hasDnFilter) {
        if (!this.matchesDnSearch(p, searchTerm)) return false;
      }
      return true;
    });
  }

  // Render bảng chi tiết các đợt giải ngân (DN / IPC) của Hợp đồng đang chọn (không render lại bảng danh sách HĐ)
  renderDisburseDetailTable(contractId = this.selectedDisburseContractId) {
    const detailContainer = document.getElementById('disburse-detail-container');
    if (!detailContainer) return;

    if (contractId) {
      this.selectedDisburseContractId = contractId;
    }

    const selectedContract = this.contracts.find((c) => c.id === this.selectedDisburseContractId);
    if (!selectedContract) {
      detailContainer.innerHTML = `
        <div style="text-align: center; padding: 40px 20px; color: var(--text-muted);">
          <i class="fa-solid fa-arrow-pointer" style="font-size: 2rem; color: #cbd5e1; margin-bottom: 8px;"></i>
          <p style="margin: 0;">Vui lòng chọn một hợp đồng ở bảng trên để xem chi tiết các đợt giải ngân (DN / IPC).</p>
        </div>
      `;
      return;
    }

    const filteredRealPayments = this.getFilteredDisbursePayments();
    const hasAnyFilter = Boolean(
      this.disburseFilterYear ||
      this.disburseFilterMonth ||
      this.disburseFilterDateFrom ||
      this.disburseFilterDateTo ||
      this.disburseContractSearchTerm ||
      (this.disburseCategoryFilter && this.disburseCategoryFilter !== 'ALL')
    );

    const isSelectedCpb = selectedContract.is_cpb || selectedContract.contract_type === 'CPB' || (selectedContract.id && selectedContract.id.startsWith('B-CPB-'));
    const selectedFamilyIds = this.getAllDescendantIds(selectedContract.id);
    const isSelectedCanceled = this.isContractOrAncestorCanceled(selectedContract);
    const selectedTargetIds = isSelectedCanceled
      ? selectedFamilyIds
      : selectedFamilyIds.filter((fid) => {
          const fc = this.contracts.find((x) => x.id === fid);
          return !fc || !this.isContractOrAncestorCanceled(fc);
        });

    // Lấy các đợt thanh toán của hợp đồng / phụ lục đang chọn (khớp theo bộ lọc hiện tại)
    const contractPayments = filteredRealPayments.filter((p) => selectedTargetIds.includes(p.contract_id));

    // Sắp xếp các đợt giải ngân theo ngày thanh toán hoặc item_index
    contractPayments.sort((a, b) => {
      const dateA = this.parsePaymentDateObj(a.payment_date)?.getTime() || 0;
      const dateB = this.parsePaymentDateObj(b.payment_date)?.getTime() || 0;
      return dateA - dateB;
    });

    // Tính tổng kết Total của hợp đồng này (đã quy đổi VNĐ theo tỷ giá thời điểm thanh toán)
    let dnTotalActual = 0;
    let dnTotalAccepted = 0;
    let dnTotalRequested = 0;
    let dnTotalPaid = 0;
    contractPayments.forEach((p) => {
      dnTotalActual += this.getPaymentVnd(p, 'actual_volume');
      dnTotalAccepted += this.getPaymentVnd(p, 'accepted_vol');
      dnTotalRequested += this.getPaymentVnd(p, 'request_amount');
      dnTotalPaid += this.getPaymentVnd(p, 'paid_amount');
    });
    const dnTotalRemaining = dnTotalRequested - dnTotalPaid;

    const MIN_DETAIL_ROWS = 5;
    let detailRowsHtml = '';
    if (contractPayments.length === 0) {
      detailRowsHtml += `
        <tr class="disburse-empty-row">
          <td style="text-align: center; color: #94a3b8; font-size: 0.74rem;">-</td>
          <td style="text-align: center; color: #94a3b8; font-size: 0.76rem;">-</td>
          <td style="color: #94a3b8; font-style: italic; font-size: 0.78rem; padding: 3px 6px !important;">
            ${hasAnyFilter ? 'Không có đợt thanh toán nào khớp với bộ lọc' : '(Hợp đồng chưa phát sinh đợt thanh toán nào)'}
          </td>
          <td style="text-align: right; color: #cbd5e1;">-</td>
          <td style="text-align: right; color: #cbd5e1;">-</td>
          <td style="text-align: right; color: #cbd5e1;">-</td>
          <td style="text-align: right; color: #cbd5e1;">-</td>
          <td style="text-align: right; color: #cbd5e1;">-</td>
        </tr>
      `;
      for (let r = 1; r < MIN_DETAIL_ROWS; r++) {
        detailRowsHtml += `
          <tr class="disburse-empty-row">
            <td style="text-align: center; color: #cbd5e1; font-size: 0.74rem;">-</td>
            <td style="text-align: center; color: #cbd5e1; font-size: 0.76rem;">-</td>
            <td>&nbsp;</td>
            <td style="text-align: right; color: #cbd5e1;">-</td>
            <td style="text-align: right; color: #cbd5e1;">-</td>
            <td style="text-align: right; color: #cbd5e1;">-</td>
            <td style="text-align: right; color: #cbd5e1;">-</td>
            <td style="text-align: right; color: #cbd5e1;">-</td>
          </tr>
        `;
      }
    } else {
      contractPayments.forEach((p, idx) => {
        const pRate = this.getPaymentRate(p);
        const pActRaw = Number(p.actual_volume) || 0;
        const pAccRaw = Number(p.accepted_vol) || 0;
        const pReqRaw = Number(p.request_amount) || 0;
        const pPaidRaw = Number(p.paid_amount) || 0;
        const pActVnd = Math.round(pActRaw * pRate);
        const pAccVnd = Math.round(pAccRaw * pRate);
        const pReqVnd = Math.round(pReqRaw * pRate);
        const pPaidVnd = Math.round(pPaidRaw * pRate);
        const pRemVnd = pReqVnd - pPaidVnd;
        const rawNo = String(p.payment_no || (idx + 1)).trim();
        const cleanNo = rawNo.replace(/^ĐN[-_\s]*/i, '').replace(/^DN[-_\s]*/i, '') || String(idx + 1);

        const tooltipParts = [];
        if (p.invoice_ref && String(p.invoice_ref).trim()) {
          tooltipParts.push(`Chứng từ / HĐ: ${String(p.invoice_ref).trim()}`);
        }
        if (p.note && String(p.note).trim()) {
          tooltipParts.push(`Ghi chú: ${String(p.note).trim()}`);
        }
        const hasNoteInfo = tooltipParts.length > 0;
        const safeTooltipAttr = tooltipParts
          .join('\n')
          .replace(/&/g, '&amp;')
          .replace(/"/g, '&quot;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');
        const safeItemName = this.getPaymentContent(p)
          .replace(/&/g, '&amp;')
          .replace(/"/g, '&quot;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');

        detailRowsHtml += `
          <tr class="disburse-dn-row${hasNoteInfo ? ' has-dn-note' : ''}" data-pay-id="${p.id}" data-dn-note="${safeTooltipAttr}" title="Bấm để xem & chỉnh sửa đợt thanh toán này${hasNoteInfo ? '\n' + safeTooltipAttr : ''}">
            <td style="font-weight: 700; white-space: nowrap; text-align: center;">
              <span style="background: #e0f2fe; color: #0369a1; padding: 2px 7px; border-radius: 4px; font-size: 0.74rem;">
                ${cleanNo}
              </span>
            </td>
            <td style="white-space: nowrap; font-size: 0.76rem; color: #475569; text-align: center;">
              ${this.formatDate(p.payment_date)}
            </td>
            <td style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${safeItemName}">
              <span style="font-weight: 600; color: #1e293b; font-size: 0.78rem;">${safeItemName}</span>${hasNoteInfo ? ' <i class="fa-regular fa-comment-dots" style="color: #6366f1; margin-left: 3px; font-size: 0.72rem;" title="Rê chuột để xem chứng từ / ghi chú"></i>' : ''}${pRate !== 1 ? ` <span style="font-size: 0.68rem; color: #64748b; font-weight: normal;">(TG: ${pRate})</span>` : ''}
            </td>
            <td style="text-align: right; white-space: nowrap;">
              ${this.formatMoneyPlain(pActVnd)}
            </td>
            <td style="text-align: right; white-space: nowrap;">
              ${this.formatMoneyPlain(pAccVnd)}
            </td>
            <td style="text-align: right; white-space: nowrap; font-weight: 600;">
              ${this.formatMoneyPlain(pReqVnd)}
            </td>
            <td style="text-align: right; white-space: nowrap; font-weight: 700; color: ${pPaidVnd < 0 ? 'var(--danger)' : (this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2')};">
              ${this.formatMoneyPlain(pPaidVnd)}
            </td>
            <td style="text-align: right; white-space: nowrap; color: ${pRemVnd > 0 ? 'var(--danger)' : '#64748b'};">
              ${this.formatMoneyPlain(pRemVnd)}
            </td>
          </tr>
        `;
      });

      // Nếu số dòng thanh toán thực tế ít hơn 5 dòng, render thêm các dòng placeholder để cố định đủ 5 dòng
      if (contractPayments.length < MIN_DETAIL_ROWS) {
        for (let r = contractPayments.length; r < MIN_DETAIL_ROWS; r++) {
          detailRowsHtml += `
            <tr class="disburse-empty-row">
              <td style="text-align: center; color: #cbd5e1; font-size: 0.74rem;">-</td>
              <td style="text-align: center; color: #cbd5e1; font-size: 0.76rem;">-</td>
              <td>&nbsp;</td>
              <td style="text-align: right; color: #cbd5e1;">-</td>
              <td style="text-align: right; color: #cbd5e1;">-</td>
              <td style="text-align: right; color: #cbd5e1;">-</td>
              <td style="text-align: right; color: #cbd5e1;">-</td>
              <td style="text-align: right; color: #cbd5e1;">-</td>
            </tr>
          `;
        }
      }
    }

    const partnerLabel = this.disburseGroup === 'A' ? (selectedContract.company_a || 'Chủ đầu tư') : (selectedContract.company_b || 'Ban QLDA / Thầu phụ');
    const partnerRole = this.disburseGroup === 'A' ? 'Bên A' : 'Bên B';
    const selStats = this.calculateTotals(selectedContract);
    const selRate = Number(selectedContract.exchange_rate) || 1.0;
    const selValVnd = selStats.totalSigned;

    detailContainer.innerHTML = `
      <div class="disburse-detail-header">
        <div style="width: 100%; overflow: hidden;">
          <div style="display: flex; align-items: center; gap: 8px; flex-wrap: nowrap; overflow: hidden;">
            <span class="badge ${this.disburseGroup === 'A' ? 'badge-primary' : 'badge-warning'}" style="font-size: 0.73rem; padding: 1.5px 6px; flex-shrink: 0;">
              <i class="fa-solid ${this.disburseGroup === 'A' ? 'fa-arrow-down-long' : 'fa-arrow-up-long'}"></i> 
              ${this.disburseGroup === 'A' ? 'BẢNG THEO DÕI THU TIỀN (NHÓM A)' : 'BẢNG THEO DÕI CHI TIỀN (NHÓM B)'}
            </span>
            <strong style="font-size: 0.88rem; color: #1e293b; white-space: nowrap; flex-shrink: 0;">
              ${selectedContract.contract_no || selectedContract.id}
            </strong>
            <span style="color: var(--text-muted); font-size: 0.77rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0;" title="${String(selectedContract.scope_of_work || 'Hợp đồng thực hiện dự án').replace(/"/g, '&quot;')}">— ${selectedContract.scope_of_work || 'Hợp đồng thực hiện dự án'}</span>
          </div>
          <div style="font-size: 0.74rem; color: var(--text-muted); margin-top: 2px; display: flex; gap: 12px; flex-wrap: nowrap; overflow: hidden; text-overflow: ellipsis;">
            <span style="white-space: nowrap; flex-shrink: 0;"><i class="fa-solid fa-building"></i> ${partnerRole}: <strong>${partnerLabel}</strong></span>
            <span style="white-space: nowrap; flex-shrink: 0;"><i class="fa-solid fa-file-contract"></i> Giá trị HĐ: <strong>${this.formatMoneyPlain(selValVnd)}</strong>${selRate !== 1 ? ` <small style="font-weight: normal; color: #64748b;">(${this.formatNumberWithDots(selectedContract.contract_val)} nguyên tệ × TG ${selRate})</small>` : ''}</span>
            <span style="white-space: nowrap; flex-shrink: 0;"><i class="fa-solid fa-circle-check" style="color: var(--success);"></i> Đã ${this.disburseGroup === 'A' ? 'thu' : 'thanh toán'}: <strong style="color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatMoneyPlain(dnTotalPaid)}</strong></span>
            <span style="white-space: nowrap; flex-shrink: 0;"><i class="fa-solid fa-clock-rotate-left"></i> Còn lại: <strong style="color: ${dnTotalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(dnTotalRemaining)}</strong></span>
          </div>
        </div>
      </div>

      <div class="disburse-scroll-container disburse-scroll-detail-5">
        <table class="disburse-compact-table disburse-summary-table disburse-detail-table" style="min-width: 1040px; width: 100%;">
          <thead>
            <tr style="background: #f8fafc;">
              <th style="width: 60px; min-width: 60px; padding: 4px 4px !important; text-align: center !important;">Đợt TT</th>
              <th style="width: 85px; min-width: 85px; padding: 4px 4px !important; text-align: center !important; white-space: nowrap;">Ngày TT</th>
              <th style="width: 170px; max-width: 210px; min-width: 120px; padding: 4px 6px !important; text-align: center !important;">Nội dung thanh toán</th>
              <th style="width: 135px; min-width: 130px; padding: 4px 6px !important; text-align: center !important; white-space: nowrap;">KL hoàn thành</th>
              <th style="width: 135px; min-width: 130px; padding: 4px 6px !important; text-align: center !important; white-space: nowrap;">KL chấp nhận TT</th>
              <th style="width: 135px; min-width: 130px; padding: 4px 6px !important; text-align: center !important; white-space: nowrap;">Đề nghị TT</th>
              <th style="width: 135px; min-width: 130px; padding: 4px 6px !important; text-align: center !important; white-space: nowrap;">Đã ${this.disburseGroup === 'A' ? 'thu' : 'thanh toán'}</th>
              <th style="width: 125px; min-width: 120px; padding: 4px 6px !important; text-align: center !important; white-space: nowrap;">Còn lại</th>
            </tr>
          </thead>
          <tbody>
            ${detailRowsHtml}
          </tbody>
          <tfoot>
            <tr style="background: #f8fafc; font-weight: 700;">
              <td colspan="3" style="padding: 4px 6px; color: #1e293b; text-align: left; font-size: 0.77rem;">
                TỔNG CỘNG (${contractPayments.length} đợt ${this.disburseGroup === 'A' ? 'thu' : 'chi'})
              </td>
              <td style="text-align: right; padding: 4px 6px; font-size: 0.77rem;">${this.formatMoneyPlain(dnTotalActual)}</td>
              <td style="text-align: right; padding: 4px 6px; font-size: 0.77rem;">${this.formatMoneyPlain(dnTotalAccepted)}</td>
              <td style="text-align: right; padding: 4px 6px; font-size: 0.77rem;">${this.formatMoneyPlain(dnTotalRequested)}</td>
              <td style="text-align: right; padding: 4px 6px; font-size: 0.77rem; color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatMoneyPlain(dnTotalPaid)}</td>
              <td style="text-align: right; padding: 4px 6px; font-size: 0.77rem; color: ${dnTotalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(dnTotalRemaining)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    // Tooltip nổi tức thì khi rê chuột vào dòng ĐN có chứng từ / ghi chú
    let noteTooltip = document.getElementById('maio-dn-note-tooltip');
    if (!noteTooltip) {
      noteTooltip = document.createElement('div');
      noteTooltip.id = 'maio-dn-note-tooltip';
      noteTooltip.style.cssText = 'position: fixed; z-index: 99999; max-width: 420px; background: #1e293b; color: #f8fafc; padding: 8px 12px; border-radius: 6px; font-size: 0.78rem; line-height: 1.45; box-shadow: 0 10px 25px rgba(15, 23, 42, 0.25); pointer-events: none; display: none; white-space: pre-wrap; word-break: break-word;';
      document.body.appendChild(noteTooltip);
    }
    detailContainer.querySelectorAll('.disburse-dn-row.has-dn-note').forEach((row) => {
      row.addEventListener('mouseenter', (e) => {
        const text = row.getAttribute('data-dn-note');
        if (!text) return;
        noteTooltip.textContent = text;
        noteTooltip.style.display = 'block';
        const x = Math.min(window.innerWidth - noteTooltip.offsetWidth - 16, e.clientX + 14);
        const y = Math.min(window.innerHeight - noteTooltip.offsetHeight - 16, e.clientY + 14);
        noteTooltip.style.left = `${Math.max(12, x)}px`;
        noteTooltip.style.top = `${Math.max(12, y)}px`;
      });
      row.addEventListener('mousemove', (e) => {
        if (noteTooltip.style.display !== 'block') return;
        const x = Math.min(window.innerWidth - noteTooltip.offsetWidth - 16, e.clientX + 14);
        const y = Math.min(window.innerHeight - noteTooltip.offsetHeight - 16, e.clientY + 14);
        noteTooltip.style.left = `${Math.max(12, x)}px`;
        noteTooltip.style.top = `${Math.max(12, y)}px`;
      });
      row.addEventListener('mouseleave', () => {
        noteTooltip.style.display = 'none';
      });
    });

    // Bấm vào bất kỳ dòng đợt thanh toán nào để mở modal chỉnh sửa (kèm tùy chọn xóa)
    detailContainer.querySelectorAll('.disburse-dn-row').forEach((row) => {
      row.addEventListener('click', () => {
        const payId = row.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (pay) {
          const c = this.contracts.find((item) => item.id === pay.contract_id) || selectedContract;
          this.openEditPaymentModal(pay, c);
        }
      });
    });
  }

  // ==========================================
  // SỔ THEO DÕI GIẢI NGÂN TOÀN DỰ ÁN (DISBURSEMENTS TAB & FINANCIAL SUMMARY)
  // ==========================================
  renderDisbursementsTab() {
    this.updateEmptyStateUI('disbursements');
    if (!this.isProjectReady()) return;

    const summaryContainer = document.getElementById('disburse-summary-container');
    const contractsTableContainer = document.getElementById('disburse-contracts-table-container');
    const detailContainer = document.getElementById('disburse-detail-container');
    if (!summaryContainer || !contractsTableContainer || !detailContainer) return;

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || {};
    const prjName = curPrj.name || curPrj.id || 'Dự án';

    // 1. TÍNH TOÁN BỘ LỌC & BẢNG TỔNG HỢP SUMMARY (FINANCIAL SUMMARY)
    const realPayments = this.payments.filter((p) => !p.is_header);

    // Populate dropdown Năm thanh toán tự động từ dữ liệu thực tế
    const yearSelect = document.getElementById('disburse-filter-year');
    if (yearSelect) {
      const yearsSet = new Set();
      realPayments.forEach((p) => {
        const d = this.parsePaymentDateObj(p.payment_date);
        if (d) yearsSet.add(d.getFullYear());
      });
      const sortedYears = Array.from(yearsSet).sort((a, b) => b - a);
      const curYearVal = this.disburseFilterYear || '';
      yearSelect.innerHTML = `<option value="">Tất cả năm</option>` + sortedYears.map((y) => `<option value="${y}">${y}</option>`).join('');
      yearSelect.value = sortedYears.includes(Number(curYearVal)) ? String(curYearVal) : '';
      if (yearSelect.value !== curYearVal) this.disburseFilterYear = yearSelect.value;
    }

    const catFilter = this.disburseCategoryFilter || 'ALL';
    const searchTerm = (this.disburseContractSearchTerm || '').trim().toLowerCase();
    const dnOnly = Boolean(this.disburseSearchDnOnly);
    const yearFilter = this.disburseFilterYear || '';
    const monthFilter = this.disburseFilterMonth || '';
    const dateFrom = this.disburseFilterDateFrom || '';
    const dateTo = this.disburseFilterDateTo || '';

    const hasDateFilter = Boolean(yearFilter || monthFilter || dateFrom || dateTo);
    const hasDnFilter = Boolean(dnOnly && searchTerm);
    const hasContractTextFilter = Boolean(!dnOnly && searchTerm);
    const hasCatFilter = catFilter !== 'ALL';
    const hasAnyFilter = hasDateFilter || hasDnFilter || hasContractTextFilter || hasCatFilter;

    // Luôn hiển thị nút Reset; làm nổi bật màu đỏ nhẹ khi đang có bộ lọc hoạt động
    const resetFiltersBtn = document.getElementById('btn-reset-disburse-filters');
    if (resetFiltersBtn) {
      resetFiltersBtn.classList.remove('hidden');
      resetFiltersBtn.style.color = hasAnyFilter ? 'var(--danger)' : '';
      resetFiltersBtn.style.borderColor = hasAnyFilter ? '#fca5a5' : '';
      resetFiltersBtn.style.background = hasAnyFilter ? '#fef2f2' : '';
    }

    // Lọc các đợt thanh toán theo Ngày / Tháng / Năm và Checkbox ĐN
    const filteredRealPayments = this.getFilteredDisbursePayments();

    // Helper kiểm tra 1 hợp đồng có thỏa mãn bộ lọc Loại HĐ / Tìm kiếm / ĐN / Ngày không
    const matchedPaymentContractIds = new Set(filteredRealPayments.map((p) => p.contract_id));
    const matchesContractFilters = (c) => {
      if (hasCatFilter) {
        const effCat = this.getEffectiveCategory(c);
        if (catFilter === 'OTHER') {
          if (effCat) return false;
        } else if (effCat.toLowerCase() !== catFilter.toLowerCase()) {
          return false;
        }
      }
      if (hasContractTextFilter) {
        const cNo = String(c.contract_no || c.id || '').toLowerCase();
        const cDisplay = String(c.display_no || '').toLowerCase();
        const cScope = String(c.scope_of_work || '').toLowerCase();
        const cPartA = String(c.company_a || '').toLowerCase();
        const cPartB = String(c.company_b || '').toLowerCase();
        if (!(cNo.includes(searchTerm) || cDisplay.includes(searchTerm) || cScope.includes(searchTerm) || cPartA.includes(searchTerm) || cPartB.includes(searchTerm))) {
          return false;
        }
      }
      if (hasDnFilter || hasDateFilter) {
        const famIds = this.getAllDescendantIds(c.id);
        const hasMatchPay = famIds.some((fid) => matchedPaymentContractIds.has(fid));
        if (!hasMatchPay) return false;
      }
      return true;
    };

    // Hàm tính tổng hợp 1 nhóm (A hoặc B) theo bộ lọc hiện tại (loại trừ hoàn toàn hợp đồng Cancel)
    const calcGroupSummary = (groupCode) => {
      const groupActiveContracts = this.contracts.filter((c) => c.group === groupCode && !this.isContractOrAncestorCanceled(c));
      const matchedContracts = hasAnyFilter
        ? groupActiveContracts.filter((c) => matchesContractFilters(c))
        : groupActiveContracts;

      // Tập ID hợp đồng khớp (tránh cộng trùng cha-con khi tính Giá trị HĐ)
      const matchedIds = new Set(matchedContracts.map((c) => c.id));
      const topMatched = matchedContracts.filter((c) => !c.parent_id || !matchedIds.has(c.parent_id));

      const summary = {
        contractVal: 0,
        actualVol: 0,
        acceptedVol: 0,
        requestedAmt: 0,
        paidAmt: 0,
        remaining: 0,
      };
      topMatched.forEach((c) => {
        const st = this.calculateTotals(c);
        summary.contractVal += st.totalSigned;
      });

      // Tập ID dùng để gom thanh toán (bao gồm toàn bộ con cháu không Cancel của topMatched)
      const payScopeIds = new Set();
      topMatched.forEach((c) => {
        this.getAllDescendantIds(c.id).forEach((fid) => {
          const fc = this.contracts.find((x) => x.id === fid);
          if (!fc || !this.isContractOrAncestorCanceled(fc)) {
            payScopeIds.add(fid);
          }
        });
      });

      filteredRealPayments.forEach((p) => {
        if (payScopeIds.has(p.contract_id)) {
          summary.actualVol += this.getPaymentVnd(p, 'actual_volume');
          summary.acceptedVol += this.getPaymentVnd(p, 'accepted_vol');
          summary.requestedAmt += this.getPaymentVnd(p, 'request_amount');
          summary.paidAmt += this.getPaymentVnd(p, 'paid_amount');
        }
      });
      summary.remaining = summary.requestedAmt - summary.paidAmt;
      return summary;
    };

    const summaryA = calcGroupSummary('A');
    const summaryB = calcGroupSummary('B');

    // Cân đối dòng tiền A - B
    const balance = summaryA.paidAmt - summaryB.paidAmt;

    // 1.1 RENDER BẢNG SUMMARY CHUẨN CPMS PLUS (Bỏ đơn vị đ, tổng hợp theo bộ lọc)
    summaryContainer.innerHTML = `
      <div class="disburse-summary-header">
        <div style="font-weight: 800; font-size: 0.9rem; color: #1e293b; display: flex; align-items: center; gap: 8px;">
          <i class="fa-solid fa-chart-pie" style="color: var(--primary);"></i>
          <span>BẢNG TỔNG HỢP TÀI CHÍNH & GIẢI NGÂN</span>
          ${hasAnyFilter ? `<span class="badge badge-primary" style="font-size: 0.72rem; padding: 2px 7px;"><i class="fa-solid fa-filter"></i> Đang lọc</span>` : ''}
        </div>
        <span style="font-size: 0.76rem; color: var(--text-muted); font-weight: 500;">Đơn vị: VNĐ</span>
      </div>
      <div style="overflow-x: auto;">
        <table class="disburse-summary-table">
          <thead>
            <tr>
              <th style="width: 110px; text-align: center;">Phân nhóm</th>
              <th style="text-align: center;">Giá trị hợp đồng</th>
              <th style="text-align: center;">Khối lượng hoàn thành</th>
              <th style="text-align: center;">Khối lượng chấp nhận TT</th>
              <th style="text-align: center;">Đề nghị thanh toán</th>
              <th style="text-align: center;">Đã thanh toán</th>
              <th style="text-align: center;">Còn lại</th>
            </tr>
          </thead>
          <tbody>
            <tr style="background: #f8fafc; font-weight: 600;">
              <td style="text-align: center;">
                <span class="badge badge-primary" style="font-size: 0.76rem; padding: 3px 8px; display: inline-flex; align-items: center; gap: 5px;"><i class="fa-solid fa-landmark"></i> Nhóm A</span>
              </td>
              <td style="text-align: right; font-weight: 700; color: var(--primary);">${this.formatMoneyPlain(summaryA.contractVal)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryA.actualVol)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryA.acceptedVol)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryA.requestedAmt)}</td>
              <td style="text-align: right; font-weight: 700; color: var(--success);">${this.formatMoneyPlain(summaryA.paidAmt)}</td>
              <td style="text-align: right; color: ${summaryA.remaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(summaryA.remaining)}</td>
            </tr>

            <tr style="background: #ffffff; font-weight: 600;">
              <td style="text-align: center;">
                <span class="badge badge-warning" style="font-size: 0.76rem; padding: 3px 8px; display: inline-flex; align-items: center; gap: 5px;"><i class="fa-solid fa-handshake"></i> Nhóm B</span>
              </td>
              <td style="text-align: right; font-weight: 700; color: #4f46e5;">${this.formatMoneyPlain(summaryB.contractVal)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryB.actualVol)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryB.acceptedVol)}</td>
              <td style="text-align: right;">${this.formatMoneyPlain(summaryB.requestedAmt)}</td>
              <td style="text-align: right; font-weight: 700; color: #0891b2;">${this.formatMoneyPlain(summaryB.paidAmt)}</td>
              <td style="text-align: right; color: ${summaryB.remaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(summaryB.remaining)}</td>
            </tr>

            <tr style="background: #f0fdf4; border-top: 2px solid #86efac; font-weight: 700;">
              <td colspan="5" style="color: #166534; font-size: 0.85rem;">
                <i class="fa-solid fa-scale-balanced" style="color: #16a34a; margin-right: 6px;"></i> A - B
              </td>
              <td style="text-align: right; font-size: 0.92rem; color: ${balance >= 0 ? 'var(--success)' : 'var(--danger)'};">
                ${balance >= 0 ? '+' : ''}${this.formatMoneyPlain(balance)}
              </td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;

    // 2. CẬP NHẬT THANH TOOLBAR & TAB BUTTONS
    const groupBtns = document.querySelectorAll('.disburse-group-btn');
    groupBtns.forEach((btn) => {
      const g = btn.getAttribute('data-disburse-group');
      if (g === this.disburseGroup) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    const btnDisburseAddCpb = document.getElementById('btn-disburse-add-cpb');
    if (btnDisburseAddCpb) {
      if (this.disburseGroup === 'B') {
        btnDisburseAddCpb.classList.remove('hidden');
      } else {
        btnDisburseAddCpb.classList.add('hidden');
      }
    }

    // Cập nhật dropdown bộ lọc Loại hợp đồng cho Tab Giải ngân
    this.updateDisburseCategoryFilterUI();

    // 3. LẤY DANH SÁCH HỢP ĐỒNG THUỘC NHÓM ĐANG CHỌN (KÈM CPB TRONG NHÓM B)
    const curGroupContracts = this.contracts.filter((c) => c.group === this.disburseGroup);
    
    // Lọc theo Loại hợp đồng, Từ khóa, Checkbox ĐN và Ngày/Tháng/Năm
    const filteredContracts = curGroupContracts.filter((c) => matchesContractFilters(c));

    // 3.1 XÂY DỰNG DANH SÁCH CÓ THỨ BẬC CÂY & ĐÁNH SỐ THỨ TỰ (STT) RIÊNG CHO PHỤ LỤC CỦA TỪNG HĐ
    const isCpbFn = (c) => Boolean(c.is_cpb || c.contract_type === 'CPB' || (c.id && String(c.id).startsWith('B-CPB-')));

    const sortByDate = (a, b) => {
      const isA = isCpbFn(a);
      const isB = isCpbFn(b);
      if (isA !== isB) return isA ? 1 : -1;

      const dateA = this.parsePaymentDateObj(a.signing_date || a.created_at);
      const dateB = this.parsePaymentDateObj(b.signing_date || b.created_at);
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;

      if (timeA && timeB) {
        if (timeA !== timeB) return timeA - timeB;
      } else if (timeA && !timeB) {
        return -1;
      } else if (!timeA && timeB) {
        return 1;
      }

      const codeA = String(a.display_no || a.id || '').trim();
      const codeB = String(b.display_no || b.id || '').trim();
      return codeA.localeCompare(codeB, 'vi', { numeric: true, sensitivity: 'base' });
    };

    // Phân loại HĐ chính (root), phụ lục mồ côi (khi filter chỉ khớp con mà không khớp cha) và CPB
    const filteredIdSet = new Set(filteredContracts.map((c) => c.id));
    const rootContracts = [];
    const orphanAppendices = [];
    const cpbList = [];

    filteredContracts.forEach((c) => {
      if (isCpbFn(c)) {
        cpbList.push(c);
      } else if (!c.parent_id) {
        rootContracts.push(c);
      } else {
        if (!filteredIdSet.has(c.parent_id)) {
          orphanAppendices.push(c);
        }
      }
    });

    rootContracts.sort(sortByDate);
    cpbList.sort(sortByDate);
    orphanAppendices.sort(sortByDate);

    // Xây dựng danh sách hiển thị có cấu trúc thứ bậc cây và gán STT (HĐ: 1, 2, ... 11; PL: 11.1, 11.2; HĐ tiếp: 12)
    const displayContracts = [];
    const sttMap = new Map();

    const appendSubTree = (parentId, parentStt) => {
      const children = filteredContracts.filter((c) => c.parent_id === parentId);
      children.sort(sortByDate);
      children.forEach((child, chIdx) => {
        const childStt = `${parentStt}.${chIdx + 1}`;
        sttMap.set(child.id, childStt);
        displayContracts.push(child);
        appendSubTree(child.id, childStt);
      });
    };

    rootContracts.forEach((root, rIdx) => {
      const rootStt = String(rIdx + 1);
      sttMap.set(root.id, rootStt);
      displayContracts.push(root);
      appendSubTree(root.id, rootStt);
    });

    // Các phụ lục mồ côi (khi tìm kiếm chỉ khớp phụ lục, cha bị ẩn)
    orphanAppendices.forEach((orph, oIdx) => {
      if (!displayContracts.some((x) => x.id === orph.id)) {
        const parentContract = this.contracts.find((x) => x.id === orph.parent_id);
        const parentDisplay = parentContract ? (parentContract.contract_no || parentContract.display_no || 'PL') : 'PL';
        const orphStt = `${parentDisplay}.${oIdx + 1}`;
        sttMap.set(orph.id, orphStt);
        displayContracts.push(orph);
        appendSubTree(orph.id, orphStt);
      }
    });

    // Cuối cùng là các khoản Chi phí Ban (CPB)
    cpbList.forEach((cpb) => {
      sttMap.set(cpb.id, 'CPB');
      displayContracts.push(cpb);
    });

    // Xác định mục đang chọn (nếu chưa có hoặc không nằm trong nhóm thì chọn mục đầu tiên)
    const existsSelected = displayContracts.some((c) => c.id === this.selectedDisburseContractId);
    if (!existsSelected && displayContracts.length > 0) {
      // Ưu tiên chọn mục đã có thanh toán
      const firstWithPay = displayContracts.find((c) => realPayments.some((p) => p.contract_id === c.id));
      this.selectedDisburseContractId = firstWithPay ? firstWithPay.id : displayContracts[0].id;
    } else if (displayContracts.length === 0) {
      this.selectedDisburseContractId = null;
    }

    // 3.2 RENDER BẢNG DANH SÁCH CÁC HỢP ĐỒNG CỦA NHÓM ĐANG CHỌN
    if (displayContracts.length === 0) {
      contractsTableContainer.innerHTML = `
        <div style="text-align: center; padding: 36px 20px; color: var(--text-muted);">
          <i class="fa-solid fa-folder-open" style="font-size: 2rem; color: #cbd5e1; margin-bottom: 8px;"></i>
          <p style="margin: 0; font-size: 0.85rem;">Không tìm thấy hợp đồng nào phù hợp trong Nhóm ${this.disburseGroup}.</p>
        </div>
      `;
    } else {
      let contractRowsHtml = '';
      displayContracts.forEach((c) => {
        const isSelected = c.id === this.selectedDisburseContractId;
        const isCpb = isCpbFn(c);
        const isAppendix = Boolean(c.parent_id);
        const itemStt = sttMap.get(c.id) || '';
        
        // Tính tổng hợp tài chính của mục này (bao gồm cả các gói con / phụ lục con trực thuộc, loại trừ gói con Cancel)
        // Lưu ý: chỉ dùng để hiển thị trên dòng này, KHÔNG cộng lặp vào bảng Tổng hợp Tài chính phía trên
        const familyIds = this.getAllDescendantIds(c.id);
        const isSelfOrParentCanceled = this.isContractOrAncestorCanceled(c);
        const targetFamilyIds = isSelfOrParentCanceled
          ? familyIds
          : familyIds.filter((fid) => {
              const fc = this.contracts.find((x) => x.id === fid);
              return !fc || !this.isContractOrAncestorCanceled(fc);
            });
        const cPayments = filteredRealPayments.filter((p) => targetFamilyIds.includes(p.contract_id));
        
        let cActual = 0;
        let cAccepted = 0;
        let cRequested = 0;
        let cPaid = 0;
        cPayments.forEach((p) => {
          cActual += this.getPaymentVnd(p, 'actual_volume');
          cAccepted += this.getPaymentVnd(p, 'accepted_vol');
          cRequested += this.getPaymentVnd(p, 'request_amount');
          cPaid += this.getPaymentVnd(p, 'paid_amount');
        });

        const st = this.calculateTotals(c);
        const exRate = Number(c.exchange_rate) || 1.0;
        const cVal = st.totalSigned;
        const cRemaining = cRequested - cPaid;
        const cRate = cVal > 0 ? ((cPaid / cVal) * 100).toFixed(1) : (cRequested > 0 ? ((cPaid / cRequested) * 100).toFixed(1) : 0);

        const partnerName = this.disburseGroup === 'A' ? (c.company_a || 'Chủ đầu tư') : (c.company_b || 'Ban QLDA / Thầu phụ');
        const isCanceledRow = this.isContractOrAncestorCanceled(c);
        const isCompletedRow = !isCanceledRow && this.isContractCompleted(c);
        const rowStatusClass = isCanceledRow
          ? ' row-status-cancel'
          : isCompletedRow
          ? ' row-status-completed'
          : '';
        const statusMiniBadge = (isCanceledRow || isCompletedRow)
          ? this.getContractStatusBadgeHtml(c.status, true)
          : '';

        contractRowsHtml += `
          <tr class="disburse-contract-row ${isSelected ? 'selected-contract-row' : ''}${rowStatusClass}" data-disburse-cid="${c.id}" title="${c.display_no ? `[${c.display_no}] ` : ''}${(c.scope_of_work || 'Hợp đồng thực hiện dự án')} | ${partnerName} (${c.status || 'Active'}) — Nhấp đúp vào ô STT để mở popup HĐ, hoặc nhấp đúp vào dòng để mở Tổng thể bảng thu chi">
            <td class="col-disburse-stt" style="white-space: nowrap; font-weight: 700; padding: 3px 6px; text-align: center; color: #475569; font-size: 0.77rem; cursor: pointer; user-select: none;" title="Nhấp đúp để mở popup chi tiết Hợp đồng / Phụ lục này (giống bên tab Hợp đồng)">
              ${isCpb 
                ? `<span class="badge" style="background: #fef3c7; color: #b45309; border: 1px solid #fde68a; font-size: 0.68rem; padding: 1.5px 4px;"><i class="fa-solid fa-receipt"></i> CPB</span>` 
                : `<span style="${isAppendix ? 'color: #0369a1; font-weight: 600;' : ''}">${itemStt}</span>`}
            </td>
            <td style="font-weight: 700; color: #1e293b; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem; ${isAppendix ? 'padding-left: 18px !important;' : ''}">
              ${isAppendix ? '<span style="color: #94a3b8; font-weight: normal; margin-right: 4px;">↳</span>' : ''}<span>${c.contract_no || c.id}</span>${statusMiniBadge}
            </td>
            <td style="text-align: right; font-weight: 700; color: #1e293b; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cVal)}
              ${exRate !== 1 ? `<div style="font-size: 0.66rem; color: #64748b; font-weight: normal;">${this.formatNumberWithDots(c.contract_val)} nt</div>` : ''}
            </td>
            <td style="text-align: right; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cActual)}
            </td>
            <td style="text-align: right; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cAccepted)}
            </td>
            <td style="text-align: right; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cRequested)}
            </td>
            <td style="text-align: right; font-weight: 700; color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'}; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cPaid)}
            </td>
            <td style="text-align: right; color: ${cRemaining > 0 ? 'var(--danger)' : '#64748b'}; padding: 3px 6px; white-space: nowrap; font-size: 0.77rem;">
              ${this.formatMoneyPlain(cRemaining)}
            </td>
            <td style="text-align: center; padding: 3px 6px; white-space: nowrap;">
              <span class="badge ${Number(cRate) >= 100 ? 'badge-success' : 'badge-primary'}" style="font-size: 0.68rem; padding: 1.5px 5px;">
                ${cRate}%
              </span>
            </td>
          </tr>
        `;
      });

      const mainContractCount = displayContracts.filter((c) => !c.parent_id && !isCpbFn(c)).length;
      const appendixCount = displayContracts.filter((c) => Boolean(c.parent_id)).length;
      const cpbCount = displayContracts.filter((c) => isCpbFn(c)).length;

      const countParts = [];
      countParts.push(`<strong>${mainContractCount}</strong> hợp đồng`);
      if (appendixCount > 0) {
        countParts.push(`<strong>${appendixCount}</strong> phụ lục`);
      }
      if (cpbCount > 0) {
        countParts.push(`<strong>${cpbCount}</strong> CPB`);
      }
      const summaryCountHtml = countParts.join(' • ');

      contractsTableContainer.innerHTML = `
        <div class="disburse-contracts-card-header">
          <span style="font-weight: 700; color: #1e293b;">
            <i class="fa-solid fa-list-check" style="color: var(--primary); margin-right: 4px;"></i> DANH SÁCH HỢP ĐỒNG & GIÁ TRỊ (NHÓM ${this.disburseGroup})
          </span>
          <span style="color: var(--text-muted); font-size: 0.75rem; font-weight: 600;">
            ${summaryCountHtml}
          </span>
        </div>
        <div class="disburse-scroll-container disburse-scroll-contract-list">
          <table class="disburse-compact-table disburse-contract-list-table" style="width: 100%; border-collapse: separate !important; border-spacing: 0 !important;">
            <thead>
              <tr>
                <th style="width: 50px; padding: 4px 6px !important; text-align: center !important;">STT</th>
                <th style="min-width: 140px; padding: 4px 6px !important; text-align: center !important;">Số hiệu HĐ / Phụ lục</th>
                <th style="width: 136px; padding: 4px 6px !important; text-align: center !important;">Giá trị HĐ</th>
                <th style="width: 136px; padding: 4px 6px !important; text-align: center !important;">KL hoàn thành</th>
                <th style="width: 136px; padding: 4px 6px !important; text-align: center !important;">KL chấp nhận TT</th>
                <th style="width: 136px; padding: 4px 6px !important; text-align: center !important;">Đề nghị TT</th>
                <th style="width: 136px; padding: 4px 6px !important; text-align: center !important;">Đã thanh toán</th>
                <th style="width: 126px; padding: 4px 6px !important; text-align: center !important;">Còn lại</th>
                <th style="width: 64px; padding: 4px 6px !important; text-align: center !important;">Tỷ lệ (%)</th>
              </tr>
            </thead>
            <tbody>
              ${contractRowsHtml}
            </tbody>
          </table>
        </div>
      `;

      // Gắn sự kiện click chọn dòng hợp đồng: KHÔNG load lại bảng HĐ mà CHỈ đổi class và render bảng chi tiết
      contractsTableContainer.querySelectorAll('.disburse-contract-row').forEach((row) => {
        row.addEventListener('click', () => {
          const cid = row.getAttribute('data-disburse-cid');
          if (this.selectedDisburseContractId === cid) return;
          this.selectedDisburseContractId = cid;

          // Cập nhật class active trực tiếp trên DOM
          contractsTableContainer.querySelectorAll('.disburse-contract-row').forEach((r) => {
            r.classList.remove('selected-contract-row');
          });
          row.classList.add('selected-contract-row');

          // Chỉ nạp lại bảng theo dõi thu chi bên dưới
          this.renderDisburseDetailTable(cid);
        });

        // Click đúp (dblclick) vào cột thứ tự (ô STT) để mở popup xem chi tiết hợp đồng (giống bên tab Hợp đồng)
        const sttCell = row.querySelector('.col-disburse-stt');
        sttCell?.addEventListener('dblclick', (e) => {
          e.stopPropagation();
          const cid = row.getAttribute('data-disburse-cid');
          if (cid) {
            this.openContractDetail(cid);
          }
        });

        // Click đúp (dblclick) vào hợp đồng (ngoài ô STT) để mở popup tổng thể bảng thu chi
        row.addEventListener('dblclick', () => {
          const cid = row.getAttribute('data-disburse-cid');
          const c = this.contracts.find((item) => item.id === cid);
          if (c) {
            this.openContractDisbursementsOverviewModal(c);
          }
        });
      });
    }

    // 4. RENDER BẢNG CHI TIẾT CÁC ĐỢT GIẢI NGÂN (DN / IPC) CỦA HỢP ĐỒNG ĐANG CHỌN
    this.renderDisburseDetailTable();

    // Lắng nghe sự kiện Toolbar giải ngân 1 lần duy nhất
    if (!this.disburseEventsBound) {
      this.disburseEventsBound = true;

      // Chuyển tab Nhóm A Thu / Nhóm B Chi
      document.querySelectorAll('.disburse-group-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          this.disburseGroup = btn.getAttribute('data-disburse-group');
          this.selectedDisburseContractId = null;
          this.disburseCategoryFilter = 'ALL';
          this.renderDisbursementsTab();
        });
      });

      // Nút Thêm Chi phí Ban (CPB) trong tab Giải ngân
      document.getElementById('btn-disburse-add-cpb')?.addEventListener('click', () => {
        this.openCreateCPBModal();
      });

      // Nút Thêm Thanh toán mới trong tab Giải ngân
      document.getElementById('btn-disburse-add-payment')?.addEventListener('click', () => {
        const curContract = this.contracts.find((c) => c.id === this.selectedDisburseContractId);
        this.openPaymentEntryModal(curContract);
      });

      // Tìm kiếm hợp đồng giải ngân & nút Clear Search
      const disburseSearchInput = document.getElementById('disburse-contract-search');
      const clearDisburseSearchBtn = document.getElementById('btn-clear-disburse-search');
      disburseSearchInput?.addEventListener('input', (e) => {
        this.disburseContractSearchTerm = e.target.value.trim().toLowerCase();
        if (clearDisburseSearchBtn) {
          clearDisburseSearchBtn.classList.toggle('hidden', !e.target.value);
        }
        this.renderDisbursementsTab();
      });

      clearDisburseSearchBtn?.addEventListener('click', () => {
        if (disburseSearchInput) {
          disburseSearchInput.value = '';
          disburseSearchInput.focus();
        }
        this.disburseContractSearchTerm = '';
        clearDisburseSearchBtn.classList.add('hidden');
        this.renderDisbursementsTab();
      });

      // Lọc theo Loại hợp đồng trong tab Giải ngân
      document.getElementById('disburse-category-filter')?.addEventListener('change', (e) => {
        this.disburseCategoryFilter = e.target.value;
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      // Checkbox Chỉ tìm theo số ĐN
      document.getElementById('disburse-search-dn-only')?.addEventListener('change', (e) => {
        this.disburseSearchDnOnly = e.target.checked;
        if (disburseSearchInput) {
          disburseSearchInput.placeholder = this.disburseSearchDnOnly
            ? 'Nhập số ĐN (vd: 14, ĐN-14, 05)...'
            : 'Tìm HĐ, đối tác hoặc ĐN...';
        }
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      // Lọc theo Năm / Tháng / Khoảng ngày thanh toán
      const yearEl = document.getElementById('disburse-filter-year');
      const monthEl = document.getElementById('disburse-filter-month');
      const dateFromEl = document.getElementById('disburse-filter-date-from');
      const dateToEl = document.getElementById('disburse-filter-date-to');

      yearEl?.addEventListener('change', (e) => {
        this.disburseFilterYear = e.target.value;
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      monthEl?.addEventListener('change', (e) => {
        this.disburseFilterMonth = e.target.value;
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      dateFromEl?.addEventListener('change', (e) => {
        this.disburseFilterDateFrom = e.target.value;
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      dateToEl?.addEventListener('change', (e) => {
        this.disburseFilterDateTo = e.target.value;
        this.selectedDisburseContractId = null;
        this.renderDisbursementsTab();
      });

      // Nút Xóa tất cả bộ lọc giải ngân
      document.getElementById('btn-reset-disburse-filters')?.addEventListener('click', () => {
        this.disburseContractSearchTerm = '';
        this.disburseSearchDnOnly = false;
        this.disburseCategoryFilter = 'ALL';
        this.disburseFilterYear = '';
        this.disburseFilterMonth = '';
        this.disburseFilterDateFrom = '';
        this.disburseFilterDateTo = '';
        this.selectedDisburseContractId = null;

        if (disburseSearchInput) {
          disburseSearchInput.value = '';
          disburseSearchInput.placeholder = 'Tìm HĐ, đối tác hoặc ĐN...';
        }
        clearDisburseSearchBtn?.classList.add('hidden');
        const dnChk = document.getElementById('disburse-search-dn-only');
        if (dnChk) dnChk.checked = false;
        if (yearEl) yearEl.value = '';
        if (monthEl) monthEl.value = '';
        if (dateFromEl) dateFromEl.value = '';
        if (dateToEl) dateToEl.value = '';
        const catEl = document.getElementById('disburse-category-filter');
        if (catEl) catEl.value = 'ALL';

        this.renderDisbursementsTab();
      });

      // Icon Lịch bên trong 2 ô ngày tháng: Click mở bộ chọn ngày, nhấp đúp hoặc chuột phải mở popup Lịch tháng
      document.querySelectorAll('.disburse-date-input-wrap .disburse-date-icon').forEach((icon) => {
        icon.addEventListener('click', () => {
          const input = icon.parentElement?.querySelector('input.disburse-filter-date');
          if (input) {
            try {
              if (typeof input.showPicker === 'function') {
                input.showPicker();
              } else {
                input.focus();
              }
            } catch (err) {
              input.focus();
            }
          }
        });
        icon.addEventListener('dblclick', (e) => {
          e.stopPropagation();
          this.openDisburseMonthCalendarModal();
        });
        icon.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          this.openDisburseMonthCalendarModal();
        });
      });

      // Xuất Excel giải ngân
      document.getElementById('btn-export-disbursements-excel')?.addEventListener('click', () => {
        this.exportAllDisbursementsToExcel();
      });
    }
  }

  exportAllDisbursementsToExcel() {
    const realPayments = this.payments.filter((p) => !p.is_header);
    const prj = this.projects.find((p) => p.id === this.currentProjectId);
    const prjName = prj ? (prj.name || prj.id) : (this.currentProjectId || 'DuAn');

    const contractMap = new Map();
    this.contracts.forEach((c) => contractMap.set(c.id, c));

    let csvContent = '\uFEFF';
    csvContent += `SỔ THEO DÕI GIẢI NGÂN & THANH TOÁN (PAYMENT LEDGER / IPC)\r\n`;
    csvContent += `Dự án: ${prjName}\r\n`;
    csvContent += `Thời điểm xuất: ${new Date().toLocaleString('vi-VN')}\r\n\r\n`;
    csvContent += `STT,Số ĐN,Nhóm,Mã HĐ,Số hiệu HĐ,Đối tác,Nội dung thanh toán,Ngày TT,Đề nghị TT (VNĐ),Thực trả (VNĐ),Hóa đơn VAT,Ghi chú\r\n`;

    realPayments.forEach((p, idx) => {
      const c = contractMap.get(p.contract_id);
      const isGroupA = c ? c.group === 'A' : (String(p.contract_id).startsWith('A-') || String(p.id).startsWith('PAY-A-'));
      const groupName = isGroupA ? 'Nhóm A (Chủ đầu tư)' : 'Nhóm B (Thầu phụ / NCC)';
      const cNo = c ? (c.contract_no || '') : '';
      const partner = c ? (isGroupA ? (c.company_a || '') : (c.company_b || '')) : '';

      const pReqVnd = this.getPaymentVnd(p, 'request_amount');
      const pPaidVnd = this.getPaymentVnd(p, 'paid_amount');
      csvContent += `"${idx + 1}","${this.formatDnNo(p.payment_no || p.item_index)}","${groupName}","${p.contract_id || ''}","${cNo.replace(/"/g, '""')}","${partner.replace(/"/g, '""')}","${(this.getPaymentContent(p) || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${pReqVnd},${pPaidVnd},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `So_Giai_Ngan_${prjName}_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  openDisburseMonthCalendarModal() {
    let activeYear = parseInt(this.disburseFilterYear, 10) || new Date().getFullYear();

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.style.zIndex = '10005';

    const renderCalendarContent = () => {
      const curGroupContracts = this.contracts.filter((c) => {
        if (this.disburseGroup === 'A') return c.group === 'A' || c.id?.startsWith('A-');
        return c.group === 'B' || c.id?.startsWith('B-') || c.is_cpb || c.contract_type === 'CPB';
      });
      const validCids = new Set(curGroupContracts.map((c) => c.id));

      const monthStats = {};
      for (let m = 1; m <= 12; m++) {
        monthStats[m] = { count: 0, amount: 0 };
      }

      this.payments.forEach((p) => {
        if (p.is_header) return;
        if (!validCids.has(p.contract_id)) return;
        const d = this.parsePaymentDateObj(p.payment_date);
        if (!d) return;
        if (d.getFullYear() === activeYear) {
          const m = d.getMonth() + 1;
          if (monthStats[m]) {
            monthStats[m].count += 1;
            monthStats[m].amount += this.getPaymentVnd(p, 'paid_amount');
          }
        }
      });

      let cellsHtml = '';
      for (let m = 1; m <= 12; m++) {
        const stat = monthStats[m];
        const isCurrentActive = String(this.disburseFilterYear) === String(activeYear) && String(this.disburseFilterMonth) === String(m);
        const hasData = stat.count > 0;

        cellsHtml += `
          <div class="month-cell ${isCurrentActive ? 'active-month' : ''} ${hasData ? 'has-payments' : ''}" data-month="${m}">
            <div class="month-title">Tháng ${m}</div>
            ${hasData ? `
              <div class="month-stat-count">${stat.count} đợt TT</div>
              <div class="month-stat-amount">${this.formatMoneyPlain(stat.amount)}</div>
            ` : `
              <div style="font-size: 0.7rem; color: #94a3b8;">Chưa có TT</div>
            `}
          </div>
        `;
      }

      modal.innerHTML = `
        <div class="modal-container" style="max-width: 480px; width: 92vw; background: #ffffff; border-radius: 12px; box-shadow: 0 20px 30px rgba(0,0,0,0.2); overflow: hidden; cursor: default;">
          <div class="modal-header">
            <div style="display: flex; align-items: center; gap: 8px;">
              <i class="fa-solid fa-calendar-days" style="color: #2563eb; font-size: 1.1rem;"></i>
              <h3 class="modal-title" style="font-size: 1.05rem; margin: 0;">Lịch Tháng Giải Ngân</h3>
            </div>
            <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
          </div>

          <div class="modal-body" style="padding: 16px;">
            <!-- Year Navigator -->
            <div class="month-calendar-header">
              <button type="button" class="button button-outline button-sm btn-prev-year" style="padding: 3px 8px;">
                <i class="fa-solid fa-chevron-left"></i>
              </button>
              <div style="font-weight: 700; font-size: 1rem; color: #1e293b; display: flex; align-items: center; gap: 6px;">
                <span>Năm</span>
                <span class="active-year-label" style="color: #2563eb; font-size: 1.15rem;">${activeYear}</span>
              </div>
              <button type="button" class="button button-outline button-sm btn-next-year" style="padding: 3px 8px;">
                <i class="fa-solid fa-chevron-right"></i>
              </button>
            </div>

            <!-- 12 Months Grid -->
            <div class="month-calendar-grid">
              ${cellsHtml}
            </div>

            <!-- Action buttons -->
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 8px; border-top: 1px solid var(--border-color); padding-top: 12px; margin-top: 4px;">
              <button type="button" class="button button-outline button-sm btn-clear-all-month-filter" style="font-size: 0.78rem;">
                <i class="fa-solid fa-rotate-left"></i> Bỏ lọc thời gian
              </button>
              <button type="button" class="button button-primary button-sm btn-select-full-year" style="font-size: 0.78rem;">
                <i class="fa-solid fa-calendar-check"></i> Chọn cả năm ${activeYear}
              </button>
            </div>
          </div>
        </div>
      `;

      // Gắn sự kiện
      modal.querySelector('.btn-close-submodal').addEventListener('click', closeModal);
      this.bindModalOutsideClick(modal, closeModal);

      modal.querySelector('.btn-prev-year').addEventListener('click', () => {
        activeYear -= 1;
        renderCalendarContent();
      });

      modal.querySelector('.btn-next-year').addEventListener('click', () => {
        activeYear += 1;
        renderCalendarContent();
      });

      modal.querySelectorAll('.month-cell').forEach((cell) => {
        cell.addEventListener('click', () => {
          const selectedMonth = cell.getAttribute('data-month');
          this.disburseFilterYear = String(activeYear);
          this.disburseFilterMonth = String(selectedMonth);
          this.disburseFilterDateFrom = '';
          this.disburseFilterDateTo = '';
          this.selectedDisburseContractId = null;

          const yEl = document.getElementById('disburse-filter-year');
          const mEl = document.getElementById('disburse-filter-month');
          const dfEl = document.getElementById('disburse-filter-date-from');
          const dtEl = document.getElementById('disburse-filter-date-to');
          if (yEl) yEl.value = this.disburseFilterYear;
          if (mEl) mEl.value = this.disburseFilterMonth;
          if (dfEl) dfEl.value = '';
          if (dtEl) dtEl.value = '';

          closeModal();
          this.renderDisbursementsTab();
        });
      });

      modal.querySelector('.btn-select-full-year').addEventListener('click', () => {
        this.disburseFilterYear = String(activeYear);
        this.disburseFilterMonth = '';
        this.disburseFilterDateFrom = '';
        this.disburseFilterDateTo = '';
        this.selectedDisburseContractId = null;

        const yEl = document.getElementById('disburse-filter-year');
        const mEl = document.getElementById('disburse-filter-month');
        const dfEl = document.getElementById('disburse-filter-date-from');
        const dtEl = document.getElementById('disburse-filter-date-to');
        if (yEl) yEl.value = this.disburseFilterYear;
        if (mEl) mEl.value = '';
        if (dfEl) dfEl.value = '';
        if (dtEl) dtEl.value = '';

        closeModal();
        this.renderDisbursementsTab();
      });

      modal.querySelector('.btn-clear-all-month-filter').addEventListener('click', () => {
        this.disburseFilterYear = '';
        this.disburseFilterMonth = '';
        this.disburseFilterDateFrom = '';
        this.disburseFilterDateTo = '';
        this.selectedDisburseContractId = null;

        const yEl = document.getElementById('disburse-filter-year');
        const mEl = document.getElementById('disburse-filter-month');
        const dfEl = document.getElementById('disburse-filter-date-from');
        const dtEl = document.getElementById('disburse-filter-date-to');
        if (yEl) yEl.value = '';
        if (mEl) mEl.value = '';
        if (dfEl) dfEl.value = '';
        if (dtEl) dtEl.value = '';

        closeModal();
        this.renderDisbursementsTab();
      });
    };

    const closeModal = () => {
      document.removeEventListener('keydown', handleEsc);
      modal.remove();
    };

    const handleEsc = (e) => {
      if (e.key === 'Escape') closeModal();
    };
    document.addEventListener('keydown', handleEsc);

    renderCalendarContent();
    document.body.appendChild(modal);
  }

  openContractDisbursementsOverviewModal(contract) {
    if (!contract) return;
    const isCpb = contract.is_cpb || contract.contract_type === 'CPB' || (contract.id && contract.id.startsWith('B-CPB-'));
    const isGroupA = contract.group === 'A' || contract.id?.startsWith('A-');
    const partnerRole = isGroupA ? 'Chủ đầu tư (Bên A)' : 'Đối tác / Thầu phụ (Bên B)';
    const partnerName = isGroupA ? (contract.company_a || 'Chủ đầu tư') : (contract.company_b || 'Ban QLDA / Thầu phụ');

    const familyIds = this.getAllDescendantIds(contract.id);
    const contractPayments = this.payments.filter((p) => !p.is_header && familyIds.includes(p.contract_id));

    contractPayments.sort((a, b) => {
      const dateA = this.parsePaymentDateObj(a.payment_date)?.getTime() || 0;
      const dateB = this.parsePaymentDateObj(b.payment_date)?.getTime() || 0;
      return dateA - dateB;
    });

    const cStats = this.calculateTotals(contract);
    const cRate = Number(contract.exchange_rate) || 1.0;
    const cValVnd = cStats.totalSigned;

    let totalAct = 0;
    let totalAcc = 0;
    let totalReq = 0;
    let totalPaid = 0;
    contractPayments.forEach((p) => {
      totalAct += this.getPaymentVnd(p, 'actual_volume');
      totalAcc += this.getPaymentVnd(p, 'accepted_vol');
      totalReq += this.getPaymentVnd(p, 'request_amount');
      totalPaid += this.getPaymentVnd(p, 'paid_amount');
    });
    const totalRemaining = totalReq - totalPaid;
    const disbursePercent = cValVnd > 0 ? ((totalPaid / cValVnd) * 100).toFixed(1) : '0.0';

    let rowsHtml = '';
    if (contractPayments.length === 0) {
      rowsHtml = `
        <tr>
          <td colspan="10" style="text-align: center; padding: 30px; color: #94a3b8;">
            <i class="fa-regular fa-folder-open" style="font-size: 1.8rem; margin-bottom: 8px; display: block;"></i>
            Hợp đồng này chưa có đợt thanh toán / giải ngân nào.
          </td>
        </tr>
      `;
    } else {
      contractPayments.forEach((p, idx) => {
        const pRate = this.getPaymentRate(p);
        const pActRaw = Number(p.actual_volume) || 0;
        const pAccRaw = Number(p.accepted_vol) || 0;
        const pReqRaw = Number(p.request_amount) || 0;
        const pPaidRaw = Number(p.paid_amount) || 0;
        const pActVnd = Math.round(pActRaw * pRate);
        const pAccVnd = Math.round(pAccRaw * pRate);
        const pReqVnd = Math.round(pReqRaw * pRate);
        const pPaidVnd = Math.round(pPaidRaw * pRate);
        const pRemVnd = pReqVnd - pPaidVnd;

        const rawNo = String(p.payment_no || (idx + 1)).trim();
        const cleanNo = rawNo.replace(/^ĐN[-_\s]*/i, '').replace(/^DN[-_\s]*/i, '') || String(idx + 1);

        const safeItemName = this.getPaymentContent(p)
          .replace(/&/g, '&amp;')
          .replace(/"/g, '&quot;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');

        rowsHtml += `
          <tr class="disburse-dn-row" data-pay-id="${p.id}" style="cursor: pointer;" title="Nhấp để chỉnh sửa đợt thanh toán này">
            <td style="text-align: center; color: #64748b; font-size: 0.76rem;">${idx + 1}</td>
            <td style="text-align: center; font-weight: 700;">
              <span style="background: #e0f2fe; color: #0369a1; padding: 2px 7px; border-radius: 4px; font-size: 0.74rem;">
                ${cleanNo}
              </span>
            </td>
            <td style="text-align: center; font-size: 0.76rem; color: #475569; white-space: nowrap;">
              ${this.formatDate(p.payment_date)}
            </td>
            <td style="font-weight: 600; color: #1e293b;" title="${safeItemName}">
              ${safeItemName}
              ${pRate !== 1 ? `<span style="font-size: 0.68rem; color: #64748b; font-weight: normal;">(TG: ${pRate})</span>` : ''}
              ${p.invoice_ref ? `<div style="font-size: 0.68rem; color: #6366f1; font-weight: normal;"><i class="fa-solid fa-receipt"></i> HĐ: ${p.invoice_ref}</div>` : ''}
            </td>
            <td style="text-align: right; white-space: nowrap;">${this.formatMoneyPlain(pActVnd)}</td>
            <td style="text-align: right; white-space: nowrap;">${this.formatMoneyPlain(pAccVnd)}</td>
            <td style="text-align: right; white-space: nowrap; font-weight: 600;">${this.formatMoneyPlain(pReqVnd)}</td>
            <td style="text-align: right; white-space: nowrap; font-weight: 700; color: ${pPaidVnd < 0 ? 'var(--danger)' : (isGroupA ? 'var(--success)' : '#0891b2')};">
              ${this.formatMoneyPlain(pPaidVnd)}
            </td>
            <td style="text-align: right; white-space: nowrap; color: ${pRemVnd > 0 ? 'var(--danger)' : '#64748b'};">
              ${this.formatMoneyPlain(pRemVnd)}
            </td>
            <td style="text-align: center; white-space: nowrap;">
              <button class="button button-outline btn-edit-dn-in-overview" data-pay-id="${p.id}" title="Sửa đợt này" style="padding: 2px 6px; font-size: 0.72rem; margin-right: 4px;">
                <i class="fa-solid fa-pen"></i>
              </button>
              <button class="button button-outline btn-delete-dn-in-overview" data-pay-id="${p.id}" title="Xóa đợt này" style="padding: 2px 6px; font-size: 0.72rem; color: var(--danger); border-color: #fca5a5;">
                <i class="fa-solid fa-trash-can"></i>
              </button>
            </td>
          </tr>
        `;
      });
    }

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.style.zIndex = '10005';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 1100px; width: 95vw; background: #ffffff; border-radius: 12px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); overflow: hidden; cursor: default; display: flex; flex-direction: column; height: 88vh; max-height: 88vh;">
        <div class="modal-header" style="flex-shrink: 0; background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border-bottom: 1px solid #e2e8f0; padding: 12px 18px;">
          <div style="display: flex; align-items: center; gap: 8px;">
            <i class="fa-solid fa-table-list" style="color: var(--primary); font-size: 1.1rem;"></i>
            <h3 class="modal-title" style="font-size: 1.05rem; margin: 0; color: #0f172a;">
              Tổng thể Bảng thu chi: <span style="color: var(--primary);">${contract.contract_no || contract.id}</span>
              ${contract.scope_of_work ? `<small style="font-size: 0.82rem; color: var(--text-muted); font-weight: normal; margin-left: 6px;">— ${contract.scope_of_work}</small>` : ''}
            </h3>
          </div>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <div class="modal-body" style="padding: 16px 20px; background: #ffffff; flex: 1 1 0; min-height: 0; display: flex; flex-direction: column; overflow: hidden;">
          <!-- Banner đối tác & Thông tin Hợp đồng -->
          <div style="flex-shrink: 0; background: #f8fafc; padding: 8px 12px; border-radius: 6px; border: 1px solid #e2e8f0; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
            <div style="display: flex; align-items: center; gap: 8px; min-width: 0;">
              <i class="fa-solid fa-building" style="color: #64748b; font-size: 0.95rem;"></i>
              <span style="font-size: 0.78rem; color: #64748b; font-weight: 600;">${partnerRole}:</span>
              <strong style="font-size: 0.88rem; color: #0f172a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${partnerName}">${partnerName}</strong>
            </div>
            <div style="display: flex; align-items: center; gap: 12px; font-size: 0.78rem; color: #64748b;">
              ${contract.signing_date ? `<span><i class="fa-regular fa-calendar" style="margin-right: 4px;"></i> Ngày ký: <strong style="color: #334155;">${this.formatDate(contract.signing_date)}</strong></span>` : ''}
              <span class="badge ${contract.status === 'Completed' ? 'badge-success' : (contract.status === 'Cancel' ? 'badge-danger' : 'badge-primary')}" style="font-size: 0.72rem; padding: 2px 7px;">
                ${contract.status || 'Active'}
              </span>
            </div>
          </div>

          <!-- Lưới 7 thẻ chỉ số KPI tài chính đồng đều -->
          <div class="overview-modal-kpi-grid" style="flex-shrink: 0;">
            <div class="overview-kpi-card">
              <span class="overview-kpi-label">Giá trị HĐ</span>
              <span class="overview-kpi-value">${this.formatMoneyPlain(cValVnd)}</span>
            </div>
            <div class="overview-kpi-card">
              <span class="overview-kpi-label">KL hoàn thành</span>
              <span class="overview-kpi-value">${this.formatMoneyPlain(totalAct)}</span>
            </div>
            <div class="overview-kpi-card">
              <span class="overview-kpi-label">KL chấp nhận TT</span>
              <span class="overview-kpi-value">${this.formatMoneyPlain(totalAcc)}</span>
            </div>
            <div class="overview-kpi-card">
              <span class="overview-kpi-label">Đề nghị TT</span>
              <span class="overview-kpi-value">${this.formatMoneyPlain(totalReq)}</span>
            </div>
            <div class="overview-kpi-card" style="background: #f0fdf4; border-color: #bbf7d0;">
              <span class="overview-kpi-label" style="color: #166534;">Đã ${isGroupA ? 'thu' : 'thanh toán'}</span>
              <span class="overview-kpi-value" style="color: ${isGroupA ? '#15803d' : '#0891b2'};">${this.formatMoneyPlain(totalPaid)}</span>
            </div>
            <div class="overview-kpi-card" style="background: ${totalRemaining > 0 ? '#fef2f2' : '#f8fafc'}; border-color: ${totalRemaining > 0 ? '#fecaca' : '#e2e8f0'};">
              <span class="overview-kpi-label" style="color: ${totalRemaining > 0 ? '#b91c1c' : '#64748b'};">Còn lại</span>
              <span class="overview-kpi-value" style="color: ${totalRemaining > 0 ? '#dc2626' : '#64748b'};">${this.formatMoneyPlain(totalRemaining)}</span>
            </div>
            <div class="overview-kpi-card" style="background: #eff6ff; border-color: #bfdbfe;">
              <span class="overview-kpi-label" style="color: #1e40af;">Tỷ lệ giải ngân</span>
              <span class="overview-kpi-value" style="color: #2563eb;">${disbursePercent}%</span>
            </div>
          </div>

          <!-- Toolbar bên trên bảng -->
          <div style="flex-shrink: 0; display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <div style="font-size: 0.85rem; font-weight: 700; color: #1e293b;">
              <i class="fa-solid fa-list-check" style="color: #3b82f6; margin-right: 4px;"></i> Chi tiết tất cả các đợt giải ngân (${contractPayments.length} đợt)
            </div>
            <div style="display: flex; gap: 8px;">
              <button type="button" class="button button-outline button-sm btn-export-contract-ledger" style="font-size: 0.78rem;">
                <i class="fa-solid fa-file-excel" style="color: #16a34a;"></i> Xuất Excel HĐ này
              </button>
              <button type="button" class="button button-primary button-sm btn-add-dn-in-overview" style="font-size: 0.78rem;">
                <i class="fa-solid fa-plus"></i> Thêm đợt TT
              </button>
            </div>
          </div>

          <!-- Bảng cuộn chi tiết -->
          <div class="overview-detail-table-wrap" style="flex: 1 1 0; min-height: 0; max-height: none; overflow-y: auto;">
            <table class="overview-detail-table">
              <thead>
                <tr>
                  <th style="width: 44px; text-align: center;">STT</th>
                  <th style="width: 76px; text-align: center;">Đợt TT</th>
                  <th style="width: 96px; text-align: center;">Ngày TT</th>
                  <th style="min-width: 150px; text-align: center;">Nội dung thanh toán</th>
                  <th style="width: 110px; text-align: center;">KL hoàn thành</th>
                  <th style="width: 110px; text-align: center;">KL chấp nhận TT</th>
                  <th style="width: 115px; text-align: center;">Đề nghị TT</th>
                  <th style="width: 115px; text-align: center;">Đã ${isGroupA ? 'thu' : 'thanh toán'}</th>
                  <th style="width: 105px; text-align: center;">Còn lại</th>
                  <th style="width: 76px; text-align: center;">Thao tác</th>
                </tr>
              </thead>
              <tbody>
                ${rowsHtml}
              </tbody>
              <tfoot>
                <tr style="background: #f8fafc; font-weight: 700;">
                  <td colspan="4" style="padding: 6px 8px; text-align: left; font-size: 0.78rem; color: #1e293b;">
                    TỔNG CỘNG (${contractPayments.length} đợt)
                  </td>
                  <td style="text-align: right; padding: 6px 8px;">${this.formatMoneyPlain(totalAct)}</td>
                  <td style="text-align: right; padding: 6px 8px;">${this.formatMoneyPlain(totalAcc)}</td>
                  <td style="text-align: right; padding: 6px 8px;">${this.formatMoneyPlain(totalReq)}</td>
                  <td style="text-align: right; padding: 6px 8px; color: ${isGroupA ? 'var(--success)' : '#0891b2'};">${this.formatMoneyPlain(totalPaid)}</td>
                  <td style="text-align: right; padding: 6px 8px; color: ${totalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatMoneyPlain(totalRemaining)}</td>
                  <td></td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeModal = () => {
      document.removeEventListener('keydown', handleEsc);
      modal.remove();
    };

    const handleEsc = (e) => {
      if (e.key === 'Escape') closeModal();
    };
    document.addEventListener('keydown', handleEsc);

    modal.querySelector('.btn-close-submodal').addEventListener('click', closeModal);
    this.bindModalOutsideClick(modal, closeModal);

    // Nút Thêm đợt TT trong overview modal
    modal.querySelector('.btn-add-dn-in-overview')?.addEventListener('click', () => {
      closeModal();
      this.openPaymentEntryModal(contract);
    });

    // Nút Xuất Excel riêng cho hợp đồng này
    modal.querySelector('.btn-export-contract-ledger')?.addEventListener('click', () => {
      this.exportSingleContractDisbursementsToExcel(contract, contractPayments);
    });

    // Click vào dòng hoặc nút sửa đợt thanh toán
    modal.querySelectorAll('.disburse-dn-row').forEach((row) => {
      row.addEventListener('click', (e) => {
        if (e.target.closest('.btn-delete-dn-in-overview') || e.target.closest('.btn-edit-dn-in-overview')) return;
        const payId = row.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (pay) {
          closeModal();
          this.openEditPaymentModal(pay, contract);
        }
      });
    });

    modal.querySelectorAll('.btn-edit-dn-in-overview').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const payId = btn.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (pay) {
          closeModal();
          this.openEditPaymentModal(pay, contract);
        }
      });
    });

    // Nút xóa đợt thanh toán
    modal.querySelectorAll('.btn-delete-dn-in-overview').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const payId = btn.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (!pay) return;
        const dnLabel = this.formatDnNo(pay.payment_no || '');
        const amtLabel = this.formatCurrency(pay.paid_amount || pay.request_amount);
        this.showConfirmDialog({
          title: 'Xác nhận xóa đợt thanh toán?',
          message: `Bạn có chắc chắn muốn xóa đợt thanh toán <strong>"${pay.item_name || dnLabel}"</strong> (${amtLabel}) không?`,
          subMessage: 'Sau khi xóa, dữ liệu sẽ được loại bỏ khỏi sổ theo dõi và tự động đồng bộ ngay lập tức.',
          confirmText: 'Xóa đợt thanh toán',
          confirmVariant: 'danger',
          onConfirm: async () => {
            this.payments = this.payments.filter((p) => p.id !== pay.id);
            await this.saveData();
            closeModal();
            this.showToast('Đã xóa đợt thanh toán thành công!', 'success');
            this.renderDisbursementsTab();
          }
        });
      });
    });
  }

  exportSingleContractDisbursementsToExcel(contract, paymentsList) {
    const prj = this.projects.find((p) => p.id === this.currentProjectId);
    const prjName = prj ? (prj.name || prj.id) : (this.currentProjectId || 'DuAn');
    const isGroupA = contract.group === 'A' || contract.id?.startsWith('A-');
    const partner = isGroupA ? (contract.company_a || '') : (contract.company_b || '');

    let csvContent = '\uFEFF';
    csvContent += `BẢNG THEO DÕI THU CHI HỢP ĐỒNG: ${contract.contract_no || contract.id}\r\n`;
    csvContent += `Dự án: ${prjName}\r\n`;
    csvContent += `Đối tác: ${partner}\r\n`;
    csvContent += `Thời điểm xuất: ${new Date().toLocaleString('vi-VN')}\r\n\r\n`;
    csvContent += `STT,Số ĐN,Nội dung thanh toán,Ngày TT,KL hoàn thành (VNĐ),KL chấp nhận TT (VNĐ),Đề nghị TT (VNĐ),Đã TT (VNĐ),Còn lại (VNĐ),Hóa đơn VAT,Ghi chú\r\n`;

    paymentsList.forEach((p, idx) => {
      const pRate = this.getPaymentRate(p);
      const pActVnd = Math.round((Number(p.actual_volume) || 0) * pRate);
      const pAccVnd = Math.round((Number(p.accepted_vol) || 0) * pRate);
      const pReqVnd = Math.round((Number(p.request_amount) || 0) * pRate);
      const pPaidVnd = Math.round((Number(p.paid_amount) || 0) * pRate);
      const pRemVnd = pReqVnd - pPaidVnd;
      const rawNo = String(p.payment_no || (idx + 1)).trim();
      const cleanNo = rawNo.replace(/^ĐN[-_\s]*/i, '').replace(/^DN[-_\s]*/i, '') || String(idx + 1);

      csvContent += `"${idx + 1}","${cleanNo}","${(this.getPaymentContent(p) || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${pActVnd},${pAccVnd},${pReqVnd},${pPaidVnd},${pRemVnd},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Bang_Thu_Chi_${contract.contract_no || contract.id}_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  // ==========================================
  // PHÂN HỆ NHẮC VIỆC & THEO DÕI HỒ SƠ (REMINDERS - CPMS PLUS)
  // ==========================================

  /**
   * Kiểm tra hồ sơ đã hoàn thành / kết thúc hay chưa (gồm cả Canceled, Returned)
   */
  isDocCompleted(doc) {
    if (!doc) return false;
    const raw = String(doc.status || '').trim().toLowerCase();
    const completedList = [
      'completed',
      'hoàn thành',
      'closed',
      'done',
      'accepted',
      'approved',
      'phê duyệt',
      'chấp thuận',
      'canceled',
      'cancelled',
      'hủy',
      'đã hủy',
      'returned',
      'trả lại'
    ];
    return completedList.includes(raw);
  }

  /**
   * Cập nhật số đếm badge nhắc việc trên thanh tab Reminders
   */
  updateReminderBadge() {
    const badge = document.getElementById('reminder-tab-badge');
    if (!this.documents || this.documents.length === 0) {
      if (badge) {
        badge.textContent = '0';
        badge.style.display = 'none';
      }
      return 0;
    }
    const pendingDocs = this.documents.filter((d) => !this.isDocCompleted(d));
    const count = pendingDocs.length;
    if (badge) {
      badge.textContent = count;
      badge.style.display = count > 0 ? 'inline-flex' : 'none';
    }
    return count;
  }

  /**
   * Lấy số hiệu hoặc số đợt rút gọn cho hồ sơ (hiển thị giống cột 3 trong CPMS Plus Reminders)
   */
  getDocShortSequence(doc) {
    if (!doc) return '—';
    const docNo = String(doc.doc_no || '').trim();
    const title = String(doc.title || '').trim();

    // 1. Tìm số đợt dạng .001, .004, -001, .05 ở đuôi mã hồ sơ
    const matchNo = docNo.match(/(?:\.|\-)(\d{2,4})(?:[^\d]|$)/);
    if (matchNo) return matchNo[1];

    // 2. Tìm số trong tiêu đề (ví dụ: đợt 01, IPC 02)
    const matchTitle = title.match(/(?:đợt|ipc|dn|số|no\.)\s*(\d+)/i);
    if (matchTitle) return matchTitle[1];

    if (doc.version) return String(doc.version);
    return '—';
  }

  /**
   * Lọc danh sách hồ sơ cho Tab Reminder theo bộ lọc hiện tại
   */
  getFilteredReminderDocs() {
    if (!this.documents) return [];

    let list = [...this.documents];

    // 1. Chuẩn hóa ngày nhận cho tất cả hồ sơ nếu chưa có
    list.forEach((d) => {
      if (!d.received_date || String(d.received_date).trim() === '' || d.received_date === '—') {
        d.received_date = this.getDocReceivedDate(d);
      }
    });

    // 2. Lọc theo checkbox Show Completed
    if (!this.reminderShowCompleted) {
      list = list.filter((d) => !this.isDocCompleted(d));
    }

    // 3. Lọc theo Radio Type: ALL / IPC / OTHER
    if (this.reminderDocType === 'IPC') {
      list = list.filter((d) => {
        const cat = String(d.category || '').toUpperCase();
        const title = String(d.title || '').toUpperCase();
        const docNo = String(d.doc_no || '').toUpperCase();
        return cat.includes('IPC') || title.includes('IPC') || docNo.includes('IPC');
      });
    } else if (this.reminderDocType === 'OTHER') {
      list = list.filter((d) => {
        const cat = String(d.category || '').toUpperCase();
        const title = String(d.title || '').toUpperCase();
        const docNo = String(d.doc_no || '').toUpperCase();
        return !cat.includes('IPC') && !title.includes('IPC') && !docNo.includes('IPC');
      });
    }

    // 4. Lọc theo Category dropdown cụ thể (nếu !== 'ALL')
    if (this.reminderCategory && this.reminderCategory !== 'ALL') {
      const targetCat = this.reminderCategory.toUpperCase();
      list = list.filter((d) => String(d.category || '').toUpperCase() === targetCat);
    }

    // 5. Tìm kiếm (Search term)
    if (this.reminderSearchTerm && this.reminderSearchTerm.trim()) {
      const q = this.reminderSearchTerm.trim().toLowerCase();
      list = list.filter((d) => {
        const targetC = this.contracts.find((c) => c.id === d.contract_id);
        const partyB = String(targetC?.company_b || '').toLowerCase();
        const cNo = String(targetC?.contract_no || '').toLowerCase();
        const scope = String(targetC?.scope_of_work || '').toLowerCase();
        const docNo = String(d.doc_no || '').toLowerCase();
        const title = String(d.title || '').toLowerCase();
        const status = String(d.status || '').toLowerCase();
        const cat = String(d.category || '').toLowerCase();
        const id = String(d.id || '').toLowerCase();
        return (
          partyB.includes(q) ||
          cNo.includes(q) ||
          scope.includes(q) ||
          docNo.includes(q) ||
          title.includes(q) ||
          status.includes(q) ||
          cat.includes(q) ||
          id.includes(q)
        );
      });
    }

    // 6. Sắp xếp: Mặc định sắp xếp theo ngày nhận/xử lý (mới nhất lên đầu để tiện theo dõi nhắc việc)
    list.sort((a, b) => {
      const tA = this.parsePaymentDateObj(a.received_date)?.getTime() || 0;
      const tB = this.parsePaymentDateObj(b.received_date)?.getTime() || 0;
      if (tA && tB) {
        if (tA !== tB) return tB - tA;
      } else if (tA && !tB) {
        return -1;
      } else if (!tA && tB) {
        return 1;
      }
      return String(b.id || '').localeCompare(String(a.id || ''));
    });

    return list;
  }

  /**
   * Render giao diện Tab Reminders (Nhắc việc & Theo dõi hồ sơ - chuẩn CPMS Plus)
   */
  renderRemindersTab() {
    this.updateEmptyStateUI('reminder');
    if (!this.isProjectReady()) return;

    const container = document.getElementById('reminder-layout-container');
    if (!container) return;

    // Cập nhật số đếm badge
    this.updateReminderBadge();

    const allDocs = this.documents || [];
    const pendingDocs = allDocs.filter((d) => !this.isDocCompleted(d));
    const totalPending = pendingDocs.length;

    // Lấy danh mục categories duy nhất từ toàn bộ documents
    const catSet = new Set();
    allDocs.forEach((d) => {
      const c = String(d.category || '').trim().toUpperCase();
      if (c) catSet.add(c);
    });
    const categories = Array.from(catSet).sort();

    const categoriesOptionsHtml = categories
      .map(
        (c) => `<option value="${c}" ${this.reminderCategory === c ? 'selected' : ''}>${c}</option>`
      )
      .join('');

    const filteredDocs = this.getFilteredReminderDocs();

    // Xác định hồ sơ đang được chọn
    if (!this.selectedReminderDocId || !filteredDocs.some((d) => d.id === this.selectedReminderDocId)) {
      this.selectedReminderDocId = filteredDocs[0]?.id || null;
    }

    const selectedDoc = (allDocs.find((d) => d.id === this.selectedReminderDocId)) || filteredDocs[0] || null;
    const targetContract = selectedDoc ? this.contracts.find((c) => c.id === selectedDoc.contract_id) : null;
    const isOrphanDoc = Boolean(selectedDoc && !targetContract);

    // Render danh sách các dòng hồ sơ (Cột trái)
    const rowsHtml =
      filteredDocs.length === 0
        ? `<tr><td colspan="5" style="text-align: center; padding: 40px 16px; color: var(--text-muted);">
            <i class="fa-solid fa-clipboard-check" style="font-size: 2rem; color: #10b981; margin-bottom: 8px; display: block;"></i>
            Không có hồ sơ nào cần xử lý trong bộ lọc hiện tại.
          </td></tr>`
        : filteredDocs
            .map((d) => {
              const c = this.contracts.find((item) => item.id === d.contract_id);
              const isRowOrphan = !c;
              const partyBText = c?.company_b || c?.company_a || '[Chưa gắn hợp đồng]';
              const partyBHtml = isRowOrphan
                ? `<span style="color: #dc2626; font-weight: 600;">[Chưa gắn hợp đồng]</span>`
                : partyBText;
              const contractSub = c?.contract_no || (d.contract_id ? `Mã cũ: ${d.contract_id}` : 'Không có mã HĐ');
              const shortSeq = this.getDocShortSequence(d);
              const recDate = this.formatDate(d.received_date || this.getDocReceivedDate(d, c));
              const isSelected = d.id === this.selectedReminderDocId;

              return `
                <tr class="reminder-table-row ${isSelected ? 'selected' : ''}" data-doc-id="${d.id}" style="cursor: pointer;">
                  <td class="reminder-col-party" title="${partyBText}">
                    <div class="reminder-party-cell">
                      <span class="reminder-party-name">${partyBHtml}</span>
                      <span class="reminder-contract-sub">${contractSub}</span>
                    </div>
                  </td>
                  <td style="text-align: center; white-space: nowrap;">
                    ${this.getDocumentCategoryBadgeHtml(d.category)}
                  </td>
                  <td style="text-align: center; white-space: nowrap; font-family: monospace; font-weight: 700; color: #334155;">
                    ${shortSeq}
                  </td>
                  <td style="text-align: center; white-space: nowrap; font-size: 0.82rem; color: #475569;">
                    ${recDate}
                  </td>
                  <td style="text-align: center; white-space: nowrap;">
                    ${this.getDocumentStatusBadgeHtml(d.status)}
                  </td>
                </tr>
              `;
            })
            .join('');

    // Chuẩn bị đường link Local & link Drive cho cột chi tiết
    let locLink = '';
    let drvLink = '';
    let driveDocPath = '';

    if (selectedDoc) {
      locLink = selectedDoc.local_link || (!String(selectedDoc.file_link || '').startsWith('http') && selectedDoc.file_link) || '';
      drvLink = selectedDoc.drive_link || (String(selectedDoc.file_link || '').startsWith('http') && selectedDoc.file_link) || '';
      driveDocPath = targetContract ? this.getDriveDocPathForContract(targetContract) : '';
    }

    // 1. HTML cho Link Local + nút mở file/thư mục Local inline
    const localLinkHtml = `
      <div class="reminder-link-row-flex">
        <div class="reminder-link-text">
          ${locLink
            ? `<span class="font-mono text-muted" style="font-size: 0.78rem; word-break: break-all;" title="${locLink}">${locLink}</span>`
            : '<span style="color: #94a3b8;">—</span>'}
        </div>
        <div class="reminder-link-btn-group">
          ${locLink
            ? `<button type="button" class="button button-outline btn-reminder-inline-icon btn-copy-doc-local" data-path="${encodeURIComponent(locLink)}" title="Sao chép đường dẫn máy tính">
                 <i class="fa-regular fa-copy"></i>
               </button>`
            : ''}
          <button type="button" id="btn-reminder-open-file-local" class="button button-outline btn-reminder-inline-icon btn-action-local" title="Mở file trên máy tính">
            <i class="fa-solid fa-file-lines"></i>
          </button>
          <button type="button" id="btn-reminder-open-folder-local" class="button button-outline btn-reminder-inline-icon btn-action-local-folder" title="Mở thư mục trên máy tính">
            <i class="fa-solid fa-folder-open"></i>
          </button>
        </div>
      </div>
    `;

    // 2. HTML cho Link Drive + nút mở file/thư mục Drive inline
    const driveLinkHtml = `
      <div class="reminder-link-row-flex">
        <div class="reminder-link-text">
          ${drvLink
            ? `<a href="${drvLink}" target="_blank" class="font-mono" style="font-size: 0.78rem; color: #0284c7; word-break: break-all; text-decoration: none;" title="${drvLink}">${driveDocPath || drvLink}</a>`
            : driveDocPath
              ? `<span class="font-mono" style="font-size: 0.78rem; color: #0369a1; word-break: break-all;" title="${driveDocPath}">${driveDocPath}</span>`
              : '<span style="color: #94a3b8;">—</span>'}
        </div>
        <div class="reminder-link-btn-group">
          <button type="button" id="btn-reminder-open-file-drive" class="button button-outline btn-reminder-inline-icon btn-action-drive-file" title="Mở file trên Google Drive">
            <i class="fa-solid fa-file-lines"></i>
          </button>
          <button type="button" id="btn-reminder-open-folder-drive" class="button button-outline btn-reminder-inline-icon btn-action-drive-folder" title="Mở thư mục trên Google Drive">
            <i class="fa-solid fa-folder-open"></i>
          </button>
        </div>
      </div>
    `;

    // Render toàn bộ layout 2 cột
    container.innerHTML = `
      <div class="reminder-main-grid">
        <!-- CỘT TRÁI: DANH SÁCH HỒ SƠ & BỘ LỌC -->
        <div class="reminder-left-col">
          <!-- BỘ LỌC (FILTER CARD) GỌN GÀNG, CÂN ĐỐI -->
          <div class="card reminder-filter-card">
            <div class="reminder-filter-inner">
              <!-- 1. Ô tìm kiếm bên TRÁI -->
              <div class="search-input-wrap reminder-search-wrap">
                <i class="fa-solid fa-magnifying-glass search-icon"></i>
                <input type="text" id="reminder-search-input" class="search-input" placeholder="Tìm kiếm..." value="${this.reminderSearchTerm || ''}">
                <button type="button" id="btn-clear-reminder-search" class="btn-clear-search ${this.reminderSearchTerm ? '' : 'hidden'}" title="Xóa từ khóa">
                  <i class="fa-solid fa-xmark"></i>
                </button>
              </div>

              <span class="reminder-filter-vdivider"></span>

              <!-- 2. Checkbox Show Completed -->
              <label class="reminder-checkbox-wrap" title="Tích chọn để hiển thị cả các hồ sơ đã hoàn tất / kết thúc (Approved, Accepted, Completed, Canceled, Returned)">
                <input type="checkbox" id="reminder-filter-show-completed" ${this.reminderShowCompleted ? 'checked' : ''}>
                <span>Show Completed</span>
              </label>

              <span class="reminder-filter-vdivider"></span>

              <!-- 3. Radio Buttons -->
              <div class="reminder-radio-group">
                <label class="reminder-radio-label" title="Tất cả hồ sơ (All Documents)">
                  <input type="radio" name="reminder-doc-type-radio" value="ALL" ${this.reminderDocType === 'ALL' ? 'checked' : ''}>
                  <span>Tất cả</span>
                </label>
                <label class="reminder-radio-label" title="Chỉ hồ sơ thanh toán đợt (Interim Payment Certificate)">
                  <input type="radio" name="reminder-doc-type-radio" value="IPC" ${this.reminderDocType === 'IPC' ? 'checked' : ''}>
                  <span>IPC</span>
                </label>
                <label class="reminder-radio-label" title="Các loại tài liệu khác (Other Documents)">
                  <input type="radio" name="reminder-doc-type-radio" value="OTHER" ${this.reminderDocType === 'OTHER' ? 'checked' : ''}>
                  <span>Khác</span>
                </label>
              </div>

              <span class="reminder-filter-vdivider"></span>

              <!-- 4. Dropdown chọn Loại bên PHẢI -->
              <select id="reminder-filter-category" class="form-control reminder-category-select" title="Lọc theo loại hồ sơ">
                <option value="ALL">Tất cả loại</option>
                ${categoriesOptionsHtml}
              </select>
            </div>
          </div>

          <!-- BẢNG DANH SÁCH (TABLE CARD) -->
          <div class="card reminder-table-card">
            <div class="reminder-table-scroll-wrap">
              <table class="data-table reminder-data-table">
                <thead>
                  <tr>
                    <th style="text-align: left; min-width: 220px;">Party B (Đối tác)</th>
                    <th style="text-align: center; width: 75px;">Type</th>
                    <th style="text-align: center; width: 85px;">Số đợt/hiệu</th>
                    <th style="text-align: center; width: 105px;">Ngày nhận</th>
                    <th style="text-align: center; width: 120px;">Tình trạng</th>
                  </tr>
                </thead>
                <tbody id="reminder-table-tbody">
                  ${rowsHtml}
                </tbody>
              </table>
            </div>

            <!-- FOOTER DANH SÁCH -->
            <div class="reminder-list-footer">
              <div class="reminder-pending-count-text">
                Pending [All]: <strong id="reminder-pending-number" class="text-primary" style="font-size: 1.05rem;">${totalPending}</strong> document(s)
                ${this.reminderShowCompleted ? `<span style="font-size: 0.8rem; color: #64748b; margin-left: 8px;">(Đang hiện ${filteredDocs.length} / ${allDocs.length} hồ sơ)</span>` : ''}
              </div>
              <div class="reminder-footer-hint">
                <i class="fa-solid fa-circle-info"></i> Nhấp đúp vào dòng để mở chi tiết hợp đồng
              </div>
            </div>
          </div>
        </div>

        <!-- CỘT PHẢI: KHUNG CHI TIẾT HỒ SƠ (REMINDER DETAIL CARD - CPMS PLUS) -->
        <div class="reminder-right-col" id="reminder-right-col">
          <div class="card reminder-detail-card">
            <!-- HEADER CHI TIẾT -->
            <div class="reminder-detail-header">
              <div class="reminder-detail-title-box">
                ${
                  isOrphanDoc
                    ? `<span class="badge" style="background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca;">CHƯA GÁN HĐ</span>`
                    : `<span class="badge ${targetContract?.group === 'B' ? 'badge-warning' : 'badge-primary'}">GROUP ${targetContract?.group || 'A'}</span>`
                }
                <span class="reminder-detail-contract-no">
                  ${
                    targetContract?.contract_no ||
                    (selectedDoc
                      ? selectedDoc.contract_id
                        ? `Mã tham chiếu: ${selectedDoc.contract_id}`
                        : 'Chưa gắn hợp đồng / phụ lục'
                      : 'Chưa chọn hồ sơ')
                  }
                </span>
              </div>
              <div>${selectedDoc ? this.getDocumentStatusBadgeHtml(selectedDoc.status) : ''}</div>
            </div>

            <!-- NỘI DUNG CHI TIẾT -->
            ${
              selectedDoc
                ? `
              <div class="reminder-detail-fields">
                <div class="reminder-field-row">
                  <div class="reminder-field-label">Party B</div>
                  <div class="reminder-field-value"><strong>${targetContract?.company_b || targetContract?.company_a || '—'}</strong></div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Scope</div>
                  <div class="reminder-field-value">${targetContract?.scope_of_work || '—'}</div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Type</div>
                  <div class="reminder-field-value">
                    ${this.getDocumentCategoryBadgeHtml(selectedDoc.category)}
                  </div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Name</div>
                  <div class="reminder-field-value reminder-field-name"><strong>${selectedDoc.title || selectedDoc.name || '—'}</strong></div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Document Number</div>
                  <div class="reminder-field-value font-mono"><strong>${selectedDoc.doc_no || selectedDoc.id || '—'}</strong></div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Status</div>
                  <div class="reminder-field-value">${this.getDocumentStatusBadgeHtml(selectedDoc.status)}</div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Processed Date</div>
                  <div class="reminder-field-value">${this.formatDate(this.getDocReceivedDate(selectedDoc, targetContract))}</div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Supporting Docs</div>
                  <div class="reminder-field-value">${selectedDoc.supporting_docs || '—'}</div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Version</div>
                  <div class="reminder-field-value">${selectedDoc.version || '1'}</div>
                </div>

                <div class="reminder-field-row">
                  <div class="reminder-field-label">Note</div>
                  <div class="reminder-field-value reminder-field-note">${selectedDoc.note || '—'}</div>
                </div>

                <!-- LINK LOCAL (MÁY TÍNH) -->
                <div class="reminder-field-row">
                  <div class="reminder-field-label">
                    <i class="fa-solid fa-laptop" style="color: #475569; margin-right: 2px;"></i> Link Local
                  </div>
                  <div class="reminder-field-value reminder-field-link">${localLinkHtml}</div>
                </div>

                <!-- LINK GOOGLE DRIVE -->
                <div class="reminder-field-row">
                  <div class="reminder-field-label">
                    <i class="fa-brands fa-google-drive" style="color: #0284c7; margin-right: 2px;"></i> Link Drive
                  </div>
                  <div class="reminder-field-value reminder-field-link">${driveLinkHtml}</div>
                </div>
              </div>

              <!-- FOOTER THAO TÁC (ACTIONS) -->
              <div class="reminder-detail-actions ${isOrphanDoc ? 'is-orphan-actions' : ''}">
                <button type="button" id="btn-reminder-refresh" class="button button-outline" title="Làm mới danh sách dữ liệu">
                  <i class="fa-solid fa-rotate-right"></i> Làm mới
                </button>
                <button type="button" id="btn-reminder-goto" class="button button-primary" ${isOrphanDoc ? 'disabled style="opacity: 0.5; cursor: not-allowed;"' : ''} title="${isOrphanDoc ? 'Tài liệu chưa gắn hợp đồng, vui lòng gán hợp đồng trước' : 'Mở chi tiết Hợp đồng và trỏ đến hồ sơ này'}">
                  <i class="fa-solid fa-arrow-up-right-from-square"></i> Mở HĐ
                </button>
                ${
                  isOrphanDoc
                    ? `
                  <button type="button" id="btn-reminder-assign-contract" class="button button-outline btn-reminder-assign-contract" title="Gán tài liệu này vào một Hợp đồng hoặc Phụ lục hiện có">
                    <i class="fa-solid fa-link"></i> Gán HĐ
                  </button>
                  <button type="button" id="btn-reminder-delete-orphan" class="button button-outline btn-reminder-delete-orphan" title="Xóa vĩnh viễn tài liệu chưa gắn hợp đồng này">
                    <i class="fa-solid fa-trash-can"></i> Xóa
                  </button>
                `
                    : ''
                }
              </div>
            `
                : `
              <div style="text-align: center; padding: 60px 20px; color: var(--text-muted);">
                <i class="fa-regular fa-folder-open" style="font-size: 2.5rem; color: #cbd5e1; margin-bottom: 12px; display: block;"></i>
                Chọn một hồ sơ trong danh sách bên trái để xem chi tiết.
              </div>
            `
            }
          </div>
        </div>
      </div>
    `;

    // Gắn sự kiện cho các thành phần điều khiển
    this.bindReminderEvents(filteredDocs, selectedDoc, targetContract, locLink, drvLink, driveDocPath);
  }

  /**
   * Gắn sự kiện tương tác cho Tab Reminders
   */
  bindReminderEvents(filteredDocs, selectedDoc, targetContract, locLink = '', drvLink = '', driveDocPath = '') {
    const container = document.getElementById('reminder-layout-container');
    if (!container) return;

    this.bindReminderRowEvents(container);
    this.bindReminderDetailEvents(selectedDoc, targetContract, locLink, drvLink, driveDocPath);
    this.bindReminderFilterEvents(container);
  }

  bindReminderRowEvents(container = document.getElementById('reminder-layout-container')) {
    if (!container) return;
    container.querySelectorAll('.reminder-table-row').forEach((row) => {
      row.addEventListener('click', () => {
        const docId = row.getAttribute('data-doc-id');
        if (docId && docId !== this.selectedReminderDocId) {
          container.querySelectorAll('.reminder-table-row.selected').forEach((r) => r.classList.remove('selected'));
          row.classList.add('selected');
          this.selectedReminderDocId = docId;
          this.updateReminderDetailOnly();
        }
      });

      row.addEventListener('dblclick', () => {
        const docId = row.getAttribute('data-doc-id');
        const doc = (this.documents || []).find((d) => d.id === docId);
        if (doc) {
          this.handleReminderGoTo(doc);
        }
      });
    });
  }

  updateReminderDetailOnly() {
    const rightCol = document.getElementById('reminder-right-col');
    if (!rightCol) return;

    const allDocs = this.documents || [];
    const filteredDocs = this.getFilteredReminderDocs();
    const selectedDoc = allDocs.find((d) => d.id === this.selectedReminderDocId) || filteredDocs[0] || null;
    const targetContract = selectedDoc ? this.contracts.find((c) => c.id === selectedDoc.contract_id) : null;
    const isOrphanDoc = Boolean(selectedDoc && !targetContract);

    let locLink = '';
    let drvLink = '';
    let driveDocPath = '';
    if (selectedDoc) {
      locLink = selectedDoc.local_link || (!String(selectedDoc.file_link || '').startsWith('http') && selectedDoc.file_link) || '';
      drvLink = selectedDoc.drive_link || (String(selectedDoc.file_link || '').startsWith('http') && selectedDoc.file_link) || '';
      driveDocPath = targetContract ? this.getDriveDocPathForContract(targetContract) : '';
    }

    const localLinkHtml = `
      <div class="reminder-link-row-flex">
        <div class="reminder-link-text">
          ${locLink
            ? `<span class="font-mono text-muted" style="font-size: 0.78rem; word-break: break-all;" title="${locLink}">${locLink}</span>`
            : '<span style="color: #94a3b8;">—</span>'}
        </div>
        <div class="reminder-link-btn-group">
          ${locLink
            ? `<button type="button" class="button button-outline btn-reminder-inline-icon btn-copy-doc-local" data-path="${encodeURIComponent(locLink)}" title="Sao chép đường dẫn máy tính">
                 <i class="fa-regular fa-copy"></i>
               </button>`
            : ''}
          <button type="button" id="btn-reminder-open-file-local" class="button button-outline btn-reminder-inline-icon btn-action-local" title="Mở file trên máy tính">
            <i class="fa-solid fa-file-lines"></i>
          </button>
          <button type="button" id="btn-reminder-open-folder-local" class="button button-outline btn-reminder-inline-icon btn-action-local-folder" title="Mở thư mục trên máy tính">
            <i class="fa-solid fa-folder-open"></i>
          </button>
        </div>
      </div>
    `;

    const driveLinkHtml = `
      <div class="reminder-link-row-flex">
        <div class="reminder-link-text">
          ${drvLink
            ? `<a href="${drvLink}" target="_blank" class="font-mono" style="font-size: 0.78rem; color: #0284c7; word-break: break-all; text-decoration: none;" title="${drvLink}">${driveDocPath || drvLink}</a>`
            : driveDocPath
              ? `<span class="font-mono" style="font-size: 0.78rem; color: #0369a1; word-break: break-all;" title="${driveDocPath}">${driveDocPath}</span>`
              : '<span style="color: #94a3b8;">—</span>'}
        </div>
        <div class="reminder-link-btn-group">
          <button type="button" id="btn-reminder-open-file-drive" class="button button-outline btn-reminder-inline-icon btn-action-drive-file" title="Mở file trên Google Drive">
            <i class="fa-solid fa-file-lines"></i>
          </button>
          <button type="button" id="btn-reminder-open-folder-drive" class="button button-outline btn-reminder-inline-icon btn-action-drive-folder" title="Mở thư mục trên Google Drive">
            <i class="fa-solid fa-folder-open"></i>
          </button>
        </div>
      </div>
    `;

    rightCol.innerHTML = `
      <div class="card reminder-detail-card">
        <div class="reminder-detail-header">
          <div class="reminder-detail-title-box">
            ${
              isOrphanDoc
                ? `<span class="badge" style="background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca;">CHƯA GÁN HĐ</span>`
                : `<span class="badge ${targetContract?.group === 'B' ? 'badge-warning' : 'badge-primary'}">GROUP ${targetContract?.group || 'A'}</span>`
            }
            <span class="reminder-detail-contract-no">
              ${
                targetContract?.contract_no ||
                (selectedDoc
                  ? selectedDoc.contract_id
                    ? `Mã tham chiếu: ${selectedDoc.contract_id}`
                    : 'Chưa gắn hợp đồng / phụ lục'
                  : 'Chưa chọn hồ sơ')
              }
            </span>
          </div>
          <div>${selectedDoc ? this.getDocumentStatusBadgeHtml(selectedDoc.status) : ''}</div>
        </div>
        ${
          selectedDoc
            ? `
          <div class="reminder-detail-fields">
            <div class="reminder-field-row">
              <div class="reminder-field-label">Party B</div>
              <div class="reminder-field-value"><strong>${targetContract?.company_b || targetContract?.company_a || '—'}</strong></div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Scope</div>
              <div class="reminder-field-value">${targetContract?.scope_of_work || '—'}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Type</div>
              <div class="reminder-field-value">${this.getDocumentCategoryBadgeHtml(selectedDoc.category)}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Name</div>
              <div class="reminder-field-value reminder-field-name"><strong>${selectedDoc.title || selectedDoc.name || '—'}</strong></div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Document Number</div>
              <div class="reminder-field-value font-mono"><strong>${selectedDoc.doc_no || selectedDoc.id || '—'}</strong></div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Status</div>
              <div class="reminder-field-value">${this.getDocumentStatusBadgeHtml(selectedDoc.status)}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Processed Date</div>
              <div class="reminder-field-value">${this.formatDate(this.getDocReceivedDate(selectedDoc, targetContract))}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Supporting Docs</div>
              <div class="reminder-field-value">${selectedDoc.supporting_docs || '—'}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Version</div>
              <div class="reminder-field-value">${selectedDoc.version || '1'}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label">Note</div>
              <div class="reminder-field-value reminder-field-note">${selectedDoc.note || '—'}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label"><i class="fa-solid fa-laptop" style="color: #475569; margin-right: 2px;"></i> Link Local</div>
              <div class="reminder-field-value reminder-field-link">${localLinkHtml}</div>
            </div>
            <div class="reminder-field-row">
              <div class="reminder-field-label"><i class="fa-brands fa-google-drive" style="color: #0284c7; margin-right: 2px;"></i> Link Drive</div>
              <div class="reminder-field-value reminder-field-link">${driveLinkHtml}</div>
            </div>
          </div>
          <div class="reminder-detail-actions ${isOrphanDoc ? 'is-orphan-actions' : ''}">
            <button type="button" id="btn-reminder-refresh" class="button button-outline" title="Làm mới danh sách dữ liệu">
              <i class="fa-solid fa-rotate-right"></i> Làm mới
            </button>
            <button type="button" id="btn-reminder-goto" class="button button-primary" ${isOrphanDoc ? 'disabled style="opacity: 0.5; cursor: not-allowed;"' : ''} title="${isOrphanDoc ? 'Tài liệu chưa gắn hợp đồng, vui lòng gán hợp đồng trước' : 'Mở chi tiết Hợp đồng và trỏ đến hồ sơ này'}">
              <i class="fa-solid fa-arrow-up-right-from-square"></i> Mở HĐ
            </button>
            ${
              isOrphanDoc
                ? `
              <button type="button" id="btn-reminder-assign-contract" class="button button-outline btn-reminder-assign-contract" title="Gán tài liệu này vào một Hợp đồng hoặc Phụ lục hiện có">
                <i class="fa-solid fa-link"></i> Gán HĐ
              </button>
              <button type="button" id="btn-reminder-delete-orphan" class="button button-outline btn-reminder-delete-orphan" title="Xóa vĩnh viễn tài liệu chưa gắn hợp đồng này">
                <i class="fa-solid fa-trash-can"></i> Xóa
              </button>
            `
                : ''
            }
          </div>
        `
            : `
          <div style="text-align: center; padding: 60px 20px; color: var(--text-muted);">
            <i class="fa-regular fa-folder-open" style="font-size: 2.5rem; color: #cbd5e1; margin-bottom: 12px; display: block;"></i>
            Chọn một hồ sơ trong danh sách bên trái để xem chi tiết.
          </div>
        `
        }
      </div>
    `;

    this.bindReminderDetailEvents(selectedDoc, targetContract, locLink, drvLink, driveDocPath);
  }

  /**
   * Chỉ cập nhật lại tbody bảng và cột chi tiết khi tìm kiếm/lọc (giữ nguyên ô input tìm kiếm & focus)
   */
  refreshReminderListAndDetail() {
    const tbody = document.getElementById('reminder-table-tbody');
    if (!tbody) {
      this.renderRemindersTab();
      return;
    }

    const allDocs = this.documents || [];
    const totalPending = allDocs.filter((d) => !this.isDocCompleted(d)).length;
    const filteredDocs = this.getFilteredReminderDocs();

    if (!this.selectedReminderDocId || !filteredDocs.some((d) => d.id === this.selectedReminderDocId)) {
      this.selectedReminderDocId = filteredDocs[0]?.id || null;
    }

    const rowsHtml =
      filteredDocs.length === 0
        ? `<tr><td colspan="5" style="text-align: center; padding: 40px 16px; color: var(--text-muted);">
            <i class="fa-solid fa-clipboard-check" style="font-size: 2rem; color: #10b981; margin-bottom: 8px; display: block;"></i>
            Không có hồ sơ nào cần xử lý trong bộ lọc hiện tại.
          </td></tr>`
        : filteredDocs
            .map((d) => {
              const c = this.contracts.find((item) => item.id === d.contract_id);
              const isRowOrphan = !c;
              const partyBText = c?.company_b || c?.company_a || '[Chưa gắn hợp đồng]';
              const partyBHtml = isRowOrphan
                ? `<span style="color: #dc2626; font-weight: 600;">[Chưa gắn hợp đồng]</span>`
                : partyBText;
              const contractSub = c?.contract_no || (d.contract_id ? `Mã cũ: ${d.contract_id}` : 'Không có mã HĐ');
              const shortSeq = this.getDocShortSequence(d);
              const recDate = this.formatDate(d.received_date || this.getDocReceivedDate(d, c));
              const isSelected = d.id === this.selectedReminderDocId;

              return `
                <tr class="reminder-table-row ${isSelected ? 'selected' : ''}" data-doc-id="${d.id}" style="cursor: pointer;">
                  <td class="reminder-col-party" title="${partyBText}">
                    <div class="reminder-party-cell">
                      <span class="reminder-party-name">${partyBHtml}</span>
                      <span class="reminder-contract-sub">${contractSub}</span>
                    </div>
                  </td>
                  <td style="text-align: center; white-space: nowrap;">
                    ${this.getDocumentCategoryBadgeHtml(d.category)}
                  </td>
                  <td style="text-align: center; white-space: nowrap; font-family: monospace; font-weight: 700; color: #334155;">
                    ${shortSeq}
                  </td>
                  <td style="text-align: center; white-space: nowrap; font-size: 0.82rem; color: #475569;">
                    ${recDate}
                  </td>
                  <td style="text-align: center; white-space: nowrap;">
                    ${this.getDocumentStatusBadgeHtml(d.status)}
                  </td>
                </tr>
              `;
            })
            .join('');

    tbody.innerHTML = rowsHtml;

    const countEl = document.querySelector('.reminder-pending-count-text');
    if (countEl) {
      countEl.innerHTML = `
        Pending [All]: <strong id="reminder-pending-number" class="text-primary" style="font-size: 1.05rem;">${totalPending}</strong> document(s)
        ${this.reminderShowCompleted ? `<span style="font-size: 0.8rem; color: #64748b; margin-left: 8px;">(Đang hiện ${filteredDocs.length} / ${allDocs.length} hồ sơ)</span>` : ''}
      `;
    }

    this.bindReminderRowEvents();
    this.updateReminderDetailOnly();
  }

  bindReminderDetailEvents(selectedDoc, targetContract, locLink = '', drvLink = '', driveDocPath = '') {
    const rightCol = document.getElementById('reminder-right-col');
    if (!rightCol) return;

    document.getElementById('btn-reminder-refresh')?.addEventListener('click', async () => {
      const btn = document.getElementById('btn-reminder-refresh');
      if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tải...';
      await this.loadData();
      this.updateReminderBadge();
      this.refreshReminderListAndDetail();
      this.showLocalSaveStatus('Đã làm mới danh sách hồ sơ nhắc việc!');
    });

    document.getElementById('btn-reminder-goto')?.addEventListener('click', () => {
      if (selectedDoc && targetContract) {
        this.handleReminderGoTo(selectedDoc);
      }
    });

    document.getElementById('btn-reminder-assign-contract')?.addEventListener('click', () => {
      if (selectedDoc) {
        this.openAssignContractModal(selectedDoc);
      }
    });

    document.getElementById('btn-reminder-delete-orphan')?.addEventListener('click', async () => {
      if (!selectedDoc) return;
      const docLabel = selectedDoc.doc_no || selectedDoc.title || selectedDoc.name || selectedDoc.id;
      const ok = await this.confirmAsync({
        title: 'Xóa tài liệu chưa gắn hợp đồng',
        message: `Bạn có chắc chắn muốn xóa vĩnh viễn tài liệu "${docLabel}" không?`,
        subMessage: `Mã nội bộ: ${selectedDoc.id}. Thao tác này không thể hoàn tác!`,
        confirmText: 'Xóa tài liệu',
        cancelText: 'Hủy bỏ',
        confirmType: 'danger',
      });
      if (!ok) return;

      await db.delete('documents', selectedDoc.id);
      this.selectedReminderDocId = null;
      await this.loadData();
      await this.syncToLocalFile();
      this.updateReminderBadge();
      this.refreshReminderListAndDetail();
      this.showLocalSaveStatus('Đã xóa tài liệu chưa gắn hợp đồng!');
    });

    document.getElementById('btn-reminder-open-file-local')?.addEventListener('click', () => {
      if (selectedDoc) this.handleOpenDocFile(selectedDoc);
    });

    document.getElementById('btn-reminder-open-folder-local')?.addEventListener('click', () => {
      if (selectedDoc) this.handleOpenDocFolder(selectedDoc, targetContract);
    });

    document.getElementById('btn-reminder-open-file-drive')?.addEventListener('click', () => {
      if (drvLink) {
        window.open(drvLink, '_blank');
      } else if (selectedDoc?.drive_link) {
        window.open(selectedDoc.drive_link, '_blank');
      } else {
        const q = selectedDoc?.doc_no || selectedDoc?.title || selectedDoc?.id || '';
        window.open(`https://drive.google.com/drive/search?q=${encodeURIComponent(q)}`, '_blank');
      }
    });

    document.getElementById('btn-reminder-open-folder-drive')?.addEventListener('click', async () => {
      if (drvLink && drvLink.includes('/folders/')) {
        window.open(drvLink, '_blank');
      } else {
        const path = driveDocPath || (targetContract ? this.getDriveDocPathForContract(targetContract) : '');
        await this.alertAsync({
          title: 'Thư mục Google Drive',
          message: `Thư mục lưu trữ trên Google Drive:\n\n${path || 'Cấu trúc MAIO_Data / 01_QuanLy_DuAn / Doc'}\n\n(Bấm "Mở Drive" để truy cập Google Drive của bạn)`,
          okText: 'Mở Drive',
          type: 'info',
        });
        window.open('https://drive.google.com', '_blank');
      }
    });

    rightCol.querySelectorAll('.btn-copy-doc-local').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const p = decodeURIComponent(btn.getAttribute('data-path') || '');
        if (p) {
          navigator.clipboard
            .writeText(p)
            .then(() => {
              const orig = btn.innerHTML;
              btn.innerHTML = '<i class="fa-solid fa-check" style="color: var(--success);"></i>';
              setTimeout(() => {
                btn.innerHTML = orig;
              }, 1500);
            })
            .catch(() => {
              prompt('Đường dẫn file trên máy tính (sao chép Ctrl+C):', p);
            });
        }
      });
    });
  }

  bindReminderFilterEvents(container = document.getElementById('reminder-layout-container')) {
    if (!container) return;

    document.getElementById('reminder-filter-show-completed')?.addEventListener('change', (e) => {
      this.reminderShowCompleted = e.target.checked;
      this.refreshReminderListAndDetail();
    });

    container.querySelectorAll('input[name="reminder-doc-type-radio"]').forEach((radio) => {
      radio.addEventListener('change', (e) => {
        this.reminderDocType = e.target.value;
        this.refreshReminderListAndDetail();
      });
    });

    document.getElementById('reminder-filter-category')?.addEventListener('change', (e) => {
      this.reminderCategory = e.target.value;
      this.refreshReminderListAndDetail();
    });

    const searchInput = document.getElementById('reminder-search-input');
    const clearBtn = document.getElementById('btn-clear-reminder-search');

    searchInput?.addEventListener('input', (e) => {
      this.reminderSearchTerm = e.target.value;
      if (clearBtn) {
        if (this.reminderSearchTerm) clearBtn.classList.remove('hidden');
        else clearBtn.classList.add('hidden');
      }
      this.refreshReminderListAndDetail();
    });

    clearBtn?.addEventListener('click', () => {
      this.reminderSearchTerm = '';
      if (searchInput) {
        searchInput.value = '';
        searchInput.focus();
      }
      clearBtn.classList.add('hidden');
      this.refreshReminderListAndDetail();
    });
  }

  /**
   * Mở modal cho phép gán một tài liệu chưa gắn hợp đồng (orphan) vào Hợp đồng hoặc Phụ lục hiện có
   */
  openAssignContractModal(doc) {
    if (!doc) return;
    document.getElementById('reminder-assign-contract-modal')?.remove();

    const projectContracts = (this.contracts || []).filter(
      (c) => !this.currentProjectId || c.project_id === this.currentProjectId
    );

    // Sắp xếp: Nhóm A trước -> Nhóm B, Hợp đồng cha trước -> Phụ lục con
    const parents = projectContracts.filter((c) => !c.parent_id);
    const orderedContracts = [];
    parents.sort((a, b) => {
      const gCmp = String(a.group || '').localeCompare(String(b.group || ''));
      if (gCmp !== 0) return gCmp;
      return String(a.contract_no || '').localeCompare(String(b.contract_no || ''));
    });

    parents.forEach((p) => {
      orderedContracts.push({ contract: p, isChild: false });
      const children = projectContracts
        .filter((c) => c.parent_id === p.id)
        .sort((a, b) => String(a.contract_no || '').localeCompare(String(b.contract_no || '')));
      children.forEach((ch) => orderedContracts.push({ contract: ch, isChild: true }));
    });

    // Thêm các phụ lục mồ côi cha (nếu có)
    const addedIds = new Set(orderedContracts.map((item) => item.contract.id));
    projectContracts.forEach((c) => {
      if (!addedIds.has(c.id)) {
        orderedContracts.push({ contract: c, isChild: Boolean(c.parent_id) });
      }
    });

    const buildOptionsHtml = (keyword = '') => {
      const q = keyword.trim().toLowerCase();
      const matched = orderedContracts.filter(({ contract }) => {
        if (!q) return true;
        const text = `${contract.group || ''} ${contract.contract_no || ''} ${contract.company_b || ''} ${contract.company_a || ''} ${contract.scope_of_work || ''}`.toLowerCase();
        return text.includes(q);
      });

      if (matched.length === 0) {
        return '<option value="" disabled>Không tìm thấy hợp đồng / phụ lục phù hợp</option>';
      }

      return matched
        .map(({ contract, isChild }) => {
          const prefix = isChild ? '↳ [Phụ lục] ' : `[Nhóm ${contract.group || 'A'}] `;
          const partner = contract.company_b || contract.company_a || '';
          const scope = contract.scope_of_work ? ` — ${contract.scope_of_work}` : '';
          const label = `${prefix}${contract.contract_no || contract.id}${partner ? ` | ${partner}` : ''}${scope}`;
          return `<option value="${contract.id}">${label}</option>`;
        })
        .join('');
    };

    const modal = document.createElement('div');
    modal.id = 'reminder-assign-contract-modal';
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 580px;">
        <div class="modal-header">
          <h2 style="font-size: 1.08rem; font-weight: 700; display: flex; align-items: center; gap: 8px;">
            <i class="fa-solid fa-link" style="color: #4338ca;"></i>
            Gán tài liệu vào Hợp đồng / Phụ lục
          </h2>
          <button type="button" id="btn-close-assign-contract" class="button button-outline" style="padding: 4px 10px;">
            <i class="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div class="modal-body" style="padding: 16px 20px;">
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; margin-bottom: 14px; font-size: 0.82rem;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
              ${this.getDocumentCategoryBadgeHtml(doc.category)}
              <strong style="color: #0f172a;">${doc.doc_no || doc.id}</strong>
            </div>
            <div style="color: #334155;">${doc.title || doc.name || '—'}</div>
            ${doc.contract_id ? `<div style="color: #dc2626; font-size: 0.76rem; margin-top: 4px;">Mã HĐ cũ không tồn tại: <code>${doc.contract_id}</code></div>` : ''}
          </div>

          <div class="form-group" style="margin-bottom: 10px;">
            <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Tìm nhanh Hợp đồng / Phụ lục</label>
            <input type="text" id="assign-contract-search" class="form-control" placeholder="Nhập số hợp đồng, tên đối tác, phạm vi...">
          </div>

          <div class="form-group" style="margin-bottom: 16px;">
            <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Chọn Hợp đồng / Phụ lục đích <span style="color: var(--danger);">*</span></label>
            <select id="assign-contract-select" class="form-control" size="8" style="height: auto !important; padding: 6px !important; font-size: 0.82rem;">
              ${buildOptionsHtml()}
            </select>
          </div>

          <div style="display: flex; justify-content: flex-end; gap: 8px;">
            <button type="button" id="btn-cancel-assign-contract" class="button button-outline">Hủy</button>
            <button type="button" id="btn-confirm-assign-contract" class="button button-primary">
              <i class="fa-solid fa-check"></i> Xác nhận gán
            </button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-assign-contract')?.addEventListener('click', closeModal);
    modal.querySelector('#btn-cancel-assign-contract')?.addEventListener('click', closeModal);
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal();
    });

    const searchInput = modal.querySelector('#assign-contract-search');
    const selectEl = modal.querySelector('#assign-contract-select');

    searchInput?.addEventListener('input', (e) => {
      if (selectEl) {
        selectEl.innerHTML = buildOptionsHtml(e.target.value);
        if (selectEl.options.length > 0 && !selectEl.options[0].disabled) {
          selectEl.selectedIndex = 0;
        }
      }
    });

    if (selectEl && selectEl.options.length > 0 && !selectEl.options[0].disabled) {
      selectEl.selectedIndex = 0;
    }

    const doAssign = async () => {
      const targetId = selectEl?.value;
      if (!targetId) {
        alert('Vui lòng chọn một Hợp đồng hoặc Phụ lục để gán!');
        return;
      }

      const updatedDoc = {
        ...doc,
        contract_id: targetId,
        project_id: doc.project_id || this.currentProjectId,
        updated_at: new Date().toISOString()
      };

      await db.put('documents', updatedDoc);
      closeModal();
      this.selectedReminderDocId = updatedDoc.id;
      await this.loadData();
      await this.syncToLocalFile();
      this.updateReminderBadge();
      this.renderRemindersTab();
      this.showLocalSaveStatus('Đã gán tài liệu vào hợp đồng thành công!');
    };

    modal.querySelector('#btn-confirm-assign-contract')?.addEventListener('click', doAssign);
    selectEl?.addEventListener('dblclick', doAssign);
  }

  /**
   * Điều hướng trực tiếp đến Hợp đồng/Phụ lục chứa hồ sơ và trỏ vào Tab 2 Hồ sơ
   */
  handleReminderGoTo(doc) {
    if (!doc || !doc.contract_id) return;
    this.openContractDetail(doc.contract_id);

    // Mở Tab 2 (m-tab-docs) trong modal hợp đồng
    setTimeout(() => {
      const docsTabBtn = document.querySelector('.modal-tab-btn[data-target="m-tab-docs"]');
      if (docsTabBtn) {
        docsTabBtn.click();
      }

      // Cuộn và highlight dòng hồ sơ
      setTimeout(() => {
        const row =
          document.querySelector(`.doc-item-row[data-doc-id="${doc.id}"]`) ||
          Array.from(document.querySelectorAll('.doc-item-row')).find(
            (r) => r.textContent.includes(doc.id) || (doc.doc_no && r.textContent.includes(doc.doc_no))
          );

        if (row) {
          row.scrollIntoView({ behavior: 'smooth', block: 'center' });
          row.classList.add('row-highlight-active');
          setTimeout(() => {
            row.classList.remove('row-highlight-active');
          }, 3500);
        }
      }, 150);
    }, 100);
  }

  // =========================================================================
  // TRA CỨU CÔNG TY, CHỌN HỢP ĐỒNG, TỔNG HỢP DÒNG TIỀN & XUẤT BÁO CÁO (OVERVIEW)
  // =========================================================================

  /**
   * Lấy danh sách tất cả công ty/đối tác (Nhóm A & Nhóm B) trong dự án hiện tại
   */
  getProjectCompanies() {
    const map = new Map();
    (this.contracts || []).forEach((c) => {
      if (c.is_cpb) return;
      const names = new Set();
      const compA = String(c.company_a || '').trim();
      const compB = String(c.company_b || '').trim();
      if (compA) names.add(compA);
      if (compB) names.add(compB);

      names.forEach((name) => {
        if (!map.has(name)) {
          map.set(name, { name, count: 0, groups: new Set() });
        }
        const entry = map.get(name);
        entry.count += 1;
        if (c.group) entry.groups.add(c.group);
      });
    });

    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, 'vi'));
  }

  /**
   * Sắp xếp và phân cấp danh sách hợp đồng/phụ lục theo thời gian ký tăng dần ở từng cấp
   */
  buildCompanyContractTreeRows(contractsList) {
    const list = Array.isArray(contractsList) ? contractsList : [];
    const idSet = new Set(list.map((c) => c.id));

    const compareBySignDate = (a, b) => {
      if ((a.group || '') !== (b.group || '')) {
        return (a.group || '').localeCompare(b.group || '');
      }
      const dA = this.parsePaymentDateObj(a.signing_date || a.created_at)?.getTime();
      const dB = this.parsePaymentDateObj(b.signing_date || b.created_at)?.getTime();
      const tA = Number.isFinite(dA) ? dA : Number.MAX_SAFE_INTEGER;
      const tB = Number.isFinite(dB) ? dB : Number.MAX_SAFE_INTEGER;
      if (tA !== tB) return tA - tB;
      return String(a.contract_no || '').localeCompare(String(b.contract_no || ''), 'vi');
    };

    const childrenMap = new Map();
    const roots = [];

    list.forEach((c) => {
      if (c.parent_id && idSet.has(c.parent_id)) {
        if (!childrenMap.has(c.parent_id)) childrenMap.set(c.parent_id, []);
        childrenMap.get(c.parent_id).push(c);
      } else {
        roots.push(c);
      }
    });

    roots.sort(compareBySignDate);
    childrenMap.forEach((arr) => arr.sort(compareBySignDate));

    const result = [];
    const traverse = (node, level, prefixNo) => {
      result.push({
        contract: node,
        level,
        treeNo: prefixNo,
        isSub: !!node.parent_id,
      });
      const children = childrenMap.get(node.id) || [];
      children.forEach((ch, idx) => {
        traverse(ch, level + 1, `${prefixNo}.${idx + 1}`);
      });
    };

    roots.forEach((r, idx) => traverse(r, 0, String(idx + 1)));
    return result;
  }

  /**
   * Lấy toàn bộ hợp đồng (gồm cả HĐ chính & phụ lục) liên quan tới một công ty, đã sắp xếp theo cây & ngày ký
   */
  getContractsByCompany(companyName) {
    const target = String(companyName || '').trim().toLowerCase();
    if (!target) return [];
    const matched = (this.contracts || []).filter((c) => {
      if (c.is_cpb) return false;
      const a = String(c.company_a || '').trim().toLowerCase();
      const b = String(c.company_b || '').trim().toLowerCase();
      return a === target || b === target;
    });
    return this.buildCompanyContractTreeRows(matched).map((item) => item.contract);
  }

  /**
   * Lấy toàn bộ hợp đồng (gồm cả HĐ chính & phụ lục) liên quan tới nhiều công ty, loại bỏ trùng lặp và sắp xếp theo cây & ngày ký
   */
  getContractsByCompanies(companyNames) {
    const targets = new Set(
      Array.from(companyNames || [])
        .map((n) => String(n || '').trim().toLowerCase())
        .filter(Boolean)
    );
    if (targets.size === 0) return [];
    const matched = (this.contracts || []).filter((c) => {
      if (c.is_cpb) return false;
      const a = String(c.company_a || '').trim().toLowerCase();
      const b = String(c.company_b || '').trim().toLowerCase();
      return targets.has(a) || targets.has(b);
    });
    return this.buildCompanyContractTreeRows(matched).map((item) => item.contract);
  }

  /**
   * Gắn sự kiện 1 lần cho bộ lọc nhiều công ty ở Tab Tổng quan (giữ nguyên focus input khi gõ)
   */
  bindOverviewCompanyEvents() {
    if (this.overviewCompanyEventsBound) return;
    const searchInput = document.getElementById('overview-company-search');
    const clearBtn = document.getElementById('btn-clear-overview-company-search');
    const selectEl = document.getElementById('overview-company-select');
    const selectAllCompBtn = document.getElementById('btn-overview-select-all-companies');
    const unselectAllCompBtn = document.getElementById('btn-overview-unselect-all-companies');
    const exportBtn = document.getElementById('btn-export-company-report');
    const bodyEl = document.getElementById('overview-company-report-body');

    if (!searchInput || !selectEl || !bodyEl) return;
    this.overviewCompanyEventsBound = true;

    const addCompany = (companyName) => {
      const name = String(companyName || '').trim();
      if (!name) return;
      this.overviewSelectedCompanies.add(name);
      this.getContractsByCompany(name).forEach((c) => this.overviewSelectedContractIds.add(c.id));
      this.renderOverviewCompanyReport(true);
    };

    const removeCompany = (companyName) => {
      const name = String(companyName || '').trim();
      if (!name || !this.overviewSelectedCompanies.has(name)) return;
      this.overviewSelectedCompanies.delete(name);
      const stillCoveredIds = new Set(
        this.getContractsByCompanies(this.overviewSelectedCompanies).map((c) => c.id)
      );
      this.getContractsByCompany(name).forEach((c) => {
        if (!stillCoveredIds.has(c.id)) {
          this.overviewSelectedContractIds.delete(c.id);
        }
      });
      this.renderOverviewCompanyReport(true);
    };

    const toggleCompany = (companyName) => {
      if (this.overviewSelectedCompanies.has(companyName)) {
        removeCompany(companyName);
      } else {
        addCompany(companyName);
      }
    };

    searchInput.addEventListener('input', (e) => {
      this.overviewCompanyQuery = e.target.value || '';
      if (clearBtn) {
        clearBtn.classList.toggle('hidden', !this.overviewCompanyQuery.trim());
      }
      this.renderOverviewCompanyReport(true);
    });

    if (clearBtn) {
      clearBtn.addEventListener('click', () => {
        this.overviewCompanyQuery = '';
        searchInput.value = '';
        clearBtn.classList.add('hidden');
        this.renderOverviewCompanyReport(true);
        searchInput.focus();
      });
    }

    selectEl.addEventListener('change', (e) => {
      if (e.target.value) {
        addCompany(e.target.value);
        e.target.value = '';
      }
    });

    if (selectAllCompBtn) {
      selectAllCompBtn.addEventListener('click', () => {
        const q = (this.overviewCompanyQuery || '').trim().toLowerCase();
        const companies = this.getProjectCompanies();
        const matched = q ? companies.filter((c) => c.name.toLowerCase().includes(q)) : companies;
        matched.forEach((c) => this.overviewSelectedCompanies.add(c.name));
        this.getContractsByCompanies(this.overviewSelectedCompanies).forEach((c) =>
          this.overviewSelectedContractIds.add(c.id)
        );
        this.renderOverviewCompanyReport(true);
      });
    }

    if (unselectAllCompBtn) {
      unselectAllCompBtn.addEventListener('click', () => {
        this.overviewSelectedCompanies.clear();
        this.overviewSelectedContractIds.clear();
        this.renderOverviewCompanyReport(true);
      });
    }

    if (exportBtn) {
      exportBtn.addEventListener('click', () => {
        this.exportCompanyCashflowReportCsv();
      });
    }

    bodyEl.addEventListener('click', (e) => {
      const removeChipBtn = e.target.closest('.btn-remove-selected-company');
      if (removeChipBtn) {
        removeCompany(removeChipBtn.dataset.company || '');
        return;
      }
      const chipBtn = e.target.closest('.overview-company-chip');
      if (chipBtn) {
        toggleCompany(chipBtn.dataset.company || '');
        return;
      }
      if (e.target.closest('#btn-overview-select-all-contracts')) {
        const related = this.getContractsByCompanies(this.overviewSelectedCompanies);
        this.overviewSelectedContractIds = new Set(related.map((c) => c.id));
        this.renderOverviewCompanyReport(true);
        return;
      }
      if (e.target.closest('#btn-overview-unselect-all-contracts')) {
        this.overviewSelectedContractIds = new Set();
        this.renderOverviewCompanyReport(true);
        return;
      }
    });

    bodyEl.addEventListener('change', (e) => {
      const cb = e.target.closest('.overview-contract-checkbox');
      if (!cb) return;
      const cid = cb.dataset.contractId;
      if (!cid) return;
      if (cb.checked) {
        this.overviewSelectedContractIds.add(cid);
      } else {
        this.overviewSelectedContractIds.delete(cid);
      }
      this.renderOverviewCompanyReport(true);
    });

    bodyEl.addEventListener('dblclick', (e) => {
      if (e.target.closest('input[type="checkbox"]')) return;
      const row = e.target.closest('.overview-cashflow-contract-row, .overview-contract-check-item');
      const cid = row?.dataset?.contractId;
      if (cid) {
        this.openContractDetail(cid);
      }
    });
  }

  /**
   * Render khu vực Tra cứu Công ty & Tổng hợp Dòng tiền trên Tab Tổng quan
   */
  renderOverviewCompanyReport(skipInputSync = false) {
    this.bindOverviewCompanyEvents();

    const searchInput = document.getElementById('overview-company-search');
    const clearBtn = document.getElementById('btn-clear-overview-company-search');
    const selectEl = document.getElementById('overview-company-select');
    const exportBtn = document.getElementById('btn-export-company-report');
    const bodyEl = document.getElementById('overview-company-report-body');
    if (!bodyEl || !selectEl) return;

    const companies = this.getProjectCompanies();
    const companyNameSet = new Set(companies.map((c) => c.name));

    // Loại bỏ các công ty không còn tồn tại nếu đổi dự án
    let removedAny = false;
    Array.from(this.overviewSelectedCompanies).forEach((name) => {
      if (!companyNameSet.has(name)) {
        this.overviewSelectedCompanies.delete(name);
        removedAny = true;
      }
    });
    if (removedAny) {
      const validIds = new Set(this.getContractsByCompanies(this.overviewSelectedCompanies).map((c) => c.id));
      this.overviewSelectedContractIds = new Set(
        Array.from(this.overviewSelectedContractIds).filter((id) => validIds.has(id))
      );
    }

    if (!skipInputSync && searchInput) {
      searchInput.value = this.overviewCompanyQuery || '';
      if (clearBtn) {
        clearBtn.classList.toggle('hidden', !(this.overviewCompanyQuery || '').trim());
      }
    }

    const q = (this.overviewCompanyQuery || '').trim().toLowerCase();
    const matchedCompanies = q
      ? companies.filter((c) => c.name.toLowerCase().includes(q))
      : companies;

    // Cập nhật dropdown thêm nhanh công ty
    selectEl.innerHTML =
      `<option value="">-- Thêm công ty vào danh sách (${matchedCompanies.length}) --</option>` +
      matchedCompanies
        .map((c) => {
          const isChosen = this.overviewSelectedCompanies.has(c.name);
          return `<option value="${c.name.replace(/"/g, '&quot;')}">${isChosen ? '✓ ' : ''}${c.name} (${c.count} HĐ)</option>`;
        })
        .join('');

    let html = '';
    if (this.overviewSelectedCompanies.size === 0) {
      html = matchedCompanies.length === 0 && q
        ? `
          <div class="overview-company-empty">
            <i class="fa-solid fa-magnifying-glass-minus"></i>
            <span>Không tìm thấy công ty nào khớp với "<strong>${q}</strong>".</span>
          </div>
        `
        : `
          <div class="overview-company-empty">
            <i class="fa-solid fa-building-user"></i>
            <span>Tìm kiếm hoặc chọn công ty từ danh sách, hoặc bấm <strong>"Chọn tất cả"</strong> để xem hợp đồng và tổng hợp dòng tiền.</span>
          </div>
        `;
      bodyEl.innerHTML = html;
      if (exportBtn) exportBtn.disabled = true;
      return;
    }

    const selectedCompanyList = Array.from(this.overviewSelectedCompanies);
    const companyContracts = this.getContractsByCompanies(this.overviewSelectedCompanies);
    const selectedContracts = companyContracts.filter((c) => this.overviewSelectedContractIds.has(c.id));
    const isMultiCompany = selectedCompanyList.length > 1;

    if (exportBtn) {
      exportBtn.disabled = selectedContracts.length === 0;
    }

    // Banner các công ty đang chọn & danh sách hợp đồng có checkbox
    html += `
      <div class="overview-selected-company-panel">
        <div class="overview-selected-company-top">
          <div class="overview-selected-company-title">
            <i class="fa-solid fa-building" style="color: var(--primary);"></i>
            <strong>${isMultiCompany ? `Đã chọn ${selectedCompanyList.length} công ty` : selectedCompanyList[0]}</strong>
            <span class="badge badge-primary">${companyContracts.length} Hợp đồng / Phụ lục</span>
            <span class="badge badge-success">Đã chọn: ${selectedContracts.length}</span>
          </div>
          <div style="display: flex; gap: 6px; flex-wrap: wrap;">
            <button type="button" id="btn-overview-select-all-contracts" class="button button-outline button-sm">
              <i class="fa-solid fa-check-double"></i> Chọn tất cả HĐ
            </button>
            <button type="button" id="btn-overview-unselect-all-contracts" class="button button-outline button-sm">
              <i class="fa-regular fa-square"></i> Bỏ chọn HĐ
            </button>
          </div>
        </div>

        ${
          isMultiCompany
            ? `<div style="display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px;">
                ${selectedCompanyList
                  .map(
                    (name) => `
                  <span class="badge badge-primary" style="display: inline-flex; align-items: center; gap: 5px; font-size: 0.74rem; padding: 3px 8px;">
                    ${name}
                    <button type="button" class="btn-remove-selected-company" data-company="${name.replace(/"/g, '&quot;')}" style="background: none; border: none; color: inherit; cursor: pointer; padding: 0; font-size: 0.82rem; line-height: 1;" title="Bỏ chọn công ty này">&times;</button>
                  </span>
                `
                  )
                  .join('')}
              </div>`
            : ''
        }

        ${(() => {
          const compA = companyContracts.filter((c) => (c.group || 'B') === 'A');
          const compB = companyContracts.filter((c) => (c.group || 'B') !== 'A');
          const hasBothGroups = compA.length > 0 && compB.length > 0;
          const renderChecklistItems = (list) =>
            this.buildCompanyContractTreeRows(list)
              .map(({ contract: c, level, treeNo, isSub }) => {
                const checked = this.overviewSelectedContractIds.has(c.id) ? 'checked' : '';
                const st = this.calculateTotals(c);
                const signDateStr = c.signing_date ? this.formatDate(c.signing_date) : '—';
                return `
                  <label class="overview-contract-check-item ${checked ? 'is-checked' : ''}" data-contract-id="${c.id}" title="Nhấp đúp để mở popup chi tiết hợp đồng" style="${level > 0 ? `border-left: 3px solid var(--primary);` : ''}">
                    <input type="checkbox" class="overview-contract-checkbox" data-contract-id="${c.id}" ${checked}>
                    <div class="overview-contract-check-info">
                      <div class="overview-contract-check-row1">
                        <span class="badge" style="font-size: 0.66rem; background: rgba(59,130,246,0.12); color: var(--primary);">${treeNo}</span>
                        <span class="badge ${c.group === 'A' ? 'badge-primary' : 'badge-warning'}" style="font-size: 0.68rem;">Nhóm ${c.group || 'B'}</span>
                        <strong class="overview-contract-no">${level > 0 ? '↳ ' : ''}${c.contract_no || c.id}</strong>
                        ${isSub ? '<span class="badge" style="font-size: 0.65rem; background: #e2e8f0; color: #475569;">Phụ lục</span>' : ''}
                        <span class="overview-contract-val">${this.formatCurrency(st.totalSigned)}</span>
                      </div>
                      <div class="overview-contract-scope">${c.signing_date ? `[${signDateStr}] ` : ''}${c.scope_of_work || '—'}</div>
                    </div>
                  </label>
                `;
              })
              .join('');

          if (!hasBothGroups) {
            return `<div class="overview-contract-checklist">${renderChecklistItems(companyContracts)}</div>`;
          }
          return `
            <div style="margin-bottom: 10px;">
              <div style="font-size: 0.78rem; font-weight: 700; color: var(--primary); margin-bottom: 6px; display: flex; align-items: center; gap: 6px;">
                <span class="badge badge-primary" style="font-size: 0.68rem;"><i class="fa-solid fa-landmark"></i> NHÓM A — THU (${compA.length})</span>
              </div>
              <div class="overview-contract-checklist">${renderChecklistItems(compA)}</div>
            </div>
            <div>
              <div style="font-size: 0.78rem; font-weight: 700; color: #d97706; margin-bottom: 6px; display: flex; align-items: center; gap: 6px;">
                <span class="badge badge-warning" style="font-size: 0.68rem;"><i class="fa-solid fa-handshake"></i> NHÓM B — CHI (${compB.length})</span>
              </div>
              <div class="overview-contract-checklist">${renderChecklistItems(compB)}</div>
            </div>
          `;
        })()}
      </div>
    `;

    // Bảng tổng hợp thông tin dòng tiền cho các hợp đồng đã chọn (phân tách rõ Nhóm A & Nhóm B)
    if (selectedContracts.length === 0) {
      html += `
        <div class="overview-company-empty" style="margin-top: 10px;">
          <i class="fa-regular fa-square-check"></i>
          <span>Vui lòng tích chọn ít nhất 1 hợp đồng ở trên để xem tổng hợp dòng tiền và xuất báo cáo.</span>
        </div>
      `;
    } else {
      const selectedA = selectedContracts.filter((c) => (c.group || 'B') === 'A');
      const selectedB = selectedContracts.filter((c) => (c.group || 'B') !== 'A');
      const hasBothSelected = selectedA.length > 0 && selectedB.length > 0;

      const buildGroupBlock = (groupContracts) => {
        const totals = { signed: 0, actual: 0, accepted: 0, requested: 0, paid: 0, remaining: 0, ipc: 0, count: groupContracts.length };
        const treeRows = this.buildCompanyContractTreeRows(groupContracts);
        const rowsHtml = treeRows
          .map(({ contract: c, level, treeNo, isSub }) => {
            const st = this.calculateTotals(c);
            const realPayments = this.getRealPayments(c.id);
            const actualVnd = realPayments.reduce((s, p) => s + this.getPaymentVnd(p, 'actual_volume', 'actual_vol'), 0);

            totals.signed += st.totalSigned;
            totals.actual += actualVnd;
            totals.accepted += st.totalAccepted;
            totals.requested += st.totalRequested;
            totals.paid += st.totalPaid;
            totals.remaining += st.remaining;
            totals.ipc += st.ipcCount;

            const signDateText = c.signing_date ? ` • Ký: ${this.formatDate(c.signing_date)}` : '';
            const partnerTag = isMultiCompany ? ` • ${c.group === 'A' ? c.company_a : c.company_b || c.company_a || ''}` : '';
            const indentPx = level * 18;

            return `
              <tr class="overview-cashflow-contract-row ${level === 0 ? 'overview-tree-parent-row' : 'overview-tree-child-row'}" data-contract-id="${c.id}" style="cursor: pointer;" title="Nhấp đúp để mở popup chi tiết & chỉnh sửa hợp đồng">
                <td style="text-align: center; font-weight: ${level === 0 ? '700' : '600'}; white-space: nowrap;">${treeNo}</td>
                <td style="padding-left: ${10 + indentPx}px;">
                  <div style="font-weight: 700; color: var(--primary); display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
                    ${level > 0 ? '<span style="color: var(--text-muted); font-weight: 400;">↳</span>' : ''}
                    <span>${c.contract_no || c.id}</span>
                    ${isSub ? '<span class="badge" style="font-size: 0.65rem; background: #e2e8f0; color: #475569;">Phụ lục</span>' : ''}
                    ${signDateText ? `<span style="font-size: 0.72rem; font-weight: 500; color: var(--text-muted);">${signDateText}</span>` : ''}
                  </div>
                  <div style="font-size: 0.74rem; color: var(--text-muted);">${c.scope_of_work || '—'}${partnerTag}</div>
                </td>
                <td style="text-align: center;">
                  <span class="badge ${c.group === 'A' ? 'badge-primary' : 'badge-warning'}" style="font-size: 0.68rem;">Nhóm ${c.group || 'B'}</span>
                </td>
                <td style="text-align: right; font-weight: 600;">${this.formatMoneyPlain(st.totalSigned)}</td>
                <td style="text-align: right;">${this.formatMoneyPlain(actualVnd)}</td>
                <td style="text-align: right;">${this.formatMoneyPlain(st.totalAccepted)}</td>
                <td style="text-align: right;">${this.formatMoneyPlain(st.totalRequested)}</td>
                <td style="text-align: right; font-weight: 700; color: var(--success);">${this.formatMoneyPlain(st.totalPaid)}</td>
                <td style="text-align: right; color: ${st.remaining > 0 ? 'var(--danger)' : 'var(--text-muted)'}; font-weight: 600;">${this.formatMoneyPlain(st.remaining)}</td>
                <td style="text-align: center;">${st.ipcCount} đợt</td>
              </tr>
            `;
          })
          .join('');
        return { rowsHtml, totals };
      };

      const blockA = selectedA.length > 0 ? buildGroupBlock(selectedA) : null;
      const blockB = selectedB.length > 0 ? buildGroupBlock(selectedB) : null;

      let tbodyHtml = '';
      let tfootHtml = '';

      if (hasBothSelected && blockA && blockB) {
        const formatSubtotalRow = (label, badgeClass, tot) => `
          <tr class="overview-group-subtotal-row" style="background: rgba(59, 130, 246, 0.06); font-weight: 700;">
            <td colspan="3" style="font-weight: 700;">
              <span class="badge ${badgeClass}" style="font-size: 0.72rem;">${label}</span>
            </td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(tot.signed)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(tot.actual)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(tot.accepted)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(tot.requested)}</td>
            <td style="text-align: right; font-weight: 700; color: var(--success);">${this.formatMoneyPlain(tot.paid)}</td>
            <td style="text-align: right; font-weight: 700; color: ${tot.remaining > 0 ? 'var(--danger)' : 'inherit'};">${this.formatMoneyPlain(tot.remaining)}</td>
            <td style="text-align: center; font-weight: 700;">${tot.ipc} đợt</td>
          </tr>
        `;

        tbodyHtml = `
          <tr class="overview-group-banner-row" style="background: rgba(37, 99, 235, 0.08);">
            <td colspan="10" style="font-weight: 700; color: var(--primary); font-size: 0.8rem;">
              <i class="fa-solid fa-landmark"></i> NHÓM A — HỢP ĐỒNG CHỦ ĐẦU TƯ / THU (${selectedA.length} hợp đồng)
            </td>
          </tr>
          ${blockA.rowsHtml}
          ${formatSubtotalRow(`TỔNG NHÓM A — THU (${selectedA.length} HĐ)`, 'badge-primary', blockA.totals)}
          <tr class="overview-group-banner-row" style="background: rgba(245, 158, 11, 0.09);">
            <td colspan="10" style="font-weight: 700; color: #d97706; font-size: 0.8rem;">
              <i class="fa-solid fa-handshake"></i> NHÓM B — HỢP ĐỒNG THẦU PHỤ / NCC / CHI (${selectedB.length} hợp đồng)
            </td>
          </tr>
          ${blockB.rowsHtml}
          ${formatSubtotalRow(`TỔNG NHÓM B — CHI (${selectedB.length} HĐ)`, 'badge-warning', blockB.totals)}
        `;

        const diffSigned = blockA.totals.signed - blockB.totals.signed;
        const diffActual = blockA.totals.actual - blockB.totals.actual;
        const diffAccepted = blockA.totals.accepted - blockB.totals.accepted;
        const diffRequested = blockA.totals.requested - blockB.totals.requested;
        const diffPaid = blockA.totals.paid - blockB.totals.paid;
        const diffRemaining = blockA.totals.remaining - blockB.totals.remaining;
        const fmtDiff = (v) => `${v > 0 ? '+' : ''}${this.formatMoneyPlain(v)}`;

        tfootHtml = `
          <tr>
            <td colspan="3" style="font-weight: 700;">
              <i class="fa-solid fa-scale-balanced" style="color: var(--success); margin-right: 4px;"></i>
              CHÊNH LỆCH DÒNG TIỀN (THU A - CHI B)
            </td>
            <td style="text-align: right; font-weight: 700; color: ${diffSigned >= 0 ? 'var(--success)' : 'var(--danger)'};">${fmtDiff(diffSigned)}</td>
            <td style="text-align: right; font-weight: 700;">${fmtDiff(diffActual)}</td>
            <td style="text-align: right; font-weight: 700;">${fmtDiff(diffAccepted)}</td>
            <td style="text-align: right; font-weight: 700;">${fmtDiff(diffRequested)}</td>
            <td style="text-align: right; font-weight: 700; color: ${diffPaid >= 0 ? 'var(--success)' : 'var(--danger)'};">${fmtDiff(diffPaid)}</td>
            <td style="text-align: right; font-weight: 700;">${fmtDiff(diffRemaining)}</td>
            <td style="text-align: center; font-weight: 700;">${blockA.totals.ipc + blockB.totals.ipc} đợt</td>
          </tr>
        `;
      } else {
        const activeBlock = blockA || blockB;
        const groupTitle = blockA ? `NHÓM A — THU (${selectedA.length} hợp đồng)` : `NHÓM B — CHI (${selectedB.length} hợp đồng)`;
        tbodyHtml = activeBlock.rowsHtml;
        tfootHtml = `
          <tr>
            <td colspan="3" style="font-weight: 700;">TỔNG CỘNG ${groupTitle}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(activeBlock.totals.signed)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(activeBlock.totals.actual)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(activeBlock.totals.accepted)}</td>
            <td style="text-align: right; font-weight: 700;">${this.formatMoneyPlain(activeBlock.totals.requested)}</td>
            <td style="text-align: right; font-weight: 700; color: var(--success);">${this.formatMoneyPlain(activeBlock.totals.paid)}</td>
            <td style="text-align: right; font-weight: 700; color: ${activeBlock.totals.remaining > 0 ? 'var(--danger)' : 'inherit'};">${this.formatMoneyPlain(activeBlock.totals.remaining)}</td>
            <td style="text-align: center; font-weight: 700;">${activeBlock.totals.ipc} đợt</td>
          </tr>
        `;
      }

      html += `
        <div class="overview-cashflow-summary-wrap">
          <div class="overview-cashflow-summary-title">
            <span><i class="fa-solid fa-money-bill-trend-up" style="color: var(--success);"></i> TỔNG HỢP THÔNG TIN DÒNG TIỀN (${selectedContracts.length} hợp đồng đã chọn)</span>
            <span style="font-size: 0.76rem; color: var(--text-muted);">Đvt: VNĐ</span>
          </div>
          <div class="data-table-wrap" style="overflow-x: auto;">
            <table class="data-table overview-cashflow-table" style="width: 100%; margin: 0;">
              <thead>
                <tr>
                  <th style="width: 44px; text-align: center;">TT</th>
                  <th style="min-width: 220px;">Số hiệu HĐ & Nội dung</th>
                  <th style="width: 80px; text-align: center;">Nhóm</th>
                  <th style="text-align: right;">Giá trị Hợp đồng</th>
                  <th style="text-align: right;">KL hoàn thành</th>
                  <th style="text-align: right;">KL nghiệm thu</th>
                  <th style="text-align: right;">Đề nghị TT</th>
                  <th style="text-align: right;">Đã thanh toán / Tạm ứng</th>
                  <th style="text-align: right;">Còn lại / Chênh lệch</th>
                  <th style="width: 80px; text-align: center;">Số đợt</th>
                </tr>
              </thead>
              <tbody>
                ${tbodyHtml}
              </tbody>
              <tfoot>
                ${tfootHtml}
              </tfoot>
            </table>
          </div>
        </div>
      `;
    }

    bodyEl.innerHTML = html;
  }

  /**
   * Xuất Báo cáo Tình hình Thực hiện Hợp đồng và Thanh toán theo mẫu Excel (.xlsx) định dạng chuẩn đẹp
   */
  async exportCompanyCashflowReportCsv() {
    if (!this.overviewSelectedCompanies || this.overviewSelectedCompanies.size === 0) return;
    const selectedCompanyList = Array.from(this.overviewSelectedCompanies);
    const companyContracts = this.getContractsByCompanies(this.overviewSelectedCompanies);
    const selectedContracts = companyContracts.filter((c) => this.overviewSelectedContractIds.has(c.id));
    if (selectedContracts.length === 0) return;

    if (!window.ExcelJS) {
      this.showAppAlert?.('Thư viện xuất báo cáo Excel (ExcelJS) chưa sẵn sàng. Vui lòng tải lại trang.');
      return;
    }

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId);
    const rawPrjName = String(curPrj?.name || this.currentProjectId || 'DỰ ÁN').trim();
    const prjUpper = rawPrjName.toUpperCase();
    const prjTitle = prjUpper.startsWith('DỰ ÁN') ? prjUpper : `DỰ ÁN ${prjUpper}`;
    const isMultiCompany = selectedCompanyList.length > 1;
    const allCompaniesCount = this.getProjectCompanies().length;
    const isAllCompanies = selectedCompanyList.length >= allCompaniesCount && allCompaniesCount > 1;
    const partnerLabel = !isMultiCompany
      ? `ĐỐI TÁC ${selectedCompanyList[0].trim().toUpperCase()}`
      : isAllCompanies
      ? `TẤT CẢ ĐỐI TÁC (${selectedCompanyList.length} CÔNG TY)`
      : `TỔNG HỢP ${selectedCompanyList.length} ĐỐI TÁC (${selectedCompanyList.slice(0, 3).map((n) => n.trim().toUpperCase()).join(', ')}${selectedCompanyList.length > 3 ? ', ...' : ''})`;

    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'MAIO CPMS';
    wb.created = new Date();

    const ws = wb.addWorksheet('BaoCao_DongTien', {
      pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    ws.columns = [
      { key: 'tt', width: 8 },
      { key: 'content', width: 56 },
      { key: 'contract_val', width: 20 },
      { key: 'actual_vol', width: 20 },
      { key: 'accepted_vol', width: 20 },
      { key: 'request_amt', width: 20 },
      { key: 'paid_amt', width: 20 },
      { key: 'remaining', width: 20 },
      { key: 'invoice', width: 16 },
      { key: 'note', width: 26 },
    ];

    const baseFont = { name: 'Times New Roman', size: 11 };
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };

    // Dòng 1: Tiêu đề báo cáo
    const titleRow = ws.addRow([
      `BÁO CÁO TÌNH HÌNH THỰC HIỆN HỢP ĐỒNG VÀ THANH TOÁN CỦA ${prjTitle} - ${partnerLabel}`,
    ]);
    ws.mergeCells('A1:J1');
    titleRow.height = 32;
    titleRow.getCell(1).font = { ...baseFont, size: 13, bold: true, color: { argb: 'FF0F172A' } };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // Dòng 2: Đơn vị tính
    const unitRow = ws.addRow(['', '', '', '', '', '', '', '', '', 'Đvt: VNĐ']);
    unitRow.height = 18;
    unitRow.getCell(10).font = { ...baseFont, size: 10, italic: true };
    unitRow.getCell(10).alignment = { horizontal: 'right', vertical: 'middle' };

    // Dòng 3: Header bảng
    const headers = [
      'TT',
      'Nội dung',
      'Giá trị Hợp đồng\n(đã bao gồm VAT)',
      'Giá trị sản lượng thực hiện\n(đã bao gồm VAT)',
      'Giá trị sản lượng đã được nghiệm thu\n(đã bao gồm VAT)',
      'Giá trị đề nghị thanh toán / tạm ứng\n(đã bao gồm VAT)',
      'Giá trị đã được thanh toán / tạm ứng\n(đã bao gồm VAT)',
      'Chênh lệch sau khi giải ngân\n(VNĐ - đã bao gồm VAT)',
      'Số hóa đơn VAT',
      'Ghi chú',
    ];
    const headerRow = ws.addRow(headers);
    headerRow.height = 44;
    headerRow.eachCell((cell) => {
      cell.font = { ...baseFont, size: 10.5, bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    });

    const estimateRowHeight = (text, colChars = 50, minH = 20) => {
      const s = String(text || '');
      const lines = Math.max(1, Math.ceil(s.length / colChars));
      return Math.max(minH, lines * 15 + 6);
    };

    const styleDataRow = (row, { isContract = false, isRoot = false, isTotal = false } = {}) => {
      const contentLen = String(row.getCell(2).value || '').length;
      row.height = isTotal ? 24 : estimateRowHeight(row.getCell(2).value, 50, isContract ? 24 : 20);
      for (let c = 1; c <= 10; c++) {
        const cell = row.getCell(c);
        cell.font = {
          ...baseFont,
          size: isTotal ? 11 : 10.5,
          bold: isTotal || isContract,
          color: { argb: isRoot ? 'FF1E3A8A' : 'FF0F172A' },
        };
        if (isTotal) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        } else if (isRoot) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF2FF' } };
        } else if (isContract) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }
        cell.border = thinBorder;

        if (c === 1) {
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        } else if (c >= 3 && c <= 8) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          if (typeof cell.value === 'number') {
            cell.numFmt = '#,##0';
          }
        } else if (c === 9) {
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
          if (c === 10 && contentLen > 0) {
            row.height = Math.max(row.height, estimateRowHeight(cell.value, 24, 20));
          }
        }
      }
    };

    const selectedA = selectedContracts.filter((c) => (c.group || 'B') === 'A');
    const selectedB = selectedContracts.filter((c) => (c.group || 'B') !== 'A');
    const hasBothGroups = selectedA.length > 0 && selectedB.length > 0;
    const selectedIdSet = new Set(selectedContracts.map((item) => item.id));

    const writeGroupSection = (groupContracts, sectionTitle, totalLabel) => {
      let gSigned = 0;
      let gActual = 0;
      let gAccepted = 0;
      let gRequested = 0;
      let gPaid = 0;
      let gRemaining = 0;
      let gRootCount = 0;

      if (sectionTitle) {
        const secRow = ws.addRow(['', sectionTitle]);
        ws.mergeCells(`B${secRow.number}:J${secRow.number}`);
        styleDataRow(secRow, { isContract: true, isRoot: true });
      }

      const treeRows = this.buildCompanyContractTreeRows(groupContracts);
      treeRows.forEach(({ contract: c, level, treeNo }) => {
        if (level === 0) gRootCount += 1;
        const st = this.calculateTotals(c);
        const realPayments = this.getRealPayments(c.id);
        const actualVnd = realPayments.reduce((s, p) => s + this.getPaymentVnd(p, 'actual_volume', 'actual_vol'), 0);
        const displayPayments = realPayments.filter(
          (p) => p.contract_id === c.id || !selectedIdSet.has(p.contract_id)
        );

        gSigned += st.totalSigned;
        gActual += actualVnd;
        gAccepted += st.totalAccepted;
        gRequested += st.totalRequested;
        gPaid += st.totalPaid;
        gRemaining += st.remaining;

        const signDateStr = c.signing_date ? ` ký ngày ${this.formatDate(c.signing_date)}` : '';
        const indentPrefix = level > 0 ? `${'  '.repeat(level)}↳ ` : '';
        const kindLabel = level > 0 ? 'Phụ lục số' : 'Hợp đồng số';
        const partnerSuffix = isMultiCompany ? ` [${c.group === 'A' ? c.company_a : c.company_b || c.company_a || ''}]` : '';
        const contractTitle = `${indentPrefix}${kindLabel}: ${c.contract_no || c.id}${signDateStr}${partnerSuffix}${c.scope_of_work ? ' - ' + c.scope_of_work : ''}`;

        const contractRow = ws.addRow([
          treeNo,
          contractTitle,
          st.totalSigned,
          actualVnd,
          st.totalAccepted,
          st.totalRequested,
          st.totalPaid,
          st.remaining,
          '',
          c.status || '',
        ]);
        styleDataRow(contractRow, { isContract: true, isRoot: level === 0 });

        displayPayments.forEach((p, pIdx) => {
          const pAct = this.getPaymentVnd(p, 'actual_volume', 'actual_vol');
          const pAcc = this.getPaymentVnd(p, 'accepted_vol');
          const pReq = this.getPaymentVnd(p, 'request_amount');
          const pPaid = this.getPaymentVnd(p, 'paid_amount');
          const pRem = pReq - pPaid;
          const pLabel = this.getPaymentContent(p) || `Đợt thanh toán ${pIdx + 1}`;
          const pDateStr = p.payment_date ? ` (${this.formatDate(p.payment_date)})` : '';

          const payRow = ws.addRow([
            `${treeNo}.${pIdx + 1}`,
            `${'  '.repeat(level + 1)}${pLabel}${pDateStr}`,
            null,
            pAct || null,
            pAcc || null,
            pReq || null,
            pPaid || null,
            pRem || null,
            p.invoice_ref || '',
            p.note || '',
          ]);
          styleDataRow(payRow);
        });
      });

      const groupTotalRow = ws.addRow([
        gRootCount + 1,
        totalLabel,
        gSigned,
        gActual,
        gAccepted,
        gRequested,
        gPaid,
        gRemaining,
        '',
        '',
      ]);
      styleDataRow(groupTotalRow, { isTotal: true });

      return { gSigned, gActual, gAccepted, gRequested, gPaid, gRemaining };
    };

    if (hasBothGroups) {
      const totA = writeGroupSection(selectedA, 'NHÓM A — HỢP ĐỒNG CHỦ ĐẦU TƯ (THU)', 'Tổng cộng Nhóm A (Thu)');
      const totB = writeGroupSection(selectedB, 'NHÓM B — HỢP ĐỒNG THẦU PHỤ / NCC (CHI)', 'Tổng cộng Nhóm B (Chi)');
      const diffRow = ws.addRow([
        '',
        'Chênh lệch dòng tiền (Thu A - Chi B)',
        totA.gSigned - totB.gSigned,
        totA.gActual - totB.gActual,
        totA.gAccepted - totB.gAccepted,
        totA.gRequested - totB.gRequested,
        totA.gPaid - totB.gPaid,
        totA.gRemaining - totB.gRemaining,
        '',
        '',
      ]);
      styleDataRow(diffRow, { isTotal: true });
    } else if (selectedA.length > 0) {
      writeGroupSection(selectedA, '', 'Tổng cộng Nhóm A (Thu)');
    } else {
      writeGroupSection(selectedB, '', 'Tổng cộng Nhóm B (Chi)');
    }

    ws.addRow([]);
    const signRow = ws.addRow([
      '',
      '',
      'LÃNH ĐẠO PHÊ DUYỆT',
      '',
      '',
      'BAN ĐẦU TƯ VÀ QUẢN LÝ DỰ ÁN',
      '',
      '',
      '',
      'Người đề nghị',
    ]);
    const signRowIdx = signRow.number;
    ws.mergeCells(`C${signRowIdx}:D${signRowIdx}`);
    ws.mergeCells(`F${signRowIdx}:G${signRowIdx}`);
    signRow.height = 24;
    [3, 6, 10].forEach((colIdx) => {
      const cell = signRow.getCell(colIdx);
      cell.font = { ...baseFont, size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    const asciiComp = !isMultiCompany
      ? selectedCompanyList[0]
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/đ/g, 'd')
          .replace(/Đ/g, 'D')
          .replace(/[^a-zA-Z0-9]+/g, '_')
          .replace(/^_+|_+$/g, '')
          .slice(0, 36) || 'DoiTac'
      : isAllCompanies
      ? 'TatCa_DoiTac'
      : `${selectedCompanyList.length}_DoiTac`;
    const today = new Date().toISOString().slice(0, 10);
    const fileName = `BaoCao_DongTien_${asciiComp}_${today}.xlsx`;

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);

    this.showLocalSaveStatus(`Đã xuất báo cáo Excel "${fileName}" (${selectedContracts.length} hợp đồng)!`);
  }
}

if (!window.__MAIO_CONTRACTS_MANAGER__) {
  window.__MAIO_CONTRACTS_MANAGER__ = new ContractsManager();
}
export const contractsManager = window.__MAIO_CONTRACTS_MANAGER__;
