/**
 * MAIO Application Entry Point
 * Quản lý khởi tạo IndexedDB, nạp dữ liệu hạt giống CHC1, điều hướng phân hệ và kết nối đồng bộ.
 */

import { gdrive } from './core/gdrive.js';
import { db } from './core/db.js';
import { CHC1_SEED_DATA } from './data/chc1_seed.js';
import { contractsManager } from './modules/cpms/contracts.js';

class App {
  constructor() {
    this.currentView = 'dashboard';
    this.localDirectoryHandle = null;
  }

  async init() {
    console.log('Khởi động MAIO Application...');

    // 1. Khởi tạo IndexedDB
    try {
      await db.init();
      console.log('IndexedDB đã sẵn sàng.');

      // 2. Nạp dữ liệu hạt giống (Seed Data) nếu chưa có
      await this.ensureSeedData();
    } catch (e) {
      console.warn('Lỗi khởi tạo IndexedDB:', e);
    }

    // 3. Khởi tạo CPMS Contracts Manager
    try {
      await contractsManager.init();
    } catch (e) {
      console.warn('Lỗi khởi tạo CPMS Manager:', e);
    }

    // 4. Thiết lập điều hướng (Navigation)
    this.bindNavigation();

    // 5. Khởi tạo dịch vụ Google Drive
    this.setupGoogleDrive();

    // 6. Thiết lập chọn thư mục máy tính (Local Folder)
    this.setupLocalDirectoryPicker();

    // 7. Cập nhật số liệu Dashboard
    await this.updateDashboardStats();

    // 8. Cập nhật cấu hình lên màn hình Settings
    this.renderSettingsInfo();
  }

  async ensureSeedData() {
    const SEED_VERSION = 'v2_strict_hierarchy';
    const currentVersion = localStorage.getItem('maio_seed_version');

    if (currentVersion !== SEED_VERSION && CHC1_SEED_DATA) {
      console.log('Đang nạp lại dữ liệu chuẩn xác dự án CHC1 (Strict Hierarchy & Anti-Double-Counting)...');
      
      // Xóa dữ liệu cũ
      const oldContracts = await db.getAll('contracts');
      for (const c of oldContracts) await db.delete('contracts', c.id);
      const oldPayments = await db.getAll('payments');
      for (const p of oldPayments) await db.delete('payments', p.id);
      const oldDocs = await db.getAll('documents');
      for (const d of oldDocs) await db.delete('documents', d.id);

      // Nạp dự án
      await db.put('projects', CHC1_SEED_DATA.project);

      // Nạp hợp đồng
      for (const c of CHC1_SEED_DATA.contracts) {
        await db.put('contracts', c);
      }

      // Nạp thanh toán (đã có cờ is_header chuẩn)
      for (const p of CHC1_SEED_DATA.payments) {
        await db.put('payments', p);
      }

      // Nạp hồ sơ
      for (const d of CHC1_SEED_DATA.documents) {
        await db.put('documents', d);
      }

      localStorage.setItem('maio_seed_version', SEED_VERSION);
      console.log('Nạp hoàn tất dữ liệu chuẩn xác v2.');
    }
  }

  async updateDashboardStats() {
    if (!contractsManager) return;
    await contractsManager.loadData();

    const curPrjId = contractsManager.currentProjectId;
    const curPrj = contractsManager.projects.find((p) => p.id === curPrjId);

    // Cập nhật tên dự án hiện tại trên Dashboard
    const activePrjName = document.getElementById('active-project-name');
    if (activePrjName) {
      activePrjName.textContent = curPrj?.name || curPrj?.id || 'Dự án CHC1 (Thiết kế & Thi công)';
    }

    const contracts = contractsManager.contracts;
    const realPayments = contractsManager.payments.filter((p) => !p.is_header);
    const docs = contractsManager.documents;

    // 1. Nhóm A (Chủ đầu tư / Tổng thầu)
    const mainA = contracts.filter((c) => !c.parent_id && c.group === 'A');
    const appA = contracts.filter((c) => c.parent_id && c.group === 'A');
    let totalSignedA = 0;
    let totalPaidA = 0;
    mainA.forEach((c) => {
      const st = contractsManager.calculateTotals(c);
      totalSignedA += st.totalSigned;
      totalPaidA += st.totalPaid;
    });

    const familyAIds = new Set();
    mainA.forEach((c) => {
      contractsManager.getAllDescendantIds(c.id).forEach((id) => familyAIds.add(id));
    });
    const paymentsA = realPayments.filter((p) => familyAIds.has(p.contract_id));

    // 2. Nhóm B (Nhà thầu phụ / Nhà cung cấp)
    const mainB = contracts.filter((c) => !c.parent_id && c.group === 'B');
    const appB = contracts.filter((c) => c.parent_id && c.group === 'B');
    let totalSignedB = 0;
    let totalPaidB = 0;
    mainB.forEach((c) => {
      const st = contractsManager.calculateTotals(c);
      totalSignedB += st.totalSigned;
      totalPaidB += st.totalPaid;
    });

    const familyBIds = new Set();
    mainB.forEach((c) => {
      contractsManager.getAllDescendantIds(c.id).forEach((id) => familyBIds.add(id));
    });
    const paymentsB = realPayments.filter((p) => familyBIds.has(p.contract_id));

    // 3. Gán lên giao diện Dashboard
    const el = (id) => document.getElementById(id);

    if (el('stat-group-a-count')) {
      el('stat-group-a-count').textContent = `${mainA.length} Hợp đồng (${appA.length} PL)`;
    }
    if (el('stat-group-a-val')) {
      el('stat-group-a-val').textContent = contractsManager.formatCurrency(totalSignedA);
    }
    if (el('stat-group-a-paid')) {
      el('stat-group-a-paid').textContent = contractsManager.formatCurrency(totalPaidA);
    }
    if (el('stat-group-a-rate')) {
      const rateA = totalSignedA > 0 ? ((totalPaidA / totalSignedA) * 100).toFixed(1) : 0;
      el('stat-group-a-rate').textContent = `Đạt ${rateA}% giá trị hợp đồng A`;
    }
    if (el('stat-group-a-ipc')) {
      el('stat-group-a-ipc').textContent = `${paymentsA.length} đợt`;
    }

    if (el('stat-group-b-count')) {
      el('stat-group-b-count').textContent = `${mainB.length} Hợp đồng (${appB.length} PL)`;
    }
    if (el('stat-group-b-val')) {
      el('stat-group-b-val').textContent = contractsManager.formatCurrency(totalSignedB);
    }
    if (el('stat-group-b-paid')) {
      el('stat-group-b-paid').textContent = contractsManager.formatCurrency(totalPaidB);
    }
    if (el('stat-group-b-rate')) {
      const rateB = totalSignedB > 0 ? ((totalPaidB / totalSignedB) * 100).toFixed(1) : 0;
      el('stat-group-b-rate').textContent = `Đạt ${rateB}% giá trị giao thầu B`;
    }
    if (el('stat-group-b-ipc')) {
      el('stat-group-b-ipc').textContent = `${paymentsB.length} đợt`;
    }

    // Cân đối dòng tiền
    const cashDiff = totalPaidA - totalPaidB;
    if (el('stat-cashflow-diff')) {
      const diffFormatted = contractsManager.formatCurrency(Math.abs(cashDiff));
      el('stat-cashflow-diff').textContent = (cashDiff >= 0 ? '+' : '-') + diffFormatted;
      el('stat-cashflow-diff').style.color = cashDiff >= 0 ? 'var(--success)' : 'var(--danger)';
    }

    if (el('stat-docs-count')) el('stat-docs-count').textContent = docs.length;
    if (el('stat-payments-count')) el('stat-payments-count').textContent = `${realPayments.length} đợt`;
  }

  bindNavigation() {
    const navPills = document.querySelectorAll('.nav-pill');
    navPills.forEach((pill) => {
      pill.addEventListener('click', () => {
        const targetView = pill.getAttribute('data-view');
        this.switchView(targetView);
      });
    });

    document.getElementById('brand-link')?.addEventListener('click', (e) => {
      e.preventDefault();
      this.switchView('dashboard');
    });

    document.getElementById('btn-quick-to-cpms')?.addEventListener('click', () => {
      this.switchView('cpms');
    });

    // Xuất file backup JSON
    document.getElementById('btn-export-backup')?.addEventListener('click', async () => {
      const snapshot = await db.exportSnapshot();
      const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `MAIO_Backup_CHC1_${new Date().toISOString().slice(0, 10)}.json`;
      a.click();
      URL.revokeObjectURL(url);
    });
  }

  switchView(viewName) {
    this.currentView = viewName;

    // Cập nhật active pill
    document.querySelectorAll('.nav-pill').forEach((pill) => {
      if (pill.getAttribute('data-view') === viewName) {
        pill.classList.add('active');
      } else {
        pill.classList.remove('active');
      }
    });

    // Ẩn tất cả views
    const allViews = ['dashboard', 'cpms', 'office', 'finance', 'ai', 'settings'];
    allViews.forEach((v) => {
      const el = document.getElementById(`view-${v}`);
      if (v === viewName) {
        el?.classList.remove('hidden');
      } else {
        el?.classList.add('hidden');
      }
    });

    // Nếu vào Dashboard thì cập nhật số liệu
    if (viewName === 'dashboard') {
      this.updateDashboardStats();
    }

    // Nếu vào CPMS thì render lại danh sách hợp đồng
    if (viewName === 'cpms') {
      contractsManager.loadData().then(() => contractsManager.render());
    }
  }

  setupLocalDirectoryPicker() {
    const pickBtn = document.getElementById('btn-pick-local-folder');
    const pathLabel = document.getElementById('local-folder-path');

    pickBtn?.addEventListener('click', async () => {
      if (!('showDirectoryPicker' in window)) {
        alert('Trình duyệt của bạn chưa hỗ trợ File System Access API. Vui lòng sử dụng Chrome hoặc Edge trên máy tính.');
        return;
      }

      try {
        this.localDirectoryHandle = await window.showDirectoryPicker();
        if (pathLabel && this.localDirectoryHandle) {
          pathLabel.textContent = `Thư mục đã chọn: ${this.localDirectoryHandle.name} (Sẵn sàng đồng bộ 2 chiều)`;
          pathLabel.style.color = 'var(--success)';
        }
        alert(`Đã liên kết thành công với thư mục máy tính: ${this.localDirectoryHandle.name}`);
      } catch (err) {
        if (err.name !== 'AbortError') {
          console.error('Lỗi chọn thư mục:', err);
        }
      }
    });
  }

  setupGoogleDrive() {
    const loginBtn = document.getElementById('btn-gdrive-login');
    const logoutBtn = document.getElementById('btn-gdrive-logout');
    const settingsLoginBtn = document.getElementById('btn-settings-login');
    const testSyncBtn = document.getElementById('btn-test-sync');
    const syncNowBtn = document.getElementById('btn-sync-now');

    const handleLogin = () => {
      try {
        gdrive.signIn();
      } catch (err) {
        alert(err.message || 'Chưa thể kết nối Google Drive. Vui lòng kiểm tra Client ID.');
      }
    };

    loginBtn?.addEventListener('click', handleLogin);
    settingsLoginBtn?.addEventListener('click', handleLogin);

    logoutBtn?.addEventListener('click', () => {
      gdrive.signOut();
    });

    // Lắng nghe thay đổi trạng thái đăng nhập Google
    gdrive.onAuthStateChanged((user) => {
      this.updateGoogleDriveUI(user);
    });

    // Khởi tạo dịch vụ Google Drive
    gdrive.init().catch((err) => {
      console.warn('Google Drive Service chưa sẵn sàng:', err);
    });

    // Xử lý nút đồng bộ ngay
    syncNowBtn?.addEventListener('click', async () => {
      try {
        syncNowBtn.disabled = true;
        syncNowBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đồng bộ...';
        await db.syncToGoogleDrive();
        alert('Đã đồng bộ thành công dữ liệu lên Google Drive!');
      } catch (err) {
        alert('Lỗi đồng bộ: ' + err.message);
      } finally {
        syncNowBtn.disabled = false;
        syncNowBtn.innerHTML = '<i class="fa-solid fa-rotate"></i> Đồng bộ ngay';
      }
    });

    testSyncBtn?.addEventListener('click', async () => {
      try {
        testSyncBtn.disabled = true;
        testSyncBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tải lên...';
        const res = await gdrive.saveJsonFile('MAIO_Test_Connection.json', {
          test: true,
          timestamp: new Date().toISOString(),
          message: 'Kết nối Google Drive thành công từ hệ thống MAIO!',
        });
        alert(`Tải file lên Google Drive thành công! File ID: ${res.id}`);
      } catch (err) {
        alert('Lỗi kiểm tra tải file: ' + err.message);
      } finally {
        testSyncBtn.disabled = false;
        testSyncBtn.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Kiểm tra Tải file lên Drive';
      }
    });
  }

  updateGoogleDriveUI(user) {
    const loginBtn = document.getElementById('btn-gdrive-login');
    const userProfile = document.getElementById('user-profile');
    const userAvatar = document.getElementById('user-avatar');
    const userName = document.getElementById('user-name');
    const cloudStatusText = document.getElementById('cloud-status-text');
    const cloudBadge = document.getElementById('cloud-badge');
    const syncNowBtn = document.getElementById('btn-sync-now');
    const testSyncBtn = document.getElementById('btn-test-sync');

    if (user) {
      loginBtn?.classList.add('hidden');
      userProfile?.classList.remove('hidden');
      if (userAvatar && user.picture) userAvatar.src = user.picture;
      if (userName) userName.textContent = user.name || user.email;

      if (cloudStatusText) {
        cloudStatusText.innerHTML = `Đã kết nối với Google Drive của <strong>${user.email}</strong>. Thư mục: <code>${
          window.MAIO_GDRIVE_CONFIG?.appFolderName || 'MAIO_Data'
        }</code>.`;
      }
      if (cloudBadge) {
        cloudBadge.className = 'badge badge-success';
        cloudBadge.textContent = 'Đã kết nối Cloud';
      }
      syncNowBtn?.classList.remove('hidden');
      if (testSyncBtn) testSyncBtn.disabled = false;
    } else {
      loginBtn?.classList.remove('hidden');
      userProfile?.classList.add('hidden');
      if (cloudStatusText) {
        cloudStatusText.textContent =
          'Dữ liệu đang được lưu trữ an toàn trong bộ nhớ cục bộ (IndexedDB). Sẵn sàng đồng bộ 2 chiều với Google Drive.';
      }
      if (cloudBadge) {
        cloudBadge.className = 'badge badge-warning';
        cloudBadge.textContent = 'Cục bộ (Offline)';
      }
      syncNowBtn?.classList.add('hidden');
      if (testSyncBtn) testSyncBtn.disabled = true;
    }
  }

  renderSettingsInfo() {
    const cfgClientId = document.getElementById('cfg-client-id');
    const config = window.MAIO_GDRIVE_CONFIG || {};

    if (cfgClientId) {
      cfgClientId.textContent = config.clientId || 'Chưa thiết lập trong gdrive-config.js';
    }
  }
}

// Khởi chạy khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
  const app = new App();
  app.init();
});
