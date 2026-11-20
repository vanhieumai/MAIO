/**
 * MAIO Application Entry Point
 * Quản lý khởi tạo IndexedDB, nạp dữ liệu hạt giống CHC1, điều hướng phân hệ và kết nối đồng bộ.
 */

import { gdrive } from './core/gdrive.js?v=5.0';
import { db } from './core/db.js?v=5.0';
import { contractsManager } from './modules/cpms/contracts.js?v=5.0';

class App {
  constructor() {
    this.currentView = 'cpms';
    this.currentCpmsTab = 'overview';
    this.localDirectoryHandle = null;
  }

  async init() {
    console.log('Khởi động MAIO Application...');

    // 1. Thiết lập điều hướng (Navigation)
    this.bindNavigation();

    // 2. Thiết lập dịch vụ Google Drive & Cổng Đăng nhập Auth Gate
    this.setupGoogleDrive();

    // 3. Thiết lập chọn thư mục máy tính (Local Folder)
    this.setupLocalDirectoryPicker();

    // 4. Cập nhật cấu hình lên màn hình Settings
    this.renderSettingsInfo();
  }

  async handleUserAuthChange(user) {
    const authGate = document.getElementById('auth-gate');
    const mainContent = document.querySelector('.main-content');
    const headerNav = document.querySelector('.header-nav');

    if (user && user.email) {
      console.log('Đã đăng nhập người dùng:', user.email, user.isOffline ? '(Offline)' : '(Online Google)');
      authGate?.classList.add('hidden');
      if (mainContent) mainContent.style.display = '';
      if (headerNav) headerNav.style.pointerEvents = '';

      // Chuyển đổi IndexedDB sang cơ sở dữ liệu riêng của tài khoản này
      await db.switchUser(user.email);

      // Nếu là tài khoản Google vừa đăng nhập từ chế độ offline, sao chép dữ liệu cục bộ sang tài khoản Google nếu DB mới rỗng
      if (!user.isOffline) {
        await db.copyLocalDataIfEmpty(user.email);
      }

      // Nạp dữ liệu và khởi tạo phân hệ Quản lý Dự án & Hợp đồng
      try {
        await contractsManager.init();
      } catch (e) {
        console.warn('Lỗi khởi tạo CPMS Manager:', e);
      }

      // Cập nhật số liệu Dashboard
      await this.updateDashboardStats();

      // Cập nhật giao diện Header & Cloud
      this.updateGoogleDriveUI(user);

      // Nếu có hành động đồng bộ đang chờ sau khi đăng nhập Google Drive thành công
      if (this.pendingSyncAction === 'to-drive') {
        this.pendingSyncAction = null;
        setTimeout(() => this.handleSyncToDrive(), 600);
      } else if (this.pendingSyncAction === 'from-drive') {
        this.pendingSyncAction = null;
        setTimeout(() => this.handleSyncFromDrive(), 600);
      }
    } else {
      console.log('Chưa đăng nhập Google Drive - Hiển thị Auth Gate');
      authGate?.classList.remove('hidden');
      if (mainContent) mainContent.style.display = 'none';
      if (headerNav) headerNav.style.pointerEvents = 'none';

      // Đóng database
      db.close();

      // Cập nhật giao diện Header & Cloud về trạng thái chưa đăng nhập
      this.updateGoogleDriveUI(null);
      this.resetDashboardStats();
    }
  }

  resetDashboardStats() {
    const el = (id) => document.getElementById(id);
    if (el('top-project-name')) el('top-project-name').textContent = 'Chưa có dự án nào';
    if (el('top-project-location')) el('top-project-location').textContent = '—';
    if (el('active-project-name')) el('active-project-name').textContent = 'Chưa có dự án nào';
    if (el('stat-group-a-count')) el('stat-group-a-count').textContent = '0 Hợp đồng';
    if (el('stat-group-a-val')) el('stat-group-a-val').textContent = '0 ₫';
    if (el('stat-group-a-paid')) el('stat-group-a-paid').textContent = '0 ₫';
    if (el('stat-group-a-rate')) el('stat-group-a-rate').textContent = 'Đạt 0% giá trị HĐ';
    if (el('stat-group-a-ipc')) el('stat-group-a-ipc').textContent = '0 đợt';
    if (el('stat-group-b-count')) el('stat-group-b-count').textContent = '0 Hợp đồng';
    if (el('stat-group-b-val')) el('stat-group-b-val').textContent = '0 ₫';
    if (el('stat-group-b-paid')) el('stat-group-b-paid').textContent = '0 ₫';
    if (el('stat-group-b-rate')) el('stat-group-b-rate').textContent = 'Đạt 0% giá trị giao thầu';
    if (el('stat-group-b-ipc')) el('stat-group-b-ipc').textContent = '0 đợt';
    if (el('stat-cash-in')) el('stat-cash-in').textContent = '0 ₫';
    if (el('stat-cash-out')) el('stat-cash-out').textContent = '0 ₫';
    if (el('stat-cash-balance')) el('stat-cash-balance').textContent = '0 ₫';
  }

  async updateDashboardStats() {
    if (!contractsManager) return;
    await contractsManager.loadData();

    const curPrjId = contractsManager.currentProjectId;
    const curPrj = contractsManager.projects.find((p) => p.id === curPrjId);

    // Cập nhật tên và địa điểm dự án hiện tại trên Banner và Dashboard
    const activePrjName = document.getElementById('active-project-name');
    const topProjName = document.getElementById('top-project-name');
    const topProjLoc = document.getElementById('top-project-location');

    if (curPrj) {
      if (activePrjName) activePrjName.textContent = curPrj.name || curPrj.id;
      if (topProjName) topProjName.textContent = curPrj.name || curPrj.id;
      if (topProjLoc) topProjLoc.textContent = curPrj.location ? String(curPrj.location).trim() : '—';
    } else {
      if (activePrjName) activePrjName.textContent = 'Chưa có dự án nào';
      if (topProjName) topProjName.textContent = 'Chưa có dự án nào';
      if (topProjLoc) topProjLoc.textContent = '—';
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
      this.switchView('cpms');
      this.switchCpmsTab('overview');
    });

    // Subnav tabs bên trong Quản lý Dự án (Tổng quan, Hợp đồng, Giải ngân)
    document.querySelectorAll('.cpms-subtab-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const targetTab = btn.getAttribute('data-cpms-tab');
        this.switchCpmsTab(targetTab);
      });
    });

    // Mobile Navigation Drawer (Mở / Đóng Menu trên màn hình dọc)
    const mobileDrawer = document.getElementById('mobile-menu-drawer');
    const openMenuBtn = document.getElementById('btn-mobile-menu');
    const closeMenuBtn = document.getElementById('btn-close-mobile-menu');

    openMenuBtn?.addEventListener('click', () => {
      mobileDrawer?.classList.remove('hidden');
    });

    closeMenuBtn?.addEventListener('click', () => {
      mobileDrawer?.classList.add('hidden');
    });

    mobileDrawer?.addEventListener('click', (e) => {
      if (e.target === mobileDrawer) {
        mobileDrawer.classList.add('hidden');
      }
    });

    // Click chuyển phân hệ từ Drawer Menu
    document.querySelectorAll('.mobile-nav-item').forEach((item) => {
      item.addEventListener('click', () => {
        const targetView = item.getAttribute('data-view');
        this.switchView(targetView);
        mobileDrawer?.classList.add('hidden');
      });
    });

    document.getElementById('btn-mobile-drawer-logout')?.addEventListener('click', () => {
      mobileDrawer?.classList.add('hidden');
      this.handleLogout();
    });

    // Sự kiện nút đăng nhập / thoát ở Header
    document.getElementById('btn-header-login')?.addEventListener('click', () => {
      try {
        gdrive.signIn();
      } catch (err) {
        alert(err.message || 'Chưa thể kết nối Google Drive.');
      }
    });

    document.getElementById('btn-header-logout')?.addEventListener('click', () => {
      this.handleLogout();
    });

    // Sự kiện trong trang Profile
    document.getElementById('btn-profile-login-gdrive')?.addEventListener('click', () => {
      try {
        gdrive.signIn();
      } catch (err) {
        alert(err.message || 'Chưa thể kết nối Google Drive.');
      }
    });

    document.getElementById('btn-profile-logout')?.addEventListener('click', () => {
      this.handleLogout();
    });

    document.getElementById('btn-profile-clear-cache')?.addEventListener('click', async () => {
      if (confirm('LƯU Ý QUAN TRỌNG:\n\nThao tác này sẽ xóa sạch dữ liệu IndexedDB lưu tạm trên trình duyệt của tài khoản này.\n\nNếu bạn đã sao lưu lên file .cpms hoặc Google Drive, dữ liệu vẫn an toàn.\n\nBạn có chắc chắn muốn xóa không?')) {
        await db.clear();
        alert('Đã xóa bộ nhớ tạm trình duyệt. Trang sẽ được tải lại.');
        window.location.reload();
      }
    });
  }

  handleLogout() {
    if (confirm('Bạn có chắc muốn thoát phiên làm việc hiện tại không?')) {
      gdrive.signOut();
    }
  }

  switchView(viewName) {
    this.currentView = viewName;

    // Cập nhật active pill trên Header & Drawer
    document.querySelectorAll('.nav-pill').forEach((pill) => {
      if (pill.getAttribute('data-view') === viewName) {
        pill.classList.add('active');
      } else {
        pill.classList.remove('active');
      }
    });

    document.querySelectorAll('.mobile-nav-item').forEach((item) => {
      if (item.getAttribute('data-view') === viewName) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    // Ẩn tất cả views
    const allViews = ['cpms', 'office', 'finance', 'profile', 'settings'];
    allViews.forEach((v) => {
      const el = document.getElementById(`view-${v}`);
      if (v === viewName) {
        el?.classList.remove('hidden');
      } else {
        el?.classList.add('hidden');
      }
    });

    // Thanh 3 Sub-tab chỉ hiển thị khi ở phân hệ Quản lý dự án
    const cpmsSubnavBar = document.getElementById('cpms-subnav-bar');
    if (viewName === 'cpms') {
      cpmsSubnavBar?.classList.remove('hidden');
      this.switchCpmsTab(this.currentCpmsTab || 'overview');
    } else {
      cpmsSubnavBar?.classList.add('hidden');
      if (viewName === 'profile') {
        this.updateProfileView();
      }
    }
  }

  switchCpmsTab(tabName) {
    this.currentCpmsTab = tabName || 'overview';

    // Cập nhật active subtab button
    document.querySelectorAll('.cpms-subtab-btn').forEach((btn) => {
      if (btn.getAttribute('data-cpms-tab') === this.currentCpmsTab) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    // Ẩn/hiện các tab pane trong CPMS
    const panes = ['overview', 'contracts', 'disbursements'];
    panes.forEach((p) => {
      const paneEl = document.getElementById(`cpms-tab-${p}`);
      if (p === this.currentCpmsTab) {
        paneEl?.classList.remove('hidden');
      } else {
        paneEl?.classList.add('hidden');
      }
    });

    // Thực hiện render dữ liệu cho tab tương ứng
    if (this.currentCpmsTab === 'overview') {
      this.updateDashboardStats();
    } else if (this.currentCpmsTab === 'contracts') {
      contractsManager.loadData().then(() => contractsManager.render());
    } else if (this.currentCpmsTab === 'disbursements') {
      contractsManager.loadData().then(() => contractsManager.renderDisbursementsTab());
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
    const gateLoginBtn = document.getElementById('btn-login-gate');
    const gateOfflineBtn = document.getElementById('btn-login-offline');
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
    gateLoginBtn?.addEventListener('click', handleLogin);

    gateOfflineBtn?.addEventListener('click', () => {
      // 1. Vào thẳng chế độ Cục bộ (Offline) - Không hỏi bất kỳ email demo nào
      gdrive.signInOffline();

      // 2. Chuyển sang phân hệ Quản lý Dự án (CPMS)
      this.switchView('cpms');

      // 3. Tự động mở File System Picker để dẫn file Data CPMS local
      setTimeout(() => {
        contractsManager.pickAndLoadLocalCPMSFile();
      }, 150);
    });

    logoutBtn?.addEventListener('click', () => {
      if (confirm('Bạn có chắc muốn thoát phiên làm việc này không?')) {
        gdrive.signOut();
      }
    });

    // Lắng nghe thay đổi trạng thái đăng nhập Google (mỗi tài khoản gắn với DB riêng)
    gdrive.onAuthStateChanged(async (user) => {
      await this.handleUserAuthChange(user);
    });

    // Khởi tạo dịch vụ Google Drive
    gdrive.init().catch((err) => {
      console.warn('Google Drive Service chưa sẵn sàng:', err);
    });

    // 1. Nút Đồng bộ lên Drive (Local -> Drive)
    const syncToDriveBtn = document.getElementById('btn-sync-to-drive');
    syncToDriveBtn?.addEventListener('click', async () => {
      if (gdrive.currentUser?.isOffline || !gdrive.isSignedIn()) {
        const wantLogin = confirm(
          'Bạn đang ở chế độ làm việc Cục bộ (Offline).\n\nĐể sao lưu dữ liệu lên Google Drive, bạn cần kết nối tài khoản Google Drive của mình.\n\nBạn có muốn đăng nhập Google Drive ngay bây giờ không?'
        );
        if (wantLogin) {
          this.pendingSyncAction = 'to-drive';
          handleLogin();
        }
        return;
      }
      await this.handleSyncToDrive();
    });

    // 2. Nút Đồng bộ về máy (Drive -> Local)
    const syncFromDriveBtn = document.getElementById('btn-sync-from-drive');
    syncFromDriveBtn?.addEventListener('click', async () => {
      if (gdrive.currentUser?.isOffline || !gdrive.isSignedIn()) {
        const wantLogin = confirm(
          'Bạn đang ở chế độ làm việc Cục bộ (Offline).\n\nĐể đồng bộ dữ liệu từ Google Drive về máy tính, bạn cần kết nối tài khoản Google Drive của mình.\n\nBạn có muốn đăng nhập Google Drive ngay bây giờ không?'
        );
        if (wantLogin) {
          this.pendingSyncAction = 'from-drive';
          handleLogin();
        }
        return;
      }
      await this.handleSyncFromDrive();
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

  /**
   * Đồng bộ dữ liệu hiện tại từ máy tính lên Google Drive
   */
  async handleSyncToDrive() {
    const btn = document.getElementById('btn-sync-to-drive');
    const origHtml = btn ? btn.innerHTML : '';
    try {
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đồng bộ...';
      }

      const curPrjId = contractsManager.currentProjectId || 'CHC1';
      await db.syncToGoogleDrive();

      // Nếu đang mở file CPMS cục bộ, tự động ghi lưu và sao lưu bản sao lên Drive
      if (contractsManager.activeFileHandle) {
        try {
          const file = await contractsManager.activeFileHandle.getFile();
          const arrayBuffer = await file.arrayBuffer();
          const cpmsFolderId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${curPrjId}`, 'CPMS_Data']);
          const fileName = file.name || `MAIO_Data_${curPrjId}.cpms`;
          
          // Ghi đè file CPMS lên Drive
          const existingFiles = await gdrive.listProjectFiles(cpmsFolderId);
          const exist = existingFiles.find((f) => f.name === fileName);

          const metadata = {
            name: fileName,
            mimeType: 'application/octet-stream',
            ...(!exist && { parents: [cpmsFolderId] }),
          };

          const boundary = '-------314159265358979323846';
          const delimiter = `\r\n--${boundary}\r\n`;
          const closeDelimiter = `\r\n--${boundary}--`;
          const metaBlob = new Blob([delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) + delimiter + 'Content-Type: application/octet-stream\r\n\r\n']);
          const endBlob = new Blob([closeDelimiter]);
          const multipartBlob = new Blob([metaBlob, arrayBuffer, endBlob], { type: `multipart/related; boundary=${boundary}` });

          let url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
          let method = 'POST';
          if (exist) {
            url = `https://www.googleapis.com/upload/drive/v3/files/${exist.id}?uploadType=multipart`;
            method = 'PATCH';
          }

          await fetch(url, {
            method,
            headers: { Authorization: `Bearer ${gdrive.accessToken}` },
            body: multipartBlob,
          });
        } catch (e) {
          console.warn('Lỗi tải kèm file nhị phân CPMS lên Drive:', e);
        }
      }

      alert(`Đã đồng bộ thành công toàn bộ dữ liệu dự án "${curPrjId}" lên Google Drive!`);
    } catch (err) {
      alert('Lỗi đồng bộ lên Google Drive: ' + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
    }
  }

  /**
   * Đồng bộ dữ liệu từ Google Drive về máy tính & Cập nhật lại đường dẫn hồ sơ local
   */
  async handleSyncFromDrive() {
    const btn = document.getElementById('btn-sync-from-drive');
    const origHtml = btn ? btn.innerHTML : '';
    try {
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tải về...';
      }

      // 1. Quét tìm các snapshot trên Drive
      const snapshots = await gdrive.listProjectSnapshots();
      if (!snapshots || snapshots.length === 0) {
        alert('Chưa tìm thấy bản sao lưu dự án (MAIO_CPMS_Snapshot_*.json) nào trên Google Drive của bạn.');
        return;
      }

      // 2. Xác định snapshot cần nạp
      let targetFile = snapshots[0];
      const curPrjId = contractsManager.currentProjectId;
      if (curPrjId) {
        const matched = snapshots.find((s) => s.name.includes(`_${curPrjId}.json`));
        if (matched) targetFile = matched;
      }

      if (snapshots.length > 1) {
        const optionsList = snapshots.map((s, idx) => `${idx + 1}. ${s.name}`).join('\n');
        const chosen = prompt(`Tìm thấy ${snapshots.length} bản sao lưu trên Drive:\n\n${optionsList}\n\nNhập số thứ tự file bạn muốn đồng bộ về máy:`, '1');
        if (!chosen) return;
        const idx = parseInt(chosen, 10) - 1;
        if (snapshots[idx]) targetFile = snapshots[idx];
      }

      // 3. Cho phép chọn thư mục máy tính để lưu hồ sơ tài liệu
      let selectedDirHandle = this.localDirectoryHandle;
      let selectedDirName = selectedDirHandle ? selectedDirHandle.name : '';

      if ('showDirectoryPicker' in window) {
        const confirmPickDir = confirm(
          `Bạn có muốn chọn thư mục trên máy tính để đồng bộ cây thư mục hồ sơ tài liệu về và cập nhật đường dẫn local không?\n\n- Bấm OK để chọn thư mục lưu trữ trên máy tính\n- Bấm Cancel để giữ nguyên dữ liệu local hiện tại`
        );
        if (confirmPickDir) {
          try {
            selectedDirHandle = await window.showDirectoryPicker();
            this.localDirectoryHandle = selectedDirHandle;
            selectedDirName = selectedDirHandle.name;
          } catch (e) {
            if (e.name !== 'AbortError') console.warn('Lỗi chọn thư mục:', e);
          }
        }
      }

      // 4. Đọc nội dung snapshot từ Drive và import vào DB
      const snapshotData = await gdrive.readJsonFile(targetFile.id);
      await db.importSnapshot(snapshotData);

      // 5. Cập nhật lại đường dẫn local cho các hồ sơ tài liệu theo thư mục đã chọn
      const prjId = snapshotData.projects?.[0]?.id || curPrjId || 'CHC1';
      let updatedDocCount = 0;
      if (selectedDirName) {
        updatedDocCount = await contractsManager.updateDocumentLocalPaths(selectedDirName, prjId);
      }

      // 6. Tải cây hồ sơ trên Drive về local nếu có thư mục local đã chọn
      if (selectedDirHandle) {
        try {
          const prjFolderId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn', `DuAn_${prjId}`]);
          if (prjFolderId) {
            const prjSubHandle = await selectedDirHandle.getDirectoryHandle(`DuAn_${prjId}`, { create: true });
            await gdrive.downloadDriveFolderToLocal(prjFolderId, prjSubHandle);
          }
        } catch (e) {
          console.warn('Lỗi tải cây thư mục từ Drive về máy:', e);
        }
      }

      // 7. Reload lại dữ liệu CPMS & Dashboard
      await contractsManager.loadData();
      contractsManager.render();
      await this.updateDashboardStats();

      const dirMsg = selectedDirName
        ? `\n- Thư mục lưu trữ: ${selectedDirName}\n- Đã cập nhật lại đường dẫn ${updatedDocCount} hồ sơ tài liệu theo thư mục mới.`
        : '';
      alert(`Đã đồng bộ thành công dữ liệu dự án từ Google Drive về máy!${dirMsg}`);
    } catch (err) {
      alert('Lỗi đồng bộ từ Google Drive về máy: ' + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
    }
  }

  updateGoogleDriveUI(user) {
    const headerLoginBtn = document.getElementById('btn-header-login');
    const headerUserStatus = document.getElementById('header-user-status');
    const headerUserAvatarIcon = document.getElementById('header-user-avatar-icon');
    const headerUserName = document.getElementById('header-user-name');
    const testSyncBtn = document.getElementById('btn-test-sync');

    const mobileDrawerAuthIcon = document.getElementById('mobile-drawer-auth-icon');
    const mobileDrawerAuthText = document.getElementById('mobile-drawer-auth-text');

    if (user) {
      // Khi đã vào phiên làm việc: ẨN nút Đăng nhập, CHỈ HIỆN nhãn (Local hoặc email) + nút Thoát
      headerLoginBtn?.classList.add('hidden');
      headerUserStatus?.classList.remove('hidden');

      if (user.isOffline) {
        // Chế độ Cục bộ: Hiển thị nhãn Local
        if (headerUserAvatarIcon) headerUserAvatarIcon.className = 'fa-solid fa-laptop-file';
        if (headerUserName) headerUserName.textContent = 'Local';
        if (mobileDrawerAuthIcon) mobileDrawerAuthIcon.className = 'fa-solid fa-laptop-file';
        if (mobileDrawerAuthText) mobileDrawerAuthText.textContent = 'Local';
        if (testSyncBtn) testSyncBtn.disabled = true;
      } else {
        // Đã kết nối Google Drive: Hiển thị email
        if (headerUserAvatarIcon) headerUserAvatarIcon.className = 'fa-brands fa-google';
        if (headerUserName) {
          headerUserName.textContent = user.email || user.name || 'Drive';
        }
        if (mobileDrawerAuthIcon) mobileDrawerAuthIcon.className = 'fa-brands fa-google';
        if (mobileDrawerAuthText) {
          mobileDrawerAuthText.textContent = user.email || user.name || 'Drive';
        }
        if (testSyncBtn) testSyncBtn.disabled = false;
      }
    } else {
      // Khi chưa đăng nhập (Auth Gate): Chỉ hiện nút Đăng nhập, ẩn nút Thoát
      headerLoginBtn?.classList.remove('hidden');
      headerUserStatus?.classList.add('hidden');
      if (testSyncBtn) testSyncBtn.disabled = true;
    }

    this.updateProfileView(user);
  }

  async updateProfileView(user = gdrive.currentUser) {
    const pName = document.getElementById('profile-user-name');
    const pEmail = document.getElementById('profile-user-email');
    const pBadge = document.getElementById('profile-status-badge');
    const pImg = document.getElementById('profile-avatar-img');
    const pIcon = document.getElementById('profile-avatar-icon');
    const pLoginBtn = document.getElementById('btn-profile-login-gdrive');
    const pLoginText = document.getElementById('btn-profile-login-text');
    const statProjects = document.getElementById('profile-stat-projects');
    const statStorage = document.getElementById('profile-stat-storage');

    if (user) {
      if (user.isOffline) {
        if (pName) pName.textContent = 'Chế độ Cục bộ (Offline)';
        if (pEmail) pEmail.textContent = 'Lưu trữ trực tiếp trên máy tính qua file CPMS & IndexedDB';
        if (pBadge) {
          pBadge.className = 'badge badge-warning';
          pBadge.textContent = 'Cục bộ (Offline)';
        }
        if (pImg) pImg.classList.add('hidden');
        if (pIcon) pIcon.classList.remove('hidden');
        if (pLoginBtn) pLoginBtn.classList.remove('hidden');
        if (pLoginText) pLoginText.textContent = 'Đăng nhập Google Drive để Đồng bộ Online';
      } else {
        if (pName) pName.textContent = user.name || 'Người dùng Google';
        if (pEmail) pEmail.textContent = user.email || '—';
        if (pBadge) {
          pBadge.className = 'badge badge-success';
          pBadge.textContent = 'Đã kết nối Google Drive';
        }
        if (pImg && user.picture) {
          pImg.src = user.picture;
          pImg.classList.remove('hidden');
        }
        if (pIcon) pIcon.classList.add('hidden');
        if (pLoginBtn) pLoginBtn.classList.remove('hidden');
        if (pLoginText) pLoginText.textContent = 'Đổi / Đăng nhập tài khoản Google khác';
      }
    } else {
      if (pName) pName.textContent = 'Chưa đăng nhập';
      if (pEmail) pEmail.textContent = '—';
      if (pBadge) {
        pBadge.className = 'badge badge-warning';
        pBadge.textContent = 'Chưa đăng nhập';
      }
    }

    if (statProjects && contractsManager) {
      statProjects.textContent = `${contractsManager.projects?.length || 0} dự án`;
    }

    if (statStorage) {
      if (navigator.storage && navigator.storage.estimate) {
        try {
          const estimate = await navigator.storage.estimate();
          const usedMB = (estimate.usage / (1024 * 1024)).toFixed(2);
          statStorage.textContent = `~${usedMB} MB`;
        } catch {
          statStorage.textContent = 'Đang hoạt động';
        }
      } else {
        statStorage.textContent = 'Đang hoạt động';
      }
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
