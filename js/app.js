/**
 * MAIO Application Entry Point
 * Quản lý khởi tạo IndexedDB, nạp dữ liệu hạt giống CHC1, điều hướng phân hệ và kết nối đồng bộ.
 */

import { gdrive } from './core/gdrive.js?v=5.2';
import { db } from './core/db.js?v=5.1';
import { contractsManager } from './modules/cpms/contracts.js?v=7.9';

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

      // Lưu lại dự án & file handle đang mở ở phiên trước để không bị mất khi chuyển sang Google Drive
      const prevProjectId = contractsManager?.currentProjectId || null;
      const prevFileHandle = contractsManager?.activeFileHandle || null;

      // Chuyển đổi IndexedDB sang cơ sở dữ liệu riêng của tài khoản này
      await db.switchUser(user.email);

      // Nếu là tài khoản Google vừa đăng nhập từ chế độ offline, hợp nhất dữ liệu cục bộ (đặc biệt là dự án đang mở) sang tài khoản Google
      if (!user.isOffline) {
        await db.copyLocalDataIfEmpty(user.email, prevProjectId);
      }

      // Nạp dữ liệu và khởi tạo phân hệ Quản lý Dự án & Hợp đồng
      try {
        await contractsManager.init();
        if (prevFileHandle) {
          contractsManager.activeFileHandle = prevFileHandle;
        }
        if (prevProjectId && (contractsManager.projects || []).some((p) => p.id === prevProjectId)) {
          contractsManager.currentProjectId = prevProjectId;
          await contractsManager.loadData();
          contractsManager.render();
        }
      } catch (e) {
        console.warn('Lỗi khởi tạo CPMS Manager:', e);
      }

      // Cập nhật số liệu Dashboard
      await this.updateDashboardStats();

      // Cập nhật giao diện Header & Cloud & Badge nguồn dữ liệu
      this.updateGoogleDriveUI(user);
      this.updateProjectSourceBadge();

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

    // Kiểm tra tính sẵn sàng của dự án và cập nhật Empty State UI (giữ nguyên tab hiện tại)
    contractsManager.updateEmptyStateUI(this.currentCpmsTab || 'overview');
    if (!contractsManager.isProjectReady()) {
      return;
    }

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

    // 1. Nhóm A (Chủ đầu tư / Tổng thầu - loại trừ hợp đồng Cancel)
    const mainA = contracts.filter((c) => !c.parent_id && c.group === 'A' && !contractsManager.isContractCanceled(c));
    const appA = contracts.filter((c) => c.parent_id && c.group === 'A' && !contractsManager.isContractOrAncestorCanceled(c));
    let totalSignedA = 0;
    let totalPaidA = 0;
    mainA.forEach((c) => {
      const st = contractsManager.calculateTotals(c);
      totalSignedA += st.totalSigned;
      totalPaidA += st.totalPaid;
    });

    const familyAIds = new Set();
    contracts
      .filter((c) => c.group === 'A' && !contractsManager.isContractOrAncestorCanceled(c))
      .forEach((c) => familyAIds.add(c.id));
    const paymentsA = realPayments.filter((p) => familyAIds.has(p.contract_id));

    // 2. Nhóm B (Nhà thầu phụ / Nhà cung cấp - loại trừ hợp đồng Cancel)
    const mainB = contracts.filter((c) => !c.parent_id && c.group === 'B' && !contractsManager.isContractCanceled(c));
    const appB = contracts.filter((c) => c.parent_id && c.group === 'B' && !contractsManager.isContractOrAncestorCanceled(c));
    let totalSignedB = 0;
    let totalPaidB = 0;
    mainB.forEach((c) => {
      const st = contractsManager.calculateTotals(c);
      totalSignedB += st.totalSigned;
      totalPaidB += st.totalPaid;
    });

    const familyBIds = new Set();
    contracts
      .filter((c) => c.group === 'B' && !contractsManager.isContractOrAncestorCanceled(c))
      .forEach((c) => familyBIds.add(c.id));
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

    // Cập nhật số đếm badge nhắc việc
    contractsManager.updateReminderBadge?.();
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

    // Left Side Menu Drawer (Ẩn / Hiện Menu phân hệ & tài khoản bên trái)
    const mobileDrawer = document.getElementById('mobile-menu-drawer');
    const openMenuBtn = document.getElementById('btn-mobile-menu');
    const closeMenuBtn = document.getElementById('btn-close-mobile-menu');

    const toggleSideMenu = () => {
      mobileDrawer?.classList.toggle('hidden');
    };

    openMenuBtn?.addEventListener('click', toggleSideMenu);
    document.querySelectorAll('.btn-side-menu-toggle').forEach((btn) => {
      btn.addEventListener('click', toggleSideMenu);
    });

    closeMenuBtn?.addEventListener('click', () => {
      mobileDrawer?.classList.add('hidden');
    });

    mobileDrawer?.addEventListener('click', (e) => {
      if (e.target === mobileDrawer) {
        mobileDrawer.classList.add('hidden');
      }
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && mobileDrawer && !mobileDrawer.classList.contains('hidden')) {
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
      const ok = await (window.maioConfirm
        ? window.maioConfirm({
            title: 'Xóa bộ nhớ tạm trình duyệt',
            message: 'Thao tác này sẽ xóa sạch dữ liệu IndexedDB lưu tạm trên trình duyệt của tài khoản này.',
            subMessage: 'Nếu bạn đã sao lưu lên file .cpms hoặc Google Drive, dữ liệu vẫn an toàn.',
            confirmText: 'Xóa bộ nhớ tạm',
            cancelText: 'Hủy bỏ',
            confirmType: 'danger',
          })
        : Promise.resolve(confirm('Bạn có chắc chắn muốn xóa bộ nhớ tạm trình duyệt không?')));
      if (ok) {
        await db.clear();
        alert('Đã xóa bộ nhớ tạm trình duyệt. Trang sẽ được tải lại.');
        window.location.reload();
      }
    });
  }

  async handleLogout() {
    const ok = await (window.maioConfirm
      ? window.maioConfirm({
          title: 'Thoát phiên làm việc',
          message: 'Bạn có chắc muốn thoát phiên làm việc hiện tại không?',
          confirmText: 'Đăng xuất',
          cancelText: 'Ở lại',
          confirmType: 'primary',
        })
      : Promise.resolve(confirm('Bạn có chắc muốn thoát phiên làm việc hiện tại không?')));
    if (ok) {
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

    // Phân hệ Quản lý dự án (CPMS) quản lý riêng các sub-tab và banner dự án
    if (viewName === 'cpms') {
      this.switchCpmsTab(this.currentCpmsTab || 'overview');
    } else if (viewName === 'profile') {
      this.updateProfileView();
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

    // Kiểm tra trạng thái dự án: Nếu chưa có dự án, ẩn toàn bộ 3 tab và hiển thị Empty State
    const isReady = contractsManager?.isProjectReady();
    contractsManager?.updateEmptyStateUI(this.currentCpmsTab);

    if (!isReady) {
      return;
    }

    // Ẩn/hiện các tab pane trong CPMS khi dự án đã sẵn sàng
    const panes = ['overview', 'contracts', 'disbursements', 'reminder'];
    panes.forEach((p) => {
      const paneEl = document.getElementById(`cpms-tab-${p}`);
      if (p === this.currentCpmsTab) {
        paneEl?.classList.remove('hidden');
      } else {
        paneEl?.classList.add('hidden');
      }
    });

    // Cập nhật badge số lượng nhắc việc
    contractsManager?.updateReminderBadge?.();

    // Thực hiện render dữ liệu cho tab tương ứng
    if (this.currentCpmsTab === 'overview') {
      this.updateDashboardStats();
    } else if (this.currentCpmsTab === 'contracts') {
      contractsManager.loadData().then(() => contractsManager.render());
    } else if (this.currentCpmsTab === 'disbursements') {
      contractsManager.loadData().then(() => contractsManager.renderDisbursementsTab());
    } else if (this.currentCpmsTab === 'reminder') {
      contractsManager.loadData().then(() => contractsManager.renderRemindersTab());
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

    const handleLogin = async () => {
      try {
        await gdrive.signIn();
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

    logoutBtn?.addEventListener('click', async () => {
      const ok = await (window.maioConfirm
        ? window.maioConfirm({
            title: 'Thoát phiên làm việc',
            message: 'Bạn có chắc muốn thoát phiên làm việc này không?',
            confirmText: 'Đăng xuất',
            cancelText: 'Ở lại',
            confirmType: 'primary',
          })
        : Promise.resolve(confirm('Bạn có chắc muốn thoát phiên làm việc này không?')));
      if (ok) {
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
        const wantLogin = await (window.maioConfirm
          ? window.maioConfirm({
              title: 'Kết nối Google Drive',
              message: 'Bạn đang ở chế độ làm việc Cục bộ (Offline). Để sao lưu dữ liệu lên Google Drive, bạn cần đăng nhập tài khoản Google Drive của mình.',
              confirmText: 'Đăng nhập Google',
              cancelText: 'Để sau',
              confirmType: 'primary',
            })
          : Promise.resolve(confirm('Bạn có muốn đăng nhập Google Drive ngay bây giờ không?')));
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
        const wantLogin = await (window.maioConfirm
          ? window.maioConfirm({
              title: 'Kết nối Google Drive',
              message: 'Bạn đang ở chế độ làm việc Cục bộ (Offline). Để đồng bộ dữ liệu từ Google Drive về máy tính, bạn cần đăng nhập tài khoản Google Drive của mình.',
              confirmText: 'Đăng nhập Google',
              cancelText: 'Để sau',
              confirmType: 'primary',
            })
          : Promise.resolve(confirm('Bạn có muốn đăng nhập Google Drive ngay bây giờ không?')));
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
   * Lập chỉ mục các file trong thư mục Local (FileSystemDirectoryHandle) để đối chiếu nhanh theo tên/đường dẫn
   */
  async indexLocalFiles(dirHandle, currentRelPath = '', fileIndex = null) {
    const index = fileIndex || { byRelPath: new Map(), byFileName: new Map() };
    try {
      for await (const [name, handle] of dirHandle.entries()) {
        const relPath = currentRelPath ? `${currentRelPath}/${name}` : name;
        if (handle.kind === 'file') {
          const normRel = relPath.replace(/\\/g, '/').toLowerCase();
          index.byRelPath.set(normRel, handle);
          const lowerName = name.toLowerCase();
          if (!index.byFileName.has(lowerName)) {
            index.byFileName.set(lowerName, []);
          }
          index.byFileName.get(lowerName).push({ relPath: normRel, handle });
        } else if (handle.kind === 'directory') {
          await this.indexLocalFiles(handle, relPath, index);
        }
      }
    } catch (e) {
      console.warn('Lỗi duyệt thư mục local:', e);
    }
    return index;
  }

  /**
   * Tìm FileHandle khớp với đường dẫn local_link trong chỉ mục đã quét
   */
  matchLocalFileHandle(localLink, index) {
    if (!localLink || !index) return null;
    const normLink = String(localLink).replace(/\\/g, '/').trim().toLowerCase();
    const parts = normLink.split('/').filter(Boolean);
    const fileName = parts[parts.length - 1];
    if (!fileName) return null;

    // 1. Ưu tiên khớp phần đuôi đường dẫn tương đối dài nhất
    const candidates = index.byFileName.get(fileName) || [];
    if (candidates.length === 1) return candidates[0].handle;
    if (candidates.length > 1) {
      let bestMatch = candidates[0].handle;
      let bestScore = 0;
      for (const cand of candidates) {
        const candParts = cand.relPath.split('/');
        let score = 0;
        let i = parts.length - 1;
        let j = candParts.length - 1;
        while (i >= 0 && j >= 0 && parts[i] === candParts[j]) {
          score++;
          i--;
          j--;
        }
        if (score > bestScore) {
          bestScore = score;
          bestMatch = cand.handle;
        }
      }
      return bestMatch;
    }
    return null;
  }

  /**
   * Cập nhật badge hiển thị rõ nguồn dữ liệu dự án đang mở (Local vs Online) ngay bên cạnh tên dự án
   */
  updateProjectSourceBadge() {
    const badge = document.getElementById('top-project-source-badge');
    if (!badge) return;

    const curPrj = (contractsManager?.projects || []).find((p) => p.id === contractsManager?.currentProjectId);
    if (!curPrj) {
      badge.style.display = 'none';
      return;
    }
    badge.style.display = 'inline-flex';

    const mode = contractsManager?.getProjectOriginMode ? contractsManager.getProjectOriginMode(curPrj) : 'local';
    const localFileName = contractsManager?.activeFileHandle?.name || '';
    const isDriveOnline = Boolean(gdrive.isSignedIn() && !gdrive.currentUser?.isOffline);

    badge.classList.remove('source-local', 'source-drive', 'source-hybrid', 'source-browser');

    if (mode === 'online') {
      badge.classList.add('source-drive');
      badge.innerHTML = `<i class="fa-brands fa-google-drive"></i> Online`;
      badge.title = isDriveOnline
        ? `Đang làm việc với dữ liệu Online trên Google Drive (${gdrive.currentUser?.email || ''})`
        : 'Dự án đồng bộ từ Google Drive Online';
    } else {
      badge.classList.add('source-local');
      badge.innerHTML = localFileName
        ? `<i class="fa-solid fa-laptop-file"></i> Local (${localFileName})`
        : `<i class="fa-solid fa-laptop-file"></i> Local`;
      badge.title = localFileName
        ? `Đang làm việc với file CPMS Local trên máy tính: ${localFileName}`
        : 'Đang làm việc với dữ liệu Local trên máy tính';
    }
  }

  /**
   * Lập chỉ mục các file từ danh sách FileList (<input webkitdirectory>)
   */
  indexFileList(fileList) {
    const index = { byRelPath: new Map(), byFileName: new Map() };
    Array.from(fileList || []).forEach((file) => {
      const rel = String(file.webkitRelativePath || file.name).replace(/\\/g, '/').toLowerCase();
      index.byRelPath.set(rel, file);
      const lowerName = file.name.toLowerCase();
      if (!index.byFileName.has(lowerName)) {
        index.byFileName.set(lowerName, []);
      }
      index.byFileName.get(lowerName).push({ relPath: rel, file });
    });
    return index;
  }

  /**
   * Tìm FileHandle hoặc File object khớp với đường dẫn local_link
   */
  matchLocalFileEntry(localLink, index) {
    if (!localLink || !index) return null;
    const normLink = String(localLink).replace(/\\/g, '/').trim().toLowerCase();
    const parts = normLink.split('/').filter(Boolean);
    const fileName = parts[parts.length - 1];
    if (!fileName) return null;

    const candidates = index.byFileName.get(fileName) || [];
    if (candidates.length === 1) return candidates[0].handle || candidates[0].file || null;
    if (candidates.length > 1) {
      let bestItem = candidates[0];
      let bestScore = -1;
      for (const cand of candidates) {
        const candParts = cand.relPath.split('/');
        let score = 0;
        let i = parts.length - 1;
        let j = candParts.length - 1;
        while (i >= 0 && j >= 0 && parts[i] === candParts[j]) {
          score++;
          i--;
          j--;
        }
        if (score > bestScore) {
          bestScore = score;
          bestItem = cand;
        }
      }
      return bestItem.handle || bestItem.file || null;
    }
    return null;
  }

  /**
   * Tìm đường dẫn gốc chung từ danh sách local_link trong data CPMS để gợi ý cho người dùng
   */
  detectCommonLocalRoot(docs) {
    const paths = (docs || [])
      .map((d) => String(d.local_link || '').trim())
      .filter((p) => p && !p.startsWith('http'));
    if (paths.length === 0) return '';
    const splitPaths = paths.map((p) => p.replace(/\\/g, '/').split('/').slice(0, -1));
    let common = splitPaths[0] || [];
    for (let i = 1; i < splitPaths.length; i++) {
      const cur = splitPaths[i];
      let j = 0;
      while (j < common.length && j < cur.length && common[j].toLowerCase() === cur[j].toLowerCase()) {
        j++;
      }
      common = common.slice(0, j);
      if (common.length === 0) break;
    }
    return common.join('\\');
  }

  /**
   * Mở Modal Đồng bộ lên Google Drive:
   * - Tự động dùng lại quyền thư mục gốc đã lưu theo dự án (chỉ cấp quyền 1 lần duy nhất)
   * - Kiểm tra trùng thư mục theo cả mã & tên dự án để hỏi ghi đè
   * - Upload file từ Link Local vào HoSo / HopDong_PL / ThanhToan & cập nhật drive_link + file .cpms
   */
  async handleSyncToDrive() {
    const curPrjId = contractsManager.currentProjectId || db.activeProjectId || 'CHC1';
    db.activeProjectId = curPrjId;
    const curPrj = (contractsManager.projects || []).find((p) => p.id === curPrjId);
    if (!curPrj) {
      alert('Chưa có dự án nào đang mở để đồng bộ lên Google Drive.');
      return;
    }

    const safePrjId = String(curPrjId).replace(/[\/\\?%*:|"<>]/g, '_');
    const defaultFolderName = `DuAn_${safePrjId}`;
    const projectDocs = (contractsManager.documents || []).filter(
      (d) => (!d.project_id || d.project_id === curPrjId) && d.local_link && !String(d.local_link).startsWith('http')
    );
    const commonRootHint = this.detectCommonLocalRoot(projectDocs);

    // Khôi phục dirHandle đã lưu của dự án nếu trong phiên hiện tại chưa có
    if (!this.localDirectoryHandle && projectDocs.length > 0) {
      const savedHandle = await db.getProjectDirHandle(curPrjId);
      if (savedHandle) {
        this.localDirectoryHandle = savedHandle;
      }
    }

    document.getElementById('drive-sync-upload-modal')?.remove();
    const modal = document.createElement('div');
    modal.id = 'drive-sync-upload-modal';
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 600px;">
        <div class="modal-header">
          <h2 style="font-size: 1.08rem; font-weight: 700; display: flex; align-items: center; gap: 8px;">
            <i class="fa-brands fa-google-drive" style="color: #2563eb;"></i>
            Đồng bộ Dự án lên Google Drive
          </h2>
          <button type="button" id="btn-close-drive-sync-modal" class="button button-outline" style="padding: 4px 10px;">
            <i class="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div class="modal-body" style="padding: 16px 20px; font-size: 0.86rem;">
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 14px; margin-bottom: 14px;">
            <div><strong>Dự án:</strong> ${curPrj.name || curPrjId} (<code>${curPrjId}</code>)</div>
            <div style="color: #64748b; font-size: 0.78rem; margin-top: 3px;">
              Cấu trúc Drive: <code>MAIO_Data / 01_QuanLy_DuAn / &lt;Thư mục dự án&gt; / { CPMS_Data, HoSo, HopDong_PL, ThanhToan }</code>
            </div>
          </div>

      <!-- BƯỚC 1: KIỂM TRA TRÙNG THƯ MỤC TRÊN DRIVE -->
          <div style="margin-bottom: 14px;">
            <div style="font-weight: 600; margin-bottom: 6px; color: #0f172a;">1. Thư mục Dự án trên Google Drive</div>
            <div id="drive-folder-check-status" style="padding: 8px 12px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; color: #1e40af; font-size: 0.82rem;">
              <i class="fa-solid fa-spinner fa-spin"></i> Đang kiểm tra thư mục dự án hiện có trên Google Drive...
            </div>
          </div>

          <!-- BƯỚC 2: TRẠNG THÁI TÀI LIỆU TỪ DATA CPMS -->
          <div style="margin-bottom: 14px;">
            <div style="font-weight: 600; margin-bottom: 6px; color: #0f172a;">
              2. Tài liệu hồ sơ theo Link Local trong CPMS (${projectDocs.length} hồ sơ)
            </div>
            ${
              commonRootHint
                ? `<div style="font-size: 0.76rem; color: #475569; margin-bottom: 6px;">Thư mục gốc trong CPMS: <code class="font-mono">${commonRootHint}</code></div>`
                : ''
            }
            <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
              <span id="sync-local-dir-match-info" style="font-size: 0.8rem; color: #475569; flex: 1; min-width: 180px;">
                ${
                  this.localDirectoryHandle
                    ? `<span style="color: #166534; font-weight: 600;"><i class="fa-solid fa-circle-check"></i> Đã liên kết thư mục gốc: "${this.localDirectoryHandle.name}"</span>`
                    : projectDocs.length > 0
                      ? 'Trình duyệt cần cấp quyền thư mục gốc 1 lần để đọc file từ máy tính.'
                      : 'Không có hồ sơ đính kèm Local nào cần tải lên.'
                }
              </span>
              <button type="button" id="btn-pick-sync-local-dir" class="button button-outline button-sm" ${projectDocs.length === 0 ? 'disabled' : ''}>
                <i class="fa-solid fa-folder-open" style="color: #d97706;"></i> ${this.localDirectoryHandle ? 'Đổi thư mục gốc' : 'Cấp quyền thư mục gốc (1 lần)'}
              </button>
              <button type="button" id="btn-add-sync-local-dir" class="button button-outline button-sm" title="Quét gộp thêm thư mục khác nếu có file lưu rải rác" ${projectDocs.length === 0 ? 'disabled' : ''}>
                <i class="fa-solid fa-folder-plus" style="color: #2563eb;"></i> Quét thêm thư mục
              </button>
              <input type="file" id="input-fallback-local-dir" webkitdirectory multiple style="display: none;">
            </div>

            <div id="sync-missing-docs-preview" style="display: none;"></div>

            <label style="display: flex; align-items: center; gap: 6px; margin-top: 10px; font-size: 0.8rem; color: #334155; cursor: pointer;">
              <input type="checkbox" id="chk-skip-existing-drive-link" checked>
              <span>Bỏ qua các hồ sơ đã có Link Drive (tiếp tục tải các hồ sơ còn thiếu)</span>
            </label>
          </div>

          <div style="display: flex; justify-content: flex-end; gap: 8px; padding-top: 8px; border-top: 1px solid #e2e8f0;">
            <button type="button" id="btn-cancel-drive-sync" class="button button-outline">Đóng</button>
            <button type="button" id="btn-start-drive-sync" class="button button-primary">
              <i class="fa-solid fa-cloud-arrow-up"></i> Bắt đầu Đồng bộ ngầm lên Drive
            </button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-drive-sync-modal')?.addEventListener('click', closeModal);
    modal.querySelector('#btn-cancel-drive-sync')?.addEventListener('click', closeModal);

    let quanLyDuAnId = null;
    let matchedFolder = null;
    let activeLocalIndex = null;
    const manualPickedFiles = new Map(); // docId -> FileSystemFileHandle | File

    const renderMissingDocsPreview = (missingDocs) => {
      const box = modal.querySelector('#sync-missing-docs-preview');
      if (!box) return;
      if (!missingDocs || missingDocs.length === 0) {
        box.style.display = 'none';
        box.innerHTML = '';
        return;
      }
      box.style.display = 'block';
      box.innerHTML = `
        <div style="font-size: 0.78rem; font-weight: 600; color: #b45309; margin-top: 8px;">
          <i class="fa-solid fa-triangle-exclamation"></i> Còn ${missingDocs.length} hồ sơ chưa tìm thấy file (bạn có thể bấm "Quét thêm thư mục" hoặc "Chọn file lẻ"):
        </div>
        <div class="missing-doc-list">
          ${missingDocs
            .slice(0, 20)
            .map(
              (d) => `
            <div class="missing-doc-item">
              <div class="missing-doc-info">
                <div><strong>${d.doc_no || d.id}</strong> — ${d.title || d.name || ''}</div>
                <div class="missing-doc-path">${d.local_link}</div>
              </div>
              <button type="button" class="button button-outline button-sm btn-pick-single-missing" data-doc-id="${d.id}" style="flex-shrink: 0; font-size: 0.72rem; padding: 3px 8px;">
                <i class="fa-solid fa-file-arrow-up"></i> Chọn file lẻ
              </button>
            </div>
          `
            )
            .join('')}
          ${missingDocs.length > 20 ? `<div style="font-size: 0.74rem; color: #64748b;">...và ${missingDocs.length - 20} hồ sơ khác.</div>` : ''}
        </div>
      `;

      box.querySelectorAll('.btn-pick-single-missing').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const docId = btn.getAttribute('data-doc-id');
          if ('showOpenFilePicker' in window) {
            try {
              const [fh] = await window.showOpenFilePicker({ multiple: false });
              if (fh) {
                manualPickedFiles.set(docId, fh);
                await refreshMatchInfo(false);
              }
            } catch (e) {
              if (e.name !== 'AbortError') console.warn(e);
            }
          }
        });
      });
    };

    const refreshMatchInfo = async (allowPromptPermission = false) => {
      const infoEl = modal.querySelector('#sync-local-dir-match-info');
      if (!infoEl) return;

      if (!activeLocalIndex && this.localDirectoryHandle) {
        try {
          let perm = await this.localDirectoryHandle.queryPermission({ mode: 'read' });
          if (perm !== 'granted' && allowPromptPermission) {
            perm = await this.localDirectoryHandle.requestPermission({ mode: 'read' });
          }
          if (perm === 'granted') {
            infoEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Đang quét thư mục "${this.localDirectoryHandle.name}"...`;
            activeLocalIndex = await this.indexLocalFiles(this.localDirectoryHandle);
          } else {
            infoEl.innerHTML = `<span style="color: #b45309;"><i class="fa-solid fa-lock"></i> Đã nhớ thư mục "${this.localDirectoryHandle.name}" (sẽ xin quyền đọc khi bấm Bắt đầu)</span>`;
            return;
          }
        } catch (err) {
          console.warn('Lỗi kiểm tra quyền thư mục đã lưu:', err);
          return;
        }
      }

      if (activeLocalIndex) {
        let matched = 0;
        const missingDocs = [];
        for (const d of projectDocs) {
          if (manualPickedFiles.has(d.id) || this.matchLocalFileEntry(d.local_link, activeLocalIndex)) {
            matched++;
          } else {
            missingDocs.push(d);
          }
        }
        const dirLabel = this.localDirectoryHandle?.name || 'Thư mục đã chọn';
        infoEl.innerHTML = `<span style="color: #166534; font-weight: 600;"><i class="fa-solid fa-circle-check"></i> "${dirLabel}": Tìm thấy ${matched}/${projectDocs.length} file khớp</span>`;
        renderMissingDocsPreview(missingDocs);
      }
    };

    // Nếu đã có sẵn localDirectoryHandle từ trước -> tự động quét nếu đã có quyền
    if (this.localDirectoryHandle && projectDocs.length > 0) {
      refreshMatchInfo(false);
    }

    // Kiểm tra trùng thư mục dự án trên Google Drive (theo cả Mã dự án, Tên dự án và từ khóa chính)
    (async () => {
      const statusBox = modal.querySelector('#drive-folder-check-status');
      try {
        quanLyDuAnId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn']);
        const existingFolders = await gdrive.findFoldersInParent(quanLyDuAnId);
        const prjNameLower = String(curPrj?.name || '').trim().toLowerCase();
        const nameTokens = prjNameLower
          .replace(/dự án|nhà ở|chung cư|cao tầng|khu đô thị/gi, '')
          .trim()
          .split(/\s+/)
          .filter((w) => w.length >= 3);

        matchedFolder = existingFolders.find((f) => {
          const fname = String(f.name || '').toLowerCase();
          if (fname === defaultFolderName.toLowerCase() || fname.includes(safePrjId.toLowerCase())) return true;
          if (prjNameLower && fname.includes(prjNameLower)) return true;
          if (nameTokens.length >= 2 && nameTokens.every((t) => fname.includes(t))) return true;
          return false;
        });

        // Nếu chỉ có đúng 1 thư mục dự án trên Drive và chưa khớp tên -> vẫn gợi ý kiểm tra nếu tên là DuAn_CHC1 cho Westlake
        if (!matchedFolder && existingFolders.length === 1 && prjNameLower.includes('westlake') && existingFolders[0].name === 'DuAn_CHC1') {
          matchedFolder = existingFolders[0];
        }

        if (statusBox) {
          if (matchedFolder) {
            statusBox.innerHTML = `
              <div style="margin-bottom: 6px; font-weight: 600; color: #b45309;">
                <i class="fa-solid fa-triangle-exclamation"></i> Dự án đã tồn tại trên Drive: <code>${matchedFolder.name}</code>
              </div>
              <label style="display: flex; align-items: center; gap: 6px; cursor: pointer; margin-bottom: 4px;">
                <input type="radio" name="drive-folder-mode" value="overwrite" checked>
                <span><strong>Ghi đè / Cập nhật</strong> vào thư mục <code>${matchedFolder.name}</code> (Khuyên dùng)</span>
              </label>
              <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
                <input type="radio" name="drive-folder-mode" value="create_new">
                <span>Tạo thư mục bản sao mới (<code>${defaultFolderName}_${new Date().toISOString().slice(0, 10)}</code>)</span>
              </label>
            `;
            statusBox.style.background = '#fffbeb';
            statusBox.style.borderColor = '#fde68a';
            statusBox.style.color = '#92400e';
          } else {
            statusBox.innerHTML = `<i class="fa-solid fa-folder-plus"></i> Sẽ tạo mới 1 thư mục dự án: <code>${defaultFolderName}</code>`;
            statusBox.style.background = '#f0fdf4';
            statusBox.style.borderColor = '#bbf7d0';
            statusBox.style.color = '#166534';
          }
        }
      } catch (e) {
        if (statusBox) {
          const msg = String(e?.message || '');
          if (msg.includes('[401]')) {
            gdrive.accessToken = null;
            sessionStorage.removeItem('maio_gdrive_token');
            sessionStorage.removeItem('maio_gdrive_token_exp');
            statusBox.innerHTML = `
              <div style="color: #dc2626; margin-bottom: 6px;">
                <i class="fa-solid fa-circle-xmark"></i> Phiên đăng nhập Google Drive đã hết hạn (401).
              </div>
              <button type="button" id="btn-reauth-drive-in-modal" class="button button-primary button-sm">
                <i class="fa-brands fa-google"></i> Đăng nhập lại Google Drive
              </button>
            `;
            statusBox.querySelector('#btn-reauth-drive-in-modal')?.addEventListener('click', () => {
              closeModal();
              this.pendingSyncAction = 'to-drive';
              gdrive.signIn();
            });
          } else {
            statusBox.innerHTML = `<span style="color: #dc2626;"><i class="fa-solid fa-circle-xmark"></i> Lỗi kiểm tra Drive: ${msg}</span>`;
          }
        }
      }
    })();

    // Nút cấp quyền / đổi thư mục gốc Local & Quét thêm thư mục
    const pickLocalBtn = modal.querySelector('#btn-pick-sync-local-dir');
    const addLocalBtn = modal.querySelector('#btn-add-sync-local-dir');
    const fallbackInput = modal.querySelector('#input-fallback-local-dir');

    pickLocalBtn?.addEventListener('click', async () => {
      if ('showDirectoryPicker' in window) {
        try {
          const handle = await window.showDirectoryPicker();
          this.localDirectoryHandle = handle;
          await db.saveProjectDirHandle(curPrjId, handle);
          activeLocalIndex = null;
          pickLocalBtn.innerHTML = '<i class="fa-solid fa-folder-open" style="color: #d97706;"></i> Đổi thư mục gốc';
          await refreshMatchInfo(true);
        } catch (e) {
          if (e.name !== 'AbortError') {
            alert('Không thể mở thư mục máy tính: ' + e.message);
          }
        }
      } else {
        fallbackInput?.click();
      }
    });

    addLocalBtn?.addEventListener('click', async () => {
      if ('showDirectoryPicker' in window) {
        try {
          const extraHandle = await window.showDirectoryPicker();
          if (extraHandle) {
            const infoEl = modal.querySelector('#sync-local-dir-match-info');
            if (infoEl) infoEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Đang quét gộp thêm "${extraHandle.name}"...`;
            activeLocalIndex = await this.indexLocalFiles(extraHandle, extraHandle.name, activeLocalIndex);
            await refreshMatchInfo(false);
          }
        } catch (e) {
          if (e.name !== 'AbortError') alert('Không thể mở thư mục bổ sung: ' + e.message);
        }
      } else {
        fallbackInput?.click();
      }
    });

    fallbackInput?.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files.length > 0) {
        const extraIdx = this.indexFileList(e.target.files);
        if (!activeLocalIndex) {
          activeLocalIndex = extraIdx;
        } else {
          for (const [k, v] of extraIdx.byRelPath.entries()) activeLocalIndex.byRelPath.set(k, v);
          for (const [k, list] of extraIdx.byFileName.entries()) {
            const cur = activeLocalIndex.byFileName.get(k) || [];
            activeLocalIndex.byFileName.set(k, cur.concat(list));
          }
        }
        await refreshMatchInfo(false);
      }
    });

    // Nút Bắt đầu đồng bộ ngầm
    const startBtn = modal.querySelector('#btn-start-drive-sync');
    startBtn?.addEventListener('click', async () => {
      if (!activeLocalIndex && this.localDirectoryHandle && projectDocs.length > 0) {
        await refreshMatchInfo(true);
      }

      const mode = modal.querySelector('input[name="drive-folder-mode"]:checked')?.value || 'overwrite';
      const skipExisting = Boolean(modal.querySelector('#chk-skip-existing-drive-link')?.checked);

      closeModal();
      this.runBackgroundDriveSync({
        curPrjId,
        curPrj,
        safePrjId,
        defaultFolderName,
        quanLyDuAnId,
        matchedFolder,
        mode,
        skipExisting,
        projectDocs,
        activeLocalIndex,
        manualPickedFiles,
      });
    });
  }

  /**
   * Thực thi đồng bộ lên Google Drive ở chế độ chạy ngầm (concurrency = 3),
   * phân loại đúng cây thư mục 3 trụ cột và hiển thị Báo cáo khi hoàn tất
   */
  async runBackgroundDriveSync(ctx) {
    const {
      curPrjId,
      curPrj,
      safePrjId,
      defaultFolderName,
      matchedFolder,
      mode,
      skipExisting,
      projectDocs,
      activeLocalIndex,
      manualPickedFiles,
    } = ctx;

    document.getElementById('drive-bg-sync-toast')?.remove();
    const toast = document.createElement('div');
    toast.id = 'drive-bg-sync-toast';
    toast.className = 'drive-bg-sync-toast';
    toast.innerHTML = `
      <div class="drive-bg-sync-header">
        <span><i class="fa-solid fa-cloud-arrow-up fa-beat-fade" style="color: #2563eb;"></i> Đang đồng bộ ngầm lên Drive</span>
        <span id="bg-sync-pct" style="font-size: 0.78rem; color: #2563eb;">0%</span>
      </div>
      <div id="bg-sync-sub" class="drive-bg-sync-sub">Đang chuẩn bị cấu trúc thư mục...</div>
      <div class="drive-sync-progress-wrap" style="margin: 2px 0 0;">
        <div id="bg-sync-bar" class="drive-sync-progress-bar"></div>
      </div>
    `;
    document.body.appendChild(toast);

    const setBgProgress = (pct, text) => {
      const pctEl = document.getElementById('bg-sync-pct');
      const subEl = document.getElementById('bg-sync-sub');
      const barEl = document.getElementById('bg-sync-bar');
      if (pctEl) pctEl.textContent = `${Math.min(100, Math.max(0, Math.round(pct)))}%`;
      if (subEl) subEl.textContent = text;
      if (barEl) barEl.style.width = `${Math.min(100, Math.max(0, pct))}%`;
    };

    try {
      let quanLyDuAnId = ctx.quanLyDuAnId || (await gdrive.ensureFolderPath(['01_QuanLy_DuAn']));
      let targetProjectFolderId = null;
      let targetFolderName = defaultFolderName;

      if (matchedFolder && mode === 'overwrite') {
        targetProjectFolderId = matchedFolder.id;
        targetFolderName = matchedFolder.name;
      } else if (matchedFolder && mode === 'create_new') {
        targetFolderName = `${defaultFolderName}_${new Date().toISOString().slice(0, 10)}`;
        targetProjectFolderId = await gdrive.createFolder(targetFolderName, quanLyDuAnId);
      } else {
        targetProjectFolderId = await gdrive.createFolder(defaultFolderName, quanLyDuAnId);
      }

      const pathCache = new Map();
      const inFlightMap = new Map();
      const folderFilesCache = new Map();

      setBgProgress(5, 'Đang chuẩn bị thư mục CPMS_Data...');
      const cpmsFolderId = await gdrive.ensureFolderPath(['CPMS_Data'], targetProjectFolderId, pathCache, inFlightMap);

      let uploadedDocsCount = 0;
      let skippedAlreadySynced = 0;
      const missingDocs = [];
      const failedDocs = [];
      let authExpiredDuringSync = false;

      const docsToUpload = [];
      for (const doc of projectDocs) {
        if (skipExisting && doc.drive_link && String(doc.drive_link).includes('drive.google.com')) {
          skippedAlreadySynced++;
          continue;
        }
        const entry =
          (manualPickedFiles && manualPickedFiles.get(doc.id)) ||
          (activeLocalIndex ? this.matchLocalFileEntry(doc.local_link, activeLocalIndex) : null);

        if (!entry) {
          missingDocs.push(doc);
        } else {
          docsToUpload.push({ doc, entry });
        }
      }

      const totalToUpload = docsToUpload.length;
      let processedCount = 0;
      let nextIdx = 0;
      const CONCURRENCY = 3;

      const getCachedFolderFiles = async (folderId) => {
        if (!folderFilesCache.has(folderId)) {
          const files = await gdrive.listProjectFiles(folderId);
          folderFilesCache.set(folderId, files);
        }
        return folderFilesCache.get(folderId);
      };

      const worker = async () => {
        while (nextIdx < totalToUpload && !authExpiredDuringSync) {
          const currentIndex = nextIdx++;
          const { doc, entry } = docsToUpload[currentIndex];
          try {
            const fileObj = typeof entry.getFile === 'function' ? await entry.getFile() : entry;
            const pathSegs = contractsManager.getDrivePathSegmentsForDoc(doc);
            const targetFolderId = await gdrive.ensureFolderPath(pathSegs, targetProjectFolderId, pathCache, inFlightMap);
            const preloadedFiles = await getCachedFolderFiles(targetFolderId);

            const uploaded = await gdrive.uploadBinaryFile(
              fileObj.name,
              fileObj,
              fileObj.type || 'application/octet-stream',
              targetFolderId,
              preloadedFiles
            );

            if (uploaded && uploaded.webViewLink) {
              doc.drive_link = uploaded.webViewLink;
              doc.drive_file_id = uploaded.id;
              doc.updated_at = new Date().toISOString();
              await db.put('documents', doc);
              uploadedDocsCount++;
            }
          } catch (uploadErr) {
            const msg = String(uploadErr?.message || '');
            if (msg.includes('[401]')) {
              authExpiredDuringSync = true;
              gdrive.accessToken = null;
              sessionStorage.removeItem('maio_gdrive_token');
              sessionStorage.removeItem('maio_gdrive_token_exp');
            }
            failedDocs.push({ doc, error: msg });
          } finally {
            processedCount++;
            const pct = 8 + Math.round((processedCount / Math.max(1, totalToUpload)) * 78);
            setBgProgress(pct, `Đã xử lý ${processedCount}/${totalToUpload}: ${doc.doc_no || doc.title || doc.id}`);
          }
        }
      };

      if (totalToUpload > 0) {
        const workers = Array.from({ length: Math.min(CONCURRENCY, totalToUpload) }, () => worker());
        await Promise.all(workers);
      }

      if (uploadedDocsCount > 0) {
        setBgProgress(88, 'Đang cập nhật Link Drive vào Tab Hồ sơ & file CPMS...');
        await contractsManager.loadData();
        contractsManager.render();
        if (contractsManager.activeFileHandle) {
          await contractsManager.syncToLocalFile();
        }
      }

      let cpmsSavedOnline = false;
      if (!authExpiredDuringSync && gdrive.isSignedIn()) {
        setBgProgress(94, 'Đang lưu MAIO_CPMS_Snapshot và file .cpms lên CPMS_Data...');
        try {
          await db.syncToGoogleDrive({
            projectId: curPrjId,
            targetProjectFolderId,
            cpmsFolderId,
          });
          const cpmsBuffer = await contractsManager.buildProjectCPMSBuffer(curPrjId);
          const cpmsFileName = contractsManager.activeFileHandle?.name || `MAIO_Data_${safePrjId}.cpms`;
          await gdrive.uploadBinaryFile(cpmsFileName, cpmsBuffer, 'application/octet-stream', cpmsFolderId);
          cpmsSavedOnline = true;
        } catch (cpmsErr) {
          console.warn('Lỗi lưu CPMS_Data lên Drive:', cpmsErr);
        }
      }

      setBgProgress(100, authExpiredDuringSync ? 'Tạm dừng do hết phiên đăng nhập Drive!' : 'Đồng bộ hoàn tất!');
      this.updateProjectSourceBadge();
      setTimeout(() => toast.remove(), 1200);

      this.openDriveSyncReportModal({
        curPrjId,
        curPrj,
        targetFolderName,
        targetProjectFolderId,
        uploadedDocsCount,
        skippedAlreadySynced,
        missingDocs,
        failedDocs,
        cpmsSavedOnline,
        authExpiredDuringSync,
        pathCache,
        folderFilesCache,
        cpmsFolderId,
      });
    } catch (err) {
      toast.remove();
      alert('Lỗi đồng bộ ngầm lên Google Drive: ' + err.message);
    }
  }

  /**
   * Hiển thị Modal Báo cáo Kết quả Đồng bộ & cho phép tải bù ngay các file thiếu
   */
  openDriveSyncReportModal(report) {
    document.getElementById('drive-sync-report-modal')?.remove();
    const modal = document.createElement('div');
    modal.id = 'drive-sync-report-modal';
    modal.className = 'modal-overlay';

    const {
      curPrjId,
      curPrj,
      targetFolderName,
      targetProjectFolderId,
      uploadedDocsCount,
      skippedAlreadySynced,
      missingDocs,
      failedDocs,
      cpmsSavedOnline,
      authExpiredDuringSync,
      cpmsFolderId,
    } = report;

    modal.innerHTML = `
      <div class="modal-container" style="max-width: 620px;">
        <div class="modal-header">
          <h2 style="font-size: 1.08rem; font-weight: 700; display: flex; align-items: center; gap: 8px;">
            <i class="fa-solid ${authExpiredDuringSync ? 'fa-triangle-exclamation' : 'fa-circle-check'}" style="color: ${authExpiredDuringSync ? '#d97706' : '#16a34a'};"></i>
            Báo cáo Kết quả Đồng bộ Google Drive
          </h2>
          <button type="button" id="btn-close-sync-report" class="button button-outline" style="padding: 4px 10px;">
            <i class="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div class="modal-body" style="padding: 16px 20px; font-size: 0.85rem;">
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 14px; margin-bottom: 12px;">
            <div><strong>Dự án:</strong> ${curPrj?.name || curPrjId}</div>
            <div><strong>Thư mục Drive:</strong> <code>MAIO_Data / 01_QuanLy_DuAn / ${targetFolderName}</code></div>
            <div><strong>Dữ liệu CPMS_Data:</strong> ${cpmsSavedOnline ? '<span style="color: #166534; font-weight: 600;">Đã cập nhật Snapshot & .cpms</span>' : '<span style="color: #dc2626;">Chưa cập nhật được do gián đoạn kết nối</span>'}</div>
          </div>

          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 12px; text-align: center;">
            <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 8px;">
              <div style="font-size: 1.15rem; font-weight: 800; color: #15803d;">${uploadedDocsCount}</div>
              <div style="font-size: 0.74rem; color: #166534;">Đã tải lên mới</div>
            </div>
            <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 8px;">
              <div style="font-size: 1.15rem; font-weight: 800; color: #1d4ed8;">${skippedAlreadySynced}</div>
              <div style="font-size: 0.74rem; color: #1e40af;">Đã có sẵn Link</div>
            </div>
            <div style="background: #fffbeb; border: 1px solid #fde68a; border-radius: 6px; padding: 8px;">
              <div style="font-size: 1.15rem; font-weight: 800; color: #b45309;">${missingDocs.length + failedDocs.length}</div>
              <div style="font-size: 0.74rem; color: #92400e;">Chưa tìm thấy / Lỗi</div>
            </div>
          </div>

          ${
            authExpiredDuringSync
              ? `<div style="background: #fef2f2; border: 1px solid #fecaca; color: #b91c1c; border-radius: 6px; padding: 10px 12px; margin-bottom: 12px;">
                  <strong>Phiên đăng nhập Google Drive đã hết hạn giữa chừng!</strong> Tiến độ ${uploadedDocsCount} hồ sơ vừa tải lên đã được lưu. Hãy bấm <strong>Đăng nhập lại Google Drive</strong> và bấm Đồng bộ lại để tải tiếp phần còn lại.
                </div>`
              : ''
          }

          ${
            missingDocs.length > 0
              ? `<div>
                  <div style="font-weight: 600; color: #b45309; margin-bottom: 4px;">
                    Danh sách ${missingDocs.length} hồ sơ chưa tìm thấy file trên máy tính (có thể tải bù từng file):
                  </div>
                  <div class="missing-doc-list">
                    ${missingDocs
                      .map(
                        (d) => `
                      <div class="missing-doc-item" data-missing-row="${d.id}">
                        <div class="missing-doc-info">
                          <div><strong>${d.doc_no || d.id}</strong> — ${d.title || d.name || ''}</div>
                          <div class="missing-doc-path">${d.local_link}</div>
                        </div>
                        <button type="button" class="button button-outline button-sm btn-upload-missing-single" data-doc-id="${d.id}" style="flex-shrink: 0; font-size: 0.72rem; padding: 3px 8px;">
                          <i class="fa-solid fa-upload"></i> Tải bù
                        </button>
                      </div>
                    `
                      )
                      .join('')}
                  </div>
                </div>`
              : ''
          }

          ${
            failedDocs.length > 0
              ? `<div style="margin-top: 10px; font-size: 0.78rem; color: #dc2626;">
                  <strong>${failedDocs.length} hồ sơ gặp lỗi khi upload:</strong> ${failedDocs[0]?.error || ''}
                </div>`
              : ''
          }

          <div style="display: flex; justify-content: flex-end; gap: 8px; padding-top: 12px; margin-top: 12px; border-top: 1px solid #e2e8f0;">
            ${
              authExpiredDuringSync
                ? `<button type="button" id="btn-report-reauth-drive" class="button button-primary"><i class="fa-brands fa-google"></i> Đăng nhập lại & Chạy tiếp</button>`
                : ''
            }
            <button type="button" id="btn-ok-sync-report" class="button button-outline">Đóng</button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const close = () => modal.remove();
    modal.querySelector('#btn-close-sync-report')?.addEventListener('click', close);
    modal.querySelector('#btn-ok-sync-report')?.addEventListener('click', close);

    modal.querySelector('#btn-report-reauth-drive')?.addEventListener('click', () => {
      close();
      this.pendingSyncAction = 'to-drive';
      gdrive.signIn();
    });

    modal.querySelectorAll('.btn-upload-missing-single').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const docId = btn.getAttribute('data-doc-id');
        const doc = missingDocs.find((d) => d.id === docId);
        if (!doc || !('showOpenFilePicker' in window)) return;
        try {
          const [fh] = await window.showOpenFilePicker({ multiple: false });
          if (!fh) return;
          btn.disabled = true;
          btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
          const fileObj = await fh.getFile();
          const pathSegs = contractsManager.getDrivePathSegmentsForDoc(doc);
          const folderId = await gdrive.ensureFolderPath(pathSegs, targetProjectFolderId);
          const uploaded = await gdrive.uploadBinaryFile(fileObj.name, fileObj, fileObj.type || 'application/octet-stream', folderId);
          if (uploaded && uploaded.webViewLink) {
            doc.drive_link = uploaded.webViewLink;
            doc.drive_file_id = uploaded.id;
            doc.updated_at = new Date().toISOString();
            await db.put('documents', doc);
            await contractsManager.loadData();
            contractsManager.render();
            if (contractsManager.activeFileHandle) await contractsManager.syncToLocalFile();
            await db.syncToGoogleDrive({ projectId: curPrjId, targetProjectFolderId, cpmsFolderId });
            const cpmsBuffer = await contractsManager.buildProjectCPMSBuffer(curPrjId);
            await gdrive.uploadBinaryFile(contractsManager.activeFileHandle?.name || `MAIO_Data_${curPrjId}.cpms`, cpmsBuffer, 'application/octet-stream', cpmsFolderId);
            btn.outerHTML = '<span style="color: #16a34a; font-weight: 600;"><i class="fa-solid fa-check"></i> Đã tải</span>';
          }
        } catch (e) {
          if (e.name !== 'AbortError') {
            alert('Lỗi tải bù file: ' + e.message);
            btn.disabled = false;
            btn.innerHTML = '<i class="fa-solid fa-upload"></i> Tải bù';
          }
        }
      });
    });
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
        const confirmPickDir = await (window.maioConfirm
          ? window.maioConfirm({
              title: 'Chọn thư mục lưu trữ trên máy tính',
              message: 'Bạn có muốn chọn thư mục trên máy tính để tải cây thư mục hồ sơ từ Drive về và cập nhật lại đường dẫn Local trong CPMS không?',
              confirmText: 'Chọn thư mục',
              cancelText: 'Bỏ qua',
              confirmType: 'primary',
            })
          : Promise.resolve(confirm('Bạn có muốn chọn thư mục trên máy tính để đồng bộ cây thư mục hồ sơ về không?')));
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
  window.__MAIO_APP__ = app;
  app.init();
});
