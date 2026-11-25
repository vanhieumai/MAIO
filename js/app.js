/**
 * MAIO Application Entry Point
 * Quản lý khởi tạo IndexedDB, nạp dữ liệu hạt giống CHC1, điều hướng phân hệ và kết nối đồng bộ.
 */

import { gdrive } from './core/gdrive.js?v=5.1';
import { db } from './core/db.js?v=5.1';
import { contractsManager } from './modules/cpms/contracts.js?v=7.3';

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
   * Đồng bộ dữ liệu hiện tại từ máy tính lên Google Drive:
   * - Đảm bảo chỉ tạo 1 thư mục dự án duy nhất; hỏi ghi đè nếu dự án đã tồn tại trên Drive
   * - Tạo cấu trúc chuẩn: CPMS_Data, HoSo, HopDong_PL, ThanhToan
   * - Upload file từ Link Local lên đúng thư mục Drive, cập nhật drive_link vào Tab Hồ sơ & file CPMS
   */
  async handleSyncToDrive() {
    const btn = document.getElementById('btn-sync-to-drive');
    const origHtml = btn ? btn.innerHTML : '';
    try {
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đồng bộ...';
      }

      const curPrjId = contractsManager.currentProjectId || db.activeProjectId || 'CHC1';
      db.activeProjectId = curPrjId;
      const curPrj = (contractsManager.projects || []).find((p) => p.id === curPrjId);
      const safePrjId = String(curPrjId).replace(/[\/\\?%*:|"<>]/g, '_');
      const defaultFolderName = `DuAn_${safePrjId}`;

      // 1. Kiểm tra thư mục dự án trên Google Drive để tránh tạo trùng
      const quanLyDuAnId = await gdrive.ensureFolderPath(['01_QuanLy_DuAn']);
      const existingFolders = await gdrive.findFoldersInParent(quanLyDuAnId);
      const prjNameLower = String(curPrj?.name || '').trim().toLowerCase();
      const matchedFolder = existingFolders.find((f) => {
        const fname = String(f.name || '').toLowerCase();
        return (
          fname === defaultFolderName.toLowerCase() ||
          fname.includes(safePrjId.toLowerCase()) ||
          (prjNameLower && fname.includes(prjNameLower))
        );
      });

      let targetProjectFolderId = null;
      let targetFolderName = defaultFolderName;

      if (matchedFolder) {
        const overwrite = confirm(
          `Phát hiện dự án "${curPrj?.name || curPrjId}" đã tồn tại trên Google Drive trong thư mục "${matchedFolder.name}".\n\n` +
            `• Bấm OK để GHI ĐÈ (cập nhật) vào thư mục "${matchedFolder.name}".\n` +
            `• Bấm Cancel để tạo thư mục mới hoặc hủy thao tác.`
        );
        if (overwrite) {
          targetProjectFolderId = matchedFolder.id;
          targetFolderName = matchedFolder.name;
        } else {
          const createNew = confirm(
            `Bạn có muốn tạo một thư mục bản sao mới cho dự án "${curPrjId}" trên Google Drive không?\n\n` +
              `• Bấm OK để tạo thư mục mới\n` +
              `• Bấm Cancel để Hủy đồng bộ`
          );
          if (!createNew) return;
          const stamp = new Date().toISOString().slice(0, 10);
          targetFolderName = `${defaultFolderName}_${stamp}`;
          targetProjectFolderId = await gdrive.createFolder(targetFolderName, quanLyDuAnId);
        }
      } else {
        targetProjectFolderId = await gdrive.createFolder(defaultFolderName, quanLyDuAnId);
      }

      // 2. Tạo/Đảm bảo 4 thư mục con chuẩn bên trong thư mục dự án
      const cpmsFolderId = await gdrive.ensureFolderPath(['CPMS_Data'], targetProjectFolderId);
      const hoSoFolderId = await gdrive.ensureFolderPath(['HoSo'], targetProjectFolderId);
      const hopDongFolderId = await gdrive.ensureFolderPath(['HopDong_PL'], targetProjectFolderId);
      const thanhToanFolderId = await gdrive.ensureFolderPath(['ThanhToan'], targetProjectFolderId);
      const subFolderIds = {
        HoSo: hoSoFolderId,
        HopDong_PL: hopDongFolderId,
        ThanhToan: thanhToanFolderId,
      };

      // 3. Đồng bộ các tài liệu từ Link Local lên Google Drive & cập nhật Link Drive
      const projectDocs = (contractsManager.documents || []).filter(
        (d) => (!d.project_id || d.project_id === curPrjId) && d.local_link && !String(d.local_link).startsWith('http')
      );

      let uploadedDocsCount = 0;
      let missingLocalCount = 0;

      if (projectDocs.length > 0) {
        if (!this.localDirectoryHandle && 'showDirectoryPicker' in window) {
          const wantPick = confirm(
            `Tìm thấy ${projectDocs.length} hồ sơ có đường dẫn file trên máy tính (Link Local).\n\n` +
              `Bạn có muốn chọn thư mục gốc trên máy tính chứa các tài liệu này để tự động tải lên Google Drive (HoSo / HopDong_PL / ThanhToan) và cập nhật Link Drive vào Tab Hồ sơ không?\n\n` +
              `• Bấm OK để chọn thư mục chứa tài liệu trên máy tính\n` +
              `• Bấm Cancel để chỉ đồng bộ dữ liệu CPMS (bỏ qua file đính kèm)`
          );
          if (wantPick) {
            try {
              this.localDirectoryHandle = await window.showDirectoryPicker();
            } catch (e) {
              if (e.name !== 'AbortError') console.warn('Lỗi chọn thư mục local:', e);
            }
          }
        }

        if (this.localDirectoryHandle) {
          if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang quét & tải hồ sơ...';
          const localIndex = await this.indexLocalFiles(this.localDirectoryHandle);

          for (let i = 0; i < projectDocs.length; i++) {
            const doc = projectDocs[i];
            const fileHandle = this.matchLocalFileHandle(doc.local_link, localIndex);
            if (!fileHandle) {
              missingLocalCount++;
              continue;
            }

            try {
              if (btn) {
                btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Đang tải (${i + 1}/${projectDocs.length})...`;
              }
              const fileObj = await fileHandle.getFile();
              const subKey = contractsManager.getDriveSubfolderKeyForDoc(doc);
              const targetFolder = subFolderIds[subKey] || hoSoFolderId;
              const uploaded = await gdrive.uploadBinaryFile(
                fileObj.name,
                fileObj,
                fileObj.type || 'application/octet-stream',
                targetFolder
              );

              if (uploaded && uploaded.webViewLink) {
                doc.drive_link = uploaded.webViewLink;
                doc.drive_file_id = uploaded.id;
                doc.updated_at = new Date().toISOString();
                await db.put('documents', doc);
                uploadedDocsCount++;
              }
            } catch (uploadErr) {
              console.warn(`Lỗi tải tài liệu "${doc.doc_no || doc.id}" lên Drive:`, uploadErr);
            }
          }
        }
      }

      // 4. Cập nhật lại giao diện Tab Hồ sơ & file data CPMS trên máy (nếu có file local đang mở)
      if (uploadedDocsCount > 0) {
        await contractsManager.loadData();
        contractsManager.render();
        if (contractsManager.activeFileHandle) {
          await contractsManager.syncToLocalFile();
        }
      }

      // 5. Lưu snapshot JSON mới nhất (đã gồm toàn bộ drive_link) vào CPMS_Data
      if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang lưu CPMS Data...';
      await db.syncToGoogleDrive({
        projectId: curPrjId,
        targetProjectFolderId,
        cpmsFolderId,
      });

      // 6. Nếu đang mở file CPMS cục bộ, upload bản mới nhất (đã cập nhật drive_link) vào CPMS_Data
      if (contractsManager.activeFileHandle) {
        try {
          const file = await contractsManager.activeFileHandle.getFile();
          const fileName = file.name || `MAIO_Data_${safePrjId}.cpms`;
          await gdrive.uploadBinaryFile(fileName, file, 'application/octet-stream', cpmsFolderId);
        } catch (e) {
          console.warn('Lỗi tải kèm file nhị phân CPMS lên Drive:', e);
        }
      }

      let summaryMsg = `Đã đồng bộ thành công dự án "${curPrj?.name || curPrjId}" lên Google Drive!\n\n`;
      summaryMsg += `• Thư mục Drive: MAIO_Data / 01_QuanLy_DuAn / ${targetFolderName}\n`;
      summaryMsg += `• Cấu trúc: CPMS_Data, HoSo, HopDong_PL, ThanhToan\n`;
      if (uploadedDocsCount > 0) {
        summaryMsg += `• Đã tải lên & cập nhật Link Drive cho ${uploadedDocsCount} tài liệu hồ sơ.\n`;
      }
      if (missingLocalCount > 0) {
        summaryMsg += `• Có ${missingLocalCount} tài liệu không tìm thấy file trong thư mục máy tính đã chọn.\n`;
      }
      alert(summaryMsg);
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
