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

import { db } from '../../core/db.js';
import { CPMSImporterExporter } from '../../core/cpms-importer.js';

export class ContractsManager {
  constructor() {
    this.currentProjectId = 'CHC1';
    this.projects = [];
    this.currentGroup = 'A';
    this.searchQuery = '';
    this.contracts = [];
    this.payments = [];
    this.documents = [];
    this.selectedContract = null;
    this.expandedContracts = new Set();
  }

  async init() {
    await this.loadData();
    this.bindEvents();
    this.render();
  }

  async loadData() {
    this.projects = await db.getAll('projects');
    this.updateProjectSelectorUI();

    const allContracts = await db.getAll('contracts');
    const allPayments = await db.getAll('payments');
    const allDocs = await db.getAll('documents');

    // Lọc theo currentProjectId
    this.contracts = allContracts.filter(
      (c) => !c.project_id || c.project_id === this.currentProjectId
    );
    this.payments = allPayments.filter(
      (p) => !p.project_id || p.project_id === this.currentProjectId
    );
    this.documents = allDocs.filter(
      (d) => !d.project_id || d.project_id === this.currentProjectId
    );
  }

  updateProjectSelectorUI() {
    const selector = document.getElementById('project-selector');
    if (!selector) return;

    let optionsHtml = '';
    const projectsList = this.projects.length > 0 ? this.projects : [{ id: 'CHC1', name: 'Dự án CHC1 (Thiết kế & Thi công)' }];

    // Đảm bảo currentProjectId hợp lệ
    if (!projectsList.some((p) => p.id === this.currentProjectId)) {
      this.currentProjectId = projectsList[0].id;
    }

    projectsList.forEach((p) => {
      const isSel = p.id === this.currentProjectId ? 'selected' : '';
      optionsHtml += `<option value="${p.id}" ${isSel}>📁 ${p.name || p.id}</option>`;
    });

    optionsHtml += `<option value="IMPORT_NEW">+ Nhập dữ liệu CPMS mới (.cpms)...</option>`;
    selector.innerHTML = optionsHtml;

    // Cập nhật tên dự án trên Dashboard nếu có
    const activePrjName = document.getElementById('active-project-name');
    if (activePrjName) {
      const curPrj = projectsList.find((p) => p.id === this.currentProjectId);
      if (curPrj) activePrjName.textContent = curPrj.name || curPrj.id;
    }
  }

  bindEvents() {
    const tabBtns = document.querySelectorAll('.group-tab-btn');
    tabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        tabBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        this.currentGroup = btn.getAttribute('data-group');
        this.render();
      });
    });

    const searchInput = document.getElementById('contract-search-input');
    searchInput?.addEventListener('input', (e) => {
      this.searchQuery = e.target.value.toLowerCase().trim();
      this.render();
    });

    document.getElementById('btn-add-contract')?.addEventListener('click', () => {
      this.openContractFormModal();
    });

    // Lắng nghe chọn dự án
    const projectSelector = document.getElementById('project-selector');
    projectSelector?.addEventListener('change', async (e) => {
      const val = e.target.value;
      if (val === 'IMPORT_NEW') {
        // Trả lại giá trị trước đó trên UI
        projectSelector.value = this.currentProjectId;
        this.openProjectModal();
      } else {
        this.currentProjectId = val;
        await this.loadData();
        this.render();
      }
    });

    // Nút Import .cpms
    document.getElementById('btn-import-project')?.addEventListener('click', () => {
      this.openProjectModal();
    });

    // Nút Sao chép Dự án
    document.getElementById('btn-clone-project')?.addEventListener('click', () => {
      this.openCloneProjectModal();
    });

    // Nút Xuất CPMS (Tạo file .cpms chuẩn để Import vào CPMS Plus)
    document.getElementById('btn-export-cpms')?.addEventListener('click', async () => {
      try {
        const btn = document.getElementById('btn-export-cpms');
        const origText = btn.innerHTML;
        btn.disabled = true;
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang xuất...';
        const exportedFile = await CPMSImporterExporter.exportToCPMSFormat(this.currentProjectId, 'cpms');
        btn.disabled = false;
        btn.innerHTML = origText;
        alert(`Đã xuất thành công file "${exportedFile}"!\n\nFile này tương thích 100% với tính năng Import của CPMS Plus, đồng thời chứa đầy đủ các cột và công thức tính toán.`);
      } catch (err) {
        alert('Lỗi xuất file CPMS: ' + err.message);
      }
    });

    // Nút Xóa Dự án
    document.getElementById('btn-delete-project')?.addEventListener('click', () => {
      this.deleteCurrentProject();
    });
  }

  async deleteCurrentProject() {
    const curProject = this.projects.find((p) => p.id === this.currentProjectId) || { id: this.currentProjectId, name: this.currentProjectId };
    const prjName = curProject.name || curProject.id;

    const confirmed = confirm(
      `⚠️ CẢNH BÁO XÓA DỰ ÁN!\n\nBạn có chắc chắn muốn xóa vĩnh viễn dự án "${prjName}" (Mã: ${this.currentProjectId}) không?\n\nHành động này sẽ xóa vĩnh viễn:\n- Toàn bộ danh mục Hợp đồng & Phụ lục của dự án\n- Toàn bộ Sổ theo dõi thanh toán (IPC)\n- Toàn bộ Hồ sơ & Công văn liên quan\n\nDữ liệu sau khi xóa sẽ không thể phục hồi!`
    );

    if (!confirmed) return;

    try {
      if (typeof db.deleteProjectCascade === 'function') {
        await db.deleteProjectCascade(this.currentProjectId);
      } else {
        // Fallback an toàn nếu trình duyệt chưa reload module db.js
        await db.delete('projects', this.currentProjectId);
        const allContracts = await db.getAll('contracts');
        for (const c of allContracts) {
          if (c.project_id === this.currentProjectId) await db.delete('contracts', c.id);
        }
        const allPayments = await db.getAll('payments');
        for (const p of allPayments) {
          if (p.project_id === this.currentProjectId) await db.delete('payments', p.id);
        }
        const allDocs = await db.getAll('documents');
        for (const d of allDocs) {
          if (d.project_id === this.currentProjectId) await db.delete('documents', d.id);
        }
      }

      // Cập nhật lại danh sách dự án
      this.projects = await db.getAll('projects');
      if (this.projects.length > 0) {
        this.currentProjectId = this.projects[0].id;
      } else {
        // Nếu không còn dự án nào, tạo một dự án mặc định mới
        const defaultProj = {
          id: 'PRJ_DEFAULT',
          code: 'PRJ_DEFAULT',
          name: 'Dự án Mới (Chưa có dữ liệu)',
          created_at: new Date().toISOString(),
        };
        await db.put('projects', defaultProj);
        this.projects = [defaultProj];
        this.currentProjectId = defaultProj.id;
      }

      await this.loadData();
      this.render();

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
    }
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
      d = new Date(Math.round((val - 25569) * 86400 * 1000));
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
      d = new Date(Math.round((val - 25569) * 86400 * 1000));
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

  getDrivePathForContract(contract) {
    const curPrjId = this.currentProjectId || 'CHC1';
    const isMain = !contract.parent_id;
    const parent = !isMain ? this.contracts.find((c) => c.id === contract.parent_id) : null;
    const cleanNo = (no) => String(no || '').replace(/[\/\\?%*:|"<>]/g, '_');

    const groupFolder = contract.group === 'A' ? '01_Nhom_A_ChuDauTu' : '02_Nhom_B_ThauPhu_NCC';

    if (contract.group === 'A') {
      if (isMain) {
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cleanNo(contract.contract_no)} / HopDong_Goc /`;
      } else {
        const parentNo = parent ? cleanNo(parent.contract_no) : cleanNo(contract.parent_id);
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${parentNo} / PhuLuc_${cleanNo(contract.contract_no)} /`;
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
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${cleanNo(contract.contract_no)}_${partnerName} / HopDong_Goc /`;
      } else {
        const rootMain = parent && parent.parent_id ? this.contracts.find((c) => c.id === parent.parent_id) || parent : parent;
        const parentNo = rootMain ? cleanNo(rootMain.contract_no) : cleanNo(contract.parent_id);
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${parentNo} / PhuLuc_${cleanNo(contract.contract_no)} /`;
      }
    }
  }

  renderProjectInfo() {
    const infoContainer = document.getElementById('project-info-bar');
    if (!infoContainer) return;

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || {
      id: this.currentProjectId,
      code: this.currentProjectId,
      name: 'Dự án ' + this.currentProjectId,
    };

    // Nếu dự án có mã CHC1 mà chưa có địa điểm hoặc địa điểm mặc định
    if ((!curPrj.location || curPrj.location === 'Việt Nam' || curPrj.location === '—') && String(curPrj.id).includes('CHC1')) {
      curPrj.location = 'Ô đất CHC1, Khu D6, Phú Thượng, Tây Hồ, Hà Nội';
    }
    // Nếu tên dự án chưa có hoặc chỉ là mã ID CHC1...
    if ((!curPrj.name || curPrj.name === curPrj.id || curPrj.name.startsWith('CHC1_')) && String(curPrj.id).includes('CHC1')) {
      curPrj.name = 'Dự án Nhà ở Chung cư Cao tầng Westlake Residential';
    }

    const sampleContract = this.contracts[0] || {};
    const clientName = curPrj.client || sampleContract.company_a || 'Chủ đầu tư';
    const location = curPrj.location || 'Việt Nam';

    // 1. Cập nhật thông tin cơ bản trên Top Banner (Vùng hiển thị cơ bản phía trên)
    const topProjName = document.getElementById('top-project-name');
    const topProjLoc = document.getElementById('top-project-location');
    if (topProjName) topProjName.textContent = curPrj.name || curPrj.id;
    if (topProjLoc) topProjLoc.textContent = location;

    // 2. Thống kê chi tiết Hợp đồng & Phụ lục nhóm A (Chủ đầu tư / Tổng thầu)
    const mainA = this.contracts.filter((c) => !c.parent_id && c.group === 'A');
    const appA = this.contracts.filter((c) => c.parent_id && c.group === 'A');
    let totalSignedA = 0;
    let totalPaidA = 0;
    mainA.forEach((c) => {
      const st = this.calculateTotals(c);
      totalSignedA += st.totalSigned;
      totalPaidA += st.totalPaid;
    });

    // 3. Thống kê chi tiết Hợp đồng & Phụ lục nhóm B (Nhà thầu phụ / Nhà cung cấp)
    const mainB = this.contracts.filter((c) => !c.parent_id && c.group === 'B');
    const appB = this.contracts.filter((c) => c.parent_id && c.group === 'B');
    let totalSignedB = 0;
    let totalPaidB = 0;
    mainB.forEach((c) => {
      const st = this.calculateTotals(c);
      totalSignedB += st.totalSigned;
      totalPaidB += st.totalPaid;
    });

    // 4. Render thanh Project Info Bar chi tiết bên dưới
    infoContainer.innerHTML = `
      <div class="project-info-meta">
        <span class="proj-meta-item">
          <i class="fa-solid fa-hashtag" style="color: var(--primary);"></i> Mã DA: <strong>${curPrj.code || curPrj.id}</strong>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-file-contract" style="color: #2563eb;"></i> HĐ với A: <strong>${mainA.length} HĐ gốc${appA.length > 0 ? ` (${appA.length} PL)` : ''}</strong> — <span style="color: var(--primary); font-weight: 700;">${this.formatCurrency(totalSignedA)}</span>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-handshake" style="color: #6366f1;"></i> HĐ với B: <strong>${mainB.length} HĐ gốc${appB.length > 0 ? ` (${appB.length} PL)` : ''}</strong> — <span style="color: #4f46e5; font-weight: 700;">${this.formatCurrency(totalSignedB)}</span>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-circle-down" style="color: #10b981;"></i> Bên A giải ngân: <strong style="color: var(--success);">${this.formatCurrency(totalPaidA)}</strong>
        </span>
        <span class="proj-meta-divider">|</span>
        <span class="proj-meta-item">
          <i class="fa-solid fa-circle-up" style="color: #06b6d4;"></i> Giải ngân cho bên B: <strong style="color: #0891b2;">${this.formatCurrency(totalPaidB)}</strong>
        </span>
      </div>
      <div style="flex-shrink: 0; margin-left: auto;">
        <button id="btn-edit-project-info" class="button button-outline" style="font-size: 0.75rem; padding: 4px 10px; height: 28px;" title="Chỉnh sửa thông tin dự án">
          <i class="fa-solid fa-pen-to-square"></i> Sửa thông tin
        </button>
      </div>
    `;

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
              <label class="form-label" style="font-size: 0.8rem; font-weight: 600;">Mã dự án (Project ID)</label>
              <input type="text" class="form-control" name="project_id" value="${project.id}" readonly style="background: #f1f5f9; cursor: not-allowed;">
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
      const updatedPrj = {
        ...project,
        name: fd.get('project_name').trim(),
        client: fd.get('client').trim(),
        location: fd.get('location').trim(),
        contractor: fd.get('contractor').trim(),
        updated_at: new Date().toISOString(),
      };
      await db.put('projects', updatedPrj);
      await this.loadData();
      closeModal();
      this.render();
      alert('Đã cập nhật thông tin dự án thành công!');
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
            <i class="fa-solid fa-copy" style="color: #6366f1;"></i> Sao chép (Clone) ${isMain ? 'Hợp đồng gốc' : 'Phụ lục'}
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
                <input type="number" class="form-control" name="contract_val" value="${sourceContract.contract_val || 0}">
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

    const closeModal = () => modal.remove();
    modal.querySelector('#btn-close-clone-ctr').onclick = closeModal;
    modal.querySelector('#btn-cancel-clone-ctr').onclick = closeModal;

    modal.querySelector('#form-clone-contract').onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const newId = fd.get('new_contract_id').trim();
      const newNo = fd.get('new_contract_no').trim();
      const newScope = fd.get('new_scope').trim();
      const val = parseFloat(fd.get('contract_val') || 0);
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

        alert(`🎉 Sao chép Hợp đồng thành công!\n\nĐã tạo "${newNo}" (Mã: ${newId}).`);
        this.openContractDetail(newContract.id);
      } catch (err) {
        alert('Lỗi sao chép Hợp đồng: ' + err.message);
      }
    };
  }

  getMainContracts() {
    return this.contracts.filter((c) => {
      const matchGroup = c.group === this.currentGroup;
      const isMain = !c.parent_id;
      const matchSearch =
        !this.searchQuery ||
        c.contract_no?.toLowerCase().includes(this.searchQuery) ||
        c.scope_of_work?.toLowerCase().includes(this.searchQuery) ||
        c.company_b?.toLowerCase().includes(this.searchQuery) ||
        c.id?.toLowerCase().includes(this.searchQuery);

      return matchGroup && isMain && matchSearch;
    });
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
  calculateTotals(contract) {
    const isMain = !contract.parent_id;
    const directChildren = this.getDirectChildren(contract.id);

    const baseVal = contract.contract_val || 0;
    const appendixSum = directChildren.reduce((sum, ch) => sum + (ch.contract_val || 0), 0);

    // Giá trị hợp đồng chuẩn xác:
    // Nếu HĐ đã có giá trị tạm tính ban đầu (như 1.200 tỷ) thì giữ nguyên, không cộng lẫn phụ lục.
    // Nếu HĐ khung không có giá trị (baseVal === 0) thì mới lấy tổng các phụ lục con.
    let totalSigned = baseVal > 0 ? baseVal : appendixSum;

    // Lấy thanh toán thực tế
    const realPayments = this.getRealPayments(contract.id);
    const totalAccepted = realPayments.reduce((sum, p) => sum + (p.accepted_vol || 0), 0);
    const totalRequested = realPayments.reduce((sum, p) => sum + (p.request_amount || 0), 0);
    const totalPaid = realPayments.reduce((sum, p) => sum + (p.paid_amount || 0), 0);
    const remaining = totalRequested - totalPaid;
    const paymentRate = totalSigned > 0 ? (totalPaid / totalSigned) * 100 : 0;

    return {
      baseVal,
      appendixSum,
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
    this.renderProjectInfo();

    const container = document.getElementById('contracts-list-container');
    if (!container) return;

    const mainContracts = this.getMainContracts();

    if (mainContracts.length === 0) {
      container.innerHTML = `
        <div style="text-align: center; padding: 48px; background: #fff; border-radius: var(--radius-md); border: 1px dashed var(--border-color);">
          <i class="fa-solid fa-folder-open" style="font-size: 2.5rem; color: var(--text-light); margin-bottom: 12px;"></i>
          <p style="color: var(--text-muted); font-weight: 500;">Không tìm thấy hợp đồng nào phù hợp trong Nhóm ${this.currentGroup}.</p>
        </div>
      `;
      return;
    }

    container.innerHTML = mainContracts
      .map((c) => {
        const stats = this.calculateTotals(c);
        const directChildren = this.getDirectChildren(c.id);
        const isExpanded = this.expandedContracts.has(c.id);

        return `
          <div class="contract-group-wrapper" data-contract-id="${c.id}">
            <!-- Card Hợp đồng gốc -->
            <div class="contract-card-main" data-open-detail="${c.id}">
              <div class="contract-header">
                <div>
                  <div class="contract-no">
                    <i class="fa-solid fa-file-signature"></i>
                    <span>${c.contract_no || c.id}</span>
                    <span class="badge ${c.group === 'A' ? 'badge-primary' : 'badge-warning'}">Nhóm ${c.group}</span>
                    <span class="badge badge-success">${c.status || 'Active'}</span>
                  </div>
                  <div class="contract-title">${c.scope_of_work || 'Hợp đồng thực hiện dự án'}</div>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                  <button class="button button-outline btn-clone-contract-action" data-clone-cid="${c.id}" style="font-size: 0.75rem; padding: 4px 8px; height: 28px;" title="Sao chép Hợp đồng này">
                    <i class="fa-solid fa-copy" style="color: #6366f1;"></i> Sao chép HĐ
                  </button>
                  <div style="text-align: right;">
                    <span style="font-size: 0.725rem; color: var(--text-muted); text-transform: uppercase;">Giá trị Hợp đồng gốc</span>
                    <div style="font-size: 1.15rem; font-weight: 800; color: var(--primary);">
                      ${this.formatCurrency(stats.totalSigned)}
                    </div>
                    ${directChildren.length > 0 ? `<div style="font-size: 0.725rem; color: var(--text-muted); font-weight: 500;">(Đã giao ${directChildren.length} PL: ${this.formatCurrency(stats.appendixSum)})</div>` : ''}
                  </div>
                </div>
              </div>

              <div class="contract-meta">
                <span><i class="fa-solid fa-handshake"></i> Bên A: <strong>${c.company_a || 'Chủ đầu tư'}</strong></span>
                <span><i class="fa-solid fa-building"></i> Bên B: <strong>${c.company_b || 'Đơn vị thực hiện'}</strong></span>
                <span><i class="fa-solid fa-calendar-check"></i> Ngày ký: ${this.formatDate(c.signing_date)}</span>
              </div>

              <div class="contract-stats">
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Phụ lục cấp 1</span>
                  <span class="contract-stat-val" style="color: var(--primary);">${stats.directChildrenCount} phụ lục</span>
                </div>
                <div class="contract-stat-item">
                  <span class="contract-stat-label">Đợt thanh toán (IPC)</span>
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
                  <span>${isExpanded ? 'Thu gọn' : 'Xem danh sách'} <strong>${directChildren.length} phụ lục con</strong> (Bấm vào bất kỳ dòng nào để vào riêng phụ lục đó)</span>
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

    // Bắt sự kiện Toggle Cây phụ lục
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

    // Bắt sự kiện Click vào HĐ gốc hoặc Phụ lục
    container.querySelectorAll('[data-open-detail]').forEach((el) => {
      el.addEventListener('click', (e) => {
        if (e.target.closest('[data-toggle-appendix]') || e.target.closest('.btn-clone-contract-action')) return;
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
              <span class="badge ${child.status === 'Active' ? 'badge-success' : 'badge-warning'}">${child.status || 'Active'}</span>
            </td>
            <td class="col-app-val">${this.formatCurrency(child.contract_val)}</td>
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

    let modal = document.getElementById('contract-detail-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'contract-detail-modal';
      modal.className = 'modal-overlay';
      document.body.appendChild(modal);
    }

    modal.innerHTML = `
      <div class="modal-container">
        <!-- Modal Header kèm Breadcrumb -->
        <div class="modal-header">
          <div>
            ${
              !isMain && parentContract
                ? `
              <div class="modal-breadcrumb">
                <span class="breadcrumb-link" id="btn-back-to-parent" data-parent-id="${parentContract.id}">
                  <i class="fa-solid fa-arrow-left"></i> ${parentContract.contract_no || parentContract.id} (${parentContract.parent_id ? 'Phụ lục cha' : 'HĐ gốc'})
                </span>
                <span>/</span>
                <span><strong>${contract.contract_no}</strong></span>
              </div>
            `
                : ''
            }
            <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
              <h2 style="font-size: 1.25rem; font-weight: 800;">
                ${isMain ? '' : '<span style="color: var(--primary);">[' + (directChildren.length > 0 ? 'Phụ lục cha' : 'Phụ lục / Gói con') + ']</span> '}
                ${contract.contract_no || contract.id}
              </h2>
              <span class="badge ${contract.group === 'A' ? 'badge-primary' : 'badge-warning'}">Nhóm ${contract.group}</span>
              <span class="badge badge-success">${contract.status || 'Active'}</span>
            </div>
            <p style="font-size: 0.85rem; color: var(--text-muted); margin-top: 2px;">${contract.scope_of_work || ''}</p>
          </div>
          <div style="display: flex; gap: 8px;">
            <button id="btn-clone-current-contract" class="button button-outline" style="font-size: 0.8rem; padding: 6px 12px;" title="Sao chép hợp đồng này">
              <i class="fa-solid fa-copy" style="color: #6366f1;"></i> Sao chép
            </button>
            <button id="btn-edit-contract-info" class="button button-outline" style="font-size: 0.8rem; padding: 6px 12px;">
              <i class="fa-solid fa-pen-to-square"></i> Sửa thông tin
            </button>
            <button id="btn-close-modal" class="button button-outline" style="padding: 6px 12px;"><i class="fa-solid fa-xmark"></i></button>
          </div>
        </div>

        <!-- 4 TABS NGHIỆP VỤ -->
        <div style="padding: 0 20px; background: #fff; border-bottom: 1px solid var(--border-color);">
          <div class="tab-nav" style="margin-bottom: 0;">
            <button class="tab-btn active modal-tab-btn" data-target="m-tab-info">
              <i class="fa-solid fa-circle-info"></i> 1. Thông tin ${isMain ? 'Hợp đồng' : 'Phụ lục'}
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-docs">
              <i class="fa-solid fa-folder-closed"></i> 2. Hồ sơ (${docs.length})
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-ipc">
              <i class="fa-solid fa-money-bill-wave"></i> 3. Thanh toán (${realPayments.length} đợt)
            </button>
            <button class="tab-btn modal-tab-btn" data-target="m-tab-dash">
              <i class="fa-solid fa-chart-pie"></i> 4. Dashboard & Báo cáo
            </button>
          </div>
        </div>

        <div class="modal-body">
          <!-- PHÂN KHU 1: THÔNG TIN -->
          <div id="m-tab-info" class="modal-tab-pane">
            <div class="card" style="margin-bottom: 20px;">
              <h3 class="card-title"><i class="fa-solid fa-file-contract"></i> Thông tin ${isMain ? 'Hợp đồng gốc' : 'Phụ lục'}</h3>
              <div class="form-grid-2">
                <div><strong>Mã nội bộ:</strong> <p style="color: var(--text-muted);">${contract.id}</p></div>
                <div><strong>Số hiệu chính thức:</strong> <p style="font-weight: 700; color: var(--primary);">${contract.contract_no}</p></div>
                <div><strong>Bên A (Chủ đầu tư / Tổng thầu):</strong> <p style="color: var(--text-muted);">${contract.company_a || '—'}</p></div>
                <div><strong>Bên B (Thầu phụ / Nhà cung cấp):</strong> <p style="color: var(--text-muted);">${contract.company_b || '—'}</p></div>
                <div><strong>Ngày ký kết:</strong> <p style="color: var(--text-muted);">${this.formatDate(contract.signing_date)}</p></div>
                <div><strong>Giá trị ký kết HĐ gốc:</strong> <p style="font-weight: 700; color: var(--primary);">${this.formatCurrency(contract.contract_val)}</p></div>
                ${
                  isMain && directChildren.length > 0
                    ? `<div><strong>Tổng giá trị các Phụ lục đã giao:</strong> <p style="font-weight: 700; color: #0284c7;">${this.formatCurrency(stats.appendixSum)}</p></div>`
                    : (!isMain ? `<div><strong>Thuộc HĐ / Phụ lục cha:</strong> <p style="color: var(--primary); font-weight: 600;">${parentContract?.contract_no || contract.parent_id}</p></div>` : '')
                }
                <div><strong>Tài khoản ngân hàng:</strong> <p style="color: var(--text-muted);">${contract.bank_account || '—'} - ${contract.bank_name || ''}</p></div>
              </div>
            </div>

            <!-- Nếu có phụ lục / gói con trực tiếp thì hiển thị bảng các con -->
            ${
              directChildren.length > 0
                ? `
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px;">
                <h3 style="font-size: 1rem; font-weight: 700;">
                  <i class="fa-solid fa-layer-group"></i> ${isMain ? 'Danh sách Phụ lục trực tiếp' : 'Các Gói con trực thuộc'} (${directChildren.length})
                  <small style="font-weight: normal; color: var(--text-muted); font-size: 0.8rem;">(Bấm vào bất kỳ dòng nào để vào riêng gói đó)</small>
                </h3>
                <button id="btn-add-appendix" class="button button-outline" style="font-size: 0.8rem; padding: 4px 10px;">
                  <i class="fa-solid fa-plus"></i> ${isMain ? 'Thêm Phụ lục mới' : 'Thêm Gói con mới'}
                </button>
              </div>

              <div class="data-table-wrap">
                <table class="data-table">
                  <thead>
                    <tr>
                      <th style="width: 120px;">Số hiệu</th>
                      <th>Nội dung / Phạm vi công việc</th>
                      <th style="width: 110px; text-align: center;">Tình trạng</th>
                      <th class="num-col" style="width: 160px;">Giá trị ký kết</th>
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
                          <td style="text-align: center;"><span class="badge ${a.status === 'Active' ? 'badge-success' : 'badge-warning'}">${a.status || 'Active'}</span></td>
                          <td class="num-col" style="color: var(--primary); font-weight: 700;">${this.formatCurrency(a.contract_val)}</td>
                          <td style="text-align: center;">${this.formatDate(a.signing_date)}</td>
                        </tr>
                      `
                      )
                      .join('')}
                  </tbody>
                </table>
              </div>
            `
                : ''
            }
          </div>

          <!-- PHÂN KHU 2: HỒ SƠ RIÊNG -->
          <div id="m-tab-docs" class="modal-tab-pane hidden">
            <div style="background: var(--bg-card-subtle); padding: 14px 18px; border-radius: var(--radius-md); margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
              <div>
                <i class="fa-brands fa-google-drive" style="color: #2563eb; font-size: 1.25rem;"></i>
                <strong style="margin-left: 8px;">Thư mục Lưu trữ trên Drive / Máy tính:</strong>
                <div style="font-family: monospace; font-size: 0.825rem; color: var(--text-muted); margin-top: 4px;">
                  ${this.getDrivePathForContract(contract)}
                </div>
              </div>
              <button id="btn-open-drive-folder" class="button button-primary" style="font-size: 0.8rem;">
                <i class="fa-solid fa-arrow-up-right-from-square"></i> Mở Thư mục Drive
              </button>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
              <h3 style="font-size: 1rem; font-weight: 700;">
                <i class="fa-solid fa-list-check"></i> Danh mục Hồ sơ & Chứng từ (${docs.length})
              </h3>
              <button id="btn-add-document" class="button button-primary" style="font-size: 0.8rem; padding: 6px 12px;">
                <i class="fa-solid fa-plus"></i> Thêm Hồ sơ mới
              </button>
            </div>

            <div class="data-table-wrap">
              <table class="data-table">
                <thead>
                  <tr>
                    <th>Mã hồ sơ</th>
                    <th>Số công văn / Hồ sơ</th>
                    <th>Phân loại</th>
                    <th>Tiêu đề tài liệu</th>
                    <th>Ngày nhận</th>
                    <th>Tình trạng</th>
                    <th>File scan</th>
                    <th>Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  ${
                    docs.length === 0
                      ? '<tr><td colspan="8" style="text-align:center; padding: 20px;">Chưa có hồ sơ nào. Bấm "+ Thêm Hồ sơ mới" để tạo.</td></tr>'
                      : docs
                          .map(
                            (d) => `
                        <tr>
                          <td><strong>${d.id}</strong></td>
                          <td>${d.doc_no || '—'}</td>
                          <td><span class="badge badge-primary">${d.category || 'Hồ sơ'}</span></td>
                          <td>${d.title || '—'}</td>
                          <td>${this.formatDate(d.received_date)}</td>
                          <td><span class="badge badge-success">${d.status || 'Approved'}</span></td>
                          <td>
                            ${
                              d.file_link
                                ? `<a href="${d.file_link}" target="_blank" class="button button-outline" style="padding: 2px 8px; font-size: 0.75rem;"><i class="fa-solid fa-file-pdf" style="color: var(--danger);"></i> Xem</a>`
                                : '<span style="color: var(--text-light); font-size: 0.8rem;">Chưa có file</span>'
                            }
                          </td>
                          <td>
                            <button class="button button-outline btn-delete-doc" data-doc-id="${d.id}" style="padding: 2px 6px; font-size: 0.75rem; color: var(--danger);" title="Xóa hồ sơ">
                              <i class="fa-solid fa-trash"></i>
                            </button>
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

          <!-- PHÂN KHU 3: THANH TOÁN (PAYMENT LEDGER / IPC) -->
          <div id="m-tab-ipc" class="modal-tab-pane hidden">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; flex-wrap: wrap; gap: 10px;">
              <div>
                <span style="font-size: 0.9rem; color: var(--text-muted);">
                  Tổng đã giải ngân thực tế: <strong style="color: var(--success); font-size: 1.05rem;">${this.formatCurrency(stats.totalPaid)}</strong> / Đề nghị: <strong>${this.formatCurrency(stats.totalRequested)}</strong>
                </span>
              </div>
              <button id="btn-add-payment-entry" class="button button-primary" style="font-size: 0.8rem; padding: 6px 14px;">
                <i class="fa-solid fa-plus"></i> Nhập đợt Thanh toán (Payment Entry)
              </button>
            </div>

            <div class="data-table-wrap">
              <table class="data-table">
                <thead>
                  <tr>
                    <th>Đợt (IPC)</th>
                    <th>Nội dung thanh toán</th>
                    <th>Ngày thanh toán</th>
                    <th class="num-col">Nghiệm thu</th>
                    <th class="num-col">Đề nghị</th>
                    <th class="num-col">Thực trả</th>
                    <th class="num-col">Chênh lệch</th>
                    <th>Hóa đơn VAT</th>
                    <th>Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  ${
                    realPayments.length === 0
                      ? '<tr><td colspan="9" style="text-align:center; padding: 20px;">Chưa có đợt thanh toán nào. Bấm "+ Nhập đợt Thanh toán" để tạo.</td></tr>'
                      : realPayments
                          .map(
                            (p) => `
                        <tr>
                          <td><strong>IPC-${p.payment_no || p.item_index || '01'}</strong></td>
                          <td>${p.item_name || 'Nghiệm thu giai đoạn'}</td>
                          <td>${this.formatDate(p.payment_date)}</td>
                          <td class="num-col">${this.formatCurrency(p.accepted_vol)}</td>
                          <td class="num-col">${this.formatCurrency(p.request_amount)}</td>
                          <td class="num-col" style="color: var(--success);">${this.formatCurrency(p.paid_amount)}</td>
                          <td class="num-col" style="color: ${p.remaining > 0 ? 'var(--danger)' : 'var(--text-main)'};">
                            ${this.formatCurrency(p.remaining)}
                          </td>
                          <td>
                            ${p.invoice_ref ? `<span class="badge badge-primary">HĐ: ${p.invoice_ref}</span>` : ''}
                            <small style="color: var(--text-muted); display: block;">${p.note || ''}</small>
                          </td>
                          <td>
                            <button class="button button-outline btn-delete-payment" data-pay-id="${p.id}" style="padding: 2px 6px; font-size: 0.75rem; color: var(--danger);" title="Xóa đợt thanh toán">
                              <i class="fa-solid fa-trash"></i>
                            </button>
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

          <!-- PHÂN KHU 4: DASHBOARD TIẾN ĐỘ & XUẤT BÁO CÁO -->
          <div id="m-tab-dash" class="modal-tab-pane hidden">
            <div class="card" style="margin-bottom: 20px;">
              <h3 class="card-title"><i class="fa-solid fa-chart-line"></i> Tiến độ Thực hiện & Giải ngân</h3>
              <div>
                <div style="display: flex; justify-content: space-between; font-size: 0.9rem; font-weight: 600;">
                  <span>Tỷ lệ Giải ngân:</span>
                  <span style="color: var(--primary);">${stats.paymentRate}%</span>
                </div>
                <div class="progress-bar-bg">
                  <div class="progress-bar-fill" style="width: ${stats.paymentRate}%;"></div>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 0.8rem; color: var(--text-muted);">
                  <span>Đã giải ngân: ${this.formatCurrency(stats.totalPaid)}</span>
                  <span>Tổng giá trị ký kết: ${this.formatCurrency(stats.totalSigned)}</span>
                </div>
              </div>

              <div class="grid-4" style="margin-top: 16px;">
                <div class="stat-box">
                  <div class="stat-label">Tổng nghiệm thu</div>
                  <div style="font-weight: 700; font-size: 1.05rem;">${this.formatCurrency(stats.totalAccepted)}</div>
                </div>
                <div class="stat-box">
                  <div class="stat-label">Tổng đề nghị</div>
                  <div style="font-weight: 700; font-size: 1.05rem;">${this.formatCurrency(stats.totalRequested)}</div>
                </div>
                <div class="stat-box">
                  <div class="stat-label">Công nợ tồn</div>
                  <div style="font-weight: 700; font-size: 1.05rem; color: ${stats.remaining > 0 ? 'var(--danger)' : 'var(--text-main)'};">
                    ${this.formatCurrency(stats.remaining)}
                  </div>
                </div>
                <div class="stat-box">
                  <div class="stat-label">Số đợt IPC thực tế</div>
                  <div style="font-weight: 700; font-size: 1.05rem;">${stats.ipcCount} đợt</div>
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

        <!-- Modal Footer -->
        <div class="modal-footer">
          <button id="btn-close-modal-bottom" class="button button-outline">Đóng</button>
        </div>
      </div>
    `;

    const closeModal = () => modal.classList.add('hidden');
    document.getElementById('btn-close-modal')?.addEventListener('click', closeModal);
    document.getElementById('btn-close-modal-bottom')?.addEventListener('click', closeModal);
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
    modalTabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        modalTabBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        const targetId = btn.getAttribute('data-target');
        modal.querySelectorAll('.modal-tab-pane').forEach((pane) => {
          if (pane.id === targetId) pane.classList.remove('hidden');
          else pane.classList.add('hidden');
        });
      });
    });

    // Mở Drive Folder
    document.getElementById('btn-open-drive-folder')?.addEventListener('click', () => {
      const drivePath = this.getDrivePathForContract(contract);
      alert(`Đường dẫn thư mục Google Drive:\n\n${drivePath}\n\n(Tự động đồng bộ với Google Drive Cloud & Máy tính)`);
    });

    // Sao chép Hợp đồng
    document.getElementById('btn-clone-current-contract')?.addEventListener('click', () => {
      this.openCloneContractModal(contract.id);
    });

    // Sửa thông tin
    document.getElementById('btn-edit-contract-info')?.addEventListener('click', () => {
      this.openContractFormModal(contract);
    });

    // Thêm phụ lục mới
    document.getElementById('btn-add-appendix')?.addEventListener('click', () => {
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

    // Xóa hồ sơ
    modal.querySelectorAll('.btn-delete-doc').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        const docId = btn.getAttribute('data-doc-id');
        if (confirm(`Bạn có chắc muốn xóa hồ sơ ${docId} không?`)) {
          await db.delete('documents', docId);
          await this.loadData();
          this.openContractDetail(contract.id);
        }
      });
    });

    // Xóa thanh toán
    modal.querySelectorAll('.btn-delete-payment').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        const payId = btn.getAttribute('data-pay-id');
        if (confirm(`Bạn có chắc muốn xóa đợt thanh toán này không?`)) {
          await db.delete('payments', payId);
          await this.loadData();
          this.openContractDetail(contract.id);
        }
      });
    });

    // Xuất báo cáo theo mẫu
    document.getElementById('btn-export-excel-report')?.addEventListener('click', () => {
      this.exportContractReport(contract, isMain);
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
            <i class="fa-solid fa-folder-plus" style="color: var(--primary);"></i> Thêm Dự án Mới hoặc Nhập dữ liệu CPMS
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>

        <div style="padding: 0 20px; background: #fff; border-bottom: 1px solid var(--border-color);">
          <div class="tab-nav" style="margin-bottom: 0;">
            <button class="tab-btn active prj-tab-btn" data-target="tab-import-cpms">
              <i class="fa-solid fa-file-import"></i> 1. Nhập dữ liệu CPMS (.cpms / .xlsx)
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
              Chọn file sao lưu dữ liệu dự án (có đuôi <code>.cpms</code>, <code>.xlsb</code>, hoặc <code>.xlsx</code> được xuất từ phần mềm CPMS Plus).
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

    btnBrowse.addEventListener('click', () => fileInput.click());
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
      this.currentProjectId = newId;
      await this.loadData();
      this.render();
      closeSubModal();
      alert(`Đã tạo thành công dự án "${newName}"!`);
    });
  }

  // ==========================================
  // FORM NHẬP / SỬA HỢP ĐỒNG GỐC HOẶC PHỤ LỤC
  // ==========================================
  openContractFormModal(editContract = null) {
    const isEdit = !!editContract;
    const isAppendix = editContract && !!editContract.parent_id;
    let currentSelectGroup = editContract ? editContract.group : this.currentGroup;

    // Sinh mã tự động chuẩn CPMS
    const formId = editContract
      ? editContract.id
      : CPMSImporterExporter.generateNextContractID(null, this.contracts, currentSelectGroup);

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 700px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700;">
            ${isEdit ? (isAppendix ? 'Chỉnh sửa Phụ lục' : 'Chỉnh sửa Hợp đồng') : 'Thêm Hợp đồng gốc mới'}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-contract">
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Mã định danh (ID chuẩn CPMS)</label>
                <input type="text" class="form-control" id="contract-input-id" name="id" value="${formId}" ${isEdit ? 'readonly' : ''} required>
                <small style="color: var(--text-muted); font-size: 0.75rem;">Quy tắc: Nhóm-Năm-SốThứTự (Vd: A-2026-01)</small>
              </div>
              <div class="form-group">
                <label class="form-label">Phân nhóm</label>
                <select class="form-control" id="contract-select-group" name="group" ${isAppendix ? 'disabled' : ''}>
                  <option value="A" ${currentSelectGroup === 'A' ? 'selected' : ''}>Nhóm A: Chủ đầu tư / Tổng thầu</option>
                  <option value="B" ${currentSelectGroup === 'B' ? 'selected' : ''}>Nhóm B: Thầu phụ / Nhà cung cấp</option>
                </select>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Số hiệu chính thức</label>
              <input type="text" class="form-control" name="contract_no" value="${editContract?.contract_no || ''}" placeholder="Vd: 68/2022/HĐKT/CHC1/VEC" required>
            </div>

            <div class="form-group">
              <label class="form-label">Phạm vi công việc / Nội dung gói thầu</label>
              <textarea class="form-control" name="scope_of_work" rows="2" placeholder="Tên gói thầu, phạm vi thiết kế hoặc thi công...">${editContract?.scope_of_work || ''}</textarea>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Bên A (Chủ đầu tư / Tổng thầu)</label>
                <input type="text" class="form-control" name="company_a" value="${editContract?.company_a || ''}" placeholder="Tên cơ quan Bên A">
              </div>
              <div class="form-group">
                <label class="form-label">Bên B (Nhà thầu / Đối tác)</label>
                <input type="text" class="form-control" name="company_b" value="${editContract?.company_b || ''}" placeholder="Tên nhà thầu Bên B">
              </div>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Giá trị ký kết (VNĐ)</label>
                <input type="number" class="form-control" name="contract_val" value="${editContract?.contract_val || 0}" required>
              </div>
              <div class="form-group">
                <label class="form-label">Ngày ký</label>
                <input type="date" class="form-control" name="signing_date" value="${this.formatDateForInput(editContract?.signing_date)}">
              </div>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary">Lưu Hợp đồng</button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    // Nếu tạo mới và thay đổi nhóm A/B thì tự sinh lại ContractID tương ứng
    if (!isEdit) {
      const grpSelect = modal.querySelector('#contract-select-group');
      const idInput = modal.querySelector('#contract-input-id');
      grpSelect?.addEventListener('change', (e) => {
        const nextId = CPMSImporterExporter.generateNextContractID(null, this.contracts, e.target.value);
        idInput.value = nextId;
      });
    }

    modal.querySelector('#form-contract').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const updatedData = {
        ...(editContract || {}),
        id: formData.get('id'),
        project_id: editContract?.project_id || this.currentProjectId,
        group: editContract?.group || formData.get('group'),
        level: 0,
        parent_id: null,
        contract_no: formData.get('contract_no'),
        contract_type: 'MAIN',
        scope_of_work: formData.get('scope_of_work'),
        company_a: formData.get('company_a'),
        company_b: formData.get('company_b'),
        contract_val: parseFloat(formData.get('contract_val') || 0),
        signing_date: formData.get('signing_date'),
        status: editContract?.status || 'Active',
      };

      await db.put('contracts', updatedData);
      await this.loadData();
      this.render();
      closeSubModal();
      alert('Đã lưu thành công!');
      this.openContractDetail(updatedData.id);
    });
  }

  // ==========================================
  // FORM THÊM PHỤ LỤC MỚI HOẶC GÓI CON CẤP 2
  // ==========================================
  openAppendixFormModal(parentContract) {
    const isLevel1 = !parentContract.parent_id;
    const levelTitle = isLevel1 ? 'Phụ lục cấp 1' : 'Gói con cấp 2';
    const newAppId = CPMSImporterExporter.generateNextContractID(parentContract, this.contracts, parentContract.group);

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 600px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700;">
            Thêm ${levelTitle} cho ${parentContract.contract_no}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-appendix">
            <div class="form-group">
              <label class="form-label">Mã định danh (ID chuẩn CPMS phân cấp)</label>
              <input type="text" class="form-control" name="id" value="${newAppId}" readonly>
              <small style="color: var(--text-muted); font-size: 0.75rem;">Định dạng chuẩn cha-con: ${parentContract.id}-XX</small>
            </div>
            <div class="form-group">
              <label class="form-label">Số hiệu ${levelTitle}</label>
              <input type="text" class="form-control" name="contract_no" placeholder="Vd: ${isLevel1 ? 'APP-01 hoặc PL-01' : 'PKG-01'}" required>
            </div>
            <div class="form-group">
              <label class="form-label">Nội dung công việc / Phạm vi</label>
              <textarea class="form-control" name="scope_of_work" rows="2" placeholder="Hạng mục thi công bổ sung, gia hạn thời gian..." required></textarea>
            </div>
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Giá trị ký kết (VNĐ)</label>
                <input type="number" class="form-control" name="contract_val" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label">Ngày ký</label>
                <input type="date" class="form-control" name="signing_date" value="${new Date().toISOString().slice(0, 10)}" required>
              </div>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary">Lưu ${levelTitle}</button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    modal.querySelector('#form-appendix').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const appendixData = {
        id: formData.get('id'),
        project_id: parentContract.project_id || this.currentProjectId,
        parent_id: parentContract.id,
        group: parentContract.group,
        level: (parentContract.level || 0) + 1,
        contract_no: formData.get('contract_no'),
        contract_type: isLevel1 ? 'Appendix' : 'Sub-Package',
        scope_of_work: formData.get('scope_of_work'),
        contract_val: parseFloat(formData.get('contract_val') || 0),
        signing_date: formData.get('signing_date'),
        company_a: parentContract.company_a,
        company_b: parentContract.company_b,
        status: 'Active',
      };

      await db.put('contracts', appendixData);
      await this.loadData();
      this.render();
      closeSubModal();
      this.openContractDetail(parentContract.id);
      alert(`Đã thêm thành công ${levelTitle}!`);
    });
  }

  // ==========================================
  // FORM THÊM HỒ SƠ
  // ==========================================
  openDocumentFormModal(targetContract) {
    const docs = this.getDocuments(targetContract.id);
    const docId = `DOC-${new Date().getFullYear()}-${String(docs.length + 1).padStart(3, '0')}`;

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 600px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700;">Thêm Hồ sơ cho ${targetContract.contract_no}</h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-doc">
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Mã Hồ sơ</label>
                <input type="text" class="form-control" name="id" value="${docId}" readonly>
              </div>
              <div class="form-group">
                <label class="form-label">Phân loại hồ sơ</label>
                <select class="form-control" name="category">
                  <option value="Hồ sơ nghiệm thu">Hồ sơ nghiệm thu (WIR/MIR)</option>
                  <option value="Hồ sơ thanh toán">Hồ sơ thanh toán (IPC)</option>
                  <option value="Hợp đồng & Pháp lý">Hợp đồng & Pháp lý (CON/APP)</option>
                  <option value="Công văn đến / đi">Công văn đến / đi (LTR)</option>
                  <option value="Bản vẽ thiết kế">Bản vẽ thiết kế (DWG/PDF)</option>
                </select>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Số hiệu công văn / Tên hồ sơ</label>
              <input type="text" class="form-control" name="doc_no" placeholder="Vd: 125/CV-CHC1 hoặc BB-NT-01" required>
            </div>

            <div class="form-group">
              <label class="form-label">Tiêu đề / Nội dung tài liệu</label>
              <textarea class="form-control" name="title" rows="2" placeholder="Mô tả tóm tắt nội dung hồ sơ..." required></textarea>
            </div>

            <div class="form-group">
              <label class="form-label">Đường dẫn tệp tin / Link Google Drive</label>
              <input type="url" class="form-control" name="file_link" placeholder="https://drive.google.com/file/d/...">
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Ngày tiếp nhận</label>
                <input type="date" class="form-control" name="received_date" value="${new Date().toISOString().slice(0, 10)}">
              </div>
              <div class="form-group">
                <label class="form-label">Tình trạng</label>
                <select class="form-control" name="status">
                  <option value="Approved" selected>Đã duyệt (Approved)</option>
                  <option value="UnderReview">Đang xem xét (UnderReview)</option>
                  <option value="Draft">Dự thảo (Draft)</option>
                </select>
              </div>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary">Lưu Hồ sơ</button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    modal.querySelector('#form-doc').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const docData = {
        id: formData.get('id'),
        project_id: targetContract.project_id || this.currentProjectId,
        contract_id: targetContract.id,
        category: formData.get('category'),
        doc_no: formData.get('doc_no'),
        title: formData.get('title'),
        file_link: formData.get('file_link'),
        received_date: formData.get('received_date'),
        status: formData.get('status'),
      };

      await db.put('documents', docData);
      await this.loadData();
      closeSubModal();
      this.openContractDetail(targetContract.id);
      alert('Đã lưu thành công Hồ sơ!');
    });
  }

  // ==========================================
  // FORM NHẬP THANH TOÁN (PAYMENT ENTRY)
  // ==========================================
  openPaymentEntryModal(targetContract) {
    const realPayments = this.getRealPayments(targetContract.id);
    const nextIdx = realPayments.length + 1;
    const ledgerId = `LG-${Date.now()}`;

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 650px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700;">Nhập Thanh toán cho ${targetContract.contract_no}</h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-payment">
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Kỳ / Đợt thanh toán (IPC No)</label>
                <input type="text" class="form-control" name="payment_no" value="${String(nextIdx).padStart(2, '0')}" required>
              </div>
              <div class="form-group">
                <label class="form-label">Ngày thanh toán</label>
                <input type="date" class="form-control" name="payment_date" value="${new Date().toISOString().slice(0, 10)}" required>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Diễn giải nội dung đợt thanh toán</label>
              <input type="text" class="form-control" name="item_name" placeholder="Vd: Tạm ứng đợt 1 / Nghiệm thu giai đoạn 2" required>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Khối lượng Nghiệm thu (VNĐ)</label>
                <input type="number" class="form-control" id="pay-accepted" name="accepted_vol" value="0">
              </div>
              <div class="form-group">
                <label class="form-label">Bên B Đề nghị thanh toán (VNĐ)</label>
                <input type="number" class="form-control" id="pay-requested" name="request_amount" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label">Bên A Thực tế đã thanh toán (VNĐ)</label>
                <input type="number" class="form-control" id="pay-paid" name="paid_amount" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label">Số Hóa đơn VAT / Tham chiếu</label>
                <input type="text" class="form-control" name="invoice_ref" placeholder="Vd: 0012458">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú thanh toán</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú đối trừ tạm ứng, số UNC..."></textarea>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary">Ghi sổ Thanh toán</button>
            </div>
          </form>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    const closeSubModal = () => modal.remove();
    modal.querySelectorAll('.btn-close-submodal').forEach((b) => b.addEventListener('click', closeSubModal));

    modal.querySelector('#form-payment').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const reqAmt = parseFloat(formData.get('request_amount') || 0);
      const paidAmt = parseFloat(formData.get('paid_amount') || 0);

      const paymentData = {
        id: ledgerId,
        is_header: false, // Dòng giao dịch thực tế
        project_id: targetContract.project_id || this.currentProjectId,
        contract_id: targetContract.id,
        parent_id: targetContract.parent_id || targetContract.id,
        payment_no: formData.get('payment_no'),
        item_index: nextIdx,
        item_name: formData.get('item_name'),
        payment_date: formData.get('payment_date'),
        accepted_vol: parseFloat(formData.get('accepted_vol') || 0),
        request_amount: reqAmt,
        paid_amount: paidAmt,
        remaining: reqAmt - paidAmt,
        invoice_ref: formData.get('invoice_ref'),
        note: formData.get('note'),
      };

      await db.put('payments', paymentData);
      await this.loadData();
      closeSubModal();
      this.openContractDetail(targetContract.id);
      alert('Đã ghi sổ thành công đợt thanh toán IPC!');
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
    csvContent += `BÁO CÁO QUẢN LÝ ${isMain ? 'HỢP ĐỒNG GỐC' : 'PHỤ LỤC'} & THANH TOÁN (CPMS)\r\n`;
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

    csvContent += `--- BẢNG KÊ CÁC ĐỢT THANH TOÁN (IPC THỰC TẾ) ---\r\n`;
    csvContent += `Đợt (IPC),Nội dung thanh toán,Ngày thanh toán,Giá trị nghiệm thu,Đề nghị thanh toán,Thực trả,Chênh lệch,Hóa đơn VAT,Ghi chú\r\n`;
    realPayments.forEach((p) => {
      csvContent += `"IPC-${p.payment_no || p.item_index}","${(p.item_name || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${p.accepted_vol},${p.request_amount},${p.paid_amount},${p.remaining},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
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
}

export const contractsManager = new ContractsManager();
