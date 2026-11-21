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

import { db } from '../../core/db.js?v=5.0';
import { CPMSImporterExporter } from '../../core/cpms-importer.js?v=5.0';

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
   * Tự động lưu mọi thay đổi trực tiếp vào file máy tính đã liên kết
   */
  async syncToLocalFile(isManual = false) {
    if (!this.activeFileHandle) {
      if (isManual) {
        alert('Chưa có liên kết tệp trực tiếp trên máy tính. Hãy bấm "Nhập dữ liệu CPMS (Dẫn file Local)" để chọn tệp trên máy tính trước.');
      }
      return false;
    }

    try {
      const buffer = await CPMSImporterExporter.buildCPMSBinaryBuffer(this.currentProjectId);
      const writable = await this.activeFileHandle.createWritable();
      await writable.write(buffer);
      await writable.close();

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
          const result = await CPMSImporterExporter.importCPMSFile(file);
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
    const projectInfoBar = document.getElementById('project-info-bar');

    if (!ready) {
      // Ẩn toàn bộ 3 tab nội dung
      overviewPane?.classList.add('hidden');
      contractsPane?.classList.add('hidden');
      disbursementsPane?.classList.add('hidden');
      if (projectInfoBar) projectInfoBar.style.display = 'none';

      // Hiển thị khung Empty State
      emptyContainer?.classList.remove('hidden');

      // Ẩn các nút thao tác dự án không cần thiết (sao chép, xuất file, xóa), chỉ giữ icon dẫn file local
      const btnClone = document.getElementById('btn-clone-project');
      const btnExport = document.getElementById('btn-export-cpms');
      const btnDelete = document.getElementById('btn-delete-project');
      if (btnClone) btnClone.style.display = 'none';
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

      // Hiện lại các nút thao tác dự án
      const btnClone = document.getElementById('btn-clone-project');
      const btnExport = document.getElementById('btn-export-cpms');
      const btnDelete = document.getElementById('btn-delete-project');
      if (btnClone) btnClone.style.display = '';
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
    this.payments = allPayments.filter(
      (p) => p.project_id === this.currentProjectId
    );
    this.documents = allDocs.filter(
      (d) => d.project_id === this.currentProjectId
    );
  }

  getBaseName(path) {
    if (!path) return '';
    const clean = String(path).replace(/[/\\]+$/, '');
    const parts = clean.split(/[/\\]/);
    return parts[parts.length - 1] || clean;
  }

  updateProjectSelectorUI() {
    const selector = document.getElementById('project-selector');
    if (!selector) return;

    if (!this.projects || this.projects.length === 0) {
      this.currentProjectId = null;
      selector.innerHTML = `
        <option value="" disabled selected>-- Chưa có dự án nào --</option>
        <option value="IMPORT_NEW">+ Dẫn link file CPMS từ máy tính (.cpms)...</option>
      `;
      const activePrjName = document.getElementById('active-project-name');
      if (activePrjName) activePrjName.textContent = 'Chưa có dự án nào';
      const topProjName = document.getElementById('top-project-name');
      if (topProjName) topProjName.textContent = 'Chưa có dự án nào';
      const topProjLoc = document.getElementById('top-project-location');
      if (topProjLoc) topProjLoc.textContent = '—';
      return;
    }

    let optionsHtml = '';
    // Đảm bảo currentProjectId hợp lệ
    if (!this.projects.some((p) => p.id === this.currentProjectId)) {
      this.currentProjectId = this.projects[0].id;
    }

    this.projects.forEach((p) => {
      const isSel = p.id === this.currentProjectId ? 'selected' : '';
      optionsHtml += `<option value="${p.id}" ${isSel}>📁 ${p.name || p.id}</option>`;
    });

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
    searchInput?.addEventListener('input', (e) => {
      this.searchQuery = e.target.value.trim();
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
      if (val === 'IMPORT_NEW') {
        // Trả lại giá trị trước đó trên UI
        projectSelector.value = this.currentProjectId;
        this.pickAndLoadLocalCPMSFile();
      } else {
        this.currentProjectId = val;
        await this.loadData();
        this.render();
      }
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

      const confirmed = confirm(
        `⚠️ CẢNH BÁO XÓA DỰ ÁN!\n\nBạn có chắc chắn muốn xóa vĩnh viễn dự án "${prjName}" (Mã: ${this.currentProjectId}) không?\n\nHành động này sẽ xóa vĩnh viễn:\n- Toàn bộ danh mục Hợp đồng & Phụ lục của dự án\n- Toàn bộ Sổ theo dõi thanh toán (IPC)\n- Toàn bộ Hồ sơ & Công văn liên quan\n\nDữ liệu sau khi xóa sẽ không thể phục hồi!`
      );

      if (!confirmed) return;

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
        // Nếu không còn dự án nào, đặt rỗng hoàn toàn
        this.currentProjectId = null;
      }

      await this.loadData();
      this.render();
      this.updateEmptyStateUI();
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

  formatNumberWithDots(val) {
    if (val === undefined || val === null || val === '') return '0';
    if (typeof val === 'number') {
      return new Intl.NumberFormat('vi-VN').format(Math.round(val));
    }
    const cleanDigits = String(val).replace(/\D/g, '');
    if (!cleanDigits) return '0';
    try {
      return new Intl.NumberFormat('vi-VN').format(BigInt(cleanDigits));
    } catch (e) {
      return cleanDigits;
    }
  }

  parseNumberFromDots(val) {
    if (!val && val !== 0) return 0;
    const cleanStr = String(val).replace(/\./g, '').replace(/,/g, '').replace(/[^\d.-]/g, '').trim();
    const num = parseFloat(cleanStr);
    return isNaN(num) ? 0 : num;
  }

  setupCurrencyInputListeners(container = document) {
    container.querySelectorAll('.currency-input').forEach((input) => {
      input.type = 'text';
      input.setAttribute('inputmode', 'numeric');

      if (input.value !== undefined && input.value !== '') {
        input.value = this.formatNumberWithDots(input.value);
      }

      input.addEventListener('input', () => {
        const cursorPosition = input.selectionStart || 0;
        const originalLength = input.value.length;
        const rawDigits = input.value.replace(/\D/g, '');

        if (!rawDigits) {
          input.value = '';
          return;
        }

        let formatted = '';
        try {
          formatted = new Intl.NumberFormat('vi-VN').format(BigInt(rawDigits));
        } catch (e) {
          formatted = rawDigits;
        }
        input.value = formatted;

        const newLength = formatted.length;
        const newCursorPosition = Math.max(0, cursorPosition + (newLength - originalLength));
        try {
          input.setSelectionRange(newCursorPosition, newCursorPosition);
        } catch (err) {}
      });

      input.addEventListener('blur', () => {
        if (!input.value.trim()) {
          input.value = '0';
        } else {
          input.value = this.formatNumberWithDots(input.value);
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

  getDriveDocPathForContract(contract) {
    const curPrjId = this.currentProjectId || 'CHC1';
    const isMain = !contract.parent_id;
    const parent = !isMain ? this.contracts.find((c) => c.id === contract.parent_id) : null;
    const cleanNo = (no) => String(no || '').replace(/[\/\\?%*:|"<>]/g, '_');

    const groupFolder = contract.group === 'A' ? '01_Nhom_A_ChuDauTu' : '02_Nhom_B_ThauPhu_NCC';

    if (contract.group === 'A') {
      const mainNo = isMain ? cleanNo(contract.contract_no) : (parent ? cleanNo(parent.contract_no) : cleanNo(contract.parent_id));
      return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${mainNo} / Doc /`;
    } else {
      const validCategories = [
        'Construction', 'Procurement', 'Consultancy', 'Design',
        'Equipment', 'Service', 'Maintenance', 'Rental'
      ];
      let cat = contract.category || 'Construction';
      const matched = validCategories.find((c) => c.toLowerCase() === cat.toLowerCase());
      cat = matched || 'Construction';

      if (isMain) {
        const partnerName = cleanNo(contract.company_b || 'NCC');
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${cleanNo(contract.contract_no)}_${partnerName} / Doc /`;
      } else {
        const rootMain = parent && parent.parent_id ? this.contracts.find((c) => c.id === parent.parent_id) || parent : parent;
        const parentNo = rootMain ? cleanNo(rootMain.contract_no) : cleanNo(contract.parent_id);
        return `📁 MAIO_Data / 01_QuanLy_DuAn / DuAn_${curPrjId} / ${groupFolder} / ${cat} / ${parentNo} / Doc /`;
      }
    }
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

    // Cập nhật text trong dropdown dự án nếu cần
    const selOpt = document.querySelector(`#project-selector option[value="${curPrj.id}"]`);
    if (selOpt) {
      selOpt.textContent = `📁 ${curPrj.name || curPrj.id}`;
    }

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
      await this.syncToLocalFile();
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

  getMainContracts() {
    return this.contracts.filter((c) => {
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
    this.updateEmptyStateUI('contracts');
    if (!this.isProjectReady()) return;

    this.renderProjectInfo();
    this.updateCategoryFilterUI();

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
      .map((c) => {
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
                      <i class="fa-solid fa-receipt" style="color: #d97706; font-size: 1.15rem;"></i>
                      <span style="font-size: 1.05rem; font-weight: 800; color: #92400e;">${c.contract_no || c.id}</span>
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

        return `
          <div class="contract-group-wrapper" data-contract-id="${c.id}">
            <!-- Card Hợp đồng -->
            <div class="contract-card-main" data-open-detail="${c.id}">
              <div class="contract-header">
                <div class="contract-header-left">
                  <div class="contract-no">
                    <i class="fa-solid fa-file-signature"></i>
                    <span>${c.contract_no || c.id}</span>
                    <span class="badge ${c.group === 'A' ? 'badge-primary' : 'badge-warning'}">Nhóm ${c.group}</span>
                    ${c.category ? `<span class="badge" style="background: #eef2ff; color: #4338ca; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${c.category}</span>` : ''}
                    <span class="badge badge-success">${c.status || 'Active'}</span>
                  </div>
                  <div class="contract-title">${c.scope_of_work || 'Hợp đồng thực hiện dự án'}</div>
                </div>

                <div class="contract-header-right">
                  <button class="button button-outline btn-clone-contract-action" data-clone-cid="${c.id}" title="Sao chép Hợp đồng này">
                    <i class="fa-solid fa-copy" style="color: #6366f1;"></i>
                  </button>
                  <div class="contract-val-block">
                    <span class="contract-val-label">Giá trị Hợp đồng</span>
                    <div class="contract-val-amount">
                      ${this.formatCurrency(stats.totalSigned)}
                    </div>
                    ${directChildren.length > 0 ? `<div class="contract-val-sub">(Đã giao ${directChildren.length} PL: ${this.formatCurrency(stats.appendixSum)})</div>` : ''}
                  </div>
                </div>
              </div>

              <div class="contract-meta">
                <span><i class="fa-solid fa-handshake"></i> Bên A: <strong>${c.company_a || 'Chủ đầu tư'}</strong></span>
                <span>
                  <i class="fa-solid fa-building"></i> Bên B: <strong>${c.company_b || 'Đơn vị thực hiện'}</strong>
                  ${c.company_b2 ? ` & <strong>${c.company_b2}</strong> <span class="badge" style="font-size: 0.68rem; background: #e0e7ff; color: #4338ca; padding: 1px 5px; vertical-align: middle;">Liên danh</span>` : ''}
                  ${c.tax_code ? `<span class="badge" style="font-size: 0.72rem; background: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; margin-left: 6px; padding: 2px 6px;"><i class="fa-solid fa-barcode"></i> MST: ${c.tax_code}</span>` : ''}
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
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const payId = btn.getAttribute('data-pay-id');
        const payment = this.payments.find((p) => p.id === payId);
        if (!payment) return;

        const confirmed = confirm(
          `Bạn có chắc chắn muốn xóa đợt chi "${payment.item_name || payment.payment_no}" (${this.formatCurrency(payment.paid_amount || payment.request_amount)}) không?`
        );
        if (!confirmed) return;

        await db.delete('payments', payId);
        await this.loadData();
        this.render();
        await this.syncToLocalFile();
      });
    });

    // Bắt sự kiện Click vào HĐ gốc hoặc Phụ lục
    container.querySelectorAll('[data-open-detail]').forEach((el) => {
      el.addEventListener('click', (e) => {
        if (
          e.target.closest('[data-toggle-appendix]') ||
          e.target.closest('.btn-clone-contract-action') ||
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
                  <i class="fa-solid fa-arrow-left"></i> ${parentContract.contract_no || parentContract.id} (${parentContract.parent_id ? 'Phụ lục cha' : 'Hợp đồng cha'})
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
              ${contract.category ? `<span class="badge" style="background: #eef2ff; color: #4f46e5; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${contract.category}</span>` : ''}
              <span class="badge badge-success">${contract.status || 'Active'}</span>
            </div>
            <p style="font-size: 0.85rem; color: var(--text-muted); margin-top: 2px;">${contract.scope_of_work || ''}</p>
          </div>
          <div style="display: flex; gap: 8px;">
            <button id="btn-clone-current-contract" class="button button-outline" style="font-size: 0.85rem; padding: 6px 10px;" title="Sao chép hợp đồng này">
              <i class="fa-solid fa-copy" style="color: #6366f1;"></i>
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
              <i class="fa-solid fa-money-bill-wave"></i> 3. Giải ngân (${realPayments.length} đợt)
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
              <h3 class="card-title"><i class="fa-solid fa-file-contract"></i> Thông tin ${isMain ? 'Hợp đồng' : 'Phụ lục'}</h3>
              <div class="form-grid-2">
                <div><strong>Mã nội bộ (Contract ID):</strong> <p style="color: var(--text-muted); font-family: monospace; font-weight: 600;">${contract.id}</p></div>
                <div><strong>Số hiệu chính thức:</strong> <p style="font-weight: 700; color: var(--primary);">${contract.contract_no}</p></div>
                <div><strong>Loại hợp đồng:</strong> <p><span class="badge" style="background: #eef2ff; color: #4f46e5; border: 1px solid #c7d2fe; font-weight: 600;"><i class="fa-solid fa-tag"></i> ${contract.category || 'Chưa phân loại'}</span></p></div>
                <div><strong>Tình trạng:</strong> <p><span class="badge badge-success">${contract.status || 'Active'}</span></p></div>
                <div><strong>Ngày ký kết:</strong> <p style="color: var(--text-muted);">${this.formatDate(contract.signing_date)}</p></div>
                <div><strong>Thời hạn kết thúc:</strong> <p style="color: var(--text-muted);">${contract.end_date ? this.formatDate(contract.end_date) : '—'}</p></div>
                <div><strong>Giá trị ký kết Hợp đồng:</strong> <p style="font-weight: 700; color: var(--primary);">${this.formatCurrency(contract.contract_val)} <small style="font-weight: normal; color: var(--text-muted);">(Tỷ giá: ${contract.exchange_rate || 1})</small></p></div>
                ${
                  isMain && directChildren.length > 0
                    ? `<div><strong>Tổng giá trị các Phụ lục đã giao:</strong> <p style="font-weight: 700; color: #0284c7;">${this.formatCurrency(stats.appendixSum)}</p></div>`
                    : (!isMain ? `<div><strong>Thuộc HĐ / Phụ lục cha:</strong> <p style="color: var(--primary); font-weight: 600;">${parentContract?.contract_no || contract.parent_id}</p></div>` : '')
                }
                <div><strong>Bên A (Chủ đầu tư / Tổng thầu):</strong> <p style="color: var(--text-muted);">${contract.company_a || '—'} ${contract.rep_a ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_a})</span>` : ''}</p></div>
                <div><strong>Bên B (Nhà thầu B1):</strong> <p style="color: var(--text-muted);">${contract.company_b || '—'} ${contract.rep_b ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_b})</span>` : ''}</p></div>
                <div><strong>Mã số thuế Bên B (Tax Code):</strong> <p style="color: var(--text-muted); font-weight: 600;">${contract.tax_code || '—'}</p></div>
                ${
                  contract.company_b2
                    ? `<div><strong>Bên B2 (Nhà thầu Liên danh):</strong> <p style="color: var(--text-muted);">${contract.company_b2} ${contract.rep_b2 ? `<span style="color: #64748b; font-size: 0.8rem;">(Đại diện: ${contract.rep_b2})</span>` : ''}</p></div>`
                    : ''
                }
                ${
                  contract.contact_person_b
                    ? `<div><strong>Người liên hệ Bên B:</strong> <p style="color: var(--text-muted);">${contract.contact_person_b}</p></div>`
                    : ''
                }
                ${
                  contract.linked_contract_id
                    ? `<div><strong>Hợp đồng liên kết / tương tự:</strong> <p>
                        <button type="button" class="btn-jump-linked-contract" data-jump-id="${contract.linked_contract_id}" style="border: 1px solid #c7d2fe; background: #eef2ff; color: #4338ca; border-radius: 4px; padding: 3px 8px; font-size: 0.8rem; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 5px;">
                          <i class="fa-solid fa-link"></i> Đi đến HĐ: <strong>${(this.contracts.find((tc) => tc.id === contract.linked_contract_id)?.contract_no) || contract.linked_contract_id}</strong> <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.68rem;"></i>
                        </button>
                       </p></div>`
                    : ''
                }
                <div><strong>Đơn vị thụ hưởng (Payee):</strong> <p style="color: var(--text-muted);">${contract.payee || contract.company_b || '—'}</p></div>
                <div><strong>Tài khoản ngân hàng:</strong> <p style="color: var(--text-muted);">${contract.bank_account || '—'} ${contract.bank_name ? `(${contract.bank_name})` : ''}</p></div>
              </div>
            </div>

            <!-- Danh sách phụ lục / gói con trực tiếp -->
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px;">
              <h3 style="font-size: 1rem; font-weight: 700;">
                <i class="fa-solid fa-layer-group"></i> ${isMain ? 'Danh sách Phụ lục trực tiếp' : 'Các Gói con trực thuộc'} (${directChildren.length})
                <small style="font-weight: normal; color: var(--text-muted); font-size: 0.8rem;">(Bấm vào bất kỳ dòng nào để vào riêng gói đó)</small>
              </h3>
              <button id="btn-add-appendix" class="button button-primary" style="font-size: 0.8rem; padding: 4px 10px;">
                <i class="fa-solid fa-plus"></i> ${isMain ? 'Thêm Phụ lục mới' : 'Thêm Gói con mới'}
              </button>
            </div>

            ${
              directChildren.length > 0
                ? `
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
            <div style="background: var(--bg-card-subtle); padding: 14px 18px; border-radius: var(--radius-md); margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
              <div>
                <i class="fa-brands fa-google-drive" style="color: #2563eb; font-size: 1.25rem;"></i>
                <strong style="margin-left: 8px;">Thư mục Lưu trữ trên Drive / Máy tính:</strong>
                <div style="font-family: monospace; font-size: 0.8rem; color: var(--text-muted); margin-top: 6px; line-height: 1.5;">
                  <div><span style="color: var(--primary); font-weight: 600;">[HĐ & Phụ lục]:</span> ${this.getDrivePathForContract(contract)}</div>
                  <div style="margin-top: 3px;"><span style="color: #10b981; font-weight: 600;">[Doc - Tài liệu khác]:</span> ${this.getDriveDocPathForContract(contract)}</div>
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
                    <th>Tệp tin (Local & Drive Online)</th>
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
                            <div style="display: flex; flex-direction: column; gap: 5px;">
                              <!-- Link Local -->
                              <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem;">
                                <i class="fa-solid fa-laptop" style="color: #64748b;" title="Tệp trên máy tính (Local)"></i>
                                ${(d.local_link || (!String(d.file_link || '').startsWith('http') && d.file_link)) ? `
                                  <span style="max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-family: monospace; font-size: 0.75rem; color: #1e293b;" title="${d.local_link || d.file_link}">
                                    ${this.getBaseName(d.local_link || d.file_link)}
                                  </span>
                                  <button class="button button-outline btn-copy-doc-local" data-path="${encodeURIComponent(d.local_link || d.file_link)}" style="padding: 1px 5px; font-size: 0.7rem; border-color: #cbd5e1; height: 20px;" title="Sao chép đường dẫn file trên máy tính">
                                    <i class="fa-regular fa-copy"></i>
                                  </button>
                                ` : `
                                  <span style="color: #94a3b8; font-size: 0.75rem;">Chưa có file local</span>
                                `}
                              </div>
                              <!-- Link Drive Online -->
                              <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem;">
                                ${(d.drive_link || (String(d.file_link || '').startsWith('http') && d.file_link)) ? `
                                  <a href="${d.drive_link || d.file_link}" target="_blank" class="button button-outline" style="padding: 1px 7px; font-size: 0.72rem; color: #0284c7; border-color: #bae6fd; display: inline-flex; align-items: center; gap: 4px; height: 20px;">
                                    <i class="fa-brands fa-google-drive" style="color: #0284c7;"></i> Mở Drive
                                  </a>
                                ` : `
                                  <span style="color: #94a3b8; font-size: 0.72rem; display: inline-flex; align-items: center; gap: 4px;">
                                    <i class="fa-solid fa-cloud-arrow-up"></i> Chưa lên Drive
                                  </span>
                                `}
                              </div>
                            </div>
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
                <i class="fa-solid fa-plus"></i> Nhập đợt Giải ngân (Đề nghị thanh toán)
              </button>
            </div>

            <div class="data-table-wrap">
              <table class="data-table">
                <thead>
                  <tr>
                    <th>ĐN</th>
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
                      ? '<tr><td colspan="9" style="text-align:center; padding: 20px;">Chưa có đợt giải ngân nào. Bấm "+ Nhập đợt Giải ngân" để tạo.</td></tr>'
                      : realPayments
                          .map(
                            (p) => `
                        <tr>
                          <td><strong>${this.formatDnNo(p.payment_no || p.item_index || '01')}</strong></td>
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
                          <td style="white-space: nowrap;">
                            <button class="button button-outline btn-edit-payment" data-pay-id="${p.id}" style="padding: 2px 7px; font-size: 0.75rem; color: var(--primary); margin-right: 4px;" title="Chỉnh sửa đợt giải ngân (ĐN)">
                              <i class="fa-solid fa-pen-to-square"></i> Sửa
                            </button>
                            <button class="button button-outline btn-delete-payment" data-pay-id="${p.id}" style="padding: 2px 6px; font-size: 0.75rem; color: var(--danger);" title="Xóa đợt giải ngân">
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
      const docPath = this.getDriveDocPathForContract(contract);
      alert(`Đường dẫn thư mục Google Drive:\n\n1. Hợp đồng & Phụ lục:\n${drivePath}\n\n2. Thư mục Doc (Tài liệu khác):\n${docPath}\n\n(Tự động đồng bộ với Google Drive Cloud & Máy tính)`);
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

    // Xóa hồ sơ
    modal.querySelectorAll('.btn-delete-doc').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        const docId = btn.getAttribute('data-doc-id');
        if (confirm(`Bạn có chắc muốn xóa hồ sơ ${docId} không?`)) {
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
      btn.addEventListener('click', async (e) => {
        const payId = btn.getAttribute('data-pay-id');
        if (confirm(`Bạn có chắc muốn xóa đợt thanh toán này không?`)) {
          await db.delete('payments', payId);
          await this.loadData();
          await this.syncToLocalFile();
          this.openContractDetail(contract.id);
        }
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
        const newId = fd.get('project_id').trim().toUpperCase().replace(/[^A-Z0-9_]/g, '_');
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
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary" style="background: #d97706; border-color: #b45309;">
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
      alert(`Đã lưu thành công khoản Chi phí Ban "${scopeWork}" (${cpbId})!`);
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

    const confirmMsg = realPayments.length > 0
      ? `⚠️ CẢNH BÁO XÓA CHI PHÍ BAN!\n\nKhoản "${cpb.scope_of_work || cpb.id}" đang có ${realPayments.length} đợt chi với tổng số tiền ${this.formatCurrency(totalPaid)}.\n\nNếu xóa, toàn bộ ${realPayments.length} đợt chi này sẽ bị xóa vĩnh viễn khỏi sổ theo dõi!\n\nBạn có chắc chắn muốn tiếp tục không?`
      : `Bạn có chắc chắn muốn xóa khoản Chi phí Ban "${cpb.scope_of_work || cpb.id}" (${cpbId}) không?`;

    if (!confirm(confirmMsg)) return;

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
    alert(`Đã xóa thành công khoản Chi phí Ban "${cpb.scope_of_work || cpbId}"!`);
  }

  // ==========================================
  // FORM NHẬP / SỬA HỢP ĐỒNG HOẶC PHỤ LỤC (CHUẨN CPMS PLUS)
  // ==========================================
  openContractFormModal(editContract = null, parentContract = null) {
    const isEdit = !!editContract;
    const isAddingAppendix = !isEdit && !!parentContract;
    const targetParent = isEdit ? (editContract.parent_id ? this.contracts.find((c) => c.id === editContract.parent_id) : null) : parentContract;
    const isAppendix = (isEdit && !!editContract.parent_id) || isAddingAppendix;

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

    const curPrj = this.projects.find((p) => p.id === this.currentProjectId) || {};

    // Thu thập danh sách gợi ý dữ liệu đã có trong dự án (Autocomplete Datalists)
    const getUnique = (prop) => {
      const set = new Set();
      (this.contracts || []).forEach((c) => {
        const val = c[prop];
        if (val && String(val).trim()) set.add(String(val).trim());
      });
      return Array.from(set);
    };

    const companiesA = getUnique('company_a');
    const repsA = getUnique('rep_a');
    const companiesB = getUnique('company_b');
    const repsB = getUnique('rep_b');
    const companiesB2 = getUnique('company_b2');
    const repsB2 = getUnique('rep_b2');
    const taxCodes = getUnique('tax_code');
    const payees = getUnique('payee');
    const bankAccounts = getUnique('bank_account');
    const bankNames = getUnique('bank_name');
    const contactsB = getUnique('contact_person_b');

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
      <div class="modal-container" style="max-width: 840px; max-height: 92vh; display: flex; flex-direction: column;">
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

        <div class="modal-body" style="overflow-y: auto; padding: 20px;">
          <form id="form-contract">
            ${
              !isEdit && sampleContracts.length > 0
                ? `
                <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: var(--radius-sm); padding: 10px 14px; margin-bottom: 16px;">
                  <label class="form-label" style="font-size: 0.82rem; font-weight: 700; color: #1e40af; margin-bottom: 4px; display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-copy"></i> Sao chép từ Hợp đồng mẫu đã có (Tự động điền thông tin nhanh):
                  </label>
                  <select id="contract-template-picker" class="form-control" style="font-size: 0.85rem; background: #fff;">
                    <option value="">-- Chọn một hợp đồng trong dự án để sao chép thông tin --</option>
                    ${sampleContracts.map((sc) => `<option value="${sc.id}">[Nhóm ${sc.group}] ${sc.id} — ${sc.contract_no} (${sc.company_b || sc.company_a || 'Chưa rõ đối tác'})</option>`).join('')}
                  </select>
                  <small style="color: #64748b; font-size: 0.73rem;">Khi chọn, hệ thống tự động điền: Bên A, Bên B, MST, Người đại diện, Tài khoản ngân hàng, Phân loại... Bạn chỉ cần nhập số HĐ mới, ngày ký và giá trị.</small>
                </div>
              `
                : ''
            }

            <!-- 1. CONTRACT INFORMATION -->
            <div style="font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-file-invoice"></i> Contract Information (Thông tin Hợp đồng & Phân cấp)
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label" style="display: flex; justify-content: space-between;">
                  <span>Mã định danh (Contract ID) *</span>
                  <span style="font-weight: normal; color: var(--danger); font-size: 0.75rem;"><i class="fa-solid fa-lock"></i> Tự động nhảy - Không sửa tay</span>
                </label>
                <input type="text" class="form-control" id="contract-input-id" name="id" value="${formId}" readonly style="background: #f1f5f9; color: #334155; font-weight: 700; font-family: monospace; cursor: not-allowed;" required>
                <small style="color: var(--text-muted); font-size: 0.72rem;">Phân cấp chuẩn CPMS: Root (A-2026-01) ➔ Cấp 1 (A-2026-01-08) ➔ Cấp 2 (A-2026-01-08-01)</small>
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
                <label class="form-label">Số hiệu HĐ / Phụ lục (Contract No.) *</label>
                <input type="text" class="form-control" name="contract_no" id="input-contract-no" value="${editContract?.contract_no || (isAddingAppendix ? `APP-${String((this.getDirectChildren(parentContract.id).length || 0) + 1).padStart(2, '0')}` : '')}" placeholder="Vd: 114/2023/HĐXD/VINAE&C-CTC hoặc APP-08" required>
              </div>

              <div class="form-group">
                <label class="form-label">Loại hợp đồng (Contract Type) *</label>
                <select class="form-control" name="contract_type" id="select-contract-type" ${isAddingAppendix ? '' : ''}>
                  <option value="MAIN" ${defaultContractType === 'MAIN' ? 'selected' : ''}>MAIN (Hợp đồng chính)</option>
                  <option value="Appendix" ${defaultContractType === 'Appendix' ? 'selected' : ''}>Appendix (Phụ lục hợp đồng)</option>
                  <option value="Sub-Package" ${defaultContractType === 'Sub-Package' ? 'selected' : ''}>Sub-Package (Gói thầu con)</option>
                </select>
              </div>
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">Ngày ký (Signing Date) *</label>
                <input type="date" class="form-control" name="signing_date" id="input-signing-date" value="${this.formatDateForInput(editContract?.signing_date) || new Date().toISOString().slice(0, 10)}" required>
              </div>

              <div class="form-group">
                <label class="form-label">Thời hạn / Ngày kết thúc (End Date)</label>
                <input type="date" class="form-control" name="end_date" id="input-end-date" value="${this.formatDateForInput(editContract?.end_date)}">
              </div>

              <div class="form-group">
                <label class="form-label">Tình trạng (Status)</label>
                <select class="form-control" name="status" id="select-status">
                  <option value="Active" ${editContract?.status === 'Active' || !editContract ? 'selected' : ''}>Active (Đang thực hiện)</option>
                  <option value="Completed" ${editContract?.status === 'Completed' ? 'selected' : ''}>Completed (Đã quyết toán)</option>
                  <option value="Suspended" ${editContract?.status === 'Suspended' ? 'selected' : ''}>Suspended (Tạm ngưng)</option>
                </select>
              </div>
            </div>

            <!-- HỢP ĐỒNG LIÊN KẾT / TƯƠNG TỰ (LINKED CONTRACT) -->
            <div class="form-group" style="background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: var(--radius-sm); padding: 10px 12px; margin-top: 8px;">
              <label class="form-label" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                <span><i class="fa-solid fa-link" style="color: #6366f1;"></i> Hợp đồng liên kết / tương tự (Linked Contract)</span>
                <small style="color: var(--text-muted); font-size: 0.75rem;">Liên kết giữa HĐ Bên A (Chủ đầu tư) và HĐ Bên B (Thầu phụ) để chuyển nhanh</small>
              </label>
              <select class="form-control" name="linked_contract_id" id="contract-select-linked" style="font-size: 0.85rem;">
                <option value="">-- Không có liên kết --</option>
                ${
                  contractsA.length > 0
                    ? `<optgroup label="Hợp đồng Nhóm A (Chủ đầu tư / Tổng thầu)">
                        ${contractsA.map((ca) => `<option value="${ca.id}" ${(editContract?.linked_contract_id || baseSource.linked_contract_id) === ca.id ? 'selected' : ''}>${ca.id}: ${ca.contract_no} — ${ca.scope_of_work || ca.company_a || ''}</option>`).join('')}
                       </optgroup>`
                    : ''
                }
                ${
                  contractsB.length > 0
                    ? `<optgroup label="Hợp đồng Nhóm B (Thầu phụ / Nhà cung cấp)">
                        ${contractsB.map((cb) => `<option value="${cb.id}" ${(editContract?.linked_contract_id || baseSource.linked_contract_id) === cb.id ? 'selected' : ''}>${cb.id}: ${cb.contract_no} — ${cb.company_b || ''} (${cb.scope_of_work || ''})</option>`).join('')}
                       </optgroup>`
                    : ''
                }
              </select>
            </div>

            <!-- 2. PARTNERS - PROJECT -->
            <div style="font-weight: 700; font-size: 0.95rem; color: #4f46e5; margin: 18px 0 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-handshake"></i> Partners - Project (Đối tác & Thông tin gói thầu)
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">Bên A (Company A - Chủ đầu tư / Tổng thầu)</label>
                <input type="text" class="form-control" name="company_a" id="input-company-a" list="datalist-companies-a" value="${editContract?.company_a || baseSource.company_a || curPrj.client || ''}" placeholder="Tên đơn vị Bên A">
              </div>

              <div class="form-group">
                <label class="form-label">Người đại diện Bên A (Representative A)</label>
                <input type="text" class="form-control" name="rep_a" id="input-rep-a" list="datalist-reps-a" value="${editContract?.rep_a || baseSource.rep_a || ''}" placeholder="Họ tên người đại diện pháp luật Bên A">
              </div>
            </div>

            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: var(--radius-sm); padding: 12px 14px; margin-bottom: 14px;">
              <div style="font-weight: 700; font-size: 0.85rem; color: #334155; margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
                <i class="fa-solid fa-users-gear" style="color: #6366f1;"></i> Liên danh Nhà thầu Bên B (Nhà thầu phụ / Nhà cung cấp)
              </div>
              
              <div class="form-grid-2">
                <div class="form-group">
                  <label class="form-label">Nhà thầu B1 (Company B1 - Đơn vị đứng đầu liên danh / Nhà thầu chính) *</label>
                  <input type="text" class="form-control" name="company_b" id="input-company-b" list="datalist-companies-b" value="${editContract?.company_b || baseSource.company_b || ''}" placeholder="Tên công ty Bên B1" required>
                </div>
                <div class="form-group">
                  <label class="form-label">Người đại diện B1 (Representative B1)</label>
                  <input type="text" class="form-control" name="rep_b" id="input-rep-b" list="datalist-reps-b" value="${editContract?.rep_b || baseSource.rep_b || ''}" placeholder="Họ tên đại diện Bên B1">
                </div>
              </div>

              <!-- MÃ SỐ THUẾ & LIÊN HỆ BÊN B -->
              <div class="form-grid-2" style="margin-top: 8px;">
                <div class="form-group">
                  <label class="form-label"><i class="fa-solid fa-barcode" style="color: #059669;"></i> Mã số thuế Bên B (Tax Code)</label>
                  <input type="text" class="form-control" name="tax_code" id="input-tax-code" list="datalist-tax-codes" value="${editContract?.tax_code || baseSource.tax_code || ''}" placeholder="Vd: 0101234567">
                </div>
                <div class="form-group">
                  <label class="form-label">Người liên hệ Bên B (Contact Person B)</label>
                  <input type="text" class="form-control" name="contact_person_b" id="input-contact-b" list="datalist-contacts-b" value="${editContract?.contact_person_b || baseSource.contact_person_b || ''}" placeholder="Họ tên & số điện thoại người liên hệ Bên B">
                </div>
              </div>

              <div class="form-grid-2" style="margin-top: 8px;">
                <div class="form-group">
                  <label class="form-label">Nhà thầu B2 (Company B2 - Đơn vị Liên danh nếu có)</label>
                  <input type="text" class="form-control" name="company_b2" id="input-company-b2" list="datalist-companies-b2" value="${editContract?.company_b2 || baseSource.company_b2 || ''}" placeholder="Tên công ty liên danh B2 (để trống nếu không có liên danh)">
                </div>
                <div class="form-group">
                  <label class="form-label">Người đại diện B2 (Representative B2)</label>
                  <input type="text" class="form-control" name="rep_b2" id="input-rep-b2" list="datalist-reps-b2" value="${editContract?.rep_b2 || baseSource.rep_b2 || ''}" placeholder="Họ tên đại diện Bên B2">
                </div>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Phạm vi công việc / Nội dung gói thầu (Scope of work)</label>
              <textarea class="form-control" name="scope_of_work" id="input-scope-of-work" rows="2" placeholder="Tên hạng mục thi công, mua sắm thiết bị, tư vấn...">${editContract?.scope_of_work || (isAddingAppendix ? `Phụ lục bổ sung khối lượng cho ${parentContract.contract_no}` : '')}</textarea>
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">Giá trị ký kết (Contract Value - VNĐ) *</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" name="contract_val" id="input-contract-val" value="${this.formatNumberWithDots(editContract?.contract_val || 0)}" required style="font-weight: 700; color: var(--primary);">
              </div>

              <div class="form-group">
                <label class="form-label">Tỷ giá (Exchange Rate)</label>
                <input type="number" step="any" class="form-control" name="exchange_rate" id="input-exchange-rate" value="${editContract?.exchange_rate || baseSource.exchange_rate || 1}">
              </div>

              <div class="form-group">
                <label class="form-label">Phân loại HĐ (Contract Category)</label>
                <select class="form-control" name="category" id="select-category">
                  <option value="Construction" ${(editContract?.category || baseSource.category) === 'Construction' ? 'selected' : ''}>Construction (Thi công / Xây lắp)</option>
                  <option value="Procurement" ${(editContract?.category || baseSource.category) === 'Procurement' ? 'selected' : ''}>Procurement (Mua sắm vật tư)</option>
                  <option value="Consultancy" ${(editContract?.category || baseSource.category) === 'Consultancy' ? 'selected' : ''}>Consultancy (Tư vấn giám sát / Thẩm tra)</option>
                  <option value="Design" ${(editContract?.category || baseSource.category) === 'Design' ? 'selected' : ''}>Design (Thiết kế / Khảo sát)</option>
                  <option value="Equipment" ${(editContract?.category || baseSource.category) === 'Equipment' ? 'selected' : ''}>Equipment (Thiết bị MEP / Cơ điện)</option>
                  <option value="Service" ${(editContract?.category || baseSource.category) === 'Service' ? 'selected' : ''}>Service (Dịch vụ thí nghiệm / Vệ sinh / Bảo vệ)</option>
                  <option value="Maintenance" ${(editContract?.category || baseSource.category) === 'Maintenance' ? 'selected' : ''}>Maintenance (Bảo trì / Vận hành)</option>
                  <option value="Rental" ${(editContract?.category || baseSource.category) === 'Rental' ? 'selected' : ''}>Rental (Thuê thiết bị / Nhân công)</option>
                  <option value="Design & Build" ${(editContract?.category || baseSource.category) === 'Design & Build' ? 'selected' : ''}>Design & Build (Tổng thầu EPC / D&B)</option>
                  <option value="Khác" ${(editContract?.category || baseSource.category) === 'Khác' || !(editContract?.category || baseSource.category) ? 'selected' : ''}>Khác / Chưa phân loại</option>
                </select>
              </div>
            </div>

            <!-- 3. PAYMENT - BANKING -->
            <div style="font-weight: 700; font-size: 0.95rem; color: #0891b2; margin: 18px 0 12px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px;">
              <i class="fa-solid fa-building-columns"></i> Payment - Banking (Thông tin Thụ hưởng & Tài khoản Ngân hàng)
            </div>

            <div class="form-grid-3">
              <div class="form-group">
                <label class="form-label">Đơn vị thụ hưởng (Payee)</label>
                <input type="text" class="form-control" name="payee" id="input-payee" list="datalist-payees" value="${editContract?.payee || baseSource.payee || editContract?.company_b || baseSource.company_b || ''}" placeholder="Tên đơn vị thụ hưởng thanh toán">
              </div>

              <div class="form-group">
                <label class="form-label">Số tài khoản (Bank Acc Number)</label>
                <input type="text" class="form-control" name="bank_account" id="input-bank-account" list="datalist-bank-accounts" value="${editContract?.bank_account || baseSource.bank_account || ''}" placeholder="Vd: 1221 0000 7136 52">
              </div>

              <div class="form-group">
                <label class="form-label">Ngân hàng & Chi nhánh (Bank)</label>
                <input type="text" class="form-control" name="bank_name" id="input-bank-name" list="datalist-bank-names" value="${editContract?.bank_name || baseSource.bank_name || ''}" placeholder="Vd: BIDV - Chi nhánh Hà Thành">
              </div>
            </div>

            <!-- GỢI Ý DỮ LIỆU TỰ ĐỘNG (DATALISTS CHO MỌI TRƯỜNG THÔNG TIN) -->
            <datalist id="datalist-companies-a">
              ${companiesA.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-reps-a">
              ${repsA.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-companies-b">
              ${companiesB.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-reps-b">
              ${repsB.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-companies-b2">
              ${companiesB2.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-reps-b2">
              ${repsB2.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-tax-codes">
              ${taxCodes.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-payees">
              ${payees.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-bank-accounts">
              ${bankAccounts.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-bank-names">
              ${bankNames.map((v) => `<option value="${v}">`).join('')}
            </datalist>
            <datalist id="datalist-contacts-b">
              ${contactsB.map((v) => `<option value="${v}">`).join('')}
            </datalist>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; padding-top: 14px; border-top: 1px solid var(--border-color);">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary">
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
      setVal('#input-rep-a', tpl.rep_a || '');
      setVal('#input-company-b', tpl.company_b || '');
      setVal('#input-rep-b', tpl.rep_b || '');
      setVal('#input-company-b2', tpl.company_b2 || '');
      setVal('#input-rep-b2', tpl.rep_b2 || '');
      setVal('#input-tax-code', tpl.tax_code || '');
      setVal('#input-contact-b', tpl.contact_person_b || '');
      setVal('#input-scope-of-work', tpl.scope_of_work || '');
      setVal('#select-category', tpl.category || 'Construction');
      setVal('#input-payee', tpl.payee || tpl.company_b || '');
      setVal('#input-bank-account', tpl.bank_account || '');
      setVal('#input-bank-name', tpl.bank_name || '');
      setVal('#input-exchange-rate', tpl.exchange_rate || 1);
      setVal('#contract-select-linked', tpl.linked_contract_id || '');
    });

    // Tự động gợi ý điền thông tin khi chọn Nhà thầu B1 đã có trong dữ liệu
    const compBInput = modal.querySelector('#input-company-b');
    compBInput?.addEventListener('change', (e) => {
      const val = e.target.value?.trim();
      if (!val) return;
      const matched = (this.contracts || []).find((c) => c.company_b && c.company_b.trim().toLowerCase() === val.toLowerCase());
      if (matched) {
        const fillIfEmpty = (sel, newVal) => {
          const el = modal.querySelector(sel);
          if (el && !el.value && newVal) el.value = newVal;
        };
        fillIfEmpty('#input-rep-b', matched.rep_b);
        fillIfEmpty('#input-tax-code', matched.tax_code);
        fillIfEmpty('#input-contact-b', matched.contact_person_b);
        fillIfEmpty('#input-payee', matched.payee || matched.company_b);
        fillIfEmpty('#input-bank-account', matched.bank_account);
        fillIfEmpty('#input-bank-name', matched.bank_name);
      }
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
        company_a: formData.get('company_a') || '',
        rep_a: formData.get('rep_a') || '',
        company_b: formData.get('company_b') || '',
        rep_b: formData.get('rep_b') || '',
        company_b2: formData.get('company_b2') || '',
        rep_b2: formData.get('rep_b2') || '',
        scope_of_work: formData.get('scope_of_work') || '',
        contract_val: this.parseNumberFromDots(formData.get('contract_val')),
        exchange_rate: parseFloat(formData.get('exchange_rate')) || 1.0,
        category: formData.get('category') || 'Construction',
        contact_person_b: formData.get('contact_person_b') || '',
        payee: formData.get('payee') || formData.get('company_b') || '',
        bank_account: formData.get('bank_account') || '',
        bank_name: formData.get('bank_name') || '',
        tax_code: formData.get('tax_code') || '',
        linked_contract_id: formData.get('linked_contract_id') || '',
      };

      await db.put('contracts', updatedData);
      await this.loadData();
      this.render();
      this.updateEmptyStateUI();
      if (window.app) {
        await window.app.updateDashboardStats();
      }
      closeSubModal();
      await this.syncToLocalFile();
      alert('Đã lưu dữ liệu Hợp đồng thành công!');
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
              <label class="form-label"><i class="fa-solid fa-laptop"></i> Đường dẫn tệp tin trên máy tính (Local File Path)</label>
              <input type="text" class="form-control" name="local_link" placeholder="Vd: D:\\OneDrive\\DuAn\\CHC1\\HopDong\\HD_114.pdf">
            </div>

            <div class="form-group">
              <label class="form-label"><i class="fa-brands fa-google-drive" style="color: #0284c7;"></i> Link xem trực tuyến Google Drive (Online URL)</label>
              <input type="url" class="form-control" name="drive_link" placeholder="https://drive.google.com/file/d/...">
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
      const localLink = formData.get('local_link')?.trim() || '';
      const driveLink = formData.get('drive_link')?.trim() || '';
      const docData = {
        id: formData.get('id'),
        project_id: targetContract.project_id || this.currentProjectId,
        contract_id: targetContract.id,
        category: formData.get('category'),
        doc_no: formData.get('doc_no'),
        title: formData.get('title'),
        local_link: localLink,
        drive_link: driveLink,
        file_link: driveLink || localLink || '',
        received_date: formData.get('received_date'),
        status: formData.get('status'),
      };

      await db.put('documents', docData);
      await this.loadData();
      closeSubModal();
      await this.syncToLocalFile();
      this.openContractDetail(targetContract.id);
      alert('Đã lưu thành công Hồ sơ!');
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

    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 680px;">
        <div class="modal-header">
          <h3 id="payment-modal-title" style="font-size: 1.15rem; font-weight: 700; color: ${isCpb ? '#92400e' : 'inherit'};">
            ${isCpb ? `<i class="fa-solid fa-receipt" style="color: #d97706;"></i> Nhập đợt Chi phí cho ${targetContract.scope_of_work || targetContract.contract_no}` : `Nhập đợt Giải ngân cho ${targetContract.contract_no}`}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-payment">
            <div class="form-group" style="margin-bottom: 14px;">
              <label class="form-label">Chọn Hợp đồng / Khoản chi phí cần giải ngân <span style="color: var(--danger);">*</span></label>
              <select class="form-control" name="contract_id" id="payment-select-contract" required>
                ${availableContracts.map((c) => {
                  const isC = c.id === targetContract.id;
                  const cType = c.is_cpb || c.contract_type === 'CPB' ? 'CPB' : `Nhóm ${c.group}`;
                  return `<option value="${c.id}" ${isC ? 'selected' : ''}>[${cType}] ${c.contract_no || c.id} — ${c.scope_of_work || (c.is_cpb ? 'Chi phí Ban' : 'Hợp đồng')}</option>`;
                }).join('')}
              </select>
            </div>

            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label" id="lbl-payment-no">${isCpb ? 'Số Phiếu chi / Đợt chi' : 'Số Đề nghị thanh toán (Số ĐN)'}</label>
                <input type="text" class="form-control" name="payment_no" value="${String(nextIdx).padStart(2, '0')}" placeholder="Vd: 01 hoặc PC-01" required>
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-payment-date">${isCpb ? 'Ngày chi tiền' : 'Ngày thanh toán'}</label>
                <input type="date" class="form-control" name="payment_date" value="${new Date().toISOString().slice(0, 10)}" required>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label" id="lbl-item-name">${isCpb ? 'Nội dung khoản chi' : 'Diễn giải nội dung đợt thanh toán'}</label>
              <input type="text" class="form-control" name="item_name" placeholder="${isCpb ? 'Vd: Chi tiếp khách bàn giao hồ sơ, Chi tiền nước văn phòng...' : 'Vd: Tạm ứng đợt 1 / Nghiệm thu giai đoạn 2'}" required>
            </div>

            <div class="form-grid-2">
              <div class="form-group" id="grp-accepted-vol">
                <label class="form-label">Khối lượng Nghiệm thu (VNĐ)</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" id="pay-accepted" name="accepted_vol" value="0">
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-request-amount">${isCpb ? 'Số tiền đề nghị chi (VNĐ)' : 'Bên B Đề nghị thanh toán (VNĐ)'}</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" id="pay-requested" name="request_amount" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-paid-amount">${isCpb ? 'Số tiền thực tế đã chi (VNĐ)' : 'Bên A Thực tế đã thanh toán (VNĐ)'}</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" id="pay-paid" name="paid_amount" value="0" required>
              </div>
              <div class="form-group">
                <label class="form-label" id="lbl-invoice-ref">${isCpb ? 'Số Chứng từ / Phiếu chi / Hóa đơn' : 'Số Hóa đơn VAT / Tham chiếu'}</label>
                <input type="text" class="form-control" name="invoice_ref" placeholder="Vd: PC-001 hoặc HD-1234">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú diễn giải</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú người nhận tiền, số UNC hoặc mục đích chi..."></textarea>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" id="btn-submit-payment" class="button button-primary" style="${isCpb ? 'background: #d97706; border-color: #b45309;' : ''}">
                ${isCpb ? '<i class="fa-solid fa-receipt"></i> Ghi sổ Chi phí' : 'Ghi sổ Thanh toán'}
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

    // Lắng nghe thay đổi dropdown Hợp đồng
    const contractSelect = modal.querySelector('#payment-select-contract');
    contractSelect?.addEventListener('change', (e) => {
      const selectedId = e.target.value;
      const foundContract = this.contracts.find((c) => c.id === selectedId);
      if (foundContract) {
        const isCurrentCpb = foundContract.is_cpb || foundContract.contract_type === 'CPB' || (foundContract.id && foundContract.id.startsWith('B-CPB-'));
        const titleEl = modal.querySelector('#payment-modal-title');
        const submitBtn = modal.querySelector('#btn-submit-payment');
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
          submitBtn.innerHTML = isCurrentCpb ? '<i class="fa-solid fa-receipt"></i> Ghi sổ Chi phí' : 'Ghi sổ Thanh toán';
        }
        if (acceptedGrp) {
          acceptedGrp.style.display = isCurrentCpb ? 'none' : '';
        }
      }
    });

    modal.querySelector('#form-payment').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const chosenContractId = formData.get('contract_id') || targetContract.id;
      const actualContract = this.contracts.find((c) => c.id === chosenContractId) || targetContract;
      const isActualCpb = actualContract.is_cpb || actualContract.contract_type === 'CPB' || (actualContract.id && actualContract.id.startsWith('B-CPB-'));

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
        payment_date: formData.get('payment_date'),
        accepted_vol: isActualCpb ? paidAmt : this.parseNumberFromDots(formData.get('accepted_vol')),
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
      alert(isActualCpb ? 'Đã ghi sổ thành công đợt Chi phí Ban!' : 'Đã ghi sổ thành công đợt Giải ngân!');
    });
  }

  // ==========================================
  // FORM CHỈNH SỬA ĐỢT GIẢI NGÂN (SỬA ĐN) & ĐỢT CHI CPB
  // ==========================================
  openEditPaymentModal(payment, targetContract) {
    const isCpb = targetContract.is_cpb || targetContract.contract_type === 'CPB' || (targetContract.id && targetContract.id.startsWith('B-CPB-'));
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-container" style="max-width: 650px;">
        <div class="modal-header">
          <h3 style="font-size: 1.15rem; font-weight: 700; color: ${isCpb ? '#92400e' : 'inherit'};">
            <i class="fa-solid fa-pen-to-square" style="color: ${isCpb ? '#d97706' : 'var(--primary)'};"></i>
            ${isCpb ? `Chỉnh sửa Đợt Chi phí: ${this.formatDnNo(payment.payment_no)}` : `Chỉnh sửa Đợt Giải ngân: ${this.formatDnNo(payment.payment_no)}`}
          </h3>
          <button class="button button-outline btn-close-submodal" style="padding: 4px 8px;"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="modal-body">
          <form id="form-edit-payment">
            <div class="form-grid-2">
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số Phiếu chi / Đợt chi' : 'Số Đề nghị thanh toán (Số ĐN)'} <span style="color: var(--danger);">*</span></label>
                <input type="text" class="form-control" name="payment_no" value="${payment.payment_no || ''}" placeholder="Vd: 01 hoặc PC-01" required>
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Ngày chi tiền' : 'Ngày thanh toán'}</label>
                <input type="date" class="form-control" name="payment_date" value="${this.formatDateForInput(payment.payment_date)}" required>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">${isCpb ? 'Nội dung khoản chi' : 'Diễn giải nội dung đợt thanh toán'}</label>
              <input type="text" class="form-control" name="item_name" value="${payment.item_name || ''}" placeholder="${isCpb ? 'Vd: Chi tiếp khách, chi văn phòng...' : 'Vd: Tạm ứng đợt 1 / Nghiệm thu giai đoạn 2'}" required>
            </div>

            <div class="form-grid-2">
              ${isCpb ? '' : `
                <div class="form-group">
                  <label class="form-label">Khối lượng Nghiệm thu (VNĐ)</label>
                  <input type="text" inputmode="numeric" class="form-control currency-input" id="edit-pay-accepted" name="accepted_vol" value="${this.formatNumberWithDots(payment.accepted_vol || 0)}">
                </div>
              `}
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số tiền đề nghị chi (VNĐ)' : 'Bên B Đề nghị thanh toán (VNĐ)'}</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" id="edit-pay-requested" name="request_amount" value="${this.formatNumberWithDots(payment.request_amount || 0)}" required>
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số tiền thực tế đã chi (VNĐ)' : 'Bên A Thực tế đã thanh toán (VNĐ)'}</label>
                <input type="text" inputmode="numeric" class="form-control currency-input" id="edit-pay-paid" name="paid_amount" value="${this.formatNumberWithDots(payment.paid_amount || 0)}" required>
              </div>
              <div class="form-group">
                <label class="form-label">${isCpb ? 'Số Chứng từ / Phiếu chi / Hóa đơn' : 'Số Hóa đơn VAT / Tham chiếu'}</label>
                <input type="text" class="form-control" name="invoice_ref" value="${payment.invoice_ref || ''}" placeholder="Vd: PC-001 hoặc HD-1234">
              </div>
            </div>

            <div class="form-group">
              <label class="form-label">Ghi chú diễn giải</label>
              <textarea class="form-control" name="note" rows="2" placeholder="Ghi chú người nhận tiền, số UNC...">${payment.note || ''}</textarea>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
              <button type="button" class="button button-outline btn-close-submodal">Hủy</button>
              <button type="submit" class="button button-primary" style="${isCpb ? 'background: #d97706; border-color: #b45309;' : ''}">
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

    modal.querySelector('#form-edit-payment').addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const reqAmt = this.parseNumberFromDots(formData.get('request_amount'));
      const paidAmt = this.parseNumberFromDots(formData.get('paid_amount'));

      payment.payment_no = formData.get('payment_no');
      payment.payment_date = formData.get('payment_date');
      payment.item_name = formData.get('item_name');
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
      alert(`Đã cập nhật thành công đợt Giải ngân ${this.formatDnNo(payment.payment_no)}!`);
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
      csvContent += `"${this.formatDnNo(p.payment_no || p.item_index)}","${(p.item_name || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${p.accepted_vol},${p.request_amount},${p.paid_amount},${p.remaining},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
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

    // 1. TÍNH TOÁN BẢNG TỔNG HỢP SUMMARY (FINANCIAL SUMMARY)
    const realPayments = this.payments.filter((p) => !p.is_header);

    // Summary Nhóm A (Thu)
    const contractsA = this.contracts.filter((c) => c.group === 'A');
    const mainA = contractsA.filter((c) => !c.parent_id);
    let summaryA = {
      estimatedVal: 0,
      contractVal: 0,
      acceptedVol: 0,
      certifiedVol: 0,
      requestedAmt: 0,
      paidAmt: 0,
      remaining: 0,
    };
    mainA.forEach((c) => {
      const st = this.calculateTotals(c);
      summaryA.contractVal += st.totalSigned;
      summaryA.estimatedVal += (c.estimated_val || st.totalSigned);
    });
    const familyAIds = new Set();
    contractsA.forEach((c) => familyAIds.add(c.id));
    realPayments.forEach((p) => {
      if (familyAIds.has(p.contract_id)) {
        summaryA.acceptedVol += Number(p.accepted_vol) || 0;
        summaryA.certifiedVol += Number(p.certified_vol || p.accepted_vol) || 0;
        summaryA.requestedAmt += Number(p.request_amount) || 0;
        summaryA.paidAmt += Number(p.paid_amount) || 0;
      }
    });
    summaryA.remaining = summaryA.requestedAmt - summaryA.paidAmt;

    // Summary Nhóm B (Chi - TÍNH GỘP CẢ CÁC HỢP ĐỒNG B VÀ CHI PHÍ BAN CPB)
    const contractsB = this.contracts.filter((c) => c.group === 'B');
    const mainB = contractsB.filter((c) => !c.parent_id);
    let summaryB = {
      estimatedVal: 0,
      contractVal: 0,
      acceptedVol: 0,
      certifiedVol: 0,
      requestedAmt: 0,
      paidAmt: 0,
      remaining: 0,
    };
    mainB.forEach((c) => {
      const st = this.calculateTotals(c);
      summaryB.contractVal += st.totalSigned;
      summaryB.estimatedVal += (c.estimated_val || st.totalSigned);
    });
    const familyBIds = new Set();
    contractsB.forEach((c) => familyBIds.add(c.id));
    realPayments.forEach((p) => {
      if (familyBIds.has(p.contract_id)) {
        const isCpbPay = /^B-CPB/i.test(p.contract_id);
        const pAccepted = Number(p.accepted_vol) || (isCpbPay ? Number(p.paid_amount) || 0 : 0);
        summaryB.acceptedVol += pAccepted;
        summaryB.certifiedVol += Number(p.certified_vol) || pAccepted;
        summaryB.requestedAmt += Number(p.request_amount) || 0;
        summaryB.paidAmt += Number(p.paid_amount) || 0;
      }
    });
    summaryB.remaining = summaryB.requestedAmt - summaryB.paidAmt;

    // Cân đối dòng tiền A - B
    const balance = summaryA.paidAmt - summaryB.paidAmt;

    // 1.1 RENDER BẢNG SUMMARY CHUẨN CPMS PLUS
    summaryContainer.innerHTML = `
      <div class="disburse-summary-header">
        <div style="font-weight: 800; font-size: 0.95rem; color: #1e293b; display: flex; align-items: center; gap: 8px;">
          <i class="fa-solid fa-chart-pie" style="color: var(--primary);"></i>
          <span>BẢNG TỔNG HỢP TÀI CHÍNH & GIẢI NGÂN (SUMMARY)</span>
        </div>
        <div style="font-size: 0.8rem; color: var(--text-muted);">
          Dự án: <strong style="color: var(--text-main);">${prjName}</strong>
        </div>
      </div>
      <div style="overflow-x: auto;">
        <table class="disburse-summary-table">
          <thead>
            <tr>
              <th style="text-align: left;">Phân nhóm (Group)</th>
              <th style="text-align: right;">Dự toán (Estimated Val)</th>
              <th style="text-align: right;">Giá trị HĐ (Contract Val)</th>
              <th style="text-align: right;">Khối lượng NT (Accepted Vol)</th>
              <th style="text-align: right;">Đề nghị TT (Requested Amt)</th>
              <th style="text-align: right;">Đã thanh toán (Paid Amt)</th>
              <th style="text-align: right;">Còn lại (Remaining)</th>
            </tr>
          </thead>
          <tbody>
            <tr style="background: #f8fafc; font-weight: 600;">
              <td>
                <span class="badge badge-primary" style="font-size: 0.72rem; padding: 2px 6px;"><i class="fa-solid fa-landmark"></i> Group-A</span>
                <span style="margin-left: 6px; color: #1e293b;">Nhóm A: Dòng tiền Thu (Chủ đầu tư / Tổng thầu)</span>
              </td>
              <td style="text-align: right;">${this.formatCurrency(summaryA.estimatedVal)}</td>
              <td style="text-align: right; font-weight: 700; color: var(--primary);">${this.formatCurrency(summaryA.contractVal)}</td>
              <td style="text-align: right;">${this.formatCurrency(summaryA.acceptedVol)}</td>
              <td style="text-align: right;">${this.formatCurrency(summaryA.requestedAmt)}</td>
              <td style="text-align: right; font-weight: 700; color: var(--success);">${this.formatCurrency(summaryA.paidAmt)}</td>
              <td style="text-align: right; color: ${summaryA.remaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatCurrency(summaryA.remaining)}</td>
            </tr>

            <tr style="background: #ffffff; font-weight: 600;">
              <td>
                <span class="badge badge-warning" style="font-size: 0.72rem; padding: 2px 6px;"><i class="fa-solid fa-handshake"></i> Group-B</span>
                <span style="margin-left: 6px; color: #1e293b;">Nhóm B: Dòng tiền Chi (Thầu phụ, NCC & Chi phí Ban CPB)</span>
              </td>
              <td style="text-align: right;">${this.formatCurrency(summaryB.estimatedVal)}</td>
              <td style="text-align: right; font-weight: 700; color: #4f46e5;">${this.formatCurrency(summaryB.contractVal)}</td>
              <td style="text-align: right;">${this.formatCurrency(summaryB.acceptedVol)}</td>
              <td style="text-align: right;">${this.formatCurrency(summaryB.requestedAmt)}</td>
              <td style="text-align: right; font-weight: 700; color: #0891b2;">${this.formatCurrency(summaryB.paidAmt)}</td>
              <td style="text-align: right; color: ${summaryB.remaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatCurrency(summaryB.remaining)}</td>
            </tr>

            <tr style="background: #f0fdf4; border-top: 2px solid #86efac; font-weight: 700;">
              <td colspan="5" style="color: #166534; font-size: 0.825rem;">
                <i class="fa-solid fa-scale-balanced" style="color: #16a34a; margin-right: 6px;"></i> Cân đối Dòng tiền Dự án (A Thu - B Chi)
              </td>
              <td style="text-align: right; font-size: 0.95rem; color: ${balance >= 0 ? 'var(--success)' : 'var(--danger)'};">
                ${balance >= 0 ? '+' : ''}${this.formatCurrency(balance)}
              </td>
              <td style="text-align: right; font-size: 0.8rem; color: #166534;">
                ${balance >= 0 ? 'Thặng dư dòng tiền' : 'Thâm hụt dòng tiền'}
              </td>
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

    // 3. LẤY DANH SÁCH HỢP ĐỒNG THUỘC NHÓM ĐANG CHỌN (KÈM CPB TRONG NHÓM B)
    const curGroupContracts = this.contracts.filter((c) => c.group === this.disburseGroup);
    
    // Nếu có tìm kiếm thì lọc
    const searchTerm = (this.disburseContractSearchTerm || '').trim().toLowerCase();
    const filteredContracts = curGroupContracts.filter((c) => {
      if (!searchTerm) return true;
      const cNo = String(c.contract_no || c.id || '').toLowerCase();
      const cScope = String(c.scope_of_work || '').toLowerCase();
      const cPartA = String(c.company_a || '').toLowerCase();
      const cPartB = String(c.company_b || '').toLowerCase();
      return cNo.includes(searchTerm) || cScope.includes(searchTerm) || cPartA.includes(searchTerm) || cPartB.includes(searchTerm);
    });

    // Xác định mục đang chọn (nếu chưa có hoặc không nằm trong nhóm thì chọn mục đầu tiên)
    const existsSelected = filteredContracts.some((c) => c.id === this.selectedDisburseContractId);
    if (!existsSelected && filteredContracts.length > 0) {
      // Ưu tiên chọn mục đã có thanh toán
      const firstWithPay = filteredContracts.find((c) => realPayments.some((p) => p.contract_id === c.id));
      this.selectedDisburseContractId = firstWithPay ? firstWithPay.id : filteredContracts[0].id;
    } else if (filteredContracts.length === 0) {
      this.selectedDisburseContractId = null;
    }

    // 3.1 RENDER BẢNG DANH SÁCH CÁC HỢP ĐỒNG CỦA NHÓM ĐANG CHỌN
    if (filteredContracts.length === 0) {
      contractsTableContainer.innerHTML = `
        <div style="text-align: center; padding: 36px 20px; color: var(--text-muted);">
          <i class="fa-solid fa-folder-open" style="font-size: 2rem; color: #cbd5e1; margin-bottom: 8px;"></i>
          <p style="margin: 0; font-size: 0.85rem;">Không tìm thấy hợp đồng nào phù hợp trong Nhóm ${this.disburseGroup}.</p>
        </div>
      `;
    } else {
      let contractRowsHtml = '';
      filteredContracts.forEach((c) => {
        const isSelected = c.id === this.selectedDisburseContractId;
        const isCpb = c.is_cpb || c.contract_type === 'CPB' || (c.id && c.id.startsWith('B-CPB-'));
        
        // Tính tổng hợp tài chính của riêng mục này
        const isMain = !c.parent_id;
        const familyIds = this.getAllDescendantIds(c.id);
        const cPayments = realPayments.filter((p) => isMain ? familyIds.includes(p.contract_id) : p.contract_id === c.id);
        
        let cAccepted = 0;
        let cRequested = 0;
        let cPaid = 0;
        cPayments.forEach((p) => {
          cAccepted += Number(p.accepted_vol) || (isCpb ? Number(p.paid_amount) || 0 : 0);
          cRequested += Number(p.request_amount) || 0;
          cPaid += Number(p.paid_amount) || 0;
        });
        const cRemaining = cRequested - cPaid;
        const cVal = c.contract_val || 0;
        const cRate = cVal > 0 ? ((cPaid / cVal) * 100).toFixed(1) : (cRequested > 0 ? ((cPaid / cRequested) * 100).toFixed(1) : 0);

        const partnerName = this.disburseGroup === 'A' ? (c.company_a || 'Chủ đầu tư') : (c.company_b || 'Ban QLDA / Thầu phụ');

        contractRowsHtml += `
          <tr class="disburse-contract-row ${isSelected ? 'selected-contract-row' : ''}" data-disburse-cid="${c.id}">
            <td style="white-space: nowrap; font-weight: 700; padding: 10px 14px;">
              ${isCpb 
                ? `<span class="badge" style="background: #fef3c7; color: #b45309; border: 1px solid #fde68a;"><i class="fa-solid fa-receipt"></i> CPB</span>` 
                : `<span style="font-family: monospace; color: var(--primary);">${c.display_no || c.id}</span>`}
            </td>
            <td style="font-weight: 700; color: #1e293b; padding: 10px 14px; white-space: nowrap;">
              ${c.contract_no || c.id}
            </td>
            <td style="padding: 10px 14px;">
              <div style="font-weight: 600; color: #1e293b; max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${c.scope_of_work || ''}">
                ${c.scope_of_work || 'Hợp đồng thực hiện dự án'}
              </div>
              <div style="font-size: 0.75rem; color: var(--text-muted); max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                <i class="fa-solid fa-building" style="margin-right: 4px;"></i> ${partnerName}
              </div>
            </td>
            <td style="text-align: right; font-weight: 700; color: #1e293b; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(cVal)}
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(cAccepted)}
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(cRequested)}
            </td>
            <td style="text-align: right; font-weight: 700; color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'}; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(cPaid)}
            </td>
            <td style="text-align: right; color: ${cRemaining > 0 ? 'var(--danger)' : '#64748b'}; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(cRemaining)}
            </td>
            <td style="text-align: center; padding: 10px 14px; white-space: nowrap;">
              <span class="badge ${Number(cRate) >= 100 ? 'badge-success' : 'badge-primary'}" style="font-size: 0.72rem;">
                ${cRate}%
              </span>
            </td>
          </tr>
        `;
      });

      contractsTableContainer.innerHTML = `
        <div style="padding: 10px 16px; background: #f8fafc; border-bottom: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center; font-size: 0.825rem;">
          <span style="font-weight: 700; color: #1e293b;">
            <i class="fa-solid fa-list-check" style="color: var(--primary);"></i> DANH SÁCH HỢP ĐỒNG & GIÁ TRỊ TỔNG HỢP (NHÓM ${this.disburseGroup})
          </span>
          <span style="color: var(--text-muted);">
            (Click vào một dòng để xem chi tiết các đợt giải ngân)
          </span>
        </div>
        <div style="overflow-x: auto;">
          <table class="disburse-summary-table" style="width: 100%;">
            <thead>
              <tr>
                <th style="width: 80px; text-align: left;">Mã</th>
                <th style="text-align: left;">Số hiệu HĐ / Phụ lục</th>
                <th style="text-align: left;">Nội dung / Nhà thầu</th>
                <th style="text-align: right;">Giá trị HĐ (Contract Val)</th>
                <th style="text-align: right;">Khối lượng NT (Accepted Vol)</th>
                <th style="text-align: right;">Đề nghị TT (Requested Amt)</th>
                <th style="text-align: right;">Đã thanh toán (Paid Amt)</th>
                <th style="text-align: right;">Còn lại (Remaining)</th>
                <th style="width: 80px; text-align: center;">Tỷ lệ (%)</th>
              </tr>
            </thead>
            <tbody>
              ${contractRowsHtml}
            </tbody>
          </table>
        </div>
      `;

      // Gắn sự kiện click chọn dòng hợp đồng
      contractsTableContainer.querySelectorAll('.disburse-contract-row').forEach((row) => {
        row.addEventListener('click', () => {
          const cid = row.getAttribute('data-disburse-cid');
          this.selectedDisburseContractId = cid;
          this.renderDisbursementsTab();
        });
      });
    }

    // 4. RENDER BẢNG CHI TIẾT CÁC ĐỢT GIẢI NGÂN (DN / IPC) CỦA HỢP ĐỒNG ĐANG CHỌN
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

    const isSelectedCpb = selectedContract.is_cpb || selectedContract.contract_type === 'CPB' || (selectedContract.id && selectedContract.id.startsWith('B-CPB-'));
    const isMainSelected = !selectedContract.parent_id;
    const selectedFamilyIds = this.getAllDescendantIds(selectedContract.id);
    
    // Lấy các đợt thanh toán của hợp đồng đang chọn
    const contractPayments = realPayments.filter((p) => 
      isMainSelected ? selectedFamilyIds.includes(p.contract_id) : p.contract_id === selectedContract.id
    );

    // Sắp xếp các đợt giải ngân theo ngày thanh toán hoặc item_index
    contractPayments.sort((a, b) => {
      const dateA = a.payment_date ? new Date(a.payment_date).getTime() : 0;
      const dateB = b.payment_date ? new Date(b.payment_date).getTime() : 0;
      return dateA - dateB;
    });

    // Tính tổng kết Total của hợp đồng này
    let dnTotalAccepted = 0;
    let dnTotalRequested = 0;
    let dnTotalPaid = 0;
    contractPayments.forEach((p) => {
      dnTotalAccepted += Number(p.accepted_vol) || (isSelectedCpb ? Number(p.paid_amount) || 0 : 0);
      dnTotalRequested += Number(p.request_amount) || 0;
      dnTotalPaid += Number(p.paid_amount) || 0;
    });
    const dnTotalRemaining = dnTotalRequested - dnTotalPaid;

    let detailRowsHtml = '';
    if (contractPayments.length === 0) {
      detailRowsHtml = `
        <tr>
          <td colspan="9" style="text-align: center; padding: 32px 16px; color: var(--text-muted);">
            <i class="fa-solid fa-receipt" style="font-size: 1.8rem; color: #cbd5e1; margin-bottom: 8px; display: block;"></i>
            Hợp đồng này chưa có đợt giải ngân nào. Bấm nút <strong>"+ Thêm đợt thanh toán"</strong> ở góc phải để ghi nhận đợt thanh toán đầu tiên.
          </td>
        </tr>
      `;
    } else {
      contractPayments.forEach((p, idx) => {
        detailRowsHtml += `
          <tr style="border-bottom: 1px solid #f1f5f9;">
            <td style="font-weight: 700; padding: 10px 14px; white-space: nowrap;">
              <span style="background: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 4px; font-size: 0.8rem;">
                ${this.formatDnNo(p.payment_no || String(idx + 1))}
              </span>
            </td>
            <td style="padding: 10px 14px; white-space: nowrap; font-size: 0.825rem; color: #475569;">
              <i class="fa-regular fa-calendar-days" style="margin-right: 4px;"></i> ${this.formatDate(p.payment_date)}
            </td>
            <td style="padding: 10px 14px;">
              <div style="font-weight: 600; color: #1e293b;">${p.item_name || 'Đợt thanh toán'}</div>
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap;">
              ${this.formatCurrency(p.accepted_vol || (isSelectedCpb ? p.paid_amount : 0))}
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap; font-weight: 600;">
              ${this.formatCurrency(p.request_amount || 0)}
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap; font-weight: 700; color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'};">
              ${this.formatCurrency(p.paid_amount || 0)}
            </td>
            <td style="text-align: right; padding: 10px 14px; white-space: nowrap; color: ${(p.remaining || 0) > 0 ? 'var(--danger)' : '#64748b'};">
              ${this.formatCurrency(p.remaining || 0)}
            </td>
            <td style="padding: 10px 14px; font-size: 0.8rem; color: #64748b;">
              ${p.invoice_ref ? `<span style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: monospace;">${p.invoice_ref}</span>` : ''}
              ${p.note ? `<div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">${p.note}</div>` : ''}
              ${!p.invoice_ref && !p.note ? '—' : ''}
            </td>
            <td style="padding: 10px 14px; text-align: center; white-space: nowrap;">
              <div style="display: flex; gap: 4px; justify-content: center;">
                <button class="button button-outline btn-edit-selected-dn" data-pay-id="${p.id}" style="padding: 3px 8px; font-size: 0.75rem;" title="Chỉnh sửa đợt thanh toán này">
                  <i class="fa-solid fa-pen-to-square"></i>
                </button>
                <button class="button button-outline btn-delete-selected-dn" data-pay-id="${p.id}" style="padding: 3px 8px; font-size: 0.75rem; color: var(--danger); border-color: #fca5a5;" title="Xóa đợt thanh toán này">
                  <i class="fa-solid fa-trash-can"></i>
                </button>
              </div>
            </td>
          </tr>
        `;
      });
    }

    const partnerLabel = this.disburseGroup === 'A' ? (selectedContract.company_a || 'Chủ đầu tư') : (selectedContract.company_b || 'Ban QLDA / Thầu phụ');
    const partnerRole = this.disburseGroup === 'A' ? 'Bên A' : 'Bên B';

    detailContainer.innerHTML = `
      <div class="disburse-detail-header">
        <div>
          <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
            <span class="badge ${this.disburseGroup === 'A' ? 'badge-primary' : 'badge-warning'}">
              ${this.disburseGroup === 'A' ? 'Nhóm A (Thu)' : 'Nhóm B (Chi)'}
            </span>
            <strong style="font-size: 1.05rem; color: #1e293b;">
              ${selectedContract.contract_no || selectedContract.id}
            </strong>
            <span style="color: var(--text-muted); font-size: 0.85rem;">— ${selectedContract.scope_of_work || 'Hợp đồng thực hiện dự án'}</span>
          </div>
          <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 4px; display: flex; gap: 14px; flex-wrap: wrap;">
            <span><i class="fa-solid fa-building"></i> ${partnerRole}: <strong>${partnerLabel}</strong></span>
            <span><i class="fa-solid fa-file-contract"></i> Giá trị HĐ: <strong>${this.formatCurrency(selectedContract.contract_val)}</strong></span>
            <span><i class="fa-solid fa-circle-check" style="color: var(--success);"></i> Đã thanh toán: <strong style="color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatCurrency(dnTotalPaid)}</strong></span>
          </div>
        </div>

        <button id="btn-add-dn-for-selected" class="button button-primary button-sm" style="font-weight: 600;" title="Thêm đợt thanh toán cho hợp đồng này">
          <i class="fa-solid fa-plus"></i> + Thêm Đợt thanh toán
        </button>
      </div>

      <div style="overflow-x: auto;">
        <table class="disburse-summary-table" style="width: 100%;">
          <thead>
            <tr style="background: #f8fafc;">
              <th style="width: 90px; text-align: left;">Đợt (DN)</th>
              <th style="width: 110px; text-align: left;">Ngày TT</th>
              <th style="text-align: left;">Nội dung thanh toán (Payment Reference)</th>
              <th style="text-align: right; width: 140px;">Khối lượng NT</th>
              <th style="text-align: right; width: 140px;">Đề nghị TT</th>
              <th style="text-align: right; width: 140px;">Thực thanh toán</th>
              <th style="text-align: right; width: 130px;">Còn lại</th>
              <th style="text-align: left; width: 150px;">Chứng từ / Ghi chú</th>
              <th style="width: 90px; text-align: center;">Thao tác</th>
            </tr>
          </thead>
          <tbody>
            ${detailRowsHtml}
          </tbody>
          ${contractPayments.length > 0 ? `
            <tfoot>
              <tr style="background: #f8fafc; border-top: 2px solid #cbd5e1; font-weight: 700;">
                <td colspan="3" style="padding: 10px 14px; color: #1e293b;">
                  TỔNG CỘNG (${contractPayments.length} đợt thanh toán)
                </td>
                <td style="text-align: right; padding: 10px 14px;">${this.formatCurrency(dnTotalAccepted)}</td>
                <td style="text-align: right; padding: 10px 14px;">${this.formatCurrency(dnTotalRequested)}</td>
                <td style="text-align: right; padding: 10px 14px; color: ${this.disburseGroup === 'A' ? 'var(--success)' : '#0891b2'};">${this.formatCurrency(dnTotalPaid)}</td>
                <td style="text-align: right; padding: 10px 14px; color: ${dnTotalRemaining > 0 ? 'var(--danger)' : '#64748b'};">${this.formatCurrency(dnTotalRemaining)}</td>
                <td colspan="2"></td>
              </tr>
            </tfoot>
          ` : ''}
        </table>
      </div>
    `;

    // 5. GẮN SỰ KIỆN TRONG TAB GIẢI NGÂN
    // Nút thêm đợt thanh toán cho HĐ đang chọn
    document.getElementById('btn-add-dn-for-selected')?.addEventListener('click', () => {
      this.openPaymentEntryModal(selectedContract);
    });

    // Nút sửa đợt thanh toán
    detailContainer.querySelectorAll('.btn-edit-selected-dn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const payId = btn.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (pay) {
          const c = this.contracts.find((item) => item.id === pay.contract_id) || selectedContract;
          this.openEditPaymentModal(pay, c);
        }
      });
    });

    // Nút xóa đợt thanh toán
    detailContainer.querySelectorAll('.btn-delete-selected-dn').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const payId = btn.getAttribute('data-pay-id');
        const pay = this.payments.find((p) => p.id === payId);
        if (!pay) return;
        const confirmed = confirm(`Bạn có chắc chắn muốn xóa đợt thanh toán "${pay.item_name || pay.payment_no}" (${this.formatCurrency(pay.paid_amount || pay.request_amount)}) không?`);
        if (!confirmed) return;
        await db.delete('payments', payId);
        await this.loadData();
        await this.syncToLocalFile();
        this.renderDisbursementsTab();
      });
    });

    // Lắng nghe sự kiện Toolbar giải ngân 1 lần duy nhất
    if (!this.disburseEventsBound) {
      this.disburseEventsBound = true;

      // Chuyển tab Nhóm A Thu / Nhóm B Chi
      document.querySelectorAll('.disburse-group-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          this.disburseGroup = btn.getAttribute('data-disburse-group');
          this.selectedDisburseContractId = null;
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

      // Tìm kiếm hợp đồng giải ngân
      document.getElementById('disburse-contract-search')?.addEventListener('input', (e) => {
        this.disburseContractSearchTerm = e.target.value.trim().toLowerCase();
        this.renderDisbursementsTab();
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

      csvContent += `"${idx + 1}","${this.formatDnNo(p.payment_no || p.item_index)}","${groupName}","${p.contract_id || ''}","${cNo.replace(/"/g, '""')}","${partner.replace(/"/g, '""')}","${(p.item_name || '').replace(/"/g, '""')}","${this.formatDate(p.payment_date)}",${p.request_amount || 0},${p.paid_amount || 0},"${p.invoice_ref || ''}","${(p.note || '').replace(/"/g, '""')}"\r\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `So_Giai_Ngan_${prjName}_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }
}

if (!window.__MAIO_CONTRACTS_MANAGER__) {
  window.__MAIO_CONTRACTS_MANAGER__ = new ContractsManager();
}
export const contractsManager = window.__MAIO_CONTRACTS_MANAGER__;
