/**
 * MAIO CPMS Importer & Exporter Engine
 * Đọc và xuất file dữ liệu định dạng .cpms (.xlsb / .xlsx) chuẩn xác 100% với phần mềm CPMS Plus.
 */

import { db } from './db.js?v=5.3';

export class CPMSImporterExporter {
  /**
   * Chuyển đổi đối tượng Date sang chuỗi YYYY-MM-DD theo giờ Local (tránh lệch 1 ngày do UTC+7)
   */
  static formatLocalYmd(d) {
    const adjusted = new Date(d.getTime() + 60 * 1000);
    const y = adjusted.getFullYear();
    const m = String(adjusted.getMonth() + 1).padStart(2, '0');
    const day = String(adjusted.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  /**
   * Chuyển đổi định dạng ngày trong Excel sang chuỗi YYYY-MM-DD (không bị lệch múi giờ)
   */
  static parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return null;
      return this.formatLocalYmd(val);
    }
    if (typeof val === 'number') {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (isNaN(d.getTime())) return null;
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
    const s = String(val).trim();
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      return s.slice(0, 10);
    }
    const dmY = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmY) {
      return `${dmY[3]}-${String(dmY[2]).padStart(2, '0')}-${String(dmY[1]).padStart(2, '0')}`;
    }
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return this.formatLocalYmd(parsed);
    }
    return s;
  }

  /**
   * Chuyển đổi chuỗi ngày thành ô Date chuẩn cho SheetJS với NumberFormat d/mm/yyyy
   */
  static dateToExcelCell(val) {
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
      if (!trimmed) return '';
      if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
        const [y, m, day] = trimmed.slice(0, 10).split('-').map(Number);
        d = new Date(y, m - 1, day);
      } else {
        const dmY = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dmY) {
          d = new Date(Number(dmY[3]), Number(dmY[2]) - 1, Number(dmY[1]));
        } else {
          const parsed = new Date(trimmed);
          if (!isNaN(parsed.getTime())) d = parsed;
        }
      }
    }
    if (d && !isNaN(d.getTime())) {
      return { t: 'd', v: d, z: 'd/mm/yyyy' };
    }
    return val;
  }

  /**
   * Phân tích mã ContractID để xác định cấp bậc và HĐ cha
   * Quy tắc chuẩn:
   * A-2026-01: Root (HĐ gốc nhóm A) -> level = 0, parent_id = null
   * A-2026-01-02: Phụ lục cấp 1 của A-2026-01 -> level = 1, parent_id = A-2026-01
   * A-2026-01-02-01: Gói con cấp 2 của A-2026-01-02 -> level = 2, parent_id = A-2026-01-02
   */
  static parseContractID(contractId) {
    if (!contractId) return { level: 0, parentId: null, group: 'A' };
    const parts = contractId.trim().split('-');
    const group = parts[0] === 'B' ? 'B' : 'A';

    // Đếm số đoạn:
    // A-2026-01 (3 đoạn) -> level 0 (Root)
    // A-2026-01-02 (4 đoạn) -> level 1 (Phụ lục)
    // A-2026-01-02-01 (5 đoạn) -> level 2 (Gói con)
    if (parts.length <= 3) {
      return { level: 0, parentId: null, group };
    } else {
      const level = parts.length - 3;
      const parentId = parts.slice(0, -1).join('-');
      return { level, parentId, group };
    }
  }

  /**
   * Sinh mã ContractID tự động theo quy tắc chuẩn:
   * - Hợp đồng (Cấp 0 / Root): Nhóm + Năm + Số thứ tự 2 chữ số (vd: A-2026-01, B-2026-03, B-2026-59)
   * - Phụ lục / Gói con (Cấp 1, 2, ...): Thêm -XX vào sau mã của Cha (vd: A-2026-01-08, A-2026-01-06-01)
   */
  static generateNextContractID(parentContract, allContracts = [], group = 'A') {
    const year = new Date().getFullYear();
    if (!parentContract) {
      // Hợp đồng cấp 0 (Root)
      const rootContracts = allContracts.filter((c) => !c.parent_id && c.group === group);
      
      // Quét các ID cùng năm hoặc cùng nhóm để tìm số thứ tự lớn nhất
      let maxSameYearIdx = 0;
      let maxAnyIdx = 0;
      
      const sameYearRegex = new RegExp(`^${group}-${year}-(\\d+)$`);
      const anyRootRegex = new RegExp(`^${group}-\\d{4}-(\\d+)$`);

      rootContracts.forEach((c) => {
        if (!c.id) return;
        const sameMatch = c.id.match(sameYearRegex);
        if (sameMatch) {
          const idx = parseInt(sameMatch[1], 10);
          if (!isNaN(idx) && idx > maxSameYearIdx) maxSameYearIdx = idx;
        }
        const anyMatch = c.id.match(anyRootRegex);
        if (anyMatch) {
          const idx = parseInt(anyMatch[1], 10);
          if (!isNaN(idx) && idx > maxAnyIdx) maxAnyIdx = idx;
        }
      });

      let nextIndex = 1;
      if (maxSameYearIdx > 0) {
        nextIndex = maxSameYearIdx + 1;
      } else if (maxAnyIdx > 0) {
        nextIndex = Math.max(maxAnyIdx + 1, rootContracts.length + 1);
      } else {
        nextIndex = rootContracts.length + 1;
      }

      return `${group}-${year}-${String(nextIndex).padStart(2, '0')}`;
    } else {
      // Phụ lục / Gói con (Cấp 1, 2, ...) của parentContract
      const parentId = parentContract.id;
      const prefix = `${parentId}-`;
      const directChildren = allContracts.filter((c) => c.parent_id === parentId || (c.id && c.id.startsWith(prefix)));
      
      let maxChildIdx = 0;
      directChildren.forEach((c) => {
        if (!c.id || !c.id.startsWith(prefix)) return;
        const tail = c.id.substring(prefix.length);
        // Chỉ lấy số trực tiếp (ví dụ "01" hoặc "08" không chứa dấu gạch ngang tiếp theo)
        const firstSegment = tail.split('-')[0];
        const num = parseInt(firstSegment, 10);
        if (!isNaN(num) && num > maxChildIdx) {
          maxChildIdx = num;
        }
      });

      const nextChildIndex = Math.max(maxChildIdx + 1, directChildren.length + 1);
      return `${parentId}-${String(nextChildIndex).padStart(2, '0')}`;
    }
  }

  /**
   * Đọc file .cpms / .xlsb / .xlsx tải lên từ máy tính
   */
  static async importCPMSFile(file) {
    if (!window.XLSX) {
      throw new Error('Thư viện xử lý Excel (SheetJS) chưa được nạp.');
    }

    const arrayBuffer = await file.arrayBuffer();
    // Đọc số serial Excel nguyên bản (cellDates: false) để tránh lệch 1 ngày do múi giờ UTC+7
    const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: false });

    console.log('Các sheet tìm thấy trong file .cpms:', workbook.SheetNames);

    // 1. Quét thông tin dự án từ các sheet có thể chứa cài đặt / thông tin dự án (PROJECT_SETTINGS, SYS_INFO, SETTINGS, v.v.)
    let projectId = '';
    let projectName = '';
    let clientName = '';
    let location = '';
    let contractor = '';

    // Thu thập các sheet theo thứ tự ưu tiên
    const candidateSheets = [];
    workbook.SheetNames.forEach((s) => {
      const lower = s.toLowerCase().trim();
      if (
        lower.includes('project_setting') ||
        lower.includes('project setting') ||
        lower.includes('project_info') ||
        lower.includes('project info') ||
        lower.includes('thông tin dự án') ||
        lower.includes('thong tin du an') ||
        lower.includes('thông tin công trình')
      ) {
        candidateSheets.unshift(s); // Ưu tiên hàng đầu
      } else if (
        lower === 'sys_info' ||
        lower === 'sys info' ||
        lower === 'sysinfo' ||
        lower === 'settings' ||
        lower === 'setting' ||
        lower === 'cpms' ||
        lower === 'thông tin' ||
        lower.includes('info')
      ) {
        candidateSheets.push(s); // Ưu tiên tiếp theo
      }
    });

    for (const sheetName of candidateSheets) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) continue;
      const data = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!data || data.length === 0) continue;

      // Quét 1: Dạng Key - Value 2 cột dọc (Row [Key, Value])
      for (const row of data) {
        if (!row || row.length < 2) continue;
        const key = String(row[0] || '').toLowerCase().trim();
        const rawVal = row[1];
        if (rawVal === undefined || rawVal === null) continue;
        const val = String(rawVal).trim();
        if (!val) continue;

        if (!projectId && (key.includes('projectid') || key.includes('mã dự án') || key === 'proj' || key.includes('mã da'))) {
          projectId = val;
        }
        if (!projectName && (key.includes('projectname') || key.includes('tên dự án') || key === 'proj_name' || key.includes('tên công trình') || key === 'project_name')) {
          projectName = val;
        }
        if (!clientName && (key.includes('client') || key.includes('chủ đầu tư') || key.includes('bên a'))) {
          clientName = val;
        }
        if (!location && (key.includes('location') || key.includes('địa điểm') || key.includes('địa chỉ') || key.includes('dia diem'))) {
          location = val;
        }
        if (!contractor && (key.includes('contractor') || key.includes('tổng thầu') || key.includes('nhà thầu') || key.includes('bên b') || key === 'maincon')) {
          contractor = val;
        }
      }

      // Quét 2: Dạng bảng ngang (Row 0 là Headers, Row 1 là Values)
      if (data.length >= 2 && Array.isArray(data[0])) {
        const headers = data[0].map((h) => String(h || '').toLowerCase().trim());
        const values = data[1] || [];
        headers.forEach((h, idx) => {
          const rawVal = values[idx];
          if (rawVal === undefined || rawVal === null) return;
          const val = String(rawVal).trim();
          if (!val) return;

          if (!projectId && (h.includes('projectid') || h.includes('mã dự án') || h === 'proj' || h.includes('mã da'))) {
            projectId = val;
          }
          if (!projectName && (h.includes('projectname') || h.includes('tên dự án') || h === 'proj_name' || h.includes('tên công trình') || h === 'project_name')) {
            projectName = val;
          }
          if (!clientName && (h.includes('client') || h.includes('chủ đầu tư') || h.includes('bên a'))) {
            clientName = val;
          }
          if (!location && (h.includes('location') || h.includes('địa điểm') || h.includes('địa chỉ') || h.includes('dia diem'))) {
            location = val;
          }
          if (!contractor && (h.includes('contractor') || h.includes('tổng thầu') || h.includes('nhà thầu') || h.includes('bên b') || h === 'maincon')) {
            contractor = val;
          }
        });
      }

      // Nếu đã trích xuất được cả projectName và location thì dừng quét
      if (projectName && location) break;
    }

    // Quét bổ sung: Nếu vẫn chưa có projectName hoặc location, quét tìm trong toàn bộ các ô cell của các candidateSheets
    if (!projectName || !location) {
      for (const sheetName of candidateSheets) {
        const ws = workbook.Sheets[sheetName];
        if (!ws) continue;
        const data = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
        for (let r = 0; r < data.length; r++) {
          const row = data[r];
          if (!row || !Array.isArray(row)) continue;
          for (let c = 0; c < row.length; c++) {
            const cellText = String(row[c] || '').toLowerCase().trim();
            if (!projectName && (cellText === 'tên dự án:' || cellText === 'tên dự án' || cellText === 'dự án:' || cellText === 'tên công trình:')) {
              const nextVal = row[c + 1] || (data[r + 1] ? data[r + 1][c] : '');
              if (nextVal && String(nextVal).trim()) projectName = String(nextVal).trim();
            }
            if (!location && (cellText === 'địa điểm:' || cellText === 'địa điểm' || cellText === 'địa chỉ:' || cellText === 'địa chỉ')) {
              const nextVal = row[c + 1] || (data[r + 1] ? data[r + 1][c] : '');
              if (nextVal && String(nextVal).trim()) location = String(nextVal).trim();
            }
          }
        }
        if (projectName && location) break;
      }
    }

    // Fallback: Chỉ dùng tên file khi thực sự không tìm thấy tên dự án trong bất kỳ sheet nào
    if (!projectId) projectId = 'PRJ_' + Date.now();
    if (!projectName) projectName = file.name.replace(/\.[^/.]+$/, '');
    if (!clientName) clientName = 'Chủ đầu tư';
    if (!location) location = '';

    // 2. CONTRACT_DB
    const contracts = [];
    if (workbook.SheetNames.includes('CONTRACT_DB')) {
      const ws = workbook.Sheets['CONTRACT_DB'];
      const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
      let currentGroup = 'A';

      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length === 0) continue;
        const c1 = String(r[0] || '').trim();

        if (c1 === 'GROUP-A') {
          currentGroup = 'A';
          continue;
        }
        if (c1 === 'GROUP-B') {
          currentGroup = 'B';
          continue;
        }
        if (!c1 || c1.startsWith('temp') || c1.toLowerCase().includes('contractid')) continue;

        const hierarchy = this.parseContractID(c1);

        const rawZ = r[25] ? String(r[25]).trim() : '';
        let codeA = '';
        let codeB = '';
        let codeB2 = '';
        if (rawZ) {
          const parts = rawZ.split('|');
          parts.forEach((p) => {
            const kv = p.split('=');
            if (kv.length >= 2) {
              const k = kv[0].trim().toUpperCase();
              const v = kv[1].trim();
              if (k === 'SUBCON') codeB = v;
              if (k === 'SUBCON2') codeB2 = v;
              if (k === 'CLIENT' || k === 'MAINCON') codeA = v;
            } else if (!codeB) {
              codeB = p.trim();
            }
          });
        }

        contracts.push({
          id: c1,
          project_id: projectId,
          group: hierarchy.group || currentGroup,
          parent_id: hierarchy.parentId,
          level: hierarchy.level,
          contract_no: String(r[1] || c1),
          contract_type: String(r[2] || (hierarchy.level === 0 ? 'MAIN' : 'Appendix')),
          signing_date: this.parseExcelDate(r[3]),
          end_date: this.parseExcelDate(r[4]),
          status: String(r[5] || 'Active'),
          company_a: String(r[6] || ''),
          company_b: String(r[7] || ''),
          company_b2: String(r[8] || ''),
          rep_b2: String(r[9] || ''),
          scope_of_work: String(r[10] || ''),
          exchange_rate: parseFloat(r[12] || 1.0) || 1.0,
          estimated_val: parseFloat(r[13] || 0) || 0,
          contract_val: parseFloat(r[14] || 0) || 0,
          rep_a: String(r[15] || ''),
          rep_b: String(r[16] || ''),
          category: String(r[17] || ''),
          payee: String(r[18] || ''),
          bank_account: String(r[19] || ''),
          bank_name: String(r[20] || ''),
          contact_person_b: String(r[21] || ''),
          created_at: r[22] || null,
          display_no: String(r[23] || ''),
          sort_order: parseFloat(r[24] || 0) || 0,
          code_a: codeA,
          code_b: codeB,
          code_b2: codeB2,
          company_code: codeB || rawZ,
          tax_code: r[26] ? String(r[26]).trim() : '',
          linked_contract_id: r[27] ? String(r[27]).trim() : '',
          tax_code_a: r[28] ? String(r[28]).trim() : '',
          tax_code_b2: r[29] ? String(r[29]).trim() : '',
          is_additive: r[30] !== undefined && r[30] !== null && String(r[30]).trim() !== ''
            ? (String(r[30]).trim() === '1' || String(r[30]).trim().toLowerCase() === 'true')
            : (!hierarchy.parentId),
          is_provisional_val: r[31] !== undefined && r[31] !== null && String(r[31]).trim() !== ''
            ? (String(r[31]).trim() === '1' || String(r[31]).trim().toLowerCase() === 'true')
            : false,
        });
      }
    }

    // 3. PAYMENT_LEDGER
    const payments = [];
    if (workbook.SheetNames.includes('PAYMENT_LEDGER')) {
      const ws = workbook.Sheets['PAYMENT_LEDGER'];
      const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1 });

      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length === 0) continue;
        const contractId = String(r[2] || '').trim();
        if (!contractId || contractId.startsWith('GROUP') || contractId.toLowerCase().includes('contractid')) continue;

        const rawLedgerId = r[1];
        const isHeader = rawLedgerId === undefined || rawLedgerId === null || String(rawLedgerId).trim() === '';
        const ledgerId = isHeader ? `HDR-${i}` : String(rawLedgerId).trim();

        const reqAmt = parseFloat(r[11] || 0) || 0;
        const paidAmt = parseFloat(r[12] || 0) || 0;

        const rawItemName = String(r[5] || '').trim();
        const rawPaymentRef = String(r[14] || '').trim();
        const rawNote = String(r[20] || '').trim();
        const finalContent = rawPaymentRef || rawItemName || rawNote || '';

        payments.push({
          id: ledgerId,
          is_header: isHeader,
          project_id: projectId,
          parent_id: String(r[0] || '').trim(),
          contract_id: contractId,
          contract_no: String(r[3] || ''),
          item_index: parseFloat(r[4] || 0) || 0,
          item_name: isHeader ? (rawItemName || finalContent) : (rawPaymentRef || rawItemName || finalContent),
          exchange_rate: parseFloat(r[6] || 1.0) || 1.0,
          contract_val: parseFloat(r[8] || 0) || 0,
          actual_volume: parseFloat(r[9] || 0) || 0,
          accepted_vol: parseFloat(r[10] || 0) || 0,
          request_amount: reqAmt,
          paid_amount: paidAmt,
          remaining: parseFloat(r[13] || reqAmt - paidAmt) || 0,
          payment_ref: rawPaymentRef || rawItemName || finalContent,
          payment_no: String(r[15] || ''),
          payment_date: r[16] || null,
          invoice_value: parseFloat(r[17] || 0) || 0,
          invoice_ref: String(r[18] || ''),
          invoice_date: r[19] || null,
          note: rawNote,
        });
      }

      // Tự động trích xuất các mục Chi phí Ban (CPB - Chi phí quản lý không có HĐ) từ PAYMENT_LEDGER
      const existingCpbIds = new Set(contracts.map((c) => c.id));
      const cpbPayments = payments.filter((p) => p.contract_id && /^B-CPB/i.test(p.contract_id));
      const uniqueCpbIds = [...new Set(cpbPayments.map((p) => p.contract_id))];

      uniqueCpbIds.forEach((cpbId) => {
        if (!existingCpbIds.has(cpbId)) {
          const headerRow = payments.find((p) => p.contract_id === cpbId && p.is_header);
          const firstRow = payments.find((p) => p.contract_id === cpbId);
          const cpbName = (headerRow && headerRow.item_name) || (firstRow && firstRow.contract_no) || 'Chi Phí Ban';

          contracts.push({
            id: cpbId,
            project_id: projectId,
            group: 'B',
            parent_id: null,
            level: 0,
            contract_no: cpbId,
            contract_type: 'CPB',
            signing_date: null,
            end_date: null,
            status: 'Active',
            company_a: clientName || 'Chủ đầu tư',
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
          });
          existingCpbIds.add(cpbId);
        }
      });
    }

    // 4. DOCUMENT_TRACKER
    const documents = [];
    if (workbook.SheetNames.includes('DOCUMENT_TRACKER')) {
      const ws = workbook.Sheets['DOCUMENT_TRACKER'];
      const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1 });

      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length === 0) continue;
        const docId = String(r[0] || '').trim();
        if (!docId || docId.toLowerCase().includes('documentid')) continue;

        const rawLink = String(r[5] || '').trim();
        const isOnlineDrive = rawLink.startsWith('http://') || rawLink.startsWith('https://');
        const extraDriveLink = String(r[20] || '').trim();
        const extraRelPath = String(r[21] || '').trim();

        documents.push({
          id: docId,
          project_id: projectId,
          contract_id: String(r[1] || ''),
          category: String(r[2] || 'Hồ sơ'),
          doc_no: String(r[3] || ''),
          title: String(r[4] || ''),
          file_link: rawLink,
          local_link: isOnlineDrive ? '' : rawLink,
          drive_link: extraDriveLink || (isOnlineDrive ? rawLink : ''),
          relative_path: extraRelPath,
          start_date: r[6] || null,
          due_date: r[7] || null,
          received_date: r[8] || null,
          return_date: r[9] || null,
          recipient: String(r[10] || ''),
          quantity: parseFloat(r[11] || 1) || 1,
          supporting_docs: String(r[12] || ''),
          lan_val: r[13] || 1,
          status: String(r[14] || 'Approved'),
          notes: String(r[15] || ''),
          issuer: String(r[16] || ''),
          recipients: String(r[17] || ''),
          sub_type1: String(r[18] || ''),
          sub_type2: String(r[19] || ''),
        });
      }
    }

    const projectData = {
      id: projectId,
      code: projectId,
      name: projectName,
      client: clientName,
      location: location,
      contractor: contractor,
      source_filename: file.name,
      source_file_size: file.size,
      source_file_updated: new Date().toISOString(),
      created_at: new Date().toISOString(),
      total_contracts: contracts.length,
      total_payments: payments.length,
      total_docs: documents.length,
    };

    try {
      localStorage.setItem('maio_active_local_filename', file.name);
    } catch (_) {}

    // Lưu vào IndexedDB
    await db.put('projects', projectData);
    for (const c of contracts) await db.put('contracts', c);
    for (const p of payments) await db.put('payments', p);
    for (const d of documents) await db.put('documents', d);

    return { projectData, contractsCount: contracts.length, paymentsCount: payments.length, docsCount: documents.length };
  }

  /**
   * Xuất toàn bộ dữ liệu dự án ra file tương thích chuẩn 100% với CPMS Plus
   * @param {string} projectId Mã dự án
   * @param {string} extension Định dạng xuất: 'cpms' (để nạp thẳng vào CPMS Plus) hoặc 'xlsx' (để mở trong Excel)
   */
  /**
   * Tạo toàn bộ cấu trúc Workbook CPMS chuẩn 5 sheet
   */
  static async buildCPMSWorkbook(projectId = 'CHC1') {
    if (!window.XLSX) {
      throw new Error('Thư viện XLSX chưa sẵn sàng.');
    }

    const allContracts = (await db.getAll('contracts')).filter((c) => c.project_id === projectId);
    const allPayments = (await db.getAll('payments')).filter((p) => p.project_id === projectId);
    const allDocs = (await db.getAll('documents')).filter((d) => d.project_id === projectId);
    const project = (await db.getAll('projects')).find((p) => p.id === projectId) || { id: projectId, name: projectId };

    const wb = window.XLSX.utils.book_new();

    // 1. Sheet PROJECT_SETTINGS (Khớp chính xác với cơ chế đọc metadata của CPMS Plus)
    const prjCode = project.id || project.code || 'CHC1';
    const prjName = project.name || 'Dự án ' + prjCode;
    const prjLoc = project.location || '';
    const prjClient = project.client || '';

    // Thu thập danh sách mã công ty trong dự án để lưu vào PROJECT_SETTINGS chuẩn CPMS Plus
    const allCompanyCodesSet = new Set(['CLI', 'VEC', 'VNCN', 'CTC', 'NSH', 'KD', 'TID', 'HHG', 'HD', 'ART']);
    let prjClientCode = 'CLI';
    let prjMainConCode = 'VEC';

    allContracts.forEach((c) => {
      if (c.code_a) allCompanyCodesSet.add(c.code_a.toUpperCase());
      if (c.code_b) allCompanyCodesSet.add(c.code_b.toUpperCase());
      if (c.code_b2) allCompanyCodesSet.add(c.code_b2.toUpperCase());
      if (c.group === 'A' && c.code_a) prjClientCode = c.code_a.toUpperCase();
      if (c.group === 'A' && c.code_b) prjMainConCode = c.code_b.toUpperCase();
    });

    const prjSettingsData = [
      ['Key', 'Value'],
      ['PROJ', prjCode],
      ['CLIENT', prjClientCode],
      ['MAINCON', prjMainConCode],
      ['SUPCON', 'ART'],
      ['COMPANY_CODES', Array.from(allCompanyCodesSet).join(';')],
      ['ProjectID', prjCode],
      ['ProjectName', prjName],
      ['Location', prjLoc],
      ['Client', prjClient],
      ['Contractor', project.contractor || ''],
      ['CreatedDate', project.created_at || new Date().toISOString().slice(0, 10)],
      ['LastOpened', new Date().toISOString().slice(0, 10)],
      ['DataVersion', '1.0'],
      ['ExportDate', new Date().toISOString()],
    ];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(prjSettingsData), 'PROJECT_SETTINGS');

    // 2. Sheet CONTRACT_DB (25 cột gốc chuẩn CPMS Plus + các cột mở rộng ở ngoài cùng bên phải)
    const contractHeaders = [
      'ContractID', 'ContractNo.', 'ContractType', 'SigningDate', 'EndDate', 'Status',
      'Company1', 'Company2', 'Company3', 'RepB2', 'ScopeOfWork', 'temp2', 'ExchangeRate',
      'EstimatedValue', 'ContractValue', 'RepA', 'RepB', 'Categories', 'Payee', 'BankAccountNo.',
      'BankName', 'ContactPersonB', 'CreatedDate', 'DisplayNo', 'SortOrder',
      'CompanyCode', 'TaxCode', 'LinkedContractID', 'TaxCodeA', 'TaxCodeB2', 'IsAdditive', 'IsProvisionalVal'
    ];
    const contractRows = [contractHeaders];

    const formatCompanyCodeCell = (c) => {
      const codeA = c.code_a ? String(c.code_a).trim().toUpperCase() : '';
      const codeB = c.code_b ? String(c.code_b).trim().toUpperCase() : (c.company_code ? String(c.company_code).trim().toUpperCase() : '');
      const codeB2 = c.code_b2 ? String(c.code_b2).trim().toUpperCase() : '';

      let parts = [];
      if (codeA) parts.push(`CLIENT=${codeA}`);
      if (codeB) parts.push(`SUBCON=${codeB}`);
      if (codeB2) parts.push(`SUBCON2=${codeB2}`);

      if (parts.length > 0) return parts.join('|');
      return c.company_code || '';
    };

    // Nhóm A
    contractRows.push(['GROUP-A', 'Hợp đồng với chủ đầu tư / Tổng thầu']);
    allContracts.filter(c => c.group === 'A').forEach(c => {
      contractRows.push([
        c.id, c.contract_no, c.contract_type || (c.parent_id ? 'Appendix' : 'MAIN'),
        this.dateToExcelCell(c.signing_date), this.dateToExcelCell(c.end_date), c.status || 'Active',
        c.company_a || '', c.company_b || '', c.company_b2 || '', c.rep_b2 || '', c.scope_of_work || '', '', c.exchange_rate || 1.0,
        c.estimated_val || 0, c.contract_val || 0, c.rep_a || '', c.rep_b || '', c.category || '', c.payee || '',
        c.bank_account || '', c.bank_name || '', c.contact_person_b || '', c.created_at || '', c.display_no || '', c.sort_order || 0,
        formatCompanyCodeCell(c), c.tax_code || '', c.linked_contract_id || '', c.tax_code_a || '', c.tax_code_b2 || '',
        (c.is_additive === true || c.is_additive === 1 || c.is_additive === '1' || (!c.parent_id && c.is_additive !== false)) ? '1' : '0',
        (c.is_provisional_val === true || c.is_provisional_val === 1 || c.is_provisional_val === '1') ? '1' : '0'
      ]);
    });

    // Nhóm B (Chỉ xuất các hợp đồng thực tế, CPB được quản lý trong PAYMENT_LEDGER chuẩn CPMS Plus)
    contractRows.push(['GROUP-B', 'Hợp đồng với thầu phụ / Nhà cung cấp']);
    allContracts.filter(c => c.group === 'B' && !c.is_cpb && c.contract_type !== 'CPB').forEach(c => {
      contractRows.push([
        c.id, c.contract_no, c.contract_type || (c.parent_id ? 'Appendix' : 'MAIN'),
        this.dateToExcelCell(c.signing_date), this.dateToExcelCell(c.end_date), c.status || 'Active',
        c.company_a || '', c.company_b || '', c.company_b2 || '', c.rep_b2 || '', c.scope_of_work || '', '', c.exchange_rate || 1.0,
        c.estimated_val || 0, c.contract_val || 0, c.rep_a || '', c.rep_b || '', c.category || '', c.payee || '',
        c.bank_account || '', c.bank_name || '', c.contact_person_b || '', c.created_at || '', c.display_no || '', c.sort_order || 0,
        formatCompanyCodeCell(c), c.tax_code || '', c.linked_contract_id || '', c.tax_code_a || '', c.tax_code_b2 || '',
        (c.is_additive === true || c.is_additive === 1 || c.is_additive === '1' || (!c.parent_id && c.is_additive !== false)) ? '1' : '0',
        (c.is_provisional_val === true || c.is_provisional_val === 1 || c.is_provisional_val === '1') ? '1' : '0'
      ]);
    });

    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(contractRows), 'CONTRACT_DB');

    // 3. Sheet PAYMENT_LEDGER (26 cột A -> Z với ĐẦY ĐỦ CÔNG THỨC EXCEL)
    // Đảm bảo mọi mục CPB đều có dòng Header trong PAYMENT_LEDGER
    const cpbContracts = allContracts.filter(c => c.is_cpb || c.contract_type === 'CPB' || c.id.startsWith('B-CPB-'));
    cpbContracts.forEach(cpb => {
      const hasHeader = allPayments.some(p => p.is_header && p.contract_id === cpb.id);
      if (!hasHeader) {
        allPayments.push({
          id: `HDR-${cpb.id}`,
          is_header: true,
          project_id: projectId,
          parent_id: 'GROUP-B',
          contract_id: cpb.id,
          contract_no: cpb.contract_no || cpb.id,
          item_name: cpb.scope_of_work || 'Chi Phí Ban',
          exchange_rate: 1.0,
          contract_val: cpb.contract_val || 0,
          actual_volume: 0,
          accepted_vol: 0,
          request_amount: 0,
          paid_amount: 0,
          remaining: 0,
        });
      }
    });

    const payHeaders = [
      'RowID', 'LedgerID', 'ContractID', 'ContractNo.', 'ItemIndex', 'ItemName',
      'ExchangeRate', 'EstimatedValue', 'ContractValue', 'ActualVolume', 'AcceptedVolume',
      'RequestAmount', 'PaidAmount', 'Remaining', 'PaymentRef', 'PaymentNo.', 'PaymentDate',
      'InvoiceValue', 'InvoiceRef', 'InvoiceDate', 'Note', 'CreatedDate',
      'ActualVND', 'AcceptedVND', 'RequestVND', 'PaidVND'
    ];
    const payRows = [payHeaders];

    allPayments.forEach((p) => {
      const rowNum = payRows.length + 1; // 1-based index trong Excel
      const isHeader = p.is_header || !p.id || String(p.id).startsWith('HDR-');
      const isGroupHeader = p.contract_id === 'GROUP-A' || p.contract_id === 'GROUP-B';

      if (isHeader) {
        // DÒNG HEADER: ÁP DỤNG CÔNG THỨC TỔNG HỢP VÀ SUMIFS
        const cellH = isGroupHeader
          ? { t: 'n', f: `SUMIFS(H:H,A:A,C${rowNum})`, v: p.estimated_val || 0 }
          : (p.estimated_val || 0);

        const cellI = isGroupHeader
          ? { t: 'n', f: `SUMIFS(I:I,A:A,C${rowNum})`, v: p.contract_val || 0 }
          : (p.contract_val || 0);

        const cellJ = { t: 'n', f: `W${rowNum}`, v: p.actual_volume || 0 };
        const cellK = { t: 'n', f: `X${rowNum}`, v: p.accepted_vol || 0 };
        const cellL = { t: 'n', f: `Y${rowNum}`, v: p.request_amount || 0 };
        const cellM = { t: 'n', f: `Z${rowNum}`, v: p.paid_amount || 0 };
        const cellN = { t: 'n', f: `L${rowNum}-M${rowNum}`, v: (p.request_amount || 0) - (p.paid_amount || 0) };

        const cellW = { t: 'n', f: `SUMIFS($W:$W,$A:$A,C${rowNum})`, v: p.actual_volume || 0 };
        const cellX = { t: 'n', f: `SUMIFS($X:$X,$A:$A,C${rowNum})`, v: p.accepted_vol || 0 };
        const cellY = { t: 'n', f: `SUMIFS($Y:$Y,$A:$A,C${rowNum})`, v: p.request_amount || 0 };
        const cellZ = { t: 'n', f: `SUMIFS($Z:$Z,$A:$A,C${rowNum})`, v: p.paid_amount || 0 };

        payRows.push([
          p.parent_id || '',                    // Col A: RowID
          '',                                   // Col B: LedgerID (Rỗng để nhận diện là dòng Header)
          p.contract_id,                        // Col C: ContractID
          p.contract_no || '',                  // Col D: ContractNo.
          p.item_index || '',                   // Col E: ItemIndex
          p.item_name || '',                    // Col F: ItemName
          p.exchange_rate || 1.0,               // Col G: ExchangeRate
          cellH,                                // Col H: EstimatedValue (SUMIFS nếu là Group)
          cellI,                                // Col I: ContractValue (SUMIFS nếu là Group)
          cellJ,                                // Col J: ActualVolume =W
          cellK,                                // Col K: AcceptedVolume =X
          cellL,                                // Col L: RequestAmount =Y
          cellM,                                // Col M: PaidAmount =Z
          cellN,                                // Col N: Remaining =L-M
          p.payment_ref || '',                  // Col O: PaymentRef
          p.payment_no || '',                   // Col P: PaymentNo.
          p.payment_date || '',                 // Col Q: PaymentDate
          p.invoice_value || 0,                 // Col R: InvoiceValue
          p.invoice_ref || '',                  // Col S: InvoiceRef
          p.invoice_date || '',                 // Col T: InvoiceDate
          p.note || '',                         // Col U: Note
          p.created_at || '',                   // Col V: CreatedDate
          cellW,                                // Col W: ActualVND =SUMIFS
          cellX,                                // Col X: AcceptedVND =SUMIFS
          cellY,                                // Col Y: RequestVND =SUMIFS
          cellZ,                                // Col Z: PaidVND =SUMIFS
        ]);
      } else {
        // DÒNG GIAO DỊCH THỰC TẾ (DETAIL PAYMENT ROW): ÁP DỤNG CÔNG THỨC TRỪ VÀ NHÂN TỶ GIÁ
        const rate = p.exchange_rate || 1.0;
        const actVol = p.actual_volume || 0;
        const accVol = p.accepted_vol || 0;
        const reqAmt = p.request_amount || 0;
        const paidAmt = p.paid_amount || 0;

        const cellN = { t: 'n', f: `L${rowNum}-M${rowNum}`, v: reqAmt - paidAmt };
        const cellW = { t: 'n', f: `$G${rowNum}*J${rowNum}`, v: rate * actVol };
        const cellX = { t: 'n', f: `$G${rowNum}*K${rowNum}`, v: rate * accVol };
        const cellY = { t: 'n', f: `$G${rowNum}*L${rowNum}`, v: rate * reqAmt };
        const cellZ = { t: 'n', f: `$G${rowNum}*M${rowNum}`, v: rate * paidAmt };

        payRows.push([
          p.parent_id || p.contract_id,         // Col A: RowID (Mã cha hoặc mã hợp đồng)
          p.id,                                 // Col B: LedgerID (Mã LG-...)
          p.contract_id,                        // Col C: ContractID
          p.contract_no || '',                  // Col D: ContractNo.
          p.item_index || 0,                    // Col E: ItemIndex
          p.item_name || '',                    // Col F: ItemName
          rate,                                 // Col G: ExchangeRate
          0,                                    // Col H: EstimatedValue
          p.contract_val || 0,                  // Col I: ContractValue
          actVol,                               // Col J: ActualVolume
          accVol,                               // Col K: AcceptedVolume
          reqAmt,                               // Col L: RequestAmount
          paidAmt,                              // Col M: PaidAmount
          cellN,                                // Col N: Remaining =L-M
          p.payment_ref || '',                  // Col O: PaymentRef
          p.payment_no || '',                   // Col P: PaymentNo.
          p.payment_date || '',                 // Col Q: PaymentDate
          p.invoice_value || 0,                 // Col R: InvoiceValue
          p.invoice_ref || '',                  // Col S: InvoiceRef
          p.invoice_date || '',                 // Col T: InvoiceDate
          p.note || '',                         // Col U: Note
          p.created_at || '',                   // Col V: CreatedDate
          cellW,                                // Col W: ActualVND =$G*J
          cellX,                                // Col X: AcceptedVND =$G*K
          cellY,                                // Col Y: RequestVND =$G*L
          cellZ,                                // Col Z: PaidVND =$G*M
        ]);
      }
    });

    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(payRows), 'PAYMENT_LEDGER');

    // 4. Sheet DOCUMENT_TRACKER (20 cột chuẩn CPMS Plus + cột mở rộng DriveLink, RelativePath ở cuối)
    const docHeaders = [
      'DocumentID', 'ContractID', 'Hồ sơ', 'Số', 'Tên hồ sơ', 'Link', 'Từ ngày',
      'Đến ngày', 'Tiếp nhận', 'Trả lại', 'Người nhận lại', 'Số lượng', 'Hồ sơ kèm theo',
      'Lần', 'Tình trạng', 'Ghi chú', 'Issuer', 'Recipients', 'SubType1', 'SubType2',
      'DriveLink', 'RelativePath'
    ];
    const docRows = [docHeaders];
    allDocs.forEach(d => {
      docRows.push([
        d.id,
        d.contract_id,
        d.category || '',
        d.doc_no || '',
        d.title || '',
        d.local_link || d.file_link || d.drive_link || '',
        d.start_date || '',
        d.due_date || '',
        d.received_date || '',
        d.return_date || '',
        d.recipient || '',
        d.quantity || 1,
        d.supporting_docs || '',
        d.lan_val || 1,
        d.status || 'Approved',
        d.notes || '',
        d.issuer || '',
        d.recipients || '',
        d.sub_type1 || '',
        d.sub_type2 || '',
        d.drive_link || '',
        d.relative_path || ''
      ]);
    });
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(docRows), 'DOCUMENT_TRACKER');

    // 5. Sheet SYS_INFO (Chứa metadata hệ thống và thông tin dự án đồng bộ)
    const sysInfoData = [
      ['Key', 'Value'],
      ['ProjectID', prjCode],
      ['PROJ', prjCode],
      ['PROJ_NAME', prjName],
      ['ProjectName', prjName],
      ['Location', prjLoc],
      ['Client', prjClient],
      ['System', 'MAIO Web Management System'],
      ['Version', '2.0_Strict_Hierarchy'],
      ['Engine', 'CPMS Plus Compatible Importer/Exporter'],
      ['ExportedAt', new Date().toISOString()],
    ];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(sysInfoData), 'SYS_INFO');

    return { wb, prjCode };
  }

  /**
   * Tạo mảng nhị phân ArrayBuffer của file CPMS (để ghi trực tiếp vào FileSystemFileHandle)
   * Định dạng chuẩn bắt buộc là 'xlsb' (Excel Binary Workbook BIFF12) để tương thích 100% với CPMS Plus
   */
  static async buildCPMSBinaryBuffer(projectId = 'CHC1') {
    const { wb } = await this.buildCPMSWorkbook(projectId);
    const arrayBuffer = window.XLSX.write(wb, { bookType: 'xlsb', type: 'array' });
    return arrayBuffer;
  }

  static async exportToCPMSFormat(projectId = 'CHC1', extension = 'cpms') {
    const { wb, prjCode } = await this.buildCPMSWorkbook(projectId);

    // Tải về file theo đúng định dạng tên chuẩn của CPMS Plus: PROJ_YYYYMMDD_HHMMSS.cpms
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const timeStr = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    const safeProjCode = prjCode.replace(/[\\/:*?"<>|]/g, '_');
    const fileName = `${safeProjCode}_${timeStr}.${extension}`;

    // Khi xuất định dạng .cpms hoặc .xlsb, luôn xuất nhị phân BIFF12 (xlsb) tương thích CPMS Plus
    const bookType = (extension === 'cpms' || extension === 'xlsb') ? 'xlsb' : extension;
    window.XLSX.writeFile(wb, fileName, { bookType: bookType });
    return fileName;
  }
}
