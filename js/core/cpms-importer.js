/**
 * MAIO CPMS Importer & Exporter Engine
 * Đọc và xuất file dữ liệu định dạng .cpms (.xlsb / .xlsx) chuẩn xác 100% với phần mềm CPMS Plus.
 */

import { db } from './db.js';

export class CPMSImporterExporter {
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
   * - HĐ gốc: Nhóm + Năm + Số thứ tự (vd: A-2026-01, B-2026-03)
   * - Phụ lục / Gói con: Thêm -XX vào sau mã của Cha (vd: A-2026-01-02, A-2026-01-02-01)
   */
  static generateNextContractID(parentContract, allContracts, group = 'A') {
    const year = new Date().getFullYear();
    if (!parentContract) {
      // HĐ gốc
      const rootContracts = allContracts.filter((c) => !c.parent_id && c.group === group);
      const nextIndex = rootContracts.length + 1;
      return `${group}-${year}-${String(nextIndex).padStart(2, '0')}`;
    } else {
      // Con của parentContract
      const directChildren = allContracts.filter((c) => c.parent_id === parentContract.id);
      const nextIndex = directChildren.length + 1;
      return `${parentContract.id}-${String(nextIndex).padStart(2, '0')}`;
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
    // SheetJS hỗ trợ trực tiếp đọc nhị phân XLSB / XLSX qua ArrayBuffer
    const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

    console.log('Các sheet tìm thấy trong file .cpms:', workbook.SheetNames);

    // 1. Quét thông tin dự án từ PROJECT_SETTINGS, project info, SYS_INFO...
    let projectId = 'PRJ_' + Date.now();
    let projectName = file.name.replace(/\.[^/.]+$/, '');
    let clientName = 'Chủ đầu tư';
    let location = '';
    let contractor = '';

    const infoSheetName = workbook.SheetNames.find((s) => {
      const lower = s.toLowerCase().trim();
      return (
        lower === 'project_settings' ||
        lower === 'project info' ||
        lower === 'project_info' ||
        lower === 'projectinfo' ||
        lower === 'proj info' ||
        lower === 'proj_info' ||
        lower === 'projinfo' ||
        lower === 'thông tin dự án' ||
        lower === 'sys_info' ||
        (lower.includes('proj') && lower.includes('info'))
      );
    });

    if (infoSheetName) {
      const ws = workbook.Sheets[infoSheetName];
      const data = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
      // 1. Quét theo dạng Key - Value 2 cột dọc
      for (const row of data) {
        if (!row || row.length < 2) continue;
        const key = String(row[0]).toLowerCase().trim();
        const val = String(row[1]).trim();
        if (key.includes('projectid') || key.includes('mã dự án') || key === 'proj') projectId = val;
        if (key.includes('projectname') || key.includes('tên dự án') || key === 'proj_name' || key === 'name') projectName = val;
        if (key.includes('client') || key.includes('chủ đầu tư') || key.includes('bên a')) clientName = val;
        if (key.includes('location') || key.includes('địa điểm') || key.includes('địa chỉ')) location = val;
        if (key.includes('contractor') || key.includes('tổng thầu') || key.includes('nhà thầu') || key.includes('bên b')) contractor = val;
      }
      // 2. Quét theo dạng bảng ngang (Row 0 là Headers, Row 1 là Values)
      if (data.length >= 2 && Array.isArray(data[0])) {
        const headers = data[0].map((h) => String(h || '').toLowerCase().trim());
        const values = data[1] || [];
        headers.forEach((h, idx) => {
          const val = String(values[idx] || '').trim();
          if (!val) return;
          if (h.includes('projectid') || h.includes('mã dự án') || h === 'proj') projectId = val;
          if (h.includes('projectname') || h.includes('tên dự án') || h === 'proj_name') projectName = val;
          if (h.includes('client') || h.includes('chủ đầu tư') || h.includes('bên a')) clientName = val;
          if (h.includes('location') || h.includes('địa điểm') || h.includes('địa chỉ')) location = val;
          if (h.includes('contractor') || h.includes('tổng thầu') || h.includes('nhà thầu') || h.includes('bên b')) contractor = val;
        });
      }
    }

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

        contracts.push({
          id: c1,
          project_id: projectId,
          group: hierarchy.group || currentGroup,
          parent_id: hierarchy.parentId,
          level: hierarchy.level,
          contract_no: String(r[1] || c1),
          contract_type: String(r[2] || (hierarchy.level === 0 ? 'MAIN' : 'Appendix')),
          signing_date: r[3] || null,
          end_date: r[4] || null,
          status: String(r[5] || 'Active'),
          company_a: String(r[6] || ''),
          company_b: String(r[7] || ''),
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
          display_no: String(r[23] || ''),
          sort_order: parseFloat(r[24] || 0) || 0,
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

        payments.push({
          id: ledgerId,
          is_header: isHeader,
          project_id: projectId,
          parent_id: String(r[0] || '').trim(),
          contract_id: contractId,
          contract_no: String(r[3] || ''),
          item_index: parseFloat(r[4] || 0) || 0,
          item_name: String(r[5] || ''),
          exchange_rate: parseFloat(r[6] || 1.0) || 1.0,
          contract_val: parseFloat(r[8] || 0) || 0,
          actual_volume: parseFloat(r[9] || 0) || 0,
          accepted_vol: parseFloat(r[10] || 0) || 0,
          request_amount: reqAmt,
          paid_amount: paidAmt,
          remaining: parseFloat(r[13] || reqAmt - paidAmt) || 0,
          payment_ref: String(r[14] || ''),
          payment_no: String(r[15] || ''),
          payment_date: r[16] || null,
          invoice_value: parseFloat(r[17] || 0) || 0,
          invoice_ref: String(r[18] || ''),
          invoice_date: r[19] || null,
          note: String(r[20] || ''),
        });
      }
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

        documents.push({
          id: docId,
          project_id: projectId,
          contract_id: String(r[1] || ''),
          category: String(r[2] || 'Hồ sơ'),
          doc_no: String(r[3] || ''),
          title: String(r[4] || ''),
          file_link: String(r[5] || ''),
          start_date: r[6] || null,
          due_date: r[7] || null,
          received_date: r[8] || null,
          return_date: r[9] || null,
          recipient: String(r[10] || ''),
          quantity: parseFloat(r[11] || 1) || 1,
          status: String(r[14] || 'Approved'),
          notes: String(r[15] || ''),
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
      created_at: new Date().toISOString(),
      total_contracts: contracts.length,
      total_payments: payments.length,
      total_docs: documents.length,
    };

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
  static async exportToCPMSFormat(projectId = 'CHC1', extension = 'cpms') {
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

    const prjSettingsData = [
      ['SettingKey', 'SettingValue'],
      ['ProjectID', prjCode],
      ['PROJ', prjCode],
      ['PROJ_NAME', prjName],
      ['ProjectName', prjName],
      ['Location', prjLoc],
      ['Client', prjClient],
      ['CreatedDate', project.created_at || new Date().toISOString().slice(0, 10)],
      ['LastOpened', new Date().toISOString().slice(0, 10)],
      ['DataVersion', '1.0'],
      ['ExportDate', new Date().toISOString()],
    ];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(prjSettingsData), 'PROJECT_SETTINGS');

    // 2. Sheet CONTRACT_DB (25 cột chuẩn xác)
    const contractHeaders = [
      'ContractID', 'ContractNo.', 'ContractType', 'SigningDate', 'EndDate', 'Status',
      'Company1', 'Company2', 'Company3', 'RepB2', 'ScopeOfWork', 'temp2', 'ExchangeRate',
      'EstimatedValue', 'ContractValue', 'RepA', 'RepB', 'Categories', 'Payee', 'BankAccountNo.',
      'BankName', 'ContactPersonB', 'CreatedDate', 'DisplayNo', 'SortOrder'
    ];
    const contractRows = [contractHeaders];

    // Nhóm A
    contractRows.push(['GROUP-A', 'Hợp đồng với chủ đầu tư / Tổng thầu']);
    allContracts.filter(c => c.group === 'A').forEach(c => {
      contractRows.push([
        c.id, c.contract_no, c.contract_type || (c.parent_id ? 'Appendix' : 'MAIN'), c.signing_date, c.end_date, c.status || 'Active',
        c.company_a || '', c.company_b || '', '', '', c.scope_of_work || '', '', c.exchange_rate || 1.0,
        c.estimated_val || 0, c.contract_val || 0, c.rep_a || '', c.rep_b || '', c.category || '', c.payee || '',
        c.bank_account || '', c.bank_name || '', '', c.created_at || '', c.display_no || '', c.sort_order || 0
      ]);
    });

    // Nhóm B
    contractRows.push(['GROUP-B', 'Hợp đồng với thầu phụ / Nhà cung cấp']);
    allContracts.filter(c => c.group === 'B').forEach(c => {
      contractRows.push([
        c.id, c.contract_no, c.contract_type || (c.parent_id ? 'Appendix' : 'MAIN'), c.signing_date, c.end_date, c.status || 'Active',
        c.company_a || '', c.company_b || '', '', '', c.scope_of_work || '', '', c.exchange_rate || 1.0,
        c.estimated_val || 0, c.contract_val || 0, c.rep_a || '', c.rep_b || '', c.category || '', c.payee || '',
        c.bank_account || '', c.bank_name || '', '', c.created_at || '', c.display_no || '', c.sort_order || 0
      ]);
    });

    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(contractRows), 'CONTRACT_DB');

    // 3. Sheet PAYMENT_LEDGER (26 cột A -> Z với ĐẦY ĐỦ CÔNG THỨC EXCEL)
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

    // 4. Sheet DOCUMENT_TRACKER (16 cột chuẩn)
    const docHeaders = [
      'DocumentID', 'ContractID', 'Hồ sơ', 'Số', 'Tên hồ sơ', 'Link', 'Từ ngày',
      'Đến ngày', 'Tiếp nhận', 'Trả lại', 'Người nhận lại', 'Số lượng', 'Hồ sơ kèm theo',
      'Lần', 'Tình trạng', 'Ghi chú'
    ];
    const docRows = [docHeaders];
    allDocs.forEach(d => {
      docRows.push([
        d.id, d.contract_id, d.category || '', d.doc_no || '', d.title || '', d.file_link || '',
        d.start_date || '', d.due_date || '', d.received_date || '', d.return_date || '', d.recipient || '',
        d.quantity || 1, '', 1, d.status || 'Approved', d.notes || ''
      ]);
    });
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(docRows), 'DOCUMENT_TRACKER');

    // 5. Sheet SYS_INFO
    const sysInfoData = [
      ['System', 'MAIO Web Management System'],
      ['Version', '2.0_Strict_Hierarchy'],
      ['Engine', 'CPMS Plus Compatible Importer/Exporter'],
      ['ExportedAt', new Date().toISOString()],
    ];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(sysInfoData), 'SYS_INFO');

    // Tải về file theo đúng định dạng tên chuẩn của CPMS Plus: PROJ_YYYYMMDD_HHMMSS.cpms
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const timeStr = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    const safeProjCode = prjCode.replace(/[\\/:*?"<>|]/g, '_');
    const fileName = `${safeProjCode}_${timeStr}.${extension}`;

    window.XLSX.writeFile(wb, fileName, { bookType: extension === 'cpms' ? 'xlsx' : extension });
    return fileName;
  }
}
