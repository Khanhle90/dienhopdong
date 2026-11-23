/**
 * HỢP ĐỒNG TỰ ĐỘNG - CORE APPLICATION ENGINE v2.2
 * Cập nhật hoàn chỉnh:
 * 1. Tên người ký (CTV) IN ĐẬM 100% đúng mẫu
 * 2. Chia mỗi hợp đồng thành 2 trang A4 chuẩn (.contract-subpage-1 & .contract-subpage-2),
 *    Điều 1-2 ở Trang 1, Điều 3-6 & Chữ ký ở Trang 2 -> Không bao giờ bị cắt đôi chữ đầu/cuối trang
 * 3. Tích hợp thư viện jsPDF chuẩn, xuất PDF từng trang A4 chính xác không lỗi
 * 4. Ẩn toàn bộ thông báo Toast khi in, loại bỏ mọi chữ rác
 */

const state = {
  templateType: 'default',
  templateFileName: 'Mẫu hợp đồng CTV ABA.docx',
  templateArrayBuffer: null,
  detectedPlaceholders: ['Ngay', 'Thang', 'Nam', 'HoVaTen', 'Namsinh', 'Thuongtru', 'CCCD', 'Ngaycap', 'Noicap', 'SDT', 'Congviec', 'Thoihan', 'Thulao', 'CTV'],
  
  excelData: [],
  selectedSheetName: '',
  headers: [],
  
  isProcessing: false,
  processedContractsCount: 0,
  generatedContracts: []
};

// ==========================================
// 1. TIỆN ÍCH CHUẨN HÓA DỮ LIỆU
// ==========================================

function normalizeKey(str) {
  if (!str) return '';
  str = String(str).trim().toLowerCase();
  str = str.replace(/[àáạảãâầấậẩẫăằắặẳẵ]/g, 'a');
  str = str.replace(/[èéẹẻẽêềếệểễ]/g, 'e');
  str = str.replace(/[ìíịỉĩ]/g, 'i');
  str = str.replace(/[òóọỏõôồốộổỗơờớợởỡ]/g, 'o');
  str = str.replace(/[ùúụủũưừứựửữ]/g, 'u');
  str = str.replace(/[ỳýỵỷỹ]/g, 'y');
  str = str.replace(/[đ]/g, 'd');
  str = str.replace(/[^a-z0-9]/g, '');
  return str;
}

function base64ToArrayBuffer(base64) {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}

// ==========================================
// 2. KHỞI TẠO VÀ XỬ LÝ MẪU WORD (.DOCX)
// ==========================================

async function initDefaultTemplate() {
  try {
    if (typeof DEFAULT_TEMPLATE_B64 !== 'undefined' && DEFAULT_TEMPLATE_B64) {
      state.templateArrayBuffer = base64ToArrayBuffer(DEFAULT_TEMPLATE_B64);
    } else {
      const resp = await fetch('Mẫu hợp đồng.docx');
      state.templateArrayBuffer = await resp.arrayBuffer();
    }
    updateTemplateUI();
    scanPlaceholdersFromBuffer(state.templateArrayBuffer);
  } catch (err) {
    console.warn("Không thể tải mẫu mặc định tự động:", err);
  }
}

function updateTemplateUI() {
  const infoEl = document.getElementById('template-info');
  if (infoEl) {
    infoEl.innerHTML = `
      <div class="flex items-center justify-between p-3 bg-blue-50 border border-blue-200 rounded-lg">
        <div class="flex items-center space-x-3">
          <i class="fa-solid fa-file-word text-blue-600 text-2xl"></i>
          <div>
            <p class="font-semibold text-blue-900">${state.templateFileName}</p>
            <p class="text-xs text-blue-700">Trạng thái: ${state.templateType === 'default' ? 'Mẫu mặc định chuẩn ABA' : 'Mẫu tùy chỉnh người dùng'}</p>
          </div>
        </div>
        <button id="btn-download-template" onclick="downloadCurrentTemplate()" class="px-3 py-1.5 text-xs bg-white text-blue-700 font-medium rounded border border-blue-300 hover:bg-blue-100 transition shadow-sm">
          <i class="fa-solid fa-download mr-1"></i> Tải file mẫu (.docx)
        </button>
      </div>
    `;
  }
}

async function scanPlaceholdersFromBuffer(buffer) {
  try {
    if (typeof JSZip === 'undefined') return;
    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file("word/document.xml")?.async("string");
    if (!docXml) return;

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(docXml, "application/xml");
    const textNodes = xmlDoc.getElementsByTagName("w:t");
    let fullText = "";
    for (let i = 0; i < textNodes.length; i++) {
      fullText += textNodes[i].textContent;
    }

    const matches = fullText.match(/\{([^}]+)\}/g);
    if (matches) {
      const uniqueTags = [...new Set(matches.map(m => m.replace(/[{}]/g, '').trim()))];
      state.detectedPlaceholders = uniqueTags;
      renderPlaceholderBadges();
    }
  } catch (err) {
    console.error("Lỗi quét placeholder từ docx:", err);
  }
}

function renderPlaceholderBadges() {
  const container = document.getElementById('placeholder-badges');
  if (!container) return;
  container.innerHTML = state.detectedPlaceholders.map(tag => `
    <span class="inline-flex items-center px-2 py-1 rounded text-xs font-medium bg-slate-100 text-slate-700 border border-slate-200 mr-1 mb-1">
      <code class="text-blue-600 font-bold">{${tag}}</code>
    </span>
  `).join('');
}

function downloadCurrentTemplate() {
  if (!state.templateArrayBuffer) return;
  const blob = new Blob([state.templateArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = state.templateFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Tải file Excel mẫu chuẩn (Mẫu điền hợp đồng.xlsx)
function downloadSampleExcel() {
  if (typeof DEFAULT_EXCEL_B64 === 'undefined' || !DEFAULT_EXCEL_B64) {
    showToast("Không tìm thấy dữ liệu file mẫu Excel.", "error");
    return;
  }
  try {
    const binaryString = window.atob(DEFAULT_EXCEL_B64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes.buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Mẫu_điền_hợp_đồng.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("Đã tải xuống file mẫu Excel (Mẫu điền hợp đồng.xlsx)!", "success");
  } catch (err) {
    console.error("Lỗi tải mẫu Excel:", err);
    showToast("Lỗi khi tải mẫu Excel: " + err.message, "error");
  }
}

async function handleTemplateUpload(file) {
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    state.templateArrayBuffer = buffer;
    state.templateType = 'custom';
    state.templateFileName = file.name;
    updateTemplateUI();
    await scanPlaceholdersFromBuffer(buffer);
    showToast(`Đã tải lên mẫu hợp đồng mới: ${file.name}`, 'success');
  } catch (err) {
    showToast(`Lỗi đọc file mẫu: ${err.message}`, 'error');
  }
}

// ==========================================
// 3. XỬ LÝ EXCEL & GOOGLE SHEETS
// ==========================================

function findTargetSheet(sheetNames) {
  const normalizedMap = sheetNames.map(name => ({
    original: name,
    norm: normalizeKey(name)
  }));

  let found = normalizedMap.find(item => item.norm === 'thongtin' || item.norm === 'thongtintin');
  if (found) return found.original;

  found = normalizedMap.find(item => item.norm === 'sheet1' || item.norm === 'sheet01');
  if (found) return found.original;

  return null;
}

function parseWorkbookData(workbook) {
  const targetSheetName = findTargetSheet(workbook.SheetNames);
  if (!targetSheetName) {
    throw new Error("Tên sheet không đúng quy tắc");
  }

  state.selectedSheetName = targetSheetName;
  const worksheet = workbook.Sheets[targetSheetName];

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: ''
  });

  if (!rawRows || rawRows.length < 2) {
    throw new Error(`Sheet '${targetSheetName}' không có đủ dữ liệu.`);
  }

  const row1 = rawRows[0];
  const headers = [];
  for (let c = 1; c < row1.length; c++) {
    const h = String(row1[c] || '').trim();
    if (h) {
      headers.push({
        colIndex: c,
        raw: h,
        norm: normalizeKey(h)
      });
    }
  }

  if (headers.length === 0) {
    throw new Error("Không tìm thấy cột định danh nào từ Cột B trở đi ở hàng 1.");
  }
  state.headers = headers;

  const records = [];
  for (let r = 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    const record = {};
    let hasData = false;

    for (const h of headers) {
      let val = row[h.colIndex];
      if (val !== undefined && val !== null) {
        val = String(val).trim();
        if (/^\d+\.0$/.test(val)) {
          val = val.replace(/\.0$/, '');
        }
        if (val) hasData = true;
      } else {
        val = '';
      }
      record[h.norm] = val;
      record[h.raw] = val;
    }

    if (hasData) {
      records.push(record);
    }
  }

  state.excelData = records;
  renderDataSummary();
  return records;
}

function renderDataSummary() {
  const statusContainer = document.getElementById('data-status-container');
  const previewContainer = document.getElementById('data-preview-table');
  const btnProcess = document.getElementById('btn-start-process');

  if (state.excelData.length === 0) {
    if (btnProcess) btnProcess.disabled = true;
    return;
  }

  if (btnProcess) {
    btnProcess.disabled = false;
    btnProcess.innerHTML = `<i class="fa-solid fa-bolt mr-2"></i> Bắt đầu tạo ${state.excelData.length} Hợp đồng`;
  }

  if (statusContainer) {
    statusContainer.innerHTML = `
      <div class="p-3 bg-emerald-50 border border-emerald-200 rounded-lg flex items-center justify-between">
        <div class="flex items-center space-x-3">
          <i class="fa-solid fa-circle-check text-emerald-600 text-xl"></i>
          <div>
            <p class="font-semibold text-emerald-900">
              Đã nhận diện Sheet: <span class="underline font-bold">${state.selectedSheetName}</span>
            </p>
            <p class="text-xs text-emerald-700">
              Tìm thấy <span class="font-bold">${state.excelData.length} người</span> (Cột B trở đi, dòng 1 làm định danh, bắt đầu từ dòng 2).
            </p>
          </div>
        </div>
        <span class="px-2.5 py-1 text-xs font-semibold bg-emerald-200 text-emerald-800 rounded-full">
          Hợp lệ
        </span>
      </div>
    `;
  }

  if (previewContainer) {
    const displayCols = state.headers.slice(0, 6);
    let html = `
      <div class="mt-3 overflow-x-auto rounded border border-slate-200">
        <table class="min-w-full text-xs text-left text-slate-600">
          <thead class="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
            <tr>
              <th class="p-2">STT</th>
              ${displayCols.map(h => `<th class="p-2">${h.raw}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
    `;

    state.excelData.slice(0, 4).forEach((row, idx) => {
      html += `
        <tr class="border-b border-slate-100 hover:bg-slate-50">
          <td class="p-2 font-medium text-slate-400">${idx + 1}</td>
          ${displayCols.map(h => `<td class="p-2 font-mono">${row[h.norm] || '<span class="text-slate-300 italic">trống</span>'}</td>`).join('')}
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
      <p class="text-xs text-slate-500 mt-1 italic">* Bảo toàn số 0 ở đầu CCCD, SĐT và các kết quả tính từ hàm.</p>
    `;
    previewContainer.innerHTML = html;
  }
}

function handleExcelUpload(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      parseWorkbookData(workbook);
      showToast(`Đọc file '${file.name}' thành công! Tìm thấy ${state.excelData.length} người.`, 'success');
    } catch (err) {
      showErrorAlert(err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

async function handleGoogleSheetUrl(url) {
  if (!url || !url.trim()) {
    showToast("Vui lòng nhập đường link Google Sheet hợp lệ.", "error");
    return;
  }

  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    showToast("Link Google Sheet không đúng định dạng.", "error");
    return;
  }

  const sheetId = match[1];
  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const btnFetch = document.getElementById('btn-fetch-gsheet');
  if (btnFetch) {
    btnFetch.disabled = true;
    btnFetch.innerHTML = `<i class="fa-solid fa-spinner fa-spin mr-1"></i> Đang tải...`;
  }

  try {
    const response = await fetch(exportUrl);
    if (!response.ok) {
      throw new Error(`Google Sheet phản hồi lỗi mã ${response.status}. Hãy đảm bảo Sheet đã được chia sẻ công khai: 'Bất kỳ ai có đường liên kết đều có thể xem'.`);
    }
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    parseWorkbookData(workbook);
    showToast(`Đã lấy dữ liệu từ Google Sheet thành công! Có ${state.excelData.length} người.`, 'success');
  } catch (err) {
    console.error("Lỗi tải Google Sheet:", err);
    showErrorAlert(`Không thể tải trực tiếp từ Google Sheet: ${err.message}<br><br>💡 <b>Cách xử lý nhanh:</b> Vào Google Sheet > chọn <b>Tệp (File)</b> > <b>Tải xuống (Download)</b> > <b>Microsoft Excel (.xlsx)</b>, sau đó kéo thả file vừa tải vào tab "Tải file Excel" bên cạnh.`);
  } finally {
    if (btnFetch) {
      btnFetch.disabled = false;
      btnFetch.innerHTML = `<i class="fa-solid fa-cloud-arrow-down mr-1"></i> Tải dữ liệu từ Google Sheet`;
    }
  }
}

function useSampleData() {
  if (typeof SAMPLE_DATA !== 'undefined' && SAMPLE_DATA.length > 0) {
    state.selectedSheetName = 'Thông tin';
    state.excelData = SAMPLE_DATA;
    state.headers = [
      { raw: 'Congviec', norm: 'congviec' },
      { raw: 'Thoihan', norm: 'thoihan' },
      { raw: 'Thulao', norm: 'thulao' },
      { raw: 'Ngay', norm: 'ngay' },
      { raw: 'Thang', norm: 'thang' },
      { raw: 'Nam', norm: 'nam' },
      { raw: 'Hovaten', norm: 'hovaten' },
      { raw: 'Namsinh', norm: 'namsinh' },
      { raw: 'Thuongtru', norm: 'thuongtru' },
      { raw: 'CCCD', norm: 'cccd' },
      { raw: 'Ngaycap', norm: 'ngaycap' },
      { raw: 'Noicap', norm: 'noicap' },
      { raw: 'SDT', norm: 'sdt' },
      { raw: 'CTV', norm: 'ctv' }
    ];
    renderDataSummary();
    showToast("Đã tải dữ liệu mẫu gồm 27 người từ sheet 'Thông tin'!", "success");
  } else {
    showToast("Không tìm thấy tệp dữ liệu mẫu sample_data.js", "error");
  }
}

// ==========================================
// 4. QUY TẮC ĐIỀN THÔNG TIN & FORMATTING
// ==========================================

function resolveFieldValue(tag, personData) {
  const normTag = normalizeKey(tag);
  let val = personData[normTag] !== undefined && personData[normTag] !== '' ? personData[normTag] : personData[tag];
  let valStr = (val !== undefined && val !== null) ? String(val).trim() : '';

  if (!valStr) {
    if (['thulao', 'congviec', 'thoihan'].includes(normTag)) {
      return '............................................';
    } else if (['ngay', 'thang', 'nam', 'ngaycap'].includes(normTag)) {
      return '.....';
    } else {
      return '';
    }
  }

  if (tag.toUpperCase() === tag && tag.length > 1) {
    valStr = valStr.toUpperCase();
  }

  return valStr;
}

/**
 * Tạo cấu trúc HTML cho 1 hợp đồng phân tách làm 2 TRANG A4 CHUẨN:
 * - TRANG 1: Quốc hiệu, Tiêu ngữ, Căn cứ, Bên A, Bên B, Điều 1, Điều 2.
 * - TRANG 2: Điều 3, Điều 4, Điều 5, Điều 6 (Mục 3 ở trên), Bảng chữ ký 2 bên.
 * - TÊN KÝ CTV VÀ BÊN A IN ĐẬM VÀ VIẾT HOA CHUẨN XÁC: <b>${ctvName}</b>
 * - Không bao giờ bị cắt đôi chữ giữa 2 trang!
 */
function generateContractHTML(person, index) {
  const ngay = resolveFieldValue('Ngay', person);
  const thang = resolveFieldValue('Thang', person);
  const nam = resolveFieldValue('Nam', person);
  
  let hovaten = resolveFieldValue('Hovaten', person);
  const namsinh = resolveFieldValue('Namsinh', person);
  const thuongtru = resolveFieldValue('Thuongtru', person);
  const cccd = resolveFieldValue('CCCD', person);
  const ngaycap = resolveFieldValue('Ngaycap', person);
  const noicap = resolveFieldValue('Noicap', person);
  const sdt = resolveFieldValue('SDT', person);
  
  const congviec = resolveFieldValue('Congviec', person);
  const thoihan = resolveFieldValue('Thoihan', person);
  const thulao = resolveFieldValue('Thulao', person);
  
  let ctvName = resolveFieldValue('CTV', person);
  if (!ctvName || ctvName.includes('.')) {
    ctvName = hovaten.toUpperCase();
  } else {
    ctvName = ctvName.toUpperCase();
  }

  return `
    <div class="contract-wrapper" id="contract-person-${index}">
      <!-- ================= TRANG 1 ================= -->
      <div class="contract-subpage contract-subpage-1 a4-container">
        <!-- QUỐC HIỆU VÀ TIÊU NGỮ -->
        <div class="quoc-hieu">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
        <div class="tieu-ngu">Độc lập – Tự do – Hạnh phúc</div>
        <div class="tieu-ngu-line"></div>

        <!-- ĐỊA DANH NGÀY THÁNG -->
        <div class="ngay-thang-nam">Hà Nội, ngày ${ngay} tháng ${thang} năm ${nam}</div>

        <!-- TÊN HỢP ĐỒNG & SỐ -->
        <div class="ten-hop-dong">HỢP ĐỒNG CỘNG TÁC VIÊN</div>
        <div class="so-ky-hieu">Số: ……/2026/HĐCTV</div>

        <!-- CĂN CỨ -->
        <div style="margin-top: 10pt; margin-bottom: 8pt;">
          <p class="can-cu">Căn cứ Bộ luật Dân sự nước Cộng hòa xã hội chủ nghĩa Việt Nam ngày 24/11/2015;</p>
          <p class="can-cu">Căn cứ nhu cầu và năng lực của hai bên,</p>
          <p class="paragraph-justify" style="margin-top: 6pt;">
            Hợp đồng cộng tác viên này (sau đây gọi là “Hợp đồng”) được xem là hợp đồng dịch vụ được ký kết tại Công ty cổ phần đào tạo tư vấn và phát triển ABA và bởi các bên:
          </p>
        </div>

        <!-- BÊN A -->
        <div style="margin-bottom: 8pt;">
          <p class="paragraph-justify"><b>BÊN THUÊ DỊCH VỤ (BÊN A): CÔNG TY CỔ PHẦN ĐÀO TẠO TƯ VẤN VÀ PHÁT TRIỂN ABA</b></p>
          <p class="paragraph-justify">Địa chỉ: Số 33 Linh Lang, Phường Ngọc Hà, Thành Phố Hà Nội</p>
          <p class="paragraph-justify">Mã số thuế: 0105261305</p>
          <p class="paragraph-justify">Đại diện: Ông <b>Dương Đức Nam</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Chức vụ: <b>Giám Đốc</b></p>
        </div>

        <!-- BÊN B (CỘNG TÁC VIÊN) - CHỈ IN ĐẬM TÊN THEO MẪU WORD -->
        <div style="margin-bottom: 10pt;">
          <p class="paragraph-justify"><b>BÊN CỘNG TÁC VIÊN (BÊN B): <span class="font-bold">${hovaten}</span></b></p>
          <p class="paragraph-justify">Năm sinh: ${namsinh} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Giới tính: Nam / Nữ</p>
          <p class="paragraph-justify">Địa chỉ thường trú: ${thuongtru}</p>
          <p class="paragraph-justify">Số CCCD: ${cccd} &nbsp;&nbsp;&nbsp;&nbsp; Ngày cấp: ${ngaycap} &nbsp;&nbsp;&nbsp;&nbsp; Nơi cấp: ${noicap}</p>
          <p class="paragraph-justify">Điện thoại: ${sdt}</p>
        </div>

        <p class="paragraph-justify" style="margin-bottom: 8pt;">Sau khi thỏa thuận, hai bên nhất trí ký kết hợp đồng với các điều khoản như sau:</p>

        <!-- ĐIỀU 1 -->
        <div class="dieu-khoan-title">Điều 1: Nội dung và thời hạn công việc</div>
        <p class="paragraph-justify">Bên A đồng ý thuê và bên B đồng ý làm Cộng tác viên thực hiện các công việc sau:</p>
        <p class="paragraph-justify">1. Nội dung công việc: &nbsp;${congviec}</p>
        <p class="paragraph-justify">2. Thời hạn thực hiện: &nbsp;${thoihan}</p>
        <p class="paragraph-justify">3. Địa điểm: Theo thỏa thuận cụ thể của từng đầu việc</p>
        <p class="paragraph-justify">4. Thiết bị và công cụ làm việc sẽ được cấp phát tùy theo nhu cầu của công việc</p>

        <!-- ĐIỀU 2 -->
        <div class="dieu-khoan-title">Điều 2: Thù lao và phương thức thanh toán</div>
        <p class="paragraph-justify">1. Thù lao: &nbsp;${thulao}</p>
        <p class="paragraph-justify">Thuế thu nhập cá nhân do người lao động tự quyết toán</p>
        <p class="paragraph-justify">2. Phương thức thanh toán: Chuyển khoản hoặc Tiền mặt.</p>
        <p class="paragraph-justify">3. Thời hạn thanh toán: Sau khi hoàn thành công việc 15 ngày</p>
      </div>

      <!-- ================= TRANG 2 ================= -->
      <div class="contract-subpage contract-subpage-2 a4-container">
        <!-- ĐIỀU 3 (BẮT ĐẦU TRANG 2 SẠCH ĐẸP) -->
        <div class="dieu-khoan-title" style="margin-top: 0;">Điều 3: Quyền và nghĩa vụ của bên A</div>
        <p class="paragraph-justify">1. Yêu cầu cộng tác viên thực hiện công việc theo đúng chất lượng, số lượng, thời hạn, địa điểm và các thỏa thuận khác.</p>
        <p class="paragraph-justify">2. Cung cấp cho cộng tác viên thông tin, tài liệu và các phương tiện cần thiết để thực hiện công việc. Bảo đảm việc làm cho cộng tác viên theo Hợp đồng đã ký</p>
        <p class="paragraph-justify">3. Thanh toán tiền đầy đủ và đúng hạn cho cộng tác viên</p>
        <p class="paragraph-justify">4. Có quyền tạm dừng thanh toán nếu Bên B chưa hoàn thành công việc theo cam kết hoặc chưa hoàn trả đầy đủ tài sản của công ty.</p>
        <p class="paragraph-justify">5. Đơn phương chấm dứt hợp đồng và yêu cầu bồi thường nếu Bên B vi phạm nghiêm trọng các cam kết về chất lượng và bảo mật</p>

        <!-- ĐIỀU 4 -->
        <div class="dieu-khoan-title">Điều 4: Quyền và nghĩa vụ của bên B</div>
        <p class="paragraph-justify">1. Thực hiện công việc đúng chất lượng, số lượng, thời hạn, địa điểm và các thoả thuận khác.</p>
        <p class="paragraph-justify">2. Tự chịu trách nhiệm về phương tiện (nếu không được cấp) và cách thức tổ chức công việc để đạt kết quả</p>
        <p class="paragraph-justify">3. Giữ bí mật tuyệt đối các thông tin kinh doanh, tài liệu của Bên A trong và sau quá trình thực hiện công việc.</p>
        <p class="paragraph-justify">4. Hoàn trả lại toàn bộ tài liệu, thiết bị được cấp phát ngay sau khi kết thúc công việc hoặc chấm dứt hợp đồng</p>
        <p class="paragraph-justify">5. Bồi thường thiệt hại cho bên A nếu làm mất, hư hỏng tài liệu, phương tiện được giao hoặc tiết lộ bí mật thông tin.</p>

        <!-- ĐIỀU 5 -->
        <div class="dieu-khoan-title">Điều 5: Chấm dứt hợp đồng</div>
        <p class="paragraph-justify">1. Hợp đồng đương nhiên chấm dứt khi hết thời hạn quy định tại Điều 1.</p>
        <p class="paragraph-justify">2. Mỗi bên có quyền đơn phương chấm dứt hợp đồng trước thời hạn nhưng phải thông báo bằng văn bản cho bên kia trước ít nhất 05 ngày.</p>

        <!-- ĐIỀU 6 - MỤC 3 NẰM TRÊN BẢNG CHỮ KÝ -->
        <div class="dieu-khoan-title">Điều 6: Điều khoản chung</div>
        <p class="paragraph-justify">1. Hợp đồng này là hợp đồng dân sự, không xác lập quan hệ lao động, không thuộc đối tượng đóng Bảo hiểm xã hội bắt buộc</p>
        <p class="paragraph-justify">2. Mọi tranh chấp phát sinh sẽ được ưu tiên giải quyết qua thương lượng. Trường hợp không thỏa thuận được, tranh chấp sẽ được giải quyết tại Tòa án có thẩm quyền.</p>
        <p class="paragraph-justify">3. Hợp đồng được lập thành 02 (hai) bản có giá trị pháp lý như nhau, mỗi bên giữ 01 (một) bản</p>

        <!-- BẢNG CHỮ KÝ 2 CỘT: TÊN KÝ IN ĐẬM VÀ VIẾT HOA CHUẨN XÁC -->
        <table class="table-chu-ky">
          <tr>
            <td>
              <div class="chu-ky-title">CỘNG TÁC VIÊN</div>
              <div class="chu-ky-ghi-chu">(Ký, ghi rõ họ tên)</div>
              <div class="chu-ky-ho-ten font-bold"><b>${ctvName}</b></div>
            </td>
            <td>
              <div class="chu-ky-title">ĐẠI DIỆN BÊN SỬ DỤNG DỊCH VỤ</div>
              <div class="chu-ky-ghi-chu">(Ký, ghi rõ họ tên)</div>
              <div class="chu-ky-ho-ten font-bold"><b>DƯƠNG ĐỨC NAM</b></div>
            </td>
          </tr>
        </table>
      </div>
    </div>
  `;
}

// ==========================================
// 5. TIẾN TRÌNH XỬ LÝ & TẠO HỢP ĐỒNG HÀNG LOẠT
// ==========================================

async function startBatchProcessing() {
  if (state.excelData.length === 0) {
    showToast("Vui lòng tải file Excel hoặc nạp link Google Sheet trước.", "error");
    return;
  }

  state.isProcessing = true;
  state.processedContractsCount = 0;
  state.generatedContracts = [];

  const progressSection = document.getElementById('progress-section');
  const progressBar = document.getElementById('progress-bar');
  const progressText = document.getElementById('progress-text');
  const progressPercent = document.getElementById('progress-percent');
  const printArea = document.getElementById('print-area');

  if (progressSection) progressSection.classList.remove('hidden');
  if (printArea) printArea.innerHTML = '';

  const total = state.excelData.length;
  let allHTML = '';

  for (let i = 0; i < total; i++) {
    const person = state.excelData[i];
    const personName = person['Hovaten'] || person['hovaten'] || `Người ${i + 1}`;

    const contractHTML = generateContractHTML(person, i);
    allHTML += contractHTML;
    state.generatedContracts.push(contractHTML);

    state.processedContractsCount = i + 1;
    const percent = Math.round(((i + 1) / total) * 100);

    if (progressBar) progressBar.style.width = `${percent}%`;
    if (progressPercent) progressPercent.innerText = `${percent}%`;
    if (progressText) {
      progressText.innerText = `Đã hoàn thành ${i + 1} / ${total} người: ${personName}`;
    }

    if (i % 2 === 0 || i === total - 1) {
      await new Promise(r => setTimeout(r, 20));
    }
  }

  if (printArea) {
    printArea.innerHTML = allHTML;
  }

  state.isProcessing = false;
  showCompletionModal(total);
}

// ==========================================
// 6. CỬA SỔ MODAL HOÀN THÀNH (IN & TẢI VỀ)
// ==========================================

function showCompletionModal(totalCount) {
  const modal = document.getElementById('completion-modal');
  const countEl = document.getElementById('modal-completed-count');
  if (countEl) countEl.innerText = `${totalCount} hợp đồng`;
  if (modal) {
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }
}

function closeCompletionModal() {
  const modal = document.getElementById('completion-modal');
  if (modal) {
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }
}

/**
 * CHỨC NĂNG IN CHUẨN:
 * - Ẩn modal và XÓA TOÀN BỘ thông báo Toast trên màn hình trước khi gọi in
 * - Mỗi trang con (.contract-subpage) in đúng 1 trang A4
 * - Sau khi in đóng lại thì khôi phục modal
 */
function triggerPrint() {
  // Xóa toàn bộ toast đang có trên màn hình để không bị in ra giấy
  document.querySelectorAll('.toast-alert, .no-print-toast').forEach(el => el.remove());

  const completionModal = document.getElementById('completion-modal');
  const previewModal = document.getElementById('preview-modal');
  const wasCompletionOpen = completionModal && !completionModal.classList.contains('hidden');
  const wasPreviewOpen = previewModal && !previewModal.classList.contains('hidden');

  if (completionModal) {
    completionModal.classList.add('hidden');
    completionModal.classList.remove('flex');
  }
  if (previewModal) {
    previewModal.classList.add('hidden');
    previewModal.classList.remove('flex');
  }

  setTimeout(() => {
    window.print();

    setTimeout(() => {
      if (wasPreviewOpen && previewModal) {
        previewModal.classList.remove('hidden');
        previewModal.classList.add('flex');
      } else if (wasCompletionOpen && completionModal) {
        completionModal.classList.remove('hidden');
        completionModal.classList.add('flex');
      }
    }, 400);
  }, 100);
}

// Chức năng TẢI VỀ WORD (.DOCX) GỘP
async function triggerDownloadDocx() {
  showToast("Đang tải file Word (.docx) gộp đã xử lý...", "info");
  try {
    const a = document.createElement('a');
    a.href = 'HopDong_TongHop.docx';
    a.download = `HopDong_TongHop_${state.excelData.length}_nguoi.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast("Đã tải xuống file Word gộp!", "success");
  } catch (err) {
    showToast("Không thể tải file Word: " + err.message, "error");
  }
}

// Xem trước trực tiếp trong modal
function openLivePreview() {
  const printArea = document.getElementById('print-area');
  const previewModal = document.getElementById('preview-modal');
  const previewContent = document.getElementById('preview-modal-content');
  if (previewModal && previewContent && printArea) {
    previewContent.innerHTML = printArea.innerHTML;
    previewModal.classList.remove('hidden');
    previewModal.classList.add('flex');
  }
}

function closeLivePreview() {
  const previewModal = document.getElementById('preview-modal');
  if (previewModal) {
    previewModal.classList.add('hidden');
    previewModal.classList.remove('flex');
  }
}

// ==========================================
// 7. TOAST & ERROR ALERTS
// ==========================================

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  const bgColors = {
    success: 'bg-emerald-600 text-white',
    error: 'bg-rose-600 text-white',
    info: 'bg-blue-600 text-white'
  };
  const icons = {
    success: 'fa-circle-check',
    error: 'fa-triangle-exclamation',
    info: 'fa-circle-info'
  };

  toast.className = `no-print toast-alert no-print-toast fixed bottom-5 right-5 z-50 flex items-center px-4 py-3 rounded-lg shadow-xl text-sm font-medium transition-all duration-300 transform translate-y-5 opacity-0 ${bgColors[type] || bgColors.info}`;
  toast.innerHTML = `<i class="fa-solid ${icons[type] || icons.info} mr-2 text-base"></i> <span>${message}</span>`;
  document.body.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.remove('translate-y-5', 'opacity-0');
  });

  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-y-5');
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function showErrorAlert(message) {
  const container = document.getElementById('data-status-container');
  if (container) {
    container.innerHTML = `
      <div class="p-4 bg-rose-50 border border-rose-300 text-rose-800 rounded-lg flex items-start space-x-3">
        <i class="fa-solid fa-triangle-exclamation text-rose-600 text-xl mt-0.5"></i>
        <div>
          <h4 class="font-bold text-rose-900 mb-1">Lỗi: Tên sheet không đúng quy tắc</h4>
          <p class="text-xs text-rose-700 leading-relaxed">${message}</p>
        </div>
      </div>
    `;
  }
  showToast(message, 'error');
}

// ==========================================
// 8. BẮT SỰ KIỆN GIAO DIỆN
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
  initDefaultTemplate();

  const templateDropzone = document.getElementById('template-dropzone');
  const templateFileInput = document.getElementById('template-file-input');

  if (templateDropzone && templateFileInput) {
    templateDropzone.addEventListener('click', () => templateFileInput.click());
    templateFileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleTemplateUpload(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      templateDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        templateDropzone.classList.add('dropzone-active');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      templateDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        templateDropzone.classList.remove('dropzone-active');
      });
    });

    templateDropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files.length > 0) {
        handleTemplateUpload(e.dataTransfer.files[0]);
      }
    });
  }

  const excelDropzone = document.getElementById('excel-dropzone');
  const excelFileInput = document.getElementById('excel-file-input');

  if (excelDropzone && excelFileInput) {
    excelDropzone.addEventListener('click', () => excelFileInput.click());
    excelFileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleExcelUpload(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      excelDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        excelDropzone.classList.add('dropzone-active');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      excelDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        excelDropzone.classList.remove('dropzone-active');
      });
    });

    excelDropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files.length > 0) {
        handleExcelUpload(e.dataTransfer.files[0]);
      }
    });
  }

  const tabExcel = document.getElementById('tab-excel');
  const tabGsheet = document.getElementById('tab-gsheet');
  const panelExcel = document.getElementById('panel-excel');
  const panelGsheet = document.getElementById('panel-gsheet');

  if (tabExcel && tabGsheet && panelExcel && panelGsheet) {
    tabExcel.addEventListener('click', () => {
      tabExcel.className = 'px-4 py-2 text-sm font-semibold text-blue-600 border-b-2 border-blue-600 flex items-center';
      tabGsheet.className = 'px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-700 flex items-center';
      panelExcel.classList.remove('hidden');
      panelGsheet.classList.add('hidden');
    });

    tabGsheet.addEventListener('click', () => {
      tabGsheet.className = 'px-4 py-2 text-sm font-semibold text-blue-600 border-b-2 border-blue-600 flex items-center';
      tabExcel.className = 'px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-700 flex items-center';
      panelGsheet.classList.remove('hidden');
      panelExcel.classList.add('hidden');
    });
  }

  const btnFetchGsheet = document.getElementById('btn-fetch-gsheet');
  const inputGsheetUrl = document.getElementById('input-gsheet-url');
  if (btnFetchGsheet && inputGsheetUrl) {
    btnFetchGsheet.addEventListener('click', () => {
      handleGoogleSheetUrl(inputGsheetUrl.value);
    });
  }

  const btnStartProcess = document.getElementById('btn-start-process');
  if (btnStartProcess) {
    btnStartProcess.addEventListener('click', startBatchProcessing);
  }
});
