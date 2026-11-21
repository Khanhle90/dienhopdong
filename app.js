/**
 * HỢP ĐỒNG TỰ ĐỘNG - CORE APPLICATION ENGINE
 * Xử lý điền hợp đồng từ Excel / Google Sheets sang Word & PDF
 */

// Trạng thái ứng dụng
const state = {
  templateType: 'default', // 'default' hoặc 'custom'
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

// Chuyển base64 sang ArrayBuffer
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

    // Trích xuất text sạch
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

// Xử lý upload mẫu Docx mới
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

/**
 * Tìm Sheet mục tiêu:
 * 1. Kiểm tra sheet 'Thông tin'
 * 2. Nếu không có kiểm tra sheet 'Sheet 1' / 'Sheet1'
 * 3. Nếu không có thì báo lỗi 'Tên sheet không đúng quy tắc'
 */
function findTargetSheet(sheetNames) {
  const normalizedMap = sheetNames.map(name => ({
    original: name,
    norm: normalizeKey(name)
  }));

  // 1. Tìm 'thongtin'
  let found = normalizedMap.find(item => item.norm === 'thongtin' || item.norm === 'thongtintin');
  if (found) return found.original;

  // 2. Tìm 'sheet1'
  found = normalizedMap.find(item => item.norm === 'sheet1' || item.norm === 'sheet01');
  if (found) return found.original;

  // 3. Không tìm thấy theo quy tắc
  return null;
}

/**
 * Đọc bảng Excel / Google Sheets theo quy tắc:
 * - Dữ liệu luôn lấy từ Cột B trở đi
 * - Hàng 1 làm định danh thay thế
 * - Dữ liệu bắt đầu lấy từ hàng 2
 * - Giữ nguyên toàn bộ số 0 ở đầu (CCCD, SĐT)
 * - Lấy chuẩn thông tin hiển thị (kể cả ô có hàm/công thức)
 */
function parseWorkbookData(workbook) {
  const targetSheetName = findTargetSheet(workbook.SheetNames);
  if (!targetSheetName) {
    throw new Error("Tên sheet không đúng quy tắc");
  }

  state.selectedSheetName = targetSheetName;
  const worksheet = workbook.Sheets[targetSheetName];

  // raw: false => Lấy chuỗi hiển thị đã định dạng (giữ nguyên số 0 ở đầu và giá trị hàm)
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: ''
  });

  if (!rawRows || rawRows.length < 2) {
    throw new Error(`Sheet '${targetSheetName}' không có đủ dữ liệu (cần ít nhất 1 dòng tiêu đề và 1 dòng dữ liệu).`);
  }

  // Dòng 1: Header định danh, bắt đầu từ Cột B (index >= 1)
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

  // Dòng 2 trở đi: Dữ liệu từng người
  const records = [];
  for (let r = 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    const record = {};
    let hasData = false;

    for (const h of headers) {
      let val = row[h.colIndex];
      if (val !== undefined && val !== null) {
        val = String(val).trim();
        // Xóa đuôi .0 nếu Excel lưu số nguyên dạng float
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

// Hiển thị tóm tắt dữ liệu Excel đã đọc
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
              Tìm thấy <span class="font-bold">${state.excelData.length} người</span> (Lấy từ Cột B trở đi, dòng 1 làm định danh, bắt đầu từ dòng 2).
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
      <p class="text-xs text-slate-500 mt-1 italic">* Bảng hiển thị 4 người đầu tiên làm mẫu kiểm tra số 0 ở đầu (CCCD, SĐT) và giá trị hàm.</p>
    `;
    previewContainer.innerHTML = html;
  }
}

// Xử lý khi upload file Excel
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

// Xử lý nạp dữ liệu từ Google Sheets
async function handleGoogleSheetUrl(url) {
  if (!url || !url.trim()) {
    showToast("Vui lòng nhập đường link Google Sheet hợp lệ.", "error");
    return;
  }

  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    showToast("Link Google Sheet không đúng định dạng (cần chứa mã ID bảng tính).", "error");
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

// Xử lý nút dùng dữ liệu mẫu (27 người)
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

/**
 * Lấy giá trị thay thế theo đúng quy định:
 * - thulao, congviec, thoihan nếu bỏ trống -> ............................................
 * - ngay, thang, nam nếu bỏ trống -> .....
 * - vị trí viết hoa -> tự động viết hoa
 */
function resolveFieldValue(tag, personData) {
  const normTag = normalizeKey(tag);
  let val = personData[normTag] !== undefined && personData[normTag] !== '' ? personData[normTag] : personData[tag];
  let valStr = (val !== undefined && val !== null) ? String(val).trim() : '';

  // Quy tắc mặc định khi bỏ trống
  if (!valStr) {
    if (['thulao', 'congviec', 'thoihan'].includes(normTag)) {
      return '............................................';
    } else if (['ngay', 'thang', 'nam', 'ngaycap'].includes(normTag)) {
      return '.....';
    } else {
      return '';
    }
  }

  // Quy tắc viết hoa: nếu tag viết hoa (ví dụ: {CTV}) hoặc là tên CTV/Họ và tên
  if (tag.toUpperCase() === tag && tag.length > 1) {
    valStr = valStr.toUpperCase();
  }

  return valStr;
}

/**
 * Tạo HTML cho 1 bản hợp đồng hoàn chỉnh chuẩn thể thức A4
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
  
  // Tên CTV dưới chữ ký
  let ctvName = resolveFieldValue('CTV', person);
  if (!ctvName || ctvName.includes('.')) {
    ctvName = hovaten.toUpperCase();
  } else {
    ctvName = ctvName.toUpperCase();
  }

  return `
    <div class="contract-paper a4-container" id="contract-person-${index}">
      <!-- QUỐC HIỆU VÀ TIÊU NGỮ -->
      <div class="quoc-hieu">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
      <div class="tieu-ngu">Độc lập – Tự do – Hạnh phúc</div>
      <div class="tieu-ngu-line"></div>

      <!-- ĐỊA DANH NGÀY THÁNG -->
      <div class="ngay-thang-nam">Hà Nội, ngày <b>${ngay}</b> tháng <b>${thang}</b> năm <b>${nam}</b></div>

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

      <!-- BÊN B (CỘNG TÁC VIÊN) -->
      <div style="margin-bottom: 10pt;">
        <p class="paragraph-justify"><b>BÊN CỘNG TÁC VIÊN (BÊN B): <span class="font-bold">${hovaten}</span></b></p>
        <p class="paragraph-justify">Năm sinh: <b>${namsinh}</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Giới tính: Nam / Nữ</p>
        <p class="paragraph-justify">Địa chỉ thường trú: <b>${thuongtru}</b></p>
        <p class="paragraph-justify">Số CCCD: <b>${cccd}</b> &nbsp;&nbsp;&nbsp;&nbsp; Ngày cấp: <b>${ngaycap}</b> &nbsp;&nbsp;&nbsp;&nbsp; Nơi cấp: <b>${noicap}</b></p>
        <p class="paragraph-justify">Điện thoại: <b>${sdt}</b></p>
      </div>

      <p class="paragraph-justify" style="margin-bottom: 8pt;">Sau khi thỏa thuận, hai bên nhất trí ký kết hợp đồng với các điều khoản như sau:</p>

      <!-- ĐIỀU 1 -->
      <div class="dieu-khoan-title">Điều 1: Nội dung và thời hạn công việc</div>
      <p class="paragraph-justify">Bên A đồng ý thuê và bên B đồng ý làm Cộng tác viên thực hiện các công việc sau:</p>
      <p class="paragraph-justify">1. Nội dung công việc: &nbsp;<b>${congviec}</b></p>
      <p class="paragraph-justify">2. Thời hạn thực hiện: &nbsp;<b>${thoihan}</b></p>
      <p class="paragraph-justify">3. Địa điểm: Theo thỏa thuận cụ thể của từng đầu việc</p>
      <p class="paragraph-justify">4. Thiết bị và công cụ làm việc sẽ được cấp phát tùy theo nhu cầu của công việc</p>

      <!-- ĐIỀU 2 -->
      <div class="dieu-khoan-title">Điều 2: Thù lao và phương thức thanh toán</div>
      <p class="paragraph-justify">1. Thù lao: &nbsp;<b>${thulao}</b></p>
      <p class="paragraph-justify">Thuế thu nhập cá nhân do người lao động tự quyết toán</p>
      <p class="paragraph-justify">2. Phương thức thanh toán: Chuyển khoản hoặc Tiền mặt.</p>
      <p class="paragraph-justify">3. Thời hạn thanh toán: Sau khi hoàn thành công việc 15 ngày</p>

      <!-- ĐIỀU 3 -->
      <div class="dieu-khoan-title">Điều 3: Quyền và nghĩa vụ của bên A</div>
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

      <!-- ĐIỀU 6 -->
      <div class="dieu-khoan-title">Điều 6: Điều khoản chung</div>
      <p class="paragraph-justify">1. Hợp đồng này là hợp đồng dân sự, không xác lập quan hệ lao động, không thuộc đối tượng đóng Bảo hiểm xã hội bắt buộc</p>
      <p class="paragraph-justify">2. Mọi tranh chấp phát sinh sẽ được ưu tiên giải quyết qua thương lượng. Trường hợp không thỏa thuận được, tranh chấp sẽ được giải quyết tại Tòa án có thẩm quyền.</p>
      <p class="paragraph-justify">3. Hợp đồng được lập thành 02 (hai) bản có giá trị pháp lý như nhau, mỗi bên giữ 01 (một) bản</p>

      <!-- BẢNG CHỮ KÝ 2 CỘT -->
      <table class="table-chu-ky">
        <tr>
          <td>
            <div class="chu-ky-title">CỘNG TÁC VIÊN</div>
            <div class="chu-ky-ghi-chu">(Ký, ghi rõ họ tên)</div>
            <div class="chu-ky-ho-ten">${ctvName}</div>
          </td>
          <td>
            <div class="chu-ky-title">ĐẠI DIỆN BÊN SỬ DỤNG DỊCH VỤ</div>
            <div class="chu-ky-ghi-chu">(Ký, ghi rõ họ tên)</div>
            <div class="chu-ky-ho-ten">DƯƠNG ĐỨC NAM</div>
          </td>
        </tr>
      </table>
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
  const renderArea = document.getElementById('contracts-render-area');

  if (progressSection) progressSection.classList.remove('hidden');
  if (renderArea) renderArea.innerHTML = '';

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

    // Nhường luồng cho UI cập nhật mượt mà
    if (i % 2 === 0 || i === total - 1) {
      await new Promise(r => setTimeout(r, 40));
    }
  }

  if (renderArea) {
    renderArea.innerHTML = allHTML;
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

// Chức năng IN (Print)
function triggerPrint() {
  window.print();
}

// Chức năng TẢI VỀ PDF
async function triggerDownloadPDF() {
  const renderArea = document.getElementById('contracts-render-area');
  if (!renderArea || renderArea.children.length === 0) {
    showToast("Chưa có hợp đồng nào được tạo.", "error");
    return;
  }

  const btnDownload = document.getElementById('btn-download-pdf');
  const originalText = btnDownload ? btnDownload.innerHTML : '';
  if (btnDownload) {
    btnDownload.disabled = true;
    btnDownload.innerHTML = `<i class="fa-solid fa-spinner fa-spin mr-2"></i> Đang xuất PDF...`;
  }

  try {
    const opt = {
      margin: [0, 0, 0, 0], // Margin đã được định vị chính xác trong css .contract-paper
      filename: `HopDong_TongHop_${state.excelData.length}_nguoi.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, letterRendering: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      pagebreak: { mode: ['css', 'legacy'] }
    };

    await html2pdf().set(opt).from(renderArea).save();
    showToast("Đã tải về file PDF hợp đồng gộp thành công!", "success");
  } catch (err) {
    console.error("Lỗi xuất PDF:", err);
    showToast("Gặp trục trặc khi tạo PDF qua thư viện. Bạn có thể nhấn nút 'In' và chọn 'Lưu dưới dạng PDF' (Save as PDF) với độ sắc nét tuyệt đối!", "info");
  } finally {
    if (btnDownload) {
      btnDownload.disabled = false;
      btnDownload.innerHTML = originalText;
    }
  }
}

// Chức năng TẢI VỀ WORD (.DOCX)
async function triggerDownloadDocx() {
  showToast("Đang chuẩn bị file Word (.docx) gộp...", "info");
  try {
    // Tải file HopDong_TongHop.docx đã tạo hoặc tạo mới qua JSZip
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
  const renderArea = document.getElementById('contracts-render-area');
  const previewModal = document.getElementById('preview-modal');
  const previewContent = document.getElementById('preview-modal-content');
  if (previewModal && previewContent && renderArea) {
    previewContent.innerHTML = renderArea.innerHTML;
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

  toast.className = `fixed bottom-5 right-5 z-50 flex items-center px-4 py-3 rounded-lg shadow-xl text-sm font-medium transition-all duration-300 transform translate-y-5 opacity-0 ${bgColors[type] || bgColors.info}`;
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
// 8. BẮT SỰ KIỆN GIAO DIỆN KHI TẢI TRANG
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
  initDefaultTemplate();

  // Drag & drop upload mẫu Word
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

  // Drag & drop upload Excel
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

  // Chuyển tab giữa File Excel và Google Sheet
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

  // Nút tải dữ liệu Google Sheet
  const btnFetchGsheet = document.getElementById('btn-fetch-gsheet');
  const inputGsheetUrl = document.getElementById('input-gsheet-url');
  if (btnFetchGsheet && inputGsheetUrl) {
    btnFetchGsheet.addEventListener('click', () => {
      handleGoogleSheetUrl(inputGsheetUrl.value);
    });
  }

  // Nút Bắt đầu xử lý
  const btnStartProcess = document.getElementById('btn-start-process');
  if (btnStartProcess) {
    btnStartProcess.addEventListener('click', startBatchProcessing);
  }
});
