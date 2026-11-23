"""
CHƯƠNG TRÌNH TỰ ĐỘNG ĐIỀN HỢP ĐỒNG TỪ EXCEL SANG WORD & PDF
Tác giả: Hợp đồng tự động ABA
Phiên bản: 2.1 (Đã sửa lỗi phân trang, lỗi bảng trôi và định dạng in đậm)
"""

import os
import re
import copy
import openpyxl
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn

def normalize_key(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'[^a-z0-9]', '', s)
    return s

def read_excel(excel_path):
    """
    Đọc dữ liệu từ file Excel:
    - Tìm sheet 'Thông tin', nếu không có tìm 'Sheet 1', nếu không có báo lỗi 'Tên sheet không đúng quy tắc'
    - Luôn lấy dữ liệu từ Cột B trở đi
    - Dòng 1 làm định danh
    - Dòng 2 trở đi là dữ liệu
    - Giữ nguyên số 0 ở đầu (CCCD, SĐT)
    - Lấy chuẩn giá trị hiển thị / hàm
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")
        
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    target_sheet = None
    for name in wb.sheetnames:
        norm = normalize_key(name)
        if norm in ['thongtin', 'thongtintin']:
            target_sheet = name
            break
            
    if not target_sheet:
        for name in wb.sheetnames:
            norm = normalize_key(name)
            if norm in ['sheet1', 'sheet01']:
                target_sheet = name
                break
                
    if not target_sheet:
        raise ValueError("Tên sheet không đúng quy tắc")
        
    ws = wb[target_sheet]
    print(f"[+] Đã mở sheet: '{target_sheet}' (Tổng số dòng: {ws.max_row}, Cột: {ws.max_column})")
    
    # Lấy headers từ cột B (cột index 2) trở đi ở dòng 1
    headers = {}
    for col in range(2, ws.max_column + 1):
        cell_val = ws.cell(row=1, column=col).value
        if cell_val is not None and str(cell_val).strip():
            raw_h = str(cell_val).strip()
            headers[col] = {
                'raw': raw_h,
                'norm': normalize_key(raw_h)
            }
            
    # Lấy dữ liệu từ dòng 2 trở đi
    records = []
    for row in range(2, ws.max_row + 1):
        row_dict = {}
        has_val = False
        for col, h in headers.items():
            cell = ws.cell(row=row, column=col)
            val = cell.value
            if val is not None:
                if isinstance(val, float) and val.is_integer():
                    val_str = str(int(val))
                else:
                    val_str = str(val).strip()
                if val_str:
                    has_val = True
            else:
                val_str = ""
                
            row_dict[h['norm']] = val_str
            row_dict[h['raw']] = val_str
            
        if has_val:
            records.append(row_dict)
            
    print(f"[+] Đã đọc thành công {len(records)} bản ghi từ bảng Excel.")
    return records

def get_field_value(tag, data):
    norm_tag = normalize_key(tag)
    val = data.get(norm_tag)
    if val is None or val == '':
        val = data.get(tag, '')
        
    val_str = str(val).strip() if val is not None else ''
    if not val_str:
        if norm_tag in ['thulao', 'congviec', 'thoihan']:
            return '............................................'
        elif norm_tag in ['ngay', 'thang', 'nam', 'ngaycap']:
            return '.....'
        else:
            return ''
            
    return val_str

def replace_in_paragraph(paragraph, data):
    """
    Thay thế các placeholder {tag} trong đoạn văn bản.
    QUY TẮC BẢO TOÀN IN ĐẬM:
    - CHỈ in đậm khi chính từ placeholder đó được in đậm trong mẫu (ví dụ: {HoVaTen} trong BÊN B).
    - Các từ khác (như {Thuongtru}, {CCCD}, {SDT}, {Congviec}...) nếu mẫu viết thường thì giữ NGUYÊN viết thường (không tự ý bôi đen).
    """
    full_text = "".join(r.text for r in paragraph.runs)
    if '{' not in full_text:
        return
        
    matches = list(re.finditer(r'\{([^}]+)\}', full_text))
    if not matches:
        return

    char_runs = []
    for r_idx, r in enumerate(paragraph.runs):
        for _ in r.text:
            char_runs.append(r_idx)
            
    new_runs_data = []
    current_char = 0
    for m in matches:
        start, end = m.span()
        tag = m.group(1).strip()
        
        # Phần text trước placeholder
        if start > current_char:
            for c_idx in range(current_char, start):
                r_idx = char_runs[c_idx]
                r = paragraph.runs[r_idx]
                new_runs_data.append({
                    'char': full_text[c_idx],
                    'bold': r.bold,
                    'italic': r.italic,
                    'underline': r.underline,
                    'font_name': r.font.name if r.font else None,
                    'font_size': r.font.size if r.font else None
                })
                
        # Lấy giá trị thay thế
        val = get_field_value(tag, data)
        
        # Kiểm tra xem chính phần text của tag (không tính dấu {}) có được in đậm không
        tag_start_char = start + 1
        tag_r_idx = char_runs[tag_start_char] if tag_start_char < len(char_runs) else char_runs[start]
        tag_run = paragraph.runs[tag_r_idx]
        
        # In đậm CHỈ KHI run chứa tag có bold=True (ví dụ {HoVaTen} trong Bên B) hoặc trường {CTV}
        is_bold = bool(tag_run.bold)
        is_italic = bool(tag_run.italic)
        
        # Quy tắc: nếu vị trí thay thế viết hoa thì tự động viết hoa (như {CTV})
        # và tên ký {CTV} luôn luôn in đậm đúng thể thức chữ ký
        if tag.isupper() or normalize_key(tag) == 'ctv':
            val = val.upper()
            if normalize_key(tag) == 'ctv':
                is_bold = True
            
        font_name = tag_run.font.name if tag_run.font else None
        font_size = tag_run.font.size if tag_run.font else None
        
        new_runs_data.append({
            'chunk': val,
            'bold': is_bold,
            'italic': is_italic,
            'underline': tag_run.underline,
            'font_name': font_name,
            'font_size': font_size
        })
        current_char = end
        
    # Phần text còn lại sau placeholder cuối cùng
    if current_char < len(full_text):
        for c_idx in range(current_char, len(full_text)):
            r_idx = char_runs[c_idx]
            r = paragraph.runs[r_idx]
            new_runs_data.append({
                'char': full_text[c_idx],
                'bold': r.bold,
                'italic': r.italic,
                'underline': r.underline,
                'font_name': r.font.name if r.font else None,
                'font_size': r.font.size if r.font else None
            })
            
    # Hợp nhất các ký tự có cùng định dạng
    merged_runs = []
    curr_run = None
    for item in new_runs_data:
        text_chunk = item.get('chunk', item.get('char', ''))
        b = item.get('bold', False)
        it = item.get('italic', False)
        u = item.get('underline', False)
        fn = item.get('font_name')
        fs = item.get('font_size')
        
        if curr_run and curr_run['bold'] == b and curr_run['italic'] == it and curr_run['font_name'] == fn and curr_run['font_size'] == fs:
            curr_run['text'] += text_chunk
        else:
            if curr_run:
                merged_runs.append(curr_run)
            curr_run = {
                'text': text_chunk,
                'bold': b,
                'italic': it,
                'underline': u,
                'font_name': fn,
                'font_size': fs
            }
    if curr_run:
        merged_runs.append(curr_run)
        
    # Xóa các runs cũ và thêm runs mới
    p_elem = paragraph._p
    for r_elem in p_elem.findall(qn('w:r')):
        p_elem.remove(r_elem)
        
    for mr in merged_runs:
        r = paragraph.add_run(mr['text'])
        if mr['bold']:
            r.bold = True
        if mr['italic']:
            r.italic = True
        if mr['underline']:
            r.underline = True
        if mr['font_name']:
            r.font.name = mr['font_name']
        if mr['font_size']:
            r.font.size = mr['font_size']

def fix_table_floating(table):
    """
    Xóa thuộc tính tblpPr (bảng trôi) để bảng chữ ký thành bảng thẳng hàng (inline).
    Điều này đảm bảo mục 3 ("Hợp đồng được lập thành 02 bản...") LUÔN NẰM TRÊN bảng chữ ký,
    không bị Word đẩy nhảy xuống dưới.
    """
    tblPr = table._tbl.tblPr
    tblpPr = tblPr.find(qn('w:tblpPr'))
    if tblpPr is not None:
        tblPr.remove(tblpPr)

def strip_trailing_empty_paragraphs(doc):
    """
    Xóa toàn bộ các dòng trống ở cuối văn bản để tránh phát sinh trang trắng thừa.
    """
    while len(doc.paragraphs) > 0 and not doc.paragraphs[-1].text.strip():
        p_elem = doc.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)

def process_single_contract(template_path, data):
    """
    Điền dữ liệu cho 1 người vào mẫu hợp đồng.
    """
    doc = docx.Document(template_path)
    
    # 1. Thay thế trong paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p, data)
        
    # 2. Thay thế trong tables và loại bỏ floating table
    for table in doc.tables:
        fix_table_floating(table)
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, data)
                    
    # 3. Loại bỏ trang trắng thừa cuối văn bản
    strip_trailing_empty_paragraphs(doc)
    return doc

def merge_documents(template_path, records, output_merged_path):
    """
    Tạo file Word tổng hợp chứa tất cả hợp đồng của mọi người.
    ĐẢM BẢO:
    - Tất cả phần tử được chèn TRƯỚC sectPr để không làm hỏng cấu trúc file OpenXML
    - Mỗi người được phân tách bởi ngắt trang rõ ràng (pageBreak)
    - Không phát sinh trang trắng thừa
    """
    if not records:
        print("[!] Không có dữ liệu cộng tác viên trong bảng Excel để tạo hợp đồng.")
        return

    print(f"\n[+] Bắt đầu tạo file hợp đồng tổng hợp cho {len(records)} người...")
    master_doc = None
    
    for idx, record in enumerate(records):
        name = record.get('Hovaten') or record.get('hovaten') or f"Người {idx+1}"
        doc_i = process_single_contract(template_path, record)
        
        if master_doc is None:
            master_doc = doc_i
        else:
            sectPr = master_doc._body._body.find(qn('w:sectPr'))
            
            # Thêm ngắt trang chuẩn trước sectPr
            p_br = parse_xml(r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>')
            if sectPr is not None:
                sectPr.addprevious(p_br)
            else:
                master_doc._body._body.append(p_br)
                
            # Thêm các phần tử của doc_i trước sectPr
            for element in doc_i._body._body:
                if element.tag.endswith('sectPr'):
                    continue
                new_el = copy.deepcopy(element)
                if sectPr is not None:
                    sectPr.addprevious(new_el)
                else:
                    master_doc._body._body.append(new_el)
                
        progress = int(((idx + 1) / len(records)) * 100)
        print(f"  -> [{progress}%] Đã hoàn thành {idx+1}/{len(records)}: {name}")
        
    master_doc.save(output_merged_path)
    print(f"\n[THÀNH CÔNG] Đã lưu file hợp đồng gộp tại: {output_merged_path}")

def main():
    template = "Mẫu hợp đồng.docx"
    excel_file = "Mẫu điền hợp đồng.xlsx"
    output_merged = "HopDong_TongHop.docx"
    
    print("=" * 60)
    print("PHẦN MỀM TỰ ĐỘNG ĐIỀN HỢP ĐỒNG ABA (BẢN PYTHON)")
    print("=" * 60)
    
    try:
        records = read_excel(excel_file)
        merge_documents(template, records, output_merged)
    except Exception as e:
        print(f"\n[LỖI]: {e}")

if __name__ == "__main__":
    main()
