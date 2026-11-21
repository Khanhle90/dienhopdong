# PHẦN MỀM TỰ ĐỘNG ĐIỀN HỢP ĐỒNG TỪ EXCEL SANG WORD & PDF

Ứng dụng web tự động hóa quy trình tạo hợp đồng hàng loạt từ bảng tính Excel hoặc Google Sheets sang mẫu Word chuẩn, tự động ghép nối thành file PDF hoàn chỉnh hợp quy cách, loại bỏ trang trắng thừa và hỗ trợ in/tải về ngay trên trình duyệt.

> **Đặc điểm nổi bật:**
> - 🚀 **Chạy trực tiếp trên GitHub Pages**: Hoàn toàn miễn phí, không cần thuê máy chủ hay cài đặt môi trường phức tạp.
> - 🔒 **Bảo mật tuyệt đối**: Dữ liệu cá nhân (CCCD, Số điện thoại, Thông tin công tác viên) được xử lý 100% cục bộ trên trình duyệt người dùng, không gửi dữ liệu ra máy chủ bên ngoài.
> - 📄 **Bảo toàn định dạng văn bản**: Giữ nguyên in đậm, viết hoa, căn giữa, căn đều 2 bên, bảng ký 2 cột chuẩn thể thức văn bản hành chính Việt Nam (Nghị định 30/2020/NĐ-CP).
> - 🔢 **Giữ nguyên số 0 ở đầu**: Tự động bảo toàn các số 0 ở đầu CCCD (`0012...`), Số điện thoại (`0978...`) và chuẩn hóa các giá trị tính toán từ hàm Excel.
> - 🔄 **Linh hoạt thay thế mẫu**: Có sẵn mẫu hợp đồng Cộng tác viên ABA, đồng thời cho phép người dùng tải lên mẫu Word (.docx) mới để tái sử dụng cho các hợp đồng khác sau này.

---

## 📋 CÁC QUY TẮC XỬ LÝ DỮ LIỆU ĐƯỢC THIẾT LẬP SẴN

1. **Quy tắc nhận diện Sheet:**
   - Phần mềm tự động tìm Sheet có tên là **`Thông tin`** (hỗ trợ cả chữ hoa, chữ thường, có dấu hoặc không dấu).
   - Nếu không có Sheet `Thông tin`, phần mềm sẽ tự động lấy dữ liệu từ **`Sheet 1`** (hoặc `Sheet1`).
   - Nếu cả 2 tên Sheet trên đều không tồn tại, phần mềm sẽ hiển thị thông báo lỗi chính xác: **`Tên sheet không đúng quy tắc`**.

2. **Quy tắc lấy dữ liệu các cột:**
   - Dữ liệu luôn được lấy từ **Cột B trở đi** (bỏ qua Cột A như Mã CTV / STT).
   - **Hàng 1**: Sử dụng làm từ khóa định danh để ghép với các trường `{...}` trong mẫu Word.
   - **Hàng 2 trở đi**: Dữ liệu của từng người. Mỗi dòng tương ứng với 1 người, tạo thành 1 hợp đồng riêng biệt.
   - Các giá trị sinh ra từ **Hàm/Công thức Excel** (VLOOKUP, CONCATENATE...) vẫn được đọc chuẩn xác theo giá trị hiển thị.

3. **Quy tắc điền ô trống & định dạng chữ:**
   - Các ô `thulao`, `congviec`, `thoihan` nếu bảng không có thông tin sẽ tự động điền `............................................`
   - Các ô `ngay`, `thang`, `nam` (và `ngaycap`) nếu bỏ trống sẽ tự động điền `.....`
   - Vị trí thay thế được **bôi đen** trong mẫu Word thì kết quả điền vào tự động được **bôi đen**.
   - Vị trí thay thế viết hoa (như trường `{CTV}`) thì kết quả tự động được **viết hoa**.
   - Toàn bộ dấu ngoặc `{}` được xóa sạch sau khi điền.
   - Tự động cắt bỏ các dòng trống thừa ở cuối văn bản để **loại bỏ trang trắng không có text**.

---

## 🛠️ HƯỚNG DẪN CÀI ĐẶT & TRIỂN KHAI TRÊN GITHUB (2 PHÚT)

Phần mềm được thiết kế dưới dạng **Single Page Application (SPA)** tĩnh, do đó việc triển khai lên GitHub cực kỳ đơn giản:

### Bước 1: Đưa mã nguồn lên GitHub Repository
1. Đăng nhập vào tài khoản [GitHub](https://github.com).
2. Tạo một Repository mới (ví dụ đặt tên: `hop-dong-tu-dong`), chọn chế độ **Public**.
3. Tải toàn bộ các file trong thư mục này lên Repository (bao gồm `index.html`, `app.js`, `style.css`, thư mục `lib/`, v.v.).

### Bước 2: Kích hoạt GitHub Pages
1. Tại giao diện Repository trên GitHub, bấm vào tab **Settings** (Cài đặt).
2. Ở menu bên trái, tìm và chọn mục **Pages**.
3. Tại phần **Build and deployment**:
   - **Source**: Chọn `Deploy from a branch`.
   - **Branch**: Chọn nhánh `main` (hoặc `master`), thư mục giữ nguyên `/ (root)`.
   - Bấm **Save** (Lưu).
4. Đợi khoảng 1 - 2 phút, GitHub sẽ cung cấp cho bạn một đường link trang web hoạt động 24/7, có dạng:
   ```
   https://<ten-tai-khoan-github>.github.io/hop-dong-tu-dong/
   ```
5. Bạn có thể mở đường link này trên bất kỳ máy tính, điện thoại, máy tính bảng nào để sử dụng ngay mà không cần cài đặt thêm bất kỳ phần mềm nào khác!

---

## 💻 HƯỚNG DẪN CHẠY TRỰC TIẾP TRÊN MÁY TÍNH (OFFLINE)

Nếu không có mạng Internet hoặc muốn chạy ngay trên máy tính cá nhân:
1. Tải toàn bộ thư mục về máy tính.
2. Nhấp đúp chuột vào file **`index.html`** để mở bằng Google Chrome, Microsoft Edge, Cốc Cốc, Firefox hoặc Safari.
3. Phần mềm đã tích hợp sẵn thư viện nội bộ và mẫu hợp đồng mặc định nên chạy mượt mà ngay cả khi ngắt kết nối mạng.

---

## 🐍 DÀNH CHO LẬP TRÌNH VIÊN: CHẠY BẰNG SCRIPT PYTHON

Nếu bạn muốn chạy dạng xử lý dòng lệnh (CLI) hoặc tích hợp vào hệ thống backend:
1. Cài đặt các thư viện phụ thuộc:
   ```bash
   pip install -r requirements.txt
   ```
2. Chạy tệp xử lý:
   ```bash
   python process_contracts.py
   ```
3. Kết quả: Tệp **`HopDong_TongHop.docx`** chứa tất cả hợp đồng của mọi người (phân tách trang rõ ràng) sẽ được tạo ra tại thư mục hiện tại.

---

## 📖 HƯỚNG DẪN SỬ DỤNG PHẦN MỀM

1. **Bước 1 - Kiểm tra hoặc thay đổi Mẫu hợp đồng Word (.docx):**
   - Ứng dụng tự động nạp sẵn mẫu **Hợp đồng Cộng tác viên ABA**.
   - Bấm nút **"Tải file mẫu (.docx)"** nếu bạn muốn xem hoặc sửa đổi mẫu trên máy.
   - Nếu bạn có mẫu hợp đồng khác, chỉ cần kéo thả file `.docx` mới vào khung ở Bước 1.

2. **Bước 2 - Chọn dữ liệu nguồn:**
   - **Cách 1 (File Excel):** Kéo thả file `.xlsx`, `.xls` hoặc `.csv` vào khung.
   - **Cách 2 (Google Sheet):** Mở bảng tính Google Sheet của bạn &rarr; Bấm nút **Chia sẻ** ở góc phải &rarr; Chọn quyền **"Bất kỳ ai có đường liên kết đều có thể xem"** &rarr; Dán link vào ô nhập và bấm **"Tải dữ liệu từ Google Sheet"**.
   - *(Mẹo nhanh)*: Bấm nút **"Thử dữ liệu mẫu (27 người)"** để kiểm tra tính năng ngay lập tức mà không cần chuẩn bị file.

3. **Bước 3 - Bắt đầu tạo hợp đồng:**
   - Bấm nút **"🚀 Bắt đầu tạo Hợp đồng"**.
   - Theo dõi thanh tiến độ hiển thị trực tiếp số lượng người và tên người đang được tạo.

4. **Bước 4 - In hoặc Tải về:**
   - Khi hoàn thành, cửa sổ thông báo sẽ hiện ra với 2 chức năng chính:
     - 🖨️ **In (Print) / Lưu PDF sắc nét**: Mở trực tiếp hộp thoại in của trình duyệt (chọn máy in hoặc chọn *Lưu dưới dạng PDF* - chất lượng vector sắc nét 100%).
     - 📥 **Tải về toàn bộ file PDF gộp**: Tải về tệp PDF nối tiếp tất cả hợp đồng.
     - 📄 **Tải file Word (.docx)**: Tải tệp Word tổng hợp đã gộp tất cả người để chỉnh sửa thêm nếu muốn.
     - 👁️ **Xem trước trực quan**: Xem lại từng trang hợp đồng ngay trên màn hình.

---

## 📂 DANH MỤC TỆP TRONG DỰ ÁN

| Tệp / Thư mục | Mục đích |
| :--- | :--- |
| `index.html` | Giao diện web ứng dụng hiện đại, tương thích mọi thiết bị |
| `app.js` | Logic xử lý đọc Excel/GSheet, quét placeholder, điền dữ liệu, xuất PDF & Word |
| `style.css` | Tùy biến giao diện, hiệu ứng chuyển động và căn chuẩn lề trang in A4 |
| `default_template_b64.js` | Mẫu hợp đồng chuẩn được mã hóa Base64 giúp chạy offline tức thì |
| `sample_data.js` | Dữ liệu mẫu 27 cộng tác viên trích xuất từ sheet Thông tin |
| `lib/` | Thư viện ngoại tuyến (SheetJS, JSZip, html2pdf) đảm bảo không phụ thuộc mạng |
| `Mẫu hợp đồng.docx` | File Word mẫu gốc do người dùng cung cấp |
| `Mẫu điền hợp đồng.xlsx` | File Excel mẫu gốc chứa 27 người ở sheet Thông tin |
| `process_contracts.py` | Mã nguồn Python xử lý tự động hóa dành cho môi trường máy chủ |
| `requirements.txt` | Danh sách gói Python cần thiết |
