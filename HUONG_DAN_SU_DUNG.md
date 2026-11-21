# HƯỚNG DẪN SỬ DỤNG ỨNG DỤNG ĐIỀU HÀNH & KIỂM ĐIỂM ĐẢNG VIÊN CHI BỘ 2
**Đảng bộ Ngân hàng TMCP Công thương Việt Nam - Chi nhánh Bảo Lộc (Năm 2026)**

---

## 🌟 GIỚI THIỆU TỔNG QUAN
Ứng dụng được thiết kế chuyên biệt phục vụ kỳ **Kiểm điểm, đánh giá, xếp loại chất lượng tổ chức Đảng và đảng viên năm 2026** (quy mô 12 đảng viên) và mở rộng cho các kỳ sinh hoạt chi bộ thường kỳ tại Chi bộ 2 thuộc Đảng bộ VietinBank Chi nhánh Bảo Lộc.

Ứng dụng hoạt động theo mô hình **Web App hiện đại (Zero-Dependency)**:
- Chạy trực tiếp trên máy tính phòng họp và điện thoại thông minh của 12 đảng viên.
- Không cần cài đặt phần mềm phức tạp, lưu trữ an toàn, bảo mật nội bộ.
- Phân quyền chặt chẽ: **Mỗi đảng viên chỉ được bỏ phiếu 01 lần duy nhất và thực hiện đúng quyền hạn theo Điều lệ Đảng**.
- Tự động hóa 100% khâu tính điểm, nhận xét chéo, bỏ phiếu kín, xuất **Bản tự chấm điểm 100đ** và **Biên bản kiểm điểm chuẩn Nghị định 30/2020/NĐ-CP**.

---

## 🚀 CÁCH KHỞI CHẠY ỨNG DỤNG (1-CLICK)

### Cách 1: Khởi chạy máy chủ nội bộ (Khuyến nghị để 12 đồng chí cùng vào bằng điện thoại)
1. Nhấp đúp vào file **`Chay_App.bat`**.
2. Ứng dụng sẽ tự động kích hoạt máy chủ nội bộ và mở trình duyệt web tại địa chỉ:
   - **Trên máy tính Bí thư:** `http://localhost:8080`
   - **Trên điện thoại Đảng viên qua Wi-Fi:** Quét mã QR hiển thị trên màn hình hoặc truy cập qua địa chỉ IP LAN được cấp.

### Cách 2: Mở trực tiếp không cần máy chủ (Offline)
- Nhấp đúp trực tiếp vào file **`app_kiem_diem_chi_bo_2.html`** hoặc **`index.html`** để mở trên trình duyệt web (Chrome, Edge, Cốc Cốc).

---

## 🔐 MA TRẬN PHÂN QUYỀN & NGUYÊN TẮC BỎ PHIẾU KÍN

### 1. Phân quyền theo chức danh Đảng:
| Vai trò / Đối tượng | Quyền tự kiểm điểm (Tab 3) | Quyền nhận xét chéo (Tab 4) | Quyền bỏ phiếu kín (Tab 5) | Quyền khóa hòm phiếu & kiểm phiếu |
|---|:---:|:---:|:---:|:---:|
| **Bí thư Chi bộ** (Chủ trì toàn quyền) | Toàn quyền (Tập thể & Cá nhân) | Toàn quyền điều phối phân công | **01 phiếu** | **Có (Độc quyền)** |
| **Phó Bí thư / Chi ủy viên** | Tự chấm & xem tập thể | Nhận xét theo phân công | **01 phiếu** | Có (khi được ủy quyền) |
| **Đảng viên chính thức** | Tự chấm & xem tập thể | Nhận xét 2–3 người được phân công | **01 phiếu duy nhất** | Không |
| **Đảng viên dự bị** | Tự chấm điểm, tự phê bình | Tham gia đóng góp ý kiến | **KHÔNG** *(Điều 3 Điều lệ Đảng)* | Không |

### 2. Nguyên tắc Bỏ phiếu kín "Mỗi đảng viên chỉ bỏ 1 lần":
1. **Kiểm soát 2 tầng (Client & Server):**
   - **Tầng Client:** Ngay khi gửi phiếu thành công, form chọn bị khóa vĩnh viễn, hiển thị thẻ chứng nhận màu xanh: *"Đồng chí đã hoàn thành bỏ phiếu kín lúc [thời gian]"*. Nút gửi phiếu bị ẩn/vô hiệu hóa.
   - **Tầng Server:** Endpoint API `/api/vote` kiểm tra danh sách `voters` trên máy chủ. Nếu phát hiện gửi lần thứ hai, server từ chối ngay với mã lỗi `HTTP 409 (Conflict)` và thông báo: *"Đồng chí đã bỏ phiếu lúc [thời gian]. Mỗi đảng viên chỉ được quyền bỏ phiếu 01 lần duy nhất!"*.
2. **Khống chế tỷ lệ Xuất sắc (20%):**
   - Hệ thống tự động kiểm soát số phiếu *"Hoàn thành xuất sắc nhiệm vụ"* tối đa **02 đồng chí** trên tổng số 12 đảng viên của chi bộ. Bỏ quá 2 người sẽ bị chặn gửi.
3. **Quy định Đảng viên dự bị:**
   - Căn cứ **Điều 3 Điều lệ Đảng Cộng sản Việt Nam**: Đảng viên dự bị không có quyền biểu quyết/bỏ phiếu xếp loại. Hệ thống tự động vô hiệu hóa quyền bỏ phiếu của đảng viên dự bị và thông báo căn cứ pháp lý rõ ràng.
4. **Quyền Khóa hòm phiếu & Kiểm phiếu:**
   - Chỉ đồng chí **Bí thư Chi bộ** (hoặc Phó Bí thư) mới có quyền bấm nút *"Khóa Hòm Phiếu & Kiểm Phiếu Tức Thì"*. Đảng viên thông thường chỉ theo dõi tiến độ bỏ phiếu và xem kết quả sau khi đã công bố.

---

## 🔑 DANH SÁCH TÀI KHOẢN ĐĂNG NHẬP 12 ĐỒNG CHÍ CHI BỘ 2
*(Quy tắc: Tên đăng nhập là tên đảng viên viết liền không dấu, chữ thường • Mật khẩu mặc định: `123456`)*

| STT | Họ và Tên Đảng viên | Chức vụ Đảng | Chức vụ chuyên môn | Tên đăng nhập | Mật khẩu |
|:---:|---|---|---|:---:|:---:|
| 1 | **BÙI KHẮC DUY** | Đảng viên | Cán bộ | `buikhacduy` | `123456` |
| 2 | **VŨ HOÀNG NGUYÊN** | Đảng viên | Cán bộ | `vuhoangnguyen` | `123456` |
| 3 | **LÊ THANH NGHỊ** | Đảng viên | Cán bộ | `lethanhnghi` | `123456` |
| 4 | **HOÀNG THỊ THỦY** | Đảng viên | Cán bộ | `hoangthithuy` | `123456` |
| 5 | **NGUYỄN THỊ TỐ NGA** | Đảng viên | Cán bộ | `nguyenthitonga` | `123456` |
| 6 | **NGUYỄN THI THAO** | **Bí thư Chi bộ** | Phó Giám đốc CN | `nguyenthithao` | `123456` |
| 7 | **VÕ VĂN MẠNH** | **Chi ủy viên** | Cán bộ | `vovanmanh` | `123456` |
| 8 | **NGÔ THỊ KIM NGÂN** | **Phó Bí thư Chi bộ** | Trưởng phòng | `ngothikimngan` | `123456` |
| 9 | **NGUYỄN VĂN CHÂU** | Đảng viên | Cán bộ | `nguyenvanchau` | `123456` |
| 10 | **TRẦN ĐÌNH TIẾN DŨNG** | Đảng viên | Cán bộ | `trandinhtiendung` | `123456` |
| 11 | **ĐỖ THU HIỀN** | Đảng viên | Cán bộ | `dothuhien` | `123456` |
| 12 | **NGUYỄN DUY LỢI** | Đảng viên | Cán bộ | `nguyenduyloi` | `123456` |

---

## 📋 6 PHÂN HỆ NGHIỆP VỤ CỐT LÕI

### 1. Kỳ Họp & Điểm Danh (Tab 1)
- Thiết lập thông tin kỳ họp (Thời gian, địa điểm, Chủ trì, Thư ký).
- Điểm danh 12 đồng chí (Xác nhận Có mặt, Vắng có lý do, quyền biểu quyết).
- **Chế độ Trình chiếu (Host View):** Bí thư bấm nút "Trình chiếu" để phóng to toàn màn hình lên TV/Máy chiếu phòng họp.

### 2. Quản Lý & Import Danh Sách Đảng Viên (Tab 2)
- Tích hợp sẵn danh sách 12 đồng chí chuẩn VietinBank Bảo Lộc.
- Hỗ trợ **Import file Excel (`.xlsx`, `.csv`)** hoặc **Dán nhanh từ Excel (Ctrl+V)** 3 cột: `STT`, `Tên Đảng viên`, `Chức vụ`.
- Nút **"Xuất Mã PIN Gửi Zalo"**: Sao chép nhanh danh sách tài khoản kèm mã PIN để gửi riêng cho từng đồng chí.

### 3. Tự Kiểm Điểm & Kết Xuất Bản Chấm Điểm 100đ (Tab 3)
- Tự động tính điểm và phân loại xếp loại theo Khung tiêu chí 2026:
  + Tập thể Chi bộ 2 (30đ chung + 70đ chuyên môn/nhiệm vụ).
  + Cán bộ lãnh đạo, cấp ủy (Phụ lục 5).
  + Đảng viên không giữ chức vụ (Phụ lục 6).
- Form viết bài tự kiểm điểm 3 phần: *Ưu điểm - Hạn chế/Khuyết điểm - Phương hướng khắc phục*.
- **Bộ 3 công cụ kết xuất chuyên nghiệp:**
  + `📄 Xuất Bản Đang Chọn (.docx)`: Tải file Word của đối tượng đang xem.
  + `📦 Xuất Trọn Bộ (Tập Thể + 12 Cá Nhân)`: Xuất file Word gộp 13 bản tự chấm điểm hoàn chỉnh.
  + `🖨️ Xem Trước & In (PDF)`: Mô phỏng tờ A4 chuẩn thể thức Đảng để in hoặc lưu PDF.

### 4. Phân Công & Nhận Xét Chéo Online (Tab 4)
- **Thuật toán chia cân bằng:** Tự động phân công mỗi đồng chí nhận xét đúng 2 người và nhận đúng 2 lượt nhận xét từ người khác, tuyệt đối đồng bộ chuẩn xác tên cán bộ.
- Nhập góp ý 3 mục mang tính xây dựng, tự động lưu vết và đồng bộ vào Biên bản kỳ họp.

### 5. Bỏ Phiếu Kín Điện Tử (Tab 5)
- Bảo mật 100% lá phiếu, chỉ ghi nhận danh sách ai đã bỏ phiếu.
- **Ràng buộc cứng:** Mỗi đồng chí chỉ bỏ 1 lần duy nhất, khống chế không quá 2 người Xuất sắc.
- Nút **"Khóa Hòm Phiếu & Kiểm Phiếu Tức Thì"** (dành cho Bí thư) để chốt số liệu và công bố bảng kết quả kiểm phiếu.

### 6. Biên Bản Họp & Xuất File Word (Tab 6)
- Tự động lắp ráp toàn bộ dữ liệu thành Biên bản hoàn chỉnh chuẩn **Nghị định 30/2020/NĐ-CP**.
- Tích hợp nút **"Xuất File Word (.docx) Chuẩn NĐ 30"** và **"In Biên Bản (PDF)"**.

---

## 📁 DANH MỤC CÁC TẬP TIN DỰ ÁN
| Tên tập tin | Chức năng |
|---|---|
| `app_kiem_diem_chi_bo_2.html` | Giao diện Web App chính (FinTech 3D, VietinBank Navy & Red, responsive PC & Điện thoại) |
| `index.html` | Bản sao đồng bộ phục vụ Web Server gốc |
| `server.py` | Máy chủ nội bộ Python (API dữ liệu, API bỏ phiếu an toàn, API xuất Word) |
| `Chay_App.bat` | Trình khởi chạy 1-click tự động |
| `ChiaSe_Mang_LAN.bat` | Trình tạo cổng chia sẻ mạng Wi-Fi LAN nội bộ |
| `xuat_ban_tu_cham_diem.py` | Module tạo file Word bản tự chấm điểm Tập thể & Cá nhân |
| `xuat_bien_ban_word.py` | Module tạo file Word Biên bản kiểm điểm chuẩn Nghị định 30 |
| `UngDung_KiemDiem_ChiBo2_VietinBank.zip` | Gói nén trọn bộ sẵn sàng gửi qua Zalo / Email nội bộ |
| `Ma_QR_ChiBo2_VietinBank.png` | Mã QR Code kết nối nhanh qua Wi-Fi |
| `Mau_Import_Danh_Sach_Dang_Vien.xlsx` | File Excel mẫu danh sách 12 đảng viên |
| `criteria_data.json` | Bộ tiêu chí đánh giá Khung 2026 chính thức |
