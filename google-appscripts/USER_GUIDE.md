# Hướng dẫn Sử dụng Hệ thống Gửi Phiếu Lương Tự động 2025

Chào mừng bạn đến với phiên bản nâng cấp 2025 của công cụ gửi phiếu lương. Phiên bản này hỗ trợ cơ chế **Template Sheet**, cho phép bạn tự thiết kế mẫu phiếu lương chuyên nghiệp ngay trên Google Sheets.

## 1. Chuẩn bị Bảng tính lương
Bảng lương của bạn cần có cấu trúc 40 cột để khớp với các thẻ Tag có sẵn.

**Cách nhanh nhất**:
1. Tải file mẫu [payroll_sample_2025.csv](./payroll_sample_2025.csv).
2. Mở Google Sheets > **Tệp (File)** > **Nhập (Import)** > **Tải lên** file CSV này.
3. Đổi tên sheet vừa nhập thành `T + <số tháng>` (Ví dụ: `T1`).

*Lưu ý: Dữ liệu nhân viên bắt đầu từ hàng 2.*

## 2. Thiết kế Mẫu Phiếu Lương (Tab TEMPLATE)
Thay vì vẽ bằng code hoặc tự kẻ bảng thủ công, giờ đây bạn có thể khởi tạo nhanh:

1. Mở menu **SME Tools** > **Tạo Sheet Mẫu (TEMPLATE)**.
2. Script sẽ tự động tạo một Tab tên là `TEMPLATE` với đầy đủ định dạng chuyên nghiệp và các thẻ Tag chuẩn.
3. Bạn có thể thay đổi logo, font chữ hoặc di chuyển các ô theo ý thích, miễn là giữ nguyên các thẻ **{{TAG}}**.

> [!TIP]
> **Mẹo thiết kế cho A4:** Để phiếu lương PDF trông vừa vặn nhất, bạn nên giữ tổng độ rộng các cột trong khoảng **700px - 750px**. Script v1.5 mặc định tạo mẫu với độ rộng 300px + 200px + 200px = 700px, rất chuẩn cho khổ A4 Portrait.

### Danh sách các thẻ (Tags) khả dụng:
- **Thông tin chung**: `{{THANGNAM}}`, `{{SENDER_NAME}}`, `{{SENDER_ADDRESS}}`, `{{SENDER_HOTLINE}}`, `{{CONTACT_EMAIL}}`
- **Thông tin nhân viên**: `{{HOTEN}}`, `{{VITRI}}`, `{{STK}}`, `{{NGANHANG}}`, `{{NGAYGUI}}`
- **Công xá**: `{{NGAYCONGCHUAN}}`, `{{TONGGIOTT}}`, `{{GIO150}}`, `{{GIO200}}`, `{{GIO300}}`, `{{RO}}`, `{{P}}`, `{{L}}`
- **Lương & Phụ cấp**: `{{L_COBAN}}`, `{{PC_ANTRUA}}`, `{{PC_DIENTHOAI}}`, `{{PC_DILAI}}`, `{{L_NGAYCONG}}`, `{{L_NGHIPHEP}}`, `{{L_OT}}`, `{{L_KPI}}`, `{{PC_MAYTINH}}`, `{{CONGTACPHI}}`, `{{THUONG}}`
- **Khấu trừ & Thực lĩnh**: `{{BHXH}}`, `{{THUETNCN}}`, `{{TONGLUONG}}`, `{{THUCLINH}}`, `{{TAMUNG}}`, `{{DANHAN}}`, `{{LUONGCK}}`

> [!TIP]
> Script sẽ tự động định dạng số tiền (thêm dấu chấm phân cách hàng nghìn) cho các thẻ liên quan đến lương thưởng.

---

## 3. Cài đặt Script
1. Mở Sheet > **Tiện ích mở rộng (Extensions)** > **Apps Script**.
2. Copy toàn bộ nội dung file `salary_send.js` vào trình soạn thảo.
3. Nhấn **Lưu** và làm mới trang Google Sheets.

## 4. Cách chạy Gửi Phiếu Lương
1. Bấm vào menu **SME Tools** > **Gửi Phiếu Lương PDF (Tháng Hiện Tại)**.
2. Cấp quyền (lần đầu) và chờ thông báo kết quả.

---

## 5. Cơ chế Tự động chạy lại (Resilience)
Phiên bản v1.3 hỗ trợ gửi email cho danh sách rất lớn (hàng trăm/nghìn người) mà không lo bị lỗi giới hạn thời gian:

- **Trạng thái Gửi**: Script sẽ tự động cập nhật cột cuối cùng (mặc định là `AO`) thành "Thành công" sau khi gửi xong.
- **Tự động Resume**: Nếu danh sách quá dài, script sẽ tự động tạm dừng sau 5 phút và đặt lịch hẹn giờ chạy tiếp phần còn lại sau 1 phút.
- **Quota Gmail**: Nếu hết hạn mức gửi của Google trong ngày, script sẽ dừng lại và bảo vệ tài khoản của bạn.
- **Báo cáo Nhật ký tóm tắt**: Ngay sau khi hoàn tất hoặc tạm dừng, script sẽ gửi một email báo cáo tới chính bạn, liệt kê danh sách các nhân viên bị lỗi (nếu có) để bạn tiện theo dõi và xử lý.

## 6. Lưu ý về Hạn mức (Email Quotas)
Google áp đặt giới hạn số lượng email bạn có thể gửi trong 24 giờ thông qua Script. Con số này thay đổi tùy loại tài khoản:

| Loại Tài Khoản | Hạn mức Gửi (Recipients/ngày) |
| :--- | :--- |
| **Gmail miễn phí (@gmail.com)** | ~100 người/ngày |
| **Google Workspace (Trả phí)** | ~1.500 - 2.000 người/ngày |
| **Workspace (Dùng thử)** | ~500 người/ngày |

- **Cơ chế báo hết**: Script v1.3 sử dụng hàm `MailApp.getRemainingDailyQuota()` để kiểm tra hạn mức còn lại trước khi gửi mỗi email.
- **Xử lý khi hết**: Nếu số lượng nhân viên lớn hơn hạn mức (ví dụ bạn có 150 nhân viên nhưng dùng Gmail miễn phí), script sẽ gửi xong 100 người, sau đó dừng lại và báo lỗi Quota. Bạn chỉ cần đợi 24h sau rồi nhấn gửi lại, script sẽ tự động gửi nốt cho 50 người còn lại nhờ cột **SentStatus**.

---
*Phát triển bởi: AI Assistant - SME SOLUTIONS*
