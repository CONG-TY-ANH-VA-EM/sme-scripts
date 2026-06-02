# Mẫu Phiếu Lương (tham chiếu thẻ chuẩn) — SME SOLUTIONS 2025

> ⚠️ **QUAN TRỌNG:** Đây là các thẻ `{{...}}` mà công cụ **thực sự nhận diện** (VIẾT HOA, theo `MAP` trong code). Bản cũ từng dùng thẻ kiểu `{{HoTen}}`, `{{TongLuong24}}` — **những thẻ đó KHÔNG hoạt động** và sẽ làm phiếu bị trống. Hãy dùng đúng các thẻ dưới đây. Bảng tra đầy đủ 41 cột: xem [CAM_NANG_CHI_TIET.md › Phần M](../google-appscripts/CAM_NANG_CHI_TIET.md#phần-m--bảng-tra-thẻ).

---

{{SENDER_NAME}}
{{SENDER_ADDRESS}}
Hotline: {{SENDER_HOTLINE}} | Email: {{CONTACT_EMAIL}}

---

## PHIẾU LƯƠNG {{THANGNAM}}

### THÔNG TIN NHÂN VIÊN
- **Họ và tên**: {{HOTEN}} (STT: {{SOTT}})
- **Vị trí**: {{VITRI}}
- **Ngày gửi**: {{NGAYGUI}}
- **STK nhận**: {{STK}} ({{NGANHANG}})
- **Email**: {{EMAIL}}

---

### CHI TIẾT LƯƠNG & CÔNG XÁ
- **Số ngày công chuẩn**: {{NGAYCONGCHUAN}}
- **Số giờ làm việc thực tế**: {{TONGGIOTT}}
- **OT 150%/200%/300%**: {{GIO150}} / {{GIO200}} / {{GIO300}}
- **Nghỉ: Ro/N/P/L**: {{RO}} / {{N}} / {{P}} / {{L}}

### THU NHẬP
- **Lương cơ bản**: {{L_COBAN}}
- **Phụ cấp: Ăn trưa/Điện thoại/Đi lại**: {{PC_ANTRUA}} / {{PC_DIENTHOAI}} / {{PC_DILAI}}
- **Lương ngày công thực tế**: {{L_NGAYCONG}}
- **Lương ngày nghỉ phép/lễ**: {{L_NGHIPHEP}}
- **Lương ngoài giờ OT**: {{L_OT}}
- **Lương KPI**: {{L_KPI}}
- **Lương phép năm**: {{L_PHEPNAM}}
- **Phụ cấp Máy tính/Gửi xe**: {{PC_MAYTINH}}
- **Công tác phí**: {{CONGTACPHI}}
- **Thưởng/Phúc lợi**: {{THUONG}}
- **Các khoản trừ (Chịu thuế/Không chịu thuế)**: {{TRU_CHIU_THUE}} / {{TRU_KG_CHIU_THUE}}

**TỔNG LƯƠNG**: {{TONGLUONG}}

### KHẤU TRỪ & THỰC LĨNH
- **Giảm trừ gia cảnh**: {{GIAMTRU}}
- **Bảo hiểm xã hội**: {{BHXH}}
- **Thuế TNCN**: {{THUETNCN}}
- **THỰC LĨNH**: {{THUCLINH}}

- **Phạt/Cọc / Tạm ứng**: {{PHAT}} / {{TAMUNG}}
- **Đã quyết toán/Đã nhận**: {{DANHAN}}
- **Ngày phép còn lại**: {{PHEPCONLAI}}
- **SỐ TIỀN CHUYỂN KHOẢN**: {{LUONGCK}}

---
*Mọi thắc mắc vui lòng liên hệ Phòng Nhân sự trong vòng 3 ngày làm việc.*

---

### Ghi chú cho người thiết kế mẫu
- Các thẻ tiền tệ (`{{L_COBAN}}`, `{{TONGLUONG}}`, `{{THUCLINH}}`, `{{LUONGCK}}`…) sẽ được công cụ **tự định dạng** kiểu `1.000.000` (làm tròn, không thập phân).
- `{{THANGNAM}}`, `{{SENDER_NAME}}`, `{{SENDER_ADDRESS}}`, `{{SENDER_HOTLINE}}`, `{{CONTACT_EMAIL}}` lấy từ tab **CONFIG** (không phải từ cột dữ liệu nhân viên).
- Muốn có sẵn khung mẫu chuẩn A4 với đúng các thẻ này: trong Google Sheets vào menu **SME Tools → Tạo Sheet Mẫu (TEMPLATE)**.
