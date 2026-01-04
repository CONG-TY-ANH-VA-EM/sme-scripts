/**
 * Tên hàm: sendMonthlyPayrollEmails_WithPDF
 * Mô tả: Đọc dữ liệu từ Google Sheet bảng lương (tên sheet động theo tháng, không có header),
 *        tạo file PDF phiếu lương cho mỗi nhân viên và gửi email đính kèm PDF.
 *        Dữ liệu nhân viên từ hàng 9 đến hàng 35.
 *        Cột Tên Nhân Viên: B
 *        Cột Email: CL
 *        Cột Tổng Lương: CJ
 * Được phát triển bởi: AI Assistant - SME SOLUTIONS
 */
/**
 * =============================================================================
 *                      CẤU HÌNH (CONFIGURATION)
 * =============================================================================
 * Bạn có thể thay đổi các tham số dưới đây để phù hợp với file Sheet của mình.
 * HOẶC: Bạn có thể tạo 1 Tab tên là "CONFIG" trong Google Sheets để lưu cấu hình.
 */
const GLOBAL_CONFIG = {
    // 1. Vị trí dữ liệu (Mặc định cho file CSV mẫu 2025)
    START_ROW: 2,              // Hàng bắt đầu có dữ liệu (Hàng 1 là Header)
    COL_EMAIL: "F",            // Cột Email
    COL_NAME: "B",             // Cột Tên

    // 2. Thông tin Công ty (Anonimized)
    SENDER_NAME: "SME SOLUTIONS JOINT STOCK COMPANY",
    SENDER_ADDRESS: "Tòa nhà Alpha, số 123 Đường Beta, Quận Gamma, Hà Nội",
    SENDER_HOTLINE: "1900 xxxx",
    CONTACT_EMAIL: "hr@sme-solutions.vn",

    // 3. Mapping cột cho template 2025 (Cột A=1, B=2...)
    MAP: {
        SOTT: "A", HOTEN: "B", VITRI: "C", NGAYGUI: "D", STK: "E", EMAIL: "F", NGANHANG: "G",
        NGAYCONGCHUAN: "H", TONGGIOTT: "I", GIO150: "J", GIO200: "K", GIO300: "L",
        RO: "M", N: "N", P: "O", L: "P",
        L_COBAN: "Q", PC_ANTRUA: "R", PC_DIENTHOAI: "S", PC_DILAI: "T",
        L_NGAYCONG: "U", L_NGHIPHEP: "V", L_OT: "W", L_KPI: "X", L_PHEPNAM: "Y",
        PC_MAYTINH: "Z", CONGTACPHI: "AA", THUONG: "AB", TRU_CHIU_THUE: "AC", TRU_KG_CHIU_THUE: "AD",
        TONGLUONG: "AE", PHEPCONLAI: "AF", GIAMTRU: "AG", BHXH: "AH", THUETNCN: "AI",
        THUCLINH: "AJ", PHAT: "AK", TAMUNG: "AL", DANHAN: "AM", LUONGCK: "AN"
    },

    // 4. Tiêu đề và File
    EMAIL_SUBJECT_PREFIX: "Phiếu Lương",
    SHEET_NAME_PREFIX: "T",
    PDF_FILE_NAME_PREFIX: "PhieuLuong_2025"
};

/**
 * Hàm chính để gửi Email bảng lương.
 */
function sendMonthlyPayrollEmails_WithPDF() {
    const settings = getSettings();
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;
    const currentYear = currentDate.getFullYear();
    const payrollMonthYear = `Tháng ${currentMonth}/${currentYear}`;
    const sheetName = settings.SHEET_NAME_PREFIX + currentMonth;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sheetName);

    if (!sourceSheet) {
        SpreadsheetApp.getUi().alert("Lỗi", `Không tìm thấy sheet "${sheetName}".`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const startRow = settings.START_ROW;
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < startRow) {
        SpreadsheetApp.getUi().alert("Thông báo", "Không có dữ liệu.");
        return;
    }

    // Lấy toàn bộ dữ liệu bảng lương
    const dataRange = sourceSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0]; // Giả định hàng 1 là headers

    let emailsSentCount = 0;
    let emailsFailedCount = 0;

    // Tạo helper function để lấy value theo tên cột (Letter)
    const getVal = (rowValues, colLetter) => {
        const idx = columnLetterToIndex(colLetter);
        return rowValues[idx];
    };

    const formatVND = (val) => {
        if (!val || isNaN(parseFloat(val))) return "0";
        return new Intl.NumberFormat('vi-VN').format(parseFloat(val));
    };

    let tempSpreadsheet = SpreadsheetApp.create(`Temp_Payroll_${new Date().getTime()}`);
    let tempSheet = tempSpreadsheet.getSheets()[0];

    try {
        for (let i = startRow - 1; i < values.length; i++) {
            const row = values[i];
            const recipientEmail = getVal(row, settings.COL_EMAIL);
            const employeeName = getVal(row, settings.COL_NAME);

            if (!recipientEmail || !String(recipientEmail).includes('@')) continue;

            try {
                tempSheet.clear();
                tempSheet.clearFormats();
                tempSheet.setColumnWidth(1, 200);
                tempSheet.setColumnWidth(2, 150);
                tempSheet.setColumnWidth(3, 150);

                // --- VẼ PHIẾU LƯƠNG ---
                // Header Công ty
                tempSheet.getRange("A1:C1").merge().setValue(settings.SENDER_NAME).setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
                tempSheet.getRange("A2:C2").merge().setValue(settings.SENDER_ADDRESS).setFontSize(9).setHorizontalAlignment("center");
                tempSheet.getRange("A3:C3").merge().setValue(`Hotline: ${settings.SENDER_HOTLINE} | Email: ${settings.CONTACT_EMAIL}`).setFontSize(9).setHorizontalAlignment("center");

                // Tiêu đề Phiếu Lương
                tempSheet.getRange("A5:C5").merge().setValue(`PHIẾU LƯƠNG ${payrollMonthYear.toUpperCase()}`).setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center").setBackground("#f3f3f3");

                // 1. THÔNG TIN NHÂN VIÊN
                tempSheet.getRange("A7").setValue("THÔNG TIN NHÂN VIÊN").setFontWeight("bold").setBackground("#e9e9e9");
                tempSheet.getRange("A8").setValue("Họ và tên:"); tempSheet.getRange("B8:C8").merge().setValue(employeeName).setFontWeight("bold");
                tempSheet.getRange("A9").setValue("Vị trí:"); tempSheet.getRange("B9:C9").merge().setValue(getVal(row, settings.MAP.VITRI));
                tempSheet.getRange("A10").setValue("Số tài khoản:"); tempSheet.getRange("B10:C10").merge().setValue(`${getVal(row, settings.MAP.STK)} (${getVal(row, settings.MAP.NGANHANG)})`);

                // 2. CHI TIẾT CÔNG XÁ
                tempSheet.getRange("A12").setValue("CHI TIẾT CÔNG XÁ").setFontWeight("bold").setBackground("#e9e9e9");
                tempSheet.getRange("A13").setValue("Công chuẩn / Thực tế:"); tempSheet.getRange("B13").setValue(getVal(row, settings.MAP.NGAYCONGCHUAN)); tempSheet.getRange("C13").setValue(getVal(row, settings.MAP.TONGGIOTT) + " giờ");
                tempSheet.getRange("A14").setValue("Tăng ca (150%/200%/300%):"); tempSheet.getRange("B14:C14").merge().setValue(`${getVal(row, settings.MAP.GIO150)} / ${getVal(row, settings.MAP.GIO200)} / ${getVal(row, settings.MAP.GIO300)}`);
                tempSheet.getRange("A15").setValue("Nghỉ (Ro/N/P/L):"); tempSheet.getRange("B15:C15").merge().setValue(`${getVal(row, settings.MAP.RO)} / ${getVal(row, settings.MAP.N)} / ${getVal(row, settings.MAP.P)} / ${getVal(row, settings.MAP.L)}`);

                // 3. CHI TIẾT THU NHẬP
                tempSheet.getRange("A17").setValue("CHI TIẾT THU NHẬP").setFontWeight("bold").setBackground("#e9e9e9");
                const incomeData = [
                    ["Lương cơ bản (HĐ)", formatVND(getVal(row, settings.MAP.L_COBAN))],
                    ["Phụ cấp (Ăn/ĐT/NL)", `${formatVND(getVal(row, settings.MAP.PC_ANTRUA))} / ${formatVND(getVal(row, settings.MAP.PC_DIENTHOAI))} / ${formatVND(getVal(row, settings.MAP.PC_DILAI))}`],
                    ["Lương ngày công thực tế", formatVND(getVal(row, settings.MAP.L_NGAYCONG))],
                    ["Lương OT / KPI", `${formatVND(getVal(row, settings.MAP.L_OT))} / ${formatVND(getVal(row, settings.MAP.L_KPI))}`],
                    ["Thưởng / Phúc lợi / Khác", formatVND(getVal(row, settings.MAP.THUONG))],
                    ["TỔNG THU NHẬP (24)", formatVND(getVal(row, settings.MAP.TONGLUONG))]
                ];
                let rowIdx = 18;
                incomeData.forEach(d => {
                    tempSheet.getRange(rowIdx, 1).setValue(d[0]);
                    tempSheet.getRange(rowIdx, 2, 1, 2).merge().setValue(d[1]).setHorizontalAlignment("right");
                    if (d[0].includes("TỔNG")) tempSheet.getRange(rowIdx, 1, 1, 3).setFontWeight("bold");
                    rowIdx++;
                });

                // 4. GIẢM TRỪ & THỰC LĨNH
                rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("GIẢM TRỪ & KHẤU TRỪ").setFontWeight("bold").setBackground("#e9e9e9"); rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("BHXH / Thuế TNCN"); tempSheet.getRange(rowIdx, 2, 1, 2).merge().setValue(`${formatVND(getVal(row, settings.MAP.BHXH))} / ${formatVND(getVal(row, settings.MAP.THUETNCN))}`).setHorizontalAlignment("right"); rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("THỰC LĨNH (29)").setFontWeight("bold"); tempSheet.getRange(rowIdx, 2, 1, 2).merge().setValue(formatVND(getVal(row, settings.MAP.THUCLINH))).setFontWeight("bold").setHorizontalAlignment("right"); rowIdx++;

                // 5. THANH TOÁN CUỐI CÙNG
                rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("THANH TOÁN CUỐI CÙNG").setFontWeight("bold").setBackground("#d9ead3"); rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("Tạm ứng / Đã nhận:"); tempSheet.getRange(rowIdx, 2, 1, 2).merge().setValue(`${formatVND(getVal(row, settings.MAP.TAMUNG))} / ${formatVND(getVal(row, settings.MAP.DANHAN))}`).setHorizontalAlignment("right"); rowIdx++;
                tempSheet.getRange(rowIdx, 1).setValue("SỐ TIỀN CHUYỂN KHOẢN").setFontWeight("bold").setFontSize(12);
                tempSheet.getRange(rowIdx, 2, 1, 2).merge().setValue(formatVND(getVal(row, settings.MAP.LUONGCK)) + " VND").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right").setBackground("#fff2cc");

                tempSheet.getRange("A7:C" + rowIdx).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

                SpreadsheetApp.flush();
                const pdfBlob = tempSpreadsheet.getAs('application/pdf').setName(`${settings.PDF_FILE_NAME_PREFIX}_${currentMonth}_${employeeName.replace(/\s+/g, '_')}.pdf`);

                const subject = `${settings.EMAIL_SUBJECT_PREFIX} ${payrollMonthYear} - ${employeeName}`;
                const body = `Kính gửi anh/chị ${employeeName},\n\n` +
                    `${settings.SENDER_NAME} xin gửi anh/chị phiếu lương ${payrollMonthYear} trong file PDF đính kèm.\n\n` +
                    `Trân trọng,\nPhòng Nhân sự - ${settings.SENDER_NAME}`;

                MailApp.sendEmail({ to: String(recipientEmail).trim(), subject: subject, body: body, attachments: [pdfBlob] });
                emailsSentCount++;
            } catch (e) {
                console.error(`Lỗi cho ${employeeName}: ${e.toString()}`);
                emailsFailedCount++;
            }
        }
    } finally {
        if (tempSpreadsheet) DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    }

    SpreadsheetApp.getUi().alert("Hoàn tất", `Đã gửi thành công: ${emailsSentCount}\nThất bại: ${emailsFailedCount}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Hàm lấy cấu hình từ tab "CONFIG" hoặc từ GLOBAL_CONFIG.
 */
function getSettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("CONFIG");
    let settings = { ...GLOBAL_CONFIG };

    if (configSheet) {
        const data = configSheet.getDataRange().getValues();
        // Giả định Tab CONFIG có dạng: Cột A là Key, Cột B là Value
        for (let i = 0; i < data.length; i++) {
            const key = String(data[i][0]).trim().toUpperCase();
            const value = data[i][1];
            if (key && value !== undefined && value !== "") {
                if (settings.hasOwnProperty(key)) {
                    settings[key] = value;
                }
            }
        }
    }
    return settings;
}

/**
 * Hàm tiện ích chuyển đổi tên cột dạng chữ (A, B, AA, CL) sang chỉ số 0-based.
 */
function columnLetterToIndex(columnLetter) {
    columnLetter = columnLetter.toUpperCase();
    let column = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        column *= 26;
        column += (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return column - 1;
}

/**
 * Hàm tiện ích chuyển đổi chỉ số cột 0-based sang tên cột dạng chữ.
 */
function columnNumberToLetter(column) {
    let temp, letter = '';
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}


/**
 * Hàm tạo menu tùy chỉnh trong Google Sheet để dễ dàng chạy script.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('SME Tools')
        .addItem('Gửi Phiếu Lương PDF (Tháng Hiện Tại)', 'sendMonthlyPayrollEmails_WithPDF')
        .addToUi();
}