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
    // 1. Vị trí dữ liệu
    START_ROW: 2,              // Hàng bắt đầu có dữ liệu (Hàng 1 là Header)
    COL_EMAIL: "F",            // Cột Email
    COL_NAME: "B",             // Cột Tên

    // 2. Tên các Sheet quan trọng
    SHEET_NAME_PREFIX: "T",    // Tiền tố tên sheet tháng (Vd: T5)
    TEMPLATE_SHEET_NAME: "TEMPLATE", // Tên sheet chứa mẫu phiếu lương

    // 3. Thông tin Công ty (Sẽ dùng làm tag {{SENDER_NAME}}, {{SENDER_ADDRESS}}...)
    SENDER_NAME: "SME SOLUTIONS JOINT STOCK COMPANY",
    SENDER_ADDRESS: "Tòa nhà Alpha, số 123 Đường Beta, Quận Gamma, Hà Nội",
    SENDER_HOTLINE: "1900 xxxx",
    CONTACT_EMAIL: "hr@sme-solutions.vn",

    // 4. Mapping cột (Sẽ dùng làm tag {{HOTEN}}, {{VITRI}}...)
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

    // 5. Cấu hình Email
    EMAIL_SUBJECT_PREFIX: "Phiếu Lương",
    PDF_FILE_NAME_PREFIX: "PhieuLuong_2025"
};

/**
 * Hàm chính để gửi Email bảng lương sử dụng Template Sheet.
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
    const templateSheet = ss.getSheetByName(settings.TEMPLATE_SHEET_NAME);

    if (!sourceSheet) {
        SpreadsheetApp.getUi().alert("Lỗi", `Không tìm thấy sheet dữ liệu "${sheetName}".`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }
    if (!templateSheet) {
        SpreadsheetApp.getUi().alert("Lỗi", `Không tìm thấy sheet mẫu "${settings.TEMPLATE_SHEET_NAME}". Bạn cần tạo một sheet tên là "${settings.TEMPLATE_SHEET_NAME}" và thiết kế mẫu phiếu lương tại đó.`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const startRow = settings.START_ROW;
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < startRow) {
        SpreadsheetApp.getUi().alert("Thông báo", "Không có dữ liệu nhân viên.");
        return;
    }

    const dataRange = sourceSheet.getDataRange();
    const values = dataRange.getValues();

    let emailsSentCount = 0;
    let emailsFailedCount = 0;

    const getVal = (rowValues, colLetter) => {
        const idx = columnLetterToIndex(colLetter);
        return rowValues[idx];
    };

    const formatVND = (val) => {
        if (val === null || val === undefined || val === "" || isNaN(parseFloat(val))) return "0";
        return new Intl.NumberFormat('vi-VN').format(parseFloat(val));
    };

    // Tạo Spreadsheet tạm để xuất PDF
    const tempSS = SpreadsheetApp.create(`Temp_Payroll_v2_${new Date().getTime()}`);
    const defaultSheet = tempSS.getSheets()[0];
    defaultSheet.setName("DUMMY");
    defaultSheet.hideSheet();

    try {
        for (let i = startRow - 1; i < values.length; i++) {
            const row = values[i];
            const recipientEmail = getVal(row, settings.COL_EMAIL);
            const employeeName = getVal(row, settings.COL_NAME);

            if (!recipientEmail || !String(recipientEmail).includes('@')) continue;

            try {
                // 1. Copy template sang Spreadsheet tạm
                const currentTempSheet = templateSheet.copyTo(tempSS);
                currentTempSheet.setName(employeeName.substring(0, 30)); // Giới hạn độ dài tên sheet

                // 2. Thay thế các thẻ TAG
                const textFinder = currentTempSheet.createTextFinder("{{.*}}").useRegularExpression(true);

                // Tạo một bản đồ thay thế (Key -> Value)
                const replacements = {
                    "{{THANGNAM}}": payrollMonthYear,
                    "{{SENDER_NAME}}": settings.SENDER_NAME,
                    "{{SENDER_ADDRESS}}": settings.SENDER_ADDRESS,
                    "{{SENDER_HOTLINE}}": settings.SENDER_HOTLINE,
                    "{{CONTACT_EMAIL}}": settings.CONTACT_EMAIL
                };

                // Thêm các tag từ MAP
                for (const key in settings.MAP) {
                    let val = getVal(row, settings.MAP[key]);
                    // Nếu là các cột liên quan đến tiền bạc (bắt đầu bằng L_ hoặc PC_ hoặc các cột đặc biệt)
                    if (key.startsWith("L_") || key.startsWith("PC_") ||
                        ["THUONG", "TRU_CHIU_THUE", "TRU_KG_CHIU_THUE", "TONGLUONG", "BHXH", "THUETNCN", "THUCLINH", "PHAT", "TAMUNG", "DANHAN", "LUONGCK"].includes(key)) {
                        val = formatVND(val);
                    }
                    replacements[`{{${key}}}`] = val;
                }

                // Thực hiện thay thế thực tế
                for (const tag in replacements) {
                    currentTempSheet.createTextFinder(tag).replaceAllWith(String(replacements[tag]));
                }

                SpreadsheetApp.flush();

                // 3. Xuất PDF từ sheet hiện tại
                // Lưu ý: getAs('application/pdf') sẽ xuất TOÀN BỘ spreadsheet tạm. 
                // Do đó mỗi lần ta chỉ nên có MỘT sheet trong tempSS hoặc ẩn các sheet khác.
                const allSheets = tempSS.getSheets();
                allSheets.forEach(s => { if (s.getName() !== currentTempSheet.getName()) s.hideSheet(); });

                const pdfBlob = tempSS.getAs('application/pdf').setName(`${settings.PDF_FILE_NAME_PREFIX}_${currentMonth}_${employeeName.replace(/\s+/g, '_')}.pdf`);

                // Hiển thị lại các sheet để xóa/quản lý nếu cần (thực ra không cần vì ta ẩn để xuất PDF)
                allSheets.forEach(s => { if (s !== currentTempSheet && s.getName() !== "Sheet1") s.showSheet(); });

                // 4. Gửi Email
                const subject = `${settings.EMAIL_SUBJECT_PREFIX} ${payrollMonthYear} - ${employeeName}`;
                const body = `Kính gửi anh/chị ${employeeName},\n\n` +
                    `${settings.SENDER_NAME} xin gửi anh/chị phiếu lương ${payrollMonthYear} trong file PDF đính kèm.\n\n` +
                    `Trân trọng,\nPhòng Nhân sự - ${settings.SENDER_NAME}`;

                MailApp.sendEmail({
                    to: String(recipientEmail).trim(),
                    subject: subject,
                    body: body,
                    attachments: [pdfBlob]
                });

                // 5. Dọn dẹp sheet vừa dùng
                tempSS.deleteSheet(currentTempSheet);
                emailsSentCount++;

            } catch (e) {
                console.error(`Lỗi cho ${employeeName}: ${e.toString()}`);
                emailsFailedCount++;
            }
        }
    } finally {
        if (tempSS) DriveApp.getFileById(tempSS.getId()).setTrashed(true);
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
        .addSeparator()
        .addItem('Tạo Sheet Mẫu (TEMPLATE)', 'createSampleTemplate')
        .addToUi();
}

/**
 * Hàm tự động tạo Sheet TEMPLATE với thiết kế chuyên nghiệp và các thẻ Tag chuẩn.
 */
function createSampleTemplate() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = GLOBAL_CONFIG.TEMPLATE_SHEET_NAME;
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert('Cảnh báo', `Sheet "${sheetName}" đã tồn tại. Bạn có muốn xóa đi và tạo lại mẫu mới không?`, ui.ButtonSet.YES_NO);
        if (response !== ui.Button.YES) return;
        ss.deleteSheet(sheet);
    }

    sheet = ss.insertSheet(sheetName);

    // --- Thiết kế Giao diện ---
    // 1. Cấu hình cột
    sheet.setColumnWidth(1, 220); // Cột nhãn
    sheet.setColumnWidth(2, 150); // Cột giá trị 1
    sheet.setColumnWidth(3, 150); // Cột giá trị 2

    // 2. Tiêu đề Công ty
    sheet.getRange("A1:C1").merge().setValue("{{SENDER_NAME}}").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    sheet.getRange("A2:C2").merge().setValue("{{SENDER_ADDRESS}}").setFontSize(9).setHorizontalAlignment("center");
    sheet.getRange("A3:C3").merge().setValue("Hotline: {{SENDER_HOTLINE}} | Email: {{CONTACT_EMAIL}}").setFontSize(9).setHorizontalAlignment("center");

    // 3. Tiêu đề Phiếu Lương
    sheet.getRange("A5:C5").merge().setValue("PHIẾU LƯƠNG {{THANGNAM}}").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center").setBackground("#f0f0f0");

    // 4. Khối thông tin nhân viên
    sheet.getRange("A7").setValue("THÔNG TIN NHÂN VIÊN").setFontWeight("bold").setBackground("#e0e0e0");
    sheet.getRange("A8").setValue("Họ và tên:"); sheet.getRange("B8:C8").merge().setValue("{{HOTEN}}").setFontWeight("bold");
    sheet.getRange("A9").setValue("Vị trí:"); sheet.getRange("B9:C9").merge().setValue("{{VITRI}}");
    sheet.getRange("A10").setValue("Tài khoản nhận:"); sheet.getRange("B10:C10").merge().setValue("{{STK}} ({{NGANHANG}})");

    // 5. Khối Chi tiết công xá
    sheet.getRange("A12").setValue("CHI TIẾT CÔNG XÁ").setFontWeight("bold").setBackground("#e0e0e0");
    sheet.getRange("A13").setValue("Công chuẩn / Thực tế:"); sheet.getRange("B13").setValue("{{NGAYCONGCHUAN}}"); sheet.getRange("C13").setValue("{{TONGGIOTT}} giờ");
    sheet.getRange("A14").setValue("Tăng ca (150%/200%/300%):"); sheet.getRange("B14:C14").merge().setValue("{{GIO150}} / {{GIO200}} / {{GIO300}}");
    sheet.getRange("A15").setValue("Nghỉ (Ro/N/P/L):"); sheet.getRange("B15:C15").merge().setValue("{{RO}} / {{N}} / {{P}} / {{L}}");

    // 6. Khối Thu nhập
    sheet.getRange("A17").setValue("CHI TIẾT THU NHẬP").setFontWeight("bold").setBackground("#e0e0e0");
    const incomeTags = [
        ["Lương cơ bản (HĐ)", "{{L_COBAN}}"],
        ["Phụ cấp (Ăn/ĐT/NL)", "{{PC_ANTRUA}} / {{PC_DIENTHOAI}} / {{PC_DILAI}}"],
        ["Lương ngày công thực tế", "{{L_NGAYCONG}}"],
        ["Lương OT", "{{L_OT}}"],
        ["Lương KPI / Hiệu quả", "{{L_KPI}}"],
        ["Thưởng / Phúc lợi khác", "{{THUONG}}"],
        ["TỔNG THU NHẬP", "{{TONGLUONG}}"]
    ];
    let row = 18;
    incomeTags.forEach(pair => {
        sheet.getRange(row, 1).setValue(pair[0]);
        sheet.getRange(row, 2, 1, 2).merge().setValue(pair[1]).setHorizontalAlignment("right");
        if (pair[0].includes("TỔNG")) sheet.getRange(row, 1, 1, 3).setFontWeight("bold").setBackground("#fff2cc");
        row++;
    });

    // 7. Khối Khấu trừ
    row++;
    sheet.getRange(row, 1).setValue("GIẢM TRỪ & KHẤU TRỪ").setFontWeight("bold").setBackground("#e0e0e0"); row++;
    sheet.getRange(row, 1).setValue("BHXH / Thuế TNCN:"); sheet.getRange(row, 2, 1, 2).merge().setValue("{{BHXH}} / {{THUETNCN}}").setHorizontalAlignment("right"); row++;
    sheet.getRange(row, 1).setValue("THỰC LĨNH").setFontWeight("bold"); sheet.getRange(row, 2, 1, 2).merge().setValue("{{THUCLINH}}").setFontWeight("bold").setHorizontalAlignment("right"); row++;

    // 8. Khối Thanh toán
    row++;
    sheet.getRange(row, 1).setValue("THANH TOÁN CUỐI CÙNG").setFontWeight("bold").setBackground("#d9ead3"); row++;
    sheet.getRange(row, 1).setValue("Tạm ứng / Đã quyết toán:"); sheet.getRange(row, 2, 1, 2).merge().setValue("{{TAMUNG}} / {{DANHAN}}").setHorizontalAlignment("right"); row++;
    sheet.getRange(row, 1).setValue("SỐ TIỀN CHUYỂN KHOẢN").setFontWeight("bold").setFontSize(12);
    sheet.getRange(row, 2, 1, 2).merge().setValue("{{LUONGCK}} VND").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right").setBackground("#fff2cc");

    // Border toàn bộ vùng nội dung
    sheet.getRange("A7:C" + row).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    // Format hiển thị cho người dùng
    sheet.getRange("A1:C" + row).setFontFamily("Roboto");

    SpreadsheetApp.getUi().alert("Thành công", `Đã tạo xong sheet "${sheetName}". Bạn có thể tùy chỉnh thêm font chữ hoặc logo nếu muốn.`, SpreadsheetApp.getUi().ButtonSet.OK);
}