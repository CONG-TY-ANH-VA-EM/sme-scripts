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
        .addToUi();
}