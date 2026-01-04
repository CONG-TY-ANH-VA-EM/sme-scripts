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
        THUCLINH: "AJ", PHAT: "AK", TAMUNG: "AL", DANHAN: "AM", LUONGCK: "AN",
        SENT_STATUS: "AO"      // Cột theo dõi trạng thái gửi
    },

    // 5. Cấu hình Resilience (Tự động resume & Quota)
    MAX_RUNTIME_MS: 5 * 60 * 1000, // 5 phút (Giới hạn thực thi Google là 6p)
    MIN_QUOTA: 10,                 // Dừng gửi nếu quota Gmail < 10

    // 6. Cấu hình Email
    EMAIL_SUBJECT_PREFIX: "Phiếu Lương",
    PDF_FILE_NAME_PREFIX: "PhieuLuong"
};

/**
 * Hàm chính để gửi Email bảng lương sử dụng Template Sheet.
 * Hỗ trợ tự động Resume nếu gặp lỗi Timeout hoặc hết thời gian.
 */
function sendMonthlyPayrollEmails_WithPDF() {
    const startTime = Date.now();
    const settings = getSettings();
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;
    const currentYear = currentDate.getFullYear();
    const payrollMonthYear = `Tháng ${currentMonth}/${currentYear}`;
    const sheetName = settings.SHEET_NAME_PREFIX + currentMonth;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sheetName);
    const templateSheet = ss.getSheetByName(settings.TEMPLATE_SHEET_NAME);

    if (!sourceSheet || !templateSheet) {
        SpreadsheetApp.getUi().alert("Lỗi", `Cần có sheet "${sheetName}" và "${settings.TEMPLATE_SHEET_NAME}".`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const startRow = settings.START_ROW;
    const lastRow = sourceSheet.getLastRow();
    const values = sourceSheet.getDataRange().getValues();
    const statusColIdx = columnLetterToIndex(settings.MAP.SENT_STATUS);

    const stats = {
        total: values.length - (startRow - 1),
        success: 0,
        failed: [],
        skipped: 0,
        isResuming: false,
        monthYear: payrollMonthYear
    };

    // Tạo Spreadsheet tạm
    const tempSS = SpreadsheetApp.create(`Temp_Payroll_Resilient_${new Date().getTime()}`);
    const defaultSheet = tempSS.getSheets()[0];
    defaultSheet.setName("DUMMY");
    defaultSheet.hideSheet();

    try {
        for (let i = startRow - 1; i < values.length; i++) {
            const row = values[i];
            const currentRowNum = i + 1;
            const currentStatus = row[statusColIdx];

            // 1. Kiểm tra trạng thái đã gửi chưa
            if (currentStatus === "Thành công") {
                stats.skipped++;
                continue;
            }

            // 2. Kiểm tra Quota Gmail
            if (MailApp.getRemainingDailyQuota() < settings.MIN_QUOTA) {
                console.warn("Hết hạn mức Quota Gmail.");
                sendSummaryEmail(stats, "TẠM DỪNG (Hết Quota Gmail)");
                deleteResumeTrigger();
                return;
            }

            // 3. Kiểm tra Thời gian thực thi (Resilience)
            if (Date.now() - startTime > settings.MAX_RUNTIME_MS) {
                console.log("Sắp hết thời gian thực thi. Đang khởi tạo cơ chế Resume tự động...");
                stats.isResuming = true;
                createResumeTrigger();
                sendSummaryEmail(stats, "TẠM DỪNG (Đang chạy tiếp...)");
                SpreadsheetApp.getUi().alert("Đang tạm dừng", "Script đã chạy gần 5 phút. Hệ thống sẽ tự động chạy tiếp phần còn lại sau 1 phút nữa và đã gửi báo cáo tạm thời cho bạn.", SpreadsheetApp.getUi().ButtonSet.OK);
                return;
            }

            const recipientEmail = row[columnLetterToIndex(settings.MAP.EMAIL)];
            const employeeName = row[columnLetterToIndex(settings.MAP.HOTEN)];

            if (!recipientEmail || !String(recipientEmail).includes('@')) continue;

            try {
                // Render và gửi (Sử dụng cơ chế đã tối ưu v1.1)
                const currentTempSheet = templateSheet.copyTo(tempSS);
                currentTempSheet.setName(employeeName.substring(0, 30) + "_" + currentRowNum);

                const replacements = {
                    "{{THANGNAM}}": payrollMonthYear,
                    "{{SENDER_NAME}}": settings.SENDER_NAME,
                    "{{SENDER_ADDRESS}}": settings.SENDER_ADDRESS,
                    "{{SENDER_HOTLINE}}": settings.SENDER_HOTLINE,
                    "{{CONTACT_EMAIL}}": settings.CONTACT_EMAIL
                };

                for (const key in settings.MAP) {
                    let val = row[columnLetterToIndex(settings.MAP[key])];
                    if (key.startsWith("L_") || key.startsWith("PC_") ||
                        ["THUONG", "TRU_CHIU_THUE", "TRU_KG_CHIU_THUE", "TONGLUONG", "BHXH", "THUETNCN", "THUCLINH", "PHAT", "TAMUNG", "DANHAN", "LUONGCK"].includes(key)) {
                        val = (val === null || val === undefined || val === "" || isNaN(parseFloat(val))) ? "0" : new Intl.NumberFormat('vi-VN').format(parseFloat(val));
                    }
                    replacements[`{{${key}}}`] = val;
                }

                for (const tag in replacements) {
                    currentTempSheet.createTextFinder(tag).replaceAllWith(String(replacements[tag]));
                }

                SpreadsheetApp.flush();
                const allSheets = tempSS.getSheets();
                allSheets.forEach(s => { if (s.getName() !== currentTempSheet.getName()) s.hideSheet(); });
                const pdfBlob = tempSS.getAs('application/pdf').setName(`${settings.PDF_FILE_NAME_PREFIX}_${employeeName.replace(/\s+/g, '_')}.pdf`);
                allSheets.forEach(s => { if (s !== currentTempSheet && s.getName() !== "DUMMY") s.showSheet(); });

                MailApp.sendEmail({
                    to: String(recipientEmail).trim(),
                    subject: `${settings.EMAIL_SUBJECT_PREFIX} ${payrollMonthYear} - ${employeeName}`,
                    body: `Kính gửi anh/chị ${employeeName},\n\nMẫu phiếu lương chi tiết đính kèm.\n\nTrân trọng.`,
                    attachments: [pdfBlob]
                });

                // CẬP NHẬT TRẠNG THÁI NGAY LẬP TỨC
                sourceSheet.getRange(currentRowNum, statusColIdx + 1).setValue("Thành công").setBackground("#d9ead3");
                tempSS.deleteSheet(currentTempSheet);
                stats.success++;
            } catch (err) {
                console.error(`Lỗi tại hàng ${currentRowNum}: ${err.message}`);
                sourceSheet.getRange(currentRowNum, statusColIdx + 1).setValue("Lỗi: " + err.message).setBackground("#f4cccc");
                stats.failed.push({ name: employeeName, row: currentRowNum, error: err.message });
            }
        }

        // Hoàn tất toàn bộ
        deleteResumeTrigger();
        sendSummaryEmail(stats, "HOÀN TẤT");
        SpreadsheetApp.getUi().alert("Hoàn tất", `Đã gửi thành công ${stats.success} email. Vui lòng kiểm tra email của bạn để xem báo cáo chi tiết.`, SpreadsheetApp.getUi().ButtonSet.OK);

    } finally {
        if (tempSS) DriveApp.getFileById(tempSS.getId()).setTrashed(true);
    }
}

/**
 * Tạo Trigger để chạy lại script sau 1 phút.
 */
function createResumeTrigger() {
    deleteResumeTrigger(); // Xóa cái cũ nếu có
    ScriptApp.newTrigger("sendMonthlyPayrollEmails_WithPDF")
        .timeBased()
        .after(1 * 60 * 1000)
        .create();
}

/**
 * Xóa tất cả Trigger liên quan đến hàm gửi lương để tránh chạy lặp.
 */
function deleteResumeTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === "sendMonthlyPayrollEmails_WithPDF") {
            ScriptApp.deleteTrigger(t);
        }
    });
}

/**
 * Xóa trạng thái gửi để có thể gửi lại từ đầu.
 */
function resetSentStatus() {
    const settings = getSettings();
    const currentMonth = new Date().getMonth() + 1;
    const sheetName = settings.SHEET_NAME_PREFIX + currentMonth;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        const lastRow = sheet.getLastRow();
        const startRow = settings.START_ROW;
        if (lastRow >= startRow) {
            const statusColIdx = columnLetterToIndex(settings.MAP.SENT_STATUS);
            sheet.getRange(startRow, statusColIdx + 1, lastRow - startRow + 1, 1).clearContent().setBackground(null);
            SpreadsheetApp.getUi().alert("Đã Reset", "Đã xóa trạng thái gửi của tháng hiện tại.", SpreadsheetApp.getUi().ButtonSet.OK);
        }
    }
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
        .addItem('Xóa trạng thái gửi (Reset)', 'resetSentStatus')
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

/**
 * Gửi email báo cáo tóm tắt cho người vận hành.
 */
function sendSummaryEmail(stats, statusMessage) {
    const userEmail = Session.getActiveUser().getEmail();
    const subject = `[BÁO CÁO] Kết quả gửi lương ${stats.monthYear} - ${statusMessage}`;

    let body = `Chào bạn,\n\nHệ thống SME Scripts gửi báo cáo kết quả thực hiện gửi phiếu lương:\n\n`;
    body += `--------------------------------------------------\n`;
    body += `Trạng thái: ${statusMessage}\n`;
    body += `Tổng số nhân viên trong danh sách: ${stats.total}\n`;
    body += `Số ca đã hoàn tất thành công: ${stats.success + stats.skipped}\n`;
    body += `  - Mới thành công: ${stats.success}\n`;
    body += `  - Đã xong từ trước (bỏ qua): ${stats.skipped}\n`;
    body += `Số ca thất bại: ${stats.failed.length}\n`;
    body += `--------------------------------------------------\n\n`;

    if (stats.failed.length > 0) {
        body += `DANH SÁCH CÁC TRƯỜNG HỢP LỖI:\n`;
        stats.failed.forEach(item => {
            body += `- Hàng ${item.row} | ${item.name}: ${item.error}\n`;
        });
        body += `\nVui lòng kiểm tra lại dữ liệu tại các hàng trên và bấm "Gửi lại" sau khi sửa lỗi.\n\n`;
    }

    if (stats.isResuming) {
        body += `\nLƯU Ý: Script đã chạm giới hạn thời gian và đang chạy tiếp phần còn lại. Bạn sẽ nhận được báo cáo cuối cùng sau khi hoàn tất.\n`;
    }

    body += `\nTrân trọng,\nSME Solutions AI Assistant.`;

    try {
        MailApp.sendEmail(userEmail, subject, body);
    } catch (e) {
        console.error("Không thể gửi email báo cáo: " + e.message);
    }
}