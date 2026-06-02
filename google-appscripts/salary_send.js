/**
 * =============================================================================
 *  CÔNG CỤ GỬI PHIẾU LƯƠNG TỰ ĐỘNG (SME Payroll Tool) — v2.0
 * =============================================================================
 *  Mô tả:
 *    - Đọc dữ liệu lương từ tab tháng (vd: T5 = lương tháng 5), KHÔNG dùng tab này thì sửa SHEET_NAME_PREFIX.
 *    - Dòng 1 là tiêu đề, dữ liệu nhân viên bắt đầu từ START_ROW (mặc định dòng 2).
 *    - Với mỗi nhân viên: render mẫu (tab TEMPLATE) -> xuất PDF -> (tùy chọn) lưu Drive -> gửi email.
 *    - Tự động RESUME khi chạm giới hạn 6 phút của Google, và dừng an toàn khi sắp hết quota Gmail.
 *
 *  Vị trí cột mặc định (đổi được qua tab CONFIG, dòng có key dạng "MAP.XXX"):
 *    - Tên nhân viên : cột B  (MAP.HOTEN)
 *    - Email         : cột F  (MAP.EMAIL)
 *    - Tổng lương     : cột AE (MAP.TONGLUONG)
 *    - Thực lĩnh      : cột AJ (MAP.THUCLINH)
 *    - Lương chuyển khoản: cột AN (MAP.LUONGCK)
 *    - Trạng thái gửi : cột AO (MAP.SENT_STATUS)
 *
 *  Phát triển bởi: AI Assistant - SME SOLUTIONS
 *
 *  GHI CHÚ BẢO MẬT (PDF password):
 *    Google Apps Script KHÔNG hỗ trợ đặt mật khẩu / mã hoá PDF native.
 *    Thay vào đó dùng SECURE_SHARE = true: file PDF lưu trên Drive chỉ được chia sẻ
 *    riêng cho đúng email nhân viên, và email sẽ kèm link truy cập riêng tư.
 * =============================================================================
 */

const GLOBAL_CONFIG = {
    // 1. Vị trí dữ liệu
    START_ROW: 2,                    // Hàng bắt đầu có dữ liệu (Hàng 1 là Header)

    // 2. Tên các Sheet quan trọng
    SHEET_NAME_PREFIX: "T",          // Tiền tố tên tab tháng (Vd: T5 cho lương tháng 5)
    TEMPLATE_SHEET_NAME: "TEMPLATE", // Tên tab chứa mẫu phiếu lương
    LOG_SHEET_NAME: "LOG",           // Tên tab ghi nhật ký gửi

    // 3. Thông tin Công ty (dùng làm tag {{SENDER_NAME}}, {{SENDER_ADDRESS}}...)
    SENDER_NAME: "SME SOLUTIONS JOINT STOCK COMPANY",
    SENDER_ADDRESS: "Tòa nhà Alpha, số 123 Đường Beta, Quận Gamma, Hà Nội",
    SENDER_HOTLINE: "1900 xxxx",
    CONTACT_EMAIL: "hr@sme-solutions.vn",
    LOGO_URL: "https://raw.githubusercontent.com/CONG-TY-ANH-VA-EM/sme-scripts/main/assets/logo-ane-white.png", // logo header phiếu (URL ảnh tải được; đổi thành logo công ty bạn)

    // 4. Mapping cột (dùng làm tag {{HOTEN}}, {{VITRI}}...). Đổi qua CONFIG bằng key "MAP.HOTEN", "MAP.EMAIL"...
    MAP: {
        SOTT: "A", HOTEN: "B", VITRI: "C", NGAYGUI: "D", STK: "E", EMAIL: "F", NGANHANG: "G",
        NGAYCONGCHUAN: "H", TONGGIOTT: "I", GIO150: "J", GIO200: "K", GIO300: "L",
        RO: "M", N: "N", P: "O", L: "P",
        L_COBAN: "Q", PC_ANTRUA: "R", PC_DIENTHOAI: "S", PC_DILAI: "T",
        L_NGAYCONG: "U", L_NGHIPHEP: "V", L_OT: "W", L_KPI: "X", L_PHEPNAM: "Y",
        PC_MAYTINH: "Z", CONGTACPHI: "AA", THUONG: "AB", TRU_CHIU_THUE: "AC", TRU_KG_CHIU_THUE: "AD",
        TONGLUONG: "AE", PHEPCONLAI: "AF", GIAMTRU: "AG", BHXH: "AH", THUETNCN: "AI",
        THUCLINH: "AJ", PHAT: "AK", TAMUNG: "AL", DANHAN: "AM", LUONGCK: "AN",
        SENT_STATUS: "AO"            // Cột theo dõi trạng thái gửi
    },

    // 5. Resilience (tự động resume & quota)
    MAX_RUNTIME_MS: 5 * 60 * 1000,   // 5 phút (giới hạn thực thi Google là 6p)
    MIN_QUOTA: 10,                   // Dừng gửi nếu quota Gmail < số này

    // 6. Email
    EMAIL_SUBJECT_PREFIX: "Phiếu Lương",
    PDF_FILE_NAME_PREFIX: "PhieuLuong",
    ATTACH_PDF_TO_EMAIL: true,       // true = đính kèm PDF vào email

    // 7. Lưu trữ Drive (mới v2.0)
    SAVE_TO_DRIVE: true,             // true = lưu mỗi phiếu vào Drive theo Năm/Tháng
    DRIVE_ROOT_FOLDER_NAME: "Phiếu lương SME", // Tên thư mục gốc (tạo tự động ở My Drive nếu chưa có)
    DRIVE_ROOT_FOLDER_ID: "",        // Nếu điền ID thư mục có sẵn sẽ ưu tiên dùng (bỏ qua tên ở trên)

    // 8. Bảo mật (thay cho "password PDF" — xem ghi chú đầu file)
    SECURE_SHARE: false,             // true = file Drive chỉ chia sẻ cho đúng email nhân viên + gửi link

    // 9. Nhật ký
    ENABLE_LOG: true                 // true = ghi lịch sử gửi vào tab LOG
};

// =============================================================================
//  ENTRY POINTS (các hàm chạy từ menu / trigger)
// =============================================================================

/**
 * MENU + giữ tên cũ để tương thích nút "Run" trong Apps Script.
 * Gửi phiếu lương cho THÁNG HIỆN TẠI (bắt đầu mới, không phải resume).
 */
function sendMonthlyPayrollEmails_WithPDF() {
    const now = new Date();
    setTargetMonth(now.getMonth() + 1, now.getFullYear());
    deleteResumeTrigger(); // bắt đầu mới: dọn trigger resume cũ
    processPayroll();
}

/**
 * MENU: chọn tháng thủ công (vd gửi lương T5 trong khi đang là T6).
 */
function sendPayrollForSelectedMonth() {
    const ui = SpreadsheetApp.getUi();
    const now = new Date();
    const resp = ui.prompt(
        "Chọn kỳ lương",
        `Nhập tháng cần gửi (vd: 5  hoặc  5/2025).\nĐể trống = tháng hiện tại (${now.getMonth() + 1}/${now.getFullYear()}).`,
        ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    let raw = resp.getResponseText().trim();
    let month, year;
    if (!raw) {
        month = now.getMonth() + 1;
        year = now.getFullYear();
    } else {
        const parts = raw.split(/[\/\-\s]+/);
        month = parseInt(parts[0], 10);
        year = parts[1] ? parseInt(parts[1], 10) : now.getFullYear();
    }
    if (!month || month < 1 || month > 12) {
        ui.alert("Lỗi", "Tháng không hợp lệ. Vui lòng nhập 1-12.", ui.ButtonSet.OK);
        return;
    }
    setTargetMonth(month, year);
    deleteResumeTrigger();
    processPayroll();
}

/**
 * TRIGGER (chỉ chạy ngầm): chạy tiếp phần còn lại. KHÔNG đặt lại tháng đích.
 */
function resumePayroll() {
    // Resume PHẢI có tháng đích rõ ràng. Nếu thiếu (vd trigger mồ côi sau khi đã hoàn tất),
    // TUYỆT ĐỐI không mặc định về tháng hiện tại — sẽ gửi nhầm tab / nhầm người.
    const p = PropertiesService.getDocumentProperties();
    if (!p.getProperty('PAYROLL_MONTH')) {
        deleteResumeTrigger();
        return;
    }
    processPayroll();
}

/**
 * MENU: gửi thử 1 phiếu về email người vận hành (hoặc email nhập tay) để xem layout.
 * KHÔNG đánh dấu trạng thái, KHÔNG lưu Drive thật, KHÔNG ghi log.
 */
function sendTestPayslip() {
    const ui = SpreadsheetApp.getUi();
    const settings = getSettings();
    const operator = Session.getActiveUser().getEmail();

    const emailResp = ui.prompt(
        "Gửi thử phiếu lương",
        `Email nhận bản thử (để trống = email của bạn: ${operator}):`,
        ui.ButtonSet.OK_CANCEL
    );
    if (emailResp.getSelectedButton() !== ui.Button.OK) return;
    const testEmail = emailResp.getResponseText().trim() || operator;

    const startRow = Number(settings.START_ROW) || 2;
    const rowResp = ui.prompt(
        "Gửi thử phiếu lương",
        `Lấy dữ liệu từ HÀNG SỐ mấy? (mặc định ${startRow}):`,
        ui.ButtonSet.OK_CANCEL
    );
    if (rowResp.getSelectedButton() !== ui.Button.OK) return;
    const rowNum = parseInt(rowResp.getResponseText().trim(), 10) || startRow;

    const { month, year } = getTargetMonth();
    const payrollMonthYear = `Tháng ${month}/${year}`;
    const sheetName = settings.SHEET_NAME_PREFIX + month;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sheetName);
    const templateSheet = ss.getSheetByName(settings.TEMPLATE_SHEET_NAME);
    if (!sourceSheet || !templateSheet) {
        ui.alert("Lỗi", `Cần có tab "${sheetName}" và "${settings.TEMPLATE_SHEET_NAME}".`, ui.ButtonSet.OK);
        return;
    }

    const values = sourceSheet.getDataRange().getValues();
    if (rowNum - 1 >= values.length || rowNum < startRow) {
        ui.alert("Lỗi", `Hàng ${rowNum} không có dữ liệu.`, ui.ButtonSet.OK);
        return;
    }
    const row = values[rowNum - 1];
    const employeeName = String(row[columnLetterToIndex(settings.MAP.HOTEN)] || ("NV_" + rowNum));

    const tempSS = SpreadsheetApp.create(`Temp_Payroll_Test_${new Date().getTime()}`);
    tempSS.getSheets()[0].setName("DUMMY"); // không hide sheet duy nhất (sẽ lỗi "can't hide all sheets")
    try {
        const pdfBlob = renderEmployeePdf(templateSheet, tempSS, row, settings, payrollMonthYear, rowNum, employeeName);
        MailApp.sendEmail({
            to: testEmail,
            subject: `[THỬ] ${settings.EMAIL_SUBJECT_PREFIX} ${payrollMonthYear} - ${employeeName}`,
            body: `Đây là BẢN GỬI THỬ phiếu lương của "${employeeName}" (hàng ${rowNum}).\n` +
                  `Không gửi cho nhân viên, không ghi nhận trạng thái.\n\nKiểm tra layout PDF đính kèm rồi gửi thật khi đã ổn.`,
            attachments: [pdfBlob]
        });
        ui.alert("Đã gửi thử", `Đã gửi bản thử phiếu của "${employeeName}" tới ${testEmail}.`, ui.ButtonSet.OK);
    } catch (err) {
        ui.alert("Lỗi gửi thử", err.message, ui.ButtonSet.OK);
    } finally {
        DriveApp.getFileById(tempSS.getId()).setTrashed(true);
    }
}

// =============================================================================
//  CORE
// =============================================================================

/**
 * Vòng xử lý chính. Dùng cho cả gửi mới lẫn resume (gọi từ trigger).
 * Khóa tài liệu (LockService) để tránh chạy chồng (manual + trigger) gây gửi trùng phiếu.
 */
function processPayroll() {
    const settings = getSettings();
    const { month, year } = getTargetMonth();
    const payrollMonthYear = `Tháng ${month}/${year}`;
    const sheetName = settings.SHEET_NAME_PREFIX + month;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sheetName);
    const templateSheet = ss.getSheetByName(settings.TEMPLATE_SHEET_NAME);

    if (!sourceSheet || !templateSheet) {
        safeAlert("Lỗi", `Cần có tab "${sheetName}" và "${settings.TEMPLATE_SHEET_NAME}".`);
        return;
    }

    // Chống chạy chồng: nếu một tiến trình gửi lương khác đang chạy, LÊN LỊCH THỬ LẠI
    // (không bỏ luôn — nếu không, phần nhân viên còn lại có thể không bao giờ được gửi).
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(2000)) {
        console.warn("Một tiến trình gửi lương khác đang chạy — lên lịch chạy lại sau ~1 phút.");
        createResumeTrigger();
        safeAlert("Đang bận", "Có một tiến trình gửi lương khác đang chạy. Hệ thống sẽ tự thử lại sau khoảng 1 phút.");
        return;
    }

    const startTime = Date.now();
    const startRow = Number(settings.START_ROW) || 2;
    const minQuota = Number(settings.MIN_QUOTA) || 10;
    const values = sourceSheet.getDataRange().getValues();
    const statusColIdx = columnLetterToIndex(settings.MAP.SENT_STATUS);
    const emailIdx = columnLetterToIndex(settings.MAP.EMAIL);

    // Đếm SỐ NHÂN VIÊN CÓ EMAIL hợp lệ (không tính hàng trống)
    let totalEligible = 0;
    for (let i = startRow - 1; i < values.length; i++) {
        const e = values[i][emailIdx];
        if (e && String(e).includes('@')) totalEligible++;
    }

    const stats = {
        total: totalEligible,
        success: 0,
        failed: [],
        skipped: 0,
        isResuming: false,
        monthYear: payrollMonthYear
    };

    // Thư mục Drive cho kỳ lương (tạo lười khi cần)
    let monthFolder = null;
    if (settings.SAVE_TO_DRIVE) {
        try {
            monthFolder = getMonthFolder(settings, month, year);
        } catch (e) {
            console.error("Không tạo được thư mục Drive: " + e.message);
        }
    }

    // Tạo spreadsheet tạm TRONG try để chắc chắn lock luôn được nhả và file tạm luôn được dọn nếu lỗi.
    let tempSS = null;

    try {
        tempSS = SpreadsheetApp.create(`Temp_Payroll_Resilient_${new Date().getTime()}`);
        // KHÔNG hide sheet "DUMMY": hide sheet hiển-thị DUY NHẤT sẽ ném lỗi "You can't hide all the sheets".
        tempSS.getSheets()[0].setName("DUMMY");
        for (let i = startRow - 1; i < values.length; i++) {
            const row = values[i];
            const currentRowNum = i + 1;
            const currentStatus = row[statusColIdx];

            // 1. Bỏ qua hàng đã gửi thành công
            if (currentStatus === "Thành công") {
                stats.skipped++;
                continue;
            }

            // 2. Kiểm tra Quota Gmail
            if (MailApp.getRemainingDailyQuota() < minQuota) {
                console.warn("Hết hạn mức Quota Gmail.");
                sendSummaryEmail(stats, "TẠM DỪNG (Hết Quota Gmail)");
                deleteResumeTrigger();
                safeAlert("Tạm dừng", "Đã hết hạn mức gửi email trong ngày của Gmail. Hãy chạy lại vào ngày mai.");
                return;
            }

            // 3. Kiểm tra thời gian thực thi -> tạo trigger chạy tiếp
            if (Date.now() - startTime > settings.MAX_RUNTIME_MS) {
                console.log("Sắp hết thời gian thực thi. Khởi tạo cơ chế Resume tự động...");
                stats.isResuming = true;
                createResumeTrigger();
                sendSummaryEmail(stats, "TẠM DỪNG (Đang chạy tiếp...)");
                // safeAlert: nếu đang trong trigger sẽ tự bỏ qua, không crash
                safeAlert("Đang tạm dừng", "Script đã chạy gần 5 phút. Hệ thống sẽ tự động chạy tiếp phần còn lại sau ~1 phút và đã gửi báo cáo tạm thời cho bạn.");
                return;
            }

            const recipientEmail = row[emailIdx];
            const employeeName = String(row[columnLetterToIndex(settings.MAP.HOTEN)] || ("NV_" + currentRowNum));

            if (!recipientEmail || !String(recipientEmail).includes('@')) continue;

            try {
                const pdfBlob = renderEmployeePdf(templateSheet, tempSS, row, settings, payrollMonthYear, currentRowNum, employeeName);

                // Lưu Drive (tùy chọn)
                let driveUrl = "";
                let secureLinkReady = false; // chỉ true khi đã chia sẻ riêng cho NV thành công
                if (monthFolder) {
                    const file = monthFolder.createFile(pdfBlob);
                    driveUrl = file.getUrl();
                    if (settings.SECURE_SHARE) {
                        try {
                            file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
                            file.addViewer(String(recipientEmail).trim());
                            secureLinkReady = true;
                        } catch (shareErr) {
                            console.warn(`Không chia sẻ riêng được cho ${employeeName}: ${shareErr.message}. Sẽ đính kèm PDF thay thế.`);
                        }
                    }
                }

                // Soạn email
                const attachments = settings.ATTACH_PDF_TO_EMAIL ? [pdfBlob] : [];
                // An toàn: nếu không đính kèm VÀ không có link an toàn -> ép đính kèm để NV luôn nhận được phiếu
                if (attachments.length === 0 && !secureLinkReady) {
                    attachments.push(pdfBlob);
                }

                let body = `Kính gửi anh/chị ${employeeName},\n\n`;
                body += `Phiếu lương ${payrollMonthYear} của anh/chị đã sẵn sàng.\n`;
                if (attachments.length > 0) body += `Vui lòng xem file PDF đính kèm.\n`;
                if (secureLinkReady) {
                    body += `\nXem/tải phiếu lương (chỉ riêng anh/chị truy cập được):\n${driveUrl}\n`;
                }
                body += `\nMọi thắc mắc xin liên hệ ${settings.CONTACT_EMAIL} | Hotline: ${settings.SENDER_HOTLINE}.\n\n`;
                body += `Trân trọng,\n${settings.SENDER_NAME}`;

                MailApp.sendEmail({
                    to: String(recipientEmail).trim(),
                    subject: `${settings.EMAIL_SUBJECT_PREFIX} ${payrollMonthYear} - ${employeeName}`,
                    body: body,
                    attachments: attachments
                });

                // Cập nhật trạng thái ngay
                sourceSheet.getRange(currentRowNum, statusColIdx + 1).setValue("Thành công").setBackground("#d9ead3");
                logToSheet(ss, settings, {
                    monthYear: payrollMonthYear, row: currentRowNum, name: employeeName,
                    email: recipientEmail, status: "Thành công", url: driveUrl
                });
                stats.success++;
            } catch (err) {
                console.error(`Lỗi tại hàng ${currentRowNum}: ${err.message}`);
                sourceSheet.getRange(currentRowNum, statusColIdx + 1).setValue("Lỗi: " + err.message).setBackground("#f4cccc");
                logToSheet(ss, settings, {
                    monthYear: payrollMonthYear, row: currentRowNum, name: employeeName,
                    email: recipientEmail, status: "Lỗi: " + err.message, url: ""
                });
                stats.failed.push({ name: employeeName, row: currentRowNum, error: err.message });
            }
        }

        // Hoàn tất toàn bộ
        deleteResumeTrigger();
        clearTargetMonth();
        sendSummaryEmail(stats, "HOÀN TẤT");
        safeAlert("Hoàn tất", `Đã gửi thành công ${stats.success} email (bỏ qua ${stats.skipped} đã gửi trước). Vui lòng kiểm tra email báo cáo chi tiết.`);

    } finally {
        if (tempSS) {
            try { DriveApp.getFileById(tempSS.getId()).setTrashed(true); }
            catch (e) { console.warn("Không xóa được file tạm: " + e.message); }
        }
        lock.releaseLock();
    }
}

/**
 * Render mẫu TEMPLATE với dữ liệu 1 nhân viên -> trả về blob PDF.
 * Tự dọn sheet tạm sau khi xuất.
 */
function renderEmployeePdf(templateSheet, tempSS, row, settings, payrollMonthYear, rowNum, employeeName) {
    const tempSheet = templateSheet.copyTo(tempSS);
    const safeName = employeeName.substring(0, 25).replace(/[^\w\sÀ-ỹ]/g, '').trim() || ("NV_" + rowNum);
    tempSheet.setName(safeName + "_" + rowNum);

    const replacements = {
        "{{THANGNAM}}": payrollMonthYear,
        "{{SENDER_NAME}}": settings.SENDER_NAME,
        "{{SENDER_ADDRESS}}": settings.SENDER_ADDRESS,
        "{{SENDER_HOTLINE}}": settings.SENDER_HOTLINE,
        "{{CONTACT_EMAIL}}": settings.CONTACT_EMAIL
    };

    const MONEY_KEYS = ["THUONG", "TRU_CHIU_THUE", "TRU_KG_CHIU_THUE", "TONGLUONG", "GIAMTRU", "BHXH",
        "THUETNCN", "THUCLINH", "PHAT", "TAMUNG", "DANHAN", "LUONGCK"];
    const moneyFmt = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 0 });

    for (const key in settings.MAP) {
        let val = row[columnLetterToIndex(settings.MAP[key])];
        if (key.startsWith("L_") || key.startsWith("PC_") || MONEY_KEYS.includes(key)) {
            if (val === null || val === undefined || val === "") {
                val = "0";
            } else if (typeof val === 'number') {
                val = moneyFmt.format(Math.round(val)); // tiền VND: làm tròn, bỏ phần thập phân
            } else {
                // Ô tiền dạng TEXT: giữ nguyên thay vì parseFloat — parseFloat("1,234,567") = 1 sẽ ra số sai
                val = String(val);
            }
        }
        replacements[`{{${key}}}`] = val;
    }

    for (const tag in replacements) {
        tempSheet.createTextFinder(tag).replaceAllWith(String(replacements[tag]));
    }

    SpreadsheetApp.flush();

    // Ẩn mọi sheet khác để PDF chỉ chứa phiếu của nhân viên này.
    // tempSheet vẫn hiển thị nên không vi phạm quy tắc "không được ẩn hết sheet".
    const others = tempSS.getSheets().filter(s => s.getName() !== tempSheet.getName());
    let blob;
    try {
        others.forEach(s => s.hideSheet());
        blob = tempSS.getAs('application/pdf')
            .setName(`${settings.PDF_FILE_NAME_PREFIX}_${employeeName.replace(/\s+/g, '_')}_${payrollMonthYear.replace(/[\/\s]+/g, '-')}.pdf`);
    } finally {
        // Luôn hiện lại sheet nền (kể cả khi getAs lỗi) để không kẹt ở trạng thái ẩn hết sheet
        others.forEach(s => s.showSheet());
        tempSS.deleteSheet(tempSheet);
    }
    return blob;
}

// =============================================================================
//  DRIVE
// =============================================================================

/**
 * Lấy/ tạo thư mục kỳ lương: ROOT / "Năm YYYY" / "Tháng MM-YYYY".
 * Mỗi nhân viên = 1 file PDF trong thư mục tháng (tên file gồm tên NV).
 */
function getMonthFolder(settings, month, year) {
    let root = null;
    if (settings.DRIVE_ROOT_FOLDER_ID) {
        try { root = DriveApp.getFolderById(String(settings.DRIVE_ROOT_FOLDER_ID).trim()); }
        catch (e) { root = null; }
    }
    if (!root) {
        root = getOrCreateChildFolder(DriveApp.getRootFolder(), settings.DRIVE_ROOT_FOLDER_NAME);
    }
    const yearFolder = getOrCreateChildFolder(root, "Năm " + year);
    const mm = ("0" + month).slice(-2);
    return getOrCreateChildFolder(yearFolder, "Tháng " + mm + "-" + year);
}

function getOrCreateChildFolder(parent, name) {
    const it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
}

// =============================================================================
//  TRIGGER / TRẠNG THÁI THÁNG
// =============================================================================

function createResumeTrigger() {
    deleteResumeTrigger();
    ScriptApp.newTrigger("resumePayroll")
        .timeBased()
        .after(1 * 60 * 1000)
        .create();
}

function deleteResumeTrigger() {
    // CHỈ xóa trigger resume tự tạo. KHÔNG đụng tới trigger định kỳ mà người dùng có thể đã
    // tự đặt cho sendMonthlyPayrollEmails_WithPDF (vd auto-gửi hằng tháng).
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === "resumePayroll") {
            ScriptApp.deleteTrigger(t);
        }
    });
}

function setTargetMonth(month, year) {
    PropertiesService.getDocumentProperties().setProperties({
        PAYROLL_MONTH: String(month),
        PAYROLL_YEAR: String(year)
    });
}

function getTargetMonth() {
    const p = PropertiesService.getDocumentProperties();
    const now = new Date();
    const m = p.getProperty('PAYROLL_MONTH');
    const y = p.getProperty('PAYROLL_YEAR');
    return {
        month: m ? Number(m) : (now.getMonth() + 1),
        year: y ? Number(y) : now.getFullYear()
    };
}

function clearTargetMonth() {
    const p = PropertiesService.getDocumentProperties();
    p.deleteProperty('PAYROLL_MONTH');
    p.deleteProperty('PAYROLL_YEAR');
}

/**
 * Xóa trạng thái gửi của THÁNG ĐÍCH để gửi lại từ đầu.
 */
function resetSentStatus() {
    const settings = getSettings();
    const { month } = getTargetMonth();
    const sheetName = settings.SHEET_NAME_PREFIX + month;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        const lastRow = sheet.getLastRow();
        const startRow = Number(settings.START_ROW) || 2;
        if (lastRow >= startRow) {
            const statusColIdx = columnLetterToIndex(settings.MAP.SENT_STATUS);
            sheet.getRange(startRow, statusColIdx + 1, lastRow - startRow + 1, 1).clearContent().setBackground(null);
            SpreadsheetApp.getUi().alert("Đã Reset", `Đã xóa trạng thái gửi của tab "${sheetName}".`, SpreadsheetApp.getUi().ButtonSet.OK);
        }
    } else {
        SpreadsheetApp.getUi().alert("Lỗi", `Không tìm thấy tab "${sheetName}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

// =============================================================================
//  CONFIG / LOG / TIỆN ÍCH
// =============================================================================

/**
 * Lấy cấu hình: GLOBAL_CONFIG hợp nhất với tab "CONFIG" (nếu có).
 * Tab CONFIG: cột A = THAM SỐ, cột B = GIÁ TRỊ.
 *   - Key thường (vd SENDER_NAME, START_ROW, SAVE_TO_DRIVE) ghi đè trực tiếp.
 *   - Key dạng "MAP.HOTEN", "MAP.EMAIL"... ghi đè cột trong MAP.
 */
function getSettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("CONFIG");
    // Deep-copy MAP để không sửa GLOBAL_CONFIG gốc
    const settings = { ...GLOBAL_CONFIG, MAP: { ...GLOBAL_CONFIG.MAP } };

    if (configSheet) {
        const data = configSheet.getDataRange().getValues();
        for (let i = 0; i < data.length; i++) {
            const rawKey = String(data[i][0]).trim();
            const value = data[i][1];
            if (!rawKey || value === undefined || value === "") continue;

            if (rawKey.toUpperCase().startsWith("MAP.")) {
                const sub = rawKey.substring(4).toUpperCase();
                settings.MAP[sub] = String(value).trim();
            } else {
                const key = rawKey.toUpperCase();
                if (settings.hasOwnProperty(key)) {
                    settings[key] = coerceValue(value);
                }
            }
        }
    }
    return settings;
}

/**
 * Ép kiểu giá trị đọc từ CONFIG: boolean / number / string.
 */
function coerceValue(v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v;
    const s = String(v).trim();
    if (/^(true|có|yes|bật)$/i.test(s)) return true;
    if (/^(false|không|no|tắt)$/i.test(s)) return false;
    // Chỉ ép số khi toàn ký tự số (tránh phá ID thư mục Drive, tiền tố...)
    if (s !== "" && /^-?\d+$/.test(s)) return Number(s);
    return v;
}

/**
 * Ghi nhật ký gửi vào tab LOG (tạo nếu chưa có).
 */
function logToSheet(ss, settings, entry) {
    if (!settings.ENABLE_LOG) return;
    try {
        let sheet = ss.getSheetByName(settings.LOG_SHEET_NAME);
        if (!sheet) {
            sheet = ss.insertSheet(settings.LOG_SHEET_NAME);
            sheet.appendRow(["Thời gian", "Kỳ lương", "Hàng", "Họ tên", "Email", "Trạng thái", "Link Drive"]);
            sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#cfe2f3");
            sheet.setFrozenRows(1);
        }
        sheet.appendRow([
            new Date(), entry.monthYear, entry.row, entry.name,
            String(entry.email || ""), entry.status, entry.url || ""
        ]);
    } catch (e) {
        console.error("Không ghi được LOG: " + e.message);
    }
}

/**
 * Alert an toàn: nếu đang chạy trong trigger (không có UI) thì chỉ log, không ném lỗi.
 */
function safeAlert(title, message) {
    try {
        const ui = SpreadsheetApp.getUi();
        ui.alert(title, message, ui.ButtonSet.OK);
    } catch (e) {
        console.log(`[THÔNG BÁO] ${title}: ${message}`);
    }
}

/**
 * Chuyển tên cột (A, B, AA, CL) -> chỉ số 0-based.
 */
function columnLetterToIndex(columnLetter) {
    columnLetter = String(columnLetter).toUpperCase();
    let column = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        column *= 26;
        column += (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return column - 1;
}

// =============================================================================
//  MENU
// =============================================================================

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('SME Tools')
        .addItem('Gửi Phiếu Lương (Tháng Hiện Tại)', 'sendMonthlyPayrollEmails_WithPDF')
        .addItem('Gửi Phiếu Lương (Chọn Tháng...)', 'sendPayrollForSelectedMonth')
        .addItem('Gửi THỬ về email tôi...', 'sendTestPayslip')
        .addItem('Xóa trạng thái gửi (Reset)', 'resetSentStatus')
        .addSeparator()
        .addItem('Tạo Sheet Mẫu (TEMPLATE)', 'createSampleTemplate')
        .addItem('Tạo Sheet Cấu hình (CONFIG)', 'createConfigSheet')
        .addToUi();
}

// =============================================================================
//  TẠO SHEET MẪU TEMPLATE
// =============================================================================

function createSampleTemplate() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = getSettings();
    const sheetName = settings.TEMPLATE_SHEET_NAME;
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert('Cảnh báo', `Sheet "${sheetName}" đã tồn tại. Bạn có muốn xóa đi và tạo lại mẫu mới không?`, ui.ButtonSet.YES_NO);
        if (response !== ui.Button.YES) return;
        ss.deleteSheet(sheet);
    }

    sheet = ss.insertSheet(sheetName);
    sheet.setHideGridlines(true);

    // ── Bảng màu thương hiệu ANH & EM ─────────────────────────────────────
    const C = {
        INDIGO: "#415097", INK: "#242d56", INK_SOFT: "#2a304e", TINT: "#d0d4e7",
        TITLE: "#ebedf5", LINE: "#e2e5ef", MUTED: "#6c769e", WHITE: "#ffffff"
    };
    const SOLID = SpreadsheetApp.BorderStyle.SOLID;

    sheet.setColumnWidth(1, 290);
    sheet.setColumnWidth(2, 205);
    sheet.setColumnWidth(3, 205);
    sheet.getRange("A1:C35").setVerticalAlignment("middle").setFontFamily("Inter");

    // Helper: hàng nhãn → giá trị (căn trái = info, căn phải = money)
    const labelValue = (r, label, value, align) => {
        sheet.getRange(r, 1).setValue(label).setFontColor(C.INK_SOFT).setFontSize(10);
        sheet.getRange(r, 2, 1, 2).merge().setValue(value).setFontColor(C.INK).setFontSize(10).setHorizontalAlignment(align);
        sheet.setRowHeight(r, 24);
    };
    const sectionBand = (r, text) => {
        sheet.getRange(r, 1, 1, 3).merge().setBackground(C.INDIGO);
        sheet.getRange(r, 1).setValue(text).setFontColor(C.WHITE).setFontWeight("bold").setFontSize(10).setFontFamily("Figtree");
        sheet.setRowHeight(r, 24);
    };
    const totalRow = (r, label, value) => {
        sheet.getRange(r, 1, 1, 3).setBackground(C.TINT);
        sheet.getRange(r, 1).setValue(label).setFontColor(C.INK).setFontWeight("bold").setFontSize(11).setFontFamily("Figtree");
        sheet.getRange(r, 2, 1, 2).merge().setValue(value).setFontColor(C.INK).setFontWeight("bold").setFontSize(11).setHorizontalAlignment("right");
        sheet.setRowHeight(r, 24);
    };

    // ── Header band (logo + thông tin công ty) trên nền indigo ────────────
    sheet.getRange("A1:C3").setBackground(C.INDIGO);
    sheet.getRange("A1:A3").merge();
    sheet.getRange("A1").setValue("ANH & EM").setFontColor(C.WHITE).setFontWeight("bold").setFontSize(14).setFontFamily("Figtree").setHorizontalAlignment("center"); // fallback nếu không tải được logo
    sheet.getRange("B1:C1").merge().setValue("{{SENDER_NAME}}").setFontColor(C.WHITE).setFontWeight("bold").setFontSize(13).setFontFamily("Figtree");
    sheet.getRange("B2:C2").merge().setValue("{{SENDER_ADDRESS}}").setFontColor(C.WHITE).setFontSize(9);
    sheet.getRange("B3:C3").merge().setValue("Hotline: {{SENDER_HOTLINE}}  ·  {{CONTACT_EMAIL}}").setFontColor(C.WHITE).setFontSize(9);
    sheet.setRowHeights(1, 3, 26);
    // Nhúng logo (bytes) — render chắc chắn trong PDF, không vướng chặn fetch URL ngoài
    try {
        if (settings.LOGO_URL) {
            const blob = UrlFetchApp.fetch(settings.LOGO_URL).getBlob();
            sheet.insertImage(blob, 1, 1, 45, 8).setWidth(200).setHeight(64);
        }
    } catch (e) {
        console.warn("Không tải được logo (" + settings.LOGO_URL + "): " + e.message);
    }

    // ── Tiêu đề phiếu ─────────────────────────────────────────────────────
    sheet.setRowHeight(5, 38);
    sheet.getRange("A5:C5").merge().setValue("PHIẾU LƯƠNG  ·  {{THANGNAM}}")
        .setBackground(C.TITLE).setFontColor(C.INK).setFontWeight("bold").setFontSize(16).setFontFamily("Figtree").setHorizontalAlignment("center");

    // ── Thông tin nhân viên ───────────────────────────────────────────────
    sectionBand(7, "THÔNG TIN NHÂN VIÊN");
    labelValue(8, "Họ và tên", "{{HOTEN}}", "left");
    sheet.getRange(8, 2).setFontWeight("bold");
    labelValue(9, "Vị trí", "{{VITRI}}", "left");
    labelValue(10, "Tài khoản nhận", "{{STK}} ({{NGANHANG}})", "left");
    labelValue(11, "Email", "{{EMAIL}}", "left");

    // ── Chi tiết công xá ──────────────────────────────────────────────────
    sectionBand(13, "CHI TIẾT CÔNG XÁ");
    sheet.getRange(14, 1).setValue("Công chuẩn / thực tế").setFontColor(C.INK_SOFT).setFontSize(10);
    sheet.getRange(14, 2).setValue("{{NGAYCONGCHUAN}}").setFontColor(C.INK).setFontSize(10);
    sheet.getRange(14, 3).setValue("{{TONGGIOTT}} giờ").setFontColor(C.INK).setFontSize(10);
    sheet.setRowHeight(14, 24);
    labelValue(15, "Tăng ca (150% / 200% / 300%)", "{{GIO150}} / {{GIO200}} / {{GIO300}}", "left");
    labelValue(16, "Nghỉ (Ro / N / P / L)", "{{RO}} / {{N}} / {{P}} / {{L}}", "left");

    // ── Chi tiết thu nhập ─────────────────────────────────────────────────
    sectionBand(18, "CHI TIẾT THU NHẬP");
    const income = [
        ["Lương cơ bản", "{{L_COBAN}}"],
        ["Phụ cấp (ăn / điện thoại / đi lại)", "{{PC_ANTRUA}} / {{PC_DIENTHOAI}} / {{PC_DILAI}}"],
        ["Lương ngày công thực tế", "{{L_NGAYCONG}}"],
        ["Lương tăng ca (OT)", "{{L_OT}}"],
        ["Lương KPI / hiệu quả", "{{L_KPI}}"],
        ["Thưởng / phúc lợi", "{{THUONG}}"]
    ];
    income.forEach((pair, i) => labelValue(19 + i, pair[0], pair[1], "right"));
    totalRow(25, "TỔNG THU NHẬP", "{{TONGLUONG}}");

    // ── Khấu trừ & thực lĩnh ──────────────────────────────────────────────
    sectionBand(27, "KHẤU TRỪ & THỰC LĨNH");
    labelValue(28, "BHXH / thuế TNCN", "{{BHXH}} / {{THUETNCN}}", "right");
    labelValue(29, "Giảm trừ gia cảnh", "{{GIAMTRU}}", "right");
    totalRow(30, "THỰC LĨNH", "{{THUCLINH}}");

    // ── Thanh toán ────────────────────────────────────────────────────────
    labelValue(32, "Tạm ứng / đã quyết toán", "{{TAMUNG}} / {{DANHAN}}", "right");
    sheet.getRange(33, 1, 1, 3).setBackground(C.INDIGO);
    sheet.getRange(33, 1).setValue("SỐ TIỀN CHUYỂN KHOẢN").setFontColor(C.WHITE).setFontWeight("bold").setFontSize(12).setFontFamily("Figtree");
    sheet.getRange(33, 2, 1, 2).merge().setValue("{{LUONGCK}} VND").setFontColor(C.WHITE).setFontWeight("bold").setFontSize(15).setHorizontalAlignment("right");
    sheet.setRowHeight(33, 34);

    // ── Footer ────────────────────────────────────────────────────────────
    sheet.getRange("A35:C35").merge().setValue("Mọi thắc mắc vui lòng liên hệ {{CONTACT_EMAIL}} trong vòng 3 ngày làm việc.")
        .setFontColor(C.MUTED).setFontStyle("italic").setFontSize(8).setHorizontalAlignment("center");
    sheet.setRowHeight(35, 22);

    // Hàng cách (spacer)
    [4, 6, 12, 17, 26, 31, 34].forEach(r => sheet.setRowHeight(r, 8));

    // Đường kẻ ngang mảnh giữa các hàng dữ liệu
    [["A8:C11"], ["A14:C16"], ["A19:C25"], ["A28:C30"]].forEach(([rg]) =>
        sheet.getRange(rg).setBorder(null, null, null, null, null, true, C.LINE, SOLID));

    SpreadsheetApp.getUi().alert("Thành công",
        `Đã tạo xong mẫu "${sheetName}" theo nhận diện ANH & EM.\n` +
        `Logo lấy từ LOGO_URL trong CONFIG (đổi được). Bạn có thể tùy chỉnh thêm màu/font nếu muốn.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
}

// =============================================================================
//  TẠO SHEET CONFIG (đã mở rộng — đổi được cả mapping cột, Drive, bảo mật)
// =============================================================================

function createConfigSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "CONFIG";
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert('Cảnh báo', `Sheet "${sheetName}" đã tồn tại. Bạn có muốn xóa đi và tạo cấu hình mặc định không?`, ui.ButtonSet.YES_NO);
        if (response !== ui.Button.YES) return;
        ss.deleteSheet(sheet);
    }

    sheet = ss.insertSheet(sheetName);

    const header = [["THAM SỐ", "GIÁ TRỊ", "MÔ TẢ / GIẢI THÍCH"]];
    sheet.getRange(1, 1, 1, 3).setValues(header).setFontWeight("bold").setBackground("#cfe2f3");

    const rows = [
        ["--- THÔNG TIN CÔNG TY ---", "", ""],
        ["SENDER_NAME", GLOBAL_CONFIG.SENDER_NAME, "Tên công ty hiển thị trên phiếu lương"],
        ["SENDER_ADDRESS", GLOBAL_CONFIG.SENDER_ADDRESS, "Địa chỉ công ty"],
        ["SENDER_HOTLINE", GLOBAL_CONFIG.SENDER_HOTLINE, "Hotline hỗ trợ"],
        ["CONTACT_EMAIL", GLOBAL_CONFIG.CONTACT_EMAIL, "Email liên hệ hiển thị cho nhân viên (KHÔNG phải email gửi đi)"],
        ["LOGO_URL", GLOBAL_CONFIG.LOGO_URL, "URL ảnh logo hiển thị trên đầu phiếu (để trống = chỉ hiện tên công ty)"],

        ["--- CẤU TRÚC DỮ LIỆU ---", "", ""],
        ["SHEET_NAME_PREFIX", GLOBAL_CONFIG.SHEET_NAME_PREFIX, "Tiền tố tab tháng (T -> tab T1, T2...)"],
        ["START_ROW", GLOBAL_CONFIG.START_ROW, "Hàng bắt đầu có dữ liệu nhân viên (thường là 2)"],

        ["--- EMAIL ---", "", ""],
        ["EMAIL_SUBJECT_PREFIX", GLOBAL_CONFIG.EMAIL_SUBJECT_PREFIX, "Tiền tố tiêu đề email gửi đi"],
        ["PDF_FILE_NAME_PREFIX", GLOBAL_CONFIG.PDF_FILE_NAME_PREFIX, "Tiền tố tên file PDF"],
        ["ATTACH_PDF_TO_EMAIL", GLOBAL_CONFIG.ATTACH_PDF_TO_EMAIL, "true = đính kèm PDF vào email"],
        ["MIN_QUOTA", GLOBAL_CONFIG.MIN_QUOTA, "Số email tối thiểu còn lại để tiếp tục chạy"],

        ["--- LƯU TRỮ DRIVE ---", "", ""],
        ["SAVE_TO_DRIVE", GLOBAL_CONFIG.SAVE_TO_DRIVE, "true = lưu mỗi phiếu vào Drive theo Năm/Tháng"],
        ["DRIVE_ROOT_FOLDER_NAME", GLOBAL_CONFIG.DRIVE_ROOT_FOLDER_NAME, "Tên thư mục gốc (tạo tự động ở My Drive)"],
        ["DRIVE_ROOT_FOLDER_ID", GLOBAL_CONFIG.DRIVE_ROOT_FOLDER_ID, "(Tùy chọn) ID thư mục có sẵn — điền thì ưu tiên dùng"],

        ["--- BẢO MẬT & NHẬT KÝ ---", "", ""],
        ["SECURE_SHARE", GLOBAL_CONFIG.SECURE_SHARE, "true = file Drive chỉ chia sẻ cho đúng email NV + gửi link riêng (thay cho password PDF)"],
        ["ENABLE_LOG", GLOBAL_CONFIG.ENABLE_LOG, "true = ghi lịch sử gửi vào tab LOG"],

        ["--- VỊ TRÍ CỘT (chỉ sửa khi bảng lương khác mẫu) ---", "", ""],
        ["MAP.HOTEN", GLOBAL_CONFIG.MAP.HOTEN, "Cột chứa Họ tên nhân viên"],
        ["MAP.EMAIL", GLOBAL_CONFIG.MAP.EMAIL, "Cột chứa Email nhận phiếu"],
        ["MAP.TONGLUONG", GLOBAL_CONFIG.MAP.TONGLUONG, "Cột Tổng thu nhập"],
        ["MAP.THUCLINH", GLOBAL_CONFIG.MAP.THUCLINH, "Cột Thực lĩnh"],
        ["MAP.LUONGCK", GLOBAL_CONFIG.MAP.LUONGCK, "Cột Lương chuyển khoản"],
        ["MAP.SENT_STATUS", GLOBAL_CONFIG.MAP.SENT_STATUS, "Cột ghi trạng thái đã gửi (script tự ghi)"]
    ];

    sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    // Định dạng các dòng tiêu đề nhóm "--- ... ---"
    for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).startsWith("---")) {
            sheet.getRange(i + 2, 1, 1, 3).merge().setFontWeight("bold").setBackground("#e6e6e6").setFontStyle("italic");
        } else {
            sheet.getRange(i + 2, 2).setBackground("#fff2cc"); // ô giá trị màu vàng
        }
    }

    sheet.setColumnWidth(1, 230);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 430);
    sheet.getRange(1, 1, rows.length + 1, 3).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);

    SpreadsheetApp.getUi().alert("Thành công",
        `Đã tạo xong tab "${sheetName}". Chỉ cần sửa GIÁ TRỊ ở cột B (ô màu vàng).\n\n` +
        `Lưu ý: GAS không đặt được mật khẩu PDF — dùng SECURE_SHARE=true để chia sẻ riêng tư qua Drive.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
}
