// กำหนดค่าคงที่ใหม่ทั้งหมด
const SS_ID = '1OIByv0R5uv6LfUXH7MEgSMFU_UKIX8O3ujfYs5qJ-0Y'; // <--- กรุณาเปลี่ยนเป็น ID ของ Google Sheet ของคุณ
const SHEET_USERS = 'ข้อมูลเจ้าหน้าที่';
const SHEET_ACCESS = 'บันทึกการเข้าออกเขตพื้นที่';
const SHEET_SETTINGS = 'การตั้งค่า';
const ADMIN_EMAIL = 'sakonnakhondoa@gmail.com'; // อีเมลแอดมินสำหรับสำเนาอีเมลยืนยัน

/**
 * 1. ฟังก์ชันหลักสำหรับ WebApp
 */
function doGet(e) {
  if (e.parameter.page === 'scanner') {
    const template = HtmlService.createTemplateFromFile('scanner');
    template.scanMode = e.parameter.mode || 'Gate A';
    return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('ระบบสแกนเข้า-ออกพื้นที่')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  }
  if (e.parameter.page === 'register') {
    return HtmlService.createHtmlOutputFromFile('RegistrationForm')
      .setTitle('ลงทะเบียนเจ้าหน้าที่')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('AdminDashboard')
    .setTitle('ระบบแอดมินและแดชบอร์ด')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. การจัดการฐานข้อมูล (Google Sheets)
 */

// สร้างชีทและคอลัมน์อัตโนมัติ (รันครั้งแรกเท่านั้นจาก Apps Script Editor)
function createSheetsAndHeaders() {
    const ss = SpreadsheetApp.openById(SS_ID);
// ตรวจสอบและสร้าง/อัปเดตชีทข้อมูลเจ้าหน้าที่
    let userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) {
      userSheet = ss.insertSheet(SHEET_USERS);
    }
    

    // *** โค้ดที่กำหนดรูปแบบคอลัมน์ 'รหัสเจ้าหน้าที่' ให้เป็น Plain Text (@) ***
    if (userSheet) {
      const headers = userSheet.getRange('1:1').getValues()[0];
      const regIdColIndex = headers.indexOf('รหัสเจ้าหน้าที่');

      if (regIdColIndex !== -1) {
        const columnToFormat = regIdColIndex + 1;
// Apps Script คอลัมน์เริ่มจาก 1
        // กำหนดรูปแบบข้อความสำหรับคอลัมน์ รหัสเจ้าหน้าที่ (เริ่มตั้งแต่แถวที่ 2 ถึงบรรทัดสูงสุด)
        userSheet.getRange(2, columnToFormat, userSheet.getMaxRows() - 1, 1).setNumberFormat('@');
      }
    
    // โค้ดส่วนนี้จะถูกรันหาก userSheet ถูกสร้างขึ้นแล้ว (ถึงแม้จะดูเหมือนซ้ำกับ header ด้านบน แต่ถูกเก็บไว้ตามโค้ดเดิม)
    userSheet.appendRow([
      'รูปโปรไฟล์', 
      'รหัสเจ้าหน้าที่', 
      'ชื่อ-นามสกุล', 
      'หน่วยงาน/สังกัด', 
      'ตำแหน่ง', 
      'เบอร์โทรศัพท์', 
      'เลขที่', 
      'หมายเลขเขตพื้นที่', 
      'วันหมดอายุ', 
      'ประเภทบัตร', 
      'Bar Code',
      'QR Code',
      'Email Status', 
      'Timestamp',
      'หมายเหตุ'
    ]);
    }
  
  // บันทึกการเข้าออกเขตพื้นที่
  let accessSheet = ss.getSheetByName(SHEET_ACCESS);
  if (!accessSheet) {
    accessSheet = ss.insertSheet(SHEET_ACCESS);
    accessSheet.appendRow([
      'รหัสเจ้าหน้าที่', 
      'วันเวลาที่สแกน', 
      'สถานะ (เข้า/ออก)', 
      'ประตู', 
      'ชื่อ-นามสกุล', 
      'เลขที่', 
      'วันหมดอายุ', 
      'หน่วยงาน/สังกัด', 
      'ตำแหน่ง',
      'หมายเลขเขตพื้นที่',
      'ประเภทบัตร', 
      'วิทยุ', 
      'มือถือ', 
      'อุปกรณ์/เครื่องมือ', 
      'Escort No', 
      'Escort Name',
      'เลขที่ Escost', 
      'ประเภท Escost', 
      'หมายเหตุ'
    ]);
  }
}

// Helper function: ตรวจสอบรหัสเจ้าหน้าที่ซ้ำซ้อน
function isRegIdDuplicate(regId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const userSheet = ss.getSheetByName(SHEET_USERS);
// *** เพิ่มการตรวจสอบว่าชีทมีอยู่หรือไม่
  if (!userSheet) return false; 
  
  const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
  const regIdColIndex = headers.indexOf('รหัสเจ้าหน้าที่') + 1;
  if (regIdColIndex <= 0) return false;
  
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // ใช้ String() เพื่อเปรียบเทียบค่าที่อาจมาจากคอลัมน์ที่เป็นตัวเลข/ข้อความ
    if (String(data[i][regIdColIndex - 1]) === regId) {
      return true;
    }
  }
  return false;
}

// Function to generate BarCode and QRCode 
function generateCodes(regId) {
  const barcodeUrl = `https://barcode.tec-it.com/barcode.ashx?data=${regId}&code=Code128&dpi=96`;
  const qrcodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${regId}`;
  return { barcodeUrl, qrcodeUrl };
}

// **ฟังก์ชันประมวลผลการลงทะเบียน**
function processRegistration(formData) {
  // *** โค้ดที่ถูกแก้ไข 1: แปลงวันหมดอายุเป็นรูปแบบ DD/MM/YYYY ***
  if (formData.expiryDate) {
    // แยกส่วนวันที่ YYYY-MM-DD
    const dateParts = formData.expiryDate.split('-');
    // สร้าง Date object จาก YYYY, MM-1, DD (เพื่อหลีกเลี่ยงปัญหา Timezone/Locale)
    const dateObject = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]); 
    // ใช้ Utilities.formatDate เพื่อให้ได้รูปแบบ dd/MM/yyyy (Text String)
    formData.expiryDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  // REMOVED: createSheetsAndHeaders();
  const ss = SpreadsheetApp.openById(SS_ID);
  const userSheet = ss.getSheetByName(SHEET_USERS);

  // ตรวจสอบว่าชีทมีอยู่หรือไม่
  if (!userSheet) {
    return { success: false, message: 'ไม่พบชีท "ข้อมูลเจ้าหน้าที่" กรุณารัน createSheetsAndHeaders() ก่อน' };
  }

  // 1. ตรวจสอบความซ้ำซ้อน
  if (isRegIdDuplicate(formData.regId)) {
    return { success: false, message: `รหัสเจ้าหน้าที่ ${formData.regId} มีอยู่แล้วในระบบ` };
  }

  // 2. สร้าง Bar Code และ QR Code URL
  const { barcodeUrl, qrcodeUrl } = generateCodes(formData.regId);
// 3. จัดเรียงข้อมูลตามคอลัมน์ใหม่
  const newRow = [
    formData.profileUrl,
    String(formData.regId), 
    formData.fullName,
    formData.department,
    formData.position,
    formData.phone,
    formData.driverLicenseNo, 
    formData.aviationAreas, 
    formData.expiryDate, 
    formData.cardType, 
    barcodeUrl,
    qrcodeUrl,
    'Pending', 
    new Date(),
    formData.note ||
''
  ];
  
  userSheet.appendRow(newRow);

  // 4. ส่งอีเมลแจ้งเตือนเฉพาะแอดมิน (ใช้ Email Template)
  try {
    // โหลดแม่แบบอีเมลจากไฟล์ 'EmailTemplate.html'
    const emailTemplate = HtmlService.createTemplateFromFile('EmailTemplate');
// กำหนดตัวแปรสำหรับแทนที่ในแม่แบบ
    emailTemplate.formData = {
      profileUrl: formData.profileUrl,
      regId: formData.regId,
      fullname: formData.fullName, 
      department: formData.department, 
      position: formData.position,
      phone: formData.phone,
      driverLicenseNo: formData.driverLicenseNo, 
      aviationAreas: formData.aviationAreas,
      cardType: formData.cardType,
      expiryDate: formData.expiryDate // ** ใช้ค่าที่ถูกแปลงเป็น DD/MM/YYYY แล้ว **
    };
    emailTemplate.barcodeUrl = barcodeUrl;
    emailTemplate.qrcodeUrl = qrcodeUrl;
    
    // สร้างเนื้อหาอีเมล HTML และกำหนดหัวเรื่อง
    const htmlBody = emailTemplate.evaluate().getContent();
    const subject = `[สำเร็จ] ผู้ใช้งานใหม่: ${formData.regId} - ${formData.fullName}`;

    // ส่งอีเมลถึงแอดมิน
    MailApp.sendEmail({ 
      to: ADMIN_EMAIL, 
      subject: subject, 
      htmlBody: htmlBody, // ใช้ htmlBody แทน body
      name: 'ระบบลงทะเบียนท่าอากาศยานสกลนคร' 
    });
  } catch (error) {
    Logger.log(`Failed to send email to admin: ${error.message}`);
  }

  return { success: true, regId: formData.regId, barcodeUrl: barcodeUrl, qrcodeUrl: qrcodeUrl };
}

/**
 * 2. การจัดการฐานข้อมูล (Google Sheets)
 */

// ฟังก์ชันดึงข้อมูลผู้ใช้จากชีทหลัก
function getUserByRegId(regId) {
    if (!regId) return null;
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const userSheet = ss.getSheetByName(SHEET_USERS);
        if (!userSheet) return null;

        const dataRange = userSheet.getDataRange().getValues();
        if (dataRange.length <= 1) return null;

        const headers = dataRange[0];
        const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');
        const fullNameCol = headers.indexOf('ชื่อ-นามสกุล');
        const driverLicenseNoCol = headers.indexOf('เลขที่');
        const expiryDateCol = headers.indexOf('วันหมดอายุ');
        const departmentCol = headers.indexOf('หน่วยงาน/สังกัด');
        const positionCol = headers.indexOf('ตำแหน่ง');
        const aviationAreasCol = headers.indexOf('หมายเลขเขตพื้นที่');
        const cardTypeCol = headers.indexOf('ประเภทบัตร');
        const profileUrlCol = headers.indexOf('รูปโปรไฟล์');
        const phoneCol = headers.indexOf('เบอร์โทรศัพท์'); // เพิ่มเบอร์โทรศัพท์

        for (let i = 1; i < dataRange.length; i++) {
            const row = dataRange[i];
            // ใช้ String(row[index]).trim() เพื่อให้มั่นใจว่าข้อมูลตรงกันและไม่มีช่องว่าง
            if (String(row[regIdCol]).trim() === regId) {
                const expiry = row[expiryDateCol];
                const today = new Date();
                
                let isExpired = false;
                let expiryDateString = 'N/A';
                
                if (expiry && expiry instanceof Date) {
                    expiryDateString = Utilities.formatDate(expiry, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
                    isExpired = expiry < today;
                } else if (typeof expiry === 'string' && expiry.length > 0) {
                    expiryDateString = expiry; // เก็บเป็นข้อความหากไม่ใช่ Date Object
                    // ไม่สามารถตรวจสอบการหมดอายุได้ถ้าเป็น String
                }

                const cardType = row[cardTypeCol];
                const requiresEscort = cardType === 'ชั่วคราว'; 

                return {
                    regId: regId,
                    fullName: row[fullNameCol] || 'ไม่ระบุชื่อ',
                    driverLicenseNo: row[driverLicenseNoCol] || 'N/A',
                    expiryDate: expiryDateString,
                    department: row[departmentCol] || 'N/A',
                    position: row[positionCol] || 'N/A',
                    aviationAreas: row[aviationAreasCol] || 'N/A',
                    cardType: cardType || 'N/A',
                    profileUrl: row[profileUrlCol] || 'https://placehold.co/150x150/0F1115/E6E9EF/png?text=NO+IMG',
                    phone: row[phoneCol] || 'N/A', // เพิ่มเบอร์โทรศัพท์
                    isExpired: isExpired,
                    requiresEscort: requiresEscort,
                };
            }
        }
    } catch(error) {
        Logger.log("Error in getUserByRegId: " + error.message);
        throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลผู้ใช้: " + error.message);
    }
    return null;
}

// ฟังก์ชันดึงสถานะล่าสุดจากชีทบันทึก
function getLastAccessStatus(regId) {
    if (!regId) return 'N/A';
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const accessSheet = ss.getSheetByName(SHEET_ACCESS);
        if (!accessSheet) return 'N/A';

        const dataRange = accessSheet.getDataRange().getValues();
        if (dataRange.length <= 1) return 'N/A';

        const headers = dataRange[0];
        const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');
        const statusCol = headers.indexOf('สถานะ (เข้า/ออก)');
        const radioCol = headers.indexOf('วิทยุ');
        const mobileCol = headers.indexOf('มือถือ');
        const equipmentCol = headers.indexOf('อุปกรณ์/เครื่องมือ');
        const escortNameCol = headers.indexOf('Escort Name');
        const escortRegIdCol = headers.indexOf('Escort RegId');
        
        let latestStatus = 'ออก';
        let latestEquipmentData = null; 
        let escortDataForCheckout = {};

        // ค้นหาจากล่างขึ้นบน
        for (let i = dataRange.length - 1; i >= 1; i--) {
            const row = dataRange[i];
            if (String(row[regIdCol]).trim() === regId) {
                latestStatus = row[statusCol] || 'ออก';
                
                // เก็บข้อมูลอุปกรณ์และ Escort ล่าสุดที่สถานะ 'เข้า'
                if (latestStatus === 'เข้า') {
                    latestEquipmentData = {
                        radio: (radioCol !== -1 && String(row[radioCol]).trim() === '/'),
                        mobile: (mobileCol !== -1 && String(row[mobileCol]).trim() === '/'),
                        equipment: (equipmentCol !== -1 && String(row[equipmentCol]).trim() === '/'),
                    };
                    
                    if (escortNameCol !== -1 && row[escortNameCol]) {
                         escortDataForCheckout = {
                            escortName: row[escortNameCol],
                            escortRegId: row[escortRegIdCol] || ''
                        };
                    }
                }
                
                return {
                    currentStatus: latestStatus,
                    latestEquipmentData: latestEquipmentData,
                    escortDataForCheckout: escortDataForCheckout
                };
            }
        }
    } catch(error) {
        Logger.log("Error in getLastAccessStatus: " + error.message);
    }
    return { 
        currentStatus: 'ออก',
        latestEquipmentData: null,
        escortDataForCheckout: {}
    };
}

/**
 * 3. ฟังก์ชันสำหรับ Scanner.html
 */

// ฟังก์ชันสำหรับสแกนบัตร (Main Logic)
function getUserForScan(regId) {
    if (!regId) {
        return { success: false, message: 'กรุณาสแกนรหัสเจ้าหน้าที่' };
    }
    
    // 1. ดึงข้อมูลผู้ใช้จากชีทหลัก
    const user = getUserByRegId(regId);
    if (!user) {
        return { success: false, message: `ไม่พบรหัสเจ้าหน้าที่ ${regId} ในระบบ`, data: null };
    }

    if (user.isExpired) {
        return { success: false, message: `บัตรหมดอายุแล้ว: ${user.expiryDate}`, data: user };
    }

    // 2. ดึงสถานะการเข้าออกล่าสุด
    const accessInfo = getLastAccessStatus(regId);
    
    // 3. กำหนดสถานะใหม่
    const newStatus = (accessInfo.currentStatus === 'เข้า') ? 'ออก' : 'เข้า';
    
    // 4. ตรวจสอบเงื่อนไข Escort
    // บัตรชั่วคราว (Temporary) ต้องมีการรับรองเมื่อ 'เข้า' เท่านั้น
    const requiresEscort = (newStatus === 'เข้า' && user.requiresEscort);
    
    return {
        success: true,
        data: user,
        currentStatus: accessInfo.currentStatus,
        newStatus: newStatus,
        requiresEscort: requiresEscort,
        latestEquipmentData: accessInfo.latestEquipmentData,
        escortDataForCheckout: accessInfo.escortDataForCheckout
    };
}

// ==========================================
// แก้ไขฟังก์ชัน 1: ตรวจสอบสิทธิ์คนรับรอง + นับโควตา (Logic ใหม่)
// ==========================================
function checkEscort(escortRegId) {
    const MAX_QUOTA = 5; // กำหนดโควตาสูงสุด

    // 1. ตรวจสอบข้อมูลคนรับรองพื้นฐาน
    const escortUser = getUserByRegId(escortRegId);
    if (!escortUser) {
        return { success: false, message: `ไม่พบรหัสผู้รับรอง ${escortRegId} ในระบบ` };
    }
    
    // ตรวจสอบประเภทบัตร (ต้องเป็น "ถาวร" เท่านั้น)
    if (escortUser.cardType !== 'ถาวร') {
        return { success: false, message: `เจ้าหน้าที่ที่จะรับรองต้องเป็นผู้ถือบัตร "ถาวร" เท่านั้น (บัตรใบนี้ : ${escortUser.cardType})` };
    }
    
    // ตรวจสอบวันหมดอายุ
    if (escortUser.isExpired) {
        return { success: false, message: `บัตรผู้รับรองหมดอายุแล้ว: ${escortUser.expiryDate}` };
    }

    // 2. นับจำนวนผู้ที่ Escort คนนี้กำลังรับรองอยู่ในพื้นที่ (Status = เข้า)
    const insideUsers = getInsideUsers();
    // กรองหาคนที่ Escort RegId ตรงกัน
    const currentEscortCount = insideUsers.filter(u => String(u.escortRegId).trim() === String(escortRegId).trim()).length;

    // 3. ตรวจสอบโควตา (แบบ B: สะสมยอด)
    if (currentEscortCount >= MAX_QUOTA) {
        return { 
            success: false, 
            message: `❌ โควตาเต็ม! (${currentEscortCount}/${MAX_QUOTA}) ❌ คุณรับรองครบ 5 คนแล้ว  กรุณาเปลี่ยนคนรับรอง` 
        };
    }

    // 4. ผ่านทุกเงื่อนไข
    return {
        success: true,
        escortData: {
            escortRegId: escortRegId,
            escortName: escortUser.fullName,
            escortDriverLicenseNo: escortUser.driverLicenseNo,
            escortCardType: escortUser.cardType,
            escortCount: currentEscortCount, // ส่งยอดปัจจุบันไปแสดงผล
            quotaMax: MAX_QUOTA
        }
    };
}

// **FIXED & UPDATED** บันทึกข้อมูลการเข้าออก
function recordAccess(regId, newStatus, scanMode, radio, mobile, equipment, escortData = {}) {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const accessSheet = ss.getSheetByName(SHEET_ACCESS);
        
        // 1. ตรวจสอบว่าชีทมีอยู่หรือไม่
        if (!accessSheet) {
            // โยนข้อผิดพลาดหากไม่พบชีทเพื่อ trigger onRecordFailure
            throw new Error(`ไม่พบชีท "${SHEET_ACCESS}" กรุณาตรวจสอบชื่อชีทหรือ SS_ID`);
        }

        const user = getUserByRegId(regId);
        if (!user) {
            throw new Error('ไม่พบข้อมูลผู้ใช้งานที่ต้องการบันทึก');
        }
        
        const timestamp = new Date();
        
        // ข้อมูล Escort (ใช้ข้อมูลที่ส่งมาจาก UI)
        const escortRegId = escortData.escortRegId || '';
        const escortName = escortData.escortName || '';
        const escortDriverLicenseNo = escortData.escortDriverLicenseNo || '';
        const escortCardType = escortData.escortCardType || '';

        // จัดเรียงข้อมูลตามคอลัมน์ของชีท "บันทึกการเข้าออกเขตพื้นที่"
        // ********* ตรวจสอบลำดับคอลัมน์ใน Sheet ให้ตรงกับโค้ดนี้ *********
        const newRow = [
            String(regId), 
            timestamp,
            newStatus,
            scanMode,
            user.fullName,
            user.driverLicenseNo, 
            user.expiryDate, 
            user.department,
            user.position,
            user.aviationAreas, // เพิ่มเขตพื้นที่
            user.cardType, 
            radio ? '/' : '',
            mobile ? '/' : '', 
            equipment ? '/' : '',
            escortRegId, 
            escortName, 
            escortDriverLicenseNo, 
            escortCardType, 
            '' // หมายเหตุ
        ];
        
        // 2. บันทึกข้อมูล
        accessSheet.appendRow(newRow); 
        
        Logger.log(`บันทึกสำเร็จสำหรับ ${user.fullName}, สถานะ: ${newStatus}`);

        return { 
            success: true, 
            message: `บันทึกการ ${newStatus} สำเร็จ!`,
            fullName: user.fullName,
            newStatus: newStatus
        };
        
    } catch(error) {
        // 3. ดักจับข้อผิดพลาดและโยนกลับไปให้ Scanner.html ทันที
        Logger.log(`[RECORD ACCESS ERROR] RegID: ${regId}, Status: ${newStatus}, Error: ${error.message}`);
        throw new Error(`บันทึกไม่สำเร็จ: ${error.message}`); 
    }
}


/**
 * 4. ฟังก์ชันสำหรับ AdminDashboard.html
 */

// ==========================================
// แก้ไขฟังก์ชัน 2: ดึงรายชื่อคนในพื้นที่ (แก้ Bug ชื่อคอลัมน์)
// ==========================================
function getInsideUsers() {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const accessSheet = ss.getSheetByName(SHEET_ACCESS);
        const userSheet = ss.getSheetByName(SHEET_USERS); // ใช้เพื่อ join ข้อมูลเพิ่มถ้าจำเป็น

        if (!accessSheet) return [];

        const accessData = accessSheet.getDataRange().getValues();
        if (accessData.length <= 1) return [];

        const headers = accessData[0];
        
        // Mapping Column Index (แก้ชื่อ 'Escort RegId' เป็น 'Escort No' หรือหาทั้งคู่เพื่อความชัวร์)
        const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');
        const statusCol = headers.indexOf('สถานะ (เข้า/ออก)');
        const timestampCol = headers.indexOf('วันเวลาที่สแกน');
        const gateCol = headers.indexOf('ประตู');
        const escortNameCol = headers.indexOf('Escort Name');
        
        // ** จุดที่แก้ไข: หาคอลัมน์ Escort ID ให้เจอไม่ว่าจะชื่ออะไร **
        let escortRegIdCol = headers.indexOf('Escort No'); 
        if (escortRegIdCol === -1) escortRegIdCol = headers.indexOf('Escort RegId'); // เผื่อกรณีชื่อไม่ตรง

        if (regIdCol === -1 || statusCol === -1) return [];

        const currentStatusMap = {};

        // 1. วนลูปจากล่างขึ้นบนเพื่อหาสถานะล่าสุดของแต่ละคน
        for (let i = accessData.length - 1; i >= 1; i--) {
            const row = accessData[i];
            const regId = String(row[regIdCol]).trim();
            const status = row[statusCol] || 'ออก';

            if (!currentStatusMap.hasOwnProperty(regId)) {
                // เก็บข้อมูลล่าสุดของคนนี้
                currentStatusMap[regId] = {
                    status: status,
                    timestamp: row[timestampCol],
                    gate: row[gateCol] || '-',
                    escortName: row[escortNameCol] || '-',
                    escortRegId: (escortRegIdCol !== -1) ? row[escortRegIdCol] : '' // เก็บ ID คนรับรอง
                };
            }
        }
        
        // 2. กรองเอาเฉพาะคนที่สถานะล่าสุดคือ 'เข้า'
        const insideUsersList = [];
        for (const [regId, info] of Object.entries(currentStatusMap)) {
            if (info.status === 'เข้า') {
                // ดึงข้อมูล User เพิ่มเติม (Optional)
                const userProfile = getUserByRegId(regId);
                
                insideUsersList.push({
                    regId: regId,
                    fullName: userProfile ? userProfile.fullName : 'ไม่ระบุชื่อ',
                    department: userProfile ? userProfile.department : '-',
                    position: userProfile ? userProfile.position : '-',
                    cardType: userProfile ? userProfile.cardType : '-',
                    profileUrl: userProfile ? userProfile.profileUrl : '',
                    timeIn: info.timestamp instanceof Date ? Utilities.formatDate(info.timestamp, ss.getSpreadsheetTimeZone(), 'HH:mm:ss dd/MM/yyyy') : String(info.timestamp),
                    gate: info.gate,
                    escortName: info.escortName,
                    escortRegId: info.escortRegId // สำคัญมากสำหรับการนับโควตา
                });
            }
        }
        
        return insideUsersList;

    } catch(error) {
        Logger.log("Error in getInsideUsers: " + error.message);
        return [];
    }
}

// ดึงรายละเอียดผู้ใช้ที่อยู่ในพื้นที่ (สำหรับ Pop-up)
function getInsideUserDetail(regId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const user = getUserByRegId(regId);
  if (!user) {
    return { success: false, message: 'ไม่พบผู้ใช้ในชีท "ข้อมูลเจ้าหน้าที่"' };
  }

  const accessSheet = ss.getSheetByName(SHEET_ACCESS);
  if (!accessSheet) return { success: true, ...user, lastAccess: null };

  const dataRange = accessSheet.getDataRange().getValues();
  const headers = dataRange[0];
  
  const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');
  const statusCol = headers.indexOf('สถานะ (เข้า/ออก)');
  const timestampCol = headers.indexOf('วันเวลาที่สแกน');
  const gateCol = headers.indexOf('ประตู');
  const escortNameCol = headers.indexOf('Escort Name');
  const radioCol = headers.indexOf('วิทยุ');
  const mobileCol = headers.indexOf('มือถือ');
  const equipmentCol = headers.indexOf('อุปกรณ์/เครื่องมือ');
  
  // *** เพิ่มการตรวจสอบคอลัมน์หลักเพื่อป้องกันการ Crash ***
  if ([regIdCol, statusCol, timestampCol, gateCol].includes(-1)) {
     Logger.log("Missing essential access log headers for getInsideUserDetails.");
     return { success: true, ...user, lastAccess: null, message: "ขาดคอลัมน์หลักในชีทบันทึกการเข้าออก" };
  }

  let lastCheckIn = null;
  
  // ค้นหาจากล่างขึ้นบน
  for (let i = dataRange.length - 1; i >= 1; i--) {
    const row = dataRange[i];
    if (String(row[regIdCol]).trim() === regId) {
        if (row[statusCol] === 'เข้า') {
            lastCheckIn = {
                timestamp: Utilities.formatDate(row[timestampCol], ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
                gate: row[gateCol] || '-',
                
                // FIX: ตรวจสอบ Index ก่อนเข้าถึงค่าในแถว (row)
                escortName: (escortNameCol !== -1 ? row[escortNameCol] : '') || 'ไม่มี',
                radio: (radioCol !== -1 ? String(row[radioCol]).trim() === '/' : false),
                mobile: (mobileCol !== -1 ? String(row[mobileCol]).trim() === '/' : false),
                equipment: (equipmentCol !== -1 ? String(row[equipmentCol]).trim() === '/' : false),
            };
            break; // พบรายการ 'เข้า' ล่าสุดแล้ว ออกจากลูป
        }
    }
  }

  // รวมข้อมูล User และข้อมูลการเข้าล่าสุด
  return { 
    success: true, 
    ...user, 
    lastAccess: lastCheckIn 
  };
}

// ฟังก์ชันสำหรับดึงบันทึกการเข้าออกทั้งหมด (ต้องใช้ใน Access Logs Tab)
function getAllAccessLogs() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const accessSheet = ss.getSheetByName(SHEET_ACCESS);

  if (!accessSheet) return [];

  const dataRange = accessSheet.getDataRange().getValues();
  if (dataRange.length <= 1) return [];

  const headers = dataRange[0];
  
  // Mapping Indices
  const colMap = {
    timestamp: headers.indexOf('วันเวลาที่สแกน'),
    status: headers.indexOf('สถานะ (เข้า/ออก)'),
    regId: headers.indexOf('รหัสเจ้าหน้าที่'),
    fullName: headers.indexOf('ชื่อ-นามสกุล'),
    gate: headers.indexOf('ประตู'),
    department: headers.indexOf('หน่วยงาน/สังกัด'), 
    position: headers.indexOf('ตำแหน่ง'), 
    aviationAreas: headers.indexOf('หมายเลขเขตพื้นที่'), 
    cardType: headers.indexOf('ประเภทบัตร'),
    escortName: headers.indexOf('Escort Name'),
  };
  
  const logs = [];
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    
    // ตรวจสอบความถูกต้องของ Timestamp ก่อน Format
    let formattedTimestamp = '';
    let rawTimestamp = row[colMap.timestamp];
    if (rawTimestamp instanceof Date) {
        formattedTimestamp = Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone(), 'HH:mm:ss dd/MM/yyyy');
    } else {
        formattedTimestamp = String(rawTimestamp);
    }
    
    logs.push({
      timestamp: formattedTimestamp,
      status: row[colMap.status] || '',
      regId: String(row[colMap.regId]) || '',
      fullName: row[colMap.fullName] || '',
      gate: row[colMap.gate] || '',
      department: row[colMap.department] || '',
      position: row[colMap.position] || '',
      aviationAreas: row[colMap.aviationAreas] || '',
      cardType: row[colMap.cardType] || '',
      escortName: row[colMap.escortName] || '',
    });
  }
  // ส่งบันทึกย้อนหลัง 500 รายการล่าสุด
  return logs.reverse().slice(0, 500); 
}


// ---------------------------------------------------------------- //
// 5. ฟังก์ชันสำหรับค้นหาบันทึกการเข้าออกรายบุคคล (แก้ไขล่าสุด + เพิ่ม Date Range Filter)
// ---------------------------------------------------------------- //
function getIndividualAccessLogs(regId, dateFromStr, dateToStr) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const accessSheet = ss.getSheetByName(SHEET_ACCESS);

    if (!accessSheet) {
      throw new Error(`ไม่พบชีท "${SHEET_ACCESS}" กรุณาตรวจสอบชื่อชีท`);
    }

    const searchRegId = String(regId).trim();
    if (!searchRegId) return [];

    // แปลงวันที่เริ่มต้นและสิ้นสุด (ถ้ามี)
    // การใช้ 'T00:00:00' และ 'T23:59:59' เพื่อให้รวมข้อมูลตั้งแต่ต้นวันถึงสิ้นวัน
    let dateFrom = dateFromStr ? new Date(dateFromStr + 'T00:00:00') : null; 
    let dateTo = dateToStr ? new Date(dateToStr + 'T23:59:59') : null;   

    const dataRange = accessSheet.getDataRange().getValues();
    if (dataRange.length <= 1) return []; 

    const headers = dataRange[0];
    
    const colMap = {
      timestamp: headers.indexOf('วันเวลาที่สแกน'),
      status: headers.indexOf('สถานะ (เข้า/ออก)'),
      regId: headers.indexOf('รหัสเจ้าหน้าที่'),
      fullName: headers.indexOf('ชื่อ-นามสกุล'),
      gate: headers.indexOf('ประตู'),
      department: headers.indexOf('หน่วยงาน/สังกัด'),
      position: headers.indexOf('ตำแหน่ง'),
      cardType: headers.indexOf('ประเภทบัตร'),
      escortName: headers.indexOf('Escort Name'),
    };
    
    if (colMap.regId === -1 || colMap.timestamp === -1 || colMap.status === -1) {
       throw new Error("ขาดหัวคอลัมน์สำคัญในชีทบันทึกการเข้าออก");
    }

    const logs = [];
    
    for (let i = 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      const currentRegId = String(row[colMap.regId]).trim();

      if (currentRegId === searchRegId) {
        
        let rawTimestamp = row[colMap.timestamp];
        let formattedTimestamp = '';
        
        if (rawTimestamp instanceof Date) {
            // กรองตามช่วงวันที่ที่กำหนด
            if (dateFrom && rawTimestamp < dateFrom) continue;
            if (dateTo && rawTimestamp > dateTo) continue;

            formattedTimestamp = Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone(), 'HH:mm:ss dd/MM/yyyy');
        } else {
           formattedTimestamp = String(rawTimestamp);
        }

        logs.push({
          timestamp: formattedTimestamp,
          status: row[colMap.status] || '',
          regId: currentRegId,
          fullName: (colMap.fullName !== -1 ? row[colMap.fullName] : '') || '',
          gate: (colMap.gate !== -1 ? row[colMap.gate] : '') || '',
          department: (colMap.department !== -1 ? row[colMap.department] : '') || '',
          position: (colMap.position !== -1 ? row[colMap.position] : '') || '',
          cardType: (colMap.cardType !== -1 ? row[colMap.cardType] : '') || '',
          escortName: (colMap.escortName !== -1 ? row[colMap.escortName] : '') || '',
        });
      }
    }
    
    // ส่งบันทึกย้อนหลัง โดยรายการล่าสุดอยู่ด้านบน
    return logs.reverse();
    
  } catch(error) {
    Logger.log(`[getIndividualAccessLogs ERROR] RegID: ${regId}, Error: ${error.message}`);
    throw new Error(`การค้นหาบันทึกไม่สำเร็จ: ${error.message}`);
  }
}


// ฟังก์ชันสำหรับ Autocomplete (ค้นหาชื่อ/รหัส)
function getUserSuggestions(searchTerm) {
    if (!searchTerm || searchTerm.length < 2) return [];

    const ss = SpreadsheetApp.openById(SS_ID);
    const userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) return [];

    const dataRange = userSheet.getDataRange().getValues();
    if (dataRange.length <= 1) return [];

    const headers = dataRange[0];
    const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');
    const fullNameCol = headers.indexOf('ชื่อ-นามสกุล');

    const suggestions = [];
    const lowerSearchTerm = searchTerm.toLowerCase();

    for (let i = 1; i < dataRange.length; i++) {
        const row = dataRange[i];
        const regId = String(row[regIdCol]).trim();
        const fullName = String(row[fullNameCol]).trim();
        const displayValue = `${regId} - ${fullName}`;

        if (regId.includes(lowerSearchTerm) || fullName.toLowerCase().includes(lowerSearchTerm)) {
            suggestions.push(displayValue);
        }
        
        if (suggestions.length >= 20) break; // จำกัดจำนวน
    }

    return suggestions;
}

// ---------------------------------------------------------------- //
// 6. CRUD Functions (เพิ่มเติม/แก้ไข)
// ---------------------------------------------------------------- //

// ฟังก์ชันดึงผู้ใช้ทั้งหมด
function getAllUsers() {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const userSheet = ss.getSheetByName(SHEET_USERS);
        if (!userSheet) return [];

        const dataRange = userSheet.getDataRange().getValues();
        if (dataRange.length <= 1) return [];

        const headers = dataRange[0];
        const colMap = {
            regId: headers.indexOf('รหัสเจ้าหน้าที่'),
            fullName: headers.indexOf('ชื่อ-นามสกุล'),
            department: headers.indexOf('หน่วยงาน/สังกัด'),
            position: headers.indexOf('ตำแหน่ง'),
            cardType: headers.indexOf('ประเภทบัตร'),
            expiryDate: headers.indexOf('วันหมดอายุ'),
            profileUrl: headers.indexOf('รูปโปรไฟล์'),
            phone: headers.indexOf('เบอร์โทรศัพท์'),
            driverLicenseNo: headers.indexOf('เลขที่'),
            aviationAreas: headers.indexOf('หมายเลขเขตพื้นที่'),
        };

        const users = [];
        for (let i = 1; i < dataRange.length; i++) {
            const row = dataRange[i];
            
            let expiryDateValue = row[colMap.expiryDate];
            if (expiryDateValue instanceof Date) {
                expiryDateValue = Utilities.formatDate(expiryDateValue, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy'); // Format for input[type=date]
            } else if (typeof expiryDateValue === 'string') {
                // ถ้าเป็น string อยู่แล้ว อาจจะอยู่ในรูปแบบ YYYY-MM-DD
            } else {
                 expiryDateValue = ''; 
            }

            users.push({
                rowNumber: i + 1, // เก็บหมายเลขแถวไว้ใช้ในการแก้ไข/ลบ
                regId: String(row[colMap.regId]).trim(),
                fullName: row[colMap.fullName] || '',
                department: row[colMap.department] || '',
                position: row[colMap.position] || '',
                cardType: row[colMap.cardType] || '',
                expiryDate: expiryDateValue,
                profileUrl: row[colMap.profileUrl] || 'https://placehold.co/150x150/0F1115/E6E9EF/png?text=NO+IMG',
                phone: row[colMap.phone] || '',
                driverLicenseNo: row[colMap.driverLicenseNo] || '',
                aviationAreas: row[colMap.aviationAreas] || '',
            });
        }
        return users;
    } catch(error) {
        Logger.log("Error in getAllUsers: " + error.message);
        return [];
    }
}

// ฟังก์ชันนับจำนวนผู้ใช้ทั้งหมด
function getTotalUserCount() {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const userSheet = ss.getSheetByName(SHEET_USERS);
        if (!userSheet) return 0;

        // นับจำนวนแถวที่มีข้อมูล (ลบแถว Header ออก 1 แถว)
        const rowCount = userSheet.getLastRow();
        return rowCount > 0 ? rowCount - 1 : 0;
    } catch(error) {
        Logger.log("Error in getTotalUserCount: " + error.message);
        return -1;
    }
}

// ฟังก์ชันอัปเดตผู้ใช้งาน
function updateUser(formData) {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const userSheet = ss.getSheetByName(SHEET_USERS);
        
        if (!userSheet) {
            return { success: false, message: 'ไม่พบชีทข้อมูลผู้ใช้งาน' };
        }
        
        const row = parseInt(formData.rowNumber);
        const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
        
        // ตรวจสอบว่ามี RegId ใหม่ซ้ำหรือไม่ (ยกเว้นรหัสเดิม)
        if (formData.newRegId !== formData.originalRegId) {
            const users = userSheet.getRange(2, headers.indexOf('รหัสเจ้าหน้าที่') + 1, userSheet.getLastRow() - 1, 1).getValues();
            const isDuplicate = users.some(userRow => String(userRow[0]).trim() === formData.newRegId);
            if (isDuplicate) {
                return { success: false, message: `รหัสเจ้าหน้าที่ ${formData.newRegId} ถูกใช้แล้ว` };
            }
        }

        // Mapping index ของคอลัมน์
        const colMap = {
            regId: headers.indexOf('รหัสเจ้าหน้าที่') + 1,
            fullName: headers.indexOf('ชื่อ-นามสกุล') + 1,
            department: headers.indexOf('หน่วยงาน/สังกัด') + 1,
            position: headers.indexOf('ตำแหน่ง') + 1,
            cardType: headers.indexOf('ประเภทบัตร') + 1,
            expiryDate: headers.indexOf('วันหมดอายุ') + 1,
            profileUrl: headers.indexOf('รูปโปรไฟล์') + 1,
            phone: headers.indexOf('เบอร์โทรศัพท์') + 1,
            driverLicenseNo: headers.indexOf('เลขที่') + 1,
            aviationAreas: headers.indexOf('หมายเลขเขตพื้นที่') + 1,
        };

        // เตรียมข้อมูลที่จะอัปเดต
        const dataToUpdate = [
            { col: colMap.regId, value: formData.newRegId },
            { col: colMap.profileUrl, value: formData.profileUrl },
            { col: colMap.fullName, value: formData.fullName },
            { col: colMap.department, value: formData.department },
            { col: colMap.position, value: formData.position },
            { col: colMap.phone, value: formData.phone },
            { col: colMap.driverLicenseNo, value: formData.driverLicenseNo },
            { col: colMap.aviationAreas, value: formData.aviationAreas },
            { col: colMap.cardType, value: formData.cardType },
            // แปลงวันที่ YYYY-MM-DD กลับเป็น Date object เพื่อให้ Google Sheet จัดเก็บถูกต้อง
            { col: colMap.expiryDate, value: new Date(formData.expiryDate) }, 
        ];

        // ทำการอัปเดตแต่ละคอลัมน์
        dataToUpdate.forEach(item => {
            if (item.col > 0) {
                userSheet.getRange(row, item.col).setValue(item.value);
            }
        });
        
        // (Optional: Implement logic to update RegId in Access Log if it changed)
        // Note: For simplicity and data integrity, we usually leave old logs as is.

        return { success: true, message: `อัปเดตข้อมูล ${formData.fullName} สำเร็จ` };

    } catch(error) {
        Logger.log("Error in updateUser: " + error.message);
        return { success: false, message: 'เกิดข้อผิดพลาดในการอัปเดต: ' + error.message };
    }
}

// ฟังก์ชันลบผู้ใช้งาน
function deleteUser(regId) {
    try {
        const ss = SpreadsheetApp.openById(SS_ID);
        const userSheet = ss.getSheetByName(SHEET_USERS);
        
        if (!userSheet) {
            return { success: false, message: 'ไม่พบชีทข้อมูลผู้ใช้งาน' };
        }
        
        const dataRange = userSheet.getDataRange().getValues();
        const headers = dataRange[0];
        const regIdCol = headers.indexOf('รหัสเจ้าหน้าที่');

        if (regIdCol === -1) {
             return { success: false, message: 'ไม่พบหัวคอลัมน์ "รหัสเจ้าหน้าที่"' };
        }
        
        // ค้นหาแถวที่จะลบ (เริ่มจากแถวที่ 2)
        for (let i = dataRange.length - 1; i >= 1; i--) {
            const row = dataRange[i];
            if (String(row[regIdCol]).trim() === regId) {
                userSheet.deleteRow(i + 1); // แถวใน sheet เป็น index + 1
                return { success: true, message: `ลบผู้ใช้งานรหัส ${regId} สำเร็จ` };
            }
        }

        return { success: false, message: `ไม่พบผู้ใช้งานรหัส ${regId} ที่ต้องการลบ` };
    } catch(error) {
        Logger.log("Error in deleteUser: " + error.message);
        return { success: false, message: 'เกิดข้อผิดพลาดในการลบ: ' + error.message };
    }
}