/** ===================== CONFIGURATION CONSTANTS ===================== */
const SHEET_MAIN_NAME   = 'Data';
const SHEET_ARCHIVE     = 'Data_Archive';
const SHEET_SETTING     = 'setting';
const SHEET_ADMIN       = 'Admin';
const SHEET_FUEL        = 'Fuel';
const SHEET_INSURANCE   = 'Insurance';
const SHEET_MAINTENANCE = 'Maintenance';
const SHEET_LOG         = 'Log';
const SHEET_VEHICLES    = 'Vehicles'; 
const SHEET_AVAILABILITY = 'Availability'; // 🍓 [BERRY ADD] เพิ่มแค่ชื่อชีตก็พอค่ะ

const HEADER_ROW        = 1;
const MAX_COLS          = 22;
const CACHE_SEC         = 120;
const TZ                = 'Asia/Bangkok';
const SHEET_VEHICLE_STATUS = 'VehicleStatus';

const COLMAP = {
  // คงไว้เหมือนเดิม ไม่ต้องเอาคอลัมน์ของ Availability มาปนค่ะ
  name:        ['ชื่อ-สกุล', 'ชื่อผู้จอง'],
  status:['สถานะ', 'Status'],
  phone:       ['เบอร์โทร', 'เบอร์โทรศัพท์'],
  position:    ['ตำแหน่ง'],
  department:  ['สังกัด', 'หน่วยงาน'],
  email:['email', 'อีเมล'],
  workType:    ['ประเภทงาน', 'jobType'],
  workName:    ['งาน/โครงการ', 'ชื่อโครงการ/งาน', 'projectName', 'project'],
  destination:['สถานที่', 'สถานที่ปลายทาง'],
  carType:     ['ประเภทรถ'],
  vehicle:['เลขทะเบียนรถ', 'ทะเบียนรถ'],  
  requestedVehicle: ['รถที่เลือก'],            
  driver:      ['พนักงานขับรถ'],               
  startDate:   ['วันเริ่มต้น'],
  startTime:['เวลาเริ่มต้น'],
  endDate:     ['วันสิ้นสุด'],
  endTime:     ['เวลาสิ้นสุด'],
  passengers:  ['จำนวนผู้ร่วมเดินทาง'],
  bookingId:   ['Booking ID', 'ID'],
  fileUrl:     ['File', 'ไฟล์แนบ'],
  reason:      ['Reason', 'หมายเหตุ'],
  cancelReason:['CancelReason', 'เหตุผลยกเลิก'],
  vehicleCount:['จำนวนรถที่ต้องการ', 'Vehicle Count', 'จำนวนคัน']
};

const VB_CFG = {
  TZ: TZ,
  DATA_SHEET: SHEET_MAIN_NAME,
  INS_SHEET: SHEET_INSURANCE,
  MAINT_SHEET: SHEET_MAINTENANCE,
  ADVANCE_DAYS: 3
};

// ===================== CORE FUNCTIONS =====================
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('index');
  
  return tpl.evaluate()
    .setTitle('ระบบจองยานพาหนะ มหาวิทยาลัยสวนดุสิต ศูนย์การศึกษาลำปาง')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


function include_(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function _probeHtmlTemplate_(name) {
  try {
    var tpl = HtmlService.createTemplateFromFile(name);
    var now = new Date();
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

    if (name === 'DashboardReport') {
      // ใส่ค่าจำลองให้ตรงกับที่ _defaultGenerateDashboardPdf_ ใช้
      var month = now.getMonth() + 1;
      var year  = now.getFullYear();
      var monthNames = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

      tpl.month       = month;
      tpl.year        = year;
      tpl.monthNames  = monthNames;
      tpl.generatedAt = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');

      tpl.data = {
        totalBookings: 0,
        vehiclesReady: '0/0',
        alerts: 0,
        fuel: 0,
        topDrivers: [],
        topVehicles: []
      };
      tpl.systemName = 'V-Berry Fleet (SelfTest)';
    } else if (name === 'FuelReport') {
      // ใส่ค่าจำลองให้ตรงกับ FuelReport.html
      var year2  = now.getFullYear();
      var month2 = now.getMonth() + 1;

      tpl.title       = 'รายงานสรุปน้ำมัน (SelfTest)';
      tpl.period      = 'เดือน ' + month2 + '/' + (year2 + 543);
      tpl.generatedAt = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');

      tpl.summary = [];      // ไม่มีข้อมูลจริง ใช้แค่ให้ template รันผ่าน
      tpl.detail  = [];
      tpl.totalLiters = 0;
      tpl.totalCost   = 0;
      tpl.systemName  = 'V-Berry Fleet (SelfTest)';
    }

    var html = tpl.evaluate().getContent();
    return 'OK(' + name + ', len=' + html.length + ')';
  } catch (e) {
    // โยน error กลับให้ selfTest log ต่อ
    throw new Error("Template '" + name + "' error: " + e.message);
  }
}

// ===================== DATA MANAGEMENT =====================
function getMainData() {
  return getMainData_();
}

function getMainData_() {
  const CK = 'mainDataCache_v13_BerryFix'; 
  try { cacheDelete_(CK); } catch(e) {}
  
  const cached = cacheGet_(CK);
  if (cached) return { ok:true, data:cached };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME); 
    if (!sh) throw new Error("ไม่พบชีต 'Data' ค่ะ!");
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
    const idx = headerIndex_(headers);

    const lastRow = sh.getLastRow();
    const startRow = 2; 
    
    if (lastRow < startRow) {
        return { 
            ok: true, 
            data: { bookings: [], vehicles: {}, drivers:[], totalBookings: 0, projects:[] } 
        };
    }

    const numRows = (lastRow - startRow) + 1;
    const values = sh.getRange(startRow, 1, numRows, headers.length).getValues();
    const seen = new Set(); 

    const recentBookingsData = values.map((row, i) => {
      const id = String(row[idx.bookingId] || '').trim();
      if (!id || seen.has(id)) return null;
      seen.add(id);

      const startISO = parseDateToISO_(row[idx.startDate]);
      
      return {
        bookingId: id,
        name: String(row[idx.name]||'').trim(),
        status: getStatusKeySafe_(row[idx.status]),
        
        plate: String(row[idx.vehicle]||'').trim(), 
        carName: String(row[idx.requestedVehicle]||'').trim(), 
        driver: String(row[idx.driver]||'').trim(),
        fileUrl: String(row[idx.fileUrl]||'').trim(),
        reason: String(row[idx.reason]||'').trim(),

        phone: formatPhoneNumber_(row[idx.phone]),
        position: String(row[idx.position]||'').trim(),
        org: String(row[idx.department]||'').trim(),
        email: String(row[idx.email]||'').trim(),
        
        // 🍓[BERRY FIX] ดึงค่า ประเภทงาน และ งาน/โครงการ โดยใช้ Fallback ไปหาชื่อเก่าด้วย
        workType: String(row[idx.workType] || row[idx.jobType] || '').trim(),
        workName: String(row[idx.workName] || row[idx.projectName] || row[idx.project] || row[idx.purpose] || '').trim(),
        
        destination: String(row[idx.destination]||'').trim(),
        carType: String(row[idx.carType]||'').trim(),
        
        startDate: startISO,
        startTime: parseTimeSafe_(row[idx.startTime]),
        endDate: parseDateToISO_(row[idx.endDate]) || startISO,
        endTime: parseTimeSafe_(row[idx.endTime]),
        
        passengers: String(row[idx.passengers]||'').trim(),
        cancelReason: String(row[idx.cancelReason]||'').trim(),
        
        dateNum: startISO ? new Date(startISO).getTime() : 0,
        rowNumber: startRow + i
      };
    }).filter(Boolean); 

    const sortedBookings = recentBookingsData.sort((a,b) => {
      const aNum = parseInt(a.bookingId) || 0;
      const bNum = parseInt(b.bookingId) || 0;
      return bNum - aNum; 
    });


// 🍓 [BERRY FIX] Merge Availability Blocks into Calendar (Data UI Safe)
    try {
       const shAvail = ss.getSheetByName('Availability');
       if(shAvail && shAvail.getLastRow() > 1) {
          const avData = shAvail.getDataRange().getValues();
          for(let i=1; i<avData.length; i++) {
             const resType = String(avData[i][0] || '').trim();
             const resId = String(avData[i][1] || '').trim();
             const isDriver = (resType === 'driver');
             
             sortedBookings.push({
                bookingId: 'BLK-' + i,
                status: isDriver ? 'driver_block' : 'vehicle_block',
                name: resId,          // ให้ name ถือค่า resourceId ไปโชว์
                driver: isDriver ? resId : '-', 
                plate: isDriver ? '-' : resId,  
                vehicle: isDriver ? '-' : resId,
                destination: '-',     // ป้องกัน UI Tooltip undefined
                place: '-',
                workType: isDriver ? 'ลาพักงาน' : 'ส่งซ่อมบำรุง',
                workName: avData[i][6] || 'งดใช้งาน', // reason
                project: avData[i][6] || 'งดใช้งาน',
                startDate: parseDateToISO_(avData[i][2]),
                startTime: parseTimeSafe_(avData[i][3]),
                endDate: parseDateToISO_(avData[i][4]) || parseDateToISO_(avData[i][2]),
                endTime: parseTimeSafe_(avData[i][5]),
                dateNum: new Date(parseDateToISO_(avData[i][2])).getTime()
             });
          }
       }
    } catch(ex) {
       Logger.log("Availability Merge Error: " + ex.message);
    }

    const driversRes = getDriversFromAdmin_();
    const drivers = (driversRes.ok && Array.isArray(driversRes.drivers)) ? driversRes.drivers :[];
    
    const platesRes = getAllVehiclePlatesFromSettings();
    const vehicles  = platesRes.ok ? {
      vans: platesRes.vans ||[],
      trucks: platesRes.trucks || [],
      all: platesRes.all ||[]
    } : { vans:[], trucks: [], all:[] };
    
    // ดึงรายชื่อโครงการที่ไม่ซ้ำมาทำ Auto-complete (ใช้อิงจาก workName)
    const projects = Array.from(new Set(sortedBookings
      .map(r => String(r.workName || '').trim())
      .filter(Boolean)));
      
    const payload = {
      ok: true,
      data: {
        bookings: sortedBookings,
        isPartial: false,
        totalBookings: lastRow - 1,
        drivers,
        projects,
        vehicles
      }
    };
    
    cachePut_(CK, payload, 120); 
    return payload;
    
  } catch (e) {
    Logger.log('getMainData_ ERROR: ' + e.stack);
    return { ok:false, error:e.message };
  }
}


function fixEmptyFileColumn() {
  // Keep List:
  // - ทำงานกับชีต Data
  // - วนทั้งคอลัมน์ File แล้ว setValues ครั้งเดียว
  // - return {ok, updated} เหมือนเดิม

  Logger.log('===== fixEmptyFileColumn START =====');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) throw new Error('ไม่พบชีต Data');

  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('INFO: No data rows');
    return { ok: true, updated: 0 };
  }

  // CHANGE: หา index คอลัมน์ File จาก header (กันชีตขยับคอลัมน์)
  var headers = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var idx = headerIndex_(headers);
  if (idx.fileUrl === undefined || idx.fileUrl === -1) {
    throw new Error('ไม่พบคอลัมน์ File/ไฟล์แนบ ในชีต Data');
  }

  var col = idx.fileUrl + 1; // 1-based
  var rng = sh.getRange(2, col, lastRow - 1, 1);
  var vals = rng.getValues();

  var updated = 0;
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] == null ? '' : vals[i][0]).trim();

    // CHANGE: เป้าหมายคือ “ค่าว่าง” ไม่ใช่ "-"
    // - ถ้าเป็น "-" (ของเก่า) ให้ล้างเป็นว่าง
    // - ถ้าว่างอยู่แล้ว ไม่ต้องทำอะไร
    if (v === '-' || v === '–' || v === '—') { // CHANGE
      vals[i][0] = '';                          // CHANGE
      updated++;                                // CHANGE
    }
  }

  if (updated > 0) {
    rng.setValues(vals);
  }

  Logger.log('✅ Normalized File cells: "-" => "" count=' + updated); // CHANGE
  Logger.log('===== fixEmptyFileColumn END =====');
  return { ok: true, updated: updated };
}


function normalizeCarTypeKeyFromUi(raw) {
  var s = String(raw || '').trim().toLowerCase();

  // รองรับทั้งค่าจาก UI (van/truck) และค่าจากชีต ("รถตู้", "รถบรรทุก/รถกระบะบรรทุก")
  if (!s) return '';

  if (s === 'van') return 'van';
  if (s === 'truck') return 'truck';

  if (s.indexOf('รถตู้') > -1 || s.indexOf('ตู้') > -1) return 'van';
  if (s.indexOf('รถบรรทุก') > -1 || s.indexOf('บรรทุก') > -1 || s.indexOf('กระบะ') > -1) return 'truck';

  return '';
}


// ===================== BOOKING MANAGEMENT =====================
// --- Helper Functions (Global Scope for access by other functions if needed) ---

function normalizePhoneText_(raw) {
  let s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  s = s.replace(/[^\d+]/g, '');
  if (s.indexOf('+66') === 0) s = '0' + s.substring(3);
  else if (s.indexOf('66') === 0 && s.length >= 11) s = '0' + s.substring(2);
  s = s.replace(/\D/g, '');
  return s;
}

function toOptText_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s || s === '-' || s === '–' || s === '—') return "";
  return s;
}

function toHHmm_(v) {
  if (v == null || v === '') return '';
  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, 'HH:mm');
  const s = String(v).trim();
  // Try HH:mm:ss or HH:mm
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) return s;
  return String(m[1]).padStart(2, '0') + ':' + m[2];
}

function coerceDateOnly_(v) {
  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    // Return Date object set to midnight in script TZ
    const iso = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.parseDate(iso, tz, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  
  // Try ISO YYYY-MM-DD
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/); // Allow - or /
  if (m) {
    return Utilities.parseDate(`${m[1]}-${m[2]}-${m[3]}`, tz, 'yyyy-MM-dd');
  }
  
  // Try DMY (Thai) dd/mm/yyyy or dd-mm-yyyy
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) {
     let y = parseInt(m[3]);
     // Fix BE year if > 2400
     if(y > 2400) y -= 543; 
     // Format as ISO string then parse back to Date object to ensure correct date
     // Note: Using ISO string YYYY-MM-DD for parsing is safer
     return Utilities.parseDate(`${y}-${m[2]}-${m[1]}`, tz, 'yyyy-MM-dd');
  }
  return '';
}

function createBookingAndBroadcast(payload) {
  const cache = CacheService.getScriptCache();
  const sigBase = String(payload.name) + String(payload.startDate) + String(payload.startTime) + String(payload.workName || payload.projectName || payload.project || '');
  const signature = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sigBase));

  if (cache.get(signature)) {
    return { ok: false, error: "รายการจองนี้กำลังถูกประมวลผลหรือเพิ่งส่งเข้ามาเมื่อครู่ค่ะ" };
  }
  cache.put(signature, "processing", 60);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: "ระบบทำงานหนัก รบกวนลองกดอีกครั้งนะคะ" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error("ไม่พบชีตชื่อ 'Data'");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const idx = headerIndex_(headers);
    const rowData = new Array(headers.length).fill("");
    const setV = (key, val) => { if (idx[key] !== undefined && idx[key] !== -1) rowData[idx[key]] = val; };

    // 🍓 [STEP 3] ใช้ Key ใหม่ workType และ workName
    let workType = String(payload.workType || payload.jobType || "").trim();
    let workName = String(payload.workName || payload.projectName || "").trim();

    const requestedCount = parseInt(payload.vehicleCount, 10) || 1;
    const availabilityCheck = getAvailableVehicles({
      startDate: payload.startDate,
      endDate: payload.endDate || payload.startDate,
      startTime: payload.startTime,
      endTime: payload.endTime,
      carTypes: (payload.carType || '').split(',')
    });

    if (availabilityCheck.ok) {
      const availableCount = availabilityCheck.vehicles.filter(v => v.available).length;
      const maxAllowed = Math.min(5, availableCount);
      if (requestedCount > maxAllowed) {
        return {
          ok: false,
          error: `ขออภัยค่ะ รถว่างไม่พอ (ขณะนี้เหลือว่าง ${maxAllowed} คัน แต่พี่ขอมา ${requestedCount} คัน) รบกวนปรับจำนวนหรือช่วงเวลาใหม่นะคะ`
        };
      }
    }

    const bookingId = reserveNextBookingId();
    const sIso = parseDateToISO_(payload.startDate);
    const eIso = parseDateToISO_(payload.endDate || payload.startDate);

    const buildDateObj = (iso) => {
      if (!iso) return null;
      const p = iso.split('-');
      return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 0, 0, 0);
    };

    const sDateObj = buildDateObj(sIso);
    const eDateObj = buildDateObj(eIso);

    const dayDiff = (eDateObj - sDateObj) / (1000 * 60 * 60 * 24);
    if (dayDiff > 30) {
      return { ok: false, error: `ไม่สามารถจองยาวเกิน 30 วันได้ค่ะ (คุณเลือก ${dayDiff} วัน) กรุณาตรวจสอบวันที่อีกครั้ง` };
    }

    const sTimeStr = parseTimeSafe_(payload.startTime);
    const eTimeStr = parseTimeSafe_(payload.endTime);

    const carTypeMap = { 'van': 'รถตู้', 'truck': 'รถบรรทุก' };
    const typeLabel = (payload.carType || '').split(',')
      .map(t => carTypeMap[String(t).trim().toLowerCase()] || t)
      .filter(Boolean).join(' + ');

    // 🍓 [STEP 6] หยอดข้อมูลลงตัวแปรแถว (ใช้ Key ใหม่)
    setV('bookingId', bookingId);
    setV('status', 'pending');
    setV('name', payload.name);
    setV('phone', payload.phone ? ("'" + String(payload.phone).replace(/\D/g, '')) : "");
    setV('position', payload.position || "-");
    setV('department', payload.department || payload.org || "-");
    setV('email', payload.email || "");

    setV('workType', workType);
    setV('workName', workName);

    setV('destination', payload.place || payload.destination);
    setV('carType', typeLabel);
    setV('vehicleCount', requestedCount);
    setV('startDate', sDateObj);
    setV('startTime', sTimeStr);
    setV('endDate', eDateObj);
    setV('endTime', eTimeStr);
    setV('passengers', payload.passengers);

    if (!payload.fileUrl && payload.fileData && payload.fileName) {
      try {
        var dataStr = String(payload.fileData);
        var base64 = (dataStr.indexOf(',') !== -1) ? dataStr.split(',')[1] : dataStr;
        var bytes = Utilities.base64Decode(base64);
        var mime = payload.fileMime || 'application/octet-stream';
        var blob = Utilities.newBlob(bytes, mime, payload.fileName);

        var up = DriveApp.createFile(blob);
        up.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        payload.fileUrl = up.getUrl();
      } catch (fileErr) {
        Logger.log("Upload file error: " + (fileErr && fileErr.stack ? fileErr.stack : fileErr));
        payload.fileUrl = "";
      }
    }

    setV('fileUrl', payload.fileUrl || payload.file || "");
    setV('reason', payload.reason || "");

    sh.appendRow(rowData);
    const r = sh.getLastRow();
    try {
      const fmtDate = '[$-th-TH-u-ca-buddhist]dd/MM/yyyy';
      if (idx.startDate !== -1) sh.getRange(r, idx.startDate + 1).setNumberFormat(fmtDate);
      if (idx.endDate !== -1) sh.getRange(r, idx.endDate + 1).setNumberFormat(fmtDate);
      if (idx.startTime !== -1) sh.getRange(r, idx.startTime + 1).setNumberFormat('@');
      if (idx.endTime !== -1) sh.getRange(r, idx.endTime + 1).setNumberFormat('@');
      if (idx.phone !== -1) sh.getRange(r, idx.phone + 1).setNumberFormat('@');
    } catch (e) {}

    SpreadsheetApp.flush();

    if (!payload.noTelegram) {
      const notifyPayload = {
        ...payload,
        bookingId: bookingId,
        workType: workType,
        workName: workName,
        status: 'pending',
        vehicleCount: requestedCount,
        carType: typeLabel,
        startDate: sDateObj,
        endDate: eDateObj,
        fileUrl: payload.fileUrl || payload.file || "" 
      };
      sendTelegramNotify(notifyPayload, payload.testMode === true);
    }

    return { ok: true, id: bookingId, message: "บันทึกข้อมูลการจองสำเร็จแล้วค่ะ 🎉" };

  } catch (e) {
    Logger.log("Create Error: " + e.stack);
    return { ok: false, error: "เกิดข้อผิดพลาด: " + e.message };
  } finally {
    lock.releaseLock();
  }
}





/**
 * 🛠️ ฟังก์ชันเสริม: ปรับวันที่เป็น พ.ศ. dd/MM/yyyy
 */
function normalizeDateToDMY_(v) {
  if (!v) return "";
  var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  var dateObj;

  if (v instanceof Date) {
    dateObj = v;
  } else {
    var s = String(v).trim();
    if (s.match(/^\d{4}-\d{2}-\d{2}$/)) {
      var p = s.split('-');
      dateObj = new Date(p[0], p[1] - 1, p[2]);
    } else {
      dateObj = new Date(s);
    }
  }

  if (dateObj && !isNaN(dateObj.getTime())) {
    var year = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'), 10);
    var finalYear = year < 2400 ? year + 543 : year;
    return Utilities.formatDate(dateObj, tz, 'dd/MM/') + finalYear;
  }
  return String(v);
}

function admin_resetBookingIdCounter() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // รอ 15 วินาที
  try {
    Logger.log('--- 🔧 เริ่มต้น Reset Booking ID Counter ---');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) {
      Logger.log(`❌ ไม่พบชีต '${SHEET_MAIN_NAME}'`);
      throw new Error(`ไม่พบชีต '${SHEET_MAIN_NAME}'`);
    }

    // 1. ค้นหา ID สูงสุดในชีต (ใช้ฟังก์ชันเดิมที่เรามี)
    const maxInSheet = detectMaxBookingId_(sh);
    Logger.log(`ℹ️ ID สูงสุดที่พบในชีต Data คือ: ${maxInSheet}`);

    // 2. ดึง ID ที่เก็บไว้ใน Counter
    const props = PropertiesService.getScriptProperties();
    const lastUsed = Number(props.getProperty('COUNTER_BOOKING_ID') || '0');
    Logger.log(`ℹ️ ID ที่เก็บไว้ใน Counter (Properties) คือ: ${lastUsed}`);

    if (maxInSheet < 100) {
        // ป้องกันกรณีชีตว่าง
        Logger.log('⚠️ คำเตือน: ID สูงสุดในชีตน้อยกว่า 100 (อาจจะยังไม่มีข้อมูล) - กำลังยกเลิกการ Reset');
        return { ok: false, error: 'Max ID in sheet is too low, aborting reset.' };
    }

    // 3. เขียนทับ Counter ด้วยค่าที่ถูกต้อง (ที่หาได้จากในชีต)
    props.setProperty('COUNTER_BOOKING_ID', String(maxInSheet));
    
    // 4. ตรวจสอบ
    const finalValue = props.getProperty('COUNTER_BOOKING_ID');
    Logger.log(`✅ Reset สำเร็จ! Counter ถูกตั้งค่าเป็น: ${finalValue}`);
    
    Logger.log('--- 🏁 Reset Booking ID Counter เสร็จสิ้น ---');
    
    return { ok: true, newCounterValue: finalValue };

  } catch(e) {
    Logger.log(`❌ เกิดข้อผิดพลาด أثناء Reset: ${e.stack}`);
    return { ok: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ===== [ANCHOR] Telegram Utilities (new) =====
// โหลด key-value จากชีต setting (คอลัมน์ A:B)
function getSettingMap_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('setting');
  if (!sh) return {};
  var last = sh.getLastRow();
  if (last < 1) return {};

  var vals = sh.getRange(1, 1, last, 2).getValues();
  var map = {};
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0] || '').trim();
    var v = String(vals[i][1] || '').trim();
    if (k) map[k] = v;
  }
  return map;
}

function getTelegramConfig() {
  var map = (typeof getSettingMap === 'function')
    ? getSettingMap()
    : (typeof getSettingMap_ === 'function' ? getSettingMap_() : {});

  var token = String(map['Telegram Bot Token'] || '').trim();
  var chatId = String(map['Telegram Chat Id'] || '').trim();

  if (!token || !chatId) {
    try {
      var p = PropertiesService.getScriptProperties();
      token = token || String(p.getProperty('TELEGRAM_TOKEN') || '').trim();
      chatId = chatId || String(p.getProperty('TELEGRAM_CHAT_ID') || '').trim();
    } catch (_) {}
  }

  return { token: token, chatId: chatId };
}


function postTelegram(text, opts) {
  var cfg = getTelegramConfig(); // ✅ ต้องไม่มี underscore
  if (!cfg.token || !cfg.chatId) {
    return { ok: false, error: 'TELEGRAM not configured (setting: Telegram Bot Token / Telegram Chat Id)' };
  }

  var url = 'https://api.telegram.org/bot' + cfg.token + '/sendMessage';
  var payload = {
    chat_id: cfg.chatId,
    text: String(text || ''),
    parse_mode: (opts && opts.parse_mode) ? String(opts.parse_mode) : 'HTML',
    disable_web_page_preview: !!(opts && (opts.disable_preview || opts.disable_web_page_preview))
  };

  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var body = res.getContentText();
    Logger.log('Telegram RESP ' + code + ': ' + body);

    return { ok: code >= 200 && code < 300, code: code, body: body };
  } catch (e) {
    Logger.log('postTelegram ERROR: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: e.message };
  }
}

function sendTelegramOnce(text, options) {
  options = options || {};

  var dedupeKey = String(options.dedupeKey || '').trim();
  var parseMode = options.parse_mode || options.parseMode || 'HTML';
  var disablePreview = (options.disable_preview !== undefined) ? !!options.disable_preview : true;
  var force = !!options.force; // ✅ support force

  if (!String(text || '').trim()) return { ok: false, error: 'Empty telegram text' };
  if (!dedupeKey) return { ok: false, error: 'Missing dedupeKey' };

  var lock = LockService.getScriptLock();
  var gotLock = false;

  try {
    gotLock = lock.tryLock(15000);
    if (!gotLock) return { ok: false, error: 'Lock timeout (telegram dedupe busy)' };

    var props = PropertiesService.getScriptProperties();
    var sentKey = 'TG_SENT_' + dedupeKey;
    var existing = props.getProperty(sentKey);

    // ✅ skip only when not forcing
    if (existing && !force) {
      return {
        ok: true,
        skipped: true,
        dedupeKey: dedupeKey,
        reason: 'already_sent',
        at: existing
      };
    }

    // ✅ Resolve sender function
    var sender = null;
    if (typeof postTelegram === 'function') sender = postTelegram;
    else if (typeof sendTelegram === 'function') sender = sendTelegram;
    else if (typeof postTelegram_ === 'function') sender = postTelegram_;

    if (!sender) {
      return { ok: false, error: 'Missing telegram sender function: postTelegram / sendTelegram / postTelegram_' };
    }

    var res = sender(text, {
      parse_mode: parseMode,
      disable_preview: disablePreview
    });

    var ok = !!(res && res.ok);

    // ✅ mark as sent only when telegram ok
    if (ok) {
      props.setProperty(sentKey, new Date().toISOString());
    }

    return {
      ok: ok,
      forced: force,
      dedupeKey: dedupeKey,
      response: res || null
    };

  } catch (e) {
    return { ok: false, error: e.message };

  } finally {
    try { if (gotLock) lock.releaseLock(); } catch (_) {}
  }
}

// ANCHOR: ForceThaiBuddhistDateDisplay
function forceThaiBuddhistDateDisplay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) throw new Error("ไม่พบชีตชื่อ 'Data'");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim());
  const idx = headerIndex_(headers);

  if (idx.startDate === undefined || idx.endDate === undefined) {
    throw new Error('ไม่พบคอลัมน์วันเริ่มต้น/วันสิ้นสุด');
  }

  // CHANGE: พยายามตั้ง locale เป็นไทย (ถ้าโดเมนอนุญาต)
  try {
    const cur = ss.getSpreadsheetLocale();
    if (cur !== 'th_TH') {
      ss.setSpreadsheetLocale('th_TH'); // CHANGE
      Logger.log('Spreadsheet locale set to th_TH');
    }
  } catch (e) {
    Logger.log('Locale change blocked by domain: ' + e.message);
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, rows: 0 };

  const fmtBE = '[$-th-TH-u-ca-buddhist]dd/MM/yyyy';
  const fmtFallback = '[$-th-TH]dd/MM/yyyy';

  const rows = lastRow - 1;

  // CHANGE: ฟอร์แมตทั้งคอลัมน์ (ไม่แตะค่า Date)
  try {
    sh.getRange(2, idx.startDate + 1, rows, 1).setNumberFormat(fmtBE);
    sh.getRange(2, idx.endDate + 1, rows, 1).setNumberFormat(fmtBE);
  } catch (e) {
    sh.getRange(2, idx.startDate + 1, rows, 1).setNumberFormat(fmtFallback);
    sh.getRange(2, idx.endDate + 1, rows, 1).setNumberFormat(fmtFallback);
  }

  SpreadsheetApp.flush();
  return { ok: true, rows: rows };
}

// ANCHOR: NormalizeDateOnlyColumnsFull
function normalizeDateOnlyColumns(opt) {
  opt = opt || {};
  var dryRun = opt.dryRun !== false; // default true
  var limit = Number(opt.limit || 10000);
  var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) throw new Error("ไม่พบชีตชื่อ 'Data'");

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var idx = headerIndex_(headers);

  if (idx.startDate === undefined || idx.startDate === -1) throw new Error('ไม่พบ header วันเริ่มต้น');
  if (idx.endDate === undefined || idx.endDate === -1) throw new Error('ไม่พบ header วันสิ้นสุด');

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, dryRun: dryRun, scanned: 0, changed: 0, samples: [] };

  // CHANGE: สแกนทั้งชีตจากแถว 2 ไล่ลงมา (กันพลาดเหมือนเคสที่ sample ไปอยู่ row 2-19)
  var scanN = Math.min(limit, lastRow - 1);
  var startRow = 2;

  var range = sh.getRange(startRow, 1, scanN, sh.getLastColumn());
  var values = range.getValues();

  function isDateObj(v) { return (v instanceof Date) && !isNaN(v.getTime()); }
  function hhmmss(v) { return isDateObj(v) ? Utilities.formatDate(v, tz, 'HH:mm:ss') : ''; }
  function toDateOnly(v) {
    if (!isDateObj(v)) return null;
    var iso = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.parseDate(iso, tz, 'yyyy-MM-dd'); // 00:00
  }

  var changed = 0;
  var samples = [];

  for (var i = 0; i < values.length; i++) {
    var rowIdx = startRow + i;

    var sd = values[i][idx.startDate];
    var ed = values[i][idx.endDate];

    var sdTime = hhmmss(sd);
    var edTime = hhmmss(ed);

    var needSd = (sdTime && sdTime !== '00:00:00');
    var needEd = (edTime && edTime !== '00:00:00');

    if (needSd || needEd) {
      changed++;
      var bid = (idx.bookingId !== undefined && idx.bookingId !== -1) ? values[i][idx.bookingId] : '';

      if (samples.length < 15) {
        samples.push({
          row: rowIdx,
          bookingId: String(bid || ''),
          startWas: isDateObj(sd) ? Utilities.formatDate(sd, tz, 'dd/MM/yyyy HH:mm:ss') : String(sd),
          endWas: isDateObj(ed) ? Utilities.formatDate(ed, tz, 'dd/MM/yyyy HH:mm:ss') : String(ed)
        });
      }

      if (!dryRun) {
        if (needSd) values[i][idx.startDate] = toDateOnly(sd);
        if (needEd) values[i][idx.endDate] = toDateOnly(ed);
      }
    }
  }

  Logger.log('normalizeDateOnlyColumns: dryRun=' + dryRun + ' scanned=' + scanN + ' changed=' + changed);
  Logger.log('samples=' + JSON.stringify(samples));

  if (!dryRun && changed > 0) {
    range.setValues(values);

    // CHANGE: ตั้ง format ให้เป็นวันล้วน (แสดงผลไทย)
    var fmtBE = '[$-th-TH-u-ca-buddhist]dd/MM/yyyy';
    var fmtFallback = '[$-th-TH]dd/MM/yyyy';
    try {
      sh.getRange(2, idx.startDate + 1, lastRow - 1, 1).setNumberFormat(fmtBE);
      sh.getRange(2, idx.endDate + 1, lastRow - 1, 1).setNumberFormat(fmtBE);
    } catch (e) {
      sh.getRange(2, idx.startDate + 1, lastRow - 1, 1).setNumberFormat(fmtFallback);
      sh.getRange(2, idx.endDate + 1, lastRow - 1, 1).setNumberFormat(fmtFallback);
    }

    SpreadsheetApp.flush();
  }

  return { ok: true, dryRun: dryRun, scanned: scanN, changed: changed, samples: samples };
}



// ANCHOR: RunNormalizeDateOnlyAllFull
function runNormalizeDateOnlyAllFull() {
  // CHANGE: รันแก้จริงทั้งชีต (แก้เฉพาะวันเริ่มต้น/วันสิ้นสุดเท่านั้น)
  return normalizeDateOnlyColumns({
    dryRun: false,
    limit: 10000
  });
}



// ===== Thai DateTime: DD/MM/พ.ศ. HH:mm น. OR "-" =====
function formatThaiDateTime_(dateText, timeText) {
  var tz = Session.getScriptTimeZone() || (typeof TZ !== 'undefined' ? TZ : 'Asia/Bangkok');

  function pad2(n) { return String(n).padStart(2, '0'); }
  function isValidHM_(hh, mm) {
    return isFinite(hh) && isFinite(mm) && hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59;
  }

  function parseDate_(d) {
    if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime())) return new Date(d.getTime());

    var s = String(d == null ? '' : d).trim();
    if (!s || s === '-') return null;

    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) {
      var dt0 = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
      return isNaN(dt0.getTime()) ? null : dt0;
    }

    var dm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (dm) {
      var dd = Number(dm[1]), mm = Number(dm[2]), yy = Number(dm[3]);
      if (yy > 2400) yy -= 543;
      var dt1 = new Date(yy, mm - 1, dd);
      return isNaN(dt1.getTime()) ? null : dt1;
    }

    var tmp = new Date(s);
    return isNaN(tmp.getTime()) ? null : tmp;
  }

  var dt = parseDate_(dateText);
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '-';

  var adYear = dt.getFullYear();
  var beYear = (adYear < 2400) ? (adYear + 543) : adYear;

  var datePart = Utilities.formatDate(dt, tz, 'dd/MM/') + beYear;

  function normalizeTimeStr_(s) {
    s = String(s == null ? '' : s).trim();
    if (!s || s === '-') return '';

    s = s.replace(/น\.\s*$/i, '').trim();

    var onlyH = s.match(/^(\d{1,2})$/);
    if (onlyH) {
      var hh0 = Number(onlyH[1]);
      if (!isValidHM_(hh0, 0)) return '';
      return pad2(hh0) + ':00';
    }

    var m = s.match(/^(\d{1,2})[:.](\d{1,2})(?::(\d{1,2}))?$/);
    if (m) {
      var hh = Number(m[1]);
      var mm = Number(m[2]);
      if (!isValidHM_(hh, mm)) return '';
      return pad2(hh) + ':' + pad2(mm);
    }

    if (typeof parseTimeSafe_ === 'function') {
      var r = String(parseTimeSafe_(s) || '').trim();
      if (r) return normalizeTimeStr_(r);
    }
    return '';
  }

  function timeFromSerial_(num) {
    if (typeof num !== 'number' || !isFinite(num)) return '';
    if (num < 0) return '';

    num = num % 1;
    var totalMinutes = Math.round(num * 24 * 60);
    if (!isFinite(totalMinutes)) return '';

    if (totalMinutes >= 24 * 60) totalMinutes = 24 * 60 - 1;
    if (totalMinutes < 0) totalMinutes = 0;

    var hh = Math.floor(totalMinutes / 60);
    var mi = totalMinutes % 60;
    if (!isValidHM_(hh, mi)) return '';
    return pad2(hh) + ':' + pad2(mi);
  }

  function coerceHHmm_(t) {
    if (t == null || t === '' || t === '-') return '';

    if (Object.prototype.toString.call(t) === '[object Date]' && !isNaN(t.getTime())) {
      return Utilities.formatDate(t, tz, 'HH:mm');
    }
    if (typeof t === 'number' && isFinite(t)) return timeFromSerial_(t);
    return normalizeTimeStr_(t);
  }

  var hhmm = '';
  try {
    hhmm = coerceHHmm_(timeText);
    if (hhmm && !/^\d{2}:\d{2}$/.test(hhmm)) hhmm = '';
  } catch (_) {}

  return datePart + ' ' + (hhmm ? (hhmm + ' น.') : '-');
}

// ANCHOR: CoerceDateOnlyInTz
function coerceDateOnlyInTz(v, tzOpt) {
  // CHANGE: centralize date-only coercion for Data sheet write
  var tz = tzOpt || (Session.getScriptTimeZone() || 'Asia/Bangkok');

  if (v == null || v === '') return '';

  // If already Date -> strip time in TZ
  if (v instanceof Date && !isNaN(v.getTime())) {
    var iso0 = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.parseDate(iso0, tz, 'yyyy-MM-dd');
  }

  var s = String(v).trim();
  if (!s) return '';

  // accept "yyyy-MM-dd" or "yyyy-MM-ddTHH:mm:ss" -> take date part
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    var isoDate = iso[1] + '-' + iso[2] + '-' + iso[3];
    return Utilities.parseDate(isoDate, tz, 'yyyy-MM-dd');
  }

  // accept "dd/MM/yyyy" or "dd/MM/BBBB"
  var dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) {
    var dd = Number(dmy[1]), mm = Number(dmy[2]), yy = Number(dmy[3]);
    if (yy >= 2400) yy -= 543;
    var iso2 = String(yy) + '-' + String(mm).padStart(2, '0') + '-' + String(dd).padStart(2, '0');
    return Utilities.parseDate(iso2, tz, 'yyyy-MM-dd');
  }

  // fallback (best-effort) -> normalize by formatting in TZ
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var iso3 = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return Utilities.parseDate(iso3, tz, 'yyyy-MM-dd');
  }

  return '';
}

// ANCHOR: NormalizeDateOnlyByBookingId
function normalizeDateOnlyByBookingId(opt) {
  opt = opt || {};
  var bookingId = String(opt.bookingId || '').trim();
  var dryRun = opt.dryRun !== false; // default true
  var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  if (!bookingId) throw new Error('bookingId is required');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) throw new Error("ไม่พบชีตชื่อ 'Data'");

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var idx = headerIndex_(headers);

  if (idx.bookingId === undefined || idx.bookingId === -1) throw new Error('ไม่พบ header Booking ID');
  if (idx.startDate === undefined || idx.startDate === -1) throw new Error('ไม่พบ header วันเริ่มต้น');
  if (idx.endDate === undefined || idx.endDate === -1) throw new Error('ไม่พบ header วันสิ้นสุด');

  var lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('no data rows');

  var ids = sh.getRange(2, idx.bookingId + 1, lastRow - 1, 1).getValues();
  var targetRow = -1;
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0] || '').trim() === bookingId) { targetRow = i + 2; break; }
  }
  if (targetRow < 2) throw new Error('bookingId not found: ' + bookingId);

  var sd = sh.getRange(targetRow, idx.startDate + 1).getValue();
  var ed = sh.getRange(targetRow, idx.endDate + 1).getValue();

  function isDateObj(v) { return (v instanceof Date) && !isNaN(v.getTime()); }
  function hhmmss(v) { return isDateObj(v) ? Utilities.formatDate(v, tz, 'HH:mm:ss') : ''; }
  function toDateOnly(v) {
    if (!isDateObj(v)) return null;
    var iso = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.parseDate(iso, tz, 'yyyy-MM-dd');
  }

  var sdTime = hhmmss(sd);
  var edTime = hhmmss(ed);

  var needSd = (sdTime && sdTime !== '00:00:00');
  var needEd = (edTime && edTime !== '00:00:00');

  var preview = {
    row: targetRow,
    bookingId: bookingId,
    startWas: isDateObj(sd) ? Utilities.formatDate(sd, tz, 'dd/MM/yyyy HH:mm:ss') : String(sd),
    endWas: isDateObj(ed) ? Utilities.formatDate(ed, tz, 'dd/MM/yyyy HH:mm:ss') : String(ed),
    startNeedFix: needSd,
    endNeedFix: needEd,
    dryRun: dryRun
  };

  Logger.log('normalizeDateOnlyByBookingId preview=' + JSON.stringify(preview));

  if (!dryRun) {
    if (needSd) sh.getRange(targetRow, idx.startDate + 1).setValue(toDateOnly(sd));
    if (needEd) sh.getRange(targetRow, idx.endDate + 1).setValue(toDateOnly(ed));

    // set format ให้เห็นเป็นวันล้วน
    sh.getRange(targetRow, idx.startDate + 1).setNumberFormat('[$-th-TH]dd/MM/yyyy');
    sh.getRange(targetRow, idx.endDate + 1).setNumberFormat('[$-th-TH]dd/MM/yyyy');

    SpreadsheetApp.flush();
  }

  return { ok: true, preview: preview };
}

// ANCHOR: RunNormalize1352
function runNormalize1352() {
  return normalizeDateOnlyByBookingId({
    bookingId: '1352',
    dryRun: false   // เปลี่ยนเป็น true ถ้าจะดู preview ก่อน
  });
}

/**
 * 🍓 [BERRY FULL IMPROVED] ฟังก์ชันสร้างรายงานสรุปงานเดินรถประจำวัน
 * ปรับปรุง: แก้ไขบั๊กเวลากลับหาย และจัดการรูปแบบเวลาปี 1899 ให้สมบูรณ์
 */
function getIntegratedDailyReport(targetDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Data');
    if (!sh) return "❌ ไม่พบชีต Data";

    const tz = Session.getScriptTimeZone() || "Asia/Bangkok";
    const d = (targetDate instanceof Date) ? targetDate : new Date();
    const reportDateISO = Utilities.formatDate(d, tz, "yyyy-MM-dd");

    const TH_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
    const dateHeader = `${Utilities.formatDate(d, tz, "dd")} ${TH_MONTHS[d.getMonth()]} ${d.getFullYear() + 543}`;

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return `📋 <b>รายงานประจำวัน</b>\n📅 ${dateHeader}\n\n🍃 วันนี้ไม่มีงานเดินรถค่ะ`;

    // 1. Mapping Headers (แก้ไข: เพิ่ม endT เพื่อดึงเวลาสิ้นสุด)
    const h = data[0].map(x => String(x).trim());
    const idx = {
      id: h.indexOf('Booking ID'),
      status: h.indexOf('สถานะ'),
      name: h.indexOf('ชื่อ-สกุล'),
      place: h.indexOf('สถานที่'),
      vehicle: h.indexOf('เลขทะเบียนรถ'),
      driver: h.indexOf('พนักงานขับรถ'),
      startD: h.indexOf('วันเริ่มต้น'),
      endD: h.indexOf('วันสิ้นสุด'),
      startT: h.indexOf('เวลาเริ่มต้น'),
      endT: h.indexOf('เวลาสิ้นสุด') // 👈 เพิ่มจุดนี้เพื่อให้ดึงเวลากลับได้จริง
    };

    const groups = { pending: [], approved: [], driver_special_approved: [], rejected: [], cancelled: [] };
    let totalFound = 0;

    // 2. Filter & Group Data
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const sIso = (typeof parseDateToISO_ === 'function') ? parseDateToISO_(r[idx.startD]) : "";
      const eIso = (typeof parseDateToISO_ === 'function') ? parseDateToISO_(r[idx.endD]) : sIso || "";
      
      const finalEndIso = eIso || sIso;
      if (!sIso || reportDateISO < sIso || reportDateISO > finalEndIso) continue;

      let key = (typeof getStatusKeySafe_ === 'function') ? getStatusKeySafe_(r[idx.status]) : 'pending';
      if (key === 'driver_claimed') key = 'pending';
      if (!groups[key]) key = 'pending';

      groups[key].push(r);
      totalFound++;
    }

    if (totalFound === 0) return `📋 <b>รายงานประจำวัน</b>\n📅 ${dateHeader}\n──────────────\n\n🍃 วันนี้ไม่มีภารกิจเดินทางค่ะ`;

    // 3. Build Message Content
    let lines = [`📋 <b>สรุปงานยานพาหนะประจำวัน</b>`, `📅 ${dateHeader}`, '', '📊 <b>สถิติจำนวนงานวันนี้</b>'];

    const statusConfig = [
      { k: 'pending', label: '⏳ รออนุมัติ' },
      { k: 'approved', label: '✅ อนุมัติ' },
      { k: 'driver_special_approved', label: '⚡ อนุมัติกรณีพิเศษ' },
      { k: 'rejected', label: '⛔ ไม่อนุมัติ' },
      { k: 'cancelled', label: '🚫 ยกเลิก' }
    ];

    statusConfig.forEach(item => {
      const count = groups[item.k].length;
      if (count > 0) lines.push(`${item.label} : ${count}`);
    });

    lines.push('━━━━━━━━━━━━━━');

    const activeJobs = [].concat(groups.approved, groups.driver_special_approved);

    if (activeJobs.length > 0) {
      activeJobs.sort((a, b) => String(a[idx.startT] || '').localeCompare(String(b[idx.startT] || '')));

      lines.push('📍 <b>รายละเอียดงานวันนี้:</b>');
      
      // Helper ภายในสำหรับจัดการเวลาให้สะอาด
      const parseT = (val) => {
        if (val instanceof Date) return Utilities.formatDate(val, tz, "HH:mm");
        if (!val || val === '-') return null;
        const m = String(val).match(/(\d{1,2})[:.](\d{2})/);
        return m ? `${m[1].padStart(2, '0')}:${m[2]}` : null;
      };

      activeJobs.forEach(r => {
        const st = (typeof getStatusKeySafe_ === 'function') ? getStatusKeySafe_(r[idx.status]) : 'approved';
        const icon = (st === 'driver_special_approved') ? '⚡' : '🔹';

        const tGo = parseT(r[idx.startT]) || "--:--";
        const tBack = parseT(r[idx.endT]);
        
        // รูปแบบเวลา: "09:00 น. - 16:30 น." หรือ "09:00 น."
        const timeRange = tBack ? `${tGo} น. - ${tBack} น.` : `${tGo} น.`;

        const place = String(r[idx.place] || '-');
        const name = String(r[idx.name] || '-');
        const plate = String(r[idx.vehicle] || '').trim();
        const driver = String(r[idx.driver] || '').trim();

        lines.push(`${icon} ${timeRange} : ${place}`);
        lines.push(`   👤 ${name}`);

        if ((plate && plate !== '-') || (driver && driver !== '-')) {
          const car = (plate && plate !== '-') ? plate : 'รอยืนยันรถ';
          const drv = (driver && driver !== '-') ? driver : 'รอยืนยันคนขับ';
          lines.push(`   🚐 ${car} (${drv})`);
        }
        lines.push('');
      });
    }

    lines.push('🤖 รายงานอัตโนมัติ 05:00 น.');

    // 4. Final Cleanup (Berry Regular Expression)
    return lines.join('\n')
      .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/gi, '')
      .replace(/D\s*น\./gi, '')
      .replace(/[ \t]{2,}/g, ' ')
      .trim();

  } catch (e) {
    console.error(`ERROR: ${e.message}`);
    return "❌ เกิดข้อผิดพลาดในการสร้างรายงาน: " + e.message;
  }
}

function getNoJobTemplate_(dateHeader) {
  return[
    '📋 รายงานประจำวันระบบจองยานพาหนะ',
    '📅 ' + dateHeader,
    '',
    '━━━━━━━━━━━━━━',
    '🍃 วันนี้ไม่มีงานที่ต้องใช้รถ',
    '━━━━━━━━━━━━━━',
    '',
    '🤖 รายงานอัตโนมัติ 05:00'
  ].join('\n');
}

function sanitizeBerryMessage(text) {
  if (!text) return "";

  return String(text)
    .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/gi, '')
    .replace(/D\s*น\./gi, '')
    .replace(/[ \t]+\n/g, '\n') // ลบ space ท้ายบรรทัด
    .replace(/\n{3,}/g, '\n\n') // บีบบรรทัดว่างที่เกิน 2 ให้เหลือ 2
    .trim();
}


/* =========================================
   CORE: TELEGRAM MESSAGE BUILDER (Fixed Time)
   ========================================= */
/* [ANCHOR: Thai DateTime Formatter (Strict BE for Server)] */
function getThaiDateTimeString(dateInput, timeInput) {
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var dateStr = '-';
    var timeStr = '-';

    // 1. แปลงวันที่ -> DD/MM/YYYY (พ.ศ.)
    if (dateInput) {
      // รองรับทั้ง Date Object และ String
      var d = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
      
      // ตรวจสอบว่าเป็นวันที่ที่ถูกต้องหรือไม่
      if (!isNaN(d.getTime())) {
        var y = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
        // Logic: ถ้าปีน้อยกว่า 2400 (เช่น 2026) ให้บวก 543 เพื่อเป็น พ.ศ. (2569)
        var beYear = (y < 2400) ? y + 543 : y; 
        dateStr = Utilities.formatDate(d, tz, 'dd/MM/') + beYear;
      }
    }

    // 2. แปลงเวลา -> HH:mm
    if (timeInput) {
      if (timeInput instanceof Date) {
        timeStr = Utilities.formatDate(timeInput, tz, 'HH:mm');
      } else {
        var s = String(timeInput).trim();
        // พยายามจับรูปแบบ 9:00 หรือ 09:00
        var m = s.match(/(\d{1,2})[:.](\d{2})/);
        if (m) {
          timeStr = ('0' + m[1]).slice(-2) + ':' + m[2];
        }
      }
    }

    // 3. รวมร่าง: "20/01/2569 09:00 น."
    var tPart = (timeStr !== '-' && timeStr !== '') ? (timeStr + ' น.') : '';
    return (dateStr + ' ' + tPart).trim();

  } catch (e) {
    Logger.log("Date convert error: " + e.message);
    return '-';
  }
}

// [TELEGRAM] Status mapping (single source for all Telegram messages)
function getTelegramStatusMeta(statusKey) {
  var k = String(statusKey || 'pending').toLowerCase().trim();
  var map = {
    pending:          { icon:'⏳', th:'รออนุมัติ' },
    approved_full:    { icon:'✅', th:'อนุมัติครบ' },
    approved_partial: { icon:'🟠', th:'อนุมัติบางส่วน' },
    rejected:         { icon:'⛔', th:'ไม่อนุมัติ' },
    cancelled:        { icon:'🚫', th:'ยกเลิก' }
  };
  return map[k] || map.pending;
}

function parseApprovedVehicles(rowObj) {
  rowObj = rowObj || {};
  var raw = String(rowObj['รถที่เลือก'] || rowObj['vehicleSelected'] || '').trim();
  if (!raw) raw = String(rowObj['เลขทะเบียนรถ'] || rowObj['vehicle'] || rowObj['plate'] || '').trim();
  var plates = raw.split(',').map(function(x){ return String(x || '').trim(); }).filter(Boolean);
  var driversRaw = String(rowObj['พนักงานขับรถ'] || rowObj['driver'] || '').trim();
  var drivers = driversRaw.split(',').map(function(x){ return String(x || '').trim(); }).filter(Boolean);
  return { plates: plates, drivers: drivers };
}

function normalizeTelegramStatusKey(statusKey, reqCount, approvedCount) {
  var k = String(statusKey || 'pending').toLowerCase().trim();
  if (k === 'approved') {
    if (approvedCount > 0 && approvedCount < reqCount) return 'approved_partial';
    if (approvedCount >= reqCount && reqCount > 0) return 'approved_full';
    return 'approved_partial';
  }
  if (k === 'pending' || k === 'approved_full' || k === 'approved_partial' || k === 'rejected' || k === 'cancelled') return k;
  return k;
}

function formatPassengers(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s) return 'ไม่ระบุ';
  if (s === '0') return '0 คน';
  var n = Number(s);
  if (isFinite(n) && String(n) === s.replace(/^0+(\d)/, '$1')) return (n + ' คน');
  return s;
}

function normalizePosition(value) {
  var s = String(value == null ? '' : value).trim();
  if (s === 'อาจารย์' || s === 'เจ้าหน้าที่' || s === 'นักศึกษา') return s;
  return '';
}

/* =========================================
   CORE: TELEGRAM MESSAGE BUILDER (Fixed Thai Date BE)
   ========================================= */
/* [ANCHOR: Date Range Generator (Noon Safe)] */
/**
 * สร้าง Array ของวันที่ (String) จากช่วงวัน
 * Logic: ใช้เวลาเที่ยงวัน (12:00) เพื่อป้องกัน Timezone Shift
 */
function buildLocalDateRangeList(startDateObj, endDateObj, tz) {
  tz = tz || (Session.getScriptTimeZone() || 'Asia/Bangkok');
  
  // [BERRY] ปรับ Format เป็น d/M/yyyy + พ.ศ. ให้ตรงกับ UI ส่วนอื่น (หรือแก้ pattern ตามต้องการ)
  function iso(d) { 
    return Utilities.formatDate(d, tz, 'd/M/') + (parseInt(Utilities.formatDate(d, tz, 'yyyy')) + 543);
  }

  if (!(startDateObj instanceof Date) || isNaN(startDateObj.getTime())) return [];

  // [BERRY] Clone วันที่และตั้งเวลาเป็น 12:00:00 (เที่ยงวัน)
  // เพื่อให้ข้าม Timezone Offset ได้ปลอดภัย ไม่ว่าจะ Run Server ที่ไหน
  var s = new Date(startDateObj.getFullYear(), startDateObj.getMonth(), startDateObj.getDate(), 12, 0, 0);
  
  var e;
  if (endDateObj instanceof Date && !isNaN(endDateObj.getTime())) {
    e = new Date(endDateObj.getFullYear(), endDateObj.getMonth(), endDateObj.getDate(), 12, 0, 0);
  } else {
    e = new Date(s.getTime());
  }

  // กรณีวันจบ < วันเริ่ม -> ให้ถือว่าเป็นวันเดียว (One Day Trip)
  if (e.getTime() < s.getTime()) e = new Date(s.getTime());

  var out = [];
  var cur = new Date(s.getTime());
  
  // [BERRY] Safety Limit: ป้องกัน Loop ตาย (จำกัดไม่เกิน 366 วัน)
  var limit = 0;
  var MAX_DAYS = 366; 

  while (cur.getTime() <= e.getTime() && limit < MAX_DAYS) {
    out.push(iso(cur));
    cur.setDate(cur.getDate() + 1); // บวก 1 วัน
    limit++;
  }
  
  return out;
}

// --- Helper: แยกวัตถุประสงค์ และ รายละเอียดโครงการ ---
function getProjectParts_(rawStr, rawDetail) {
  // กรณี 1: มีฟิลด์ detail แยกมาให้ชัดเจน
  if (rawDetail && rawDetail !== '-' && rawDetail !== '') {
    return { purpose: String(rawStr).trim(), detail: String(rawDetail).trim() };
  }

  // กรณี 2: ข้อมูลรวมอยู่ใน project ในรูปแบบ "Purpose: Detail" (Legacy support)
  var s = String(rawStr || '').trim();
  if (!s) return { purpose: '-', detail: '' };
  
  var idx = s.indexOf(':');
  if (idx === -1) {
    return { purpose: s, detail: '' };
  }
  
  var purpose = s.substring(0, idx).trim();
  var detail = s.substring(idx + 1).trim();
  return { purpose: purpose, detail: detail };
}

// --- Helper: จัดการเวลา (Time Block) ---
function buildTimeBlock_(dStart, dEnd, tStart, tEnd, reportDateISO) {
  var tz = 'Asia/Bangkok';
  
  function fmtD(v) {
    if (!v) return null;
    var dObj = (v instanceof Date) ? v : new Date(v);
    if (isNaN(dObj.getTime())) return String(v);
    var dd = Utilities.formatDate(dObj, tz, 'dd');
    var mm = Utilities.formatDate(dObj, tz, 'MM');
    var yyyy = parseInt(Utilities.formatDate(dObj, tz, 'yyyy')) + 543; // บังคับ พ.ศ.
    return dd + '/' + mm + '/' + yyyy;
  }
  
  function fmtShort(v) {
    var s = fmtD(v);
    return s ? s.substring(0, 5) : '';
  }

  function fmtT(v) {
    if (!v) return '00:00';
    var s = String(v).replace('น.', '').trim();
    var p = s.split(':');
    if (p.length === 2) return ('0' + p[0]).slice(-2) + ':' + ('0' + p[1]).slice(-2);
    return s;
  }

  var ds = fmtD(dStart), de = fmtD(dEnd);
  var ts = fmtT(tStart), te = fmtT(tEnd);
  var isoStart = (dStart instanceof Date) ? Utilities.formatDate(dStart, tz, 'yyyy-MM-dd') : String(dStart);
  var isoEnd = (dEnd instanceof Date) ? Utilities.formatDate(dEnd, tz, 'yyyy-MM-dd') : String(dEnd);

  var lines = [];
  if (isoStart === isoEnd) {
    lines.push('🕒 เวลา: ' + ts + '–' + te + ' น.');
    lines.push('📅 วันที่: ' + ds);
  } else {
    lines.push('🕒 งานข้ามวัน');
    lines.push('ไป: ' + ds + ' ' + ts + ' น.');
    lines.push('กลับ: ' + de + ' ' + te + ' น.');
    lines.push('ช่วงงาน: ' + fmtShort(dStart) + ' ' + ts + ' → ' + fmtShort(dEnd) + ' ' + te);
    lines.push('🌙 ค้างคืน');
    if (reportDateISO && isoStart < reportDateISO) {
      lines.push('🔁 ต่อเนื่องจากวันที่ ' + ds);
    }
  }
  return lines.join('\n');
}

/* 🍓 [BERRY FULL FIX] ปรับปรุงการแสดงผล วันที่ (พ.ศ.) และ เวลา (HH:mm) */
function buildBookingStatusMessage(rowObj, statusKey, reasonFromPayload) {
  rowObj = rowObj || {};
  var st = getStatusKeySafe_(statusKey || rowObj.status || rowObj['สถานะ'] || 'pending');
  if (st === 'driver_claimed') st = 'pending';

  // 🍓 BERRY FIX: เช็คว่าเป็นเคสแก้ไขการมอบหมายงานหรือไม่
  var isUpdate = reasonFromPayload && (reasonFromPayload.indexOf('อัปเดต') > -1 || reasonFromPayload.indexOf('เปลี่ยน') > -1);

  var headMap = {
    pending: '🚌 ระบบจองรถ: แจ้งเตือนการจองใหม่',
    approved: isUpdate ? '🔄 ระบบจองรถ: อัปเดตการมอบหมายงาน' : '✅ ระบบจองรถ: อนุมัติรายการ',
    driver_special_approved: isUpdate ? '🔄 ระบบจองรถ: อัปเดตการมอบหมายงาน (ด่วน)' : '⚡ ระบบจองรถ: อนุมัติกรณีพิเศษ',
    rejected: '⛔ ระบบจองรถ: แจ้งผลการพิจารณา',
    cancelled: '🚫 ระบบจองรถ: แจ้งยกเลิก'
  };

  var statusLabelMap = {
    pending: '⏳ รออนุมัติ',
    approved: isUpdate ? '🔄 อัปเดตข้อมูลแล้ว' : '✅ อนุมัติ',
    driver_special_approved: isUpdate ? '🔄 อัปเดตข้อมูลแล้ว' : '⚡ อนุมัติกรณีพิเศษ',
    rejected: '⛔ ไม่อนุมัติ',
    cancelled: '🚫 ยกเลิก'
  };

  function getV(keys) {
    for (var i = 0; i < keys.length; i++) {
      var val = rowObj[keys[i]];
      if (val != null && String(val).trim() !== '' && String(val) !== '-') return val;
    }
    return '';
  }

  // ดึงค่าพื้นฐาน
  var id = getV(['Booking ID', 'id', 'bookingId']);
  var name = getV(['ชื่อ-สกุล', 'name', 'ผู้จอง']);
  var workType = getV(['ประเภทงาน', 'workType', 'jobType']);
  var workName = getV(['งาน/โครงการ', 'ชื่อโครงการ/งาน', 'workName', 'projectName', 'project']);
  var place = getV(['สถานที่', 'destination', 'place']);
  
  // 🍓 BERRY FIX: จัดการ วันที่ และ เวลา ให้เป็นรูปแบบ พ.ศ. และ HH:mm
  var sDateRaw = getV(['วันเริ่มต้น', 'startDate']);
  var sDate = (typeof fmtThaiDateBE_ === 'function') ? fmtThaiDateBE_(sDateRaw) : sDateRaw;
  
  var sTimeRaw = getV(['เวลาเริ่มต้น', 'startTime']);
  var eTimeRaw = getV(['เวลาสิ้นสุด', 'endTime']);
  var sTime = (typeof parseTimeSafe_ === 'function') ? parseTimeSafe_(sTimeRaw) : sTimeRaw;
  var eTime = (typeof parseTimeSafe_ === 'function') ? parseTimeSafe_(eTimeRaw) : eTimeRaw;

  var plate = getV(['เลขทะเบียนรถ', 'vehicle', 'plate']);
  var driver = getV(['พนักงานขับรถ', 'driver']);
  var pax = getV(['จำนวนผู้ร่วมเดินทาง', 'passengers']) || '-';
  var carType = getV(['ประเภทรถ', 'carType', 'vehicleType']) || 'รถยนต์';
  var vCount = getV(['จำนวนรถที่ต้องการ', 'vehicleCount', 'carCount']) || '1';

  var lines = [];
  lines.push('<b>' + (headMap[st] || headMap.pending) + '</b>');
  lines.push('🆔 Booking ID: ' + (id || '-'));
  lines.push('──────────────');
  lines.push('👤 ผู้จอง: ' + name);
  lines.push('🎯 ประเภทงาน: ' + (workType || '-'));
  lines.push('📝 งาน/โครงการ: ' + (workName || '-'));
  lines.push('📍 สถานที่: ' + place);
  lines.push('📅 วันที่: ' + (sDate || '-'));
  lines.push('🕒 เวลา: ' + (sTime || '--:--') + ' - ' + (eTime || '--:--') + ' น.');
  lines.push('🚐 ประเภทรถ: ' + carType);
  lines.push('🚗 จำนวนที่ขอ: ' + vCount + ' คัน');
  lines.push('👥 ผู้ร่วมเดินทาง: ' + pax + ' คน');

  if (st === 'approved' || st === 'driver_special_approved') {
    lines.push('──────────────');
    lines.push('✅ <b>มอบหมายยานพาหนะ</b>');
    lines.push('🚐 ทะเบียน: ' + (plate || 'รอระบุทะเบียน'));
    lines.push('🧑‍✈️ พนักงาน: ' + (driver || 'รอระบุชื่อ'));
  }

  lines.push('──────────────');
  lines.push('🔖 สถานะ: <b>' + (statusLabelMap[st] || statusLabelMap.pending) + '</b>');

  var note = reasonFromPayload || (st === 'cancelled' ? getV(['CancelReason', 'cancelReason']) : getV(['Reason', 'reason']));
  if (note && note !== '-') {
    lines.push('💬 ' + (st === 'rejected' || st === 'cancelled' ? 'เหตุผล: ' : 'หมายเหตุ: ') + note);
  }

  // ล้างอักขระขยะก่อนส่ง
  return lines.join('\n')
    .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/gi, '')
    .replace(/D\s*น\./gi, '')
    .trim();
}


function rowToBookingObject_(rowData, idx, headers) {
  rowData = rowData || [];
  idx = idx || {};
  const get = (k) => (idx[k] != null) ? rowData[idx[k]] : '';
  return {
    name: get('name') || get('fullname') || get('fullName') || get('ชื่อ-สกุล') || '',
    phone: get('phone') || get('tel') || get('เบอร์โทร') || '',
    email: get('email') || '',
    project: get('project') || get('งาน/โครงการ') || '',
    place: get('destination') || get('place') || get('สถานที่') || '',
    carType: get('carType') || get('ประเภทรถ') || '',
    vehicle: get('vehicle') || get('เลขทะเบียนรถ') || '',
    requestedVehicle: get('requestedVehicle') || get('รถที่เลือก') || '',
    driver: get('driver') || get('พนักงานขับรถ') || '',
    startDate: get('startDate') || get('วันเริ่มต้น') || '',
    startTime: get('startTime') || get('เวลาเริ่มต้น') || '',
    endDate: get('endDate') || get('วันสิ้นสุด') || '',
    endTime: get('endTime') || get('เวลาสิ้นสุด') || '',
    passengers: get('passengers') || get('จำนวนผู้ร่วมเดินทาง') || '',
    bookingId: get('bookingId') || get('Booking ID') || '',
    fileUrl: get('fileUrl') || get('File') || '',
    reason: get('reason') || get('Reason') || '',
    cancelReason: get('cancelReason') || get('CancelReason') || '',
    status: get('status') || get('สถานะ') || ''
  };
}



function buildBookingObjectFromRow_(headerRow, rowArr) {
  headerRow = headerRow || [];
  rowArr = rowArr || [];

  var obj = {};
  for (var c = 0; c < headerRow.length; c++) {
    var key = String(headerRow[c] || '').trim();
    if (!key) continue;
    obj[key] = rowArr[c];
  }

  // Add normalized keys that buildBookingStatusMessage_ might use
  // (still derived from header, not fixed index)
  obj.name       = obj['ชื่อ-สกุล'];
  obj.phone      = obj['เบอร์โทร'];
  obj.email      = obj['email'];
  obj.project    = obj['งาน/โครงการ'];
  obj.place      = obj['สถานที่'];
  obj.vehicleType= obj['ประเภทรถ'];
  obj.plate      = obj['เลขทะเบียนรถ'];
  obj.requestedVehicle = obj['รถที่เลือก'];
  obj.driver     = obj['พนักงานขับรถ'];
  obj.startDate  = obj['วันเริ่มต้น'];
  obj.startTime  = obj['เวลาเริ่มต้น'];
  obj.endDate    = obj['วันสิ้นสุด'];
  obj.endTime    = obj['เวลาสิ้นสุด'];
  obj.passengers = obj['จำนวนผู้ร่วมเดินทาง'];
  obj.bookingId  = obj['Booking ID'];
  obj.fileUrl    = obj['File'];
  obj.reason     = obj['Reason'];
  obj.cancelReason = obj['CancelReason'];
  obj.status     = obj['สถานะ'];

  return obj;
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


function fmtThaiDateBE_(d) {
  var tz = (typeof TZ !== 'undefined' && TZ) ? TZ : 'Asia/Bangkok';
  var x = d;

  if (!(x instanceof Date)) {
    var s = String(x || '').trim();
    if (!s || s === '-') return '-';
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) x = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    else x = new Date(s);
  }

  if (!(x instanceof Date) || isNaN(x.getTime())) return '-';

  var ad = x.getFullYear();
  var be = (ad < 2400) ? (ad + 543) : ad;
  return Utilities.formatDate(x, tz, 'dd/MM/') + be;
}

function safeTimeRange_(start, end) {
  // ✅ normalize any time-ish input into "HH:mm"
  function normalizeToHHmm_(v) {
    if (v == null) return '';

    // Date object
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
      var hh = v.getHours();
      var mm = v.getMinutes();
      return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
    }

    var s = String(v || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (!s || s === '-' || s.toLowerCase() === 'null') return '';

    // remove Thai suffix "น."
    s = s.replace(/\s*น\.\s*$/g, '').trim();

    // already "HH:mm"
    var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
    if (m24) {
      var h24 = Number(m24[1]), min24 = Number(m24[2]);
      if (h24 >= 0 && h24 <= 23 && min24 >= 0 && min24 <= 59) {
        return String(h24).padStart(2, '0') + ':' + String(min24).padStart(2, '0');
      }
      return '';
    }

    // "HH:mm:ss"
    var m24s = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m24s) {
      var h24s = Number(m24s[1]), min24s = Number(m24s[2]), sec = Number(m24s[3]);
      if (h24s >= 0 && h24s <= 23 && min24s >= 0 && min24s <= 59 && sec >= 0 && sec <= 59) {
        return String(h24s).padStart(2, '0') + ':' + String(min24s).padStart(2, '0');
      }
      return '';
    }

    // "HH:mm AM/PM"
    var m12 = s.match(/^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/);
    if (m12) {
      var h12 = Number(m12[1]);
      var m12min = Number(m12[2]);
      var ap = String(m12[3]).toUpperCase();
      if (h12 < 1 || h12 > 12 || m12min < 0 || m12min > 59) return '';
      if (ap === 'PM' && h12 < 12) h12 += 12;
      if (ap === 'AM' && h12 === 12) h12 = 0;
      return String(h12).padStart(2, '0') + ':' + String(m12min).padStart(2, '0');
    }

    // fallback: try parse as DateTime string
    var dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      return String(dt.getHours()).padStart(2, '0') + ':' + String(dt.getMinutes()).padStart(2, '0');
    }

    return '';
  }

  var a = normalizeToHHmm_(start);
  var b = normalizeToHHmm_(end);

  if (!a || !b) return '-';

  // ✅ output standard: always one "น." at end of each time
  return a + ' น.-' + b + ' น.';
}

function buildDailySummaryTelegram_(targetDate) {
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var d = targetDate ? new Date(targetDate) : new Date();
    if (targetDate instanceof Date) d = targetDate;

    function toStr(v) { return String(v == null ? '' : v).trim(); }
    function cleanText(v) { return toStr(v).replace(/[\u200B-\u200D\uFEFF]/g, '').trim(); }
    function isEmpty(v) {
      var s = cleanText(v);
      return !s || s === '–' || s === '—' || s.toLowerCase() === 'null' || s.toLowerCase() === 'undefined';
    }
    function safeText(v) { return isEmpty(v) ? 'ไม่ระบุ' : cleanText(v); }

    function fmtDateBE(dateObj) {
      var ad = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'), 10);
      var be = (ad < 2400) ? (ad + 543) : ad;
      return Utilities.formatDate(dateObj, tz, 'dd/MM/') + be;
    }

    function dateISO(dateObj) { return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'); }

    function parseDateAny(v) {
      if (v instanceof Date && !isNaN(v.getTime())) return v;
      var s = cleanText(v);
      if (!s) return null;

      s = s
        .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/g, '')
        .replace(/D\s*น\./g, '')
        .replace(/น\./g, '')
        .replace(/[ ]{2,}/g, ' ')
        .trim();

      var m1 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m1) return new Date(Number(m1[1]), Number(m1[2]) - 1, Number(m1[3]));

      var m2 = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
      if (m2) {
        var dd = Number(m2[1]), mm = Number(m2[2]), yy = Number(m2[3]);
        if (yy > 2400) yy -= 543;
        return new Date(yy, mm - 1, dd);
      }
      return null;
    }

    function parseToHM(v) {
      if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, 'HH:mm');
      var s = cleanText(v);
      if (!s) return '';
      var m = s.match(/(\d{1,2}):(\d{2})/);
      if (!m) return '';
      return String(m[1]).padStart(2, '0') + ':' + String(m[2]);
    }

    function fmtHM(hm) { return hm ? (hm + ' น.') : ''; }

    function normalizePlateList(raw) {
      var s = cleanText(raw);
      if (!s) return [];
      var parts = s.split(',').map(function (x) { return cleanText(x); }).filter(Boolean);

      var out = [];
      parts.forEach(function (p) {
        var mm = p.match(/[ก-ฮ]{1,3}-\d{3,4}/g);
        if (mm && mm.length) mm.forEach(function (x) { out.push(cleanText(x)); });
        else out.push(p);
      });

      var seen = {};
      return out.filter(function (x) {
        x = cleanText(x);
        if (!x) return false;
        if (seen[x]) return false;
        seen[x] = true;
        return true;
      });
    }

    function statusKeyFromThai(v) {
      var s = cleanText(v).toLowerCase();
      if (!s) return 'pending';
      if (s.indexOf('ไม่ผ่าน') > -1 || s.indexOf('ไม่อนุมัติ') > -1 || s === 'rejected') return 'rejected';
      if (s.indexOf('ยกเลิก') > -1 || s === 'cancelled' || s === 'canceled') return 'cancelled';
      if (s.indexOf('อนุมัติ') > -1 || s === 'approved') return 'approved';
      if (s.indexOf('รอ') > -1 || s.indexOf('pending') > -1) return 'pending';
      return 'pending';
    }

    function sortByStartTime(arr) {
      arr.sort(function (a, b) {
        var ta = a.startTime || '';
        var tb = b.startTime || '';
        if (ta === tb) return 0;
        if (!ta) return 1;
        if (!tb) return -1;
        return ta < tb ? -1 : 1;
      });
    }

    // --- read sheet (map header) ---
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Data');
    if (!sh) return null;

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return null;

    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return cleanText(h); });
    var map = {};
    headers.forEach(function (h, i) { if (h) map[h] = i; });

    function getByHeader(row, headerName) {
      var p = map[headerName];
      if (p == null) return '';
      return row[p];
    }

    var targetISO = dateISO(d);
    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var groups = { approved_full: [], approved_partial: [], pending: [], rejected: [], cancelled: [] };

    values.forEach(function (row) {
      var startDateObj = parseDateAny(getByHeader(row, 'วันเริ่มต้น'));
      if (!startDateObj) return;

      var endDateObj = parseDateAny(getByHeader(row, 'วันสิ้นสุด'));
      var isoList = buildLocalDateRangeList(startDateObj, endDateObj, tz);
      if (isoList.indexOf(targetISO) === -1) return;

      var baseKey = statusKeyFromThai(getByHeader(row, 'สถานะ'));

      var it = {
        place: cleanText(getByHeader(row, 'สถานที่')),
        user: cleanText(getByHeader(row, 'ชื่อ-สกุล')),
        position: cleanText(getByHeader(row, 'ตำแหน่ง')),
        pax: getByHeader(row, 'จำนวนผู้ร่วมเดินทาง'),
        purpose: cleanText(getByHeader(row, 'งาน/โครงการ')),
        carType: cleanText(getByHeader(row, 'ประเภทรถ')),
        vCount: cleanText(getByHeader(row, 'จำนวนรถที่ต้องการ')) || '1',
        id: cleanText(getByHeader(row, 'Booking ID')),
        startTime: parseToHM(getByHeader(row, 'เวลาเริ่มต้น')),
        endTime: parseToHM(getByHeader(row, 'เวลาสิ้นสุด')),
        vehicleSelected: cleanText(getByHeader(row, 'รถที่เลือก')),
        plate: cleanText(getByHeader(row, 'เลขทะเบียนรถ')),
        driver: cleanText(getByHeader(row, 'พนักงานขับรถ')),
        fileUrl: cleanText(getByHeader(row, 'File')),
        reason: cleanText(getByHeader(row, 'Reason')),
        cancelReason: cleanText(getByHeader(row, 'CancelReason'))
      };

      var req = parseInt(cleanText(it.vCount), 10);
      req = (isFinite(req) && req > 0) ? req : 1;
      var plates = normalizePlateList(it.vehicleSelected || it.plate || '');
      var approvedCount = plates.length;

      var finalKey = baseKey;
      if (baseKey === 'approved') {
        finalKey = (approvedCount > 0 && approvedCount < req) ? 'approved_partial' : 'approved_full';
      }

      it.reqCountNum = req;
      it.approvedPlates = plates;
      it.approvedCount = approvedCount;

      groups[finalKey].push(it);
    });

    // sort each group by start time
    sortByStartTime(groups.approved_full);
    sortByStartTime(groups.approved_partial);
    sortByStartTime(groups.pending);
    sortByStartTime(groups.rejected);
    sortByStartTime(groups.cancelled);

    function sumRequestedCars(groupsObj) {
      var total = 0;
      ['approved_full', 'approved_partial', 'pending', 'rejected', 'cancelled'].forEach(function (k) {
        (groupsObj[k] || []).forEach(function (it) {
          var n = parseInt(cleanText(it.vCount), 10);
          total += (isFinite(n) && n > 0) ? n : 1;
        });
      });
      return total;
    }

    var totalJobs =
      groups.approved_full.length +
      groups.approved_partial.length +
      groups.pending.length +
      groups.rejected.length +
      groups.cancelled.length;

    var totalCars = sumRequestedCars(groups);

    var lines = [];
    lines.push('📋 สรุปงานยานพาหนะประจำวัน');
    lines.push('📅 ' + fmtDateBE(d));

    if (totalJobs === 0) {
      lines.push('');
      lines.push('🍃 วันนี้ไม่มีรายการจองยานพาหนะค่ะ');
      lines.push('(ระบบเปิดรับจองตามปกติ)');
      lines.push('');
      lines.push('— ออกรายงานอัตโนมัติ 05:00 น. —');
      return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    }

    lines.push('📊 รวมทั้งหมด: ' + totalJobs + ' งาน');
    lines.push('🚗 รวมจำนวนรถที่ขอ: ' + totalCars + ' คัน');
    lines.push(
      '⏳ รอ: ' + groups.pending.length +
      ' | ✅ อนุมัติครบ: ' + groups.approved_full.length +
      ' | 🟠 บางส่วน: ' + groups.approved_partial.length +
      ' | ⛔ ไม่อนุมัติ: ' + groups.rejected.length +
      ' | 🚫 ยกเลิก: ' + groups.cancelled.length
    );
    lines.push('— ออกรายงานอัตโนมัติ 05:00 น. —');

    function fmtTimeRange(it) {
      var a = it.startTime ? fmtHM(it.startTime) : '';
      var b = it.endTime ? fmtHM(it.endTime) : '';
      if (a && b) return a + '-' + b;
      if (a) return a;
      if (b) return b;
      return '';
    }

    function renderGroup(titleLine, arr, kind) {
      if (!arr || arr.length === 0) return;

      lines.push('');
      lines.push(titleLine);

      for (var i = 0; i < arr.length; i++) {
        var it = arr[i];
        var idx = i + 1;

        var tr = fmtTimeRange(it);
        var place = safeText(it.place);
        var user = safeText(it.user);

        var pos = normalizePosition(it.position);
        var userLine = user + (pos ? (' (' + pos + ')') : '');

        var paxLabel = formatPassengers(it.pax);
        var purpose = safeText(it.purpose);
        var carType = safeText(it.carType);

        var approvedLine = '';
        if (kind === 'approved_full') {
          approvedLine = '✅ อนุมัติจริง: ' + it.approvedCount + ' คัน';
        } else if (kind === 'approved_partial') {
          approvedLine = '🟠 อนุมัติจริง ' + it.approvedCount + ' จาก ' + it.reqCountNum + ' คัน';
        } else {
          // pending/rejected/cancelled: show requested count
          approvedLine = '🚗 จำนวนที่ขอ: ' + it.reqCountNum + ' คัน';
        }

        var head = idx + ') ';
        if (tr) head += tr + ' : ' + place;
        else head += place;

        lines.push(head);
        lines.push('   👤 ' + userLine);
        lines.push('   👥 ผู้ร่วมเดินทาง: ' + paxLabel);
        lines.push('   📝 งาน/โครงการ: ' + purpose);
        lines.push('   🚐 ประเภทรถ: ' + carType);
        lines.push('   ' + approvedLine);

        if ((kind === 'approved_full' || kind === 'approved_partial') && it.approvedPlates && it.approvedPlates.length) {
          lines.push('   ✅ รถที่อนุมัติจริง: ' + it.approvedPlates.join(', '));
          if (!isEmpty(it.driver)) lines.push('   🧑‍✈️ คนขับ: ' + cleanText(it.driver));
        } else {
          if (!isEmpty(it.driver)) lines.push('   🧑‍✈️ คนขับ: ' + cleanText(it.driver));
        }

        if (!isEmpty(it.id)) lines.push('   🆔 Booking ID: ' + cleanText(it.id));

        if (kind === 'rejected' || kind === 'cancelled') {
          var rs = (kind === 'cancelled') ? (cleanText(it.cancelReason) || cleanText(it.reason)) : cleanText(it.reason);
          if (!isEmpty(rs)) lines.push('   💬 เหตุผล: ' + rs);
        }

        if (!isEmpty(it.fileUrl)) lines.push('   📎 ไฟล์แนบ: ' + cleanText(it.fileUrl));
      }
    }

    // order: approved_full -> approved_partial -> pending -> rejected -> cancelled
    renderGroup('✅ อนุมัติครบ', groups.approved_full, 'approved_full');
    renderGroup('🟠 อนุมัติบางส่วน', groups.approved_partial, 'approved_partial');
    renderGroup('⏳ รออนุมัติ', groups.pending, 'pending');
    renderGroup('⛔ ไม่อนุมัติ', groups.rejected, 'rejected');
    renderGroup('🚫 ยกเลิก', groups.cancelled, 'cancelled');

    return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  } catch (e) {
    Logger.log('buildDailySummaryTelegram_ error: ' + e);
    return null;
  }
}



function _uniqCsv_(text) {
  var s = String(text || '').trim();
  if (!s) return '';
  var parts = s.split(',').map(function(x){ return String(x || '').trim(); }).filter(function(x){ return !!x; });
  var seen = {};
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var k = parts[i];
    if (!seen[k]) { seen[k] = true; out.push(k); }
  }
  return out.join(', ');
}

function _shouldShowRequestedVehicle_(carType, vCount) {
  var ct = String(carType || '').trim();
  var vc = String(vCount || '').trim();
  if (vc && vc !== '1') return true;
  if (ct.indexOf('+') >= 0) return true;
  return false;
}

function fmtThaiDateBE(d) {
  var tz =
    (typeof VB_CFG !== 'undefined' && VB_CFG && VB_CFG.TZ) ? VB_CFG.TZ :
    ((typeof TZ !== 'undefined' && TZ) ? TZ : 'Asia/Bangkok');

  var x = d;

  if (!(x instanceof Date)) {
    var s = String(x || '').trim();
    if (!s || s === '-') return '-';

    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) x = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    else x = new Date(s);
  }

  if (!(x instanceof Date) || isNaN(x.getTime())) return '-';

  var ad = x.getFullYear();
  var be = (ad < 2400) ? (ad + 543) : ad;
  return Utilities.formatDate(x, tz, 'dd/MM/') + be;
}

// ANCHOR: dailyVehicleSummaryIntegratedFix

function buildDailyReport(targetDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Data');
  if (!sh) return "ไม่พบชีต Data";

  var tz = Session.getScriptTimeZone() || "Asia/Bangkok";
  var d = (targetDate instanceof Date) ? targetDate : new Date();
  var targetISO = Utilities.formatDate(d, tz, "yyyy-MM-dd");

  // 1. แมป Header (ใช้ฟังก์ชันเดิมที่มีในระบบ)
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var idx = headerIndex_(headers);

  // 2. ตัวช่วยอ่านค่าจาก Header แบบปลอดภัย
  function getVal(row, key) {
    var p = idx[key];
    return (p !== undefined && p !== -1) ? String(row[p] || '').trim() : '';
  }

  // 3. กรองและจัดกลุ่มข้อมูล
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var groups = { pending: [], approved: [], rejected: [], cancelled: [] };
  var stats = { totalJobs: 0, totalCars: 0 };

  values.forEach(function(row) {
    var rowDateISO = parseDateToISO_(row[idx.startDate]);
    if (rowDateISO !== targetISO) return;

    var statusRaw = getVal(row, 'status').toLowerCase();
    var stKey = "pending";
    if (statusRaw.indexOf('อนุมัติ') > -1 || statusRaw === 'approved') stKey = 'approved';
    else if (statusRaw.indexOf('ไม่') > -1 || statusRaw === 'rejected') stKey = 'rejected';
    else if (statusRaw.indexOf('ยกเลิก') > -1 || statusRaw === 'cancelled') stKey = 'cancelled';

    // ดึงจำนวนรถจากคอลัมน์ 22
    var vCount = parseInt(getVal(row, 'vehicleCount')) || 1;
    
    // ดึงทะเบียนรถ (ถ้าทะเบียนหลักว่าง ให้ไปดึงจากคอลัมน์ "รถที่เลือก")
    var plate = getVal(row, 'vehicle');
    if (!plate || plate === '-') plate = getVal(row, 'requestedVehicle');

    var item = {
      place: getVal(row, 'destination'),
      user: getVal(row, 'name'),
      pax: getVal(row, 'passengers') || '1',
      purpose: getVal(row, 'project'),
      carType: getVal(row, 'carType'),
      plate: (plate && plate !== '-') ? plate : '',
      driver: getVal(row, 'driver'),
      vCount: vCount,
      id: getVal(row, 'bookingId'),
      reason: getVal(row, 'reason'),
      file: getVal(row, 'fileUrl'),
      cancel: getVal(row, 'cancelReason'),
      startTime: getVal(row, 'startTime') // เก็บไว้ Sort
    };

    groups[stKey].push(item);
    stats.totalJobs++;
    stats.totalCars += vCount;
  });

  // 4. สร้างข้อความรายงาน
  var thYear = parseInt(Utilities.formatDate(d, tz, "yyyy")) + 543;
  var thDate = Utilities.formatDate(d, tz, "dd/MM/") + thYear;

  var lines = [
    '📋 <b>สรุปงานยานพาหนะประจำวัน</b>',
    '📅 <b>' + thDate + '</b>',
    '📊 <b>รวมทั้งหมด: ' + stats.totalJobs + ' งาน</b>',
    '🚗 <b>รวมจำนวนรถที่ขอ: ' + stats.totalCars + ' คัน</b>',
    '🟡 รอ: ' + groups.pending.length + ' | ✅ อนุมัติ: ' + groups.approved.length + ' | ⛔ ไม่ผ่าน: ' + groups.rejected.length + ' | 🚫 ยกเลิก: ' + groups.cancelled.length,
    ''
  ];

  if (groups.approved.length > 0) {
    lines.push('<b>✅ รายการที่อนุมัติแล้ว (' + groups.approved.length + ')</b>', '');
    
    // เรียงตามเวลาไป
    groups.approved.sort(function(a, b) { return a.startTime.localeCompare(b.startTime); });

    groups.approved.forEach(function(it) {
      lines.push('🔹 ' + it.place);
      lines.push('👤 ' + it.user + ' (' + it.pax + ' คน)');
      lines.push('📝 ' + it.purpose);
      lines.push('🚗 จำนวนรถที่ต้องการ: ' + it.vCount + ' คัน');
      
      // ทะเบียนรถ: ห้ามโชว์ (-)
      var plateDisplay = it.plate ? ' (' + it.plate + ')' : '';
      lines.push('🚐 ' + it.carType + plateDisplay);
      
      if (it.driver && it.driver !== '-') lines.push('🧑‍✈️ ' + it.driver);
      lines.push('🆔 ' + it.id);
      
      // ฟิลด์เสริม: ถ้าว่างห้ามแสดง
      if (it.reason && it.reason !== '-') lines.push('💬 ' + it.reason);
      if (it.file && it.file !== '-') lines.push('📎 ' + it.file);
      lines.push('');
    });
  }

  lines.push('— ออกรายงานอัตโนมัติ 05:00 น. —');

  // 5. Final Sanitize: กวาดล้าง Sat และ D น. รอบสุดท้าย
  var finalMsg = lines.join('\n')
    .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/gi, '') // ลบชื่อวัน
    .replace(/D\s*น\./g, '')                           // ลบ D น.
    .replace(/\s{2,}/g, ' ')                           // ลบช่องว่างซ้ำซ้อน
    .trim();

  return finalMsg;
}

/* 🍓 [BERRY FIXED] ฟังก์ชันสำหรับ Trigger ตอนตี 5 (รันแบบห้ามส่งซ้ำ) */
function dailySummaryAt5am() {
  Logger.log("⏰ Trigger: dailySummaryAt5am run");
  sendDailySummaryNotification(false); // false = ห้ามส่งซ้ำเด็ดขาด
}

/* 🍓 [BERRY FIXED] แกนหลักในการส่ง Report (รองรับระบบกันส่งซ้ำ) */
function sendDailySummaryNotification(forceSend) {
  try {
    var tz = Session.getScriptTimeZone() || "Asia/Bangkok";
    var now = new Date();

    Logger.log("===== DAILY REPORT TRIGGER START =====");
    Logger.log("Current Time: " + Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"));

    var msg = getIntegratedDailyReport(now);

    if (!msg || msg.indexOf('ไม่พบชีต Data') > -1) {
      Logger.log("❌ REPORT EMPTY OR ERROR → NOT SENDING");
      return;
    }

    Logger.log("📄 REPORT GENERATED LENGTH: " + msg.length);

    // กุญแจกันส่งซ้ำประจำวัน (1 วันส่งได้ครั้งเดียวต่อ 1 คีย์)
    var dedupeKey = 'DAILY_REPORT_' + Utilities.formatDate(now, tz, 'yyyyMMdd');
    Logger.log("DedupeKey: " + dedupeKey);

    // ส่งเข้า Telegram พร้อมระบบป้องกันการเบิ้ล
    var res = sendTelegramOnce(msg, {
      parse_mode: 'HTML',
      disable_preview: true,
      dedupeKey: dedupeKey,
      force: forceSend === true // 🍓 ถ้าไม่ได้สั่ง force = true จะไม่ยอมส่งซ้ำค่ะ
    });

    Logger.log("📤 TELEGRAM RESULT: " + JSON.stringify(res));
    Logger.log("===== DAILY REPORT TRIGGER END =====");

  } catch (e) {
    Logger.log("❌ ERROR sendDailySummaryNotification: " + e.message);
    Logger.log(e.stack);
  }
}

/* ====== TRIGGERS: create/list/check (05:00 daily) ====== */
function _vbListTriggers_() {
  var out = [];
  var ts = ScriptApp.getProjectTriggers() || [];
  for (var i = 0; i < ts.length; i++) {
    try {
      out.push({
        handler: ts[i].getHandlerFunction && ts[i].getHandlerFunction(),
        type: String(ts[i].getEventType && ts[i].getEventType()),
      });
    } catch (_){}
  }
  return out;
}

function installReminderTriggers() {
  var ts = ScriptApp.getProjectTriggers() || [];
  for (var i = 0; i < ts.length; i++) {
    try {
      var h = ts[i].getHandlerFunction ? ts[i].getHandlerFunction() : '';
      // 🍓 เพิ่มการค้นหาตัวที่ชื่อ sendDailySummaryNotification เพื่อลบทิ้งด้วย
      if (h === 'runAllReminders_' || h === 'dailySummaryAt5am_' || h === 'dailySummaryAt5am' || h === 'sendDailySummaryNotification') {
        ScriptApp.deleteTrigger(ts[i]);
      }
    } catch (_){}
  }
  
  // สร้างใหม่ให้เหลือแค่ 2 ตัวที่ถูกต้อง
  ScriptApp.newTrigger('runAllReminders_').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('dailySummaryAt5am').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
  
  var listed = _vbListTriggers_();
  try { Logger.log('installReminderTriggers: ' + JSON.stringify(listed)); } catch(_){}
  return { ok:true, triggers: listed };
}

/* API ให้เรียกจากหน้าเว็บ/ selfTest */
function apiInstallReminderTriggers(){
  try{
    var res = installReminderTriggers_();
    return _ok_(res);
  }catch(e){
    try { Logger.log('apiInstallReminderTriggers error: ' + e.stack); } catch(_){}
    return _err_(e);
  }
}

function apiListReminderTriggers(){
  try{
    var list = _vbListTriggers_();
    try{
      Logger.log('ReminderTriggers: ' + JSON.stringify(list));
      if (!list || !list.length) Logger.log('ReminderTriggers: EMPTY');
      else {
        for (var i=0;i<list.length;i++){
          Logger.log('Trigger['+i+'] ' + list[i].handler + ' | ' + list[i].type);
        }
      }
    }catch(_){}
    return _ok_({ triggers: list });
  }catch(e){
    try{ Logger.log('apiListReminderTriggers error: ' + e.stack); }catch(_){}
    return _err_(e);
  }
}

/* Dry-run ตรวจสอบและส่งแจ้งเตือน (ใช้ทดสอบจาก UI) */
function apiRunAllReminders(){
  try{
    var res = runAllReminders_();
    return _ok_(res);
  }catch(e){
    try { Logger.log('apiRunAllReminders error: ' + e.stack); } catch(_){}
    return _err_(e);
  }
}


function apiReminderScanDebug(){
  try{
    var leadI = _vb_getSettingNumber('InsuranceReminderDays', VB_CFG && VB_CFG.ADVANCE_DAYS || 3);
    var leadM = _vb_getSettingNumber('MaintenanceReminderDays', VB_CFG && VB_CFG.ADVANCE_DAYS || 3);
    var now = new Date();

    var insRaw = _vb_collectInsuranceDue_();
    var maiRaw = _vb_collectMaintenanceDue_();

    function mapInfo(arr, lead){
      return arr.map(function(it){
        var days = _vb_daysDiff(now, it.due);
        return {
          sheet: it.source,
          plate: it.plate,
          dueISO: Utilities.formatDate(it.due, TZ, 'yyyy-MM-dd\'T\'HH:mm'),
          daysToDue: days,
          withinWindow: (days >= 0 && days <= lead),
          reason: (days < 0 ? 'expired' : (days > lead ? 'not_in_window' : 'ok'))
        };
      });
    }

    var ins = mapInfo(insRaw, leadI);
    var mai = mapInfo(maiRaw, leadM);

    Logger.log('ReminderScanDebug insuranceRaw=' + insRaw.length + ' maintenanceRaw=' + maiRaw.length);
    if (ins.length) Logger.log(JSON.stringify(ins.slice(0,10)));
    if (mai.length) Logger.log(JSON.stringify(mai.slice(0,10)));

    return _ok_({
      leadDays: { insurance: leadI, maintenance: leadM },
      insurance: ins,
      maintenance: mai
    });
  }catch(e){
    Logger.log('apiReminderScanDebug error: ' + e.stack);
    return _err_(e);
  }
}

/** 🛡️ รายงานประกันภัยประจำปี (Server Side - Smart Robust Version) */
function apiGenerateInsuranceAnnualPdf() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Insurance');
    if (!sh) throw new Error("ไม่พบชีต Insurance ในระบบค่ะ");

    const now = new Date();
    const currentYear = now.getFullYear(); // 2026
    const tz = Session.getScriptTimeZone();
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error("ยังไม่มีข้อมูลในฐานข้อมูลประกันภัยค่ะ");

    // 1. Map Header Indices (ดึงตามชื่อคอลัมน์)
    const head = data[0].map(h => String(h||'').trim());
    const idx = {
      plate: head.indexOf('Plate'),
      provider: head.indexOf('Provider'),
      policy: head.indexOf('PolicyNumber'),
      start: head.indexOf('StartDate'),
      end: head.indexOf('EndDate'),
      status: head.indexOf('Status'),
      cost: head.indexOf('Cost'),
      remark: head.indexOf('Remark')
    };

    // เช็คคอลัมน์บังคับ
    if (idx.plate === -1 || idx.start === -1 || idx.end === -1) {
      throw new Error("หัวตารางชีต Insurance ไม่ถูกต้อง (ต้องมี Plate, StartDate, EndDate)");
    }

    let allProcessed = [];
    
    // 2. อ่านข้อมูลทั้งหมดและแปลงเป็น Object
    for (let i = 1; i < data.length; i++) {
      let r = data[i];
      let sISO = parseDateToISO_(r[idx.start]);
      let eISO = parseDateToISO_(r[idx.end]);
      if (!sISO) continue;

      let startDateObj = new Date(sISO + 'T00:00:00');
      let endDateObj = eISO ? new Date(eISO + 'T00:00:00') : startDateObj;
      let costVal = parseFloat(String(r[idx.cost]||'0').replace(/,/g,'')) || 0;

      allProcessed.push({
        plate: String(r[idx.plate] || '-'),
        provider: String(r[idx.provider] || '-'),
        policy: String(r[idx.policy] || '-'),
        startDate: startDateObj,
        endDate: endDateObj,
        costNum: costVal,
        status: String(r[idx.status] || '-'),
        remark: String(r[idx.remark] || '-')
      });
    }

    // 3. Smart Filtering (กรองที่คาบเกี่ยวปีปัจจุบัน)
    let list = allProcessed.filter(item => {
      let startY = item.startDate.getFullYear();
      let endY = item.endDate.getFullYear();
      // เงื่อนไข: เริ่มในปีนี้ OR จบในปีนี้ OR คุ้มครองยาวข้ามปีนี้
      return (startY === currentYear || endY === currentYear || (startY < currentYear && endY > currentYear));
    });

    let periodNote = "";
    // 4. Fallback: ถ้าปีปัจจุบันไม่มีเลย ให้เอา "ข้อมูลล่าสุดทั้งหมด" ในชีตมาโชว์
    if (list.length === 0 && allProcessed.length > 0) {
      list = allProcessed.sort((a,b) => b.startDate - a.startDate).slice(0, 30);
      periodNote = "(แสดงข้อมูลล่าสุดย้อนหลัง เนื่องจากปีปัจจุบันยังไม่มีรายการใหม่)";
    }

    // 5. สรุปสถิติจากรายการที่เลือก
    let stats = { total: 0, active: 0, expired: 0, cost: 0 };
    list.forEach(item => {
      stats.total++;
      stats.cost += item.costNum;
      
      let isStillActive = (item.status.toLowerCase().includes('active') || item.status.includes('คุ้มครอง') || item.endDate >= now);
      if (isStillActive) stats.active++; else stats.expired++;

      // แปลงค่าพร้อมโชว์ใน PDF
      item.startTxt = Utilities.formatDate(item.startDate, tz, "dd/MM/") + (item.startDate.getFullYear() + 543);
      item.endTxt = Utilities.formatDate(item.endDate, tz, "dd/MM/") + (item.endDate.getFullYear() + 543);
      item.costTxt = item.costNum.toLocaleString('th-TH', {minimumFractionDigits: 2});
      item.statusTxt = isStillActive ? "คุ้มครอง" : "หมดอายุ";
    });

    // 6. ส่งข้อมูลให้ Template
    const tpl = HtmlService.createTemplateFromFile('InsuranceReport');
    tpl.list = list.sort((a,b) => a.plate.localeCompare(b.plate));
    tpl.stats = stats;
    tpl.yearBE = currentYear + 543;
    tpl.periodNote = periodNote;
    tpl.generatedAt = Utilities.formatDate(now, tz, "dd/MM/") + (currentYear+543) + " " + Utilities.formatDate(now, tz, "HH:mm") + " น.";
    tpl.systemName = 'ระบบจองยานพาหนะ มหาวิทยาลัยสวนดุสิต';

    // 7. เจน PDF
    const html = tpl.evaluate().getContent();
    const url = __vb_htmlToPdfUrl__("Insurance_Annual_Report_" + (currentYear + 543), html);
    return { ok: true, url: url };

  } catch (e) {
    Logger.log("apiGenerateInsuranceAnnualPdf Error: " + e.stack);
    return { ok: false, error: e.message };
  }
}

/** 🔧 รายงานซ่อมบำรุงรายเดือน (Server Side - Full Robust Version) */
function apiGenerateMaintenanceMonthlyPdf() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Maintenance');
    if (!sh) throw new Error("ไม่พบชีต Maintenance ในระบบค่ะ");

    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const tz = Session.getScriptTimeZone();
    const monthNames = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];

    const data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error("ยังไม่มีข้อมูลการซ่อมบำรุงในฐานข้อมูลค่ะ");

    // 1. Map Header Indices
    const head = data[0].map(h => String(h||'').trim());
    const idx = {
      plate: head.indexOf('Vehicle'),
      date: head.indexOf('Date'),
      type: head.indexOf('Type'),
      cost: head.indexOf('Cost'),
      remark: head.indexOf('Remark'),
      next: head.indexOf('NextDueDate')
    };

    // ตรวจสอบคอลัมน์สำคัญ
    if (idx.plate === -1 || idx.date === -1 || idx.cost === -1) {
      throw new Error("โครงสร้างหัวตารางในชีต Maintenance ไม่ถูกต้อง (ต้องมี Vehicle, Date, Cost)");
    }

    let allProcessed = [];
    let plateStats = {};
    let typeStats = {};

    // 2. Process All Rows
    for (let i = 1; i < data.length; i++) {
      let r = data[i];
      let plate = String(r[idx.plate] || '-').trim();
      if (!plate || plate === '-') continue;

      // ใช้ฟังก์ชัน Robust Date Parsing ที่บอสเพิ่งอัปเกรดไป
      let dISO = parseDateToISO_(r[idx.date]);
      if (!dISO) continue;
      
      let dt = parseDateTime_(dISO, "00:00");
      let costVal = parseFloat(String(r[idx.cost]||'0').replace(/,/g,'')) || 0;
      let serviceType = String(r[idx.type] || 'ทั่วไป');

      allProcessed.push({
        plate: plate,
        type: serviceType,
        dateObj: dt,
        year: dt.getFullYear(),
        month: dt.getMonth(),
        costNum: costVal,
        costTxt: costVal.toLocaleString('th-TH', {minimumFractionDigits: 2}),
        next: r[idx.next] ? parseDateToISO_(r[idx.next]) : null,
        remark: String(r[idx.remark] || '-')
      });
    }

    // 3. Filtering Logic (ดึงเดือนปัจจุบัน ถ้าไม่มีเอาล่าสุด)
    let list = allProcessed.filter(item => item.month === currentMonth && item.year === currentYear);
    let reportPeriod = monthNames[currentMonth] + " " + (currentYear + 543);

    if (list.length === 0) {
      // กรณีเดือนนี้ไม่มีข้อมูล -> ดึง 20 รายการล่าสุด
      list = allProcessed.sort((a,b) => b.dateObj - a.dateObj).slice(0, 20);
      reportPeriod = "ล่าสุด (ย้อนหลัง)";
    } else {
      // เรียงจากวันที่ล่าสุดไปเก่า
      list.sort((a,b) => b.dateObj - a.dateObj);
    }

    // 4. Calculate Summary Statistics
    let stats = { total: 0, cost: 0, topPlate: '-', topType: '-' };
    list.forEach(item => {
      stats.total++;
      stats.cost += item.costNum;
      plateStats[item.plate] = (plateStats[item.plate] || 0) + 1;
      typeStats[item.type] = (typeStats[item.type] || 0) + 1;
    });

    if (stats.total > 0) {
      stats.topPlate = Object.keys(plateStats).reduce((a, b) => plateStats[a] > plateStats[b] ? a : b);
      stats.topType = Object.keys(typeStats).reduce((a, b) => typeStats[a] > typeStats[b] ? a : b);
    }

    // 5. Prepare Template Data
    const tpl = HtmlService.createTemplateFromFile('MaintenanceReport');
    tpl.list = list.map(item => ({
      plate: item.plate,
      type: item.type,
      date: Utilities.formatDate(item.dateObj, tz, "dd/MM/") + (item.year + 543),
      cost: item.costTxt,
      next: item.next ? Utilities.formatDate(new Date(item.next), tz, "dd/MM/") + (new Date(item.next).getFullYear() + 543) : '-',
      remark: item.remark.length > 150 ? item.remark.substring(0, 147) + "..." : item.remark
    }));
    
    tpl.stats = {
      total: stats.total,
      cost: stats.cost.toLocaleString('th-TH', {minimumFractionDigits: 2}),
      topPlate: stats.topPlate,
      topType: stats.topType
    };
    tpl.period = reportPeriod;
    tpl.generatedAt = Utilities.formatDate(now, tz, "dd/MM/") + (currentYear+543) + " " + Utilities.formatDate(now, tz, "HH:mm") + " น.";
    tpl.systemName = 'ระบบจองยานพาหนะ มหาวิทยาลัยสวนดุสิต';

    // 6. Generate PDF URL
    const html = tpl.evaluate().getContent();
    const fileName = "Maint_Report_" + Utilities.formatDate(now, tz, "yyyyMMdd_HHmm");
    const url = __vb_htmlToPdfUrl__(fileName, html);

    return { ok: true, url: url };

  } catch (e) {
    Logger.log("apiGenerateMaintenanceMonthlyPdf Error: " + e.stack);
    return { ok: false, error: e.message };
  }
}

// ===== Triggers: Daily 05:00 Summary (Berry Fixed) =====
function installDailySummaryTrigger_() {
  try {
    var removed = 0;
    var all = ScriptApp.getProjectTriggers();
    for (var i = 0; i < all.length; i++) {
      var t = all[i];
      try {
        var h = t.getHandlerFunction ? t.getHandlerFunction() : '';
        // ลบทั้งตัวเก่า (มี _) และตัวใหม่ทิ้งก่อนสร้างใหม่
        if (h === 'dailySummaryAt5am_' || h === 'dailySummaryAt5am') {
          ScriptApp.deleteTrigger(t);
          removed++;
        }
      } catch (_){}
    }

    // 🍓 สร้างใหม่ ชี้ไปที่ฟังก์ชันหลัก (ไม่มี _)
    ScriptApp.newTrigger('dailySummaryAt5am')
      .timeBased()
      .atHour(5)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone(Session.getScriptTimeZone())
      .create();

    var summary = listDailySummaryTriggers_();
    Logger.log('installDailySummaryTrigger_: removed=' + removed + ' now=' + JSON.stringify(summary));
    return { ok:true, removed: removed, now: summary };
  } catch (e) {
    Logger.log('installDailySummaryTrigger_ error: ' + e.stack);
    return { ok:false, error: e.message };
  }
}


// รายการ Trigger ที่เกี่ยวข้อง (ไว้ debug ใน selfTest)
function listDailySummaryTriggers_() {
  var out = [];
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    var t = all[i];
    var h = (t.getHandlerFunction && t.getHandlerFunction()) || '';
    if (h === 'dailySummaryAt5am' || h === 'dailySummaryAt5am_') {
      out.push({
        handler: h,
        type: String(t.getEventType && t.getEventType()),
        // ไม่มี method มาตรฐานให้ดึงชั่วโมง/นาทีตรง ๆ จาก trigger ที่สร้างแล้ว
        // จึงบันทึก handler ไว้ให้ตรวจด้วยชื่อแทน
      });
    }
  }
  return out;
}

// API เรียกจาก UI (google.script.run) เพื่อสร้าง Trigger และดู error ได้
function apiInstallDailySummaryTrigger() {
  try {
    var res = installDailySummaryTrigger_();
    return res && res.ok ? { ok:true, triggers: listDailySummaryTriggers_() } : res;
  } catch (e) {
    return { ok:false, error: e.message };
  }
}

// API ตรวจสุขภาพ Trigger + Telegram config
function apiDailySummaryHealth() {
  try {
    var cfg = getTelegramConfig();
    var trig = listDailySummaryTriggers_();
    return {
      ok: (!!cfg.token && !!cfg.chatId && trig.length > 0),
      telegram: { token: !!cfg.token, chatId: !!cfg.chatId },
      triggers: trig
    };
  } catch (e) {
    return { ok:false, error: e.message };
  }
}

function getById(bookingId) {
  try {
    const idToFind = String(bookingId || '').trim();
    if (!idToFind) throw new Error('ไม่ระบุ Booking ID');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error("ไม่พบชีต 'Data'");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const idx = headerIndex_(headers);

    if (idx.bookingId === undefined) throw new Error("ไม่พบคอลัมน์ 'Booking ID'");

    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('ไม่พบข้อมูลในชีต');

    // ค้นหาแถวจากล่างขึ้นบน
    const idColValues = sh.getRange(2, idx.bookingId + 1, lastRow - 1, 1).getValues();
    let foundRowIndex = -1;
    for (let i = idColValues.length - 1; i >= 0; i--) {
      if (String(idColValues[i][0]).trim() === idToFind) {
        foundRowIndex = i;
        break;
      }
    }

    if (foundRowIndex === -1) return { ok: false, error: `ไม่พบ ID: ${idToFind}` };

    const sheetRowNumber = foundRowIndex + 2;
    const rowValues = sh.getRange(sheetRowNumber, 1, 1, headers.length).getValues()[0];
    const startISO = parseDateToISO_(rowValues[idx.startDate]);

    const bookingObject = {
      bookingId: idToFind,
      name: String(rowValues[idx.name] || '').trim(),
      status: getStatusKeySafe_(rowValues[idx.status]),
      phone: formatPhoneNumber_(rowValues[idx.phone]),
      position: String(rowValues[idx.position] || '').trim(),
      org: String(rowValues[idx.department] || '').trim(),
      email: String(rowValues[idx.email] || '').trim(),
      
      // 🍓 [BERRY FIX] ใช้ Key มาตรฐานใหม่
      workType: String(rowValues[idx.workType] || rowValues[idx.jobType] || '').trim(),
      workName: String(rowValues[idx.workName] || rowValues[idx.projectName] || rowValues[idx.project] || rowValues[idx.purpose] || '').trim(),
      
      place: String(rowValues[idx.destination] || '').trim(),
      carType: String(rowValues[idx.carType] || '').trim(),
      vehicle: String(rowValues[idx.vehicle] || rowValues[idx.plate] || '').trim(),
      driver: String(rowValues[idx.driver] || '').trim(),
      requestedVehicle: String(rowValues[idx.requestedVehicle] || '').trim(),
      vehicleCount: String(rowValues[idx.vehicleCount] || '1').trim(),

      startDate: startISO,
      startTime: parseTimeSafe_(rowValues[idx.startTime]),
      endDate: parseDateToISO_(rowValues[idx.endDate]) || startISO,
      endTime: parseTimeSafe_(rowValues[idx.endTime]),
      passengers: String(rowValues[idx.passengers] || '').trim(),
      fileUrl: String(rowValues[idx.fileUrl] || '').trim(),
      reason: String(rowValues[idx.reason] || '').trim(),
      cancelReason: String(rowValues[idx.cancelReason] || '').trim(),
      rowNumber: sheetRowNumber
    };

    return { ok: true, data: bookingObject };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}


function normalizeStatusKey_(raw) {
  var s = String(raw || '').trim().toLowerCase();

  // รองรับ label / emoji
  if (s.indexOf('อนุมัติ') >= 0 || s === 'approved') return 'approved';
  if (s.indexOf('ไม่อนุมัติ') >= 0 || s.indexOf('reject') >= 0) return 'rejected';
  if (s.indexOf('ยกเลิก') >= 0 || s.indexOf('cancel') >= 0) return 'cancelled';
  if (s.indexOf('รอ') >= 0 || s.indexOf('pending') >= 0) return 'pending';

  // key ตรง ๆ
  if (s === 'pending' || s === 'approved' || s === 'rejected' || s === 'cancelled') return s;
  return '';
}

function splitCsv_(text) {
  var s = String(text || '').trim();
  if (!s) return [];
  return s.split(',').map(function (x) { return String(x || '').trim(); }).filter(Boolean);
}


function updateBookingStatus(payload) {
  var p = payload || {};
  var bookingId = String(p.bookingId || '').trim();
  var status = String(p.status || '').trim().toLowerCase();

  var dryRun = (p.dryRun === true || String(p.dryRun).toLowerCase() === 'true');
  var testMode = (p.testMode === true || String(p.testMode).toLowerCase() === 'true');

  try {
    if (!bookingId) throw new Error('INVALID_PAYLOAD: missing bookingId');
    if (!status) throw new Error('INVALID_PAYLOAD: missing status');

    var allowed = { pending: 1, approved: 1, rejected: 1, cancelled: 1 };
    if (!allowed[status]) throw new Error('INVALID_PAYLOAD: invalid status => ' + status);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error('SHEET_NOT_FOUND: ' + SHEET_MAIN_NAME);

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) throw new Error('SHEET_EMPTY');

    // ✅ map headers
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
    if (typeof headerIndex_ !== 'function') throw new Error('MISSING_HELPER: headerIndex_');
    var idx = headerIndex_(headers);

    if (typeof idx.bookingId !== 'number') throw new Error('HEADER_MISSING: Booking ID');
    if (typeof idx.status !== 'number') throw new Error('HEADER_MISSING: สถานะ');

    // optional columns
    // vehicle/driver/reason/file/cancelReason
    var hasVehicle = (typeof idx.vehicle === 'number');
    var hasDriver = (typeof idx.driver === 'number');
    var hasReason = (typeof idx.reason === 'number');
    var hasCancelReason = (typeof idx.cancelReason === 'number');
    var hasFile = (typeof idx.file === 'number');

    // ✅ find row by bookingId
    var startRow = 2;
    var values = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
    var foundRowIndex = -1;
    var rowArr = null;

    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      var idVal = String(r[idx.bookingId] || '').trim();
      if (idVal === bookingId) {
        foundRowIndex = startRow + i;
        rowArr = r;
        break;
      }
    }
    if (!rowArr) throw new Error('NOT_FOUND: bookingId=' + bookingId);

    var prevStatus = String(rowArr[idx.status] || '').trim().toLowerCase();

    // ✅ guard: if status unchanged and not testMode/dryRun -> do nothing (avoid duplicates)
    if (!testMode && !dryRun && prevStatus === status) {
      return { ok: true, skipped: true, reason: 'status_unchanged', bookingId: bookingId, status: status };
    }

    // ✅ update sheet first
    rowArr[idx.status] = status;

    // approved: update vehicles/drivers arrays (join as ", ")
    var vehicles = Array.isArray(p.vehicles) ? p.vehicles : [];
    var drivers = Array.isArray(p.drivers) ? p.drivers : [];

    if (status === 'approved') {
      if (hasVehicle) rowArr[idx.vehicle] = vehicles.length ? vehicles.join(', ') : '';
      if (hasDriver) rowArr[idx.driver] = drivers.length ? drivers.join(', ') : '';
      if (hasReason) rowArr[idx.reason] = String(p.reason || '').trim();
    }

    if (status === 'rejected') {
      if (hasReason) rowArr[idx.reason] = String(p.reason || '').trim();
    }

    if (status === 'cancelled') {
      if (hasReason) rowArr[idx.reason] = String(p.reason || '').trim();
      if (hasCancelReason) rowArr[idx.cancelReason] = String(p.cancelReason || p.reason || '').trim();
    }

    // file (optional)
    if (hasFile && p.file != null) {
      rowArr[idx.file] = String(p.file || '').trim();
    }

    // ✅ write back row (atomic row write)
    sh.getRange(foundRowIndex, 1, 1, lastCol).setValues([rowArr]);

    // ✅ build rowObj using header mapping (object)
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = rowArr[c];
    }

    // map into normalized keys for builder
    var norm = {
      bookingId: bookingId,
      status: status,
      name: String(rowObj['ชื่อ-สกุล'] || '').trim(),
      phone: String(rowObj['เบอร์โทร'] || '').trim(),
      email: String(rowObj['email'] || '').trim(),
      project: String(rowObj['งาน/โครงการ'] || '').trim(),
      place: String(rowObj['สถานที่'] || '').trim(),
      carType: String(rowObj['ประเภทรถ'] || '').trim(),
      vehicleCount: String(rowObj['จำนวนรถที่ต้องการ'] || '').trim(),
      passengers: String(rowObj['จำนวนผู้ร่วมเดินทาง'] || '').trim(),
      startDate: rowObj['วันเริ่มต้น'],
      startTime: rowObj['เวลาเริ่มต้น'],
      endDate: rowObj['วันสิ้นสุด'],
      endTime: rowObj['เวลาสิ้นสุด'],
      file: String(rowObj['File'] || '').trim(),
      reason: String(rowObj['Reason'] || '').trim(),
      cancelReason: String(rowObj['CancelReason'] || '').trim(),
      vehicles: vehicles,
      drivers: drivers
    };

    var previewText = buildBookingStatusMessage(norm, status, String(p.reason || '').trim());

    // ✅ PREVIEW ONLY mode
    if (testMode || dryRun) {
      Logger.log('updateBookingStatus: NO TELEGRAM (dryRun=' + dryRun + ', testMode=' + testMode + ')');
      Logger.log('--- TELEGRAM PREVIEW (NOT SENT) ---');
      Logger.log(previewText);
      Logger.log('----------------------------------');

      return {
        ok: true,
        bookingId: bookingId,
        status: status,
        preview: previewText,
        telegramSent: false
      };
    }

    // ✅ real send (dedupe by BookingID + status)
    var dedupeKey = 'BOOKING:' + bookingId + ':' + status;
    var sent = sendTelegramOnce(previewText, {
      parse_mode: 'HTML',
      disable_preview: true,
      dedupeKey: dedupeKey
    });

    return {
      ok: !!(sent && sent.ok),
      bookingId: bookingId,
      status: status,
      preview: previewText,
      telegramSent: !!(sent && sent.ok),
      telegram: sent || null
    };

  } catch (e) {
    Logger.log('updateBookingStatus Error: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}



// ===== helpers used by updateBookingStatus =====
function normalizeStatusKey_(s) {
  var x = String(s || '').toLowerCase().trim();
  x = x.replace(/✅|⏳|❌|🚫|🚗|⚡/g, '').trim();

  if (!x) return 'pending';

  if (x === 'driver_special_approved' || x.indexOf('กรณีพิเศษ') >= 0 || x.indexOf('เร่งด่วน') >= 0) return 'driver_special_approved';
  if (x === 'approved' || x === 'approve') return 'approved';
  if (x === 'pending' || x === 'wait' || x === 'waiting') return 'pending';
  if (x === 'rejected' || x === 'reject' || x === 'notapproved') return 'rejected';
  if (x === 'cancelled' || x === 'canceled' || x === 'cancel') return 'cancelled';

  if (x.indexOf('ไม่อนุมัติ') >= 0 || x.indexOf('ไม่ผ่าน') >= 0) return 'rejected';
  if (x.indexOf('ยกเลิก') >= 0) return 'cancelled';
  if (x.indexOf('อนุมัติ') >= 0) return 'approved';
  if (x.indexOf('รอดำเนินการ') >= 0 || x.indexOf('รอ') === 0) return 'pending';

  if (x === 'driver_claimed' || x.indexOf('รับงาน') >= 0) return 'pending';

  return 'pending';
}

function statusLabelThai_(statusKey) {
  var k = String(statusKey || '').toLowerCase().trim();
  if (k === 'approved') return '✅ อนุมัติ';
  if (k === 'pending') return '⏳ รอดำเนินการ';
  if (k === 'rejected') return '❌ ไม่อนุมัติ';
  if (k === 'cancelled') return '🚫 ยกเลิก';
  return String(statusKey || '');
}

function checkResourcesConflict_(sh, idx, sDate, sTime, eDate, eTime, excludeId, checkVehicles, checkDrivers) {
    if ((!checkVehicles || !checkVehicles.length) && (!checkDrivers || !checkDrivers.length)) return { hasConflict: false };
    
    // [BERRY FIX] ตรวจสอบ Availability Engine (การลางาน/ซ่อมบำรุง) เป็นด่านแรก
    if (checkDrivers && checkDrivers.length > 0) {
        for (let d of checkDrivers) {
           let avail = checkDriverAvailability(d, sDate, sTime, eDate, eTime);
           if (avail.conflict) return { hasConflict: true, message: `คนขับ ${d} ไม่พร้อมใช้งาน: ${avail.reason}` };
        }
    }
    if (checkVehicles && checkVehicles.length > 0) {
        for (let v of checkVehicles) {
           let avail = checkVehicleAvailability(v, sDate, sTime, eDate, eTime);
           if (avail.conflict) return { hasConflict: true, message: `รถ ${v} ไม่พร้อมใช้งาน: ${avail.reason}` };
        }
    }

    const data = sh.getDataRange().getValues();
    const reqStart = parseDateTime_(sDate, sTime);
    const reqEnd = parseDateTime_(eDate, eTime);
    
    if (!reqStart || !reqEnd) return { hasConflict: false };

    for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const rowId = String(row[idx.bookingId] || '').trim();
        
        // ข้ามตัวเอง และ ข้ามรายการที่ยังไม่ Approved
        if (rowId === String(excludeId)) continue;
        const status = getStatusKeySafe_(row[idx.status]);
        if (status !== 'approved') continue;

        // เช็คเวลาชนกัน
        const rStartISO = parseDateToISO_(row[idx.startDate]);
        const rStartTime = parseTimeSafe_(row[idx.startTime]);
        const rEndISO = parseDateToISO_(row[idx.endDate]) || rStartISO;
        const rEndTime = parseTimeSafe_(row[idx.endTime]);

        const exStart = parseDateTime_(rStartISO, rStartTime);
        const exEnd = parseDateTime_(rEndISO, rEndTime);

        if (!exStart || !exEnd) continue;
        
        // Logic Overlap: (StartA < EndB) && (EndA > StartB)
        const isOverlapping = (reqStart < exEnd && reqEnd > exStart);
        
        if (isOverlapping) {
            // เช็คทะเบียนรถซ้ำ
            const rowVehicles = String(row[idx.vehicle] || '').split(',').map(v => v.trim());
            const vehicleConflict = checkVehicles.find(v => rowVehicles.includes(v));
            if (vehicleConflict) {
                return { hasConflict: true, message: `รถ ${vehicleConflict} ติดงานอื่นในช่วงเวลานี้ (Booking ID: ${rowId})` };
            }

            // เช็คคนขับซ้ำ (ถ้ามีคนขับส่งมาเช็ค)
            if (checkDrivers && checkDrivers.length > 0) {
                const rowDrivers = String(row[idx.driver] || '').split(',').map(d => d.trim());
                const driverConflict = checkDrivers.find(d => rowDrivers.includes(d));
                if (driverConflict) {
                    return { hasConflict: true, message: `คนขับ ${driverConflict} ติดงานอื่นในช่วงเวลานี้ (Booking ID: ${rowId})` };
                }
            }
        }
    }
    return { hasConflict: false };
}

// ===================== UTILITY FUNCTIONS =====================
function normalizeStatus_(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (!s) return 'pending';
  if (['driver_special_approved', 'อนุมัติกรณีพิเศษ', 'อนุมัติเร่งด่วน', 'กรณีพิเศษ'].includes(s)) return 'driver_special_approved';
  if (['driver_claimed', 'คนขับรับงานแล้ว', 'รับงาน', 'พนักงานรับงานแล้ว'].includes(s)) return 'pending';
  if (['pending', 'รออนุมัติ', 'รอดำเนินการ', 'กำลังรอ'].includes(s)) return 'pending';
  if (['approved', 'อนุมัติ', 'ยืนยัน', 'ผ่าน'].includes(s)) return 'approved';
  if (['rejected', 'ไม่อนุมัติ', 'ปฏิเสธ', 'ไม่ผ่าน'].includes(s)) return 'rejected';
  if (['cancelled', 'ยกเลิก', 'ผู้จองยกเลิก', 'ยกเลิกการจอง'].includes(s)) return 'cancelled';
  return s;
}

/**
 * แปลงวันที่ทุกรูปแบบให้เป็น ISO AD (ค.ศ.) yyyy-MM-dd
 * [BERRY FIXED] แก้ปัญหาลบปีจนถอยไปยุคอยุธยา (1654)
 */
function parseDateToISO_(v) {
  try {
    var tz = Session.getScriptTimeZone(); 
    if (v == null || v === "") return null;

    var year, month, day;
    var now = new Date();
    var currentAD = now.getFullYear();

    // 1. แยกตัวเลข ปี เดือน วัน ออกมา
    if (v instanceof Date && !isNaN(v.getTime())) {
      year = v.getFullYear();
      month = v.getMonth() + 1;
      day = v.getDate();
    } else if (typeof v === 'number' && isFinite(v)) {
      var d = new Date(Math.round((v - 25569) * 86400 * 1000));
      year = d.getFullYear(); month = d.getMonth() + 1; day = d.getDate();
    } else {
      var s = String(v).trim().replace(/[๐-๙]/g, function(d) { return '๐๑๒๓๔๕๖๗๘๙'.indexOf(d); });
      var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      var mSlash = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
      var mThai = s.match(/(\d{1,2})\s+([\u0E00-\u0E7F\.]+)\s+(\d{4})/);

      if (mIso) {
        year = parseInt(mIso[1]); month = parseInt(mIso[2]); day = parseInt(mIso[3]);
      } else if (mSlash) {
        day = parseInt(mSlash[1]); month = parseInt(mSlash[2]); year = parseInt(mSlash[3]);
      } else if (mThai) {
        var TH_MONTHS = {'มกราคม':1,'กุมภาพันธ์':2,'มีนาคม':3,'เมษายน':4,'พฤษภาคม':5,'มิถุนายน':6,'กรกฎาคม':7,'สิงหาคม':8,'กันยายน':9,'ตุลาคม':10,'พฤศจิกายน':11,'ธันวาคม':12,'ม.ค.':1,'ก.พ.':2,'มี.ค.':3,'เม.ย.':4,'พ.ค.':5,'มิ.ย.':6,'ก.ค.':7,'ส.ค.':8,'ก.ย.':9,'ต.ค.':10,'พ.ย.':11,'ธ.ค.':12};
        day = parseInt(mThai[1]); month = TH_MONTHS[mThai[2]]; year = parseInt(mThai[3]);
      } else {
        var d2 = new Date(s);
        if (isNaN(d2.getTime())) return null;
        year = d2.getFullYear(); month = d2.getMonth() + 1; day = d2.getDate();
      }
    }

    // 🍓 [BERRY'S NEW SMART GUARD]
    // Step 1: ถ้าเป็น พ.ศ. (2400+) ให้ลบ 543 ก่อน
    if (year > 2400) year -= 543;

    // Step 2: ตรวจสอบความสมเหตุสมผล (Reasonable Year Range)
    // ระบบจองรถควรจะอยู่ในช่วงปี ค.ศ. ปัจจุบัน (บวกลบไม่เกิน 5 ปี)
    // ถ้าใครส่งปี 4369 หรือปีที่ลบแล้วเหลือ 1654 มา -> ให้ปัดมาเป็นปีปัจจุบัน (AD) ทันที
    if (year > (currentAD + 10) || year < (currentAD - 5)) {
      year = currentAD; 
    }
    
    if (!month || !day) return null;
    return year + "-" + String(month).padStart(2, '0') + "-" + String(day).padStart(2, '0');
  } catch (e) {
    return null;
  }
}

/**
 * รวมวันที่ ISO และเวลา String ให้เป็น Date Object (Local Time)
 */
function parseDateTime_(dateISO, timeStr) {
  try {
    if (!dateISO) return null;
    
    // 🍓 [BERRY FIX] ล้างค่าปีอีกครั้งก่อนสร้าง Object เพื่อความชัวร์
    const cleanISO = parseDateToISO_(dateISO); 
    if (!cleanISO) return null;

    const dParts = cleanISO.split('-');
    const year = parseInt(dParts[0]);
    const month = parseInt(dParts[1]) - 1; 
    const day = parseInt(dParts[2]);

    // จัดการเวลา (รองรับทั้ง 08:00 และ 08:00 น.)
    const tParts = String(timeStr || '00:00')
      .replace(/น\./g, '')
      .trim()
      .split(':');
      
    const hours = parseInt(tParts[0] || 0);
    const minutes = parseInt(tParts[1] || 0);

    // สร้าง Date Object แบบพารามิเตอร์เพื่อเลี่ยงปัญหา UTC offset
    const d = new Date(year, month, day, hours, minutes, 0, 0);
    
    if (isNaN(d.getTime())) return null;
    return d;
  } catch (e) {
    Logger.log(`parseDateTime_ Error: ${e.message}`);
    return null;
  }
}

/* [ANCHOR: Public Driver List for Dropdowns] */
function getDriverList() {
  try {
    // 1. ดึงข้อมูล Drivers ทั้งหมด (ใช้ฟังก์ชันเดิมของพี่)
    const driversRes = getDriversFromAdmin_(); 
    const driversRaw = driversRes.ok ? driversRes.drivers : [];

    // 2. อ่านสถานะ Active/Inactive จาก Setting
    const dStatusKv = readSettingKV_('DriverStatus'); 
    const dStatusMap = parseBoolMap_(dStatusKv.val);

    // 3. กรองเอาเฉพาะคนที่ Active (สถานะเป็น true)
    const activeDrivers = driversRaw.filter(d => {
       // ถ้าไม่มีค่าใน Setting ให้ถือว่า Active (true) โดย Default
       return dStatusMap.hasOwnProperty(d.name) ? dStatusMap[d.name] : true;
    });

    // 4. ส่งกลับเฉพาะ "ชื่อ" เป็น Array List
    return activeDrivers.map(d => d.name);

  } catch (e) {
    Logger.log("getDriverList Error: " + e.toString());
    // Fallback: ถ้า Error ให้ส่งค่าพื้นฐานไปก่อน
    return ["Admin", "Tester"]; 
  }
}

/* --- ADMIN PANEL APIS (Enhanced by Berry) --- */
function apiGetAdminPanelData() {
  try {
    // 1. ดึงข้อมูล Drivers (Master List)
    const driversRes = getDriversFromAdmin_();
    const driversRaw = driversRes.ok ? driversRes.drivers : [];

    // 2. ดึงข้อมูล Vehicles (Master List)
    const vehiclesRes = getAllVehiclePlatesFromSettings();
    const vehiclesRaw = vehiclesRes.ok ? vehiclesRes.all : [];
    
    // 3. อ่านสถานะจาก Setting (Real-time override)
    // DriverStatus: "Somchai:true,Somsri:false"
    const dStatusKv = readSettingKV_('DriverStatus'); 
    const dStatusMap = parseBoolMap_(dStatusKv.val);

    // VehicleAvailability: "ฮค-1234:true,กข-9999:false"
    const vStatusKv = readSettingKV_('VehicleAvailability');
    const vStatusMap = parseBoolMap_(vStatusKv.val);

    // 4. ผสมข้อมูล (Merge)
    const drivers = driversRaw.map(d => {
      // ถ้ามีค่าใน Map ให้ใช้ค่านั้น, ถ้าไม่มีให้ Default เป็น true (Active)
      const isActive = dStatusMap.hasOwnProperty(d.name) ? dStatusMap[d.name] : true;
      return {
        name: d.name,
        username: d.username,
        role: d.role,
        active: isActive
      };
    });

    const vehicles = vehiclesRaw.map(v => {
      const isActive = vStatusMap.hasOwnProperty(v.plate) ? vStatusMap[v.plate] : true;
      return {
        plate: v.plate,
        name: v.name,
        type: v.type,
        active: isActive
      };
    });

    return { ok: true, drivers: drivers, vehicles: vehicles };
  } catch (e) {
    Logger.log('apiGetAdminPanelData Error: ' + e.stack);
    return { ok: false, error: e.message };
  }
}

function apiToggleDriverStatus(payload) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(5000)) return {ok:false, error:'System busy'};
  
  try {
    const { name, active } = payload;
    if(!name) throw new Error('Missing name');

    // อ่านค่าเดิม
    const kv = readSettingKV_('DriverStatus');
    const map = parseBoolMap_(kv.val);
    
    // อัปเดต
    map[name] = (active === true || active === 'true');
    
    // แปลงกลับเป็น String
    const newStr = Object.keys(map).map(k => `${k}:${map[k]}`).join(',');
    _saveSettingValue_('DriverStatus', newStr);
    
    return { ok: true, name: name, active: map[name] };
  } catch(e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

function apiToggleVehicleStatus(payload) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(5000)) return {ok:false, error:'System busy'};

  try {
    const { plate, active } = payload;
    if(!plate) throw new Error('Missing plate');

    const kv = readSettingKV_('VehicleAvailability');
    const map = parseBoolMap_(kv.val);
    
    map[plate] = (active === true || active === 'true');
    
    const newStr = Object.keys(map).map(k => `${k}:${map[k]}`).join(',');
    _saveSettingValue_('VehicleAvailability', newStr);
    
    return { ok: true, plate: plate, active: map[plate] };
  } catch(e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}


// Helper: บันทึกค่าลง Setting (ใช้คู่กับ readSettingKV_)
function _saveSettingValue_(key, value) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('setting');
  if(!sh) throw new Error("No setting sheet");
  
  // หา Row เดิม
  const data = sh.getRange("A:A").getValues();
  let row = -1;
  for(let i=0; i<data.length; i++){
    if(String(data[i][0]).trim() === key) {
      row = i + 1;
      break;
    }
  }
  
  if(row > 0) {
    sh.getRange(row, 2).setValue(value); // Update Col B
  } else {
    sh.appendRow([key, value]); // Create New
  }
}

function keepPhone(v) {
  var s = (v == null) ? '' : String(v).trim();
  if (!s) return '';
  // เก็บเฉพาะตัวเลขและเครื่องหมาย +
  s = s.replace(/[^\d+]/g, '');
  return s;
}

function normalizeVehicleTypeLabel_(raw) {
  var s = (raw == null) ? '' : String(raw).trim();
  if (!s) return '';

  var low = s.toLowerCase();

  if (low === 'van' || low.indexOf('van') > -1 || s.indexOf('ตู้') > -1) return 'รถตู้';
  if (low === 'truck' || low.indexOf('truck') > -1 || s.indexOf('กระบะ') > -1 || s.indexOf('บรรทุก') > -1) return 'รถบรรทุก';

  s = s.replace('รถกระบะ/บรรทุก', 'รถบรรทุก');
  s = s.replace(/, /g, ' + ');
  s = s.replace(/\|/g, ' + ');
  return s;
}

function getBookingRowById_(bookingId) {
  var id = String(bookingId || '').trim();
  if (!id) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var bookingCol = (typeof COL !== 'undefined' && COL.BOOKING_ID) ? COL.BOOKING_ID : 18;

  var ids = sh.getRange(2, bookingCol, lastRow - 1, 1).getValues();
  var foundRow = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === id) { foundRow = i + 2; break; }
  }
  if (foundRow < 0) return null;

  var rowVals = sh.getRange(foundRow, 1, 1, sh.getLastColumn()).getValues()[0];

  function V(colIndex1Based) {
    return (colIndex1Based && colIndex1Based >= 1) ? rowVals[colIndex1Based - 1] : '';
  }

  var obj = {
    name: V(COL.NAME || 1),
    status: V(COL.STATUS || 2),
    phone: V(COL.PHONE || 3),
    position: V(COL.POSITION || 4),
    org: V(COL.ORG || 5),
    email: V(COL.EMAIL || 6),
    project: V(COL.PROJECT || 7),
    destination: V(COL.DESTINATION || 8),
    carType: V(COL.CAR_TYPE || 9),
    plate: V(COL.PLATE || 10),
    carName: V(COL.CAR_SELECTED || 11),
    driver: V(COL.DRIVER || 12),
    startDate: V(COL.START_D || 13),
    startTime: V(COL.START_T || 14),
    endDate: V(COL.END_D || 15),
    endTime: V(COL.END_T || 16),
    passengers: V(COL.PASSENGERS || 17),
    bookingId: id,
    fileUrl: V(COL.FILE || 19),
    reason: V(COL.REASON || 20),
    cancelReason: V(COL.CANCEL_REASON || 21),
    vehicleCount: V(COL.VEHICLE_COUNT || 22)
  };

  return obj;
}

function getHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]).trim();
    if (key) map[key] = i + 1; // เก็บเป็น 1-based index
  }
  return map;
}

function normalizeDriverName(name) {
  if (!name) return '-';
  var s = String(name);
  
  // 1. Remove zero-width chars COMPLETELY
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');

  // 2. Convert control chars to space
  s = s.replace(/[\r\n\t]/g, ' ');

  // 3. Collapse multiple spaces
  s = s.replace(/\s+/g, ' ');

  // 4. Trim
  s = s.trim();

  // 5. [Berry Fix] Specific Correction
  if (s === 'ปรีชา ถวิล เวช') return 'ปรีชา ถวิลเวช';

  return s || '-';
}


// ===================== CACHE MANAGEMENT =====================
function cachePut_(key, data, seconds) {
  try {
    var ttl = (typeof seconds === 'number' ? seconds : CACHE_SEC);
    CacheService.getScriptCache().put(key, JSON.stringify(data), ttl);
  } catch (e) { Logger.log('Cache put error: ' + e.toString()); }
}

function cacheGet_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { 
    Logger.log('Cache get error: ' + e.toString());
    return null; 
  }
}

function cacheDelete_(key){ 
  try{ 
    CacheService.getScriptCache().remove(key); 
  } catch(e){ 
    Logger.log('Cache delete error: ' + e.toString());
  } 
}

function cacheDelete(key) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(String(key || ''));
    return true;
  } catch (e) {
    Logger.log('cacheDelete fail: ' + (e && e.message ? e.message : e));
    return false;
  }
}


// ===================== VEHICLE MANAGEMENT =====================
function getAllVehiclePlatesFromSettings() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_VEHICLES);
    if (!sh) return { ok:false, error:"ไม่พบชีต 'Vehicles'" };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, vans:[], trucks:[], all:[] };

    const vs = sh.getRange(1,1,lastRow,sh.getLastColumn()).getDisplayValues();
    const headers = vs[0].map(v => String(v||'').trim().toLowerCase());

    // 💖 Helper หา Index หัวตาราง (รองรับไทย/อังกฤษ)
    function findIdx(keywords) {
      for (var k of keywords) {
        var ix = headers.indexOf(k);
        if (ix > -1) return ix;
      }
      return -1;
    }

    const ixPlate = findIdx(['plate', 'ทะเบียน', 'ทะเบียนรถ', 'เลขทะเบียน']);
    const ixName  = findIdx(['name', 'ยี่ห้อ', 'ชื่อรถ', 'รุ่น', 'brand']);
    const ixType  = findIdx(['type', 'ประเภท', 'ชนิด', 'car_type']);
    
    // ถ้าหาไม่เจอ ให้ลองเดา (Col 1=Plate, Col 2=Name, Col 3=Type)
    const pIdx = ixPlate > -1 ? ixPlate : 0;
    const nIdx = ixName  > -1 ? ixName  : 1;
    const tIdx = ixType  > -1 ? ixType  : 2;

    // อ่านสถานะซ่อมบำรุงจาก Setting (VehicleAvailability)
    const vStatusKv = readSettingKV_('VehicleAvailability');
    const vStatusMap = parseBoolMap_(vStatusKv.val);

    const rows = vs.slice(1).filter(r => (r[pIdx] || '').trim());
    const all = rows.map(r => {
      const plate = String(r[pIdx]||'').trim();
      const rawType = String(r[tIdx]||'').trim().toLowerCase();
      
      // เช็คสถานะ (ถ้าไม่มีใน Map ให้ถือว่า True/Active)
      const isActive = vStatusMap.hasOwnProperty(plate) ? vStatusMap[plate] : true;

      return {
        plate: plate,
        name:  String(r[nIdx]||'').trim(),
        type:  rawType || 'van',
        active: isActive // ส่งสถานะไปด้วย
      };
    });
    
    const vans   = all.filter(x => x.type.includes('van') || x.type.includes('ตู้'));
    const trucks = all.filter(x => x.type.includes('truck') || x.type.includes('กระบะ') || x.type.includes('บรรทุก'));
    
    return { ok:true, vans, trucks, all };

  } catch (e) {
    Logger.log('getAllVehiclePlatesFromSettings Error: ' + e.stack);
    return { ok:false, error:e.message, vans:[], trucks:[], all:[] };
  }
}

function getDriversFromVehicles_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_VEHICLES);
    if (!sh) throw new Error("Sheet 'Vehicles' not found");
    var vs = sh.getDataRange().getDisplayValues();
    if (vs.length < 2) return [];

    var headers = vs[0].map(function (v) { return String(v || '').trim().toLowerCase(); });
    var ixDriver = (function (hs) {
      var keys = ['driverlist', 'drivers', 'คนขับ', 'พนักงานขับรถ'];
      for (var i = 0; i < hs.length; i++) {
        var h = hs[i];
        for (var k = 0; k < keys.length; k++) if (h === keys[k]) return i;
      }
      return -1;
    })(headers);
    if (ixDriver === -1) return [];

    var out = {};
    for (var r = 1; r < vs.length; r++) {
      var raw = String(vs[r][ixDriver] || '').trim();
      if (!raw) continue;
      raw.split(/[\n,;]/).forEach(function (x) {
        var s = String(x || '').trim();
        if (s) out[s] = true;
      });
    }
    return Object.keys(out).sort();
  } catch (e) {
    Logger.log('getDriversFromVehicles_ error: ' + e.toString());
    return [];
  }
}

function getDistinctProjectsFromData_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error("Sheet Data not found");
    var rng = sh.getDataRange().getValues();
    if (rng.length < 2) return [];

    var headers = rng[0];
    var idx = headerIndex_(headers);
    var col = idx.project;
    if (col === undefined) return [];

    var set = {};
    for (var r = 1; r < rng.length; r++) {
      var val = String(rng[r][col] || '').trim();
      if (val) set[val] = true;
    }
    return Object.keys(set).sort();
  } catch (e) {
    Logger.log('getDistinctProjectsFromData_ error: ' + e.toString());
    return [];
  }
}

function getUsageCountsThisMonth_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error('Data sheet not found');
    var vs = sh.getDataRange().getValues();
    if (vs.length < 2) return {};

    var headers = vs[0];
    var idx = headerIndex_(headers);
    var ixStatus = idx.status;
    var ixPlate = idx.vehicle;
    var ixStartDate = idx.startDate;

    if (ixPlate === undefined || ixStartDate === undefined) return {};

    var now = new Date();
    var y = now.getFullYear();
    var m = now.getMonth();
    var start = new Date(y, m, 1);
    var end = new Date(y, m + 1, 1);

    var counts = {};
    for (var r = 1; r < vs.length; r++) {
      var row = vs[r];
      var plate = String(row[ixPlate] || '').trim();
      if (!plate) continue;

      var d = row[ixStartDate];
      var dt = (d instanceof Date) ? d : new Date(d);
      if (!(dt instanceof Date) || isNaN(dt.getTime())) continue;
      if (dt < start || dt >= end) continue;

      var s = (ixStatus !== undefined) ? String(row[ixStatus] || '').trim().toUpperCase() : '';
      // นับเฉพาะงานที่อนุมัติ/ใช้งานจริง
      if (s && s !== 'A' && s !== 'APPROVED' && s !== 'อนุมัติ') continue;

      counts[plate] = (counts[plate] || 0) + 1;
    }
    return counts;
  } catch (e) {
    Logger.log('getUsageCountsThisMonth_ error: ' + e.toString());
    return {};
  }
}

function apiGetDashboardData() {
  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) throw new Error('Spreadsheet not found');

    var sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error('Data sheet not found');

    var vs = sh.getDataRange().getValues();
    if (!vs || vs.length < 2) {
      return {
        ok: true,
        pending: 0,
        approved: 0,
        rejected: 0,
        cancelled: 0
      };
    }

    var headers = vs[0];
    var idx = headerIndex_(headers);
    var ixStatus = idx.status;

    if (ixStatus === undefined) {
      throw new Error('ไม่พบคอลัมน์สถานะในชีต Data (status)');
    }

    var counts = {
      pending: 0,
      approved: 0,
      rejected: 0,
      cancelled: 0
    };

    for (var r = 1; r < vs.length; r++) {
      var row = vs[r];
      if (!row) continue;

      var rawStatus = row[ixStatus];
      var norm = normalizeStatus_(rawStatus);  // ใช้ mapping เดิมในระบบ

      if (counts.hasOwnProperty(norm)) {
        counts[norm] = (counts[norm] || 0) + 1;
      }
    }

    Logger.log(
      'apiGetDashboardData: P=' + counts.pending +
      ', A=' + counts.approved +
      ', R=' + counts.rejected +
      ', C=' + counts.cancelled
    );

    return {
      ok: true,
      pending: counts.pending || 0,
      approved: counts.approved || 0,
      rejected: counts.rejected || 0,
      cancelled: counts.cancelled || 0
    };

  } catch (e) {
    Logger.log('apiGetDashboardData error: ' + e.toString());
    return {
      ok: false,
      error: e.message
    };
  }
}


function apiGetFuelFormOptions() {
  try {
    var platesRes = getAllVehiclePlatesFromSettings();
    if (!platesRes.ok) throw new Error(platesRes.error || 'Load plates failed');
    var plates = platesRes.all.map(function (v) { return v.plate; });

    var drivers = getDriversFromVehicles_();
    var projects = getDistinctProjectsFromData_();
    var counts = getUsageCountsThisMonth_();

    return { ok: true, plates: plates, drivers: drivers, projects: projects, counts: counts };
  } catch (e) {
    Logger.log('apiGetFuelFormOptions error: ' + e.toString());
    return { ok: false, error: e.message, plates: [], drivers: [], projects: [], counts: {} };
  }
}

/* ===== REMINDER: Insurance & Maintenance (3 days early, TH locale) ===== */
function _vb_norm(s){ return String(s||'').replace(/\s+/g,'').toLowerCase(); }
function _vb_idx(headers, aliases){
  var map = headers.map(function(h){ return _vb_norm(h); });
  for (var i=0;i<map.length;i++){
    for (var j=0;j<aliases.length;j++){
      if (map[i] === _vb_norm(aliases[j])) return i;
    }
  }
  return -1;
}

function _vb_parseDate(v){
  if (v instanceof Date) return v;
  if (v == null) return null;
  var s = String(v).trim();
  if (!s) return null;

  // dd/MM/yyyy or dd/MM/yyyy HH:mm (รองรับ พ.ศ. 25xx)
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2})[:\.](\d{2}))?$/);
  if (m){
    var d = +m[1], mo = +m[2], y = +m[3];
    if (y > 2400) y -= 543; // พ.ศ. -> ค.ศ.
    var hh = +m[4] || 0, mm = +m[5] || 0;
    return new Date(y, mo - 1, d, hh, mm);
  }

  // yyyy-MM-dd or yyyy-MM-dd HH:mm (กันเคส export เป็นแบบ ISO)
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2})[:\.](\d{2}))?$/);
  if (m){
    var y2 = +m[1]; if (y2 > 2400) y2 -= 543;
    var mo2 = +m[2], d2 = +m[3], hh2 = +m[4] || 0, mm2 = +m[5] || 0;
    return new Date(y2, mo2 - 1, d2, hh2, mm2);
  }

  var dflt = new Date(s);
  return isNaN(dflt) ? null : dflt;
}


function _vb_fmtThaiDate(d){
  if (!(d instanceof Date)) d = new Date(d);
  var thYear = d.getFullYear() + 543;
  return Utilities.formatDate(d, TZ, 'dd/MM/') + thYear;
}
function _vb_fmtThaiTime(d){
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, TZ, 'HH.mm') + ' น.';
}
function _vb_daysDiff(a, b){ // b - a (in days)
  var ms = new Date(b.getFullYear(), b.getMonth(), b.getDate()) - new Date(a.getFullYear(), a.getMonth(), a.getDate());
  return Math.round(ms / 86400000);
}

function _vb_getSettingNumber(key, fallback){
  var fb = (fallback != null) ? Number(fallback) : ((VB_CFG && VB_CFG.ADVANCE_DAYS) ? Number(VB_CFG.ADVANCE_DAYS) : 3);
  try{
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_SETTING); // ใช้คอนสแตนต์
    if (!sh) return fb;
    var last = Math.max(2, sh.getLastRow());
    var vals = sh.getRange(1,1,last,2).getValues();
    for (var i=0;i<vals.length;i++){
      if (_vb_norm(vals[i][0]) === _vb_norm(key)) return Number(vals[i][1] || fb);
    }
    return fb;
  }catch(e){ return fb; }
}

function _vb_getSettingString(key, fallback){
  try{
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTING);
    if (!sh) return String(fallback||'');
    var rng = sh.getRange(1,1,Math.max(2, sh.getLastRow()), 2).getValues();
    for (var i=0;i<rng.length;i++){
      if (String(rng[i][0]).trim().toLowerCase() === String(key||'').trim().toLowerCase()){
        return String(rng[i][1]||'');
      }
    }
    return String(fallback||'');
  }catch(e){ return String(fallback||''); }
}

function _vb_csvToArray(s){
  return String(s||'').split(/[,\|]/).map(function(x){return x.trim();}).filter(function(x){return !!x;});
}

function _vb_findFirstDateColumn_(vals){
  var head = vals[0] || [];
  var last = Math.min(vals.length, Math.max(2, Math.min(50, vals.length)));
  var bestIdx = -1, bestScore = 0;
  for (var c=0; c<head.length; c++){
    var score = 0;
    for (var r=1; r<last; r++){
      var d = _vb_parseDate(vals[r][c]);
      if (d) score++;
    }
    if (score > bestScore){ bestScore = score; bestIdx = c; }
  }
  return (bestScore >= 1) ? bestIdx : -1;
}


/* ---------- Read DUEs: Insurance ---------- */
function _vb_collectInsuranceDue_(){
  var out = [];
  var ss = SpreadsheetApp.getActive();

  var shVS = ss.getSheetByName(SHEET_VEHICLE_STATUS);
  if (shVS && shVS.getLastRow() >= 2){
    var vals = shVS.getRange(1,1,shVS.getLastRow(), shVS.getLastColumn()).getValues();
    var head = vals[0] || [];
    var ixPlate = _vb_idx(head, ['ทะเบียน','plate','เลขทะเบียนรถ']);
    var pref = _vb_csvToArray(_vb_getSettingString('InsuranceDueHeader','วันหมดอายุประกัน,วันสิ้นสุดประกัน,สิ้นสุดประกัน,ครบกำหนดประกัน,insuranceend,enddate'));
    var ixDue = _vb_idx(head, pref);
    if (ixDue < 0) ixDue = _vb_findFirstDateColumn_(vals);
    for (var r=1;r<vals.length;r++){
      var plate = (ixPlate>=0)? String(vals[r][ixPlate]||'').trim() : '';
      var due   = (ixDue  >=0)? _vb_parseDate(vals[r][ixDue]) : null;
      if (plate && due) out.push({ type:'insurance', plate:plate, due:due, source:'VehicleStatus' });
    }
  }

  var shI = ss.getSheetByName(SHEET_INSURANCE);
  if (shI && shI.getLastRow() >= 2){
    var valsI = shI.getRange(1,1,shI.getLastRow(), shI.getLastColumn()).getValues();
    var headI = valsI[0] || [];
    var ixPlateI = _vb_idx(headI, ['ทะเบียนรถ','ทะเบียน','plate','เลขทะเบียนรถ']);
    var prefI = _vb_csvToArray(_vb_getSettingString('InsuranceSheetDueHeader','วันสิ้นสุด,วันสิ้นสุดประกัน,วันที่สิ้นสุด,ครบกำหนด,end,enddate,หมดอายุ'));
    var ixDueI = _vb_idx(headI, prefI);
    if (ixDueI < 0) ixDueI = _vb_findFirstDateColumn_(valsI);
    for (var i=1;i<valsI.length;i++){
      var plateI = (ixPlateI>=0)? String(valsI[i][ixPlateI]||'').trim() : '';
      var dueI   = (ixDueI  >=0)? _vb_parseDate(valsI[i][ixDueI]) : null;
      if (plateI && dueI) out.push({ type:'insurance', plate:plateI, due:dueI, source:'Insurance' });
    }
  }
  return out;
}

/* ---------- Read DUEs: Maintenance NextDueDate ---------- */
function _vb_collectMaintenanceDue_(){
  var out = [];
  var ss = SpreadsheetApp.getActive();

  var shVS = ss.getSheetByName(SHEET_VEHICLE_STATUS);
  if (shVS && shVS.getLastRow() >= 2){
    var vals = shVS.getRange(1,1,shVS.getLastRow(), shVS.getLastColumn()).getValues();
    var head = vals[0] || [];
    var ixPlate = _vb_idx(head, ['ทะเบียน','plate','เลขทะเบียนรถ']);
    var pref = _vb_csvToArray(_vb_getSettingString('MaintenanceNextDueHeader','รอบบริการถัดไป,nextduedate,nextservice,กำหนดครั้งถัดไป,กำหนดเข้าซ่อม'));
    var ixDue = _vb_idx(head, pref);
    if (ixDue < 0) ixDue = _vb_findFirstDateColumn_(vals);
    for (var r=1;r<vals.length;r++){
      var plate = (ixPlate>=0)? String(vals[r][ixPlate]||'').trim() : '';
      var due   = (ixDue  >=0)? _vb_parseDate(vals[r][ixDue]) : null;
      if (plate && due) out.push({ type:'maintenance', plate:plate, due:due, source:'VehicleStatus' });
    }
  }

  var shM = ss.getSheetByName(SHEET_MAINTENANCE);
  if (shM && shM.getLastRow() >= 2){
    var valsM = shM.getRange(1,1,shM.getLastRow(), shM.getLastColumn()).getValues();
    var headM = valsM[0] || [];
    var ixPlateM = _vb_idx(headM, ['plate','ทะเบียน','ทะเบียนรถ','เลขทะเบียนรถ']);
    var prefM = _vb_csvToArray(_vb_getSettingString('MaintenanceSheetNextDueHeader','nextduedate,รอบบริการถัดไป,กำหนดครั้งถัดไป,กำหนดเข้าซ่อม'));
    var ixDueM = _vb_idx(headM, prefM);
    if (ixDueM < 0) ixDueM = _vb_findFirstDateColumn_(valsM);
    for (var j=1;j<valsM.length;j++){
      var plateM = (ixPlateM>=0)? String(valsM[j][ixPlateM]||'').trim() : '';
      var dueM   = (ixDueM  >=0)? _vb_parseDate(valsM[j][ixDueM]) : null;
      if (plateM && dueM) out.push({ type:'maintenance', plate:plateM, due:dueM, source:'Maintenance' });
    }
  }
  return out;
}



/* ---------- Core check & notify ---------- */
function runInsuranceReminder_(){
  var lead = _vb_getSettingNumber('InsuranceReminderDays', 3);
  var now = new Date();
  var list = _vb_collectInsuranceDue_();
  var sent = 0;

  for (var i = 0; i < list.length; i++){
    var it = list[i];
    var days = _vb_daysDiff(now, it.due);
    if (days < 0) continue;      // already expired
    if (days > lead) continue;   // not yet within window

    var dd = _vb_fmtThaiDate(it.due) + ' เวลา ' + _vb_fmtThaiTime(it.due);
    var msg = '🛡️ แจ้งเตือนประกันใกล้หมดอายุ\n'
            + 'ทะเบียน: ' + it.plate + '\n'
            + 'ครบกำหนด: ' + dd + '\n'
            + 'แหล่งข้อมูล: ' + it.source;

    var key = 'REM:INS:' + it.plate + ':' + Utilities.formatDate(it.due, TZ, 'yyyyMMdd');
    var r = sendTelegramOnce(msg, { parse_mode:'HTML', disable_preview:true, dedupeKey:key, force:true });
    if (r && r.ok) sent++;
  }

  try { Logger.log('Insurance reminder sent: ' + sent + '/' + list.length); } catch(_){}
  return { ok:true, sent:sent, total:list.length };
}

function runMaintenanceReminder_(){
  var lead = _vb_getSettingNumber('MaintenanceReminderDays', 3);
  var now = new Date();
  var list = _vb_collectMaintenanceDue_();
  var sent = 0;

  for (var i = 0; i < list.length; i++){
    var it = list[i];
    var days = _vb_daysDiff(now, it.due);
    if (days < 0) continue;
    if (days > lead) continue;

    var dd = _vb_fmtThaiDate(it.due) + ' เวลา ' + _vb_fmtThaiTime(it.due);
    var msg = '🧰 แจ้งเตือนกำหนดเข้ารับบริการซ่อมครั้งถัดไป\n'
            + 'ทะเบียน: ' + it.plate + '\n'
            + 'กำหนด: ' + dd + '\n'
            + 'แหล่งข้อมูล: ' + it.source;

    var key = 'REM:MAINT:' + it.plate + ':' + Utilities.formatDate(it.due, TZ, 'yyyyMMdd');
    var r = sendTelegramOnce(msg, { parse_mode:'HTML', disable_preview:true, dedupeKey:key, force:true });
    if (r && r.ok) sent++;
  }

  try { Logger.log('Maintenance reminder sent: ' + sent + '/' + list.length); } catch(_){}
  return { ok:true, sent:sent, total:list.length };
}


function runAllReminders_(){
  var a = runInsuranceReminder_();
  var b = runMaintenanceReminder_();
  return { ok:true, insurance:a, maintenance:b };
}



/********************** SETTINGS & COMMON HELPERS **********************/
function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('setting');
}

function getSetting_(key) {
  try {
    const sh = getSettingsSheet_();
    if (!sh) return null;
    const rng = sh.getRange(1, 1, sh.getLastRow(), 2).getValues(); // คอลัมน์ A=key, B=value
    for (var i = 0; i < rng.length; i++) {
      if (String(rng[i][0]).trim() === String(key).trim()) return String(rng[i][1] || '').trim();
    }
    return null;
  } catch (e) { return null; }
}

function firstExistingFunction_(names) {
  for (var i = 0; i < names.length; i++) {
    var n = names[i];
    try { if (n && typeof this[n] === 'function') return n; } catch (_){}
  }
  return null;
}

function getVehiclesPlates_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Vehicles');
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const header = vals[0];
  const idxPlate = header.indexOf('plate');
  if (idxPlate < 0) return [];
  const set = {};
  for (var r = 1; r < vals.length; r++) {
    const v = String(vals[r][idxPlate] || '').trim();
    if (v) set[v] = true;
  }
  return Object.keys(set);
}

/********************** FUEL/INSURANCE/MAINTENANCE FORM APIs **********************/
function _getPlateListFromVehiclesSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Vehicles'); // ชื่อชีตต้องเป๊ะ
    if (!sh) return [];
    
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return []; 
    
    const data = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const plateIndex = headers.indexOf('plate'); // หาคอลัมน์ plate
    
    if (plateIndex === -1) return [];

    // ดึงเฉพาะคอลัมน์ plate, ตัดหัวตาราง, กรองค่าว่าง
    const plates = data.slice(1)
      .map(r => String(r[plateIndex]).trim())
      .filter(p => p !== '');
      
    return [...new Set(plates)].sort(); // ตัดซ้ำและเรียงลำดับ
  } catch (e) {
    Logger.log('Error getting plates: ' + e.message);
    return [];
  }
}

// API สำหรับ Tab ประกันภัย (ดึงทะเบียนจริงจาก Vehicles)
function apiGetInsurancePlates() {
  return { ok: true, plates: _getPlateListFromVehiclesSheet() };
}

// API สำหรับ Tab ซ่อมบำรุง (ดึงทะเบียนจริงจาก Vehicles)
function apiGetMaintenancePlates() {
  return { ok: true, plates: _getPlateListFromVehiclesSheet() };
}

/********************** THAI DATE/TIME (B.E.) **********************/
function pad2_(n){ return (n < 10 ? '0' : '') + n; }
function toThaiDate_(dt) {
  const tz = 'Asia/Bangkok';
  const y = parseInt(Utilities.formatDate(dt, tz, 'yyyy'), 10) + 543;
  const m = Utilities.formatDate(dt, tz, 'MM');
  const d = Utilities.formatDate(dt, tz, 'dd');
  return d + '/' + m + '/' + y;
}
function toThaiDateTime_(dt) {
  const tz = 'Asia/Bangkok';
  return toThaiDate_(dt) + '   ⏰ ' + Utilities.formatDate(dt, tz, 'HH:mm') + ' น.';
}

/* ===================== DASHBOARD & REPORT APIs (BERRY FIXED - COMPLETE) ===================== */
// 1. Helper: จัดการการเรียกฟังก์ชันแบบปลอดภัย
function __vb_invoke_(names) {
  for (var i = 0; i < names.length; i++) {
    var fn = this[names[i]];
    if (typeof fn === 'function') return fn;
  }
  return null;
}

function _ok_(obj){ obj = obj || {}; obj.ok = true; return obj; }
function _err_(e){ return { ok:false, error: (e && e.message) ? e.message : String(e) }; }
function __vb_pad2(n){ return (n<10 ? '0' : '') + n; }

// Helper สำหรับแปลง HTML เป็น PDF และบันทึกลง Drive (ใช้ร่วมกัน)
function __vb_htmlToPdfUrl__(filename, html) {
  var name = String(filename||'report.pdf');
  if (!/\.pdf$/i.test(name)) name += '.pdf';
  var pdfBlob = Utilities.newBlob(html, 'text/html', 'tmp.html').getAs('application/pdf').setName(name);
  
  // สร้างไฟล์ที่ Root Folder เพื่อให้ได้ URL ที่เปิดได้จริง
  var file = DriveApp.createFile(pdfBlob);
  // ตั้งค่าแชร์ให้อ่านได้ทุกคนที่มีลิงก์ (สำคัญมาก ไม่งั้น Client เปิดไม่ได้)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

// 2. API: รีเฟรชข้อมูล Dashboard (เคลียร์ Cache)
function apiRefreshDashboard() {
  try {
    const cacheKey = 'mainDataCache_v13_BerryFix'; // ต้องตรงกับที่ใช้ใน getMainData_
    cacheDelete_(cacheKey);
    return _ok_({ message: 'Refreshed' });
  } catch (e) { 
    return _err_(e); 
  }
}

function buildDashboardPdfData(year, month, fuelData) {
  const out = {
    totalBookings: 0,
    vehiclesReady: '0/0',
    readyVehicles: [],
    alerts: 0,
    fuel: 0,
    topDrivers: [],
    topVehicles: []
  };

  const y0 = Number(year);
  const m0 = Number(month) - 1;

  function normalizePlate(v) {
    if (v == null) return '';
    if (typeof v === 'string' || typeof v === 'number') return String(v).trim();

    if (typeof v === 'object') {
      const keys = ['plate', 'ทะเบียนรถ', 'เลขทะเบียนรถ', 'registration', 'reg', 'vehiclePlate', 'name', 'value', 'text', 'label'];
      for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (v[k] != null && String(v[k]).trim()) return String(v[k]).trim();
      }
      return '';
    }
    return '';
  }

  // 1) Fuel
  try {
    if (Array.isArray(fuelData) && fuelData.length) {
      out.fuel = fuelData.reduce((acc, r) => {
        const ts = r && (r.ts || r.date || r.datetime);
        const dt = ts ? new Date(ts) : null;
        if (!dt || isNaN(dt.getTime())) return acc;
        if (dt.getFullYear() !== y0 || dt.getMonth() !== m0) return acc;

        const liters = parseFloat(String(r.liters || r.liter || '0').replace(/,/g, '')) || 0;
        return acc + liters;
      }, 0);
    }
  } catch (_) {}

  // 2) Vehicles ready
  try {
    if (
      typeof getAllVehiclePlatesFromSettings === 'function' &&
      typeof readSettingKV_ === 'function' &&
      typeof parseBoolMap_ === 'function'
    ) {
      const res = getAllVehiclePlatesFromSettings();
      const allRaw = (res && Array.isArray(res.all)) ? res.all : [];

      const kv = readSettingKV_('VehicleAvailability');
      const map = parseBoolMap_(kv && kv.val);

      const seen = {};
      const allPlates = [];

      allRaw.forEach(item => {
        const plate = normalizePlate(item);
        if (!plate) return;
        if (seen[plate]) return;
        seen[plate] = true;
        allPlates.push(plate);
      });

      const total = allPlates.length;
      const readyList = [];
      let ready = 0;

      allPlates.forEach(plate => {
        const isReady = Object.prototype.hasOwnProperty.call(map, plate) ? !!map[plate] : true;
        if (isReady) {
          ready++;
          if (readyList.length < 5) readyList.push(plate);
        }
      });

      out.vehiclesReady = `${ready}/${total}`;
      out.readyVehicles = readyList;
    }
  } catch (e) {
    Logger.log('buildDashboardPdfData vehiclesReady error: ' + e);
    out.vehiclesReady = out.vehiclesReady || '0/0';
    out.readyVehicles = Array.isArray(out.readyVehicles) ? out.readyVehicles : [];
  }

  // 3) Summary from Data Sheet
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Data');
    if (!sh) {
      out.topDrivers = [{ name: 'ไม่มีข้อมูล', count: 0 }];
      out.topVehicles = [{ plate: 'ไม่มีข้อมูล', trips: 0 }];
      return out;
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      out.topDrivers = [{ name: 'ไม่มีข้อมูล', count: 0 }];
      out.topVehicles = [{ plate: 'ไม่มีข้อมูล', trips: 0 }];
      return out;
    }

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = (values[0] || []).map(h => String(h || '').trim());
    const hmap = {};
    headers.forEach((h, i) => { if (h) hmap[h] = i; });

    const ixBookingId = (hmap['Booking ID'] != null) ? hmap['Booking ID'] : -1;
    const ixStatus = (hmap['สถานะ'] != null) ? hmap['สถานะ'] : -1;
    const ixDriver = (hmap['พนักงานขับรถ'] != null) ? hmap['พนักงานขับรถ'] : -1;
    const ixStartDate = (hmap['วันเริ่มต้น'] != null) ? hmap['วันเริ่มต้น'] : -1;

    const ixVehicleSelected = (hmap['รถที่เลือก'] != null) ? hmap['รถที่เลือก'] : -1;
    const ixPlate = (hmap['เลขทะเบียนรถ'] != null) ? hmap['เลขทะเบียนรถ'] : -1;

    if (ixBookingId < 0 || ixStatus < 0 || ixDriver < 0 || ixStartDate < 0) {
      out.topDrivers = [{ name: 'ไม่มีข้อมูล', count: 0 }];
      out.topVehicles = [{ plate: 'ไม่มีข้อมูล', trips: 0 }];
      return out;
    }

    const driverCount = {};
    const vehicleTrips = {};

    let totalBookings = 0;
    let alerts = 0;

    for (let r = 1; r < values.length; r++) {
      const row = values[r] || [];

      const bookingId = String(row[ixBookingId] || '').trim();
      if (!bookingId) continue;

      let dt = null;
      const raw = row[ixStartDate];

      if (raw instanceof Date && !isNaN(raw.getTime())) {
        dt = raw;
      } else {
        const s = String(raw || '').trim();
        if (s) {
          const mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
          const mDm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (mIso) {
            dt = new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]));
          } else if (mDm) {
            let yy = Number(mDm[3]);
            if (yy > 2400) yy -= 543;
            dt = new Date(yy, Number(mDm[2]) - 1, Number(mDm[1]));
          } else {
            const t = new Date(s);
            if (!isNaN(t.getTime())) dt = t;
          }
        }
      }

      if (!dt || isNaN(dt.getTime())) continue;
      if (dt.getFullYear() !== y0 || dt.getMonth() !== m0) continue;

      totalBookings++;

      const statusRaw = row[ixStatus];
      const statusKey = (typeof getStatusKeySafe_ === 'function')
        ? String(getStatusKeySafe_(statusRaw) || '').toLowerCase().trim()
        : String(statusRaw || '').toLowerCase().trim();

      const isPending =
        statusKey === 'pending' ||
        statusKey === 'wait' ||
        statusKey === 'waiting' ||
        statusKey.indexOf('รอ') === 0;

      if (isPending) alerts++;

      const isApproved =
        statusKey === 'approved' ||
        statusKey.indexOf('อนุมัติ') !== -1;

      if (isApproved) {
        // [Berry Fix] Normalize Driver Name
        const drivers = String(row[ixDriver] || '')
          .split(',')
          .map(s => normalizeDriverName(s)) // ใช้ Helper
          .filter(Boolean);

        const plateRaw =
          ((ixVehicleSelected >= 0) ? String(row[ixVehicleSelected] || '').trim() : '') ||
          ((ixPlate >= 0) ? String(row[ixPlate] || '').trim() : '');

        const plates = plateRaw
          .split(',')
          .map(s => String(s || '').trim())
          .filter(Boolean);

        drivers.forEach(name => { driverCount[name] = (driverCount[name] || 0) + 1; });
        plates.forEach(plate => { vehicleTrips[plate] = (vehicleTrips[plate] || 0) + 1; });
      }
    }

    out.totalBookings = totalBookings;
    out.alerts = alerts;

    const topDrivers = Object.keys(driverCount)
      .map(k => ({ name: k, count: driverCount[k] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    const topVehicles = Object.keys(vehicleTrips)
      .map(k => ({ plate: k, trips: vehicleTrips[k] }))
      .sort((a, b) => b.trips - a.trips)
      .slice(0, 10);

    out.topDrivers = topDrivers.length ? topDrivers : [{ name: 'ไม่มีข้อมูล', count: 0 }];
    out.topVehicles = topVehicles.length ? topVehicles : [{ plate: 'ไม่มีข้อมูล', trips: 0 }];

    if (!Array.isArray(out.readyVehicles)) out.readyVehicles = [];
    out.readyVehicles = out.readyVehicles.map(normalizePlate).filter(Boolean).slice(0, 5);

    return out;
  } catch (e) {
    Logger.log('buildDashboardPdfData error: ' + (e && e.stack ? e.stack : e));
    out.topDrivers = Array.isArray(out.topDrivers) && out.topDrivers.length ? out.topDrivers : [{ name: 'ไม่มีข้อมูล', count: 0 }];
    out.topVehicles = Array.isArray(out.topVehicles) && out.topVehicles.length ? out.topVehicles : [{ plate: 'ไม่มีข้อมูล', trips: 0 }];
    out.readyVehicles = Array.isArray(out.readyVehicles) ? out.readyVehicles.map(normalizePlate).filter(Boolean).slice(0, 5) : [];
    out.vehiclesReady = out.vehiclesReady || '0/0';
    out.totalBookings = Number(out.totalBookings) || 0;
    out.alerts = Number(out.alerts) || 0;
    out.fuel = Number(out.fuel) || 0;
    return out;
  }
}

/* [ANCHOR: Generate PDF Common (Berry Improved - Auto Summary & Header)] */
function _generatePdfCommon_(tplName, dataOpt) {
  try {
    dataOpt = dataOpt || {};
    var now = new Date();
    var y = now.getFullYear();
    var m = now.getMonth() + 1;
    var d = now.getDate();
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

    // 1. เตรียมข้อมูลพื้นฐาน
    var mainData = {};
    var fuelData = [];
    var reportTitle = '';
    var reportPeriod = '';

    // 2. ดึงข้อมูล Dashboard (ถ้าจำเป็น)
    if (tplName === 'DashboardReport') {
      try {
        var dRes = getMainData_();
        if (dRes && dRes.ok) mainData = dRes.data || {};
        
        if (typeof buildDashboardPdfData === 'function') {
           var fRes = apiGetFuelHistory();
           var allFuel = (fRes && fRes.ok) ? fRes.data : [];
           var dashData = buildDashboardPdfData(y, m, allFuel);
           for (var k in dashData) mainData[k] = dashData[k];
        }
      } catch (e) { console.warn('Dashboard data error', e); }
    }

    // 3. ดึงและกรองข้อมูลน้ำมัน (สำหรับ FuelReport)
    if (tplName === 'FuelReport') {
      var fRes = apiGetFuelHistory();
      var allFuel = (fRes && fRes.ok) ? fRes.data : [];
      
      if (dataOpt.day) {
        // --- รายงานรายวัน ---
        reportTitle = 'รายงานสรุปน้ำมัน (รายวัน)';
        var thDay = Utilities.formatDate(now, tz, 'dd/MM/');
        var thY = parseInt(Utilities.formatDate(now, tz, 'yyyy')) + 543;
        reportPeriod = 'ประจำวันที่ ' + thDay + thY;
        
        fuelData = allFuel.filter(function(row) {
           if (!row.timestamp) return false;
           var rd = new Date(row.timestamp);
           return rd.getDate() === d && rd.getMonth() === (m-1) && rd.getFullYear() === y;
        });

      } else {
        // --- รายงานรายเดือน ---
        reportTitle = 'รายงานสรุปน้ำมัน (เดือนนี้)';
        var thY2 = parseInt(Utilities.formatDate(now, tz, 'yyyy')) + 543;
        reportPeriod = 'ประจำเดือน ' + m + '/' + thY2;
        
        fuelData = allFuel.filter(function(row) {
           if (!row.timestamp) return false;
           var rd = new Date(row.timestamp);
           return rd.getMonth() === (m-1) && rd.getFullYear() === y;
        });
      }
    }

    // 4. [BERRY FIX] คำนวณยอดรวมและ Grouping (แก้ปัญหาตารางสรุปว่าง)
    var totalCost = 0;
    var totalLiters = 0;
    var summaryMap = {}; // เก็บข้อมูลแยกตามทะเบียนรถ

    fuelData.forEach(function(r) {
       var c = (r.cost || 0);
       var l = (r.liters || 0);
       var p = r.plate || 'ไม่ระบุ';

       totalCost += c;
       totalLiters += l;

       // Grouping Logic
       if (!summaryMap[p]) {
           summaryMap[p] = { plate: p, trips: 0, liters: 0, cost: 0 };
       }
       summaryMap[p].trips++;
       summaryMap[p].liters += l;
       summaryMap[p].cost += c;
    });

    // แปลง Map กลับเป็น Array เพื่อส่งให้ Template วนลูป
    var summaryArray = Object.keys(summaryMap).map(function(k) { return summaryMap[k]; });

    // 5. เตรียม Template Data
    var templateData = {
      year: y,
      month: m,
      day: dataOpt.day || null,
      generatedAt: Utilities.formatDate(now, tz, 'dd/MM/') + (y+543) + ' ' + Utilities.formatDate(now, tz, 'HH:mm') + ' น.',
      
      // ✅ แก้ชื่อระบบตรงนี้ (ส่งไปให้ Template)
      systemName: 'ระบบจองยานพาหนะ มหาวิทยาลัยสวนดุสิต ศูนย์การศึกษาลำปาง', 
      
      // Dashboard Data
      data: mainData,
      monthNames: ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'],
      
      // Fuel Report Data
      title: reportTitle,
      period: reportPeriod,
      detail: fuelData, 
      summary: summaryArray, // ✅ ส่งข้อมูลสรุปที่คำนวณแล้วไป
      totalCost: totalCost,
      totalLiters: totalLiters
    };

    // 6. Generate HTML & PDF
    var html = '';
    try {
      var t = HtmlService.createTemplateFromFile(tplName);
      for (var key in templateData) { t[key] = templateData[key]; }
      html = t.evaluate().getContent();
    } catch (e) {
      html = '<h1>Error Generating PDF</h1><p>' + e.message + '</p>';
    }

    var fileName = tplName + '_' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmm') + '.pdf';
    var url = __vb_htmlToPdfUrl__(fileName, html);

    return { ok: true, url: url };

  } catch (e) {
    Logger.log('_generatePdfCommon_ Error: ' + e.message);
    return { ok: false, error: e.message };
  }
}



// 4. API Wrappers (เรียกใช้จากหน้าเว็บผ่าน google.script.run)
// ฟังก์ชันเหล่านี้ต้องชื่อตรงกับที่ JS ฝั่ง Client เรียกใช้เป๊ะๆ

function apiGenerateDashboardPdf() {
  return _generatePdfCommon_('DashboardReport', {});
}

function apiGenerateFuelMonthlyPdf() {
  return _generatePdfCommon_('FuelReport', {});
}

function apiGenerateFuelDailyPdf() {
  // ส่งวันที่ปัจจุบันเข้าไปเพื่อทำรายงานรายวัน
  return _generatePdfCommon_('FuelReport', { day: new Date().getDate() });
}
// ===================== DRIVER MANAGEMENT =====================
function getDriversFromAdmin_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Admin');
    if (!sh) {
      return { ok: false, error: "ไม่พบชีต 'Admin'", drivers: [] };
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      return { ok: true, drivers: [] };
    }

    const vs = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = vs[0].map(h => String(h || '').trim().toLowerCase());
    const ixUser = headers.indexOf('username');
    const ixPass = headers.indexOf('password');
    const ixName = headers.indexOf('name');
    const ixRole = headers.indexOf('role');

    if (ixUser === -1 || ixPass === -1 || ixName === -1 || ixRole === -1) {
      return {
        ok: false,
        error: "ชีต Admin ต้องมีหัวตาราง: username | password | name | Role",
        drivers: []
      };
    }

    const statusKv = readSettingKV_('DriverStatus');
    const statusMap = parseBoolMap_(statusKv.val);

    const allowedRoles = ['driver', 'admindriver'];
    const rows = vs.slice(1);
    const drivers = rows
      .map(r => {
        const rawRole = String(r[ixRole] || '').trim();
        return {
          username: String(r[ixUser] || '').trim(),
          pass:     String(r[ixPass] || '').trim(),
          name:     String(r[ixName] || '').trim(),
          role:     rawRole,
          _roleLc:  rawRole.toLowerCase()
        };
      })
      .filter(d => d.username || d.name)
      .filter(d => allowedRoles.indexOf(d._roleLc) !== -1)
      .map(d => {
        const isActive = statusMap.hasOwnProperty(d.name) 
          ? !!statusMap[d.name] 
          : true;

        return {
          username: d.username,
          pass: d.pass,
          name: d.name,
          role: d.role,
          active: isActive
        };
      });
      
    return { ok: true, drivers: drivers };
  } catch (e) {
    return { ok: false, error: e.message, drivers: [] };
  }
}

// ===================== ADDITIONAL REQUIRED FUNCTIONS =====================
const TH_MONTHS_ = {
  'มกราคม':1,'กุมภาพันธ์':2,'มีนาคม':3,'เมษายน':4,'พฤษภาคม':5,'มิถุนายน':6,
  'กรกฎาคม':7,'สิงหาคม':8,'กันยายน':9,'ตุลาคม':10,'พฤศจิกายน':11,'ธันวาคม':12,
  'ม.ค.':1,'ก.พ.':2,'มี.ค.':3,'เม.ย.':4,'พ.ค.':5,'มิ.ย.':6,
  'ก.ค.':7,'ส.ค.':8,'ก.ย.':9,'ต.ค.':10,'พ.ย.':11,'ธ.ค.':12
};

function thaiToArabic_(text) {
  if (text == null) return text;
  return String(text).replace(/[๐-๙]/g, d => '๐๑๒๓๔๕๖๗๘๙'.indexOf(d));
}

function parseTimeSafe_(timeInput) {
  if (timeInput instanceof Date) return Utilities.formatDate(timeInput, TZ, 'HH:mm');
  let s = thaiToArabic_(String(timeInput || '').trim()).toLowerCase();
  if (!s) return '00:00';

  s = s.replace(/นาฬิกา/g,'').replace(/น\./g,'').replace(/\s+/g,' ');
  s = s.replace(/[\. ]+/g, ':')
       .replace(/[^0-9:]/g,'')
       .replace(/:+/g,':')
       .replace(/^:|:$/g,'');

  if (/^\d{3,4}$/.test(s)) s = s.padStart(4,'0').replace(/(\d{2})(\d{2})/, '$1:$2');
  if (/^\d{1,2}$/.test(s)) s = s + ':00';

  const m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!m) return '00:00';
  const h = +m[1], mi = +m[2];
  if (h>23 || mi>59) return '00:00';
  return `${String(h).padStart(2,'0')}:${String(mi).padStart(2,'0')}`;
}

function formatPhoneNumber_(raw) {
  if (!raw) return '';
  let phone = String(raw).replace(/[\s\-\(\)]/g, '');
  if (phone.length === 9 && !phone.startsWith('0')) phone = '0' + phone;
  return phone;
}

function getStatusKeySafe_(raw) {
  try {
    if (typeof getStatusKey_ === 'function') {
      const k = getStatusKey_(raw);
      if (k) {
        if (k === 'driver_claimed') return 'pending';
        return k;
      }
    }
  } catch (_) {}

  const t = String(raw || '').toLowerCase().trim();

  if (/ไม่อนุมัติ|reject|ปฏิเสธ|fail/.test(t)) return 'rejected';
  if (/ยกเลิก|cancel/.test(t)) return 'cancelled';
  if (t === 'driver_special_approved' || /กรณีพิเศษ|เร่งด่วน/.test(t)) return 'driver_special_approved';
  if (t === 'driver_claimed' || /รับงาน/.test(t)) return 'pending';
  if (t === 'approved' || /อนุมัติ|approved|ok|pass/.test(t)) return 'approved';
  if (/รออนุมัติ|pending/.test(t)) return 'pending';

  return 'pending';
}

function headerIndex_(headArr, logs) {
  const head = headArr.map(s => String(s || '').trim());
  const idx = {};

  // helper: find first matching header from a list (case-sensitive first, then case-insensitive)
  const findHeader_ = (candidates) => {
    const list = (Array.isArray(candidates) ? candidates : [candidates])
      .map(x => String(x || '').trim())
      .filter(Boolean);

    // 1) exact match
    for (var i = 0; i < list.length; i++) {
      var p = head.indexOf(list[i]);
      if (p !== -1) return p;
    }

    // 2) case-insensitive match
    const headLower = head.map(h => String(h).toLowerCase());
    for (var j = 0; j < list.length; j++) {
      var q = headLower.indexOf(String(list[j]).toLowerCase());
      if (q !== -1) return q;
    }

    return -1;
  };

  // ✅ ใช้ COLMAP เป็นหลัก (รองรับหัวคอลัมน์หลายชื่อ เช่น email/อีเมล และ File/ไฟล์แนบ)
  Object.keys(COLMAP).forEach(k => {
    idx[k] = findHeader_(COLMAP[k]);
  });

  // Alias (คงเดิมเพื่อไม่ให้ของเก่าพัง)
  idx['place'] = idx['destination'];
  idx['plate'] = idx['vehicle'];
  idx['vehicleSelected'] = idx['requestedVehicle'];
  idx['file'] = idx['fileUrl'];

  return idx;
}

// ===================== BOOKING ID MANAGEMENT =====================
function reserveNextBookingId() {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Data');
    if (!sh) throw new Error('ไม่พบชีต Data');
    
    var maxInSheet = detectMaxBookingId_(sh);
    var lastUsed = Number(props.getProperty('COUNTER_BOOKING_ID') || '0');
    var base = Math.max(maxInSheet, lastUsed);
    var next = base + 1;
    
    props.setProperty('COUNTER_BOOKING_ID', String(next));
    
    Logger.log('reserveNextBookingId: MaxInSheet=' + maxInSheet + ', LastUsed=' + lastUsed + ', Base=' + base + ' -> Next ID=' + next);
    return next;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function detectMaxBookingId_(sheet) {
  var pos = findHeaderRowAndCol_(sheet);
  var startRow = pos.headerRow + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 0;
  
  var range = sheet.getRange(startRow, pos.idCol, lastRow - pos.headerRow + 1, 1);
  var values = range.getValues();
  
  var maxId = 0;
  for (var i = 0; i < values.length; i++) {
    var id = parseFloat(values[i][0]);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }
  return maxId;
}

function findHeaderRowAndCol_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var scanHeaderRows = Math.min(lastRow, 5);
  if (scanHeaderRows < 1) return { headerRow: 1, idCol: 18 };

  var headVals = sheet.getRange(1, 1, scanHeaderRows, lastCol).getValues();
  var aliases = ['booking id', 'bookingid', 'booking-id', 'id'];
  
  for (var r = 0; r < headVals.length; r++) {
    for (var c = 0; c < headVals[r].length; c++) {
      var cellValue = String(headVals[r][c] || '').replace(/\s+/g, ' ').trim().toLowerCase();
      if (aliases.includes(cellValue)) {
        return { headerRow: r + 1, idCol: c + 1 };
      }
    }
  }
  return { headerRow: 1, idCol: 18 };
}

// ===================== DATA ROW BUILDER =====================
function buildRowForDataSheet(parsed, bookingIdFinal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Data');
  if (!sh) throw new Error("ไม่พบชีต 'Data' ค่ะบอส!");

  // ✅ ดึง Header ล่าสุดจากหน้าชีตจริง
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim());

  if (typeof headerIndex_ !== 'function') {
    throw new Error('Missing function: headerIndex_');
  }
  const idx = headerIndex_(headers) || {};

  // ✅ สร้างแถวเปล่าตามจำนวนคอลัมน์จริงใน Sheet (Index-agnostic)
  const row = new Array(headers.length).fill('');

  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  const toStr = (v) => String(v == null ? '' : v).trim();

  // ✅ normalize "-" / empty
  const cleanDash = (v) => {
    const s = toStr(v);
    if (!s) return '';
    if (s === '-') return '';
    return s;
  };

  // ✅ Date formatter รองรับ Date/ISO/dd/MM/yyyy/BE
  const fmtD = (d) => {
    if (!d) return '';
    if (d instanceof Date && !isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'dd/MM/yyyy');

    const s = toStr(d);

    // ISO: yyyy-MM-dd
    const mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (mIso) return `${mIso[3]}/${mIso[2]}/${mIso[1]}`;

    // dd/MM/yyyy or dd/MM/BBBB
    const mDMY = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (mDMY) {
      const dd = String(mDMY[1]).padStart(2, '0');
      const mm = String(mDMY[2]).padStart(2, '0');
      const yy = String(mDMY[3]);
      return `${dd}/${mm}/${yy}`;
    }

    return s;
  };

  // ✅ Time formatter รองรับ "9:00", "09:00", "09:00 AM", "9.00", "0900"
  const fmtT = (t) => {
    const s0 = toStr(t);
    if (!s0) return '';

    // ถ้ามี parseTimeSafe_ ให้ใช้ก่อน
    if (typeof parseTimeSafe_ === 'function') {
      const hhmm = toStr(parseTimeSafe_(s0));
      if (hhmm) return hhmm;
    }

    let s = s0.replace('.', ':').replace(/\s+/g, ' ').trim();

    // 0900 -> 09:00
    const m4 = s.match(/^(\d{2})(\d{2})$/);
    if (m4) return `${m4[1]}:${m4[2]}`;

    // 9:00 / 09:00
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];

    // fallback: ส่งคืนเดิม (แต่ถือว่าไม่ชัวร์)
    Logger.log('[buildRowForDataSheet] ⚠️ Unrecognized time format: ' + s0);
    return s0;
  };

  // ✅ Normalize carType ให้เป็นภาษาไทยมาตรฐาน + รองรับ array
  const normalizeCarType_ = (raw) => {
    let s = raw;
    if (Array.isArray(s)) s = s.join(', ');
    s = toStr(s);

    if (!s || s === '-') return '';

    const low = s.toLowerCase();
    const typeList =[];

    if (low.includes('van') || low.includes('ตู้') || low.includes('รถตู้') || low.includes('win')) typeList.push('รถตู้');
    if (low.includes('truck') || low.includes('กระบะ') || low.includes('บรรทุก') || low.includes('รถบรรทุก')) typeList.push('รถบรรทุก');

    return typeList.length ? Array.from(new Set(typeList)).join(', ') : s;
  };

  // ✅ key alias รองรับ headerIndex_ ที่ใช้ชื่อไม่เหมือนกัน
  const ALIAS = {
    name:['name', 'fullname', 'ชื่อ-สกุล'],
    status: ['status', 'สถานะ'],
    phone:['phone', 'tel', 'เบอร์โทร'],
    position:['position', 'ตำแหน่ง'],
    department:['department', 'org', 'dept', 'หน่วยงาน', 'ฝ่าย', 'ส่วนงาน', 'สังกัด'],
    email:['email', 'Email'],
    
    // 🍓[BERRY FIX] เพิ่ม Alias สำหรับ 2 คอลัมน์ใหม่
    workType: ['workType', 'jobType', 'ประเภทงาน'], 
    workName:['workName', 'projectName', 'ชื่อโครงการ/งาน', 'project', 'purpose', 'งาน/โครงการ'],
    
    destination:['destination', 'place', 'location', 'สถานที่'],
    carType:['carType', 'vehicleType', 'ประเภทรถ'],
    startDate:['startDate', 'วันเริ่มต้น', 'วันไป'],
    startTime:['startTime', 'เวลาเริ่มต้น', 'เวลาไป'],
    endDate:['endDate', 'วันสิ้นสุด', 'วันกลับ'],
    endTime: ['endTime', 'เวลาสิ้นสุด', 'เวลากลับ'],
    passengers:['passengers', 'people', 'จำนวนผู้ร่วมเดินทาง'],
    bookingId:['bookingId', 'id', 'Booking ID'],
    vehicleCount:['vehicleCount', 'carCount', 'จำนวนรถที่ต้องการ'],
    fileUrl:['fileUrl', 'file', 'File', 'ไฟล์แนบ'],
    reason:['reason', 'Reason', 'หมายเหตุ/เหตุผล', 'หมายเหตุ'],
    cancelReason:['cancelReason', 'CancelReason', 'เหตุผลยกเลิก'],
    vehicle:['vehicle', 'plate', 'เลขทะเบียนรถ'],
    driver:['driver', 'drivers', 'พนักงานขับรถ']
  };

  // ✅ set cell by canonical key with alias fallback
  const setCellSmart_ = (canonicalKey, value) => {
    const keys = ALIAS[canonicalKey] || [canonicalKey];

    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (idx[k] !== undefined && idx[k] !== -1) {
        row[idx[k]] = value;
        return true;
      }
    }
    return false;
  };

  // ✅ warn for missing required keys (แต่ไม่ throw เพื่อไม่ทำให้ระบบเดิมพัง)
  (function assertHeaderMap_() {
    const must =['bookingId', 'name', 'status', 'startDate', 'startTime', 'endDate', 'endTime'];
    const missing = must.filter(k => !setCellSmart_(k, row[idx[k] || 0])); 
  })();

  // ✅ ตรวจ missing แบบไม่มี side effect
  (function warnMissingHeaderKeys_() {
    const must =['bookingId', 'name', 'status', 'startDate', 'startTime', 'endDate', 'endTime'];
    const missing =[];

    must.forEach(k => {
      const keys = ALIAS[k] || [k];
      const okFound = keys.some(kk => idx[kk] !== undefined && idx[kk] !== -1);
      if (!okFound) missing.push(k);
    });

    if (missing.length) {
      Logger.log('[buildRowForDataSheet] ⚠️ Missing header map for keys: ' + missing.join(', '));
    }
  })();

  // ✅ Values
  const name = toStr(parsed.name);
  const status = toStr(parsed.status || 'pending');

  const phone = (typeof formatPhoneNumber_ === 'function')
    ? formatPhoneNumber_(parsed.phone)
    : toStr(parsed.phone);

  const dept = toStr(parsed.org || parsed.department);
  
  // 🍓 [BERRY FIX] ดึงค่าประเภทงานและชื่อโครงการจาก parsed payload ให้ถูกต้อง
  const jobType = toStr(parsed.workType || parsed.jobType);
  const projectName = toStr(parsed.workName || parsed.projectName || parsed.project || parsed.purpose);
  
  const place = toStr(parsed.place || parsed.destination);

  const carType = normalizeCarType_(parsed.carType || parsed.vehicleType);

  const startDate = fmtD(parsed.startDate);
  const startTime = fmtT(parsed.startTime);
  const endDate = fmtD(parsed.endDate || parsed.startDate);
  const endTime = fmtT(parsed.endTime);

  const passengers = toStr(parsed.passengers == null ? '1' : parsed.passengers);
  const bookingId = toStr(bookingIdFinal);

  const reason = toStr(parsed.reason);
  const cancelReason = toStr(parsed.cancelReason);

  const vehicleCount = toStr(parsed.vehicleCount || '1');

  // 💖[FILE LOGIC] หยอดลงคอลัมน์ File เสมอ ถ้าว่างให้ใส่ '-'
  let fileVal = cleanDash(parsed.fileUrl || parsed.file);
  if (!fileVal) fileVal = '-';

  // ✅ Write row
  setCellSmart_('name', name);
  setCellSmart_('status', status);
  setCellSmart_('phone', phone);
  setCellSmart_('position', toStr(parsed.position));
  setCellSmart_('department', dept);
  setCellSmart_('email', toStr(parsed.email));
  
  // 🍓 [BERRY FIX] เขียนค่าลง 2 คอลัมน์ใหม่
  setCellSmart_('workType', jobType);
  setCellSmart_('workName', projectName);
  
  setCellSmart_('destination', place);
  setCellSmart_('carType', carType);

  setCellSmart_('startDate', startDate);
  setCellSmart_('startTime', startTime);
  setCellSmart_('endDate', endDate);
  setCellSmart_('endTime', endTime);

  setCellSmart_('passengers', passengers);
  setCellSmart_('vehicleCount', vehicleCount);

  setCellSmart_('bookingId', bookingId);

  setCellSmart_('fileUrl', fileVal);

  setCellSmart_('reason', reason);
  setCellSmart_('cancelReason', cancelReason);

  // ✅ ล้างค่าคอลัมน์ที่ควรว่างตอนเริ่ม
  setCellSmart_('vehicle', '');
  setCellSmart_('driver', '');

  // ✅ HARDEN: ถ้า startTime/endTime ว่าง ให้ log เตือน (ช่วยตาม bug เวลา 00:00)
  if (!startTime) Logger.log('[buildRowForDataSheet] ⚠️ startTime empty → message may fallback to 00:00');
  if (!endTime) Logger.log('[buildRowForDataSheet] ⚠️ endTime empty → message may fallback to 00:00');

  return row;
}



function getBookingObjectById_(bookingId) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_MAIN_NAME);
  if (!sh) throw new Error('ไม่พบชีต Data');

  var id = String(bookingId || '').trim();
  if (!id) throw new Error('bookingId ว่าง');

  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2) return null;

  var headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(function(x){ return String(x||'').trim(); });
  var idCol = headers.indexOf('Booking ID') + 1;
  if (idCol < 1) throw new Error('ไม่พบคอลัมน์ Booking ID');

  var ids = sh.getRange(2, idCol, lr - 1, 1).getValues();
  var rowIndex = -1;
  for (var i = 0; i < ids.length; i++) {
    var v = String(ids[i][0] || '').trim();
    if (v === id) { rowIndex = i + 2; break; }
  }
  if (rowIndex < 0) return null;

  var row = sh.getRange(rowIndex, 1, 1, lc).getValues()[0];
  var obj = {};
  for (var c = 0; c < headers.length; c++) obj[headers[c] || ('C' + (c+1))] = row[c];

  obj.bookingId = obj.bookingId || obj['Booking ID'] || id;
  obj.name = obj.name || obj['ชื่อ-สกุล'];
  obj.phone = obj.phone || obj['เบอร์โทร'];
  obj.project = obj.project || obj['งาน/โครงการ'];
  obj.place = obj.place || obj['สถานที่'];
  obj.carType = obj.carType || obj['ประเภทรถ'];
  obj.plate = obj.plate || obj['เลขทะเบียนรถ'];
  obj.driver = obj.driver || obj['พนักงานขับรถ'];
  obj.startDate = obj.startDate || obj['วันเริ่มต้น'];
  obj.startTime = obj.startTime || obj['เวลาเริ่มต้น'];
  obj.endDate = obj.endDate || obj['วันสิ้นสุด'];
  obj.endTime = obj.endTime || obj['เวลาสิ้นสุด'];
  obj.passengers = obj.passengers || obj['จำนวนผู้ร่วมเดินทาง'];
  obj.vehicleCount = obj.vehicleCount || obj['จำนวนรถที่ต้องการ'];
  obj.status = obj.status || obj['สถานะ'];
  obj.reason = obj.reason || obj['Reason'];
  obj.cancelReason = obj.cancelReason || obj['CancelReason'];

  return obj;
}


// ===================== SETTINGS MANAGEMENT =====================
function readSettingKV_(key){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('setting');
  if(!sh) throw new Error("ไม่พบชีต setting");
  const lastRow = sh.getLastRow();
  const rows = sh.getRange(1,1,lastRow,2).getValues();
  for (let i=0;i<rows.length;i++){
    if (String(rows[i][0]).trim() === key) {
      return {row:i+1, val:String(rows[i][1]||'')};
    }
  }
  return {row:-1, val:''};
}

function parseBoolMap_(s){
  const m={};
  String(s||'').split(',').forEach(pair=>{
    const [k,v] = pair.split(':').map(x=>String(x||'').trim());
    if(k) m[k]= (String(v).toLowerCase()==='true' || v==='1');
  });
  return m;
}

// ===================== ADD MISSING FUNCTIONS =====================
function logEvent_(type, bookingId, obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Log') || ss.insertSheet('Log');
  sh.appendRow([new Date(), String(type||''), String(bookingId||''), JSON.stringify(obj||{},null,0)]);
}

function logActivity_(action, id, message) {
  logEvent_(action, id, { message: message });
}

function getTZ() {
  return VB_CFG.TZ || 'Asia/Bangkok';
}


function isOverlapping_(startA, endA, startB, endB) {
  // เงื่อนไขคือ:
  // A เริ่มก่อน B จบ และ A จบหลัง B เริ่ม
  return startA < endB && endA > startB;
}

function getAvailableVehicles(payload) {
  try {
    if (!payload) throw new Error('ข้อมูลวันเวลาไม่ครบถ้วน');
    const d1 = payload.startDate || payload.date;
    const d2 = payload.endDate || payload.date || d1;
    const excludeId = String(payload.bookingId || '').trim();
    const clean = (s) => String(s || '').replace(/[\u200B-\u200D\uFEFF]/g, '').trim();

    const nd1 = parseDateToISO_(d1);
    const nd2 = parseDateToISO_(d2);
    const startTime24 = parseTimeSafe_(payload.startTime);
    const endTime24 = parseTimeSafe_(payload.endTime);
    const reqStart = parseDateTime_(nd1, startTime24);
    const reqEnd = parseDateTime_(nd2, endTime24);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = headerIndex_(headers);
    const values = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), headers.length).getValues();

    const busyPlatesMap = {}; 
    const busyDriversMap = {}; 

    for (const row of values) {
      const rowId = clean(row[idx.bookingId]);
      if (excludeId && rowId === excludeId) continue; 

      const status = getStatusKeySafe_(row[idx.status]);
      if (status === 'approved' || status === 'pending') {
        const rStartISO = parseDateToISO_(row[idx.startDate]);
        const bStart = parseDateTime_(rStartISO, parseTimeSafe_(row[idx.startTime]));
        const bEnd = parseDateTime_(parseDateToISO_(row[idx.endDate]) || rStartISO, parseTimeSafe_(row[idx.endTime]));

        if (bStart && bEnd && isOverlapping_(reqStart, reqEnd, bStart, bEnd)) {
          const job = clean(row[idx.project] || row[idx.purpose] || 'ติดงาน');
          const pCell = clean(row[idx.vehicle]);
          if (pCell) pCell.split(',').forEach(p => { if (clean(p)) busyPlatesMap[clean(p)] = job; });
          const dCell = clean(row[idx.driver]);
          if (dCell) dCell.split(',').forEach(d => { if (clean(d)) busyDriversMap[clean(d)] = job; });
        }
      }
    }

    const vStatusMap = parseBoolMap_(readSettingKV_('VehicleAvailability').val);
    const vehiclesRes = getAllVehiclePlatesFromSettings();
    const vehicleStatus = (vehiclesRes.ok ? vehiclesRes.all : []).map(v => {
      const p = clean(v.plate);
      const isM = vStatusMap[p] === false;
      if (isM) return { ...v, available: false, badge: 'งดใช้/ซ่อมบำรุง 🔧' };
      if (busyPlatesMap[p]) return { ...v, available: false, badge: busyPlatesMap[p] };
      return { ...v, available: true, badge: 'ว่าง' };
    });

    const driversRes = getDriversFromAdmin_();
    const driverList = (driversRes.ok ? driversRes.drivers : []).map(d => {
      const name = clean(typeof d === 'object' ? d.name : d);
      const busyJob = busyDriversMap[name];
      const inactive = (typeof d === 'object' && d.active === false);
      return {
        name: name,
        active: !inactive,
        busyBadge: busyJob ? `ติดงาน: ${busyJob}` : (inactive ? 'พักงาน' : ''),
        isBusy: !!(busyJob || inactive) // 🍓 บังคับส่งค่า Boolean จริงๆ ไปเลยค่ะ
      };
    });

    return { ok: true, vehicles: vehicleStatus, drivers: driverList };
  } catch (e) { return { ok: false, error: e.message }; }
}

function getTimelineData(payload) {
  try {
    if (!payload || !payload.dateISO) throw new Error('dateISO is required');
    const dateISO = String(payload.dateISO);
    const filterPlate = String(payload.plate || '').trim();

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_MAIN_NAME);
    if (!sh) throw new Error("Sheet Data not found");

    const rng = sh.getDataRange().getValues();
    if (rng.length < 2) {
      return { ok: true, dateISO: dateISO, plates:[], bookings: [] };
    }

    const header = rng[0].map(x => String(x || '').trim());
    const idx = headerIndex_(header);
    
    // ตรวจสอบคอลัมน์ที่จำเป็นพื้นฐาน
    if (idx.startDate === undefined || idx.startTime === undefined || idx.status === undefined) {
      throw new Error("Timeline failed: Missing required columns (startDate, startTime, status).");
    }

    // วันเป้าหมาย (00:00 - 23:59)
    const dayStart = parseDateTime_(dateISO, "00:00");
    const dayEnd = parseDateTime_(dateISO, "23:59");

    function overlaps(d1, t1, d2, t2) {
      const iso1 = parseDateToISO_(d1) || String(d1 || '');
      const iso2 = parseDateToISO_(d2) || iso1;
      
      const sTimeSafe = parseTimeSafe_(t1 || '00:00');
      const eTimeSafe = parseTimeSafe_(t2 || t1 || '00:00');

      const s = parseDateTime_(iso1, sTimeSafe);
      const e = parseDateTime_(iso2, eTimeSafe);
      if (!s || !e) return false;
      
      // ถ้าเวลาเริ่มและจบเท่ากัน ให้บวก 1 นาทีเพื่อให้คำนวณช่วงเวลาได้
      if (e.getTime() === s.getTime()) {
        e.setMinutes(e.getMinutes() + 1);
      }
      
      return (e >= dayStart && s <= dayEnd);
    }

    const list =[];
    for (let r = 1; r < rng.length; r++) {
      const row = rng[r];
      const st = getStatusKeySafe_(row[idx.status]);
      
      // 💖[BERRY UPDATE] ปลดล็อก: ดึงทุกสถานะ (Approved, Pending, Rejected, Cancelled)
      // เพื่อให้รายงานสรุปประจำวัน (Daily Report) ได้ข้อมูลครบถ้วน
      
      const plate = String(idx.vehicle >= 0 ? row[idx.vehicle] : '').trim();
      if (filterPlate && plate !== filterPlate) continue;

      const sDate = (idx.startDate >= 0 ? row[idx.startDate] : '');
      const sTime = (idx.startTime >= 0 ? row[idx.startTime] : '');
      const eDate = (idx.endDate >= 0 ? row[idx.endDate] : sDate);
      const eTime = (idx.endTime >= 0 ? row[idx.endTime] : sTime);

      if (!overlaps(sDate, sTime, eDate, eTime)) continue;
      
      const driver = String(idx.driver >= 0 ? row[idx.driver] : '').trim();
      
      list.push({
        bookingId: String(idx.bookingId >= 0 ? row[idx.bookingId] : '').trim(),
        status: st,
        plate: plate,
        carType: String(idx.carType >= 0 ? row[idx.carType] : '').trim(), 
        requestedVehicle: String(idx.requestedVehicle >= 0 ? row[idx.requestedVehicle] : '').trim(),
        name: String(idx.name >= 0 ? row[idx.name] : '').trim(),
        // 🍓 [BERRY FIX] ดึงจาก Key ใหม่ และมี Fallback
        purpose: String(idx.workName >= 0 ? row[idx.workName] : (idx.project >= 0 ? row[idx.project] : '')).trim(),
        destination: String(idx.destination >= 0 ? row[idx.destination] : '').trim(),
        startDate: parseDateToISO_(sDate) || dateISO,
        startTime: parseTimeSafe_(sTime),
        endDate: parseDateToISO_(eDate) || dateISO,
        endTime: parseTimeSafe_(eTime),
        driver: driver,
        passengers: String(idx.passengers >= 0 ? row[idx.passengers] : '').trim(),
        
        // 💖 [BERRY ADD] เพิ่มฟิลด์ให้ครบถ้วนสำหรับ UI และ Report
        phone: String(idx.phone >= 0 ? row[idx.phone] : '').trim(),
        org: String(idx.department >= 0 ? row[idx.department] : '').trim(),
        fileUrl: String(idx.fileUrl >= 0 ? row[idx.fileUrl] : '').trim(), // File
        reason: String(idx.reason >= 0 ? row[idx.reason] : '').trim(), // Note/Reason
        cancelReason: String(idx.cancelReason >= 0 ? row[idx.cancelReason] : '').trim() // Cancel Reason
      });
    }

    // รายการทะเบียนรถสำหรับ dropdown (เหมือนเดิม)
    let plates =[];
    try {
      const pRes = getAllVehiclePlatesFromSettings();
      if (pRes && pRes.ok && pRes.all) {
        plates = pRes.all.map(v => v.plate);
      }
    } catch (e) {
      const set = {};
      for (let r2 = 1; r2 < rng.length; r2++) {
        const p = String(idx.vehicle >= 0 ? rng[r2][idx.vehicle] : '').trim();
        if (p) set[p] = 1;
      }
      plates = Object.keys(set);
    }

    return { ok: true, dateISO: dateISO, plates: plates, bookings: list };
  } catch (err) {
    Logger.log("getTimelineData Error: " + err.stack);
    return { ok: false, error: err.message };
  }
}
// ===================== EXPORT FUNCTIONS =====================
function getWebAppInitialData() {
  return getMainData_();
}

function ping() {
  return { ok:true, ts: new Date().toISOString() };
}

function submitForm(payload) {
  return createBookingAndBroadcast_(payload);
}

// ===================== AUTH & SESSION =====================
function logoutUser() {
  // ใน GAS Web App ไม่มี Session ฝั่ง Server ถาวร 
  // แต่เราต้องมีฟังก์ชันนี้เพื่อให้ Client เรียกแล้วไม่ Error
  // และใช้สำหรับเคลียร์ Cache ฝั่ง Server ที่เกี่ยวข้องกับ User คนนั้น (ถ้ามี)
  return { ok: true, message: 'Logged out from server context' };
}

// ===================== ADMIN FUNCTIONS =====================
function checkUserSession() {
  try {
    return { ok: true, data: { isLoggedIn: false, role: 'Guest' } };
  } catch (e) {
    Logger.log(`Error in checkUserSession: ${e.message}`);
    return { ok: false, error: e.message };
  }
}

function verifyAdminLogin(formData) {
  try {
    const { username, password } = formData;
    
    if (!username || !password) {
      return { ok: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
    }

    const userData = getUserDataByUsername_(username);
    if (!userData) {
      return { ok: false, error: 'ไม่พบชื่อผู้ใช้นี้ในระบบ' };
    }

    const isPasswordMatch = checkPassword_(username, password);
    if (isPasswordMatch) {
      return {
        ok: true,
        data: {
          isLoggedIn: true,
          username: username,
          role: userData.Role,
          displayName: userData.DisplayName
        }
      };
    } else {
      return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
    }
  } catch (e) {
    Logger.log(`Error in verifyAdminLogin: ${e.message}`);
    return { ok: false, error: e.message };
  }
}

function apiUserCancelBooking(payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: "ระบบยุ่ง กรุณาลองใหม่ค่ะ" };

  try {
    var id = String(payload.bookingId).trim();
    var inputPhoneRaw = String(payload.phone || '').replace(/\D/g, '');
    var inputPhoneCheck = inputPhoneRaw.substring(inputPhoneRaw.length - 9); 

    var reason = String(payload.reason || '').trim();
    var noTelegram = payload.noTelegram === true;

    if (!id || inputPhoneRaw.length < 9 || !reason) throw new Error("ข้อมูลไม่ครบ");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data");
    if (!sheet) throw new Error("ไม่พบชีต Data");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var map = headerIndex_(headers);

    if (map.bookingId === undefined || map.phone === undefined || map.status === undefined) {
      throw new Error("โครงสร้างตารางไม่ถูกต้อง");
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var rowData = null;

    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][map.bookingId]).trim() === id) {
        var rowPhoneRaw = String(data[i][map.phone]).replace(/\D/g, '');
        var rowPhoneCheck = rowPhoneRaw.substring(rowPhoneRaw.length - 9);

        if (rowPhoneCheck !== inputPhoneCheck) {
          throw new Error("เบอร์โทรศัพท์ไม่ตรงกับข้อมูลการจอง");
        }
        
        var currentStatus = String(data[i][map.status] || '').toLowerCase();
        if (currentStatus === 'cancelled' || currentStatus === 'rejected') {
          throw new Error("รายการนี้ถูกยกเลิกหรือปฏิเสธไปแล้ว");
        }

        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) throw new Error("ไม่พบรายการจอง หรือเบอร์โทรผิด");

    // บันทึกลง Sheet
    sheet.getRange(rowIndex, map.status + 1).setValue('cancelled');
    if (map.cancelReason !== undefined) {
      sheet.getRange(rowIndex, map.cancelReason + 1).setValue(reason);
    } else if (map.reason !== undefined) {
      sheet.getRange(rowIndex, map.reason + 1).setValue("User Cancel: " + reason);
    }

    SpreadsheetApp.flush();
    try { cacheDelete_('mainDataCache_v13_BerryFix'); } catch(e) {}

    // 📢 ส่วนแจ้งเตือน Telegram
    if (!noTelegram) {
      try {
        var notifyData = {};
        for (var k = 0; k < headers.length; k++) {
          notifyData[headers[k]] = rowData[k];
        }
        
        // 🍓 Berry Edit: อัปเดตข้อมูลให้สดใหม่ก่อนส่งไป Build Message
        notifyData['สถานะ'] = 'cancelled';
        notifyData['status'] = 'cancelled'; // เพิ่มเผื่อไว้
        notifyData['Reason'] = reason;     // เพื่อให้ไปโผล่ในบรรทัดหมายเหตุของ Telegram
        notifyData['Booking ID'] = id;

        // เรียกใช้ตัวส่งที่เราทำไว้ (ส่งจริง)
        sendTelegramNotify(notifyData, false); 
      } catch (ex) {
        Logger.log("Notify Error: " + ex.message);
      }
    }

    try { logActivity_('USER_CANCEL', id, { status: 'cancelled', reason: reason }); } catch(e) {}

    return { ok: true, id: id, status: 'cancelled' };

  } catch (e) {
    Logger.log("Cancel Error: " + e.stack);
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getUserDataByUsername_(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admin"); 
    if (!sheet) {
      throw new Error("ไม่พบชีต 'Admin' ค่ะ!");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase()); 
    
    const userCol = headers.indexOf("username");
    const nameCol = headers.indexOf("name");        
    const roleCol = headers.indexOf("role");        

    if (userCol === -1 || roleCol === -1 || nameCol === -1) {
      Logger.log(`❌ [Helper Error] ไม่พบคอลัมน์ที่ต้องการค่ะ!`);
      Logger.log(`   (กำลังมองหา: 'username', 'name', 'role')`);
      Logger.log(`   (ที่เจอในชีตคือ: [${headers.join(', ')}])`);
      return null;
    }

    const cleanString = (str) => {
      if (typeof str !== 'string') str = str.toString();
      return str.replace(/[\s\u200B-\u200D\uFEFF]/g, '').toLowerCase();
    };

    const safeUsername = cleanString(username);

    for (let i = 1; i < data.length; i++) {
      const sheetUsername = cleanString(data[i][userCol]);
      if (sheetUsername === safeUsername) {
        return { 
          Role: data[i][roleCol], 
          DisplayName: data[i][nameCol] 
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log(`❌ [Helper Error] เกิดข้อผิดพลาด: ${e.message}`);
    return null;
  }
}

function checkPassword_(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admin");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const userCol = headers.indexOf("username");
    const passCol = headers.indexOf("password");

    if (passCol === -1) {
      Logger.log("❌ [checkPassword_] ไม่พบคอลัมน์ 'password' ในชีต Admin!");
      return false;
    }

    const cleanString = (str) => {
      if (typeof str !== 'string') str = str.toString();
      return str.replace(/[\s\u200B-\u200D\uFEFF]/g, '').toLowerCase();
    };
    
    const cleanUsername = cleanString(username);

    for (let i = 1; i < data.length; i++) {
      const sheetUsername = cleanString(data[i][userCol]);
      if (sheetUsername === cleanUsername) {
        return data[i][passCol].toString() === password.toString();
      }
    }
    return false;
  } catch (e) {
     Logger.log(`❌ [checkPassword_] Error: ${e.message}`);
     return false;
  }
}

// ===================== ANCHOR: File Upload & Form Data (NEW) =====================
/* [ANCHOR: Insurance & Maintenance Services (Full Safe Version)] */

// --- Helper 1: ค้นหาแถวว่างถัดไป (คงไว้เผื่อใช้) ---
function _getNextRow_(sh) {
  return sh.getLastRow() + 1;
}

// --- Helper 2: Format วันที่ไทย (พ.ศ.) ---
function _fmtThaiDateBE(d) {
  if (!d) return '-';
  try {
    var dateObj = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dateObj.getTime())) return '-';
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var y = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'), 10);
    var be = (y < 2400) ? y + 543 : y;
    return Utilities.formatDate(dateObj, tz, 'dd/MM/') + be;
  } catch (e) { return '-'; }
}

// --- Helper 3: Save Base64 File (จำเป็นสำหรับ Maintenance) ---
function _saveBase64File_(base64Data, filename) {
  if (!base64Data || !filename) return null;
  try {
    var parts = base64Data.split(',');
    var mimeType = parts[0].match(/:(.*?);/)[1];
    var bytes = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    
    // หาโฟลเดอร์ (หรือสร้างใหม่ถ้าไม่มี)
    var folders = DriveApp.getFoldersByName("V-Berry Uploads");
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("V-Berry Uploads");
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log("Save File Error: " + e.message);
    return "";
  }
}

// 1. บันทึกข้อมูลประกันภัย (เพิ่ม LockService)
function apiSaveInsurance(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวสูงสุด 10 วินาที
    
    if (!payload || !payload.plate || !payload.company || !payload.endDate) {
      throw new Error('ข้อมูลไม่ครบถ้วน (ทะเบียนรถ, บริษัท, วันสิ้นสุด)');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Insurance');
    if (!sh) throw new Error("ไม่พบชีต 'Insurance'");

    const id = 'INS-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmm');
    const end = new Date(payload.endDate);
    const now = new Date();
    const status = (end < now) ? 'Expired' : 'Active';

    // Map ตาม Header: InsuranceID | Plate | Provider | PolicyNumber | StartDate | EndDate | Status | Remark | Cost
    const newRow = [
      id,
      String(payload.plate || ''),
      String(payload.company || ''),
      String(payload.policyNo || ''),
      payload.startDate ? new Date(payload.startDate) : null,
      end,
      status,
      String(payload.note || ''),
      payload.cost ? Number(payload.cost) : 0
    ];
    
    sh.appendRow(newRow); // ใช้ appendRow ปลอดภัยกว่า
    
    return { ok: true, id: id };
  } catch (e) {
    Logger.log('❌ apiSaveInsurance ERROR: ' + e.stack);
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// 2. ดึงข้อมูลประกันภัย
function apiGetInsuranceHistory() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Insurance');
    if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };
    
    const lastRow = sh.getLastRow();
    // อ่านข้อมูล A ถึง I (9 คอลัมน์)
    const vals = sh.getRange(2, 1, lastRow - 1, 9).getValues();
    
    const data = vals.map(r => {
      const endDate = r[5] ? new Date(r[5]) : null;
      let realStatus = 'active';
      if (endDate) {
         const now = new Date();
         const diff = (endDate - now) / (1000 * 60 * 60 * 24);
         if (diff < 0) realStatus = 'expired';
         else if (diff < 30) realStatus = 'warning';
      }

      return {
        id: String(r[0]),
        plate: String(r[1]),
        company: String(r[2]),
        policyNo: String(r[3]),
        startDate: _fmtThaiDateBE(r[4] ? new Date(r[4]) : null),
        endDate: _fmtThaiDateBE(endDate),
        status: realStatus,        
        remark: String(r[7]),
        cost: Number(r[8] || 0)
      };
    }).reverse();

    return { ok: true, data: data };
  } catch (e) { 
    return { ok: false, error: e.message }; 
  }
}

// 3. บันทึกข้อมูลซ่อมบำรุง (เพิ่ม LockService + File Upload)
function apiSaveMaintenance(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!payload || !payload.plate || !payload.topic || !payload.cost || !payload.startDate) {
      throw new Error('ข้อมูลไม่ครบถ้วน (ทะเบียนรถ, รายการซ่อม, ค่าใช้จ่าย, วันที่ซ่อม)');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Maintenance');
    if (!sh) throw new Error("ไม่พบชีต 'Maintenance'");
    
    const id = 'MAINT-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmm');
    
    // จัดการไฟล์แนบ
    let fileUrl = '';
    if (payload.fileData) {
      fileUrl = _saveBase64File_(payload.fileData, payload.fileName || 'maint_upload.png');
    }
    
    // Map ตาม Header
    const newRow = [
      id,
      String(payload.plate || ''),
      'ซ่อมบำรุงทั่วไป',
      String(payload.topic || ''),
      new Date(payload.startDate),
      payload.cost ? Number(payload.cost) : 0,
      payload.nextDate ? new Date(payload.nextDate) : null,
      String(payload.note || ''),
      '1',
      new Date(),
      fileUrl
    ];
    
    sh.appendRow(newRow);
    
    return { ok: true, id: id, fileUrl: fileUrl };
  } catch (e) {
    Logger.log('❌ apiSaveMaintenance ERROR: ' + e.stack);
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// 4. ดึงข้อมูลซ่อมบำรุง
function apiGetMaintenanceHistory() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Maintenance');
    if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };
    
    const lastRow = sh.getLastRow();
    // อ่านข้อมูล A ถึง K (11 คอลัมน์)
    const vals = sh.getRange(2, 1, lastRow - 1, 11).getValues();
    
    const data = vals.map(r => {
      return {
        id: String(r[0]),
        plate: String(r[1]),
        topic: String(r[3]),
        date: _fmtThaiDateBE(r[4] ? new Date(r[4]) : null),
        cost: Number(r[5] || 0),
        nextDate: _fmtThaiDateBE(r[6] ? new Date(r[6]) : null),
        fileUrl: String(r[10] || '')
      };
    }).reverse();
    
    return { ok: true, data: data };
  } catch (e) { return { ok: false, error: e.message }; }
}

function saveInsuranceRecord(form) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Insurance');
    
    // รวม Driver เข้าไปใน Remark (เพราะชีตนี้ไม่มีคอลัมน์ Driver แยก)
    let finalRemark = form.note || form.remark || '';
    if (form.driver) {
      finalRemark = (finalRemark ? finalRemark + ' ' : '') + '(ผู้บันทึก: ' + form.driver + ')';
    }

    // ✅ เรียงข้อมูลให้ตรงกับชีต Insurance (A-I)
    // [A:ID, B:Plate, C:Provider, D:PolicyNumber, E:StartDate, F:EndDate, G:Status, H:Remark, I:Cost]
    
    const rowData = [
      Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy, HH:mm:ss"), // A
      form.plate,                        // B
      form.company,                      // C
      form.policyNo || form.policyNumber,// D
      form.startDate,                    // E
      form.endDate,                      // F
      'Active',                          // G
      finalRemark,                       // H
      form.cost                          // I: Cost (✅ แก้แล้ว: ใส่ Cost ให้ถูกช่องท้ายสุด)
    ];

    sh.appendRow(rowData);
    return { ok: true, message: 'บันทึกข้อมูลประกันภัยเรียบร้อย' };

  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

function listInsuranceRecords() {
  try {
    const sheetName = (typeof SHEET_INSURANCE !== 'undefined') ? SHEET_INSURANCE : 'Insurance';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) return { ok: true, data:[] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, data:[] };

    const rows = data.slice(1);

    const result = rows.map(row => {
      let sDate = row[4] ? _fmtThaiDateBE(row[4]) : '';
      let eDate = row[5] ? _fmtThaiDateBE(row[5]) : '';

      if (sDate === '-') sDate = '';
      if (eDate === '-') eDate = '';

      let realStatus = 'active';
      if (row[5]) {
         const now = new Date();
         const due = new Date(row[5]);
         const diff = (due - now) / (1000 * 60 * 60 * 24);
         if (diff < 0) realStatus = 'expired';
         else if (diff < 30) realStatus = 'warning';
      }

      return {
        timestamp: String(row[0] || ''), // 🍓 BERRY FIX: บังคับเป็น String ป้องกัน Payload แครช
        vehicle:   String(row[1] || ''),
        company:   String(row[2] || ''),
        policyNo:  String(row[3] || ''),
        startDate: sDate,
        endDate:   eDate,
        status:    realStatus,
        remark:    String(row[7] || ''),
        cost:      Number(row[8] || 0)
      };
    });

    return { ok: true, data: result.reverse() };

  } catch (e) {
    Logger.log("List Ins Error: " + e.toString());
    return { ok: false, message: e.toString(), data:[] };
  }
}

/* [ANCHOR: Server - listMaintenanceRecords] */
function listMaintenanceRecords() {
  try {
    const sheetName = (typeof SHEET_MAINTENANCE !== 'undefined') ? SHEET_MAINTENANCE : 'Maintenance';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) return { ok: true, data:[] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, data:[] };

    const rows = data.slice(1);
    const result = rows.map(row => ({
      timestamp: String(row[0] || ''), // 🍓 BERRY FIX: บังคับเป็น String ป้องกัน Payload แครช
      vehicle:   String(row[1] || ''),
      date:      row[2] ? _fmtThaiDateBE(row[2]) : '',
      type:      String(row[3] || ''),
      cost:      Number(row[4] || 0),
      odometer:  String(row[5] || ''),
      location:  String(row[6] || ''),
      remark:    String(row[7] || ''),
      fileUrl:   String(row[8] || '')
    }));

    return { ok: true, data: result.reverse() };

  } catch (e) {
    Logger.log("List Maint Error: " + e.toString());
    return { ok: false, message: e.toString(), data:[] };
  }
}



function saveMaintenanceRecord(form, fileData) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Maintenance');

    // จัดการไฟล์
    let fileUrl = '';
    if (fileData && fileData.data && fileData.fileName) {
       try {
         const folderName = 'Maintenance_Slip';
         const folder = DriveApp.getFoldersByName(folderName).hasNext() ? DriveApp.getFoldersByName(folderName).next() : DriveApp.createFolder(folderName);
         const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType || 'image/jpeg', fileData.fileName);
         const file = folder.createFile(blob);
         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         fileUrl = file.getUrl();
       } catch (err) { Logger.log("Maint Upload Error: " + err); }
    }

    // รวม Driver เข้าไปใน Remark (เพราะชีตนี้ไม่มีคอลัมน์ Driver แยก)
    let finalRemark = form.remark || '';
    if (form.driver) {
      finalRemark = (finalRemark ? finalRemark + ' ' : '') + '(ผู้บันทึก: ' + form.driver + ')';
    }

    // ✅ เรียงข้อมูลให้ตรงกับชีต Maintenance (A-I)
    // [A:TimeStamp, B:Vehicle, C:Date, D:Type, E:Cost, F:Odometer, G:Location, H:Remark, I:File]
    
    const rowData = [
      Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy, HH:mm:ss"), // A
      form.vehicle || form.plate,        // B
      form.service_date,                 // C
      form.service_type || form.topic,   // D
      form.cost,                         // E
      form.odometer || '',               // F: Odometer (✅ แก้แล้ว: รับค่าจากฟอร์มมาใส่)
      form.location,                     // G
      finalRemark,                       // H
      fileUrl                            // I
    ];

    sh.appendRow(rowData);
    return { ok: true, message: 'บันทึกข้อมูลซ่อมบำรุงเรียบร้อย', fileUrl: fileUrl };

  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

/* ===================== ANCHOR: Fuel Management (Fixed Columns) ===================== */
// Helper: สร้าง ID แบบสุ่ม
function _generateFuelLogId() {
  return 'FL-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmmss');
}

// Helper: บันทึกไฟล์ Base64 ลง Drive (ปรับปรุงใหม่ รองรับ PNG/JPG อัตโนมัติ)
function _saveBase64File_(base64Data, fileName) {
  try {
    if (!base64Data) return '';
    
    // แปลง Base64 เป็น Blob
    var decoded = Utilities.base64Decode(base64Data);
    
    // ⚠️ แก้ไข: ไม่ Fix MimeType เป็น JPEG เพื่อให้รองรับ PNG จาก Test Script ได้
    // ให้ Google Drive ตรวจจับจากนามสกุลไฟล์เอง หรือระบุ null
    var blob = Utilities.newBlob(decoded, null, fileName); 
    
    // บันทึกลง Drive
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
    
  } catch (e) {
    Logger.log("⚠️ Error Saving File: " + e.toString());
    return ''; 
  }
}

function apiSaveFuel(form) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Fuel'); // ตรวจสอบชื่อชีต
    
    // จัดการไฟล์แนบ (ถ้ามี)
    var fileUrl = '';
    if (form.fileData && form.fileName) {
       try {
         var folderName = 'Fuel_Receipts';
         var folder;
         var folders = DriveApp.getFoldersByName(folderName);
         if (folders.hasNext()) { folder = folders.next(); } 
         else { folder = DriveApp.createFolder(folderName); }
         
         var blob = Utilities.newBlob(Utilities.base64Decode(form.fileData), form.mimeType, form.fileName);
         var file = folder.createFile(blob);
         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         fileUrl = file.getUrl();
       } catch (err) { Logger.log("Fuel Upload Error: " + err); }
    }

    // คำนวณค่าต่างๆ
    var dist = (parseFloat(form.endMileage) || 0) - (parseFloat(form.startMileage) || 0);
    var used = (parseFloat(form.fuelPercentStart) || 0) - (parseFloat(form.fuelPercentEnd) || 0);
    var month = Utilities.formatDate(new Date(), "GMT+7", "MM/yyyy");
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy, HH:mm:ss");
    var logId = 'FL-' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");

    // ✅ เรียงข้อมูลให้ตรงกับชีต Fuel (A-P)
    // [A:LogID, B:BookingID, C:Plate, D:StartMile, E:EndMile, F:Start%, G:End%, H:Remark, I:Driver, J:Time, K:Dist, L:Used%, M:Month, N:File, O:Liters, P:Cost]
    
    var rowData = [
      logId,                       // A: FuelLogID
      form.project || '',          // B: BookingID (✅ แก้แล้ว: เอาชื่อโครงการมาใส่แทน)
      form.plate,                  // C: Plate
      form.startMileage,           // D: StartMileage
      form.endMileage,             // E: EndMileage
      form.fuelPercentStart,       // F: StartFuelLevel
      form.fuelPercentEnd,         // G: EndFuelLevel
      form.remark,                 // H: Remark
      form.driver,                 // I: Driver
      timestamp,                   // J: Timestamp
      dist,                        // K: Distance
      used,                        // L: FuelUsedPercent
      month,                       // M: Month
      fileUrl,                     // N: Receipt URL
      form.liters,                 // O: Liters
      form.cost                    // P: Cost
    ];

    sh.appendRow(rowData);
    return { ok: true, message: 'บันทึกข้อมูลเรียบร้อย', fileUrl: fileUrl };

  } catch (e) {
    return { ok: false, message: e.toString(), error: e.toString() };
  }
}

/* [ANCHOR: API Get Fuel History (Real Data + Thai Date)] */
function apiGetFuelHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Fuel'); // ✅ ใช้ชื่อ 'Fuel' โดยตรง
    if (!sh) return { ok: true, data: [] };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };

    // อ่านข้อมูล A ถึง P (16 คอลัมน์)
    var vals = sh.getRange(2, 1, lastRow - 1, 16).getValues();
    
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

    var data = vals.filter(function(r) {
      // ต้องมี ID (Col A) และ ทะเบียนรถ (Col C)
      return String(r[0]).trim() && String(r[2]).trim(); 
    }).map(function(r) {
      // จัดการวันที่ (Col J = Index 9)
      var rawDate = r[9];
      var dateObj = null;
      var dateStr = '-';
      
      if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        dateObj = rawDate;
      } else if (rawDate) {
        var parsed = new Date(rawDate);
        if (!isNaN(parsed.getTime())) dateObj = parsed;
      }

      // แปลงเป็น พ.ศ.
      if (dateObj) {
        var y = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'));
        var thYear = (y < 2400) ? y + 543 : y;
        dateStr = Utilities.formatDate(dateObj, tz, 'dd/MM/') + thYear + ' ' + Utilities.formatDate(dateObj, tz, 'HH:mm') + ' น.';
      }

      return {
        id: String(r[0]),
        bookingId: String(r[1]),
        plate: String(r[2]),
        driver: String(r[8]),        // Col I
        timestamp: dateObj ? dateObj.getTime() : 0, 
        dateDisplay: dateStr,        // ✅ วันที่แบบไทยพร้อมโชว์
        liters: Number(r[14] || 0),  // Col O
        cost: Number(r[15] || 0),    // Col P
        fileUrl: String(r[13] || '') // Col N
      };
    });

    // เรียงจากใหม่ไปเก่า
    data.sort(function(a, b) { return b.timestamp - a.timestamp; });

    return { ok: true, data: data }; // ✅ ส่งกลับรูปแบบ Object

  } catch (e) {
    Logger.log('apiGetFuelHistory Error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/* [ANCHOR: Dashboard Fuel Level Widget - Server] */
function getDashboardFuelLevels() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Fuel');
    if (!sh) return { ok: true, data: [] };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };

    var vals = sh.getRange(2, 1, lastRow - 1, 16).getValues();
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

    // Group by Plate → เลือก record ล่าสุดจาก Timestamp
    var plateMap = {};

    for (var i = 0; i < vals.length; i++) {
      var plate = String(vals[i][2] || '').trim();   // Col C: Plate
      if (!plate) continue;

      var rawEnd = vals[i][6];                        // Col G: EndFuelLevel
      var rawTs  = vals[i][9];                        // Col J: Timestamp

      // Parse timestamp
      var tsMs = 0;
      var dateObj = null;
      if (rawTs instanceof Date && !isNaN(rawTs.getTime())) {
        dateObj = rawTs;
        tsMs = rawTs.getTime();
      } else if (rawTs) {
        var parsed = new Date(rawTs);
        if (!isNaN(parsed.getTime())) {
          dateObj = parsed;
          tsMs = parsed.getTime();
        }
      }

      // เลือก record ที่ Timestamp ใหม่สุดต่อ Plate
      if (!plateMap[plate] || tsMs > plateMap[plate].tsMs) {
        // Sanitize EndFuelLevel
        var endLevel = parseFloat(rawEnd);
        if (isNaN(endLevel) || endLevel === null || endLevel === undefined) {
          endLevel = null;
        } else {
          endLevel = Math.round(endLevel);
        }

        // แปลงวันที่เป็น พ.ศ.
        var dateDisplay = '-';
        if (dateObj) {
          var y = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'));
          var thYear = (y < 2400) ? y + 543 : y;
          dateDisplay = Utilities.formatDate(dateObj, tz, 'dd/MM/') + thYear + ' ' + Utilities.formatDate(dateObj, tz, 'HH:mm') + ' น.';
        }

        plateMap[plate] = {
          plate: plate,
          endFuelLevel: endLevel,
          timestamp: tsMs,
          dateDisplay: dateDisplay
        };
      }
    }

    // แปลง Map → Array เรียงตาม Plate
    var result = [];
    for (var p in plateMap) {
      if (plateMap.hasOwnProperty(p)) {
        result.push(plateMap[p]);
      }
    }
    result.sort(function(a, b) {
      return a.plate < b.plate ? -1 : (a.plate > b.plate ? 1 : 0);
    });

    return { ok: true, data: result };

  } catch (e) {
    Logger.log('getDashboardFuelLevels Error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ===================== ANCHOR: Data Fetchers (สำหรับดึงมาแสดงหน้าเว็บ) =====================
/* [ANCHOR: Date Helper (Standard - Fixed Time)] */
function _fmtThaiDateTime(d, tStr) {
  var tz = Session.getScriptTimeZone(); // ใช้ TZ ของ Script (Asia/Bangkok)
  
  // --- 1. จัดการวันที่ (Date) ---
  var datePart = '-';
  if (d) {
    // ถ้าเป็น Date Object หรือ String ที่แปลงได้
    var dateObj = (d instanceof Date) ? d : new Date(d);
    if (!isNaN(dateObj.getTime())) {
      var y = parseInt(Utilities.formatDate(dateObj, tz, 'yyyy'));
      var thYear = (y < 2400) ? y + 543 : y; // แปลง ค.ศ. -> พ.ศ.
      datePart = Utilities.formatDate(dateObj, tz, 'dd/MM/') + thYear;
    } else {
      datePart = String(d); // ถ้าแปลงไม่ได้จริงๆ ให้โชว์ค่าเดิม
    }
  }

  // --- 2. จัดการเวลา (Time) ---
  var timePart = '';
  if (tStr) {
    var tObj = null;
    
    // กรณี A: รับมาเป็น Date Object ตรงๆ
    if (tStr instanceof Date) {
      tObj = tStr;
    } 
    // กรณี B: รับมาเป็น String (รวมถึง String ยาวๆ แบบ "Sat Dec 30 1899...")
    else {
      var s = String(tStr).trim();
      // ถ้าเป็น String ยาวๆ ของปี 1899 ให้ลองแปลงเป็น Date
      if (s.length > 10 && (s.includes('1899') || s.includes('GMT') || s.includes('Sat'))) {
         var tryDate = new Date(s);
         if (!isNaN(tryDate.getTime())) tObj = tryDate;
      }
    }

    if (tObj) {
      // ถ้าได้เป็น Date Object แล้ว ให้จัดรูปแบบเป็น HH:mm
      timePart = Utilities.formatDate(tObj, tz, 'HH:mm');
    } else {
      // กรณี C: เป็น String เวลาปกติ เช่น "09:00" หรือ "9:30"
      var m = String(tStr).match(/(\d{1,2}):(\d{2})/);
      if (m) {
        timePart = m[1].padStart(2, '0') + ':' + m[2];
      } else {
        timePart = String(tStr); // ยอมแพ้ คืนค่าเดิม
      }
    }

    // เติม "น." ถ้ามีเวลาและยังไม่มีหน่วย
    if (timePart && timePart !== '-' && !timePart.includes('น.')) {
      timePart += ' น.';
    }
  }
  
  return (datePart + ' ' + timePart).trim();
}

// 1. ดึงประวัติประกันภัย
function apiGetInsuranceHistory() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSURANCE);
    if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };
    
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
    const data = vals.map(r => ({
      plate: String(r[1]),
      company: String(r[2]),
      endDate: _fmtThaiDateBE(r[5] ? new Date(r[5]) : null), // ใช้ Helper แปลง
      status: calculateStatus_(r[5])
    })).reverse();
    
    return { ok: true, data: data };
  } catch (e) { return { ok: false, error: e.message }; }
}

// 2. ดึงประวัติซ่อมบำรุง
function apiGetMaintenanceHistory() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAINTENANCE);
    if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };
    
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    const data = vals.map(r => ({
      date: _fmtThaiDateBE(r[0] ? new Date(r[0]) : null), // ใช้ Helper แปลง
      plate: String(r[1]),
      topic: String(r[2]),
      cost: Number(r[3] || 0).toLocaleString()
    })).reverse();
    
    return { ok: true, data: data };
  } catch (e) { return { ok: false, error: e.message }; }
}


// Helper เล็กๆ สำหรับคำนวณสถานะประกัน
function calculateStatus_(dateObj) {
  if (!dateObj) return 'ไม่ระบุ';
  const now = new Date();
  const due = new Date(dateObj);
  const diff = (due - now) / (1000 * 60 * 60 * 24);
  
  if (diff < 0) return 'expired'; // หมดอายุ
  if (diff < 30) return 'warning'; // ใกล้หมด (<30 วัน)
  return 'active'; // ปกติ
}

function getRealTimeAvailableCount(arg1, arg2, arg3) {
  try {
    var dateISO, startTime, endTime;

    // 🛠️ แก้ไขการรับค่า: รองรับทั้งแบบ Object (จาก Client) และ Arguments (จาก Server Test)
    if (typeof arg1 === 'object' && arg1 !== null) {
       dateISO = arg1.dateISO;
       startTime = arg1.startTime;
       endTime = arg1.endTime;
    } else {
       dateISO = arg1;
       startTime = arg2;
       endTime = arg3;
    }

    // 1. ตรวจสอบข้อมูลนำเข้า
    if (!dateISO || !startTime || !endTime) {
      return { ok: false, error: 'ข้อมูลวันเวลาไม่ครบถ้วน (Server Received: ' + JSON.stringify(arg1) + ')' };
    }

    // แปลงเวลาให้เป็น Date Object
    const reqStart = parseDateTime_(dateISO, startTime);
    const reqEnd = parseDateTime_(dateISO, endTime);
    
    if (!reqStart || !reqEnd) {
       return { ok: false, error: 'รูปแบบเวลาไม่ถูกต้อง (Time Parse Error)' };
    }
    
    if (reqEnd <= reqStart) {
      return { ok: false, error: 'เวลาสิ้นสุดต้องหลังเวลาเริ่มต้น' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2. ดึงข้อมูลรถทั้งหมดและเช็คสถานะซ่อมบำรุง
    const vehiclesRes = getAllVehiclePlatesFromSettings();
    const allVehicles = (vehiclesRes.ok && vehiclesRes.all) ? vehiclesRes.all : [];
    
    const vStatusKv = readSettingKV_('VehicleAvailability');
    const vStatusMap = parseBoolMap_(vStatusKv.val);

    let availableList = allVehicles.filter(v => {
      const isActive = vStatusMap.hasOwnProperty(v.plate) ? vStatusMap[v.plate] : true;
      return isActive;
    }).map(v => v.plate);

    // 3. ตัดรถที่ติดจอง "อนุมัติ"
    const sh = ss.getSheetByName('Data');
    if (sh) {
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = headerIndex_(headers);
      const lastRow = sh.getLastRow();
      
      if (lastRow > 1) {
        // อ่านข้อมูลเฉพาะคอลัมน์ที่จำเป็นเพื่อความเร็ว (Load ทั้งหมดอาจช้าถ้าข้อมูลเยอะ)
        const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const status = getStatusKeySafe_(row[idx.status]);
          
          if (status !== 'approved') continue;

          const plateStr = String(row[idx.vehicle] || '').trim();
          if (!plateStr) continue;

          const rDateISO = parseDateToISO_(row[idx.startDate]);
          const rStart = parseDateTime_(rDateISO, parseTimeSafe_(row[idx.startTime]));
          const rEnd = parseDateTime_(parseDateToISO_(row[idx.endDate]) || rDateISO, parseTimeSafe_(row[idx.endTime]));

          if (rStart && rEnd) {
             if (reqStart < rEnd && reqEnd > rStart) {
                const busyPlates = plateStr.split(',').map(p => p.trim());
                availableList = availableList.filter(p => !busyPlates.includes(p));
             }
          }
        }
      }
    }

    const count = availableList.length;
    const maxAllowed = Math.min(5, count);

    return { 
      ok: true, 
      count: count, 
      maxAllowed: maxAllowed,
      debug: availableList 
    };

  } catch (e) {
    Logger.log('getRealTimeAvailableCount Error: ' + e.toString());
    return { ok: false, error: e.message };
  }
}

function apiGetBookingsByPhone(phone) {
  try {
    const searchPhone = String(phone || '').replace(/\D/g, ''); // เก็บเฉพาะตัวเลข
    if (searchPhone.length < 9) return { ok: false, error: 'เบอร์โทรศัพท์สั้นเกินไปค่ะ' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Data'); // ใช้ชื่อชีตตาม Config
    if (!sh) throw new Error("ไม่พบชีต Data");

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };

    // ดึงข้อมูลทั้งหมดมาครั้งเดียว (เพื่อความเร็ว)
    // คาดการณ์ Column Index จาก Header (ใช้ Helper เดิมที่มี)
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const idx = headerIndex_(headers); // ฟังก์ชันเดิมในระบบ

    // ตรวจสอบ Index ที่จำเป็น
    if (idx.phone === undefined || idx.bookingId === undefined) {
      throw new Error("ไม่พบคอลัมน์ Phone หรือ Booking ID");
    }

    const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    const found = [];

    // วนลูปจากล่างขึ้นบน (ล่าสุดก่อน)
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const rowPhone = String(row[idx.phone] || '').replace(/\D/g, '');
      const status = getStatusKeySafe_(row[idx.status]);

      // เงื่อนไข: เบอร์ตรงกัน และ สถานะต้องไม่ใช่ Cancelled หรือ Rejected ไปแล้ว
      if (rowPhone === searchPhone && status !== 'cancelled' && status !== 'rejected') {
        const dateRaw = row[idx.startDate];
        const dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, 'Asia/Bangkok', 'dd/MM/yyyy') : String(dateRaw);
        
        found.push({
          bookingId: String(row[idx.bookingId]),
          summary: `${dateStr} : ${String(row[idx.destination] || '-')}`
        });
      }

      if (found.length >= 5) break; // เอาแค่ 5 รายการล่าสุด
    }

    return { ok: true, data: found };

  } catch (e) {
    Logger.log("apiGetBookingsByPhone Error: " + e.message);
    return { ok: false, error: e.message };
  }
}

/* [ANCHOR: Public Vehicle List for Dropdowns] */
function getVehicleList() {
  try {
    // 1. ดึงข้อมูลรถทั้งหมด (ใช้ Helper ที่พี่มีอยู่แล้ว)
    // ถ้าไม่มีฟังก์ชันนี้ ให้ลองเช็คว่าใน Code.gs มี getAllVehiclePlatesFromSettings ไหม
    var vehiclesRes = getAllVehiclePlatesFromSettings(); 
    var allVehicles = vehiclesRes.ok ? vehiclesRes.all : [];

    // 2. อ่านสถานะ Active/Inactive จาก Setting
    var vStatusKv = readSettingKV_('VehicleAvailability');
    var vStatusMap = parseBoolMap_(vStatusKv.val);

    // 3. กรองเอาเฉพาะรถที่ Active (สถานะเป็น true)
    var activeVehicles = allVehicles.filter(function(v) {
       // ถ้าไม่มีค่าใน Setting ให้ถือว่า Active (true) โดย Default
       return vStatusMap.hasOwnProperty(v.plate) ? vStatusMap[v.plate] : true;
    });

    // 4. ส่งกลับเฉพาะ "เลขทะเบียน" เป็น Array
    return activeVehicles.map(function(v) { return v.plate; });

  } catch (e) {
    Logger.log("getVehicleList Error: " + e.toString());
    // Fallback: กรณี Error ให้ส่งค่าว่าง หรือค่าทดสอบไปก่อน
    return ["99-9999 Test", "ฮค-1234"]; 
  }
}


function apiCheckAndPatchFileColumn() {
  const log = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Data'); // ชื่อชีตต้องตรงตาม Config
    if (!sh) throw new Error("ไม่พบชีต Data");

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, msg: "ไม่มีข้อมูลให้ตรวจสอบ" };

    // ค้นหา Index ของคอลัมน์ File จาก Header
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    let fileColIdx = -1;
    
    // ตรวจหาคอลัมน์ที่มีคำว่า File (Case-insensitive)
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().trim() === 'file') {
        fileColIdx = i + 1; // 1-based index
        break;
      }
    }

    if (fileColIdx === -1) {
        // Fallback: ใช้ค่าคงที่ถ้าหาไม่เจอ (จาก index 19 หรือ Col S)
        fileColIdx = 19; 
        log.push("⚠️ ไม่พบ Header 'File' โดยตรง ใช้ Default Col index: " + fileColIdx);
    } else {
        log.push("✅ พบ Header 'File' ที่คอลัมน์: " + fileColIdx);
    }

    // อ่านข้อมูลเฉพาะคอลัมน์ File
    const range = sh.getRange(2, fileColIdx, lastRow - 1, 1);
    const values = range.getValues();
    let fixedCount = 0;

    const newValues = values.map((r, i) => {
      let val = String(r[0]).trim();
      // ถ้าว่าง หรือ เป็น null/undefined ให้แก้เป็น "-"
      if (!val || val === '') {
        fixedCount++;
        return ['-'];
      }
      return [val];
    });

    // บันทึกกลับถ้ามีการแก้ไข
    if (fixedCount > 0) {
      range.setValues(newValues);
      log.push(`🛠️ ซ่อมแซมแถวที่ว่างจำนวน ${fixedCount} รายการ เรียบร้อยแล้ว`);
    } else {
      log.push("✅ ข้อมูลคอลัมน์ File สมบูรณ์ดีอยู่แล้ว");
    }

    return { ok: true, logs: log };

  } catch (e) {
    return { ok: false, error: e.message, logs: log };
  }
}

function apiUpdateBookingStatus(payload) {
  try {
    payload = payload || {};
    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); } catch (e0) {}
    }

    // Validate payload (ถ้ามี helper อยู่แล้วให้ใช้)
    if (typeof _assertUpdateStatusPayload_ === 'function') {
      _assertUpdateStatusPayload_(payload);
    }

    var bookingId = String(payload.bookingId || payload.id || '').trim();
    var newStatus = String(payload.status || '').toLowerCase().trim();
    var reasonText = (payload.reason == null) ? '' : String(payload.reason);
    var testMode = payload.testMode === true;
    var noTelegram = payload.noTelegram === true;

    if (!bookingId) return { ok: false, error: 'กรุณาระบุ Booking ID' };
    if (!newStatus) return { ok: false, error: 'กรุณาระบุสถานะ' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Data');
    if (!sheet) return { ok: false, error: 'ไม่พบชีต "Data" ในระบบค่ะ' };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok: false, error: 'ไม่พบข้อมูลในชีต Data' };

    // ---- headers (แถว 1) ----
    var headersRow = sheet.getRange(1, 1, 1, lastCol).getValues();
    var headers = (headersRow && headersRow[0]) ? headersRow[0].map(function (h) { return String(h || '').trim(); }) : [];

    // ---- idx mapping ----
    var idx = (typeof headerIndex_ === 'function') ? headerIndex_(headers) : null;

    function colIndexByName(name) {
      var i = headers.indexOf(name);
      return (i >= 0) ? i : undefined;
    }

    function normalizeIdx(v) {
      // กัน -1 หรือค่าประหลาด ให้เป็น undefined
      if (v === -1) return undefined;
      if (v === null || v === '' || v === false) return undefined;
      return v;
    }

    if (!idx || typeof idx !== 'object') idx = {};

    // เติม index ที่สำคัญ (ถ้า headerIndex_ ไม่ได้ให้มา)
    if (idx.bookingId === undefined) idx.bookingId = colIndexByName('Booking ID');
    if (idx.status === undefined) idx.status = colIndexByName('สถานะ');
    if (idx.reason === undefined) idx.reason = colIndexByName('Reason');
    if (idx.cancelReason === undefined) idx.cancelReason = colIndexByName('CancelReason');
    if (idx.vehicleCount === undefined) idx.vehicleCount = colIndexByName('จำนวนรถที่ต้องการ');
    if (idx.vehicle === undefined) idx.vehicle = colIndexByName('เลขทะเบียนรถ');
    if (idx.driver === undefined) idx.driver = colIndexByName('พนักงานขับรถ');

    // normalize ทั้งหมด (กัน -1 จาก headerIndex_)
    idx.bookingId = normalizeIdx(idx.bookingId);
    idx.status = normalizeIdx(idx.status);
    idx.reason = normalizeIdx(idx.reason);
    idx.cancelReason = normalizeIdx(idx.cancelReason);
    idx.vehicleCount = normalizeIdx(idx.vehicleCount);
    idx.vehicle = normalizeIdx(idx.vehicle);
    idx.driver = normalizeIdx(idx.driver);

    if (idx.bookingId === undefined) return { ok: false, error: 'ไม่พบคอลัมน์ "Booking ID"' };
    if (idx.status === undefined) return { ok: false, error: 'ไม่พบคอลัมน์ "สถานะ"' };

    // 1) Find row (อ่านทีเดียวทั้ง data)
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var foundOffset = -1; // offset ใน data (0-based)
    for (var i = 0; i < data.length; i++) {
      var idCell = data[i][idx.bookingId];
      if (String(idCell == null ? '' : idCell).trim() === bookingId) {
        foundOffset = i;
        break;
      }
    }
    if (foundOffset < 0) return { ok: false, error: 'ไม่พบ Booking ID: ' + bookingId };

    var rowIndex = foundOffset + 2; // ชีตจริง (1-based) โดยเริ่มข้อมูลที่แถว 2

    // 2) Prepare vehicles/drivers strings
    var vehiclesStr = '';
    var driversStr = '';
    if (payload.vehicles != null) {
      vehiclesStr = Array.isArray(payload.vehicles) ? payload.vehicles.join(', ') : String(payload.vehicles || '');
    }
    if (payload.drivers != null) {
      driversStr = Array.isArray(payload.drivers) ? payload.drivers.join(', ') : String(payload.drivers || '');
    }

    // helper: build row object from row values
    function buildRowObj(rowVals) {
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        var key = headers[c] || ('COL_' + (c + 1));
        obj[key] = rowVals[c];
      }
      // key มาตรฐานสำหรับ notify
      obj.bookingId = (idx.bookingId !== undefined) ? rowVals[idx.bookingId] : (obj['Booking ID'] || '');
      obj.status = (idx.status !== undefined) ? rowVals[idx.status] : (obj['สถานะ'] || '');
      return obj;
    }

    // Snapshot เดิม
    var oldRowVals = sheet.getRange(rowIndex, 1, 1, lastCol).getValues();
    oldRowVals = (oldRowVals && oldRowVals[0]) ? oldRowVals[0] : [];
    var rowObj = buildRowObj(oldRowVals);

    // 🍓 [BERRY FIX] ตรวจสอบคิวรถและคนขับชนกัน (Conflict Check) ก่อนบันทึกลงชีต
    if (newStatus === 'approved') {
      if (idx.startDate === undefined) idx.startDate = colIndexByName('วันเริ่มต้น');
      if (idx.startTime === undefined) idx.startTime = colIndexByName('เวลาเริ่มต้น');
      if (idx.endDate === undefined) idx.endDate = colIndexByName('วันสิ้นสุด');
      if (idx.endTime === undefined) idx.endTime = colIndexByName('เวลาสิ้นสุด');

      var sd = rowObj['วันเริ่มต้น'] || '';
      var st = rowObj['เวลาเริ่มต้น'] || '';
      var ed = rowObj['วันสิ้นสุด'] || '';
      var et = rowObj['เวลาสิ้นสุด'] || '';

      var reqVehicles = payload.vehicles ? (Array.isArray(payload.vehicles) ? payload.vehicles : String(payload.vehicles).split(',').map(function(s){return s.trim();})) : [];
      var reqDrivers = payload.drivers ? (Array.isArray(payload.drivers) ? payload.drivers : String(payload.drivers).split(',').map(function(s){return s.trim();})) : [];

      // เรียกใช้ฟังก์ชันตรวจสอบการชนกัน
      if (typeof checkResourcesConflict_ === 'function') {
        var conflictRes = checkResourcesConflict_(sheet, idx, sd, st, ed, et, bookingId, reqVehicles, reqDrivers);
        if (conflictRes && conflictRes.hasConflict) {
          // หากพบการชนกัน จะส่ง Error กลับไปยังหน้าบ้านทันทีและไม่บันทึก
          return { ok: false, error: conflictRes.message }; 
        }
      }
    }

    // 3) Update fields
    sheet.getRange(rowIndex, idx.status + 1).setValue(newStatus);

    // ✅ แยก Reason / CancelReason ให้ตรงคอลัมน์
    if (newStatus === 'cancelled') {
      if (idx.cancelReason !== undefined && reasonText) {
        sheet.getRange(rowIndex, idx.cancelReason + 1).setValue(reasonText);
      }
    } else {
      if (idx.reason !== undefined && reasonText) {
        sheet.getRange(rowIndex, idx.reason + 1).setValue(reasonText);
      }
    }

    // vehicleCount (ถ้าส่งมา)
    var actualCount = null;
    if (payload.vehicleCount != null && payload.vehicleCount !== '') {
      var vc = parseInt(payload.vehicleCount, 10);
      if (isFinite(vc) && vc > 0) {
        if (idx.vehicleCount !== undefined) {
          sheet.getRange(rowIndex, idx.vehicleCount + 1).setValue(vc);
        }
        actualCount = vc;
      }
    }

    // Approved -> write vehicles/drivers
    if (newStatus === 'approved') {
      if (idx.vehicle !== undefined && vehiclesStr) sheet.getRange(rowIndex, idx.vehicle + 1).setValue(vehiclesStr);
      if (idx.driver !== undefined && driversStr) sheet.getRange(rowIndex, idx.driver + 1).setValue(driversStr);
    }

    // ให้แน่ใจว่าค่าถูกเขียนแล้วก่อนอ่านกลับ
    SpreadsheetApp.flush();

    // 4) Refresh rowObj for notify + return
    var freshRowVals = sheet.getRange(rowIndex, 1, 1, lastCol).getValues();
    freshRowVals = (freshRowVals && freshRowVals[0]) ? freshRowVals[0] : [];
    rowObj = buildRowObj(freshRowVals);

    // หากไม่ได้ส่ง vehicleCount มา ลองอ่านจากชีต
    if (actualCount == null) {
      if (idx.vehicleCount !== undefined) {
        var vccell = freshRowVals[idx.vehicleCount];
        var vci = parseInt(vccell, 10);
        actualCount = (isFinite(vci) ? vci : (vccell == null ? '' : vccell));
      } else {
        actualCount = '';
      }
    }

    // 5) Notify (เคารพ noTelegram + testMode)
    var notifyRes = null;
    if (!noTelegram && typeof sendTelegramNotify === 'function') {
      notifyRes = sendTelegramNotify(rowObj, testMode === true);
    }

    return {
      ok: true,
      id: bookingId,
      status: newStatus,
      actualCount: actualCount,
      testMode: testMode,
      noTelegram: noTelegram,
      telegram: notifyRes
    };

  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}



function _assertUpdateStatusPayload_(payload) {
  payload = payload || {};
  if (typeof payload === 'string') {
    try { payload = JSON.parse(payload); } catch (e) {}
  }

  function err_(msg) { throw new Error('INVALID_PAYLOAD: ' + msg); }

  // bookingId
  var bookingId = String(payload.bookingId || payload.id || '').trim();
  if (!bookingId) err_('missing bookingId');

  // status
  var statusKey = String(payload.status || '').toLowerCase().trim();
  if (!statusKey) err_('missing status');

  var ALLOWED = { pending: 1, approved: 1, rejected: 1, cancelled: 1, canceled: 1 };
  if (!ALLOWED[statusKey]) err_('invalid status => ' + statusKey);

  // dryRun/testMode type guard
  function coerceBool_(v) {
    if (v === true || v === false) return v;
    if (v === 'true') return true;
    if (v === 'false') return false;
    if (v == null || v === '') return false;
    err_('invalid boolean => ' + v);
  }
  if ('dryRun' in payload) payload.dryRun = coerceBool_(payload.dryRun);
  if ('testMode' in payload) payload.testMode = coerceBool_(payload.testMode);

  // vehicles/drivers shape (optional)
  function isCsvOrArr_(v) {
    if (v == null || v === '') return true;
    if (Array.isArray(v)) return true;
    if (typeof v === 'string') return true;
    return false;
  }
  if (!isCsvOrArr_(payload.vehicles)) err_('invalid vehicles type');
  if (!isCsvOrArr_(payload.drivers)) err_('invalid drivers type');

  return true;
}




function coerceTimeHHmm_(v, tz) {
  tz = tz || Session.getScriptTimeZone() || 'Asia/Bangkok';

  if (v == null || v === '') return '';

  // Case 1: Date object
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, 'HH:mm');
  }

  // Case 2: number serial (Sheets time fraction)
  if (typeof v === 'number' && isFinite(v)) {
    var totalMinutes = Math.round(v * 24 * 60);
    var hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, '0');
    var mi = String(totalMinutes % 60).padStart(2, '0');
    return hh + ':' + mi;
  }

  // Case 3: string
  var s = String(v).trim();
  if (!s) return '';

  // Allow "9:00" -> "09:00"
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];

  // If it's something else (like "00:00:00") try trim
  m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})/);
  if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];

  return s; // fallback
}

function assertSheetTimeSanity_(sh, hm, bookingId, ok, ng) {
  bookingId = String(bookingId || '').trim();
  if (!bookingId) {
    ng('Sheet time sanity', 'Missing bookingId');
    return;
  }

  // ✅ local helper: coerce time to HH:mm
  function coerceTimeHHmm_(val, tz) {
    try {
      if (val == null || val === '') return '';

      // Date object
      if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
        return Utilities.formatDate(val, tz, 'HH:mm');
      }

      // number serial (Sheets time fraction)
      if (typeof val === 'number' && isFinite(val)) {
        var totalMinutes = Math.round(val * 24 * 60);
        var hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, '0');
        var mi = String(totalMinutes % 60).padStart(2, '0');
        return hh + ':' + mi;
      }

      // string
      var s = String(val).trim();
      if (!s) return '';
      s = s.replace(/\s+/g, ' ').replace(/น\.$/, '').trim(); // remove thai suffix if any
      var m = s.match(/^(\d{1,2}):(\d{2})/);
      if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];

      // string with seconds
      m = s.match(/^(\d{1,2}):(\d{2}):\d{2}/);
      if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];

      return s;
    } catch (_) {
      return '';
    }
  }

 
  // ✅ find row index (duplicate tiny finder to avoid dependency)
  function findRowByBookingIdLocal_(sh2, hm2, bookingId2) {
    bookingId2 = String(bookingId2 || '').trim();
    if (!bookingId2) return -1;

    var lastRow = sh2.getLastRow();
    if (lastRow < 2) return -1;

    var bookingCol = hm2.col('Booking ID');
    var data = sh2.getRange(2, 1, lastRow - 1, hm2.lastCol).getValues();

    for (var i = 0; i < data.length; i++) {
      var v = String(data[i][bookingCol - 1] || '').trim();
      if (v === bookingId2) return i + 2;
    }
    return -1;
  }

  var rowIndex = findRowByBookingIdLocal_(sh, hm, bookingId);
  if (rowIndex <= 0) {
    ng('Sheet time sanity', 'Booking ID not found: ' + bookingId);
    return;
  }

  function pickTimeCol_(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var h = candidates[i];
      try {
        var colNo = hm.col(h);
        if (colNo) return h;
      } catch (_) {}
    }
    return '';
  }

  var goHeader = pickTimeCol_(['เวลาไป', 'เวลาเริ่มต้น', 'Start Time', 'startTime']);
  var backHeader = pickTimeCol_(['เวลากลับ', 'เวลาสิ้นสุด', 'End Time', 'endTime']);

  if (!goHeader) {
    ng('Sheet เวลาไป header exists', 'Missing header: เวลาไป/เวลาเริ่มต้น/startTime');
    return;
  }
  if (!backHeader) {
    ng('Sheet เวลากลับ header exists', 'Missing header: เวลากลับ/เวลาสิ้นสุด/endTime');
    return;
  }

  var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';

  var goRaw = sh.getRange(rowIndex, hm.col(goHeader)).getValue();
  var backRaw = sh.getRange(rowIndex, hm.col(backHeader)).getValue();

  var go = coerceTimeHHmm_(goRaw, tz);
  var back = coerceTimeHHmm_(backRaw, tz);

  function isInvalid_(t) {
    if (!t) return true;
    if (t === '-' || t === '00:00') return true;
    if (t === '00:00 น.' || t === '00:00:00') return true;
    return false;
  }

  if (!isInvalid_(go)) ok('Sheet เวลาไป valid', go);
  else ng('Sheet เวลาไป valid', 'invalid=' + (go || '(empty)') + ' | raw=' + String(goRaw));

  if (!isInvalid_(back)) ok('Sheet เวลากลับ valid', back);
  else ng('Sheet เวลากลับ valid', 'invalid=' + (back || '(empty)') + ' | raw=' + String(backRaw));
}

function sendTelegramNotify(payload, testMode) {
  function toStr(v) {
    return String(v == null ? "" : v).trim();
  }

  function normalizeStatus(raw) {
    var s = toStr(raw).toLowerCase();

    if (!s) return "pending";
    if (s.indexOf("กรณีพิเศษ") > -1 || s.indexOf("พิเศษ") > -1 || s === "driver_special_approved") return "driver_special_approved";
    if (s.indexOf("ยกเลิก") > -1 || s.indexOf("cancel") > -1 || s === "cancelled") return "cancelled";
    if (s.indexOf("ไม่") > -1 || s.indexOf("reject") > -1 || s === "rejected") return "rejected";
    if (s.indexOf("อนุมัติ") > -1 || s.indexOf("approve") > -1 || s === "approved") return "approved";

    // รองรับข้อมูลเก่าในชีต/ระบบเดิม แต่ไม่ใช้เป็น flow หลักแล้ว
    if (s.indexOf("รับงาน") > -1 || s === "driver_claimed") return "pending";

    return "pending";
  }

  var msg = "";
  var statusKey = "pending";
  var bookingKey = "SYS";

  if (typeof payload === "object" && payload !== null) {
    var rawStatus = payload.status || payload["สถานะ"] || "pending";
    statusKey = normalizeStatus(rawStatus);

    bookingKey = toStr(payload.bookingId || payload.id || payload["Booking ID"] || "SYS");

    var reason = "";
    if (statusKey === "cancelled") {
      reason = toStr(payload.cancelReason || payload["CancelReason"] || payload.reason || payload["Reason"] || "");
    } else {
      reason = toStr(payload.reason || payload["Reason"] || payload.cancelReason || payload["CancelReason"] || "");
    }

    msg = buildBookingStatusMessage(payload, statusKey, reason);
  } else {
    msg = toStr(payload || "");
  }

  msg = toStr(msg)
    .replace(/\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/gi, "")
    .replace(/D\s*น\./gi, "")
    .replace(/[ \t]{2,}/g, " ")
    .trim();

  if (testMode === true) {
    return { ok: true, log: msg };
  }

  var config = (typeof getTelegramConfig === "function") ? getTelegramConfig() : { token: "", chatId: "" };
  var token = config.token;
  var chatId = config.chatId;

  if (testMode === "send_test") {
    var map = (typeof getSettingMap_ === "function") ? getSettingMap_() : {};
    token = map["Telegram Bot Test Token ID"] || token;
    chatId = map["Telegram Test Chat ID"] || chatId;
  }

  if (!token || !chatId) {
    var err = "❌ Telegram Config Missing (Token or ChatID)";
    if (typeof appendTgLog_ === "function") appendTgLog_("ERR_CONFIG", null, err);
    return { ok: false, error: err };
  }

  try {
    var url = "https://api.telegram.org/bot" + token + "/sendMessage";
    var options = {
      method: "post",
      payload: {
        chat_id: chatId,
        text: msg,
        parse_mode: "HTML",
        disable_web_page_preview: true
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var resText = response.getContentText();
    var resJson;

    try {
      resJson = JSON.parse(resText);
    } catch (parseErr) {
      resJson = { ok: false, error: "Telegram response is not valid JSON", raw: resText };
    }

    if (typeof appendTgLog_ === "function") {
      appendTgLog_("BID_" + bookingKey, response, msg);
    }

    return resJson;
  } catch (e) {
    console.error("❌ Telegram Send Error: " + e.message);
    if (typeof appendTgLog_ === "function") {
      appendTgLog_("ERR_SEND", null, e.message);
    }
    return { ok: false, error: e.message };
  }
}

// ===================== FEATURE 1: DRIVER CLAIM =====================
function claimBooking(payload) {
  return {
    ok: false,
    error: 'ปิดการใช้งานฟังก์ชันคนขับรับงานเองแล้ว กรุณาใช้ขั้นตอนอนุมัติจากผู้ดูแลระบบ'
  };
}

function specialApproveBooking(payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { ok: false, error: 'ระบบไม่ตอบสนอง', stage: 'lock' };
  }

  try {
    payload = payload || {};
    Logger.log('[specialApproveBooking] payload=' + JSON.stringify(payload));

    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Data');
    if (!sh) throw new Error('ไม่พบชีต Data');

    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) throw new Error('ไม่พบข้อมูลการจอง');

    var h = data[0].map(function(x) { return String(x || '').trim(); });
    var idx = {
      bid: h.indexOf('Booking ID'),
      st: h.indexOf('สถานะ'),
      v: h.indexOf('เลขทะเบียนรถ'),
      d: h.indexOf('พนักงานขับรถ')
    };

    if (idx.bid < 0 || idx.st < 0 || idx.v < 0 || idx.d < 0) {
      throw new Error('โครงสร้างชีต Data ไม่ครบคอลัมน์สำคัญ');
    }

    var bookingId = String(payload.bookingId || '').trim();
    var driverName = String(payload.driverName || '').trim();
    var plate = String(payload.plate || '').trim();

    if (!bookingId) throw new Error('ไม่พบ Booking ID');
    if (!driverName) throw new Error('ไม่พบชื่อพนักงานขับรถ');
    if (!plate) throw new Error('ไม่พบเลขทะเบียนรถ');

    var rowIndex = -1;
    var rowData = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idx.bid] || '').trim() === bookingId) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1 || !rowData) {
      throw new Error('ไม่พบข้อมูลการจอง');
    }

    var currentStatus = getStatusKeySafe_(rowData[idx.st]);
    if (currentStatus !== 'pending') {
      throw new Error('รายการนี้ไม่ได้อยู่ในสถานะรออนุมัติแล้ว');
    }

    sh.getRange(rowIndex, idx.st + 1).setValue('driver_special_approved');
    sh.getRange(rowIndex, idx.v + 1).setValue(plate);
    sh.getRange(rowIndex, idx.d + 1).setValue(driverName);
    SpreadsheetApp.flush();

    Logger.log('[specialApproveBooking] sheet updated bookingId=' + bookingId + ', plate=' + plate + ', driver=' + driverName);

    var freshRow = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    var notifyObj = {};
    h.forEach(function(key, colIndex) {
      notifyObj[key] = freshRow[colIndex];
    });

    notifyObj.status = 'driver_special_approved';
    notifyObj.bookingId = bookingId;
    notifyObj.driverName = driverName;
    notifyObj.plate = plate;

    var tg = sendTelegramNotify(notifyObj, false);
    Logger.log('[specialApproveBooking] telegram result bookingId=' + bookingId + ' => ' + JSON.stringify(tg));

    var tgOk = !!(tg && (tg.ok === true || tg.result));
    if (!tgOk) {
      return {
        ok: false,
        error: 'บันทึกข้อมูลสำเร็จ แต่ส่ง Telegram ไม่สำเร็จ',
        stage: 'telegram',
        bookingId: bookingId,
        status: 'driver_special_approved',
        plate: plate,
        driverName: driverName,
        sheetSaved: true,
        telegram: tg
      };
    }

    return {
      ok: true,
      bookingId: bookingId,
      status: 'driver_special_approved',
      plate: plate,
      driverName: driverName,
      sheetSaved: true,
      telegram: tg
    };
  } catch (e) {
    Logger.log('[specialApproveBooking][ERROR] ' + (e && e.stack ? e.stack : e));
    return {
      ok: false,
      error: e.message,
      stage: 'server'
    };
  } finally {
    lock.releaseLock();
  }
}
// ===================== FEATURE 2: AVAILABILITY ENGINE =====================
function _getAvailabilitySheet() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Availability');
  if(!sh) {
    sh = ss.insertSheet('Availability');
    sh.appendRow(['resourceType','resourceId','startDate','startTime','endDate','endTime','reason','createdBy','createdAt']);
  }
  return sh;
}

function createAvailabilityBlock(payload) {
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(5000)) return {ok:false, error:'System busy'};
  try {
    var sh = _getAvailabilitySheet();
    sh.appendRow([
      payload.resourceType, payload.resourceId,
      parseDateToISO_(payload.startDate), parseTimeSafe_(payload.startTime),
      parseDateToISO_(payload.endDate), parseTimeSafe_(payload.endTime),
      payload.reason, payload.createdBy, new Date()
    ]);
    return {ok:true};
  } catch(e) { return {ok:false, error:e.message}; }
  finally { lock.releaseLock(); }
}

function _checkAvailabilityOverlap(resType, resId, reqStart, reqEnd) {
  var sh = _getAvailabilitySheet();
  var data = sh.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] === resType && data[i][1] === resId) {
      var bStart = parseDateTime_(data[i][2], data[i][3]);
      var bEnd = parseDateTime_(data[i][4], data[i][5]);
      if(isOverlapping_(reqStart, reqEnd, bStart, bEnd)) return { conflict: true, reason: data[i][6] };
    }
  }
  return { conflict: false };
}

function checkDriverAvailability(driver, sd, st, ed, et) {
  var rs = parseDateTime_(sd, st), re = parseDateTime_(ed, et);
  return _checkAvailabilityOverlap('driver', driver, rs, re);
}

function checkVehicleAvailability(plate, sd, st, ed, et) {
  var rs = parseDateTime_(sd, st), re = parseDateTime_(ed, et);
  return _checkAvailabilityOverlap('vehicle', plate, rs, re);
}

const VB_RADAR_DRIVER_MASTER = [
  'ชัชวาลย์ วงศ์มั่น',
  'ประเสริฐ หน่อแก้ว',
  'ปรีชา ถวิลเวช',
  'ปริญญา ก้อนสัมฤทธิ์',
  'อภิรัฐวุฒิ คณารักษ์'
];

const VB_RADAR_VEHICLE_MASTER = [
  'ฮล-466',
  'ฮค-4964',
  '1นช-6112',
  'ฮร-4820',
  'ห-4845'
];

function normalizeRadarName_(v) {
  var s = String(v || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (s === 'ปรีชา ถวิล เวช') s = 'ปรีชา ถวิลเวช';
  return s;
}

function normalizeRadarPlate_(v) {
  return String(v || '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseDateTimeBkk_(dateRaw, timeRaw) {
  try {
    var dISO = parseDateToISO_(dateRaw);
    if (!dISO) return null;
    var tStr = parseTimeSafe_(timeRaw || '00:00');
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var dStr = dISO + ' ' + tStr + ':00';
    var d = Utilities.parseDate(dStr, tz, 'yyyy-MM-dd HH:mm:ss');
    if (!d || isNaN(d.getTime())) {
      // Fallback
       var p = dISO.split('-');
       var tp = tStr.split(':');
       d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]), parseInt(tp[0]), parseInt(tp[1]), 0, 0);
    }
    return d;
  } catch (e) {
    Logger.log('parseDateTimeBkk_ error: ' + e.message);
    return null;
  }
}

function getServerNowBangkok_() {
  var tz = Session.getScriptTimeZone() || (typeof TZ !== 'undefined' ? TZ : 'Asia/Bangkok');
  return Utilities.parseDate(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'), tz, 'yyyy-MM-dd HH:mm:ss');
}

function isNowBetween_(now, startAt, endAt) {
  if (!now || !startAt || !endAt) return false;
  return now.getTime() >= startAt.getTime() && now.getTime() <= endAt.getTime();
}

// ANCHOR: buildRadarContext_
function buildRadarContext_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = 'Asia/Bangkok';
  var now = new Date(); // เวลาปัจจุบันระดับ Server

  // Helper สำหรับแปลง Date/Time ให้ล็อกอยู่ใน Timezone BKK ป้องกันเวลาเพี้ยน
  function getRadarDateTime_(dISO, tStr, defaultTime) {
      if (!dISO) return null;
      var cleanISO = parseDateToISO_(dISO);
      if (!cleanISO) return null;
      var cleanTime = parseTimeSafe_(tStr) || defaultTime;
      var fullStr = cleanISO + ' ' + cleanTime + ':00';
      return Utilities.parseDate(fullStr, tz, 'yyyy-MM-dd HH:mm:ss');
  }

  var availBlocks =[];
  var shAvail = ss.getSheetByName('Availability');
  if (shAvail && shAvail.getLastRow() > 1) {
    var avVals = shAvail.getRange(2, 1, shAvail.getLastRow() - 1, shAvail.getLastColumn()).getValues();
    for (var i = 0; i < avVals.length; i++) {
      var row = avVals[i];
      var resourceType = String(row[0] || '').trim().toLowerCase();
      var resourceId = String(row[1] || '').trim();
      
      // อ้างอิง index ตามหัวคอลัมน์มาตรฐาน: 0=Type, 1=Id, 2=startD, 3=startT, 4=endD, 5=endT, 6=Reason
      var startAt = getRadarDateTime_(row[2], row[3], '00:00');
      var endAt = getRadarDateTime_(row[4] || row[2], row[5], '23:59');

      if (!resourceType || !resourceId || !startAt || !endAt) continue;

      availBlocks.push({
        resourceType: resourceType,
        resourceId: resourceType === 'driver' ? normalizeRadarName_(resourceId) : normalizeRadarPlate_(resourceId),
        startAt: startAt,
        endAt: endAt,
        reason: String(row[6] || '').trim()
      });
    }
  }

  var approvedBookings =[];
  var shData = ss.getSheetByName('Data');
  if (shData && shData.getLastRow() > 1) {
    var headers = shData.getRange(1, 1, 1, shData.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
    var idx = headerIndex_(headers);
    var dataVals = shData.getRange(2, 1, shData.getLastRow() - 1, shData.getLastColumn()).getValues();

    for (var r = 0; r < dataVals.length; r++) {
      var row2 = dataVals[r];
      var statusKey = getStatusKeySafe_(row2[idx.status]);
      if (statusKey !== 'approved' && statusKey !== 'driver_special_approved') continue;

      var startAt2 = getRadarDateTime_(row2[idx.startDate], row2[idx.startTime], '00:00');
      var endAt2 = getRadarDateTime_(row2[idx.endDate] || row2[idx.startDate], row2[idx.endTime], '23:59');
      
      if (!startAt2 || !endAt2) continue;

      approvedBookings.push({
        bookingId: String(row2[idx.bookingId] || '').trim(),
        status: statusKey,
        driver: normalizeRadarName_(row2[idx.driver]),
        vehicle: normalizeRadarPlate_(row2[idx.vehicle]),
        destination: String(row2[idx.destination] || '').trim(),
        workName: String(row2[idx.workName] || '').trim(),
        startAt: startAt2,
        endAt: endAt2
      });
    }
  }

  // 🍓 BERRY FIX: ตัดการส่งค่า Global Toggle ทิ้ง บังคับให้ใช้เวลาจาก Sheet เท่านั้น
  return {
    now: now,
    tz: tz,
    availBlocks: availBlocks,
    approvedBookings: approvedBookings
  };
}

// ANCHOR: calculateVehicleStatus
function calculateVehicleStatus(plate, ctx) {
  var normPlate = normalizeRadarPlate_(plate);
  var now = ctx.now;

  function isTargetActive(startAt, endAt) {
    if (!startAt || !endAt) return false;
    // กฎ: เวลาปัจจุบัน ต้องอยู่ระหว่าง วันเวลาเริ่มต้น และ สิ้นสุด เท่านั้น
    return now.getTime() >= startAt.getTime() && now.getTime() <= endAt.getTime();
  }

  // 1. ส่งซ่อม (ตรวจสอบจาก Availability Block ตามเวลาจริง)
  var activeRepair = (ctx.availBlocks ||[]).find(function(b) {
    return b.resourceType === 'vehicle' && normalizeRadarPlate_(b.resourceId) === normPlate && isTargetActive(b.startAt, b.endAt);
  });
  
  if (activeRepair) {
     return { status: 'repair', label: 'ส่งซ่อม', color: 'red', job: activeRepair.reason || 'ซ่อมบำรุง' };
  }

  // 2. ไม่ว่าง (งาน Booking ที่กำลังเกิดขึ้น ณ วินาทีนี้)
  var activeBooking = (ctx.approvedBookings ||[]).find(function(b) {
    if (!isTargetActive(b.startAt, b.endAt)) return false;
    var plates = String(b.vehicle || '').split(',').map(function(x) { return normalizeRadarPlate_(x); });
    return plates.indexOf(normPlate) > -1;
  });
  
  if (activeBooking) {
     return { status: 'busy', label: 'ไม่ว่าง', color: 'yellow', job: activeBooking.destination || activeBooking.workName };
  }

  // 3. พร้อม (ไม่ติดทั้งซ่อมและงาน)
  return { status: 'ready', label: 'พร้อม', color: 'green', job: 'พร้อมใช้งาน' };
}

// ANCHOR: calculateDriverStatus
function calculateDriverStatus(driverName, ctx) {
  var name = normalizeRadarName_(driverName);
  var now = ctx.now;

  function isTargetActive(startAt, endAt) {
    if (!startAt || !endAt) return false;
    return now.getTime() >= startAt.getTime() && now.getTime() <= endAt.getTime();
  }

  // 1. ลา (ตรวจสอบจาก Availability Block ตามเวลาจริง)
  var activeLeave = (ctx.availBlocks ||[]).find(function(b) {
    return b.resourceType === 'driver' && normalizeRadarName_(b.resourceId) === name && isTargetActive(b.startAt, b.endAt);
  });
  
  if (activeLeave) {
     return { status: 'leave', label: 'ลา', color: 'red', job: activeLeave.reason || 'ลางาน' };
  }

  // 2. ติดภารกิจ (งาน Booking ที่กำลังเกิดขึ้น ณ วินาทีนี้)
  var activeBooking = (ctx.approvedBookings ||[]).find(function(b) {
    if (!isTargetActive(b.startAt, b.endAt)) return false;
    var drivers = String(b.driver || '').split(',').map(function(x) { return normalizeRadarName_(x); });
    return drivers.indexOf(name) > -1;
  });
  
  if (activeBooking) {
     return { status: 'busy', label: 'ติดภารกิจ', color: 'yellow', job: activeBooking.destination || activeBooking.workName };
  }

  // 3. พร้อม 
  return { status: 'ready', label: 'พร้อม', color: 'green', job: 'พร้อมปฏิบัติงาน' };
}

// ANCHOR: ฟังก์ชัน buildRadarData (เรียกใช้สถานะปัจจุบัน)
function buildRadarData() {
  var ctx = buildRadarContext_();
  var yearBE = parseInt(Utilities.formatDate(ctx.now, ctx.tz, 'yyyy'), 10) + 543;

  var drivers = VB_RADAR_DRIVER_MASTER.map(function(name) {
    var st = calculateDriverStatus(name, ctx);
    return { name: name, active: true, status: st.status, label: st.label, color: st.color, job: st.job };
  });

  var vehicles = VB_RADAR_VEHICLE_MASTER.map(function(plate) {
    var st = calculateVehicleStatus(plate, ctx);
    return { plate: plate, active: true, status: st.status, label: st.label, color: st.color, job: st.job };
  });

  return {
    ok: true,
    serverNow: Utilities.formatDate(ctx.now, ctx.tz, 'yyyy-MM-dd HH:mm:ss'),
    serverDateThai: Utilities.formatDate(ctx.now, ctx.tz, 'dd/MM/') + yearBE,
    drivers: drivers, 
    vehicles: vehicles 
  };
}


function apiGetLiveStatus() {
  try {
    return buildRadarData();
  } catch (e) {
    Logger.log('apiGetLiveStatus Error: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: e.message };
  }
}

// 🍓 BERRY FIX: Diagnostic function for Radar Vehicle Maintenance Timezone check
function selfTestRadarVehicleTime() {
  Logger.log('🚀 === START: selfTestRadarVehicleTime ===');
  
  // สร้าง Mock context แทนการดึงข้อมูลจริงจากชีต
  var startAt = parseDateTimeBkk_('2026-03-20', '00:00');
  var endAt = parseDateTimeBkk_('2026-03-20', '23:59');
  
  var mockAvailBlocks = [{
    resourceType: 'vehicle',
    resourceId: 'ฮร-4820',
    startAt: startAt,
    endAt: endAt,
    reason: 'ส่งซ่อมบำรุงตามรอบเช็คระยะ',
    status: 'repair',
    type: 'repair',
    title: ''
  }];
  
  var mockBookings = [{
    vehicle: 'ฮร-4820',
    startAt: parseDateTimeBkk_('2026-03-14', '08:00'),
    endAt: parseDateTimeBkk_('2026-03-14', '12:00'),
    status: 'approved',
    workName: 'Booking ก่อนเวลาซ่อม'
  }];

  function runCase(stepName, currentServerTimeISO, expectedStatus) {
    var testNow = parseDateTimeBkk_(currentServerTimeISO.split(' ')[0], currentServerTimeISO.split(' ')[1] || '00:00');
    
    var ctx = {
      now: testNow,
      availBlocks: mockAvailBlocks,
      approvedBookings: mockBookings,
      vehicleStatusMap: {}
    };
    
    var res = calculateVehicleStatus('ฮร-4820', ctx);
    var pass = res.status === expectedStatus;
    
    Logger.log('[' + stepName + '] ' + (pass ? '✅ PASS' : '❌ FAIL'));
    Logger.log('  resourceId: ฮร-4820');
    Logger.log('  startDateTime: ' + startAt);
    Logger.log('  endDateTime: ' + endAt);
    Logger.log('  currentServerDateTime: ' + testNow);
    Logger.log('  finalStatus: ' + res.status + ' (Expected: ' + expectedStatus + ')');
    Logger.log('-----------------------------------');
  }

  runCase('STEP1', '2026-03-14 10:00', 'busy');   // ตรงกับ Booking
  runCase('STEP1.5', '2026-03-15 10:00', 'ready'); // ว่าง ไม่มี Booking และยังไม่ถึงซ่อม
  runCase('STEP2', '2026-03-20 10:00', 'repair');  // อยู่ในช่วงซ่อม
  runCase('STEP3', '2026-03-21 00:01', 'ready');   // เลยช่วงซ่อมแล้ว
  
  Logger.log('🏁 === END: selfTestRadarVehicleTime ===');
}

function appendTgLog_(key, response, text) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('_tg_log');
    if (!sh) {
      sh = ss.insertSheet('_tg_log');
      sh.appendRow(['key', 'http_code', 'text', 'ts']);
    }
    var code = response ? response.getResponseCode() : 'FAIL';
    sh.appendRow([
      'BID_' + (key || 'SYS'),
      code,
      text ? text.substring(0, 500) : '',
      new Date()
    ]);
  } catch (e) { Logger.log("Log Error: " + e.message); }
}

function runFullTelegramLogTest() {
  Logger.log("🚀 === เริ่มต้นการทดสอบระบบ V-Berry Diagnostics (Log Only) ===\n");

  var mockPayload = {
    "Booking ID": "TEST-1391",
    "ชื่อ-สกุล": "คุณปรีชา ทดสอบระบบ",
    "ตำแหน่ง": "อาจารย์",
    "เบอร์โทร": "0812345678",
    "ประเภทงาน": "ประชุม",
    "งาน/โครงการ": "วางแผนยุทธศาสตร์ AI 2026",
    "สถานที่": "มหาวิทยาลัยสวนดุสิต (กรุงเทพฯ)",
    "ประเภทรถ": "รถตู้",
    "จำนวนรถที่ต้องการ": "2",
    "จำนวนผู้ร่วมเดินทาง": "12",
    "วันเริ่มต้น": "2026-03-12",
    "เวลาเริ่มต้น": "08:30",
    "วันสิ้นสุด": "2026-03-12",
    "เวลาสิ้นสุด": "16:30",
    "เลขทะเบียนรถ": "",
    "พนักงานขับรถ": "",
    "Reason": "",
    "CancelReason": ""
  };

  function runIndividualCase(caseTitle, status, extraData) {
    Logger.log("💬 [" + caseTitle + "]");
    var payload = Object.assign({}, mockPayload, extraData || {});
    payload.status = status;
    var res = sendTelegramNotify(payload, true); // true = test mode (log only)
    if (res && res.ok) Logger.log("\n" + res.log + "\n");
    Logger.log("----------------------------------------");
  }

  // 1-6. ทดสอบแต่ละสถานะ
  runIndividualCase("1. จองใหม่ (Pending)", "pending");
  runIndividualCase("2. [SPECIAL TEST] อนุมัติกรณีพิเศษ", "driver_special_approved", { "เลขทะเบียนรถ": "นข-9999", "พนักงานขับรถ": "นายสมชาย" });
  runIndividualCase("3. อนุมัติปกติ", "approved", { "เลขทะเบียนรถ": "ฮค-1234", "พนักงานขับรถ": "พี่ยอด" });
  
  // 🍓 BERRY FIX: เพิ่มเคสทดสอบเปลี่ยนรถ/คนขับ (Re-Assign) ตรวจจับคำว่า 'อัปเดต'
  runIndividualCase("4. [RE-ASSIGN] เปลี่ยนรถ/คนขับ", "approved", { "เลขทะเบียนรถ": "กท-5555", "พนักงานขับรถ": "นายเอกชัย", "Reason": "อัปเดตการมอบหมายรถ/คนขับใหม่" });
  
  runIndividualCase("5. ไม่อนุมัติ", "rejected", { "Reason": "รถติดภารกิจ" });
  runIndividualCase("6. ยกเลิกการจอง", "cancelled", { "CancelReason": "ยกเลิกโครงการ" });

  // 7. ทดสอบรายงานประจำวัน (Daily Summary: มีงาน)
  Logger.log("📋 [7. รายงานสรุปประจำวัน (Daily Summary: มีงาน - 05:00 AM)]\n");

  var mockHeaders = ["Booking ID", "สถานะ", "ชื่อ-สกุล", "ประเภทงาน", "งาน/โครงการ", "สถานที่", "เลขทะเบียนรถ", "พนักงานขับรถ", "วันเริ่มต้น", "เวลาเริ่มต้น", "วันสิ้นสุด", "เวลาสิ้นสุด"];
  var mockDataRowsJobs = [
    mockHeaders,
    ["BK-001", "approved", "สมชาย จองจริง", "ประชุม", "งานแผน", "ศูนย์ฯ ลำปาง", "ฮค-4964", "ประเสริฐ", "2026-03-12", "09:00", "2026-03-12", "12:00"],
    ["BK-002", "pending", "สมหญิง พึ่งพา", "อบรม", "โครงการ A", "กทม.", "", "", "2026-03-12", "07:00", "2026-03-12", "17:00"],
    ["BK-003", "driver_special_approved", "ผอ.ศูนย์", "รับรอง", "ต้อนรับแขก", "สนามบิน", "นข-1111", "พี่ยอด", "2026-03-12", "14:00", "2026-03-12", "16:00"]
  ];

  var originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  
  SpreadsheetApp.getActiveSpreadsheet = function() {
    return { getSheetByName: function() { return { getDataRange: function() { return { getValues: function() { return mockDataRowsJobs; } }; } }; } };
  };
  var dailyMsgJobs = getIntegratedDailyReport(new Date(2026, 2, 12));
  Logger.log(dailyMsgJobs);
  Logger.log("\n----------------------------------------");

  // 8. ทดสอบรายงานประจำวัน (Daily Summary: ไม่มีงาน)
  Logger.log("📋 [8. รายงานสรุปประจำวัน (Daily Summary: ไม่มีงาน - 05:00 AM)]\n");
  var mockDataRowsNoJobs = [ mockHeaders ]; // มีแค่ Header ไร้ Data
  
  SpreadsheetApp.getActiveSpreadsheet = function() {
    return { getSheetByName: function() { return { getDataRange: function() { return { getValues: function() { return mockDataRowsNoJobs; } }; } }; } };
  };
  var dailyMsgNoJobs = getIntegratedDailyReport(new Date(2026, 2, 12));
  Logger.log(dailyMsgNoJobs);

  // Restore original
  SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
  Logger.log("\n🏁 === สิ้นสุดการทดสอบระบบ V-Berry Diagnostics ===");
}



