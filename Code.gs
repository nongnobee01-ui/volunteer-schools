/**
 * เครือข่ายเยาวชน ศอญ. / โรงเรียนจิตอาสา ทภ.1 — Apps Script backend (v2)
 * วางใน Google Sheets (Extensions > Apps Script) แทนโค้ดเดิมทั้งหมด
 *
 * GET  ?action=all   → ข้อมูลทั้ง 4 เครือข่าย + รายงานกิจกรรม (หน้าเว็บใช้ตัวนี้)
 * POST (JSON แบบ text/plain) → รับรายงานติดตามกิจกรรม
 *
 * แท็บในชีต (กด Run "setup" หนึ่งครั้ง ระบบสร้างแท็บที่ยังไม่มีพร้อมหัวคอลัมน์ให้)
 *   1.1 จอส904        → เครือข่าย จอส.904
 *   1.2 Sheet1        → เครือข่าย ฉก.ทม.รอ.904 และหน่วยในกองทัพ (บัญชีโรงเรียนเดิม)
 *   1.3 OVN-DAYTRIP   → เครือข่าย overnight trip / day trip
 *   1.4 ครูก           → เครือข่าย ครู ก (กลุ่ม: ครู / ทหารตำรวจ / ปราชญ์ชาวบ้าน)
 *   ติดตามกิจกรรม       → รายงานที่ส่งจากหน้าเว็บ (สร้างอัตโนมัติ)
 */

const CONFIG = {
  ACTIVITY_SHEET: 'ติดตามกิจกรรม',
  DRIVE_FOLDER_NAME: 'ภาพกิจกรรม โรงเรียนจิตอาสา ทภ.1',
  DRIVE_FOLDER_ID: '',
  MAX_IMAGES: 3,
  MAX_IMAGE_BYTES: 4 * 1024 * 1024,
  TZ: 'Asia/Bangkok',
};

// [หัวคอลัมน์ในชีต, ชื่อ field ที่ส่งให้หน้าเว็บ]  — จับคู่แบบตัดช่องว่าง ไม่สนตัวพิมพ์
const NETWORKS = [
  { key: 'jos', tab: 'จอส904', fields: [
    ['ชื่อ-สกุล', 'name'], ['สถานะ/ตำแหน่ง', 'role'], ['รุ่น', 'batch'], ['หน่วย/สังกัด', 'unit'],
    ['อำเภอ', 'amp'], ['จังหวัด', 'prov'], ['พิกัด', 'll'], ['หมายเหตุ', 'note'] ] },
  { key: 'unit', tab: 'Sheet1', fields: [
    ['ประเภทเป้าหมาย', 'type'], ['โรงเรียน', 'name'], ['ที่ตั้ง googlemap', 'll'], ['อำเภอ', 'amp'],
    ['จังหวัด', 'prov'], ['ผู้ประสานโรงเรียน(ครู)', 'teacher'], ['หมายเลขติดต่อโรงเรียน', 'schoolTel'],
    ['หน่วยประสานงาน', 'unit'], ['POC หน่วย', 'poc'], ['หมายเลขติดต่อ POC', 'tel'] ] },
  { key: 'trip', tab: 'OVN-DAYTRIP', fields: [
    ['โรงเรียน', 'name'], ['ประเภท', 'type'], ['วันที่', 'date'], ['หน่วยดำเนินการ', 'unit'],
    ['จำนวนนักเรียน', 'students'], ['อำเภอ', 'amp'], ['จังหวัด', 'prov'], ['พิกัด', 'll'], ['รายละเอียด', 'detail'] ] },
  { key: 'kru', tab: 'ครูก', fields: [
    ['ชื่อ-สกุล', 'name'], ['กลุ่ม', 'group'], ['สังกัด/หน่วยงาน', 'org'], ['ความเชี่ยวชาญ', 'skill'],
    ['อำเภอ', 'amp'], ['จังหวัด', 'prov'], ['พิกัด', 'll'] ] },
];

const ACTIVITY_HEADERS = ['รหัส', 'เวลาส่ง', 'วันที่กิจกรรม', 'เครือข่าย', 'หน่วย/ชื่อ', 'กิจกรรม', 'ลิงก์โพสต์', 'ภาพ (file id)', 'โฟลเดอร์ภาพ'];
const NETWORK_NAMES = ['จอส.904', 'ฉก.ทม.รอ.904 และหน่วยในกองทัพ', 'Overnight trip / Day trip', 'ครู ก'];

/* ---------------- GET ---------------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const out = { ok: true, updated: new Date().toISOString(), networks: {} };
    NETWORKS.forEach(n => out.networks[n.key] = readNetwork_(n));
    out.activities = getActivities_(Number(p.limit) || 2000);
    out.schools = out.networks.unit; // ให้หน้าเดิม (weekly.html) ยังอ่านได้
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function norm_(s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); }

function readNetwork_(n) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(n.tab) || (n.key === 'unit' ? ss.getSheets()[0] : null);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getDisplayValues();
  const head = values[0].map(norm_);
  const idx = {};
  n.fields.forEach(([h, f]) => { const i = head.indexOf(norm_(h)); if (i > -1) idx[f] = i; });
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const o = {};
    Object.keys(idx).forEach(f => o[f] = String(values[r][idx[f]] || '').trim());
    if (!o.name) continue;
    if (!/^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/.test(o.ll || '')) o.ll = '';
    out.push(o);
  }
  return out;
}

function getActivities_(limit) {
  const sh = activitySheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit, last - 1);
  const rows = sh.getRange(last - n + 1, 1, n, ACTIVITY_HEADERS.length).getValues();
  return rows.reverse().map(r => {
    const date = r[2] instanceof Date ? Utilities.formatDate(r[2], CONFIG.TZ, 'yyyy-MM-dd') : String(r[2]);
    const images = String(r[7] || '').split(',').map(s => s.trim()).filter(Boolean);
    return {
      id: r[0], ts: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), date: date,
      network: r[3], who: r[4], activity: r[5], link: r[6], images: images,
      // ชื่อ field เดิม เผื่อหน้าอื่นที่ยังใช้อยู่
      unit: r[4], title: r[5], type: r[3], school: '', prov: '', detail: r[6], students: 0, staff: 0,
    };
  });
}

/* ---------------- POST ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true }); // honeypot กันบอท

    const need = { date: 'วันที่', network: 'ชื่อเครือข่าย', who: 'ชื่อหน่วย/ชื่อตัว', activity: 'ชื่อกิจกรรม', link: 'ลิงก์ที่โพสต์' };
    const miss = Object.keys(need).filter(k => !String(d[k] || '').trim()).map(k => need[k]);
    if (miss.length) return json_({ ok: false, error: 'กรอกไม่ครบ: ' + miss.join(', ') });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return json_({ ok: false, error: 'รูปแบบวันที่ไม่ถูกต้อง' });
    if (NETWORK_NAMES.indexOf(d.network) < 0) return json_({ ok: false, error: 'ไม่รู้จักชื่อเครือข่ายนี้' });
    const links = String(d.link).split(/\s+/).filter(Boolean);
    if (!links.every(u => /^https?:\/\/\S+\.\S+/.test(u))) return json_({ ok: false, error: 'ลิงก์ต้องขึ้นต้นด้วย http:// หรือ https://' });

    const id = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyMMddHHmmss') + '-' + Math.random().toString(36).slice(2, 6);
    const images = Array.isArray(d.images) ? d.images.slice(0, CONFIG.MAX_IMAGES) : [];
    let fileIds = [], folderUrl = '';
    if (images.length) {
      const sub = getOrCreate_(getOrCreate_(rootFolder_(), d.date.slice(0, 7)), clip_(`${d.date}_${d.who}`, 120));
      folderUrl = sub.getUrl();
      images.forEach((img, i) => {
        const bytes = Utilities.base64Decode(String(img.data || '').replace(/^data:[^,]+,/, ''));
        if (bytes.length > CONFIG.MAX_IMAGE_BYTES) throw new Error(`ภาพที่ ${i + 1} ใหญ่เกินไป`);
        const f = sub.createFile(Utilities.newBlob(bytes, 'image/jpeg', `${id}_${i + 1}.jpg`));
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileIds.push(f.getId());
      });
    }

    lock.waitLock(20000);
    activitySheet_().appendRow([
      id, new Date(), d.date, d.network, clip_(d.who, 150), clip_(d.activity, 300),
      clip_(links.join('\n'), 1000), fileIds.join(','), folderUrl,
    ]);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ---------------- helpers ---------------- */
function activitySheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.ACTIVITY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ACTIVITY_SHEET);
    styleHeader_(sh, ACTIVITY_HEADERS);
  }
  return sh;
}

function styleHeader_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFD21F');
  sh.setFrozenRows(1);
}

function rootFolder_() {
  if (CONFIG.DRIVE_FOLDER_ID) return DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('ROOT_FOLDER_ID');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (_) {} }
  const f = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', f.getId());
  return f;
}

function getOrCreate_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function clip_(s, n) { return String(s == null ? '' : s).trim().slice(0, n); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งเดียว: สร้างแท็บที่ยังไม่มีพร้อมหัวคอลัมน์ และแสดงจำนวนข้อมูลแต่ละเครือข่าย */
function setup() {
  const ss = SpreadsheetApp.getActive();
  NETWORKS.forEach(n => {
    if (n.key === 'unit') return; // ใช้บัญชีโรงเรียนเดิม
    if (!ss.getSheetByName(n.tab)) styleHeader_(ss.insertSheet(n.tab), n.fields.map(f => f[0]));
  });
  // ใส่ dropdown กลุ่มครู ก / ประเภท trip ให้กรอกง่าย
  const kru = ss.getSheetByName('ครูก');
  kru.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['ครู', 'ทหารตำรวจ', 'ปราชญ์ชาวบ้าน'], true).build());
  const trip = ss.getSheetByName('OVN-DAYTRIP');
  trip.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['OVN', 'DAYTRIP'], true).build());
  activitySheet_();
  rootFolder_();
  NETWORKS.forEach(n => Logger.log(`${n.tab}: ${readNetwork_(n).length} รายการ`));
}
