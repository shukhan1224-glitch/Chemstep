/**
 * 补习点名 + 收费提醒系统 —— 设定档
 *
 * 这个档案只放常数与共用小工具,不放商业逻辑。
 * 改栏位名称 / 新增出席状态,都从这里改。
 */

/** 各个工作表的名称 */
const SHEETS = {
  SETTINGS: '设置',
  STUDENTS: '学生',
  SESSIONS: '课程记录',
  PAYMENTS: '付款记录',
  ROLLCALL: '今日点名',
  DASHBOARD: '总览',
  REMINDER_LOG: '催费日志',
  AUDIT: '修改日志'
};

/** 出席状态 —— 你在点名时会选的选项 */
const ATTENDANCE_STATUS = {
  PRESENT: '出席',
  ABSENT_CHARGED: '缺席(照算)',
  EXCUSED: '请假(免扣)',
  TEACHER_CANCEL: '我取消(免扣)',
  HOLIDAY: '公假/停课(免扣)'
};

/**
 * 哪些状态要扣堂数。
 * 预设规则:学生自己不来 → 照算;你批准的请假、你自己取消、公假 → 免扣。
 * 这就是你说的「我可以选择让他缺一堂课」。
 */
const CHARGEABLE_BY_STATUS = {
  '出席': true,
  '缺席(照算)': true,
  '请假(免扣)': false,
  '我取消(免扣)': false,
  '公假/停课(免扣)': false
};

/** 学生状态 */
const STUDENT_STATUS = {
  ACTIVE: '在读',
  PAUSED: '暂停',
  ENDED: '结束'
};

const YES = '是';
const NO = '否';

/** 收款方式 —— 记录收款的视窗和「付款记录」的下拉都用这一份 */
const PAYMENT_METHODS = ['银行转账', 'DuitNow / QR', '现金', 'eWallet', '其他'];

/** 各表的栏位标题(顺序即栏序) */
const HEADERS = {
  STUDENTS: [
    '学生ID', '姓名', '班级', '家长称呼', 'WhatsApp号码',
    '每堂收费RM', '配套堂数', '状态', '备注'
  ],
  SESSIONS: [
    '记录ID', '日期', '班级', '学生ID', '学生姓名',
    '出席状态', '是否扣堂', '备注', '记录时间', '最后修改'
  ],
  PAYMENTS: [
    '付款ID', '日期', '学生ID', '学生姓名', '金额RM',
    '购买堂数', '付款方式', '备注', '记录时间'
  ],
  ROLLCALL: [
    '学生ID', '姓名', '目前余额', '欠款RM', '出席状态', '备注'
  ],
  DASHBOARD: [
    '学生ID', '姓名', '班级', '状态', '已付堂数', '已扣堂数', '余额', '欠款RM',
    '上次上课', '提醒状态', '发送催费', '已发送', '上次催费'
  ],
  REMINDER_LOG: ['时间', '学生ID', '学生姓名', '当时余额', '讯息内容'],
  AUDIT: ['时间', '工作表', '位置', '操作', '旧值', '新值']
};

/** 预设设定值。第一次初始化时会写进「设置」表,之后以表内的值为准。 */
const DEFAULT_SETTINGS = [
  ['老师名字', 'Teacher', '会出现在讯息署名'],
  ['每期配套堂数', 4, '一期几堂。你说的「补了4次」就是这个'],
  ['默认每堂收费RM', 50, '学生表没填收费时用这个'],
  ['催费门槛(余额≤)', 0, '余额小于或等于这个数 → 标红,该催费'],
  ['预告门槛(余额=)', 1, '余额等于这个数 → 标黄,温馨提示快用完'],
  ['国际区号', '60', 'wa.me 链接用。马来西亚是 60'],
  ['备份保留份数', 8, '自动备份最多留几份,超过的丢进垃圾桶'],
  ['催费讯息模板',
   '{家长}你好 👋\n\n{学生}的补习记录更新:目前这一期 {配套堂数} 堂课已经上完了(累计 {已扣堂数} 堂)。\n\n下一期 {配套堂数} 堂,费用 RM{金额}。方便的话麻烦安排缴费,谢谢你 🙏\n\n—— {老师}',
   '可用变数:{家长} {学生} {老师} {已扣堂数} {余额} {配套堂数} {金额} {每堂收费}'],
  ['预告讯息模板',
   '{家长}你好 👋\n\n温馨提示:{学生}的补习配套还剩 {余额} 堂。\n\n下一期 {配套堂数} 堂费用 RM{金额},到时再麻烦你安排,谢谢 🙏\n\n—— {老师}',
   '余额剩最后一堂时用的预告讯息']
];

/* ------------------------------------------------------------------ *
 * 共用小工具
 * ------------------------------------------------------------------ */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** 取得工作表,不存在就丢出清楚的错误(而不是 null pointer) */
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) {
    throw new Error('找不到工作表「' + name + '」。请先执行 补习管理 → 🔧 初始化 / 修复表格。');
  }
  return sh;
}

/** 依标题名称取得栏号(1-based)。找不到就报错,避免默默写错栏。 */
function colIdx_(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
  const i = headers.indexOf(headerName);
  if (i === -1) {
    throw new Error('工作表「' + sheet.getName() + '」找不到栏位「' + headerName + '」。请执行 🔧 初始化 / 修复表格。');
  }
  return i + 1;
}

/** 把整张表读成物件阵列,方便用栏名取值 */
function readTable_(sheetName) {
  const sh = sheet_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { headers: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], rows: [] };

  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const headers = values[0];
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = { _row: r + 1 };
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] !== '') obj[headers[c]] = values[r][c];
    }
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

/** 读取「设置」表,回传 key -> value 物件 */
function getSettings_() {
  const sh = sheet_(SHEETS.SETTINGS);
  const lastRow = sh.getLastRow();
  const out = {};
  if (lastRow < 2) return out;
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(function (row) {
    if (row[0] !== '') out[String(row[0]).trim()] = row[1];
  });
  return out;
}

/** 从「学生」表捞出所有班级名称(去重、排序),给点名的班级下拉用 */
function getClasses_() {
  const seen = {};
  readTable_(SHEETS.STUDENTS).rows.forEach(function (s) {
    const c = String(s['班级'] == null ? '' : s['班级']).trim();
    if (c && s['状态'] !== STUDENT_STATUS.ENDED) seen[c] = true;
  });
  return Object.keys(seen).sort();
}

function settingNumber_(settings, key, fallback) {
  const v = Number(settings[key]);
  return isNaN(v) || settings[key] === '' || settings[key] === undefined ? fallback : v;
}

/** 产生流水号,例如 S001 / L0001 / P001 */
function nextId_(sheetName, headerName, prefix, width) {
  const sh = sheet_(sheetName);
  const col = colIdx_(sh, headerName);
  const lastRow = sh.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
    values.forEach(function (r) {
      const m = String(r[0]).match(new RegExp('^' + prefix + '(\\d+)$'));
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  const n = String(max + 1);
  return prefix + (n.length >= width ? n : new Array(width - n.length + 1).join('0') + n);
}

/** 日期正规化成 yyyy-MM-dd 字串,用来比对「同一天有没有点过名」 */
function dateKey_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  if (isNaN(date.getTime())) return String(d).trim();
  return Utilities.formatDate(date, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

/** 把 012-345 6789 / +60123456789 之类的号码正规化成 wa.me 要的 60123456789 */
function normalizePhone_(raw, countryCode) {
  let digits = String(raw == null ? '' : raw).replace(/[^\d]/g, '');
  if (!digits) return '';
  const cc = String(countryCode || '60').replace(/[^\d]/g, '');
  if (digits.indexOf('00' + cc) === 0) digits = digits.substring(2);
  if (digits.indexOf(cc) === 0) return digits;
  if (digits.charAt(0) === '0') return cc + digits.substring(1);
  return cc + digits;
}

/** 写一笔修改日志(争议时可以翻旧账) */
function logAudit_(sheetName, a1, action, oldVal, newVal) {
  try {
    const sh = ss_().getSheetByName(SHEETS.AUDIT);
    if (!sh) return;
    sh.appendRow([new Date(), sheetName, a1, action, String(oldVal), String(newVal)]);
  } catch (e) {
    // 日志失败绝不能拖垮主流程
  }
}

function toast_(msg, title) {
  ss_().toast(msg, title || '补习管理', 6);
}
