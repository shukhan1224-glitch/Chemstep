/**
 * ══════════════════════════════════════════════════════════════
 *  线上补习点名 + 收费提醒系统  —— 合并版(单一档案)
 * ══════════════════════════════════════════════════════════════
 *
 *  这个档案由 tools/build-bundle.js 自动产生,请勿直接编辑。
 *  要改程式码请改 apps-script/ 底下的个别档案,再重新执行:
 *      node tools/build-bundle.js
 *
 *  安装:把整个档案内容贴进 Apps Script 编辑器,储存,
 *        重新整理试算表,然后点 📚 补习管理 → 🔧 初始化 / 修复表格。
 *
 *  合并自:00_Config.gs、01_Setup.gs、02_Attendance.gs、03_Dashboard.gs、04_Reminders.gs、05_Triggers.gs、06_Backup.gs、07_Payments.gs
 */

/* ── 00_Config.gs ──────────────────────────────────────────────── */

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


/* ── 01_Setup.gs ───────────────────────────────────────────────── */

/**
 * 初始化 / 修复表格结构。
 *
 * 这个函式设计成可以重复执行:已经存在的表不会被清空,
 * 只会补上缺少的工作表、标题列、下拉选单与格式。
 */
function setupWorkbook() {
  const ss = ss_();

  setupSettingsSheet_(ss);
  setupStudentsSheet_(ss);
  setupSessionsSheet_(ss);
  setupPaymentsSheet_(ss);
  setupRollcallSheet_(ss);
  setupDashboardSheet_(ss);
  setupPlainSheet_(ss, SHEETS.REMINDER_LOG, HEADERS.REMINDER_LOG);
  setupPlainSheet_(ss, SHEETS.AUDIT, HEADERS.AUDIT);

  // 删掉 Google 预设的空白 "Sheet1" / "工作表1"。
  // 有资料的工作表一律保留 —— 由旧 Excel 转过来的旧分页会原封不动留着。
  const kept = [];
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    const known = Object.keys(SHEETS).some(function (k) { return SHEETS[k] === name; });
    if (!known) {
      if (sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
      else kept.push(name);
    }
  });

  // 排序:常用的放前面
  const order = [SHEETS.DASHBOARD, SHEETS.ROLLCALL, SHEETS.STUDENTS, SHEETS.SESSIONS,
                 SHEETS.PAYMENTS, SHEETS.SETTINGS, SHEETS.REMINDER_LOG, SHEETS.AUDIT];
  order.forEach(function (name, i) {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));

  // 一定要跟着重画总览:上面只换了标题列,底下还是上一次的资料。
  // 栏位数量若有变动(例如版本更新新增了栏),不重画就会整排错位。
  try {
    refreshDashboard();
  } catch (e) {
    logAudit_(SHEETS.DASHBOARD, '', '初始化后重画总览失败', '', e.message);
  }

  SpreadsheetApp.getUi().alert(
    '✅ 初始化完成',
    '所有工作表已建立。\n\n下一步:\n' +
    '1. 到「设置」填你的名字、配套堂数、收费\n' +
    '2. 到「学生」加学生(姓名和班级填了就会自动配学生ID)\n' +
    '3. 上完课 → 今日点名选好日期和班级 → 载入今日点名\n\n' +
    (kept.length
      ? '📁 你原有的工作表已原封不动保留:\n   ' + kept.join('、') + '\n\n'
      : '') +
    '提醒:讯息不会自动送出,系统只会帮你写好草稿 + 产生 WhatsApp 链接,由你按下才发。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * 检查工作表的标题列跟程式里的定义是否一致。
 *
 * 更新到新版程式码后,如果没有接着执行「初始化 / 修复表格」,
 * 标题还是旧的、写进去的资料却是新的栏数,整排就会往左错位。
 * 会用到这个检查的地方会自己修好,不必依赖使用者记得跑哪一步。
 */
function headersStale_(sheetName, headers, headerRow) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh) return true;
  const row = headerRow || 1;
  if (sh.getLastRow() < row) return true;

  const width = Math.max(sh.getLastColumn(), headers.length);
  const current = sh.getRange(row, 1, 1, width).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(current[i] == null ? '' : current[i]).trim() !== headers[i]) return true;
  }
  return false;
}

/** 写入标题列并冻结 */
function writeHeaders_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#37474f')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > headers.length) {
    sh.deleteColumns(headers.length + 1, sh.getMaxColumns() - headers.length);
  }
}

function setupPlainSheet_(ss, name, headers) {
  const sh = getOrCreateSheet_(ss, name);
  writeHeaders_(sh, headers);
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function setupSettingsSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEETS.SETTINGS);
  writeHeaders_(sh, ['设定项目', '值', '说明']);

  // 只补上「还没有」的设定项,不覆盖使用者已经改过的值
  const existing = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
      if (r[0] !== '') existing[String(r[0]).trim()] = i + 2;
    });
  }
  DEFAULT_SETTINGS.forEach(function (row) {
    if (!existing[row[0]]) sh.appendRow(row);
  });

  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 380);
  sh.getRange(2, 2, Math.max(sh.getLastRow() - 1, 1), 2).setWrap(true).setVerticalAlignment('top');
  return sh;
}

function setupStudentsSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.STUDENTS, HEADERS.STUDENTS);
  const rows = Math.max(sh.getMaxRows() - 1, 1);

  applyDropdown_(sh, colIdx_(sh, '状态'), rows,
    [STUDENT_STATUS.ACTIVE, STUDENT_STATUS.PAUSED, STUDENT_STATUS.ENDED]);

  sh.getRange(2, colIdx_(sh, 'WhatsApp号码'), rows, 1).setNumberFormat('@'); // 保留开头的 0
  sh.getRange(2, colIdx_(sh, '每堂收费RM'), rows, 1).setNumberFormat('0.00');
  setWidths_(sh, {
    '学生ID': 70, '姓名': 110, '班级': 100, '家长称呼': 110,
    'WhatsApp号码': 130, '每堂收费RM': 100, '配套堂数': 80, '状态': 70, '备注': 260
  });

  // 还没填 WhatsApp 号码的标黄 —— 没有号码就发不出催费讯息
  const phone = sh.getRange(2, colIdx_(sh, 'WhatsApp号码'), rows, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty().setBackground('#fff8e1')
      .setRanges([phone]).build()
  ]);
  return sh;
}

/** 依栏位名称设定栏宽 */
function setWidths_(sh, widths) {
  Object.keys(widths).forEach(function (name) {
    sh.setColumnWidth(colIdx_(sh, name), widths[name]);
  });
}

function setupSessionsSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.SESSIONS, HEADERS.SESSIONS);
  const rows = Math.max(sh.getMaxRows() - 1, 1);

  applyDropdown_(sh, colIdx_(sh, '出席状态'), rows, objValues_(ATTENDANCE_STATUS));
  applyDropdown_(sh, colIdx_(sh, '是否扣堂'), rows, [YES, NO]);
  // 不设定对齐方式:真日期会自动靠右、被当成文字的会靠左,一眼看得出输入错误
  sh.getRange(2, colIdx_(sh, '日期'), rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, colIdx_(sh, '记录时间'), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, colIdx_(sh, '最后修改'), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');

  // 免扣的那几堂用灰色底,一眼看得出来
  const chargeCol = sh.getRange(2, colIdx_(sh, '是否扣堂'), rows, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(NO)
    .setBackground('#eceff1')
    .setFontColor('#78909c')
    .setRanges([chargeCol])
    .build();
  sh.setConditionalFormatRules([rule]);
  return sh;
}

function setupPaymentsSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.PAYMENTS, HEADERS.PAYMENTS);
  const rows = Math.max(sh.getMaxRows() - 1, 1);

  applyDropdown_(sh, colIdx_(sh, '付款方式'), rows, PAYMENT_METHODS.concat(['旧档结转']));
  sh.getRange(2, colIdx_(sh, '日期'), rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, colIdx_(sh, '金额RM'), rows, 1).setNumberFormat('0.00');
  sh.getRange(2, colIdx_(sh, '记录时间'), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  return sh;
}

function setupRollcallSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEETS.ROLLCALL);

  // 第 1 列是「上课日期 + 班级」选择区,第 2 列才是标题
  sh.getRange('A1').setValue('上课日期 →').setFontWeight('bold');
  sh.getRange('B1').setNumberFormat('yyyy-mm-dd').setBackground('#fff8e1')
    .setBorder(true, true, true, true, false, false);
  sh.getRange('C1').setValue('班级 →').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('D1').setBackground('#fff8e1')
    .setBorder(true, true, true, true, false, false);
  sh.getRange('E1').setValue('← 先选好这两格,再按「载入今日点名」').setFontColor('#90a4ae');

  sh.getRange(2, 1, 1, HEADERS.ROLLCALL.length).setValues([HEADERS.ROLLCALL])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(2);
  if (sh.getMaxColumns() > HEADERS.ROLLCALL.length) {
    sh.deleteColumns(HEADERS.ROLLCALL.length + 1, sh.getMaxColumns() - HEADERS.ROLLCALL.length);
  }

  const rows = Math.max(sh.getMaxRows() - 2, 1);
  const cStatus = HEADERS.ROLLCALL.indexOf('出席状态') + 1;
  const cOwed = HEADERS.ROLLCALL.indexOf('欠款RM') + 1;
  sh.getRange(3, cStatus, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(objValues_(ATTENDANCE_STATUS), true)
      .setAllowInvalid(false).build()
  );
  sh.setColumnWidth(1, 70);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(cOwed, 90);
  sh.setColumnWidth(cStatus, 150);
  sh.setColumnWidth(cStatus + 1, 240);
  sh.getRange(3, cOwed, rows, 1).setNumberFormat('#,##0.00');

  // 有欠款的整格标红,点名时一眼看到
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
      .setRanges([sh.getRange(3, cOwed, rows, 1)]).build()
  ]);

  refreshClassDropdown_();
  return sh;
}

/** 把「今日点名」D1 的班级下拉,依「学生」表现有的班级重建 */
function refreshClassDropdown_() {
  const sh = ss_().getSheetByName(SHEETS.ROLLCALL);
  if (!sh) return;
  const classes = getClasses_();
  const cell = sh.getRange('D1');
  if (!classes.length) { cell.clearDataValidations(); return; }
  cell.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(classes, true)
      .setAllowInvalid(false).build()
  );
  if (classes.indexOf(String(cell.getValue()).trim()) === -1) cell.clearContent();
}

function setupDashboardSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.DASHBOARD, HEADERS.DASHBOARD);
  setWidths_(sh, {
    '学生ID': 70, '姓名': 110, '班级': 100, '状态': 60,
    '已付堂数': 70, '已扣堂数': 70, '余额': 60, '欠款RM': 90, '上次上课': 100,
    '提醒状态': 110, '发送催费': 130, '已发送': 70, '上次催费': 130
  });
  sh.getRange(2, colIdx_(sh, '上次上课'), Math.max(sh.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, colIdx_(sh, '上次催费'), Math.max(sh.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');
  return sh;
}

function applyDropdown_(sh, col, numRows, options) {
  sh.getRange(2, col, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
      .build()
  );
}

function objValues_(obj) {
  return Object.keys(obj).map(function (k) { return obj[k]; });
}


/* ── 02_Attendance.gs ──────────────────────────────────────────── */

/**
 * 点名流程。
 *
 * 设计原则:点名要三秒内做完 —— 载入 → 改几个下拉 → 提交。
 * 大部分学生都是「出席」,所以预设就是出席,你只需要动有状况的那几个。
 */

/** 载入某个班级的在读学生到「今日点名」 */
function loadRollcall() {
  const ui = SpreadsheetApp.getUi();

  // 版本更新后标题列可能还是旧的,先自己修好,避免资料写进去错位
  if (headersStale_(SHEETS.ROLLCALL, HEADERS.ROLLCALL, 2)) {
    setupRollcallSheet_(ss_());
    toast_('点名表的栏位已更新到最新版。');
  }

  const sh = sheet_(SHEETS.ROLLCALL);
  refreshClassDropdown_();

  const klass = String(sh.getRange('D1').getValue()).trim();
  if (!klass) {
    ss_().setActiveSheet(sh);
    const classes = getClasses_();
    ui.alert('请先选班级',
      classes.length
        ? '在「今日点名」的 D1 选一个班级,再按一次「载入今日点名」。\n\n目前有的班级:\n• ' + classes.join('\n• ')
        : '「学生」表里还没有任何班级。请先在学生表填上「班级」栏。',
      ui.ButtonSet.OK);
    return;
  }

  const students = readTable_(SHEETS.STUDENTS).rows.filter(function (s) {
    return s['学生ID'] !== '' &&
           s['状态'] === STUDENT_STATUS.ACTIVE &&
           String(s['班级']).trim() === klass;
  });

  // 清掉上一轮残留
  if (sh.getMaxRows() > 2) {
    sh.getRange(3, 1, sh.getMaxRows() - 2, HEADERS.ROLLCALL.length)
      .clearContent().setBackground(null);
  }

  if (!students.length) {
    ss_().setActiveSheet(sh);
    ui.alert('「' + klass + '」这个班没有状态为「在读」的学生。');
    return;
  }

  // 日期空的话预设今天
  const dateCell = sh.getRange('B1');
  if (!dateCell.getValue()) dateCell.setValue(new Date());
  const targetKey = dateKey_(dateCell.getValue());

  // 当天已经点过名的学生,直接带出原本的状态,避免重复记录
  const existing = {};
  readTable_(SHEETS.SESSIONS).rows.forEach(function (r) {
    if (dateKey_(r['日期']) === targetKey) existing[r['学生ID']] = r['出席状态'];
  });

  const settings = getSettings_();
  const defaultFee = settingNumber_(settings, '默认每堂收费RM', 50);
  const balances = computeBalances_();
  const out = students.map(function (s) {
    const b = balances[s['学生ID']] || { balance: 0 };
    const perLesson = Number(s['每堂收费RM']) || defaultFee;
    return [
      s['学生ID'],
      s['姓名'],
      b.balance,
      b.balance < 0 ? -b.balance * perLesson : '',   // 点名时就看得到谁欠钱
      existing[s['学生ID']] || ATTENDANCE_STATUS.PRESENT,
      existing[s['学生ID']] ? '⚠️ 当天已有记录,提交时会略过' : ''
    ];
  });

  const cStatus = HEADERS.ROLLCALL.indexOf('出席状态') + 1;
  sh.getRange(3, 1, out.length, HEADERS.ROLLCALL.length).setValues(out);
  sh.getRange(3, 1, out.length, cStatus - 1).setBackground('#f5f5f5'); // 唯读区域淡灰
  sh.getRange(3, cStatus, out.length, 2).setBackground('#ffffff');
  ss_().setActiveSheet(sh);

  toast_(klass + ':已载入 ' + out.length + ' 位学生,预设全部「出席」。' +
         '改完有状况的那几个再按「提交点名」。');
}

/**
 * 撤销某一次点名(整班)。
 * 用「今日点名」的 B1 日期 + D1 班级当条件,把那次的记录整批删掉。
 * 删之前会先列出内容让你确认,删除的事实会写进「修改日志」。
 *
 * 只有一个学生算错的话,不必用这个 —— 直接到「课程记录」改他那一行的
 * 出席状态就好,余额会自动重算。
 */
function undoRollcall() {
  const ui = SpreadsheetApp.getUi();
  const rc = sheet_(SHEETS.ROLLCALL);

  const rawDate = rc.getRange('B1').getValue();
  const klass = String(rc.getRange('D1').getValue()).trim();
  if (!rawDate || !klass) {
    ss_().setActiveSheet(rc);
    ui.alert('请先在「今日点名」填好 B1 上课日期、D1 班级,再执行撤销。');
    return;
  }
  const date = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
  if (isNaN(date.getTime())) { ui.alert('B1 的日期看不懂。'); return; }
  const key = dateKey_(date);

  const hits = readTable_(SHEETS.SESSIONS).rows.filter(function (r) {
    return dateKey_(r['日期']) === key && String(r['班级']).trim() === klass;
  });
  if (!hits.length) {
    ui.alert('找不到「' + klass + '」在 ' + key + ' 的点名记录。\n\n' +
             '确认一下 B1 的日期和 D1 的班级对不对。');
    return;
  }

  const charged = hits.filter(function (r) {
    return String(r['是否扣堂']).trim() === YES;
  }).length;
  const list = hits.map(function (r) {
    return r['学生姓名'] + '(' + r['出席状态'] + ')';
  }).join('\n• ');

  const answer = ui.alert(
    '确定要撤销这次点名吗?',
    klass + ' · ' + key + '\n\n' +
    '共 ' + hits.length + ' 笔,其中 ' + charged + ' 笔有扣堂:\n• ' + list + '\n\n' +
    '删除后这 ' + charged + ' 堂会还给学生(余额各加 1)。\n' +
    '这个动作会记进「修改日志」。',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  // 由下往上删,不然删掉一行之后下面的列号会往上跑
  const sh = sheet_(SHEETS.SESSIONS);
  hits.map(function (r) { return r._row; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (row) { sh.deleteRow(row); });

  logAudit_(SHEETS.SESSIONS, klass + ' ' + key, '撤销整次点名',
            hits.length + ' 笔(扣堂 ' + charged + ')', '已删除');
  refreshDashboard();

  ui.alert('✅ 已撤销',
           klass + ' · ' + key + '\n\n' +
           '删除 ' + hits.length + ' 笔,' + charged + ' 堂已还给学生。',
           ui.ButtonSet.OK);
}

/** 把「今日点名」的内容写进「课程记录」 */
function submitRollcall() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SHEETS.ROLLCALL);
  const lastRow = sh.getLastRow();

  if (lastRow < 3) {
    ui.alert('点名表是空的。请先执行「载入今日点名」。');
    return;
  }

  // 表上的资料若是旧版栏位排出来的,直接提交会存错栏 —— 挡下来要求重载
  if (headersStale_(SHEETS.ROLLCALL, HEADERS.ROLLCALL, 2)) {
    ui.alert('点名表的栏位是旧版的',
             '请先执行「✅ 载入今日点名」重新载入一次(会自动更新栏位),再提交。',
             ui.ButtonSet.OK);
    return;
  }

  const rawDate = sh.getRange('B1').getValue();
  if (!rawDate) { ui.alert('请先在 B1 填上课日期。'); return; }
  const date = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
  if (isNaN(date.getTime())) { ui.alert('B1 的日期看不懂,请用日期格式(例如 2026-08-07)。'); return; }
  const targetKey = dateKey_(date);

  const klass = String(sh.getRange('D1').getValue()).trim();
  if (!klass) { ui.alert('请先在 D1 选班级。'); return; }

  // 已存在的记录:同一个学生、同一天,只算一次
  const already = {};
  readTable_(SHEETS.SESSIONS).rows.forEach(function (r) {
    if (dateKey_(r['日期']) === targetKey) already[r['学生ID']] = true;
  });

  const data = sh.getRange(3, 1, lastRow - 2, HEADERS.ROLLCALL.length).getValues();
  const cStatus = HEADERS.ROLLCALL.indexOf('出席状态');
  const now = new Date();
  const toAppend = [];
  const skipped = [];
  const counts = {};
  let seq = null;

  data.forEach(function (row) {
    const id = String(row[0]).trim();
    const name = row[1];
    const status = String(row[cStatus]).trim();
    const note = row[cStatus + 1];

    if (!id || !status) return;
    if (!(status in CHARGEABLE_BY_STATUS)) {
      skipped.push(name + '(出席状态「' + status + '」无法辨识)');
      return;
    }
    if (already[id]) { skipped.push(name + '(' + targetKey + ' 已有记录)'); return; }
    already[id] = true; // 同一次提交里若名单重复,也只记一笔

    if (seq === null) seq = parseInt(nextId_(SHEETS.SESSIONS, '记录ID', 'L', 4).substring(1), 10);
    const recId = 'L' + ('000' + seq).slice(-4);
    seq++;

    const chargeable = CHARGEABLE_BY_STATUS[status] ? YES : NO;
    counts[status] = (counts[status] || 0) + 1;

    toAppend.push([
      recId, date, klass, id, name,
      status, chargeable,
      String(note).indexOf('⚠️') === 0 ? '' : note,
      now, ''
    ]);
  });

  if (!toAppend.length) {
    ui.alert('没有新的记录可以提交。' + (skipped.length ? '\n\n略过:\n• ' + skipped.join('\n• ') : ''));
    return;
  }

  const sessions = sheet_(SHEETS.SESSIONS);
  sessions.getRange(sessions.getLastRow() + 1, 1, toAppend.length, HEADERS.SESSIONS.length)
          .setValues(toAppend);
  logAudit_(SHEETS.SESSIONS, targetKey + ' ' + klass, '提交点名', '', toAppend.length + ' 笔');

  refreshDashboard();

  // 提交完直接告诉你「谁该催费了」—— 这才是这个系统存在的理由
  const due = getStudentsNeedingReminder_();
  let msg = '✅ ' + klass + ' · ' + targetKey + ' 已记录 ' + toAppend.length + ' 笔。\n\n';
  Object.keys(counts).forEach(function (k) { msg += '• ' + k + ':' + counts[k] + '\n'; });
  const charged = toAppend.filter(function (r) { return r[6] === YES; }).length;
  msg += '\n其中扣堂 ' + charged + ' 堂、免扣 ' + (toAppend.length - charged) + ' 堂。';

  if (skipped.length) msg += '\n\n⚠️ 略过:\n• ' + skipped.join('\n• ');

  if (due.length) {
    msg += '\n\n🔴 以下学生堂数用完,该催费了:\n• ' +
           due.map(function (d) { return d.name + '(' + d.klass + ',余额 ' + d.balance + ')'; }).join('\n• ') +
           '\n\n到「总览」按「📱 发送催费」就会带出写好的 WhatsApp 讯息。';
  } else {
    msg += '\n\n🟢 目前没有人需要催费。';
  }

  ui.alert('点名完成', msg, ui.ButtonSet.OK);

  // 清空点名区,避免下次误按重复提交
  sh.getRange(3, 1, sh.getMaxRows() - 2, HEADERS.ROLLCALL.length).clearContent();
  ss_().setActiveSheet(sheet_(SHEETS.DASHBOARD));
}


/* ── 03_Dashboard.gs ───────────────────────────────────────────── */

/**
 * 余额计算 + 总览表。
 *
 * 核心公式:  余额 = 已付堂数 - 已扣堂数
 *
 * 「补了4次要催费」在这个模型下 = 买了 4 堂的配套用完了(余额 ≤ 0)。
 * 好处是学生预付 8 堂、单堂付、欠账,通通不用改逻辑。
 */

/** 算出每个学生的 已付 / 已扣 / 余额 / 上次上课 */
function computeBalances_() {
  const result = {};

  readTable_(SHEETS.STUDENTS).rows.forEach(function (s) {
    if (s['学生ID'] === '') return;
    result[s['学生ID']] = { paid: 0, charged: 0, balance: 0, lastSession: null };
  });

  readTable_(SHEETS.PAYMENTS).rows.forEach(function (p) {
    const r = result[p['学生ID']];
    if (!r) return;
    const n = Number(p['购买堂数']);
    if (!isNaN(n)) r.paid += n;
  });

  readTable_(SHEETS.SESSIONS).rows.forEach(function (l) {
    const r = result[l['学生ID']];
    if (!r) return;
    if (String(l['是否扣堂']).trim() === YES) r.charged += 1;
    const d = (l['日期'] instanceof Date) ? l['日期'] : new Date(l['日期']);
    if (!isNaN(d.getTime()) && (!r.lastSession || d > r.lastSession)) r.lastSession = d;
  });

  Object.keys(result).forEach(function (id) {
    result[id].balance = result[id].paid - result[id].charged;
  });
  return result;
}

/** 上次对某个学生发催费讯息的时间 */
function lastReminderMap_() {
  const map = {};
  readTable_(SHEETS.REMINDER_LOG).rows.forEach(function (r) {
    const t = (r['时间'] instanceof Date) ? r['时间'] : new Date(r['时间']);
    if (isNaN(t.getTime())) return;
    if (!map[r['学生ID']] || t > map[r['学生ID']]) map[r['学生ID']] = t;
  });
  return map;
}

/** 重算并重画「总览」 */
function refreshDashboard() {
  // 同上:标题列若还是旧版,先修好再画,不然整排会错位
  if (headersStale_(SHEETS.DASHBOARD, HEADERS.DASHBOARD, 1)) {
    setupDashboardSheet_(ss_());
  }

  const sh = sheet_(SHEETS.DASHBOARD);
  const settings = getSettings_();
  const balances = computeBalances_();
  const lastReminder = lastReminderMap_();

  const dueThreshold = settingNumber_(settings, '催费门槛(余额≤)', 0);
  const warnThreshold = settingNumber_(settings, '预告门槛(余额=)', 1);

  // 依 班级 → 余额 排序:同班的排在一起,快没堂数的排最上面
  const students = readTable_(SHEETS.STUDENTS).rows
    .filter(function (s) {
      return s['学生ID'] !== '' && s['状态'] !== STUDENT_STATUS.ENDED;
    })
    .sort(function (a, b) {
      const ka = String(a['班级'] || ''), kb = String(b['班级'] || '');
      if (ka !== kb) return ka < kb ? -1 : 1;
      const ba = (balances[a['学生ID']] || {}).balance || 0;
      const bb = (balances[b['学生ID']] || {}).balance || 0;
      return ba - bb;
    });

  // 清掉旧内容(含核取方块的验证规则)
  if (sh.getMaxRows() > 1) {
    const old = sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.DASHBOARD.length);
    old.clearContent().clearDataValidations().setBackground(null).setFontColor(null);
  }
  if (!students.length) { toast_('还没有学生资料。'); return; }

  let totalOwed = 0;
  let dueCount = 0;

  const rows = students.map(function (s) {
    const id = s['学生ID'];
    const b = balances[id] || { paid: 0, charged: 0, balance: 0, lastSession: null };
    const state = b.balance <= dueThreshold ? '🔴 该催费'
                : b.balance <= warnThreshold ? '🟡 快用完'
                : '🟢 正常';
    const link = b.balance <= warnThreshold ? buildWaFormula_(s, b, settings) : '';

    // 欠款 = 欠的堂数 × 该学生的每堂收费(不同班收费不同)
    const perLesson = Number(s['每堂收费RM']) ||
                      settingNumber_(settings, '默认每堂收费RM', 50);
    const owed = b.balance < 0 ? -b.balance * perLesson : 0;
    totalOwed += owed;
    if (b.balance <= dueThreshold) dueCount++;

    return [
      id, s['姓名'], s['班级'], s['状态'],
      b.paid, b.charged, b.balance, owed || '',
      b.lastSession || '',
      state, link, false,
      lastReminder[id] || ''
    ];
  });

  sh.getRange(2, 1, rows.length, HEADERS.DASHBOARD.length).setValues(rows);

  // 「已发送」做成核取方块 —— 勾了就自动写进催费日志
  sh.getRange(2, colIdx_(sh, '已发送'), rows.length, 1).insertCheckboxes();

  sh.getRange(2, colIdx_(sh, '上次上课'), rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, colIdx_(sh, '上次催费'), rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, colIdx_(sh, '欠款RM'), rows.length, 1).setNumberFormat('#,##0.00');

  applyDashboardColors_(sh, rows.length);
  refreshClassDropdown_();

  toast_('总览已更新:' + rows.length + ' 位学生 · ' + getClasses_().length + ' 个班' +
         (dueCount
           ? ' · 🔴 ' + dueCount + ' 位该催费,共欠 RM' + totalOwed.toFixed(2)
           : ' · 🟢 无人需要催费'));
}

function applyDashboardColors_(sh, numRows) {
  const stateRange = sh.getRange(2, colIdx_(sh, '提醒状态'), numRows, 1);
  const balanceRange = sh.getRange(2, colIdx_(sh, '余额'), numRows, 1);
  const owedRange = sh.getRange(2, colIdx_(sh, '欠款RM'), numRows, 1);

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('该催费').setBackground('#ffcdd2').setFontColor('#b71c1c')
      .setRanges([stateRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('快用完').setBackground('#fff9c4').setFontColor('#f57f17')
      .setRanges([stateRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('正常').setBackground('#c8e6c9').setFontColor('#1b5e20')
      .setRanges([stateRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(0).setFontColor('#b71c1c').setBold(true)
      .setRanges([balanceRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor('#b71c1c').setBold(true)
      .setRanges([owedRange]).build()
  ]);
}

/** 找出目前该催费的学生(点名完会用到) */
function getStudentsNeedingReminder_() {
  const settings = getSettings_();
  const dueThreshold = settingNumber_(settings, '催费门槛(余额≤)', 0);
  const balances = computeBalances_();

  return readTable_(SHEETS.STUDENTS).rows
    .filter(function (s) {
      if (s['学生ID'] === '' || s['状态'] !== STUDENT_STATUS.ACTIVE) return false;
      const b = balances[s['学生ID']];
      return b && b.balance <= dueThreshold;
    })
    .map(function (s) {
      return {
        id: s['学生ID'], name: s['姓名'], klass: s['班级'],
        balance: balances[s['学生ID']].balance
      };
    });
}


/* ── 04_Reminders.gs ───────────────────────────────────────────── */

/**
 * 催费讯息:产生草稿 + WhatsApp 链接。
 *
 * 刻意「不自动送出」。系统只负责算清楚 + 写好文案,
 * 按下送出的永远是你 —— 催费发错人是社交灾难,不是 bug ticket。
 */

/** 依学生资料算出讯息里要代入的数字 */
function reminderContext_(student, balance, settings) {
  const perLesson = Number(student['每堂收费RM']) ||
                    settingNumber_(settings, '默认每堂收费RM', 50);
  const packageSize = Number(student['配套堂数']) ||
                      settingNumber_(settings, '每期配套堂数', 4);
  return {
    '{家长}': student['家长称呼'] || student['姓名'] + '家长',
    '{学生}': student['姓名'],
    '{老师}': settings['老师名字'] || 'Teacher',
    '{已扣堂数}': balance.charged,
    '{余额}': balance.balance,
    '{配套堂数}': packageSize,
    '{每堂收费}': perLesson.toFixed(2),
    '{金额}': (perLesson * packageSize).toFixed(2)
  };
}

/** 套入模板产生讯息文字 */
function renderReminder_(student, balance, settings) {
  const dueThreshold = settingNumber_(settings, '催费门槛(余额≤)', 0);
  const template = balance.balance <= dueThreshold
    ? settings['催费讯息模板']
    : settings['预告讯息模板'];

  let text = String(template || '{家长}你好,{学生}的补习堂数用完了,麻烦安排缴费,谢谢。—— {老师}');
  const ctx = reminderContext_(student, balance, settings);
  Object.keys(ctx).forEach(function (key) {
    text = text.split(key).join(ctx[key]);
  });
  return text;
}

/** 产生总览里那颗「📱 发送催费」链接 */
function buildWaFormula_(student, balance, settings) {
  const phone = normalizePhone_(student['WhatsApp号码'], settings['国际区号']);
  if (!phone) return '⚠️ 缺 WhatsApp 号码';
  const url = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(renderReminder_(student, balance, settings));
  return '=HYPERLINK("' + url + '", "📱 发送催费")';
}

/**
 * 打开一个视窗,列出所有该催费的学生 + 讯息草稿。
 * 比在储存格里点链接好用:可以先读一遍文案再决定发不发。
 */
function showReminderDrafts() {
  const settings = getSettings_();
  const balances = computeBalances_();
  const warnThreshold = settingNumber_(settings, '预告门槛(余额=)', 1);
  const dueThreshold = settingNumber_(settings, '催费门槛(余额≤)', 0);

  const items = readTable_(SHEETS.STUDENTS).rows
    .filter(function (s) {
      if (s['学生ID'] === '' || s['状态'] !== STUDENT_STATUS.ACTIVE) return false;
      const b = balances[s['学生ID']];
      return b && b.balance <= warnThreshold;
    })
    .map(function (s) {
      const b = balances[s['学生ID']];
      const phone = normalizePhone_(s['WhatsApp号码'], settings['国际区号']);
      const text = renderReminder_(s, b, settings);
      return {
        id: s['学生ID'],
        name: s['姓名'],
        klass: s['班级'],
        balance: b.balance,
        urgent: b.balance <= dueThreshold,
        phone: phone,
        text: text,
        url: phone ? 'https://wa.me/' + phone + '?text=' + encodeURIComponent(text) : ''
      };
    })
    .sort(function (a, b) {
      if (a.balance !== b.balance) return a.balance - b.balance;  // 最急的排最前
      return String(a.klass) < String(b.klass) ? -1 : 1;
    });

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(reminderDialogHtml_(items)).setWidth(520).setHeight(600),
    '催费讯息草稿'
  );
}

function reminderDialogHtml_(items) {
  const esc = function (s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  };

  let body = '';
  if (!items.length) {
    body = '<div class="empty">🟢 目前没有人需要催费或预告。</div>';
  } else {
    items.forEach(function (it) {
      body +=
        '<div class="card ' + (it.urgent ? 'urgent' : 'warn') + '">' +
          '<div class="head">' +
            '<span class="name">' + esc(it.name) +
              (it.klass ? ' <span class="klass">' + esc(it.klass) + '</span>' : '') + '</span>' +
            '<span class="badge">余额 ' + it.balance + ' 堂</span>' +
          '</div>' +
          '<pre>' + esc(it.text) + '</pre>' +
          (it.url
            ? '<a class="btn" target="_blank" href="' + esc(it.url) + '">📱 在 WhatsApp 打开</a>'
            : '<div class="nophone">⚠️ 这位学生没填 WhatsApp 号码</div>') +
        '</div>';
    });
  }

  return '' +
    '<style>' +
    'body{font-family:-apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;' +
      'margin:0;padding:14px;background:#fafafa;color:#263238;}' +
    '.tip{font-size:12px;color:#78909c;margin-bottom:12px;line-height:1.5;}' +
    '.card{background:#fff;border-radius:10px;padding:12px 14px;margin-bottom:12px;' +
      'border-left:4px solid #ffb300;box-shadow:0 1px 3px rgba(0,0,0,.12);}' +
    '.card.urgent{border-left-color:#e53935;}' +
    '.head{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}' +
    '.name{font-weight:700;font-size:15px;}' +
    '.klass{font-weight:400;font-size:12px;color:#78909c;}' +
    '.badge{font-size:12px;background:#eceff1;border-radius:10px;padding:2px 9px;color:#546e7a;}' +
    'pre{white-space:pre-wrap;word-break:break-word;background:#f5f7f8;border-radius:6px;' +
      'padding:10px;font-size:13px;line-height:1.6;margin:0 0 10px;font-family:inherit;}' +
    '.btn{display:inline-block;background:#25d366;color:#fff;text-decoration:none;' +
      'padding:8px 16px;border-radius:6px;font-size:13px;font-weight:600;}' +
    '.nophone{font-size:12px;color:#e53935;}' +
    '.empty{text-align:center;padding:40px 0;color:#78909c;}' +
    '</style>' +
    '<div class="tip">讯息不会自动送出。按下按钮会打开 WhatsApp 并把文字填好,' +
      '你可以再改再送。送出后回到「总览」勾选「已发送」,系统就会记进催费日志。</div>' +
    body;
}


/* ── 05_Triggers.gs ────────────────────────────────────────────── */

/**
 * 选单 + 自动化(onEdit)。
 *
 * onEdit 负责所有「你不该手动做的杂事」:
 *   • 学生表打了姓名 → 自动配学生ID、预设在读
 *   • 付款表填了学生ID → 自动带姓名、盖时间戳
 *   • 课程记录改了出席状态 → 自动重算是否扣堂,并写进修改日志
 *   • 总览勾了「已发送」→ 自动写催费日志
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 补习管理')
    .addItem('✅ 载入今日点名', 'loadRollcall')
    .addItem('📥 提交点名', 'submitRollcall')
    .addItem('↩️ 撤销某次点名', 'undoRollcall')
    .addSeparator()
    .addItem('💵 记录收款', 'showPaymentDialog')
    .addItem('📱 查看催费草稿', 'showReminderDrafts')
    .addItem('🔄 刷新总览', 'refreshDashboard')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('💾 备份')
      .addItem('立刻备份一份', 'backupNowWithAlert')
      .addItem('开启每周自动备份', 'installWeeklyBackup')
      .addItem('关闭每周自动备份', 'removeWeeklyBackup'))
    .addSeparator()
    .addItem('🔧 初始化 / 修复表格', 'setupWorkbook')
    .addItem('❓ 使用说明', 'showHelp')
    .addToUi();
}

/**
 * 备案:如果重新整理后「📚 补习管理」还是不出现,在 Apps Script 里执行这个函式。
 *
 * onOpen 属于「简单触发器」,某些情况下(公司/学校的 Google Workspace 政策、
 * 浏览器扩充套件、档案是从别人那里复制来的)不会自动执行。
 * 改装成「可安装触发器」就会稳定跑。
 */
function installOpenTrigger() {
  const ss = ss_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onOpen') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();

  // 顺便立刻建一次,不必等重新整理
  onOpen();
  Logger.log('✅ 已为「' + ss.getName() + '」安装开启触发器\n' +
             '网址:' + ss.getUrl() + '\n' +
             '回试算表重新整理(F5),选单就会出现。');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();

  try {
    if (name === SHEETS.STUDENTS) handleStudentEdit_(sh, e);
    else if (name === SHEETS.PAYMENTS) handlePaymentEdit_(sh, e);
    else if (name === SHEETS.SESSIONS) handleSessionEdit_(sh, e);
    else if (name === SHEETS.DASHBOARD) handleDashboardEdit_(sh, e);
  } catch (err) {
    // onEdit 不能弹错误视窗打断你打字,记进日志就好
    logAudit_(name, e.range.getA1Notation(), 'onEdit 错误', '', err.message);
  }
}

function editedRows_(e) {
  const start = e.range.getRow();
  const rows = [];
  for (let i = 0; i < e.range.getNumRows(); i++) {
    if (start + i >= 2) rows.push(start + i);
  }
  return rows;
}

/** 学生表:自动编号 + 预设值 */
function handleStudentEdit_(sh, e) {
  const cId = colIdx_(sh, '学生ID');
  const cName = colIdx_(sh, '姓名');
  const cStatus = colIdx_(sh, '状态');
  const cClass = colIdx_(sh, '班级');
  const cFee = colIdx_(sh, '每堂收费RM');
  const cPkg = colIdx_(sh, '配套堂数');
  const editedCol = e.range.getColumn();
  const touchedClass = editedCol <= cClass && cClass < editedCol + e.range.getNumColumns();

  // 同一班的收费通常一样,新学生留空就照抄同班同学的,免得套到不对的预设值
  const feeByClass = {};
  readTable_(SHEETS.STUDENTS).rows.forEach(function (s) {
    const k = String(s['班级'] == null ? '' : s['班级']).trim();
    if (!k || feeByClass[k]) return;
    const fee = Number(s['每堂收费RM']);
    const pkg = Number(s['配套堂数']);
    if (fee && pkg) feeByClass[k] = { fee: fee, pkg: pkg };
  });

  editedRows_(e).forEach(function (row) {
    const hasName = String(sh.getRange(row, cName).getValue()).trim() !== '';
    if (!hasName) return;

    if (String(sh.getRange(row, cId).getValue()).trim() === '') {
      const id = nextId_(SHEETS.STUDENTS, '学生ID', 'S', 3);
      sh.getRange(row, cId).setValue(id);
      logAudit_(SHEETS.STUDENTS, 'A' + row, '新增学生', '', id);
    }
    if (String(sh.getRange(row, cStatus).getValue()).trim() === '') {
      sh.getRange(row, cStatus).setValue(STUDENT_STATUS.ACTIVE);
    }

    const klass = String(sh.getRange(row, cClass).getValue()).trim();
    const ref = feeByClass[klass];
    if (ref) {
      if (!Number(sh.getRange(row, cFee).getValue())) sh.getRange(row, cFee).setValue(ref.fee);
      if (!Number(sh.getRange(row, cPkg).getValue())) sh.getRange(row, cPkg).setValue(ref.pkg);
    }
  });

  if (touchedClass) refreshClassDropdown_(); // 新班级要马上出现在点名的下拉里

  // 改了学生资料(加人、换班、停课、改收费、补 WhatsApp 号码)都会影响总览,
  // 直接重画,不必记得手动按「刷新总览」
  refreshDashboard();
}

/** 付款表:自动编号、带出姓名、盖时间戳 */
function handlePaymentEdit_(sh, e) {
  const cPayId = colIdx_(sh, '付款ID');
  const cDate = colIdx_(sh, '日期');
  const cStuId = colIdx_(sh, '学生ID');
  const cStuName = colIdx_(sh, '学生姓名');
  const cRecorded = colIdx_(sh, '记录时间');

  const nameById = {};
  readTable_(SHEETS.STUDENTS).rows.forEach(function (s) { nameById[s['学生ID']] = s['姓名']; });

  editedRows_(e).forEach(function (row) {
    const stuId = String(sh.getRange(row, cStuId).getValue()).trim();
    if (!stuId) return;

    if (String(sh.getRange(row, cPayId).getValue()).trim() === '') {
      sh.getRange(row, cPayId).setValue(nextId_(SHEETS.PAYMENTS, '付款ID', 'P', 4));
    }
    if (!sh.getRange(row, cDate).getValue()) {
      sh.getRange(row, cDate).setValue(new Date());
    }
    if (nameById[stuId]) {
      sh.getRange(row, cStuName).setValue(nameById[stuId]);
    } else {
      sh.getRange(row, cStuName).setValue('⚠️ 找不到此学生ID');
    }
    sh.getRange(row, cRecorded).setValue(new Date());
  });

  refreshDashboard();
}

/** 课程记录:改了出席状态 → 重算是否扣堂,并留下修改痕迹 */
function handleSessionEdit_(sh, e) {
  const cStatus = colIdx_(sh, '出席状态');
  const cCharge = colIdx_(sh, '是否扣堂');
  const cModified = colIdx_(sh, '最后修改');
  const editedCol = e.range.getColumn();
  const spans = function (col) {
    return editedCol <= col && col < editedCol + e.range.getNumColumns();
  };

  if (!spans(cStatus) && !spans(cCharge)) return;

  editedRows_(e).forEach(function (row) {
    if (spans(cStatus)) {
      const status = String(sh.getRange(row, cStatus).getValue()).trim();
      if (status in CHARGEABLE_BY_STATUS) {
        const oldCharge = String(sh.getRange(row, cCharge).getValue()).trim();
        const newCharge = CHARGEABLE_BY_STATUS[status] ? YES : NO;
        if (oldCharge !== newCharge) {
          sh.getRange(row, cCharge).setValue(newCharge);
          logAudit_(SHEETS.SESSIONS, 'row ' + row, '改出席状态 → ' + status, oldCharge, newCharge);
        }
      }
    } else if (spans(cCharge)) {
      logAudit_(SHEETS.SESSIONS, 'row ' + row, '手动改是否扣堂',
                e.oldValue === undefined ? '' : e.oldValue,
                sh.getRange(row, cCharge).getValue());
    }
    sh.getRange(row, cModified).setValue(new Date());
  });

  refreshDashboard();
}

/** 总览:勾「已发送」→ 写催费日志 */
function handleDashboardEdit_(sh, e) {
  const cSent = colIdx_(sh, '已发送');
  if (e.range.getColumn() !== cSent || e.range.getNumColumns() !== 1) return;

  const settings = getSettings_();
  const balances = computeBalances_();
  const students = {};
  readTable_(SHEETS.STUDENTS).rows.forEach(function (s) { students[s['学生ID']] = s; });

  const cStuId = colIdx_(sh, '学生ID');
  const cLast = colIdx_(sh, '上次催费');
  const log = sheet_(SHEETS.REMINDER_LOG);

  editedRows_(e).forEach(function (row) {
    if (sh.getRange(row, cSent).getValue() !== true) return;
    const id = String(sh.getRange(row, cStuId).getValue()).trim();
    const student = students[id];
    if (!student) return;

    const b = balances[id] || { balance: 0, charged: 0, paid: 0 };
    const now = new Date();
    log.appendRow([now, id, student['姓名'], b.balance, renderReminder_(student, b, settings)]);
    sh.getRange(row, cLast).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, cSent).setValue(false); // 复原,下一期还能再勾
  });

  toast_('已记进催费日志。');
}

function showHelp() {
  const html =
    '<style>body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;' +
      'padding:16px;line-height:1.7;color:#263238;font-size:13px;}' +
      'h3{margin:18px 0 6px;font-size:14px;}h3:first-child{margin-top:0;}' +
      'code{background:#eceff1;padding:1px 5px;border-radius:3px;}' +
      'li{margin-bottom:4px;}</style>' +
    '<h3>每堂课的流程</h3><ol>' +
      '<li>到「今日点名」表,<b>B1 选日期、D1 选班级</b></li>' +
      '<li>选单 <code>📚 补习管理 → ✅ 载入今日点名</code> → 只会载入那一班的学生</li>' +
      '<li>预设全部「出席」,只改有状况的那几个</li>' +
      '<li><code>📥 提交点名</code> → 系统会直接告诉你谁的堂数用完了</li></ol>' +
      '<p>一天有两个班?点完一班,把 D1 换成另一班,再载入一次就好。</p>' +
    '<h3>出席状态的意思</h3><ul>' +
      '<li><b>出席</b> —— 扣一堂</li>' +
      '<li><b>缺席(照算)</b> —— 学生自己没来,照扣</li>' +
      '<li><b>请假(免扣)</b> —— 生病、比赛,你批准的,不扣</li>' +
      '<li><b>我取消(免扣)</b> —— 你自己取消的,不扣</li>' +
      '<li><b>公假/停课(免扣)</b> —— 公共假期、考试周,不扣</li></ul>' +
    '<h3>余额怎么算</h3>' +
      '<p><code>余额 = 已付堂数 − 已扣堂数</code><br>' +
      '学生付了 4 堂的钱 → 到「付款记录」填 购买堂数 = 4。<br>' +
      '上了 4 堂 → 余额变 0 → 总览标红 🔴,该催费。<br>' +
      '预付 8 堂也一样能用,不必改任何设定。</p>' +
    '<h3>催费</h3>' +
      '<p><code>📱 查看催费草稿</code> 会列出所有该催的学生和写好的讯息,' +
      '按按钮打开 WhatsApp。<b>系统不会自动送出</b>,送出的永远是你。<br>' +
      '送完回「总览」勾「已发送」,系统会记进催费日志(勾勾会自动跳回,下一期能再勾)。</p>' +
    '<h3>改错了怎么办</h3>' +
      '<p>直接到「课程记录」改出席状态,余额会自动重算,而且旧值会存进「修改日志」。' +
      '家长有异议时可以翻出来看。</p>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(560), '使用说明'
  );
}


/* ── 06_Backup.gs ──────────────────────────────────────────────── */

/**
 * 备份。
 *
 * Google Sheet 本身已经有版本记录,但版本记录救不了两种情况:
 *   1. 整个档案被删掉 / 被误清空后过了 30 天
 *   2. 你想看「上个学期结束时」的完整快照
 *
 * 所以这里做的是:定期把整份试算表复制一份到 Google Drive 的备份资料夹,
 * 并且只保留最近 N 份,不会塞爆你的云端硬碟。
 */

const BACKUP_FOLDER_NAME = '补习管理备份';

/** 立刻备份一份 */
function backupNow() {
  const folder = getOrCreateBackupFolder_();
  const ss = ss_();
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmm');
  const name = ss.getName() + ' — 备份 ' + stamp;

  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  const keep = settingNumber_(getSettings_(), '备份保留份数', 8);
  const removed = pruneBackups_(folder, keep);

  logAudit_('(备份)', '', '建立备份', '', name + (removed ? ',清掉 ' + removed + ' 份旧的' : ''));
  return { name: name, folderUrl: folder.getUrl(), removed: removed };
}

/** 选单用:备份 + 弹窗告诉你放在哪 */
function backupNowWithAlert() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = backupNow();
    ui.alert('✅ 备份完成',
      '已复制一份:\n' + r.name +
      '\n\n位置:Google 云端硬碟 → ' + BACKUP_FOLDER_NAME +
      (r.removed ? '\n\n(顺便清掉 ' + r.removed + ' 份过旧的备份)' : '') +
      '\n\n资料夹:\n' + r.folderUrl,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('备份失败', e.message, ui.ButtonSet.OK);
  }
}

/** 装上每周自动备份(每星期日晚上) */
function installWeeklyBackup() {
  const ui = SpreadsheetApp.getUi();
  removeBackupTriggers_();
  ScriptApp.newTrigger('backupNow')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(22)
    .create();
  ui.alert('✅ 已开启自动备份',
    '之后每星期日晚上会自动备份一份到 Google 云端硬碟的「' + BACKUP_FOLDER_NAME + '」资料夹。\n\n' +
    '预设保留最近 8 份,可以在「设置」表改「备份保留份数」。\n\n' +
    '就算你没打开这个试算表,备份还是会跑。',
    ui.ButtonSet.OK);
}

/** 关掉每周自动备份 */
function removeWeeklyBackup() {
  const n = removeBackupTriggers_();
  SpreadsheetApp.getUi().alert(n ? '已关闭自动备份。' : '本来就没有开启自动备份。');
}

function removeBackupTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function getOrCreateBackupFolder_() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/** 只留最近 keep 份,其余丢进垃圾桶(不是永久删除,还能救回来) */
function pruneBackups_(folder, keep) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ file: f, created: f.getDateCreated() });
  }
  if (files.length <= keep) return 0;

  files.sort(function (a, b) { return b.created - a.created; }); // 新的在前
  let removed = 0;
  for (let i = keep; i < files.length; i++) {
    files[i].file.setTrashed(true);
    removed++;
  }
  return removed;
}


/* ── 07_Payments.gs ────────────────────────────────────────────── */

/**
 * 记录收款的视窗。
 *
 * 选班级 → 选学生 → 选方式 → 送出。堂数和金额会自动带,改得动。
 * 送出后视窗不关,余额立刻更新 —— 一堂课收好几个人的钱时可以连续记。
 */

function showPaymentDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(paymentDialogHtml_()).setWidth(480).setHeight(640),
    '记录收款'
  );
}

/** 视窗开启时抓资料:班级、学生(含目前余额与收费)、收款方式 */
function getPaymentFormData() {
  const settings = getSettings_();
  const balances = computeBalances_();
  const defaultFee = settingNumber_(settings, '默认每堂收费RM', 50);
  const defaultPkg = settingNumber_(settings, '每期配套堂数', 4);

  const students = readTable_(SHEETS.STUDENTS).rows
    .filter(function (s) {
      return s['学生ID'] !== '' && s['状态'] !== STUDENT_STATUS.ENDED;
    })
    .map(function (s) {
      const b = balances[s['学生ID']] || { balance: 0 };
      return {
        id: s['学生ID'],
        name: s['姓名'],
        klass: String(s['班级'] == null ? '' : s['班级']).trim(),
        balance: b.balance,
        fee: Number(s['每堂收费RM']) || defaultFee,
        pkg: Number(s['配套堂数']) || defaultPkg
      };
    })
    .sort(function (a, b) {
      if (a.klass !== b.klass) return a.klass < b.klass ? -1 : 1;
      return a.balance - b.balance;          // 欠最多的排前面
    });

  return {
    classes: getClasses_(),
    students: students,
    methods: PAYMENT_METHODS,
    today: Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd')
  };
}

/** 真正写入「付款记录」。回传更新后的余额给视窗显示。 */
function recordPayment(p) {
  const id = String(p && p.studentId ? p.studentId : '').trim();
  if (!id) throw new Error('请先选学生。');

  const lessons = Number(p.lessons);
  if (!lessons || lessons <= 0) throw new Error('堂数要大于 0。');

  const amount = Number(p.amount);
  if (isNaN(amount) || amount < 0) throw new Error('金额不正确。');

  const student = readTable_(SHEETS.STUDENTS).rows.filter(function (s) {
    return String(s['学生ID']).trim() === id;
  })[0];
  if (!student) throw new Error('找不到学生 ' + id + ',请先在「学生」表建立。');

  let date = new Date();
  if (p.date) {
    const parsed = new Date(p.date + 'T00:00:00');
    if (!isNaN(parsed.getTime())) date = parsed;
  }

  const sh = sheet_(SHEETS.PAYMENTS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.PAYMENTS.length).setValues([[
    nextId_(SHEETS.PAYMENTS, '付款ID', 'P', 4),
    date, id, student['姓名'], amount, lessons,
    p.method || '', p.note || '', new Date()
  ]]);

  logAudit_(SHEETS.PAYMENTS, id, '记录收款', '',
            student['姓名'] + ' ' + lessons + ' 堂 RM' + amount);

  refreshDashboard();

  const after = computeBalances_()[id];
  return {
    name: student['姓名'],
    lessons: lessons,
    amount: amount,
    balance: after ? after.balance : lessons
  };
}

function paymentDialogHtml_() {
  return [
    '<style>',
    'body{font-family:-apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;',
    '  margin:0;padding:16px;background:#fafafa;color:#263238;font-size:13px;}',
    'label{display:block;font-weight:600;margin:12px 0 4px;font-size:12px;color:#546e7a;}',
    'select,input{width:100%;box-sizing:border-box;padding:8px 10px;font-size:14px;',
    '  border:1px solid #cfd8dc;border-radius:6px;background:#fff;font-family:inherit;}',
    'select:focus,input:focus{outline:2px solid #1976d2;outline-offset:-1px;}',
    '.row{display:flex;gap:10px;}.row>div{flex:1;}',
    '#info{margin-top:10px;padding:10px 12px;border-radius:8px;background:#eceff1;',
    '  font-size:13px;line-height:1.6;}',
    '#info.owe{background:#ffcdd2;color:#b71c1c;}',
    '#info.ok{background:#c8e6c9;color:#1b5e20;}',
    'button{width:100%;margin-top:16px;padding:12px;font-size:15px;font-weight:700;',
    '  color:#fff;background:#2e7d32;border:0;border-radius:8px;cursor:pointer;}',
    'button:disabled{background:#b0bec5;cursor:default;}',
    '#msg{margin-top:12px;font-size:13px;line-height:1.6;}',
    '#msg .err{color:#c62828;font-weight:600;}',
    '#done{margin-top:14px;border-top:1px solid #cfd8dc;padding-top:10px;}',
    '#done h4{margin:0 0 6px;font-size:12px;color:#78909c;font-weight:600;}',
    '#done div{padding:5px 0;border-bottom:1px solid #eceff1;font-size:12px;}',
    '</style>',

    '<label>班级</label><select id="klass"><option value="">— 全部 —</option></select>',
    '<label>学生</label><select id="student"><option value="">— 请选择 —</option></select>',
    '<div id="info">选了学生就会显示目前余额</div>',

    '<div class="row">',
    '  <div><label>堂数</label><input id="lessons" type="number" min="1" step="1"></div>',
    '  <div><label>金额 RM</label><input id="amount" type="number" min="0" step="0.01"></div>',
    '</div>',

    '<div class="row">',
    '  <div><label>收款方式</label><select id="method"></select></div>',
    '  <div><label>日期</label><input id="date" type="date"></div>',
    '</div>',

    '<label>备注(可留空)</label><input id="note" placeholder="例如:妈妈转账">',
    '<button id="go" disabled>记录收款</button>',
    '<div id="msg"></div>',
    '<div id="done" style="display:none"><h4>这次已记录</h4><div id="list"></div></div>',

    '<script>',
    'var DATA = null;',
    'var $ = function (id) { return document.getElementById(id); };',

    'function load() {',
    '  google.script.run.withSuccessHandler(onData)',
    '    .withFailureHandler(function (e) { fail(e.message); }).getPaymentFormData();',
    '}',

    'function onData(d) {',
    '  DATA = d;',
    '  d.classes.forEach(function (k) { add($("klass"), k, k); });',
    '  d.methods.forEach(function (m) { add($("method"), m, m); });',
    '  $("date").value = d.today;',
    '  fillStudents();',
    '}',

    'function add(sel, val, text) {',
    '  var o = document.createElement("option");',
    '  o.value = val; o.textContent = text; sel.appendChild(o);',
    '}',

    'function fillStudents() {',
    '  var k = $("klass").value, sel = $("student");',
    '  sel.innerHTML = "";',
    '  add(sel, "", "— 请选择 —");',
    '  DATA.students.forEach(function (s) {',
    '    if (k && s.klass !== k) return;',
    '    var tag = s.balance < 0 ? "  ⚠️ 欠 " + (-s.balance) + " 堂"',
    '            : s.balance === 0 ? "  (余额 0)" : "  (剩 " + s.balance + " 堂)";',
    '    add(sel, s.id, s.name + (k ? "" : " · " + s.klass) + tag);',
    '  });',
    '  onStudent();',
    '}',

    'function find(id) {',
    '  var r = null;',
    '  DATA.students.forEach(function (s) { if (s.id === id) r = s; });',
    '  return r;',
    '}',

    'function onStudent() {',
    '  var s = find($("student").value), box = $("info");',
    '  $("go").disabled = !s;',
    '  box.className = "";',
    '  if (!s) { box.textContent = "选了学生就会显示目前余额"; return; }',
    '  $("lessons").value = s.pkg;',
    '  $("amount").value = (s.pkg * s.fee).toFixed(2);',
    '  if (s.balance < 0) {',
    '    box.className = "owe";',
    '    box.innerHTML = s.name + " · " + s.klass + "<br>目前欠 <b>" + (-s.balance) +',
    '      " 堂</b>,折合 <b>RM" + (-s.balance * s.fee).toFixed(2) + "</b>";',
    '  } else {',
    '    box.className = "ok";',
    '    box.innerHTML = s.name + " · " + s.klass + "<br>余额 <b>" + s.balance +',
    '      " 堂</b>(每堂 RM" + s.fee.toFixed(2) + ")";',
    '  }',
    '}',

    'function onLessons() {',
    '  var s = find($("student").value);',
    '  if (s) $("amount").value = (Number($("lessons").value || 0) * s.fee).toFixed(2);',
    '}',

    'function submit() {',
    '  $("go").disabled = true;',
    '  $("msg").textContent = "记录中…";',
    '  google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {',
    '    fail(e.message); $("go").disabled = false;',
    '  }).recordPayment({',
    '    studentId: $("student").value, lessons: $("lessons").value,',
    '    amount: $("amount").value, method: $("method").value,',
    '    date: $("date").value, note: $("note").value',
    '  });',
    '}',

    'function ok(r) {',
    '  $("msg").innerHTML = "✅ " + r.name + " 收款 RM" + Number(r.amount).toFixed(2) +',
    '    "(" + r.lessons + " 堂),最新余额 <b>" + r.balance + " 堂</b>";',
    '  $("done").style.display = "block";',
    '  var d = document.createElement("div");',
    '  d.textContent = r.name + "  " + r.lessons + " 堂  RM" + Number(r.amount).toFixed(2) +',
    '    "  → 余额 " + r.balance;',
    '  $("list").insertBefore(d, $("list").firstChild);',
    '  $("note").value = "";',
    '  google.script.run.withSuccessHandler(function (d2) {',   // 重抓余额,继续记下一个
    '    DATA = d2; var keep = $("klass").value;',
    '    $("klass").value = keep; fillStudents();',
    '  }).getPaymentFormData();',
    '}',

    'function fail(m) { $("msg").innerHTML = \'<span class="err">❌ \' + m + "</span>"; }',

    'document.addEventListener("DOMContentLoaded", function () {',
    '  $("klass").addEventListener("change", fillStudents);',
    '  $("student").addEventListener("change", onStudent);',
    '  $("lessons").addEventListener("input", onLessons);',
    '  $("go").addEventListener("click", submit);',
    '  load();',
    '});',
    '</script>'
  ].join('\n');
}
