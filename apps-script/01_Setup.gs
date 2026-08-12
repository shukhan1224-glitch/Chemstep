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

  applyDropdown_(sh, colIdx_(sh, '付款方式'), rows,
    ['银行转账', 'DuitNow / QR', '现金', 'eWallet', '其他']);
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
  sh.getRange(3, cStatus, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(objValues_(ATTENDANCE_STATUS), true)
      .setAllowInvalid(false).build()
  );
  sh.setColumnWidth(cStatus, 150);
  sh.setColumnWidth(cStatus + 1, 240);
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
