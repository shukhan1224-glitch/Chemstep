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

  // 删掉 Google 预设的空白 "Sheet1" / "工作表1"
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    const known = Object.keys(SHEETS).some(function (k) { return SHEETS[k] === name; });
    if (!known && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
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
    '2. 到「学生」加学生(姓名填了就会自动配学生ID)\n' +
    '3. 每次上完课 → 补习管理 → 载入今日点名\n\n' +
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
  applyDropdown_(sh, colIdx_(sh, '上课形式'), rows, ['一对一', '小组']);

  sh.getRange(2, colIdx_(sh, 'WhatsApp号码'), rows, 1).setNumberFormat('@'); // 保留开头的 0
  sh.getRange(2, colIdx_(sh, '每堂收费RM'), rows, 1).setNumberFormat('0.00');
  sh.setColumnWidth(colIdx_(sh, '备注'), 240);
  return sh;
}

function setupSessionsSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.SESSIONS, HEADERS.SESSIONS);
  const rows = Math.max(sh.getMaxRows() - 1, 1);

  applyDropdown_(sh, colIdx_(sh, '出席状态'), rows, objValues_(ATTENDANCE_STATUS));
  applyDropdown_(sh, colIdx_(sh, '是否扣堂'), rows, [YES, NO]);
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

  // 第 1 列是上课日期,第 2 列才是标题
  sh.getRange('A1').setValue('上课日期 →').setFontWeight('bold');
  sh.getRange('B1').setNumberFormat('yyyy-mm-dd').setBackground('#fff8e1')
    .setBorder(true, true, true, true, false, false);
  sh.getRange('C1').setValue('← 改这里就能补录过去的课').setFontColor('#90a4ae');

  sh.getRange(2, 1, 1, HEADERS.ROLLCALL.length).setValues([HEADERS.ROLLCALL])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(2);
  if (sh.getMaxColumns() > HEADERS.ROLLCALL.length) {
    sh.deleteColumns(HEADERS.ROLLCALL.length + 1, sh.getMaxColumns() - HEADERS.ROLLCALL.length);
  }

  const rows = Math.max(sh.getMaxRows() - 2, 1);
  sh.getRange(3, 5, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(objValues_(ATTENDANCE_STATUS), true)
      .setAllowInvalid(false).build()
  );
  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 240);
  return sh;
}

function setupDashboardSheet_(ss) {
  const sh = setupPlainSheet_(ss, SHEETS.DASHBOARD, HEADERS.DASHBOARD);
  sh.setColumnWidth(colIdx_(sh, '提醒状态'), 110);
  sh.setColumnWidth(colIdx_(sh, '发送催费'), 130);
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
