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
    .addSeparator()
    .addItem('🔄 刷新总览', 'refreshDashboard')
    .addItem('📱 查看催费草稿', 'showReminderDrafts')
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
  });
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
    '<h3>每天的流程</h3><ol>' +
      '<li>上完课 → 选单 <code>📚 补习管理 → ✅ 载入今日点名</code></li>' +
      '<li>预设全部「出席」,只改有状况的那几个</li>' +
      '<li><code>📥 提交点名</code> → 系统会直接告诉你谁的堂数用完了</li></ol>' +
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
