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
  const sh = sheet_(SHEETS.DASHBOARD);
  const settings = getSettings_();
  const balances = computeBalances_();
  const lastReminder = lastReminderMap_();

  const dueThreshold = settingNumber_(settings, '催费门槛(余额≤)', 0);
  const warnThreshold = settingNumber_(settings, '预告门槛(余额=)', 1);

  const students = readTable_(SHEETS.STUDENTS).rows.filter(function (s) {
    return s['学生ID'] !== '' && s['状态'] !== STUDENT_STATUS.ENDED;
  });

  // 清掉旧内容(含核取方块的验证规则)
  if (sh.getMaxRows() > 1) {
    const old = sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.DASHBOARD.length);
    old.clearContent().clearDataValidations().setBackground(null).setFontColor(null);
  }
  if (!students.length) { toast_('还没有学生资料。'); return; }

  const rows = students.map(function (s) {
    const id = s['学生ID'];
    const b = balances[id] || { paid: 0, charged: 0, balance: 0, lastSession: null };
    const state = b.balance <= dueThreshold ? '🔴 该催费'
                : b.balance <= warnThreshold ? '🟡 快用完'
                : '🟢 正常';
    const link = b.balance <= warnThreshold ? buildWaFormula_(s, b, settings) : '';
    return [
      id, s['姓名'], s['科目'], s['状态'],
      b.paid, b.charged, b.balance,
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

  applyDashboardColors_(sh, rows.length);
  toast_('总览已更新(' + rows.length + ' 位学生)。');
}

function applyDashboardColors_(sh, numRows) {
  const stateRange = sh.getRange(2, colIdx_(sh, '提醒状态'), numRows, 1);
  const balanceRange = sh.getRange(2, colIdx_(sh, '余额'), numRows, 1);

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
      .setRanges([balanceRange]).build()
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
      return { id: s['学生ID'], name: s['姓名'], balance: balances[s['学生ID']].balance };
    });
}
