/**
 * 点名流程。
 *
 * 设计原则:点名要三秒内做完 —— 载入 → 改几个下拉 → 提交。
 * 大部分学生都是「出席」,所以预设就是出席,你只需要动有状况的那几个。
 */

/** 载入某个班级的在读学生到「今日点名」 */
function loadRollcall() {
  const ui = SpreadsheetApp.getUi();
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

  const balances = computeBalances_();
  const out = students.map(function (s) {
    const b = balances[s['学生ID']] || { balance: 0 };
    return [
      s['学生ID'],
      s['姓名'],
      b.balance,
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

/** 把「今日点名」的内容写进「课程记录」 */
function submitRollcall() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SHEETS.ROLLCALL);
  const lastRow = sh.getLastRow();

  if (lastRow < 3) {
    ui.alert('点名表是空的。请先执行「载入今日点名」。');
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
