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
 * 补记某位学生的历史出席。
 *
 * 用在「学生其实早就在上课,但忘了先加进系统」的情况:
 * 依他班上「课程记录」里已经有的上课日,把他缺的那几堂补上去。
 * 不会动到其他学生,也不会重复补已经有的日期。
 */
function backfillStudent() {
  const ui = SpreadsheetApp.getUi();

  const ask = ui.prompt('补记出席 (1/2)',
    '要补记哪一位学生?打学生ID(例如 S041)或姓名都可以。',
    ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;
  const query = ask.getResponseText().trim();
  if (!query) return;

  const matches = readTable_(SHEETS.STUDENTS).rows.filter(function (s) {
    return String(s['学生ID']).trim() === query ||
           String(s['姓名']).trim().toLowerCase() === query.toLowerCase();
  });
  if (!matches.length) { ui.alert('找不到「' + query + '」,请确认学生ID或姓名。'); return; }
  if (matches.length > 1) {
    ui.alert('有超过一位叫「' + query + '」的学生,请改用学生ID:\n• ' +
             matches.map(function (s) { return s['学生ID'] + ' · ' + s['班级']; }).join('\n• '));
    return;
  }

  const student = matches[0];
  const klass = String(student['班级']).trim();
  if (!klass) { ui.alert(student['姓名'] + ' 还没填班级,请先到「学生」表补上。'); return; }

  // 这个班已经点过名的日子,以及这位学生已经有记录的日子
  const sessions = readTable_(SHEETS.SESSIONS).rows;
  const classDates = {};
  const hasAlready = {};
  sessions.forEach(function (r) {
    const key = dateKey_(r['日期']);
    if (!key) return;
    if (String(r['班级']).trim() === klass) classDates[key] = r['日期'];
    if (String(r['学生ID']).trim() === student['学生ID']) hasAlready[key] = true;
  });

  let dates = Object.keys(classDates).filter(function (k) { return !hasAlready[k]; }).sort();
  if (!dates.length) {
    ui.alert(student['姓名'] + ' 在「' + klass + '」已经没有缺漏的上课记录了。');
    return;
  }

  const ask2 = ui.prompt('补记出席 (2/2)',
    student['姓名'] + ' · ' + klass + '\n\n' +
    '他缺记录的上课日有 ' + dates.length + ' 天:\n' + dates.join('、') + '\n\n' +
    '要从哪一天开始补?打日期(例如 2026-07-13)。\n' +
    '整批都要补就留空直接按 OK。',
    ui.ButtonSet.OK_CANCEL);
  if (ask2.getSelectedButton() !== ui.Button.OK) return;

  const from = ask2.getResponseText().trim();
  if (from) {
    const d = new Date(from);
    if (isNaN(d.getTime())) { ui.alert('日期看不懂:' + from + '\n请用 2026-07-13 这种格式。'); return; }
    const fromKey = dateKey_(d);
    dates = dates.filter(function (k) { return k >= fromKey; });
    if (!dates.length) { ui.alert(fromKey + ' 之后没有要补的上课日。'); return; }
  }

  const answer = ui.alert('确定要补记吗?',
    student['姓名'] + ' · ' + klass + '\n\n' +
    '会新增 ' + dates.length + ' 笔「出席」记录(都会扣堂):\n• ' + dates.join('\n• ') + '\n\n' +
    '其中有请假或没上到的,补完再到「课程记录」个别改成免扣。',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const now = new Date();
  let seq = parseInt(nextId_(SHEETS.SESSIONS, '记录ID', 'L', 4).substring(1), 10);
  const rows = dates.map(function (key) {
    const rec = ['L' + ('000' + seq).slice(-4), classDates[key], klass,
                 student['学生ID'], student['姓名'],
                 ATTENDANCE_STATUS.PRESENT, YES, '补记', now, ''];
    seq++;
    return rec;
  });

  const sh = sheet_(SHEETS.SESSIONS);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.SESSIONS.length).setValues(rows);
  logAudit_(SHEETS.SESSIONS, student['学生ID'], '补记历史出席', '',
            rows.length + ' 笔(' + dates[0] + ' → ' + dates[dates.length - 1] + ')');
  refreshDashboard();

  const b = computeBalances_()[student['学生ID']] || { balance: 0 };
  ui.alert('✅ 补记完成',
    student['姓名'] + ' 补了 ' + rows.length + ' 堂。\n\n' +
    '目前余额:' + b.balance + ' 堂' +
    (b.balance <= 0 ? '\n\n🔴 已经该催费了。' : ''),
    ui.ButtonSet.OK);
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
