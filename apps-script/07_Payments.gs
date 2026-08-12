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
