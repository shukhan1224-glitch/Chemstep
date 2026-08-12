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
