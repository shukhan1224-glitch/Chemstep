// 用假资料验证核心逻辑:余额计算、催费判定、号码正规化、讯息套版
const fs = require('fs');
const dir = '/home/user/Chemstep/apps-script/';

// 载入不依赖 SpreadsheetApp 的档案内容
let src = '';
['00_Config.gs', '03_Dashboard.gs', '04_Reminders.gs'].forEach(f => {
  src += fs.readFileSync(dir + f, 'utf8') + '\n';
});

// 假资料
const DB = {
  '学生': [
    { _row: 2, '学生ID': 'S001', '姓名': '陈小明', '家长称呼': '陈太太', 'WhatsApp号码': '012-345 6789', '班级': '初二科学', '每堂收费RM': 60, '配套堂数': 4, '状态': '在读' },
    { _row: 3, '学生ID': 'S002', '姓名': '林美玲', '家长称呼': '', 'WhatsApp号码': '+60198887777', '班级': '高二补习', '每堂收费RM': '', '配套堂数': '', '状态': '在读' },
    { _row: 4, '学生ID': 'S003', '姓名': '黄志强', '家长称呼': '黄先生', 'WhatsApp号码': '60177778888', '班级': '初二科学', '每堂收费RM': 50, '配套堂数': 8, '状态': '在读' },
    { _row: 5, '学生ID': 'S004', '姓名': '已毕业的', '家长称呼': '', 'WhatsApp号码': '', '班级': '初二科学', '每堂收费RM': 50, '配套堂数': 4, '状态': '结束' }
  ],
  '付款记录': [
    { '学生ID': 'S001', '购买堂数': 4 },
    { '学生ID': 'S002', '购买堂数': 4 },
    { '学生ID': 'S003', '购买堂数': 8 }   // 预付 8 堂 —— counter 模型会在这里出事
  ],
  '课程记录': [
    // S001:上了 4 堂(3 出席 + 1 缺席照算)→ 余额 0 → 该催费
    { '学生ID': 'S001', '日期': new Date('2026-07-01'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S001', '日期': new Date('2026-07-08'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S001', '日期': new Date('2026-07-15'), '出席状态': '缺席(照算)', '是否扣堂': '是' },
    { '学生ID': 'S001', '日期': new Date('2026-07-22'), '出席状态': '出席', '是否扣堂': '是' },
    // S002:上了 4 堂,但其中 1 堂请假免扣 → 只扣 3 → 余额 1 → 快用完
    { '学生ID': 'S002', '日期': new Date('2026-07-02'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S002', '日期': new Date('2026-07-09'), '出席状态': '请假(免扣)', '是否扣堂': '否' },
    { '学生ID': 'S002', '日期': new Date('2026-07-16'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S002', '日期': new Date('2026-07-23'), '出席状态': '出席', '是否扣堂': '是' },
    // S003:预付 8 堂,上了 4 堂 → 余额 4 → 不该被催费
    { '学生ID': 'S003', '日期': new Date('2026-07-03'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S003', '日期': new Date('2026-07-10'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S003', '日期': new Date('2026-07-17'), '出席状态': '我取消(免扣)', '是否扣堂': '否' },
    { '学生ID': 'S003', '日期': new Date('2026-07-24'), '出席状态': '出席', '是否扣堂': '是' },
    { '学生ID': 'S003', '日期': new Date('2026-07-31'), '出席状态': '出席', '是否扣堂': '是' }
  ],
  '催费日志': []
};

const SETTINGS = {
  '老师名字': 'Ms. Tan',
  '每期配套堂数': 4,
  '默认每堂收费RM': 50,
  '催费门槛(余额≤)': 0,
  '预告门槛(余额=)': 1,
  '国际区号': '60',
  '催费讯息模板': '{家长}你好 👋 {学生}已完成 {已扣堂数} 堂,配套用完。下一期 {配套堂数} 堂 RM{金额}。—— {老师}',
  '预告讯息模板': '{家长}你好 👋 {学生}的配套还剩 {余额} 堂。下一期 {配套堂数} 堂 RM{金额}。—— {老师}'
};

// 桩函式
const stubs = `
function readTable_(name){ return { headers: [], rows: (DB[name] || []) }; }
function getSettings_(){ return SETTINGS; }
function sheet_(){ throw new Error('本测试不该碰真表'); }
function ss_(){ return { getSpreadsheetTimeZone: function(){ return 'Asia/Kuala_Lumpur'; } }; }
function toast_(){}
function logAudit_(){}
var SpreadsheetApp = { getUi: function(){ throw new Error('no ui'); } };
var Utilities = { formatDate: function(d){ return d.toISOString().slice(0,10); } };
`;

const sandbox = { DB, SETTINGS, console };
require('vm').createContext(sandbox);
require('vm').runInContext(src + stubs, sandbox);

let pass = 0, fail = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((ok ? '  ✅' : '  ❌') + ' ' + label +
    (ok ? '  → ' + JSON.stringify(actual)
        : '\n       得到: ' + JSON.stringify(actual) + '\n       预期: ' + JSON.stringify(expected)));
  ok ? pass++ : fail++;
}

console.log('\n── 余额计算 ──');
const b = sandbox.computeBalances_();
check('S001 已付', b['S001'].paid, 4);
check('S001 已扣(含1堂缺席照算)', b['S001'].charged, 4);
check('S001 余额', b['S001'].balance, 0);
check('S002 已扣(请假那堂免扣)', b['S002'].charged, 3);
check('S002 余额', b['S002'].balance, 1);
check('S003 预付8堂、上4堂扣4', b['S003'].charged, 4);
check('S003 余额(counter模型会在这里误催)', b['S003'].balance, 4);
check('S001 上次上课', b['S001'].lastSession.toISOString().slice(0,10), '2026-07-22');

console.log('\n── 该催费的人 ──');
const due = sandbox.getStudentsNeedingReminder_();
check('只有 S001 该催费', due.map(d => d.name), ['陈小明']);
check('催费名单带出班级', due.map(d => d.klass), ['初二科学']);

console.log('\n── 班级清单(点名下拉用) ──');
check('去重、排序、排除已结束的学生', sandbox.getClasses_(), ['初二科学', '高二补习']);

console.log('\n── 号码正规化 (wa.me 需要 60xxxxxxxxx) ──');
check('012-345 6789', sandbox.normalizePhone_('012-345 6789', '60'), '60123456789');
check('+60198887777', sandbox.normalizePhone_('+60198887777', '60'), '60198887777');
check('60177778888', sandbox.normalizePhone_('60177778888', '60'), '60177778888');
check('0060123456789', sandbox.normalizePhone_('0060123456789', '60'), '60123456789');
check('空号码', sandbox.normalizePhone_('', '60'), '');

console.log('\n── 讯息套版 ──');
const s1 = DB['学生'][0], s2 = DB['学生'][1];
const m1 = sandbox.renderReminder_(s1, b['S001'], SETTINGS);
const m2 = sandbox.renderReminder_(s2, b['S002'], SETTINGS);
console.log('  S001 (余额0,催费版):\n    ' + m1);
console.log('  S002 (余额1,预告版):\n    ' + m2);
check('S001 用催费模板', m1.indexOf('配套用完') > -1, true);
check('S001 金额 = 60 x 4', m1.indexOf('RM240.00') > -1, true);
check('S001 家长称呼', m1.indexOf('陈太太') > -1, true);
check('S002 用预告模板', m2.indexOf('还剩 1 堂') > -1, true);
check('S002 家长称呼留空 → fallback', m2.indexOf('林美玲家长') > -1, true);
check('S002 金额用预设 50 x 4', m2.indexOf('RM200.00') > -1, true);
check('没有残留未取代的变数', /\{[^}]+\}/.test(m1 + m2), false);

console.log('\n── wa.me 链接 ──');
const f = sandbox.buildWaFormula_(s1, b['S001'], SETTINGS);
console.log('  ' + f.slice(0, 110) + '...');
check('是 HYPERLINK 公式', f.indexOf('=HYPERLINK("https://wa.me/60123456789?text=') === 0, true);
check('公式内没有裸双引号', (f.match(/"/g) || []).length, 4);
check('没号码 → 不产生链接', sandbox.buildWaFormula_(DB['学生'][3], {balance:0,charged:0}, SETTINGS), '⚠️ 缺 WhatsApp 号码');

console.log('\n' + (fail ? '❌ ' + fail + ' 项失败,' : '🎉 全数通过,') + pass + ' 项通过\n');
process.exit(fail ? 1 : 0);
