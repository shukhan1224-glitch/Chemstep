#!/usr/bin/env python3
"""验证旧 Excel 汇入时的日期修正逻辑。跑法:python3 tests/test_import.py"""

import datetime
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'tools'))
from import_excel import (parse_header_date, swap_day_month, from_serial,  # noqa: E402
                          read_sheet, fill_rgb, detect_paid_fill, CHARGEABLE, YES)

D = datetime.date
passed = failed = 0


def check(label, actual, expected):
    global passed, failed
    ok = actual == expected
    print(('  ✅ ' if ok else '  ❌ ') + label +
          ('  → %s' % (actual,) if ok else '\n       得到 %s\n       预期 %s' % (actual, expected)))
    if ok:
        passed += 1
    else:
        failed += 1


print('\n── 日月对调 ──')
check('2026-12-01 → 2026-01-12', swap_day_month(D(2026, 12, 1)), D(2026, 1, 12))
check('2026-09-03 → 2026-03-09', swap_day_month(D(2026, 9, 3)), D(2026, 3, 9))
check('日 > 12 无法对调', swap_day_month(D(2026, 1, 19)), None)
check('日月相同 → 不变', swap_day_month(D(2026, 2, 2)), D(2026, 2, 2))

print('\n── Excel 序号解码 ──')
check('46357 = 2026-12-01', from_serial(46357), D(2026, 12, 1))
check('46055 = 2026-02-02', from_serial(46055), D(2026, 2, 2))

print('\n── 标题日期解析(开启修正)──')
# 这几个值直接取自使用者的旧档,预期结果由「同班星期一致」推得
cases = [
    (46357, D(2026, 1, 12), '初二科学 周一'),
    (46268, D(2026, 3, 9), '初二科学 周一'),
    (46055, D(2026, 2, 2), '初二科学 周一(日月相同)'),
    (46144, D(2026, 2, 5), '高二补习 周四'),
    (46175, D(2026, 2, 6), '高三 周五'),
    (46025, D(2026, 3, 1), '高一 周日'),
]
for raw, expect, why in cases:
    got, _ = parse_header_date(raw, swap_fix=True)
    check('%s → %s  (%s)' % (raw, expect, why), got, expect)

print('\n── 文字日期一律照 日/月/年 读 ──')
for text, expect in [('19/1/2026', D(2026, 1, 19)),
                     ('26/1/2026', D(2026, 1, 26)),
                     ('16/1/26', D(2026, 1, 16)),
                     ('2026-03-16', D(2026, 3, 16))]:
    got, _ = parse_header_date(text, swap_fix=True)
    check('「%s」' % text, got, expect)

print('\n── 关闭修正时应保留原值 ──')
got, _ = parse_header_date(46357, swap_fix=False)
check('--no-swap-fix 时 46357 维持 2026-12-01', got, D(2026, 12, 1))

print('\n── 坏资料 ──')
check('空白', parse_header_date(None, True)[0], None)
check('乱字串', parse_header_date('第三週', True)[0], None)


# ── 端到端:组一张假的旧点名表,验证 1/0 与底色的解读 ────────────────
print('\n── 端到端:1=算钱、0=不算钱、底色=已缴 ──')
import openpyxl                                                    # noqa: E402
from openpyxl.styles import PatternFill                            # noqa: E402

YELLOW = PatternFill(fill_type='solid', fgColor='FFFFFF00')
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '测试班'

# 4 个上课日:两个文字日期、两个「被 Excel 吃掉」的数字日期
#   B1 文字 15/1/2026(四)  C1 数字 46144→2026-02-05(四)
#   D1 文字 26/2/2026(四)  E1 数字 46145→2026-03-05(四)
ws['B1'], ws['C1'], ws['D1'], ws['E1'] = '15/1/2026', 46144, '26/2/2026', 46145

# 甲:全出席,前 3 堂涂色         → 已上 4、已缴 3
# 乙:第 2 堂是 0(不算钱),整段涂色 → 已上 3、已缴 3(0 那格不计入已缴)
# 丙:中途才加入(前两格空白)     → 已上 2、已缴 0
ws['A2'], ws['A3'], ws['A4'] = '甲', '乙', '丙'
for col, v in zip('BCDE', [1, 1, 1, 1]):
    ws['%s2' % col] = v
for col, v in zip('BCDE', [1, 0, 1, 1]):
    ws['%s3' % col] = v
for col, v in zip('DE', [1, 1]):
    ws['%s4' % col] = v

for col in 'BCD':
    ws['%s2' % col].fill = YELLOW      # 甲:涂 3 格
for col in 'BCDE':
    ws['%s3' % col].fill = YELLOW      # 乙:涂满 4 格,其中一格是 0

detected, counts = detect_paid_fill(wb)
check('自动侦测底色', detected, 'FFFFFF00')
check('底色格数', counts['FFFFFF00'], 7)
check('没有底色的格子回传 None', fill_rgb(ws['E2']), None)

records, weekdays, bad = read_sheet(ws, ws, True, 'FFFFFF00', False, [])
check('日期全部还原到星期四', dict(weekdays), {'四': 4})
check('没有解析不了的日期', bad, [])


def tally(name, recs=None):
    r = (recs or records)[name]
    used = sum(1 for _, st, _ in r['entries'] if CHARGEABLE.get(st) == YES)
    paid = sum(1 for _, _, p in r['entries'] if p) + r['prepaid']
    return used, paid


check('甲 已上/已缴', tally('甲'), (4, 3))
check('乙 已上/已缴(0 不算钱、也不算已缴)', tally('乙'), (3, 3))
check('丙 已上/已缴(空白不产生记录)', tally('丙'), (2, 0))
check('乙 的 0 被记成请假(免扣)',
      sorted({st for _, st, _ in records['乙']['entries']}), ['出席', '请假(免扣)'])

records2, _, _ = read_sheet(ws, ws, True, 'FFFFFF00', True, [])   # --zero-charged
used2 = sum(1 for _, st, _ in records2['乙']['entries'] if CHARGEABLE.get(st) == YES)
check('--zero-charged 时乙的 0 改为照算', used2, 4)


# ── 预缴:涂了色但还没写 1 ────────────────────────────────────────
print('\n── 预缴(涂色但空白)──')
wb2 = openpyxl.Workbook()
w2 = wb2.active
w2.title = '预缴班'

# 区块 1:B..D 有日期,E 没日期(还没上的课)
w2['B1'], w2['C1'], w2['D1'] = '15/1/2026', '22/1/2026', '29/1/2026'
# 区块 2(第 5 列起):B..C 有日期
w2['B5'], w2['C5'] = '5/2/2026', '12/2/2026'

# 丁:区块1 三堂全涂,E2 是预缴记号,后面没有再涂 → 已缴 3+1=4
# 戊:区块1 三堂全涂,E3 预缴记号,但区块2 B7 又写了 1 且涂色
#     → 那笔预缴已经被 B7 用掉,记号不算 → 已缴 4(不是 5)
# 己:区块1 三堂全涂,E4 预缴记号,区块2 B8 是 0 且涂色
#     → 0 不消耗预缴 → 已缴 3+1=4
for row, name in ((2, '丁'), (3, '戊'), (4, '己')):
    w2.cell(row, 1).value = name
    for col in 'BCD':
        c = w2['%s%d' % (col, row)]
        c.value, c.fill = 1, YELLOW
    w2['E%d' % row].fill = YELLOW          # 预缴记号(没有值)

for row, name in ((6, '丁'), (7, '戊'), (8, '己')):
    w2.cell(row, 1).value = name
w2['B7'].value, w2['B7'].fill = 1, YELLOW  # 戊:预缴被真正上的课用掉
w2['C7'].value = 1
w2['B8'].value, w2['B8'].fill = 0, YELLOW  # 己:0 不消耗预缴
w2['C8'].value = 1
w2['B6'].value = 1                          # 丁:后面没再涂色

r2, _, _ = read_sheet(w2, w2, True, 'FFFFFF00', False, [])
check('丁 预缴记号在最后 → 算 1 堂', r2['丁']['prepaid'], 1)
check('戊 预缴已被后面的 1 用掉 → 不算', r2['戊']['prepaid'], 0)
check('己 后面只有 0,预缴仍有效 → 算 1 堂', r2['己']['prepaid'], 1)
check('丁 已上/已缴', tally('丁', r2), (4, 4))
check('戊 已上/已缴', tally('戊', r2), (5, 4))
check('己 已上/已缴', tally('己', r2), (4, 4))

print('\n%s%d 项通过\n' % ('❌ %d 项失败,' % failed if failed else '🎉 全数通过,', passed))
sys.exit(1 if failed else 0)
