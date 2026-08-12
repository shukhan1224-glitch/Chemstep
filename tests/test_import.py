#!/usr/bin/env python3
"""验证旧 Excel 汇入时的日期修正逻辑。跑法:python3 tests/test_import.py"""

import datetime
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'tools'))
from import_excel import parse_header_date, swap_day_month, from_serial  # noqa: E402

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

print('\n%s%d 项通过\n' % ('❌ %d 项失败,' % failed if failed else '🎉 全数通过,', passed))
sys.exit(1 if failed else 0)
