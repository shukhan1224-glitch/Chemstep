#!/usr/bin/env python3
"""
把旧的 Excel 点名表(矩阵式:横轴日期、纵轴学生、1=出席 0=缺席)
转成新系统可以直接贴上的 CSV。

用法:
    python3 tools/import_excel.py 旧点名表.xlsx -o out/
    python3 tools/import_excel.py 旧点名表.xlsx -o out/ --history none
    python3 tools/import_excel.py 旧点名表.xlsx -o out/ --no-swap-fix

产生:
    out/学生.csv        贴进「学生」表 A2
    out/课程记录.csv     贴进「课程记录」表 A2   (--history all 时才有)
    out/付款记录.csv     贴进「付款记录」表 A2   (期初结转,让今天的余额归零)
    out/汇入报告.txt     做了什么、跳过了什么

── 日期修正 ────────────────────────────────────────────────────────
旧档的日期栏常常一半是文字、一半是数字。原因是 Excel 的地区设定是
「月/日/年」,而使用者输入的是「日/月/年」:
    输入 19/1/2026 → 19 不可能是月份 → Excel 放弃解析 → 存成文字(正确)
    输入 12/1/2026 → 当成 12月1日   → 存成数字(日月被对调了)
本程式预设会把数字那一半的日月对调回来,并用「同一班的星期是否一致」
来验证。要关掉这个修正请加 --no-swap-fix。
"""

import argparse
import csv
import datetime
import os
import sys
from collections import Counter, OrderedDict

try:
    import openpyxl
except ImportError:
    sys.exit('需要 openpyxl:pip install openpyxl')

EPOCH = datetime.date(1899, 12, 30)
WEEK = '一二三四五六日'

STATUS_PRESENT = '出席'
STATUS_ABSENT = '缺席(照算)'
YES, NO = '是', '否'
CHARGEABLE = {STATUS_PRESENT: YES, STATUS_ABSENT: YES}

STUDENT_HEADERS = ['学生ID', '姓名', '班级', '家长称呼', 'WhatsApp号码',
                   '每堂收费RM', '配套堂数', '状态', '备注']
SESSION_HEADERS = ['记录ID', '日期', '班级', '学生ID', '学生姓名',
                   '出席状态', '是否扣堂', '备注', '记录时间', '最后修改']
PAYMENT_HEADERS = ['付款ID', '日期', '学生ID', '学生姓名', '金额RM',
                   '购买堂数', '付款方式', '备注', '记录时间']


# ── 日期解析 ────────────────────────────────────────────────────────

def from_serial(n):
    return EPOCH + datetime.timedelta(days=int(n))


def swap_day_month(d):
    """2026-12-01 → 2026-01-12。日或月超过 12 就换不了,回传 None。"""
    try:
        return datetime.date(d.year, d.day, d.month)
    except ValueError:
        return None


def parse_header_date(value, swap_fix):
    """回传 (date, 来源说明) 或 (None, 原因)"""
    if value is None:
        return None, '空白'

    if isinstance(value, datetime.datetime):
        return value.date(), '真日期'
    if isinstance(value, datetime.date):
        return value, '真日期'

    if isinstance(value, (int, float)):
        raw = from_serial(value)
        if swap_fix:
            fixed = swap_day_month(raw)
            if fixed and fixed != raw:
                return fixed, '数字%d,日月对调 %s → %s' % (int(value), raw, fixed)
            return raw, '数字%d(日月相同或无法对调)' % int(value)
        return raw, '数字%d,未修正' % int(value)

    text = str(value).strip()
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y'):
        try:
            return datetime.datetime.strptime(text, fmt).date(), '文字「%s」' % text
        except ValueError:
            continue
    return None, '看不懂的日期「%s」' % text


# ── 读取旧档 ────────────────────────────────────────────────────────

def find_blocks(ws):
    """
    旧档写满一整列后会另起一个区块继续。
    区块的特徵:A 栏是空的,但右边有日期。
    """
    return [row[0].row for row in ws.iter_rows()
            if row[0].value is None and any(c.value is not None for c in row[1:])]


def read_sheet(ws, swap_fix, report):
    """回传 {学生姓名: [(date, 出席状态), ...]},顺序保留原表的学生顺序"""
    records = OrderedDict()
    weekdays = Counter()
    bad_dates = []

    for block_row in find_blocks(ws):
        header = {}
        for cell in ws[block_row][1:]:
            if cell.value is None:
                continue
            d, why = parse_header_date(cell.value, swap_fix)
            if d is None:
                bad_dates.append('%s!%s %s' % (ws.title, cell.coordinate, why))
            else:
                header[cell.column] = d
                weekdays[WEEK[d.weekday()]] += 1

        r = block_row + 1
        while r <= ws.max_row and ws.cell(r, 1).value is not None:
            name = str(ws.cell(r, 1).value).strip()
            records.setdefault(name, [])
            for col, d in header.items():
                v = ws.cell(r, col).value
                if v is None or str(v).strip() == '':
                    continue          # 空白 = 当时还没入学 / 已退出,不是缺席
                try:
                    n = int(float(v))
                except (TypeError, ValueError):
                    report.append('  ⚠️ %s!%s 看不懂的出席值「%s」,已跳过'
                                  % (ws.title, ws.cell(r, col).coordinate, v))
                    continue
                records[name].append((d, STATUS_PRESENT if n else STATUS_ABSENT))
            r += 1

    return records, weekdays, bad_dates


# ── 主流程 ──────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx', help='旧的 Excel 点名表')
    ap.add_argument('-o', '--out', default='out', help='输出资料夹(预设 out/)')
    ap.add_argument('--history', choices=['all', 'none'], default='all',
                    help='all=连历史出席记录一起汇入(预设);none=只汇入学生名单')
    ap.add_argument('--no-swap-fix', action='store_true',
                    help='不要修正日月对调(只在你确定旧档日期是对的时候用)')
    ap.add_argument('--fee', type=float, default=50, help='每堂收费,预设 50')
    ap.add_argument('--package', type=int, default=4, help='配套堂数,预设 4')
    args = ap.parse_args()

    swap_fix = not args.no_swap_fix
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    os.makedirs(args.out, exist_ok=True)

    report = ['旧 Excel 汇入报告',
              '来源:%s' % os.path.basename(args.xlsx),
              '产生时间:%s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
              '日月对调修正:%s' % ('开启' if swap_fix else '关闭'),
              '历史记录:%s' % ('一起汇入' if args.history == 'all' else '不汇入'),
              '']

    students, sessions = [], []
    seen_names = {}
    sid = pid = lid = 0
    today = datetime.date.today()

    for ws in wb.worksheets:
        klass = ws.title.strip()
        records, weekdays, bad_dates = read_sheet(ws, swap_fix, report)
        if not records:
            report.append('【%s】没有资料,跳过' % klass)
            continue

        main_day, main_count = weekdays.most_common(1)[0]
        total_dates = sum(weekdays.values())
        report.append('【%s】%d 人 · %d 堂课 · 主要在星期%s(%d/%d)'
                      % (klass, len(records), total_dates, main_day, main_count, total_dates))
        if main_count < total_dates:
            others = ', '.join('星期%s×%d' % (d, n)
                               for d, n in weekdays.most_common()[1:])
            report.append('    其余不在固定日:%s(可能是补课)' % others)
        for b in bad_dates:
            report.append('    ⚠️ %s' % b)

        for name, entries in records.items():
            sid += 1
            student_id = 'S%03d' % sid
            if name in seen_names:
                report.append('    ⚠️ 「%s」在「%s」和「%s」都出现,已当成两个不同的学生'
                              % (name, seen_names[name], klass))
            seen_names[name] = klass

            charged = sum(1 for _, st in entries if CHARGEABLE.get(st) == YES)
            students.append([student_id, name, klass, '', '',
                             args.fee, args.package, '在读',
                             '由旧 Excel 汇入,历史 %d 堂' % charged])

            if args.history == 'all':
                for d, st in sorted(entries):
                    lid += 1
                    sessions.append(['L%04d' % lid, d.isoformat(), klass,
                                     student_id, name, st, CHARGEABLE.get(st, NO),
                                     '旧 Excel 汇入', '', ''])

    # 期初结转:补一笔付款,让「今天」的余额归零。
    # 不这样做的话,汇入历史后每个人都会立刻变成大额欠费而被误催。
    payments = []
    if args.history == 'all':
        charged_by_id = Counter(s[3] for s in sessions if s[6] == YES)
        for st in students:
            student_id, name = st[0], st[1]
            n = charged_by_id.get(student_id, 0)
            if not n:
                continue
            pid += 1
            payments.append(['P%04d' % pid, today.isoformat(), student_id, name,
                             '', n, '期初结转',
                             '汇入旧记录时的结转,视为历史已结清', ''])

    def write_csv(fname, headers, rows):
        path = os.path.join(args.out, fname)
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)
        return path

    write_csv('学生.csv', STUDENT_HEADERS, students)
    if sessions:
        write_csv('课程记录.csv', SESSION_HEADERS, sessions)
    if payments:
        write_csv('付款记录.csv', PAYMENT_HEADERS, payments)

    report += ['',
               '─' * 50,
               '学生      %d 位' % len(students),
               '课程记录  %d 笔' % len(sessions),
               '付款记录  %d 笔(期初结转)' % len(payments),
               '',
               '⚠️ 「期初结转」的用意:历史出席会扣掉堂数,若不补这笔,',
               '   每个人一汇入就变成欠一大笔而被误判该催费。补上之后,',
               '   今天的余额一律是 0 —— 也就是「旧账当作已经结清」,',
               '   从下一堂课开始重新计算。',
               '',
               '⚠️ 旧档只有 1/0,分不出「缺席」和「请假」,一律当成',
               '   「缺席(照算)」。要改的话,汇入后到「课程记录」改那几格,',
               '   余额会自动重算。',
               '',
               '下一步:把每个 CSV 的内容(不含标题列)贴到对应工作表的 A2,',
               '        然后执行 📚 补习管理 → 🔄 刷新总览。',
               '        学生的 WhatsApp 号码和家长称呼旧档没有,要自己补。']

    path = os.path.join(args.out, '汇入报告.txt')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report) + '\n')

    print('\n'.join(report))
    print('\n✅ 输出到 %s/' % args.out)


if __name__ == '__main__':
    main()
