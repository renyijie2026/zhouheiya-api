"""
运行前准备：
  1. 将各店CSV文件放入「各店数据」文件夹
  2. 确保「周黑鸭毛利测算_最终.xlsx」存在且产品清单正确

用法：
  python update_import.py              → 汇总最新日期
  python update_import.py 0428 0429    → 汇总指定日期
  python update_import.py all          → 汇总全部日期
"""
import csv, os, re, sys
import openpyxl
import shutil
from collections import defaultdict

SRC = '周黑鸭毛利测算_最终.xlsx'

wb = openpyxl.load_workbook(SRC)
ws = wb['毛利测算']
ws2 = wb['Sheet2']

# 读取17个产品 (rows 2-18): B=品名, C=单品毛利, H=应收单价
products = {}
for row in range(2, 19):
    name = ws.cell(row=row, column=2).value
    c_val = ws.cell(row=row, column=3).value
    h_val = ws.cell(row=row, column=8).value
    if name:
        products[name.strip()] = {'row': row, 'C': c_val, 'H': h_val}

# 清空数据区
for row in range(25, 80):
    for col in [2, 4, 5]:
        ws.cell(row=row, column=col).value = None

# 门店 -> Sheet2行号
store_sheet2_map = {
    '尚峰': 5, '宣化': 6, '民心': 7, '长安': 8, '定州': 9, '未来石': 10,
}

# 高毛利产品（C >= 0.60 的产品）
high_margin_names = [p for p, info in products.items() if info['C'] >= 0.60]

results = {}
csv_folder = '各店数据'

# 收集所有CSV，按门店汇总（跨日期合并）
store_files = defaultdict(list)  # store -> [(date_str, filepath)]
all_dates = set()
for fname in os.listdir(csv_folder):
    if not fname.endswith('.csv'):
        continue
    m = re.search(r'(\d{4})\.csv$', fname)
    if not m:
        continue
    date_str = m.group(1)
    store = fname[:m.start()]
    store_files[store].append((date_str, os.path.join(csv_folder, fname)))
    all_dates.add(date_str)

if not store_files:
    print('No CSV files found!')
    exit(1)

# 日期选择
args = sys.argv[1:]
sorted_all = sorted(all_dates)

if args:
    # 命令行直接指定
    if 'all' in args:
        target_dates = all_dates
    else:
        target_dates = set(args)
else:
    # 交互菜单
    print(f'\n可用日期 ({len(sorted_all)}个):')
    for i, d in enumerate(sorted_all):
        print(f'  [{i+1}] {d}')
    print(f'  [0] 全部日期')
    print(f'  [回车] 最新日期 ({sorted_all[-1]})')
    choice = input('请选择 (多个用逗号分隔，如 1,2): ').strip()

    if choice == '':
        target_dates = {sorted_all[-1]}
    elif choice == '0':
        target_dates = all_dates
    else:
        try:
            idxs = [int(x.strip()) for x in choice.replace('，', ',').split(',') if x.strip()]
            target_dates = {sorted_all[i-1] for i in idxs if 0 < i <= len(sorted_all)}
            if not target_dates:
                print('无效选择，使用最新日期')
                target_dates = {sorted_all[-1]}
        except ValueError:
            print('输入格式错误，使用最新日期')
            target_dates = {sorted_all[-1]}

invalid = target_dates - all_dates
if invalid:
    print(f'未找到日期: {", ".join(sorted(invalid))}')
    print(f'可用日期: {", ".join(sorted(all_dates))}')
    exit(1)

# 过滤：只保留指定日期
for store in list(store_files.keys()):
    store_files[store] = [(d, p) for d, p in store_files[store] if d in target_dates]
    if not store_files[store]:
        del store_files[store]

sorted_dates = sorted(target_dates)
date_range = ', '.join(sorted_dates)
print(f'汇总日期: {date_range}, 共{len(store_files)}个门店, {len(sorted_dates)}天数据')

# 构造日期显示
if len(sorted_dates) == 1:
    d = sorted_dates[0]
    date_display = f'2026/{d[:2]}/{d[2:]}-2026/{d[:2]}/{d[2:]}'
else:
    d1, d2 = sorted_dates[0], sorted_dates[-1]
    date_display = f'2026/{d1[:2]}/{d1[2:]}-2026/{d2[:2]}/{d2[2:]}'

# 清空数据区
for row in range(25, 80):
    for col in [2, 4, 5]:
        ws.cell(row=row, column=col).value = None

for store in ['尚峰', '宣化', '民心', '长安', '定州', '未来石']:
    files = store_files.get(store, [])
    if not files:
        continue

    # 汇总该门店所有日期的数据
    store_data = {}
    for date_str, fpath in files:
        with open(fpath, 'r', encoding='utf-8') as f:
            rows = list(csv.reader(f))
        for row in rows[1:]:
            if not row or len(row) < 10:
                continue
            try:
                name = row[3].strip()
                qty = float(row[5])
                amt = float(row[6])
                if name in store_data:
                    store_data[name]['qty'] += qty
                    store_data[name]['amt'] += amt
                else:
                    store_data[name] = {'qty': qty, 'amt': amt}
            except (ValueError, IndexError):
                continue

    # 写入数据区
    data_row = 25
    for pname, pdata in store_data.items():
        ws.cell(row=data_row, column=1).value = date_display
        ws.cell(row=data_row, column=2).value = pname
        ws.cell(row=data_row, column=3).value = '周黑鸭（ZHOUHEIYA）'
        ws.cell(row=data_row, column=4).value = pdata['qty']
        ws.cell(row=data_row, column=5).value = pdata['amt']
        data_row += 1

    # Calculate metrics
    total_E = 0
    total_I = 0
    total_K = 0
    total_L = 0
    all_E = {}

    for pname, pinfo in products.items():
        sd = store_data.get(pname, {'qty': 0, 'amt': 0})
        D = sd['qty']
        E = sd['amt']
        C = pinfo['C']
        H = pinfo['H']

        I_val = D * H
        K_val = D * H * (1 - C)
        L_val = E - K_val

        total_E += E
        total_I += I_val
        total_K += K_val
        total_L += L_val
        all_E[pname] = E

    F_values = {pname: all_E[pname] / total_E if total_E > 0 else 0 for pname in all_E}

    M17 = total_L / total_E if total_E > 0 else 0
    F19 = sum(F_values.get(p, 0) for p in high_margin_names)
    I19 = total_E / total_I if total_I > 0 else 0

    sr = store_sheet2_map.get(store)
    if sr:
        ws2.cell(row=sr, column=10).value = round(total_E, 2)
        ws2.cell(row=sr, column=11).value = round(M17, 6)
        ws2.cell(row=sr, column=12).value = round(F19, 6)
        ws2.cell(row=sr, column=13).value = round(I19, 6)

    results[store] = {
        'sales': total_E,
        'margin': M17,
        'high_margin_ratio': F19,
        'discount': I19,
    }

# Save (backup before overwrite)
if os.path.exists(SRC):
    shutil.copy(SRC, SRC.replace('.xlsx', '_backup.xlsx'))
wb.save(SRC)

# Report
with open('import_result.txt', 'w', encoding='utf-8') as out:
    out.write(f'=== 更新结果 (日期: {date_range}) ===\n\n')
    out.write(f'产品总数: {len(products)}\n')
    out.write('高毛利产品: ' + ', '.join(high_margin_names) + '\n\n')
    out.write(f'{"门店":<8} {"销售额":>12} {"毛利率":>10} {"高毛利占比":>10} {"折扣":>10}\n')
    out.write('-' * 55 + '\n')
    for store in ['尚峰', '宣化', '未来石', '定州', '民心', '长安']:
        r = results.get(store, {})
        out.write(f'{store:<8} {r.get("sales",0):>12,.2f} {r.get("margin",0):>10.4f} {r.get("high_margin_ratio",0):>10.4f} {r.get("discount",0):>10.4f}\n')

print('Done')
