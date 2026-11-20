#!/usr/bin/env python3
"""从 keyword_trend_*.xlsx 提取 JP-monthly 工作表，另存为单 sheet 文件供 analyze.py 使用。

自动清理：
1. 跳过文件名已含 _monthly 的文件（否则会把上次产物再加工成 _monthly_monthly）。
2. 每次运行前把残留的 *_monthly_monthly*.xlsx 移到 raw-products/_archive_prev/，
   避免 analyze.load_all 把它们当趋势重复读入。
"""
import sys, glob, os, shutil
import openpyxl

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
RAW = os.path.join(BASE, 'raw')
ARCHIVE = os.path.join(BASE, 'raw-products', '_archive_prev')

# 1. 清理历史遗留的多层 monthly 冗余文件
redundant = glob.glob(os.path.join(RAW, '*_monthly_monthly*.xlsx'))
if redundant:
    os.makedirs(ARCHIVE, exist_ok=True)
    for f in redundant:
        dst = os.path.join(ARCHIVE, os.path.basename(f))
        # 同名覆盖：先删旧的再移
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(f, dst)
    print(f'清理多层monthly冗余文件 {len(redundant)} 个 -> raw-products/_archive_prev/')

# 2. 只处理原始趋势文件，跳过已加工的 *_monthly.xlsx
files = sorted(glob.glob(os.path.join(RAW, 'keyword_trend_*.xlsx')))
files = [f for f in files if '_monthly' not in os.path.basename(f)]

for f in files:
    wb = openpyxl.load_workbook(f, read_only=True)
    monthly = [ws.title for ws in wb.worksheets if 'monthly' in ws.title.lower()]
    if not monthly:
        print(f'{os.path.basename(f)}: 无 monthly sheet, sheets={[ws.title for ws in wb.worksheets]}')
        continue
    src = wb[monthly[0]]
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = monthly[0]
    n = 0
    for row in src.iter_rows(values_only=True):
        ws.append(list(row))
        n += 1
    dst = f.replace('.xlsx', '_monthly.xlsx')
    out.save(dst)
    print(f'{os.path.basename(f)} -> {os.path.basename(dst)} rows={n} sheet={monthly[0]}')
