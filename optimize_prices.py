#!/usr/bin/env python3
"""
单价调整优化器 - 基于MILP算法
模式A (uniform):  等比例下浮，所有单价均匀调整，总价精确达标
模式B (minimal):  最少项数模式，尽量保持原单价不变，只动最少几行凑总价（发票场景）

用法:
  python optimize_prices.py <excel_file> <target_total>              # 默认模式A
  python optimize_prices.py <excel_file> <target_total> --mode minimal   # 模式B
"""

import sys, argparse, os
import pandas as pd
import pulp
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CLR_TOTAL  = "D9E1F2"
CLR_TARGET = "E2EFDA"

# ──────────────────────────────────────────────────────────────
# 列名自动识别
# ──────────────────────────────────────────────────────────────
def detect_columns(df):
    col_map = {}
    for col in df.columns:
        nl = str(col).strip().lower()
        if any(k in nl for k in ['规格','型号','名称','spec','item','材料','描述']):
            col_map.setdefault('spec', col)
        elif any(k in nl for k in ['数量','qty','quantity','用量']):
            col_map.setdefault('qty', col)
        elif any(k in nl for k in ['单价','unit price','price','综合单价']):
            col_map.setdefault('price', col)
        elif any(k in nl for k in ['合价','金额','小计','合计','total','amount']):
            col_map.setdefault('total', col)
    return col_map

# ──────────────────────────────────────────────────────────────
# 模式A：等比例下浮 + MILP微调（均匀下浮）
# ──────────────────────────────────────────────────────────────
def run_uniform(items, quantities, orig_prices, target_yuan, max_delta_cents=5):
    """所有单价等比例缩放，MILP在±max_delta_cents分钱内精确凑总价"""
    n = len(items)
    target_cents = int(round(target_yuan * 100))
    orig_total   = sum(q * p for q, p in zip(quantities, orig_prices))
    ratio        = target_yuan / orig_total
    theory_cents = [round(p * 100 * ratio) for p in orig_prices]

    prob = pulp.LpProblem("Uniform", pulp.LpMinimize)
    x = {i: pulp.LpVariable(f"P_{i}",
                             lowBound=theory_cents[i] - max_delta_cents,
                             upBound=theory_cents[i] + max_delta_cents,
                             cat='Integer') for i in range(n)}
    e_pos = pulp.LpVariable("Ep", lowBound=0, cat='Continuous')
    e_neg = pulp.LpVariable("En", lowBound=0, cat='Continuous')
    d     = {i: pulp.LpVariable(f"D_{i}", lowBound=0, cat='Continuous') for i in range(n)}

    prob += pulp.lpSum(quantities[i]*x[i] for i in range(n)) + e_neg - e_pos == target_cents
    for i in range(n):
        prob += d[i] >= x[i] - theory_cents[i]
        prob += d[i] >= theory_cents[i] - x[i]

    M = 100000
    prob += M*(e_pos+e_neg) + pulp.lpSum(d[i] for i in range(n))
    prob.solve(pulp.PULP_CBC_CMD(msg=False))

    if pulp.LpStatus[prob.status] != 'Optimal':
        return run_uniform(items, quantities, orig_prices, target_yuan, max_delta_cents*2)

    new_prices = [x[i].varValue/100.0 for i in range(n)]
    changed_items = []
    for i in range(n):
        diff = round(x[i].varValue - theory_cents[i])
        if diff != 0:
            changed_items.append({'spec': items[i], 'direction': '上调' if diff>0 else '下调',
                                   'delta_cents': abs(diff), 'qty': quantities[i]})
    return new_prices, changed_items, ratio

# ──────────────────────────────────────────────────────────────
# 模式B：最少项数（发票场景）
# ──────────────────────────────────────────────────────────────
def run_minimal(items, quantities, orig_prices, target_yuan,
                max_change_per_item_yuan=None, max_items_to_change=None):
    """
    尽量保持原单价不变，只修改最少数量的行，使总价精确等于目标值。
    
    差额 gap = target - orig_total，需要在若干行上分摊。
    每行的调整量（分钱）= delta_i，约束：sum(qty_i * delta_i) == gap_cents
    引入0/1变量 z_i：z_i=1 表示该行被修改，目标：minimize sum(z_i)
    
    参数:
      max_change_per_item_yuan: 单项最大调整金额（元），默认不限
      max_items_to_change:      最多允许修改几项，默认不限（由模型自动找最优）
    """
    n = len(items)
    orig_total   = sum(q * p for q, p in zip(quantities, orig_prices))
    gap_cents    = int(round((target_yuan - orig_total) * 100))  # 需要调整的总分钱

    if gap_cents == 0:
        return list(orig_prices), [], 0

    # 每行最大可调整分钱数（防止单价变成负数或过大）
    # 默认：最大调整到原价的±50%，但不超过 max_change_per_item_yuan
    def max_delta_for(i):
        base = int(orig_prices[i] * 100 * 0.5)  # 原价50%
        if max_change_per_item_yuan is not None:
            base = min(base, int(max_change_per_item_yuan * 100))
        return max(base, 1)

    prob = pulp.LpProblem("Minimal_Changes", pulp.LpMinimize)

    # delta[i]: 第i行单价调整量（分钱，可正可负）
    # z[i]:     0/1 是否被修改
    delta = {i: pulp.LpVariable(f"delta_{i}", cat='Integer') for i in range(n)}
    z     = {i: pulp.LpVariable(f"z_{i}", cat='Binary') for i in range(n)}

    # 约束1：总差额精确平衡
    prob += pulp.lpSum(quantities[i] * delta[i] for i in range(n)) == gap_cents

    # 约束2：Big-M 联动：delta_i != 0 => z_i = 1
    for i in range(n):
        M_i = max_delta_for(i)
        prob += delta[i] <=  M_i * z[i]
        prob += delta[i] >= -M_i * z[i]

    # 约束3：单价不能变为负数
    for i in range(n):
        orig_cents = int(round(orig_prices[i] * 100))
        prob += delta[i] >= -orig_cents + 1  # 新单价至少1分钱

    # 约束4（可选）：限制最多修改几项
    if max_items_to_change is not None:
        prob += pulp.lpSum(z[i] for i in range(n)) <= max_items_to_change

    # 目标：最小化被修改的项数
    prob += pulp.lpSum(z[i] for i in range(n))

    prob.solve(pulp.PULP_CBC_CMD(msg=False))

    if pulp.LpStatus[prob.status] != 'Optimal':
        # 放宽限制重试
        if max_items_to_change:
            return run_minimal(items, quantities, orig_prices, target_yuan,
                               max_change_per_item_yuan, None)
        raise ValueError(
            f"模式B无可行解。差额={gap_cents/100:.2f}元，"
            f"请检查目标总价是否合理，或放宽 max_change_per_item_yuan 约束。"
        )

    new_prices = []
    changed_items = []
    for i in range(n):
        d_val = int(round(delta[i].varValue or 0))
        new_p = orig_prices[i] + d_val / 100.0
        new_prices.append(new_p)
        if d_val != 0:
            changed_items.append({
                'spec': items[i],
                'direction': '上调' if d_val > 0 else '下调',
                'delta_cents': abs(d_val),
                'delta_yuan':  abs(d_val) / 100,
                'qty': quantities[i],
                'orig_price': orig_prices[i],
                'new_price':  new_p,
            })

    num_changed = int(round(sum(z[i].varValue or 0 for i in range(n))))
    return new_prices, changed_items, num_changed

# ──────────────────────────────────────────────────────────────
# Excel 读取 + 写回（两种模式共用）
# ──────────────────────────────────────────────────────────────
def process_excel(input_path, target_total, output_path=None,
                  mode='uniform',
                  # 模式A参数
                  max_delta_cents=5,
                  # 模式B参数
                  max_change_per_item_yuan=None,
                  max_items_to_change=None,
                  # 通用参数
                  spec_col=None, qty_col=None, price_col=None, total_col=None,
                  header_row=0, sheet_name=0):

    df = pd.read_excel(input_path, sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.map(str)

    col_map = detect_columns(df)
    if spec_col:  col_map['spec']  = spec_col
    if qty_col:   col_map['qty']   = qty_col
    if price_col: col_map['price'] = price_col
    if total_col: col_map['total'] = total_col

    missing = [k for k in ['qty','price'] if k not in col_map]
    if missing:
        raise ValueError(f"无法自动识别列: {missing}，请用 --qty-col / --price-col 手动指定\n已有列: {list(df.columns)}")

    qty_c, price_c = col_map['qty'], col_map['price']
    mask = pd.to_numeric(df[qty_c], errors='coerce').notna() & \
           pd.to_numeric(df[price_c], errors='coerce').notna()
    data_df = df[mask].copy()
    data_df[qty_c]   = pd.to_numeric(data_df[qty_c])
    data_df[price_c] = pd.to_numeric(data_df[price_c])
    data_df = data_df[data_df[qty_c] > 0]

    if len(data_df) == 0:
        raise ValueError("未找到有效数据行")

    quantities  = data_df[qty_c].tolist()
    orig_prices = data_df[price_c].tolist()
    specs       = data_df[col_map['spec']].tolist() if 'spec' in col_map else [f"Item_{i}" for i in range(len(quantities))]
    orig_total  = sum(q*p for q,p in zip(quantities, orig_prices))

    # ── 调用优化 ─────────────────────────────────────────────
    mode_label = ""
    ratio = None
    num_changed = None

    if mode == 'uniform':
        new_prices, changed_items, ratio = run_uniform(
            specs, quantities, orig_prices, target_total, max_delta_cents)
        mode_label = "模式A：等比例均匀下浮"
        num_changed = len(changed_items)

    elif mode == 'minimal':
        new_prices, changed_items, num_changed = run_minimal(
            specs, quantities, orig_prices, target_total,
            max_change_per_item_yuan, max_items_to_change)
        mode_label = "模式B：最少项数调整（发票场景）"

    else:
        raise ValueError(f"未知模式: {mode}")

    new_total = sum(quantities[i]*new_prices[i] for i in range(len(quantities)))

    # ── 输出路径 ─────────────────────────────────────────────
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        suffix = "_均匀下浮" if mode == 'uniform' else "_最少调整"
        output_path = base + suffix + ext

    # ── 写回 Excel ────────────────────────────────────────────
    wb = load_workbook(input_path)
    ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb[sheet_name]

    excel_header_row = header_row + 1
    col_letters = {c: get_column_letter(j+1) for j, c in enumerate(df.columns)}
    price_col_idx = df.columns.get_loc(price_c) + 1  # 1-based

    # 在单价列后插2列
    insert_at = price_col_idx + 1
    ws.insert_cols(insert_at, 2)
    new_price_letter = get_column_letter(insert_at)
    rate_letter      = get_column_letter(insert_at + 1)

    hfont = Font(bold=True, color="FFFFFF")
    ws[f"{new_price_letter}{excel_header_row}"] = "优化单价"
    ws[f"{new_price_letter}{excel_header_row}"].font = hfont
    ws[f"{new_price_letter}{excel_header_row}"].fill = PatternFill("solid", fgColor="2E75B6")
    ws[f"{new_price_letter}{excel_header_row}"].alignment = Alignment(horizontal='center')

    rate_header = "下浮率" if mode == 'uniform' else "单价变动"
    ws[f"{rate_letter}{excel_header_row}"] = rate_header
    ws[f"{rate_letter}{excel_header_row}"].font = hfont
    ws[f"{rate_letter}{excel_header_row}"].fill = PatternFill("solid", fgColor="375623" if mode=='uniform' else "7B2C2C")
    ws[f"{rate_letter}{excel_header_row}"].alignment = Alignment(horizontal='center')

    data_indices = data_df.index.tolist()
    changed_specs = {c['spec'] for c in changed_items}

    # 模式B：橙色高亮被修改行；模式A：淡黄色
    fill_changed = PatternFill("solid", fgColor="FCE4D6") if mode=='minimal' else PatternFill("solid", fgColor="FFF2CC")

    for k, (orig_idx, new_p, orig_p, spec) in enumerate(zip(data_indices, new_prices, orig_prices, specs)):
        erow = orig_idx + header_row + 2
        changed = abs(new_p - orig_p) > 0.004

        ws[f"{new_price_letter}{erow}"] = round(new_p, 2)
        ws[f"{new_price_letter}{erow}"].number_format = '#,##0.00'

        if mode == 'uniform':
            rate_val = (new_p - orig_p) / orig_p if orig_p != 0 else 0
            ws[f"{rate_letter}{erow}"] = rate_val
            ws[f"{rate_letter}{erow}"].number_format = '0.000%'
        else:
            # 模式B：显示绝对变动金额（元）
            delta_yuan = new_p - orig_p
            ws[f"{rate_letter}{erow}"] = round(delta_yuan, 2)
            ws[f"{rate_letter}{erow}"].number_format = '+#,##0.00;-#,##0.00;"-"'

        if changed:
            ws[f"{new_price_letter}{erow}"].fill = fill_changed
            ws[f"{rate_letter}{erow}"].fill = fill_changed

    ws.column_dimensions[new_price_letter].width = 12
    ws.column_dimensions[rate_letter].width = 11

    # ── 汇总行 ───────────────────────────────────────────────
    last_data_row = max(i + header_row + 2 for i in data_indices)
    sr = last_data_row + 2
    sfont = Font(bold=True)
    sfill = PatternFill("solid", fgColor=CLR_TOTAL)
    tfill = PatternFill("solid", fgColor=CLR_TARGET)

    def sumrow(row, label, value, fill, fmt='#,##0.00'):
        ws.cell(row=row, column=1, value=label).font = sfont
        ws.cell(row=row, column=1).fill = fill
        c = ws.cell(row=row, column=price_col_idx+2, value=value)
        c.number_format = fmt; c.font = sfont; c.fill = fill

    sumrow(sr,   "📊 原始总价（元）",    orig_total,  sfill)
    sumrow(sr+1, "🎯 目标总价（元）",    target_total, tfill)
    sumrow(sr+2, "✅ 优化后总价（元）",  new_total,   tfill)
    sumrow(sr+3, "⚖️ 误差（元）",        abs(new_total-target_total), sfill, '#,##0.0000')
    sumrow(sr+4, "📉 整体下浮率",        (new_total-orig_total)/orig_total, sfill, '0.000%')
    sumrow(sr+5, "✏️ 被调整项数",        num_changed, sfill, '0')

    wb.save(output_path)

    # ── 控制台报告 ───────────────────────────────────────────
    print(f"\n{'='*65}")
    print(f"  {mode_label}")
    print(f"{'='*65}")
    print(f"  原始总价:    {orig_total:>12,.2f} 元")
    print(f"  目标总价:    {target_total:>12,.2f} 元")
    print(f"  优化总价:    {new_total:>12,.2f} 元")
    print(f"  绝对误差:    {abs(new_total-target_total):>12.4f} 元")
    print(f"  整体下浮:    {(new_total-orig_total)/orig_total*100:>12.3f}%")
    print(f"  被调整项数:  {num_changed:>12} 项")
    if ratio is not None:
        print(f"  等比下浮系数: {ratio:>11.6f}")
    print(f"{'='*65}")

    if changed_items:
        print(f"\n{'橙色' if mode=='minimal' else '黄色'}高亮行明细（共 {len(changed_items)} 项）:")
        for c in changed_items:
            if mode == 'uniform':
                print(f"  • {c['spec']}: 在理论值基础上{c['direction']} {c['delta_cents']}分  (×{c['qty']}件)")
            else:
                print(f"  • {c['spec']}: {c['orig_price']:.2f} → {c['new_price']:.2f}  "
                      f"({c['direction']} {c['delta_yuan']:.2f}元/件, ×{c['qty']}件 = "
                      f"±{c['delta_yuan']*c['qty']:.2f}元)")
    else:
        print("\n✨ 原总价已等于目标，无需调整。")

    print(f"\n📁 已保存: {output_path}")
    return output_path, new_total, abs(new_total-target_total)

# ──────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(description='单价优化器 (uniform / minimal)')
    p.add_argument('input',  help='输入Excel')
    p.add_argument('target', type=float, help='目标总价（元）')
    p.add_argument('-o','--output', help='输出路径')
    p.add_argument('--mode', choices=['uniform','minimal'], default='uniform',
                   help='uniform=等比下浮(默认)  minimal=最少项数(发票场景)')
    # 模式A
    p.add_argument('--max-delta', type=int, default=5,
                   help='[uniform] 单价最大偏差分钱，默认5')
    # 模式B
    p.add_argument('--max-change-per-item', type=float, default=None,
                   help='[minimal] 单项最大调整金额（元），默认不限')
    p.add_argument('--max-items', type=int, default=None,
                   help='[minimal] 最多修改几项，默认自动最优')
    # 通用
    p.add_argument('--sheet',      default=0)
    p.add_argument('--header-row', type=int, default=0)
    p.add_argument('--spec-col')
    p.add_argument('--qty-col')
    p.add_argument('--price-col')
    p.add_argument('--total-col')
    args = p.parse_args()

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    process_excel(
        input_path=args.input,
        target_total=args.target,
        output_path=args.output,
        mode=args.mode,
        max_delta_cents=args.max_delta,
        max_change_per_item_yuan=args.max_change_per_item,
        max_items_to_change=args.max_items,
        spec_col=args.spec_col,
        qty_col=args.qty_col,
        price_col=args.price_col,
        total_col=args.total_col,
        header_row=args.header_row,
        sheet_name=sheet,
    )

if __name__ == '__main__':
    main()
