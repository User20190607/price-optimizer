---
name: price_optimizer
title: 单价调整优化器
description: "给定含数量/单价的Excel报价单和目标总价，自动调整各行单价使总价精确达标。内置GCD可行性预检、模式A（等比均匀下浮+贪心残差分配）、模式B（最少项数分层解）、模式C（单一费率+平衡行）及MILP通用兜底，覆盖投标、发票、结算等场景。"
when_to_use: "当用户需要把Excel报价单/清单的总价调整到指定目标值，或需要对单价进行等比下浮/控价/凑总价时使用。触发词：调整单价、下浮、控价、凑总价、目标总价、报价单优化、发票凑数、投标报价。"
---

# 单价调整优化器

## When to use

- 用户提供 Excel/CSV 报价单，要求把合计总价调整到某个目标值
- 投标场景：所有单价等比均匀下浮，总价精确命中中标限价
- 发票/结算场景：尽量少改行数，只动最少几行凑到目标金额
- 审计友好场景：用单一折扣率 + 一行尾差项，全程可解释

## Runtime requirements

- Browser login required: no
- Sandbox (E2B) required: yes（运行 Python 脚本、读写 Excel）
- MCP required: no
- Page script required: no
- User local folder required: optional（如用户通过 @folder 挂载本地目录则使用）

## Capability routing

| 步骤 | 工具路径 |
|---|---|
| 读取用户上传的 Excel | `read_file` / `e2b_write` 落盘后用 Python 读取 |
| 运行优化算法 | `e2b_bash` 执行 `scripts/optimize_prices.py` |
| 写回带高亮的新 Excel | Python (openpyxl) 在 E2B 内写文件 |
| 输出结果文件 | 保存到 `/mnt/cos/artifacts/`，提供下载链接 |
| 展示结果摘要 | 控制台文字回复，包含误差、调整项数、下浮率 |

## Pre-flight：GCD 可行性判断（必须先做）

在调用任何优化算法前，先执行 GCD 预检：

```
g = gcd(q1, q2, ..., qn)          # 所有数量的最大公约数（分钱单位）
可达格点步长 = g 分钱 = g/100 元
目标是否精确可达 = (target_cents % g == 0)
若不可达 → 自动吸附到最近格点，并告知用户不可消除误差为 X 分
```

详见 `references/algorithm-notes.md`。

## Workflow

### Step 1：接收输入

1. 确认用户提供了 Excel 文件（通过上传或 @folder）和目标总价（元）。
2. 询问调整模式（若用户未指定）：
   - **A（默认）**：均匀下浮——所有行等比缩放
   - **B**：最少项数——只改最少几行（发票场景）
   - **C**：单一费率——一个统一折扣 + 一行平衡项（审计场景）
   - **MILP**：通用兜底——复杂约束时使用

### Step 2：把上传文件落盘

```python
# 把 read_file 的内容以二进制写入 E2B
# 然后用 optimize_prices.py 读取
```

### Step 3：执行优化

```bash
python /mnt/work/.skills/<skill_id>/scripts/optimize_prices.py \
  <input.xlsx> <target_total> \
  --mode <uniform|minimal|single_rate|milp> \
  --output /mnt/cos/artifacts/optimized_<timestamp>.xlsx
```

### Step 4：输出结果

- 控制台打印：原始总价、目标总价、优化后总价、误差、下浮率、被调整项数
- 生成带高亮的 Excel，保存到 `/mnt/cos/artifacts/`
- 向用户展示摘要表格并提供下载

### Step 5：验证

确认 `abs(new_total - target_total) <= g/100`（即在 GCD 格点精度内）。

## Files in this skill

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 本文件，主工作流 |
| `scripts/optimize_prices.py` | 完整优化脚本（四种算法 + Excel 读写） |
| `references/algorithm-notes.md` | GCD 理论、各算法数学推导、适用边界说明 |

## Output contract

- 必须输出：优化后 Excel 文件路径
- 必须包含：原始总价、目标总价、优化总价、绝对误差（≤ g/100 元）、整体下浮率、被调整项数
- Excel 中：被修改行橙色/黄色高亮，新增"优化单价"和"变动"两列，末尾汇总行

## Failure handling

| 情况 | 处理 |
|---|---|
| 目标不在 GCD 格点 | 自动吸附到最近格点，告知不可消除误差 |
| 模式B单行/双行无解 | 自动降级到贪心多行分配，不崩溃 |
| 数量为小数 | 统一乘以100取整到最小计量单位后再建模 |
| openpyxl/pandas/pulp 未安装 | `e2b_bash` 先 `pip install` 再运行 |
| Excel 列名无法自动识别 | 提示用户用 `--qty-col` / `--price-col` 手动指定 |
| 目标总价明显超出合理范围 | 提示用户确认，不静默继续 |
