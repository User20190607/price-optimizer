# Price Optimizer · 单价优化工具

基于 MILP（混合整数线性规划）算法的 Excel 报价单价优化工具。

## 功能

- **模式A：等比均匀下浮** — 投标报价场景，所有项等比例下调
- **模式B：最少项数调整** — 发票对账场景，只改最少几项来达到目标总价

## 使用方法

```bash
python optimize_prices.py "报价单.xlsx" 目标总价

# 模式B
python optimize_prices.py "报价单.xlsx" 目标总价 --mode minimal
```

## 依赖

```bash
pip install pulp openpyxl pandas
```
