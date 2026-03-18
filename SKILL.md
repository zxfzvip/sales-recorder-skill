---
name: sales-recorder
description: 记录每日销售数据到Excel表格。用于用户发送商品销售信息（日期、商品、数量、单价、快递单数、快递费）时，自动录入到桌面「拿货记数/弟弟.xlsx」表格，包含公式自动计算小计。
---

# 销售记数

## 表格位置与对应商品

| 表格 | 商品 |
|------|------|
| `/Users/mac/Desktop/拿货记数/弟弟.xlsx` | 风流果、川井依克多因、高潮液、延时膏、润滑包等 |
| `/Users/mac/Desktop/拿货记数/央央.xlsx` | 润滑油 |

### 判断逻辑
- 如果商品是"润滑油" → 录入央央.xlsx
- 其他商品 → 录入弟弟.xlsx

## 表格结构

| A列 | B列 | C列 | D列 | E列 | F列 | G列 | H列 |
|-----|-----|-----|-----|-----|-----|-----|-----|
| 日期 | 商品 | 数量 | 单价 | 合计 | 快递单数 | 价格 | 合计 |

- E列公式：`=C列*D列`
- H列公式：`=F列*G列`
- 第101行有汇总公式：`=SUM(E2:E100)`、`=SUM(F2:F100)`、`=SUM(H2:H100)`

## 录入规则

1. **找空行**：从第2行开始扫描，找到B列（商品）第一个为空的行
2. **日期**：当天第一条记录写日期（如"18号"），后续同天记录留空
3. **商品**：必填
4. **数量**：必填
5. **单价**：必填
6. **快递单数**：用户未提及则留空
7. **快递价格**：用户未提及则留空
8. **公式**：C列和D列有值后，E列自动写入公式 `=C{row}*D{row}`
9. **快递公式**：F列和G列都有值后，H列自动写入公式 `=F{row}*G{row}`

## 解析用户消息格式

用户可能发送的格式：
- `3月18号 风流果数量10价格5快递10价格2.8`
- `风流果数量10价格5快递10`
- `高潮液数量10价格3`

解析逻辑：
1. 提取日期：`\d+月?\d*号?` → 如"18号"、"3月18号"
2. 提取商品：匹配商品名称（风流果、川井依克多因、高潮液、延时膏、润滑包等）
3. 提取数量：`数量(\d+)`
4. 提取单价：`价格(\d+\.?\d*)`
5. 提取快递单数：`快递(\d+)`
6. 提取快递价格：第二个`价格(\d+\.?\d*)`或`快递价格(\d+\.?\d*)`

## 录入流程

```python
from openpyxl import load_workbook

# 判断商品对应表格
if product == "润滑油":
    file_path = '/Users/mac/Desktop/拿货记数/央央.xlsx'
else:
    file_path = '/Users/mac/Desktop/拿货记数/弟弟.xlsx'

wb = load_workbook(file_path)
ws = wb.active

# 找空行
next_row = 2
for row in range(2, 110):
    if ws.cell(row, 2).value is None:
        next_row = row
        break

# 写入数据
ws.cell(next_row, 1).value = date或None  # A列
ws.cell(next_row, 2).value = product     # B列
ws.cell(next_row, 3).value = qty         # C列
ws.cell(next_row, 4).value = price      # D列
ws.cell(next_row, 5).value = f"=C{next_row}*D{next_row}"  # E列

if expr_count:
    ws.cell(next_row, 6).value = expr_count  # F列
    ws.cell(next_row, 7).value = expr_price  # G列
    ws.cell(next_row, 8).value = f"=F{next_row}*G{next_row}"  # H列

wb.save('/Users/mac/Desktop/拿货记数/弟弟.xlsx')
```

## 确认回复

录入完成后，回复格式：
```
✅ 已添加到第{row}行（{表格名}）！

| 商品 | 数量 | 单价 | 小计 | 快递单数 |
|------|------|------|------|----------|
| {商品} | {数量} | {单价} | ¥{小计} | {快递单数或"（留空）"} |
```

然后提示用户继续：`继续 📝`
