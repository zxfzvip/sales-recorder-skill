import os
from openpyxl import load_workbook

def record_inventory(target: str, product: str = None, qty: float = None, price: float = None, expr_count: float = None, expr_price: float = None):
    base_path = "/Users/mac/Desktop/拿货记数"
    file_path = os.path.join(base_path, f"{target}.xlsx")
    if not os.path.exists(file_path):
        return f"找不到文件：{target}.xlsx"
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        next_row = 2
        for row in range(2, 1000):
            val_b = ws.cell(row, 2).value
            val_f = ws.cell(row, 6).value
            is_b_empty = val_b is None or (isinstance(val_b, str) and val_b.startswith('='))
            is_f_empty = val_f is None or (isinstance(val_f, str) and val_f.startswith('='))
            if is_b_empty and is_f_empty:
                next_row = row
                break
        
        if product:
            ws.cell(next_row, 2).value = product
            ws.cell(next_row, 3).value = qty or 0
            ws.cell(next_row, 4).value = price or 0
            ws.cell(next_row, 5).value = f"=C{next_row}*D{next_row}"
        if expr_count is not None:
            ws.cell(next_row, 6).value = expr_count
            ws.cell(next_row, 7).value = expr_price or 0
            ws.cell(next_row, 8).value = f"=F{next_row}*G{next_row}"
        
        wb.save(file_path)
        
        # 计算小计
        subtotal = (qty or 0) * (price or 0) if qty and price else None
        expr_subtotal = (expr_count or 0) * (expr_price or 0) if expr_count and expr_price else None
        
        # 构建回复
        result = f"✅ 已添加到 {target}.xlsx 第{next_row}行！\n\n"
        
        if product and qty and price:
            result += f"| 商品 | 数量 | 单价 | 小计 |\n"
            result += f"|------|------|------|------|\n"
            result += f"| {product} | {qty} | {price} | ¥{subtotal:.2f} |\n"
        
        if expr_count and expr_price:
            result += f"\n| 快递单数 | 快递价格 | 小计 |\n"
            result += f"|----------|----------|------|\n"
            result += f"| {int(expr_count)} | {expr_price} | ¥{expr_subtotal:.2f} |\n"
        
        result += "\n继续 📝"
        return result
        
    except Exception as e:
        return f"错误: {str(e)}"

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "拿货记录"
    product = sys.argv[2] if len(sys.argv) > 2 else None
    qty = float(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3] else None
    price = float(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4] else None
    expr_count = float(sys.argv[5]) if len(sys.argv) > 5 and sys.argv[5] else None
    expr_price = float(sys.argv[6]) if len(sys.argv) > 6 and sys.argv[6] else None
    print(record_inventory(target, product, qty, price, expr_count, expr_price))
