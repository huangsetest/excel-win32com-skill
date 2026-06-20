import os
import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2 - 副本.xlsx"

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)
    sheet = wb.ActiveSheet

    last_row = sheet.UsedRange.Rows.Count
    rows_to_delete = {last_row, last_row - 1}
    print(f"总行数: {last_row}")
    print(f"要删除的行: {rows_to_delete}")

    # 收集要删除的 shape 索引（倒序）
    to_delete_indices = []
    for i in range(1, sheet.Shapes.Count + 1):
        try:
            shape = sheet.Shapes.Item(i)
            tl_row = shape.TopLeftCell.Row
            br_row = shape.BottomRightCell.Row
            name = shape.Name
            print(f"  [{i}] {name}: 行{tl_row}-{br_row}")
            if tl_row in rows_to_delete or br_row in rows_to_delete:
                to_delete_indices.append(i)
                print(f"    -> 标记删除")
        except Exception as e:
            print(f"  [{i}] 无法获取位置: {e}")

    # 按倒序删除 shapes（避免索引变化）
    for idx in sorted(to_delete_indices, reverse=True):
        try:
            sheet.Shapes.Item(idx).Delete()
            print(f"已删除 shape 索引 {idx}")
        except Exception as e:
            print(f"删除 shape {idx} 失败: {e}")

    # 删除行（从下往上）
    sheet.Rows(last_row).Delete()
    sheet.Rows(last_row - 1).Delete()
    print(f"已删除第 {last_row-1} 和 {last_row} 行")

    wb.SaveAs(dst)
    print(f"已保存为: {dst}")
finally:
    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()
    print("完成")
