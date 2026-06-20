import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2_删除第5行.xlsx"

target_row = 5

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)
    ws = wb.ActiveSheet

    # 收集 target_row 行的图片索引（倒序）
    to_delete = []
    for i in range(1, ws.Shapes.Count + 1):
        try:
            s = ws.Shapes.Item(i)
            r = s.TopLeftCell.Row
            if r == target_row:
                to_delete.append(i)
                print(f"标记删除图片: {s.Name} (索引 {i})")
        except:
            pass

    # 按倒序删除图片
    for idx in sorted(to_delete, reverse=True):
        try:
            name = ws.Shapes.Item(idx).Name
            ws.Shapes.Item(idx).Delete()
            print(f"已删除图片: {name}")
        except Exception as e:
            print(f"删除图片索引 {idx} 失败: {e}")

    # 删除行
    ws.Rows(target_row).Delete()
    print(f"已删除第 {target_row} 行")

    # 验证
    print(f"\n验证: 总行数={ws.UsedRange.Rows.Count}, 剩余图片数={ws.Shapes.Count}")
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        try:
            r = s.TopLeftCell.Row
            print(f"  [{i}] {s.Name}: 行{r}")
        except:
            print(f"  [{i}] {s.Name}: 位置未知")

    wb.SaveAs(dst)
    print(f"\n已保存为: {dst}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
    pythoncom.CoUninitialize()
    print("完成")
