import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2_交换23行.xlsx"

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)
    ws = wb.ActiveSheet

    # 给第2、3行的图片加唯一前缀，避免同名冲突
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        try:
            r = s.TopLeftCell.Row
            if r == 2:
                s.Name = f"__ROW2__{s.Name}"
            elif r == 3:
                s.Name = f"__ROW3__{s.Name}"
        except:
            pass

    # 用 temp 行（第6行）交换行数据
    ws.Rows(6).Insert()

    ws.Rows(2).Copy()
    ws.Rows(6).PasteSpecial(-4104)
    excel.CutCopyMode = False

    ws.Rows(3).Copy()
    ws.Rows(2).PasteSpecial(-4104)
    excel.CutCopyMode = False

    ws.Rows(6).Copy()
    ws.Rows(3).PasteSpecial(-4104)
    excel.CutCopyMode = False

    ws.Rows(6).Delete()

    # 关键：现在行数据已交换，读取当前第2行和第3行的 Top 值
    row2_top = ws.Rows(2).Top
    row3_top = ws.Rows(3).Top
    print(f"目标行位置: 第2行 Top={row2_top}, 第3行 Top={row3_top}")

    # 移动 __ROW2__ 图片到第3行位置
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        if s.Name.startswith("__ROW2__"):
            orig = s.Name[9:]
            s.Top = row3_top
            s.Name = orig
            print(f"  第2行图片 {orig} -> 第3行位置 (Top={row3_top})")

    # 移动 __ROW3__ 图片到第2行位置
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        if s.Name.startswith("__ROW3__"):
            orig = s.Name[9:]
            s.Top = row2_top
            s.Name = orig
            print(f"  第3行图片 {orig} -> 第2行位置 (Top={row2_top})")

    # 验证
    print("\n验证结果:")
    print(f"  总行数: {ws.UsedRange.Rows.Count}")
    print(f"  图片总数: {ws.Shapes.Count}")
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        try:
            r = s.TopLeftCell.Row
            c = s.TopLeftCell.Column
            print(f"  [{i}] {s.Name}: 行{r} 列{c}, Top={s.Top}")
        except:
            print(f"  [{i}] {s.Name}: 位置未知, Top={s.Top}")

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
