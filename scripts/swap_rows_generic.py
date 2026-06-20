import sys
import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2_交换25行.xlsx"

row_a = 2
row_b = 5

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)
    ws = wb.ActiveSheet

    # 给 row_a、row_b 行的图片加唯一前缀
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        try:
            r = s.TopLeftCell.Row
            if r == row_a:
                s.Name = f"__ROWA__{s.Name}"
            elif r == row_b:
                s.Name = f"__ROWB__{s.Name}"
        except:
            pass

    # 用 temp 行（第7行）交换
    ws.Rows(7).Insert()

    # row_a → temp
    ws.Rows(row_a).Copy()
    ws.Rows(7).PasteSpecial(-4104)
    excel.CutCopyMode = False

    # row_b → row_a
    ws.Rows(row_b).Copy()
    ws.Rows(row_a).PasteSpecial(-4104)
    excel.CutCopyMode = False

    # temp（原 row_a）→ row_b
    ws.Rows(7).Copy()
    ws.Rows(row_b).PasteSpecial(-4104)
    excel.CutCopyMode = False

    ws.Rows(7).Delete()

    # 移动 __ROWA__ 图片到 row_b 位置
    row_b_top = ws.Rows(row_b).Top
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        if s.Name.startswith("__ROWA__"):
            orig = s.Name[9:]
            s.Top = row_b_top
            s.Name = orig

    # 移动 __ROWB__ 图片到 row_a 位置
    row_a_top = ws.Rows(row_a).Top
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        if s.Name.startswith("__ROWB__"):
            orig = s.Name[9:]
            s.Top = row_a_top
            s.Name = orig

    # 验证
    print(f"交换完成: 第{row_a}行 <-> 第{row_b}行")
    print(f"总行数: {ws.UsedRange.Rows.Count}")
    print(f"图片数: {ws.Shapes.Count}")
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes.Item(i)
        try:
            r = s.TopLeftCell.Row
            print(f"  [{i}] {s.Name}: 行{r}")
        except:
            print(f"  [{i}] {s.Name}: 位置未知")

    wb.SaveAs(dst)
    print(f"已保存为: {dst}")
finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
    pythoncom.CoUninitialize()
    print("完成")
