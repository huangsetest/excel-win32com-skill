import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2_格式统一.xlsx"

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)

    for sid in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(sid)
        sheet_name = ws.Name
        print(f"\n处理 Sheet: {sheet_name}")

        # 读取 E3 的格式作为参考
        if sid == 1:
            e3 = ws.Cells(3, 5)
            ref_font_name = e3.Font.Name
            ref_font_size = e3.Font.Size
            ref_bold = e3.Font.Bold
            ref_wrap = e3.WrapText
            ref_halign = e3.HorizontalAlignment
            ref_valign = e3.VerticalAlignment
            print(f"  参考格式: 字体={ref_font_name}, 大小={ref_font_size}, 加粗={ref_bold}, 换行={ref_wrap}")

        # 处理该 Sheet 的所有数据行（跳过第1行表头）
        last_row = ws.UsedRange.Rows.Count
        for row in range(2, last_row + 1):
            cell = ws.Cells(row, 5)  # E列
            val = cell.Value
            if not val:
                continue

            # 将 * 符号替换为数字序号
            lines = val.split("\n")
            new_lines = []
            idx = 1
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # 去掉开头的 * 和多余空格
                if line.startswith("*"):
                    line = line[1:].strip()
                # 如果已经有数字序号，保留；否则加序号
                new_lines.append(f"{idx}. {line}")
                idx += 1

            new_val = "\n".join(new_lines)
            cell.Value = new_val

            # 应用 E3 的格式
            cell.Font.Name = ref_font_name
            cell.Font.Size = ref_font_size
            cell.Font.Bold = ref_bold
            cell.WrapText = ref_wrap
            cell.HorizontalAlignment = ref_halign
            cell.VerticalAlignment = ref_valign

            print(f"  行{row} E列: 已统一格式 ({idx-1} 条)")

        # 统一 E 列列宽和各行行高（与 E3 一致）
        ws.Columns(5).ColumnWidth = 136.38
        for row in range(2, last_row + 1):
            ws.Rows(row).RowHeight = 258.5

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
