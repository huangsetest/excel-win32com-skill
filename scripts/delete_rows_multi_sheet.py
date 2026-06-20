import pythoncom
from win32com.client import Dispatch

src = r"C:\Users\Administrator\Desktop\2.xlsx"
dst = r"C:\Users\Administrator\Desktop\2_删除多行.xlsx"

# 按 Sheet 索引（从1开始）定义要删除的行
# {sheet_index: [行号列表]}
delete_map = {
    1: [3],   # 第1个 Sheet 删除第3行
    2: [4],   # 第2个 Sheet 删除第4行
}

pythoncom.CoInitialize()
excel = Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(src)

    # 打印所有 sheet 名称供参考
    print("所有 Sheet（索引从1开始）:")
    for idx in range(1, wb.Worksheets.Count + 1):
        print(f"  [{idx}] {wb.Worksheets(idx).Name}")

    for sheet_idx, rows in delete_map.items():
        try:
            ws = wb.Worksheets(sheet_idx)
            sheet_name = ws.Name
            print(f"\n处理 Sheet [{sheet_idx}] {sheet_name}")

            # 倒序删除行（避免行号偏移）
            for target_row in sorted(rows, reverse=True):
                # 收集该行图片索引
                to_delete = []
                for i in range(1, ws.Shapes.Count + 1):
                    try:
                        s = ws.Shapes.Item(i)
                        r = s.TopLeftCell.Row
                        if r == target_row:
                            to_delete.append(i)
                    except:
                        pass

                print(f"  找到 {len(to_delete)} 个图片在第{target_row}行，准备删除")

                # 倒序删除图片
                for idx in sorted(to_delete, reverse=True):
                    try:
                        name = ws.Shapes.Item(idx).Name
                        ws.Shapes.Item(idx).Delete()
                    except:
                        pass

                # 删除行
                ws.Rows(target_row).Delete()
                print(f"  已删除第 {target_row} 行")

        except Exception as e:
            print(f"处理 Sheet 索引 {sheet_idx} 失败: {e}")

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
