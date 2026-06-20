import os
import pythoncom
from win32com.client import Dispatch

src = r'C:\Users\Administrator\Desktop\2.xlsx'
dst = r'C:\Users\Administrator\Desktop\2_提取到Sheet3.xlsx'

# 如果目标文件已存在，先删除
if os.path.exists(dst):
    try:
        os.remove(dst)
    except PermissionError:
        print('请先关闭文件:', dst)
        exit(1)

pythoncom.CoInitialize()
excel = Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(src)

# 确保有 Sheet3
sheet3_exists = False
for i in range(1, wb.Sheets.Count + 1):
    if wb.Sheets(i).Name == 'Sheet3':
        sheet3_exists = True
        break

if not sheet3_exists:
    wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count)).Name = 'Sheet3'
    print('已新建 Sheet3')
else:
    print('Sheet3 已存在')

ws3 = wb.Sheets('Sheet3')

# 先复制列宽（从 Sheet1）
ws1 = wb.Sheets(1)
for c in range(1, 20):
    try:
        ws3.Columns(c).ColumnWidth = ws1.Columns(c).ColumnWidth
    except:
        break

# 复制行：先处理 Sheet2 第2行 -> Sheet3 第1行
print('\n复制 Sheet2 第2行 -> Sheet3 第1行...')
ws2 = wb.Sheets(2)
ws2.Activate()
ws2.Rows(2).Copy()
ws3.Activate()
ws3.Rows(1).Insert()
print('  完成')

# 复制行：再处理 Sheet1 第5行 -> Sheet3 第2行
print('复制 Sheet1 第5行 -> Sheet3 第2行...')
ws1.Activate()
ws1.Rows(5).Copy()
ws3.Activate()
ws3.Rows(2).Insert()
print('  完成')

# 设置自动换行和行高
print('\n设置格式...')
for row in [1, 2]:
    ws3.Rows(row).RowHeight = ws1.Rows(5).RowHeight
    for col in range(1, 20):
        try:
            ws3.Cells(row, col).WrapText = True
        except:
            break

# 删除 Sheet1 和 Sheet2
print('\n删除 Sheet1 和 Sheet2...')
try:
    wb.Sheets(ws1.Name).Delete()
except Exception as e:
    print(f'  删除 Sheet1 失败: {e}')

try:
    wb.Sheets(ws2.Name).Delete()
except Exception as e:
    print(f'  删除 Sheet2 失败: {e}')

# 激活 Sheet3
ws3.Activate()

# 保存
print('\n保存到:', dst)
wb.SaveAs(dst)
wb.Close()
excel.Quit()
pythoncom.CoUninitialize()

print('\n完成！')
print('文件已保存至:', dst)
