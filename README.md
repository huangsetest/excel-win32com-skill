# excel-win32com-skill

使用 `win32com` 客户端操作 Excel 文件，支持精确的行/图片操作。

## 功能

- **删除指定行及其图片**：按行号删除行，同时删除该行锚定的所有图片
- **交换任意两行（含图片）**：交换两行数据，同时移动图片到对应位置
- **多 Sheet 批量操作**：同时对多个 Sheet 执行删除行操作
- **统一单元格格式**：将指定列的内容格式统一为参考行的格式（如序号格式、字体、对齐等）
- **提取行到新 Sheet**：将多个 Sheet 的指定行提取出来，合并到一个新 Sheet，并删除原 Sheet
- **支持同名图片处理**：通过临时重命名避免同名 Shape 冲突
- **图片自适应单元格尺寸**：提取行时自动调整图片尺寸与单元格匹配

## 依赖

```bash
pip install pywin32
```

需要 Windows + Microsoft Excel 已安装。

## 聊天命令举例（直接对 AI 说）

以下是实际对话中的命令举例，可以直接对 AI 说，AI 会自动调用对应脚本完成操作。

### 删除行

```
"C:\Users\Administrator\Desktop\2.xlsx" 删除第五行
```

```
"C:\Users\Administrator\Desktop\2.xlsx" 将 sheet1 第3行删掉 和 sheet2 的第4行删掉
```

### 交换行

```
"C:\Users\Administrator\Desktop\2.xlsx" 交换第2行和第5行
```

```
"C:\Users\Administrator\Desktop\2.xlsx" 将第2行和第3行交换一下
```

### 统一格式

```
"C:\Users\Administrator\Desktop\2.xlsx" 将所有行的 description 的文字，序号格式和 E3 一致
```

（说明：保留原内容，只把 `*` 符号改成 `1. 2. 3. ...` 数字序号格式，同时统一字体/换行/列宽/行高）

### 提取行到新 Sheet

```
"C:\Users\Administrator\Desktop\2.xlsx" 将 sheet1 第5行，sheet2 第2行提取出来
```

（说明：AI 会自动提取指定行到新 Sheet3，复制列宽、开启自动换行、调整图片尺寸，并删除原 Sheet1 和 Sheet2）

### 删除最后两行（含图片）

```
"C:\Users\Administrator\Desktop\2 - 副本.xlsx" 用 win32com 将最后两行连图片删掉，其他行的图片保留
```

### 通用说明

- 文件路径用引号包裹（路径含空格时必须加引号）
- 不修改源文件，结果自动保存为新文件（如 `2_删除第5行.xlsx`）
- 支持中文自然语言描述，AI 会自动识别意图并调用对应脚本

## 使用方法（手动运行脚本）

### 删除指定行及其图片

修改 `scripts/delete_row_generic.py` 中的 `target_row` 变量，然后运行：

```bash
python scripts/delete_row_generic.py
```

### 交换任意两行（含图片）

修改 `scripts/swap_rows_generic.py` 中的 `row_a`、`row_b` 变量，然后运行：

```bash
python scripts/swap_rows_generic.py
```

### 多 Sheet 批量删除行

修改 `scripts/delete_rows_multi_sheet.py` 中的 `delete_map`（按 Sheet 索引），然后运行：

```bash
python scripts/delete_rows_multi_sheet.py
```

示例 `delete_map` 配置：
```python
delete_map = {
    1: [3],   # 第1个 Sheet 删除第3行
    2: [4],   # 第2个 Sheet 删除第4行
}
```

### 统一单元格格式（如序号格式）

修改 `scripts/normalize_desc_format.py` 中的 `ref_row`、`target_rows`、`col`、`sheets` 变量，然后运行：

```bash
python scripts/normalize_desc_format.py
```

示例：将 E 列所有行的 `*` 符号格式统一为 E3 的数字序号格式（`1. 2. 3. ...`），同时统一字体、换行、列宽、行高。

### 提取行到新 Sheet（含图片+格式）

修改 `scripts/extract_rows_to_sheet3.py` 中的 `tasks` 列表，然后运行：

```bash
python scripts/extract_rows_to_sheet3.py
```

示例 `tasks` 配置：
```python
tasks = [
    {'sheet_idx': 2, 'row_no': 2},  # Sheet2 第2行 -> 新Sheet 第1行
    {'sheet_idx': 1, 'row_no': 5},  # Sheet1 第5行 -> 新Sheet 第2行
]
```

脚本会自动：
- 创建/使用 `Sheet3` 作为目标 Sheet
- 复制指定行（含图片）到新 Sheet
- 按源 Sheet1 的列宽设置新 Sheet
- 开启自动换行
- 调整图片尺寸与单元格匹配
- 删除原 Sheet1 和 Sheet2

### 删除最后两行（固定逻辑）

```bash
python scripts/delete_last_two_rows_fixed.py
```

### 交换第2、3行（固定逻辑）

```bash
python scripts/swap_rows_2_3.py
```

## 脚本说明

| 脚本 | 功能 |
|------|------|
| `scripts/delete_row_generic.py` | 删除 Excel 指定行，同时删除该行锚定的所有图片 |
| `scripts/swap_rows_generic.py` | 交换 Excel 任意两行，同时移动图片到交换后的位置 |
| `scripts/delete_rows_multi_sheet.py` | 多 Sheet 批量删除行，支持同时操作多个 Sheet |
| `scripts/normalize_desc_format.py` | 统一指定列的内容格式，将 `*` 符号替换为数字序号，同时统一字体、自动换行、列宽、行高 |
| `scripts/extract_rows_to_sheet3.py` | 将多个 Sheet 的指定行提取到新 Sheet（Sheet3），含图片复制、列宽设置、自动换行、图片尺寸自适应 |
| `scripts/delete_last_two_rows_fixed.py` | 删除 Excel 最后两行（固定逻辑） |
| `scripts/swap_rows_2_3.py` | 交换 Excel 第2、3行（固定逻辑） |

## 注意事项

- 原文件不会被修改，结果保存为新文件
- 同名图片（如两个 `Picture 1`）会自动重命名处理，不会误删
- 需要退出所有 Excel 进程后再运行，避免 COM 锁文件
- 修改脚本中的 `src` 和 `dst` 路径以适配你的文件
- 组合图形（Group Shape）也能正确复制
- 图片会自动调整为与单元格相同的尺寸
