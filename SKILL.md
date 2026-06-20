---
name: "excel-win32com"
description: "使用 win32com 操作 Excel 文件，支持删除指定行及其关联图片、交换任意两行（含图片）、多 Sheet 批量操作。适用于需要精确控制 Excel 行/图片位置的场景。"
---

# Excel Win32com Skill

使用 `win32com` 客户端操作 Excel，支持精确的行/图片操作。

## 功能

- **删除指定行及其图片**：按行号删除行，同时删除该行锚定的所有图片
- **交换任意两行（含图片）**：交换两行数据，同时移动图片到对应位置
- **多 Sheet 批量操作**：同时对多个 Sheet 执行删除行操作
- **支持同名图片处理**：通过临时重命名避免同名 Shape 冲突

## 依赖

```bash
pip install pywin32
```

需要 Windows + Microsoft Excel 已安装。

## 使用方法

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

### 删除最后两行（固定逻辑）

```bash
python scripts/delete_last_two_rows_fixed.py
```

### 交换第2、3行（固定逻辑）

```bash
python scripts/swap_rows_2_3.py
```

## 脚本说明

### `scripts/delete_row_generic.py`
删除 Excel 指定行，同时删除该行锚定的所有图片。

### `scripts/swap_rows_generic.py`
交换 Excel 任意两行，同时移动图片到交换后的位置。

### `scripts/delete_rows_multi_sheet.py`
多 Sheet 批量删除行，支持同时操作多个 Sheet。

### `scripts/delete_last_two_rows_fixed.py`
删除 Excel 最后两行（固定逻辑）。

### `scripts/swap_rows_2_3.py`
交换 Excel 第2、3行（固定逻辑）。

## 注意事项

- 原文件不会被修改，结果保存为新文件
- 同名图片（如两个 `Picture 1`）会自动重命名处理，不会误删
- 需要退出所有 Excel 进程后再运行，避免 COM 锁文件
- 修改脚本中的 `src` 和 `dst` 路径以适配你的文件
