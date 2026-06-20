---
name: "excel-win32com"
description: "使用 win32com 操作 Excel 文件，支持删除指定行及其关联图片、交换任意两行（含图片）。适用于需要精确控制 Excel 行/图片位置的场景。"
---

# Excel Win32com Skill

使用 `win32com` 客户端操作 Excel，支持精确的行/图片操作。

## 功能

- **删除指定行及其图片**：按行号删除行，同时删除该行锚定的所有图片
- **交换任意两行（含图片）**：交换两行数据，同时移动图片到对应位置
- **支持同名图片处理**：通过临时重命名避免同名 Shape 冲突

## 依赖

```bash
pip install pywin32
```

需要 Windows + Microsoft Excel 已安装。

## 使用方法

### 删除指定行及其图片

```python
# 修改脚本中的 target_row 变量指定行号
python scripts/delete_row_generic.py
```

### 交换任意两行（含图片）

```python
# 修改脚本中的 row_a、row_b 变量指定行号
python scripts/swap_rows_generic.py
```

### 删除最后两行（固定逻辑）

```python
python scripts/delete_last_two_rows_fixed.py
```

### 交换第2、3行（固定逻辑）

```python
python scripts/swap_rows_2_3.py
```

## 脚本说明

### `scripts/delete_row_generic.py`
删除 Excel 指定行，同时删除该行锚定的所有图片。
- 通过 `TopLeftCell.Row` 判断图片所属行
- 按索引倒序删除 Shape，避免索引偏移
- 结果保存为新文件，不修改原文件

### `scripts/swap_rows_generic.py`
交换 Excel 任意两行，同时移动图片到交换后的位置。
- 先给两行图片加 `__ROWA__` / `__ROWB__` 前缀避免同名冲突
- 用临时行完成行数据交换
- 按行的 `Top` 像素值移动图片到目标位置
- 恢复图片原名

### `scripts/delete_last_two_rows_fixed.py`
删除 Excel 最后两行（固定逻辑），同时删除这两行锚定的所有图片。

### `scripts/swap_rows_2_3.py`
交换 Excel 第2、3行（固定逻辑），同时移动图片到对应位置。

## 注意事项

- 原文件不会被修改，结果保存为新文件
- 同名图片（如两个 `Picture 1`）会自动重命名处理，不会误删
- 需要退出所有 Excel 进程后再运行，避免 COM 锁文件
- 修改脚本中的 `src` 和 `dst` 路径以适配你的文件
