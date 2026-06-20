# Excel Win32com Skill

使用 `win32com` 客户端操作 Excel 文件，精确控制行/图片位置。

## 功能

- 删除指定行及其关联图片
- 交换任意两行（含图片同步移动）
- 处理同名 Shape 冲突

## 依赖

```bash
pip install pywin32
```

需要 Windows + Microsoft Excel。

## 使用

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

### 删除最后两行（固定逻辑）

```bash
python scripts/delete_last_two_rows_fixed.py
```

### 交换第2、3行（固定逻辑）

```bash
python scripts/swap_rows_2_3.py
```

## 文件说明

- `SKILL.md` — WorkBuddy Skill 定义文件
- `README.md` — 本说明文档
- `scripts/delete_row_generic.py` — 删除指定行+图片（通用）
- `scripts/swap_rows_generic.py` — 交换任意两行（通用）
- `scripts/delete_last_two_rows_fixed.py` — 删除最后两行（固定）
- `scripts/swap_rows_2_3.py` — 交换第2、3行（固定）
