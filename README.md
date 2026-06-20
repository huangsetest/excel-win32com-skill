# Excel Win32com Skill

使用 `win32com` 客户端操作 Excel 文件，精确控制行/图片位置。

## 功能

- 删除指定行及其关联图片
- 交换两行（含图片同步移动）
- 处理同名 Shape 冲突

## 依赖

```bash
pip install pywin32
```

需要 Windows + Microsoft Excel。

## 使用

```bash
# 删除最后两行及其图片（结果存为新文件）
python scripts/delete_last_two_rows.py 文件.xlsx

# 交换第2、3行（含图片）
python scripts/swap_rows_2_3.py 文件.xlsx
```

## 文件说明

- `SKILL.md` — WorkBuddy Skill 定义文件
- `scripts/delete_last_two_rows.py` — 删除行+图片脚本
- `scripts/swap_rows_2_3.py` — 交换行脚本
