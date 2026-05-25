# Translation Engine — Implementation Details

仅当需要编写或调试翻译代码时读取此文件。

## 最长匹配优先算法

```python
candidates = sorted(dict.items(), key=lambda x: -len(x[0]))
i = 0
while i < len(text):
    matched = False
    for cn, en in candidates:
        if text[i:].startswith(cn):
            text = text[:i] + en + text[i+len(cn):]
            i += len(en)
            matched = True
            break
    if not matched:
        i += 1
```

## clear_and_set_text() 内部流程

先复制源格式 → 清除底纹 → 强制 TNR 12pt（按此顺序，标黄/加粗等不会丢失）：

```python
new_rPr = copy.deepcopy(source_rPr)   # 复制标黄、加粗等
_strip_shading(new_rPr)                # 清除底纹
_force_tnr_12pt(new_rPr)               # 强制 TNR 12pt
```

参数：
- `force_font=True`（默认）— 正文，强制 TNR 12pt
- `force_font=False` — 标题，只替换文本和清除底纹，保留原字体字号

## 行距设置

- 正文：`set_spacing(para, True)` → `line=480, lineRule=auto` = 2 倍行距
- 表格：`set_spacing(para, False)` → `line=240, lineRule=auto` = 1 倍行距

## Windows 长路径处理

输出路径可能超过 260 字符时加 `\\?\` 前缀：

```python
abs_dst = os.path.abspath(dst_path)
if len(abs_dst) > 260 and not abs_dst.startswith('\\\\?\\'):
    save_path = '\\\\?\\' + abs_dst
else:
    save_path = abs_dst
doc.save(save_path)
```

## 相关脚本

- `translate_engine.py` — 主要翻译引擎，包含上述所有逻辑
- `fix_fonts.py` — 事后字体和间距修复工具
