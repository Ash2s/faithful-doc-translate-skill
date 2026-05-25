# Debugging Guide

仅当翻译匹配失败或输出异常时读取此文件。

## Windows 终端编码问题

Windows 终端默认编码为 GBK，中文字符会显示为乱码。推荐将脚本输出写到 UTF-8 文件再用 Read 工具读取：

```python
with open('debug_output.txt', 'w', encoding='utf-8') as f:
    f.write(result)
```

## 翻译字典 key 匹配失败调试步骤

当一段原文在输出文档中保持中文未翻译时：

1. **打印 repr**：`print(repr(paragraph.text))` 查看完整文本、不可见字符、引号类型
2. **逐字符检查**：遍历段落文本，输出每个字符的 `ord()` 和 `repr()`，特别是引号位置
3. **与字典 key 对比**：用 `==` 直接比较；如果 False，找出第一个差异字符
4. **检查段落拆分**：对照源文档，确认该段是否被分成了多个物理段落
5. **检查 OCR 误差**：对差异字符做形近字判断，参见 `references/ocr-errors.md`
