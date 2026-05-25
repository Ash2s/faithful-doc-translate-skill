# Troubleshooting: Common Translation Issues

## 1. Run-Level Formatting Destruction

**Symptoms:**
- Yellow highlight on Chinese names lost in English output
- Partial bold (e.g. only the label "Abstract:" was bold) becomes full-paragraph bold

**Root cause:** OOXML splits text with mixed formatting into multiple `<w:r>` (run) elements. The old engine copied formatting from only the first run with `rPr`, blindly applying it to the entire translated text.

**Fix:** `clear_and_set_text()` in translate_engine.py now uses **consensus rules**:
- Bold: applied only if **ALL** text runs are bold
- Highlight: applied if **ANY** run has highlight
- Italic: applied only if **ALL** text runs are italic
- Underline: applied if **ANY** run has underline

**Detection:** Before/after comparison — check source paragraph's run-level formatting, verify EN output matches the consensus.

---

## 2. Bracket Labels Not Translated (Split-Run Tags)

**Symptoms:**
- `【图片】` / `【条形码】` / `【图标】` / `【盖章】` remain in Chinese in the output
- Grep shows the full string `【图片】` in `para.text` but no single run contains it

**Root cause:** Word splits these bracket labels across multiple runs (e.g. `【` + `图片` + `】`) for font/formatting reasons. Run-level matching (`'【图片】' in run.text`) fails on every run.

**Fix:** Use paragraph-level detection: `para.text.strip() == '【图片】'`. The function `fix_cn_labels(doc)` in translate_engine.py handles this.

**Verify:** Run `verify_no_cn()` after `fix_cn_labels()` — any remaining `【` characters indicate new bracket labels not yet in `CN_LABEL_MAP`.

---

## 3. Paragraph-Index Translation Gaps

**Symptoms:**
- Journal headers, article metadata, title remnants, split figure caption lines remain in Chinese
- These paragraphs are not technical content — they just weren't in the translation dictionary

**Root cause:** When using paragraph-index-based translation maps (`{idx: en_text}`), any paragraph not explicitly listed is left untouched. Edge paragraphs (page headers, split captions, table footnotes) are the most frequently missed.

**Fix:** `verify_no_cn()` on every output file is the only safety net. Review each CN residual — if it's journal metadata or a structural label, translate it. If it's a technical equation, reference entry, or other paper's TOC line, it's intentionally preserved as evidence.

---

## 4. Bash Heredoc Quote Breakage

**Symptoms:**
- `unexpected EOF while looking for matching` error when running inline Python via bash
- Apostrophes in English text (e.g. "Master's degree") break heredoc quoting

**Root cause:** bash heredoc inside `python -c "..."` has conflicting quote requirements. Single quotes, double quotes, and apostrophes in the Python string interact with bash's parsing.

**Fix:** Write Python code to a `.py` file (using the Write tool), then execute it. Never embed long Python with English text in a bash heredoc.

---

## 5. Locked File PermissionError

**Symptoms:**
- `PermissionError` when saving a .docx file
- File is open in Microsoft Word

**Fix:** Close the file in Word before re-running the translation. The file handle lock in Word prevents python-docx from writing.
