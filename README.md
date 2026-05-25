# Faithful Doc Translate Skill

A Claude Code skill for **faithful Chinese-to-English document translation** — preserving original formatting (bold, underline, yellow highlight, alignment, borders) while producing accurate, professional English translations.

## Installation

```bash
claude add skill https://github.com/Ash2s/faithful-doc-translate-skill
```

Once installed, the skill auto-triggers when you say phrases like "翻译文件", "中译英", "翻译合同", "翻译专利", etc. Or invoke it directly with `/faithful-doc-translate`.

## Core Features

- **Glossary-driven**: Longest-match-first dictionary lookup for consistent terminology across the document
- **Format preservation**: Retains yellow highlight, bold, underline, strikethrough, borders, alignment
- **Background cleanup**: Removes gray/patterned shading (shading causes artifacts when scanned by official authorities)
- **Font enforcement**: Forces Times New Roman 12pt, double line spacing for body, single spacing for tables
- **Quote normalization**: Unifies straight/curly quotes to avoid lookup mismatches
- **OCR error detection**: Catches common OCR misreadings (e.g. `203年` → `2023年`)
- **Mixed-cell translation**: Handles tables with partially translated Chinese-English content
- **Chinese residual scan**: Post-translation quality check — detects any missed Chinese characters
- **Three translation strategies**: dictionary lookup (terms/labels), full-sentence translation (narrative text), and hybrid cell translation (mixed-content tables)

## Structure

```
├── skill.md              # Skill definition (Claude Code entry point)
├── translate_engine.py   # Core translation engine
├── fix_fonts.py          # Post-processing font & spacing fixer
├── references/           # Reference documentation
│   ├── translation-engine.md
│   ├── ocr-errors.md
│   ├── debugging.md
│   └── troubleshooting.md
└── evals/                # Evaluation data
```

## Usage

1. Prepare a glossary file (.xlsx or .csv): Column A = Chinese term, Column B = English translation
2. Prepare the source document (.docx or .pdf)
3. Tell Claude: "翻译这份文件，术语表在这里：[path/to/glossary]"
4. The skill handles the rest — term lookup, format preservation, translation, quality check

## Requirements

- Claude Code
- Python 3.8+ with `python-docx`
