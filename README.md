# Faithful Doc Translate Skill

A Claude Code skill for **faithful document translation** — preserving original formatting (bold, underline, yellow highlight, alignment, borders) while producing accurate, professional English translations.

## Core Features

- **Glossary-driven**: Longest-match-first dictionary lookup for consistent terminology
- **Format preservation**: Retains yellow highlight, bold, underline, strikethrough, borders, alignment
- **Background cleanup**: Removes gray/patterned shading (required for immigration authority scanning)
- **Font enforcement**: Forces Times New Roman 12pt, double line spacing for body, single spacing for tables
- **Quote normalization**: Unifies straight/quoted quotes to avoid lookup mismatches
- **OCR error detection**: Catches common OCR misreadings (e.g. `203年` → `2023年`)
- **Mixed-cell translation**: Handles tables with partially translated Chinese-English content
- **Chinese residual scan**: Post-translation quality check — detects any missed Chinese characters
- **Three translation strategies**: dictionary lookup, full-sentence translation, and hybrid cell translation

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

## Requirements

- Claude Code (designed for Claude Code skill system)
- Python 3.8+ with `python-docx`
