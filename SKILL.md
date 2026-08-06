---
name: sab-news-audit-skill
description: Generate the complete SAB星系列新闻稿统一标准手册 as a professionally formatted DOCX binary stream, and review submitted 星动态、星故事、星分享、星标杆、星视频稿件 against the manual to return only definite problems and modification suggestions. Use when Codex needs to generate the manual or audit a SAB星系列 news article without filesystem I/O.
---

# SAB新闻稿文档处理

Use only `python-docx` and `openpyxl` beyond the Python standard library. Never import PyYAML or perform YAML loading/serialization. Never read or write local files in the callable functions; return file contents as `bytes`.

## Functions

Call `scripts/sab_news_processor.py::generate_standard_manual()` to generate the complete manual. It accepts no arguments and returns DOCX bytes.

Call `scripts/sab_news_processor.py::review_news_article(article_text, column_type)` to review an article. Pass one of `星动态、星故事、星分享、星标杆、星视频`. Return only a dictionary containing `summary` and `results`. Do not generate DOCX/XLSX files, binary streams, attachments, or download links for an article review.

Return only definite problems actually detected in the submitted text and their modification suggestions. Never include a `需人工确认` section or result item; omit uncertain checks instead of presenting them as issues. Present the review directly in the conversation from `results`.

For responsibility information, use semantic matching at the end of the article: accept `撰稿/供稿` for `文字`, `摄影/摄像/配图/图片` for `摄影`, `编辑/排版` for `编辑`, and `审核/审校` for `审核`. Do not restrict colons, spaces, slashes, or similar separators. For internal-use warnings, accept wording that conveys both “internal use” and “do not circulate externally”; do not require an exact fixed sentence.

Only reject `集团` when it is used to introduce company leadership, such as `集团副总经理××`; require `总公司` in that context. Allow `集团` in entrepreneurship anniversaries, development history, and other non-leadership contexts.

For people mentioned repeatedly, require the department/unit, complete official position, and full name only at the first introduction. Allow later references to use a concise form such as `路总`, `×总`, or the person's name without repeating the department. Do not flag later concise references as missing a title or department. For non-leaders, require department and name at first mention, then allow the name alone.

For `星动态`, determine whether the article is a meeting article from the title's core event. Only a meeting article is checked for meeting name, host, and participant scope in its first substantive paragraph. Leadership visits, inspections, guidance, research visits, and similar event articles are checked against their actual event elements and must not be forced into the meeting template.

## Function-calling schemas

Read `references/function_calling.json` when exposing the functions to a platform. Do not emit or require a YAML configuration file.

## Validation

For Function 1, confirm the returned bytes begin with the ZIP/OOXML signature, can be opened by `python-docx`, contain all seven modules, preserve every non-empty source line, use A4 geometry with 2.54 cm margins, and distinguish the three bracketed rule labels.

For Function 2, confirm the return value contains exactly `summary` and `results`, performs no document or spreadsheet rendering, and includes one structured result per definite issue. Each result must retain problem evidence, standard basis, and modification tips. Do not include any uncertain-review field in `summary` or `results`.
