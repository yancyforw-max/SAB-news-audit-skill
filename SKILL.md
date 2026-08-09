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

For responsibility information, use semantic matching at the end of the article: accept `撰稿/供稿` for `文字`, `摄影/摄像/配图/图片` for `摄影`, `编辑/排版/剪辑` for `编辑`, and `审核/审校` for `审核`. Do not restrict colons, spaces, slashes, or similar separators. Apply the same semantic rule to `星视频`; never require fixed full-width colons. For internal-use warnings, accept wording that conveys both “internal use” and “do not circulate externally”; do not require an exact fixed sentence. Treat `经营、财务、人事、资产、生产技术、订单、市场策略、年度、半年度、年底、周年` as internal-warning trigger words.

Only reject `集团` when it is used to introduce company leadership, such as `集团副总经理××`; require `总公司` in that context. Allow `集团` in entrepreneurship anniversaries, development history, and other non-leadership contexts.

For people mentioned repeatedly, require the company or unit, complete official position, and full name only at a leader's first introduction. For an industry leader at general-manager-assistant level or above—`总经理助理、副总经理、总经理、董事长`—do not require department information; use `产业＋完整正式职务＋姓名`. Allow later references to use an approved concise title without repeating the department. Write a department manager assistant as `xx部经理助理＋姓名`, never `xx部部门经理助理＋姓名`; delete the redundant `部门`. After `xx部经理助理路人甲`, use `路经理`, never `路助理`. After `总经理助理路人甲`, use `路总`, never `路助理`.

For a person without a management position, require only the department and full name at first mention, then allow the name alone. Do not add an unnecessary detailed `专员` label: change `xx部xx专员路人甲` to `xx部路人甲`. Exclude the four end-credit categories and their names (`文字/撰稿/供稿`, `摄影/摄像/配图/图片`, `编辑/排版`, `审核/审校`) from all people-title checks.

Treat ordinary lifestyle interactions and personal rapport with customers as customer-relationship maintenance, not as privacy disclosure. Allow such content when it does not expose phone numbers, email addresses, addresses, identity credentials, account information, health information, or other directly identifying sensitive data. Do not trigger an internal-warning issue merely because the word `客户` appears.

Allow subtitle text inside the opening column tag for all five columns: `星动态、星标杆、星故事、星分享、星视频`. Treat forms such as `【星标杆｜先进工作者】`, `【星视频｜主题短片】`, and the corresponding base tags as equally valid. Do not require a fixed subtitle.

For `星动态`, determine whether the article is a meeting article from the title's core event. Only a meeting article is checked for meeting name, host, and participant scope in its first substantive paragraph. Leadership visits, inspections, guidance, research visits, and similar event articles are checked against their actual event elements and must not be forced into the meeting template.

Do not claim or perform visual and layout review from `article_text`. Exclude image authenticity, image composition, Xiumi typography/layout, PPT screens, video frames, and other visual-format checks from this Skill's automatic review scope.

## Function-calling schemas

Read `references/function_calling.json` when exposing the functions to a platform. Do not emit or require a YAML configuration file.

## Validation

For Function 1, confirm the returned bytes begin with the ZIP/OOXML signature, can be opened by `python-docx`, contain all seven modules, preserve every non-empty source line, use A4 geometry with 2.54 cm margins, and distinguish the three bracketed rule labels.

For Function 2, confirm the return value contains exactly `summary` and `results`, performs no document or spreadsheet rendering, and includes one structured result per definite issue. Each result must retain problem evidence, standard basis, and modification tips. Do not include any uncertain-review field in `summary` or `results`.
