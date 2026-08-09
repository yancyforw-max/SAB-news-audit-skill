# SAB News Audit Skill

用于审核 SAB「星系列」新闻稿的 Codex Skill，支持以下栏目：

- 星动态
- 星故事
- 星分享
- 星标杆
- 星视频

## 主要功能

- 生成《SAB星系列新闻稿统一标准手册》DOCX。
- 按栏目审核新闻稿，仅输出明确存在的问题和修改建议。
- 支持五个栏目使用副标题，如 `【星标杆｜先进工作者】`。
- 检查标题、内部提示、责任信息、人物称谓、栏目结构等规则。
- 不生成审核 Word/Excel，不输出“需人工确认”项目。

> 当前只审核新闻稿纯文本，不覆盖图片、秀米、PPT、视频画面等视觉与排版内容。

## 安装

将 `sab-news-audit-skill` 文件夹复制到：

```text
~/.codex/skills/
```

重新打开 Codex 后即可使用。

## 使用示例

```text
使用 $sab-news-audit-skill 审核一篇【星动态】新闻稿：

【星动态｜质量提升】……
```

## 核心函数

```python
generate_standard_manual()
review_news_article(article_text, column_type)
```

`review_news_article()` 返回：

```json
{
  "summary": {},
  "results": []
}
```

## 运行依赖

- Python 3
- python-docx
- openpyxl

不依赖 PyYAML，配置采用 Python 内置字典及纯 JSON。
