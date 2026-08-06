# SAB-news-audit-skill
# SAB News Audit Skill

面向 SAB“星系列”新闻稿的 Codex 审核 Skill。

本 Skill 内置《SAB星系列新闻稿统一标准手册》，可根据不同栏目规范审核新闻稿，只输出能够从稿件文字中明确判断的问题、标准依据和修改建议。

## 支持栏目

- 星动态
- 星故事
- 星分享
- 星标杆
- 星视频

## 主要功能

### 1. 生成标准手册

函数：

```python
generate_standard_manual()
```

功能：

生成完整《SAB星系列新闻稿统一标准手册》DOCX 二进制内容。

该函数：

- 无入参；
- 使用 `python-docx` 生成文档；
- 不读取或写入本地文件；
- 通过二进制流返回 DOCX 内容。

### 2. 审核新闻稿

函数：

```python
review_news_article(article_text, column_type)
```

参数说明：

| 参数 | 类型 | 必填 | 说明 |
|---|---|---:|---|
| `article_text` | string | 是 | 新闻稿全文，第一行应为标题 |
| `column_type` | string | 是 | 栏目类型 |

`column_type` 可选值：

```text
星动态
星故事
星分享
星标杆
星视频
```

返回示例：

```json
{
  "summary": {
    "明确问题": 2,
    "不通过": 1,
    "需整改": 1,
    "通过": 8
  },
  "results": [
    {
      "序号": "G01",
      "规则类型": "硬性规则",
      "检查类别": "标题与栏目",
      "检查项目": "标题栏目标签与指定栏目一致",
      "结果": "不通过",
      "问题证据": "原文证据",
      "标准依据": "对应手册模块",
      "修改Tips": "具体修改建议"
    }
  ]
}
```

新闻稿审核不会生成 Word、Excel、PDF或其他附件，只返回：

- `summary`：审核结果汇总；
- `results`：明确存在的问题、证据、标准依据和修改建议。

## 审核原则

- 只指出能够从稿件文字中明确判断的问题。
- 不输出“需人工确认”清单。
- 不展示已经通过的具体检查项目。
- 不自行增加《SAB星系列新闻稿统一标准手册》以外的规范。
- 无法从文字直接判断的授权、审批、图片或排版问题不作为问题输出。
- 不生成 Word、Excel 或其他审核附件。

## 已适配规则

### 文末责任信息

文末责任信息按含义匹配，不限制冒号、空格、斜杠或换行。

允许以下对应关系：

- `文字、撰稿、供稿`：文字责任；
- `摄影、摄像、配图、图片`：摄影责任；
- `编辑、排版`：编辑责任；
- `审核、审校`：审核责任。

### 内部提示语

内部提示语不要求固定句式。

只要同时表达“内部使用”和“禁止外传”的意思即可，例如：

```text
内部通讯 请勿外传
```

或者：

```text
内部通讯稿，禁止对外转发
```

### “集团”与“总公司”

只有在介绍公司领导身份时，才要求使用“总公司”而不能使用“集团”。

例如：

```text
集团副总经理路人甲
```

应改为：

```text
总公司副总经理路人甲
```

创业周年、发展历史、企业文化等非领导身份介绍场景允许使用“集团”。

### 人物称谓

- 领导首次出现时，应写明公司或单位、完整正式职务和姓名。
- 其他人物首次出现时，应写明职务或部门和姓名。
- 后续可以使用姓名、“×总”“×经理”等简称。
- 不得将后续简称判定为职务或部门信息缺失。
- 文末文字、摄影、编辑、审核等责任人不属于人物首次介绍检查范围。

### 星动态会议稿判断

先根据标题核心事件判断稿件是否属于会议稿。

只有会议稿才检查：

- 时间；
- 地点；
- 会议名称；
- 主持人；
- 必要参会范围。

领导莅临、检查、指导、调研、考察等稿件不属于会议稿，不要求补充会议名称、主持人和参会范围。

## 项目结构

```text
sab-news-audit-skill/
├── SKILL.md
├── scripts/
│   └── sab_news_processor.py
└── references/
    └── function_calling.json
```

安装后必须确保：

- `SKILL.md` 位于 Skill 根目录；
- `scripts` 和 `references` 与 `SKILL.md` 同级；
- 不要形成两层重复的 `sab-news-audit-skill` 文件夹。

## 安装方法

### 方法一：下载 ZIP

1. 打开本 GitHub 仓库。
2. 点击右上角的 `Code`。
3. 选择 `Download ZIP`。
4. 解压下载文件。
5. 将文件夹重命名为：

```text
sab-news-audit-skill
```

6. 将整个文件夹复制到 Codex Skills 目录。

macOS 或 Linux：

```text
~/.codex/skills/sab-news-audit-skill/
```

Windows：

```text
%USERPROFILE%\.codex\skills\sab-news-audit-skill\
```

正确的安装结构：

```text
~/.codex/skills/sab-news-audit-skill/
├── SKILL.md
├── scripts/
│   └── sab_news_processor.py
└── references/
    └── function_calling.json
```

### 方法二：使用 Git

将下面的仓库地址替换为实际 GitHub 地址：

```bash
git clone https://github.com/你的用户名/你的仓库名.git ~/.codex/skills/sab-news-audit-skill
```

如果是私有仓库，需要先配置 GitHub 登录或访问权限。

## 使用方法

安装完成后，建议新建一个 Codex 任务，再输入：

```text
使用 $sab-news-audit-skill 审核一篇【星动态】新闻稿。
```

然后粘贴新闻稿全文。

完整示例：

```text
使用 $sab-news-audit-skill 审核以下【星动态】新闻稿，只指出明确存在的问题和修改建议：

【星动态】新闻稿标题

新闻稿正文……

供稿：张三
配图：李四
排版：王五
审核：赵六
```

审核其他栏目：

```text
使用 $sab-news-audit-skill 审核以下【星标杆】稿件。
```

生成标准手册：

```text
使用 $sab-news-audit-skill 生成《SAB星系列新闻稿统一标准手册》。
```

## Function Calling

平台函数配置文件位于：

```text
references/function_calling.json
```

提供两个函数：

```text
generate_standard_manual
review_news_article
```

配置使用纯 JSON，不需要 PyYAML，也不提供 YAML 配置文件。

## 运行环境

使用的第三方库：

```text
python-docx
openpyxl
```

其他部分仅使用 Python 标准库。

代码要求：

- 不导入 PyYAML；
- 不加载或序列化 YAML；
- 不依赖外部模板文件；
- 规则和手册内容内置；
- 函数执行过程中不进行本地文件读写。

## 更新方法

更新 Skill 时：

1. 替换 `SKILL.md`。
2. 替换 `scripts/sab_news_processor.py`。
3. 如函数配置发生变化，替换 `references/function_calling.json`。
4. 保持文件名和目录结构不变。
5. 更新完成后新建 Codex 任务，确保加载最新版本。

## 注意事项

- 本 Skill 用于新闻稿文字审核辅助，不能代替组织内部正式审批。
- 不要将未经授权的内部稿件提交到公开环境。
- 发布新闻稿前，应按照内部流程完成最终审核。
- 如果本仓库包含内部规范或非公开内容，请将 GitHub 仓库设置为 `Private`。

## License

本项目的使用和分发应遵守组织内部授权要求。

如未获得公开发布授权，请勿使用开放源代码许可证向外部授予复制、修改或再分发权限。
