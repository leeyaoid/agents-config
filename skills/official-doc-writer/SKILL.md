---
name: official-doc-writer
description: 生成、修改和排版符合GB/T 9704-2012要求的党政机关公文及国企正式材料，并输出真正使用Word样式的DOCX。适用于通知、报告、请示、函、通报、纪要、汇报稿和公文式提纲；生成或修改Word公文时必须同时使用minimax-docx完成结构校验和最终预览。
---

# 党政机关公文生成

使用 `scripts/generate_official_doc.py` 作为生成入口。默认以 `assets/official_document_normal_template.docx` 为基础模板；该模板基于 Word Normal 默认模板，直接修改 Word 内置样式，并为没有合适内置样式的公文要素创建无前缀语义样式。

## 强制协同

每次生成或修改 Word 公文时，同时使用 `minimax-docx`：

- 本 skill 负责内容结构、GB/T 9704 版式和 Word 样式；
- `minimax-docx` 负责结构校验、业务规则校验、运行合并和最终预览；
- 两套校验通过后才能交付。

## 样式硬要求

1. 字体、字号、颜色、粗体、缩进和行距必须定义在 `word/styles.xml` 中，不得逐段直接设置。
2. 每个非空段落必须引用 `w:pStyle`。
3. 正文、公文标题、一级至四级标题必须复用 Word 内置 `Normal`、`Title`、`Heading 1` 至 `Heading 4`，并将内置样式修改为公文要求。
4. 不得创建“公文-正文”“公文-标题”“公文-X级标题”等替代样式。
5. 主送机关、附件说明、署名、日期、版记和页码等无合适内置样式的要素，直接使用“主送机关”“附件说明”“版记”等无前缀语义样式。
6. 红色分隔线和版记线使用段落样式边框，不得用下划线或重复符号模拟。
7. 正文和三级、四级标题使用“仿宋”，二级标题和签发人姓名使用“楷体”；不得使用“仿宋_GB2312”或“楷体_GB2312”。
8. 中文单双引号统一为 `“”`、`‘’`，直接继承所在段落样式，不得为引号单独创建字符样式。

## 标准样式体系

| 内容 | Word 样式 |
|---|---|
| 正文 | `正文`（Word 内置 `Normal`） |
| 公文标题 | `标题`（Word 内置 `Title`） |
| 一级标题 | `标题 1`（Word 内置 `Heading 1`） |
| 二级标题 | `标题 2`（Word 内置 `Heading 2`） |
| 三级标题 | `标题 3`（Word 内置 `Heading 3`） |
| 四级标题 | `标题 4`（Word 内置 `Heading 4`） |
| 发文机关标志 | `发文机关标志` |
| 密级、紧急程度 | `密级与紧急程度` |
| 发文字号 | `发文字号` |
| 红色分隔线 | `红色分隔线` |
| 主送机关 | `主送机关` |
| 附件说明 | `附件说明` |
| 发文机关署名 | `发文机关署名` |
| 成文日期 | `成文日期` |
| 附注 | `附注` |
| 版记 | `版记`、`版记分隔线` |
| 页码 | `页码段落`、`页码字符` |

## 生成流程

### 1. 组织内容

`body` 支持字符串自动识别层级，也支持显式样式：

```python
body = [
    "正文自然段。",
    "一、一级标题",
    "（一）二级标题",
    "1. 三级标题",
    "（1）四级标题",
    {"text": "明确指定的正文。", "style": "正文"},
]
```

### 2. 生成 DOCX

```python
from scripts.generate_official_doc import generate_document

content = {
    "issuer": "XXX公司",
    "doc_number": "XX〔2026〕1号",
    "title": "关于有关情况的报告",
    "recipient": "上级单位",
    "body": ["现将有关情况报告如下。", "一、总体情况", "有关工作稳步推进。"],
    "issuer_signature": "XXX公司",
    "date": "2026年6月29日",
}
generate_document("报告", content, "output.docx")
```

命令行：

```bash
python3 scripts/generate_official_doc.py --type 报告 --input-json content.json --output output.docx
```

需要重建独立模板时运行：

```bash
python3 scripts/build_normal_template.py
```

不得修改用户系统中的全局 `Normal.dotm`。

### 3. 样式硬校验

```bash
python3 scripts/validate_style_usage.py output.docx
```

校验器检查必需样式、内置样式 ID、段落样式引用、直接字符格式和英文直引号。只有输出 `Style validation: PASSED` 才能继续。

### 4. minimax-docx 双重校验

```bash
bash MINIMAX_DOCX_SKILL_DIR/scripts/env_check.sh
dotnet run --project MINIMAX_DOCX_SKILL_DIR/scripts/dotnet/MiniMaxAIDocx.Cli -- merge-runs --input output.docx
dotnet run --project MINIMAX_DOCX_SKILL_DIR/scripts/dotnet/MiniMaxAIDocx.Cli -- validate --input output.docx --business
bash MINIMAX_DOCX_SKILL_DIR/scripts/docx_preview.sh output.docx
python3 scripts/validate_style_usage.py output.docx
unzip -t output.docx
```

最后转为 PDF 或图片检查全部页面，确认标题回行、分页、孤行、版记和页码位置。

## 公文格式参数

- 页面：A4；上 37mm、下 35mm、左 28mm、右 26mm；
- 标题：方正小标宋_GBK，二号；
- 正文：仿宋，三号，固定 28 磅行距；
- 一级标题：黑体，三号；
- 二级标题：楷体，三号；
- 三级标题：仿宋，三号加粗；
- 四级标题：仿宋，三号；
- 页码：宋体，四号。

格式细节不确定时读取 `references/GBT_9704-2012_党政机关公文格式.md`；字体说明见 `fonts/README.md`。
