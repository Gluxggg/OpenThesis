<p align="center">
  <img src="assets/logo.png" alt="OpenThesis 标志" width="180" />
</p>

<h1 align="center">OpenThesis</h1>

<p align="center">
  <strong>AI 驱动的文档模板引擎</strong>
</p>

<p align="center">
  <a href="README.md"><img src="https://img.shields.io/badge/Language-English-2563eb" alt="English"></a>
  <a href="README.zh-CN.md"><img src="https://img.shields.io/badge/语言-简体中文-dc2626" alt="简体中文"></a>
</p>

<p align="center">
  <a href="https://github.com/Gluxggg/OpenThesis/actions/workflows/ci.yml"><img src="https://github.com/Gluxggg/OpenThesis/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="许可证"></a>
  <img src="https://img.shields.io/badge/status-V2_complete-brightgreen" alt="状态">
  <img src="https://img.shields.io/badge/node-%3E%3D22-success" alt="Node">
  <img src="https://img.shields.io/badge/pnpm-11.x-orange" alt="pnpm">
</p>

---

**一个引擎，无限模板。** 解析 `.docx` 模板，或从文本型 `.pdf` 参照文件中推断可复用样式；使用 Markdown 或结构化 JSON 写作，导出可直接提交且可继续编辑的 DOCX。

适用于**高校学位论文**、**期刊论文**和**党政机关公文**，可处理字体、页边距、页眉页脚、页码、表格和公式等排版要求。

---

<p align="center">
  <img src="assets/how-it-works.png" alt="OpenThesis 工作流程图：DOCX 或文本型 PDF 模板参照与 Markdown 或 JSON 内容交由解析器处理，再由 DOCX 渲染器生成可直接提交的文档" width="100%" />
</p>

---

## 为什么选择 OpenThesis？

| 传统方案 | OpenThesis |
| --- | --- |
| 只向模板中填充数据 | 理解标题、摘要、发文字号等语义角色 |
| 针对单个模板硬编码 | 解析 DOCX 或推断文本型 PDF 布局并生成 JSON 样式 DSL |
| 格式参数散落在代码中 | 所有格式由解析后的模板驱动 |
| Markdown 通常导出 LaTeX | Markdown 或 JSON 直接生成可编辑 DOCX |
| 仅支持单一文档类型 | 一个引擎覆盖论文、期刊和公文 |

## 工作流程

```mermaid
graph TD
    A["DOCX 模板或文本型 PDF 参照"] --> C["模板解析器"]
    B["Markdown 或 JSON 内容"] --> D["内容解析与校验"]
    C --> E["DOCX 渲染器"]
    D --> E
    E --> F["可提交的 DOCX 文档"]
```

## PDF 模板支持

许多高校、期刊和政府部门只发布 PDF 格式的排版规范或示例文件。OpenThesis 现在可以直接解析**文本型 PDF**，推断出驱动 DOCX 输出的 OpenThesis 模板。

| 输入 | 当前支持情况 | 说明 |
| --- | --- | --- |
| `.docx` 模板 | ✅ 原生支持 | 可提取命名样式、继承关系、页面尺寸、页眉页脚和分栏设置 |
| 文本型 `.pdf` | ✅ 启发式推断 | 推断页面尺寸、页边距、分栏、字体、字号、字重、对齐、行距、语义角色、奇偶页眉与页码，以及表格线 |
| 扫描型 `.pdf` | ✅ 搭配 OCRmyPDF | PDF 没有文本层时 `parse` 会调用本机 OCRmyPDF；没有装则给出明确的安装提示 |
| Markdown / OpenThesis JSON | ✅ 原生支持 | 作为文档内容输入，并渲染为可编辑 DOCX |

```bash
node packages/cli/dist/index.js parse 格式说明.pdf \
  --type thesis --org "XX大学"
# → 格式说明.template.json
```

PDF 保存的是定位后的字符，而不是 Word 的命名样式。OpenThesis 会把文字组合成行，从页面字体对象里读出字体的 PostScript 名称（`SimSun`、`SimHei`、`TimesNewRomanPS-BoldMT`），再映射成写作规范里叫的那个字体族——宋体、黑体、楷体、仿宋——并保留字体本身声明的加粗与斜体。页边距按正文栏测量，不把页眉或段落首行缩进算进去；带编号的行只有比正文更大时才被当作标题，所以目录不会被识别成章节。页面上的线条也会被读出来：三线表或全框线表格会写进 `tableFormat`，生成的文档沿用模板的线宽和表头底纹。生成的 JSON 仍会附带警告，因为文件里没有记录的信息无法凭空恢复。

正式使用前应检查生成的模板。绘图对象、脚注和无文本层的扫描页仍不在本次 PDF 支持范围内；奇偶页页眉、页脚页码和表格线会从页面页边带与路径里读出来。扫描件在装了 OCRmyPDF 时会自动 OCR，否则 `parse` 会提示需要安装什么。

## 快速开始

```bash
# 1. 安装并构建
git clone https://github.com/Gluxggg/OpenThesis.git
cd OpenThesis
pnpm install && pnpm build

# 2. 直接构建仓库内的 Markdown 示例
node packages/cli/dist/index.js build examples/sample-thesis.md -o output.docx

# 3. 解析 DOCX 模板或文本型 PDF 参照
node packages/cli/dist/index.js parse 你的学校模板.docx --type thesis --org "XX大学"
node packages/cli/dist/index.js parse 期刊格式说明.pdf --type journal --org "期刊名称"

# 4. 创建结构化内容示例
node packages/cli/dist/index.js init --type thesis

# 5. 使用解析后的模板生成文档
node packages/cli/dist/index.js build thesis-content.json -t template.json -o output.docx
```

Markdown 也可以先导出为便于检查的 JSON：

```bash
node packages/cli/dist/index.js import manuscript.md --type thesis -o manuscript.json
```

## Agent Skill

OpenThesis 在 `.codex/skills/openthesis` 中提供标准仓库 Skill。Agent 可以通过 `$openthesis` 显式调用，也可以在论文、期刊、公文、Word/PDF 模板和 Markdown 转 DOCX 等任务中自动发现它。

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs build manuscript.md \
  --type thesis -t university.template.json -o final.docx
```

## 支持的文档类型

| `--type` | 使用场景 | 主要能力 |
| --- | --- | --- |
| `thesis` | 学士、硕士、博士论文 | 封面、章节、图表、公式和参考文献 |
| `journal` | 学术期刊投稿 | 作者、单位、摘要、关键词、章节和参考文献 |
| `official` | 党政机关公文（GB/T 9704） | 红头、发文字号、主送/抄送、附件和落款 |

## 核心包

| 包 | 功能 |
| --- | --- |
| `@openthesis/document-schema` | 论文、期刊和公文的领域模型 |
| `@openthesis/markdown-parser` | Markdown 与 front matter 转结构化文档 JSON |
| `@openthesis/template-parser` | 解析 DOCX 或推断文本型 PDF 布局，生成 JSON 样式 DSL |
| `@openthesis/docx-renderer` | 模板驱动的 DOCX 渲染器 |
| `@openthesis/equation-engine` | LaTeX 数学 AST、原生 OMML 与兼容回退 |
| `@openthesis/cli` | `parse`、`import`、`init` 和 `build` 命令行工具 |

## 主要能力

- ✅ 解析并递归合并 DOCX 样式继承关系
- ✅ 将文本型 PDF 的字体族、字重、正文字栏、页眉页码和表格线推断为可复用模板（启发式，见上方 PDF 支持说明）
- ✅ Markdown 行内格式（加粗、斜体、行内代码、删除线、超链接）保留为 Word 富文本 run 与超链接关系，支持嵌套强调和表格单元格内格式
- ✅ Markdown 脚注（`[^1]` 引用与定义）输出为真正的 Word 脚注
- ✅ 扫描件 PDF 在装有 OCRmyPDF 时由 `parse` 自动 OCR
- ✅ 根据样式名称和大纲级别识别标题、摘要等语义角色
- ✅ 提取页面尺寸、页边距和分栏设置
- ✅ 表格边框、底纹和列宽控制
- ✅ 中文与西文字体分别配置
- ✅ 页眉、页脚和页码
- ✅ 图片嵌入、尺寸处理和题注
- ✅ 列表、代码块和引用块
- ✅ Markdown 导入与直接构建 DOCX
- ✅ 标准 Agent Skill
- ✅ 分数、根式、上下标、求和、积分、符号和希腊字母的原生 Office Math
- ✅ 函数、极限、重音、可缩放括号、二项式、矩阵、分段函数和对齐公式
- 🚧 后续 V3：用户宏、数组列规格、公式数组和较少使用的 AMS 环境

## 技术栈

- TypeScript 严格模式与 ESM
- pnpm workspace 单体仓库
- dolanmiu/docx
- JSZip 与 fast-xml-parser
- PDF.js
- Node.js ≥ 22

## 路线图

| 阶段 | 目标 | 状态 |
| --- | --- | --- |
| V1 | 核心解析、渲染与 CLI | ✅ 完成 |
| V2 | 图片、双栏输出与 Markdown 解析器 | ✅ 完成 |
| V3 | 原生 LaTeX → OMML | 🚧 进行中（核心与常用高级子集） |
| V4 | 高级 PDF 重建：OCR、表格、页眉页脚和机器学习语义识别 | 📋 未来计划 |
| V5 | 自动生成论文内容的 AI Agent | 📋 未来计划 |
| SaaS | 社区贡献的模板市场 | 📋 未来计划 |

## Star History

<p align="center">
  <a href="https://www.star-history.com/#Gluxggg/OpenThesis&Date">
    <img src="https://api.star-history.com/svg?repos=Gluxggg/OpenThesis&type=Date" alt="OpenThesis Star 历史趋势图" width="700">
  </a>
</p>

## 参与贡献

如果你希望 OpenThesis 支持某个高校或期刊的 `.docx` 或文本型 `.pdf` 模板，欢迎提交 issue 并附上模板或下载链接。

欢迎提交 PR。开发流程参见 [CONTRIBUTING.md](CONTRIBUTING.md)，架构说明参见 [CLAUDE.md](CLAUDE.md)。

## 开发检查

```bash
pnpm check
pnpm audit --prod
```

## 许可证

MIT © 2026 OpenThesis
