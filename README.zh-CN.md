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

**一个引擎，无限模板。** 解析任意 `.docx` 模板，使用 Markdown 或结构化 JSON 写作，导出可直接提交且可继续编辑的 DOCX。

适用于**高校学位论文**、**期刊论文**和**党政机关公文**，可处理字体、页边距、页眉页脚、页码、表格和公式等排版要求。

---

## 为什么选择 OpenThesis？

| 传统方案 | OpenThesis |
| --- | --- |
| 只向模板中填充数据 | 理解标题、摘要、发文字号等语义角色 |
| 针对单个模板硬编码 | 解析任意 `.docx` 模板并生成 JSON 样式 DSL |
| 格式参数散落在代码中 | 所有格式由解析后的模板驱动 |
| Markdown 通常导出 LaTeX | Markdown 或 JSON 直接生成可编辑 DOCX |
| 仅支持单一文档类型 | 一个引擎覆盖论文、期刊和公文 |

## 工作流程

```mermaid
graph TD
    A["Word 模板 .docx"] --> C["模板解析器"]
    B["Markdown 或 JSON 内容"] --> D["内容解析与校验"]
    C --> E["DOCX 渲染器"]
    D --> E
    E --> F["可提交的 DOCX 文档"]
```

## 快速开始

```bash
# 1. 安装并构建
git clone https://github.com/Gluxggg/OpenThesis.git
cd OpenThesis
pnpm install && pnpm build

# 2. 直接构建仓库内的 Markdown 示例
node packages/cli/dist/index.js build examples/sample-thesis.md -o output.docx

# 3. 解析学校或期刊的 Word 模板
node packages/cli/dist/index.js parse 你的学校模板.docx --type thesis --org "XX大学"

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

OpenThesis 在 `.codex/skills/openthesis` 中提供标准仓库 Skill。Agent 可以通过 `$openthesis` 显式调用，也可以在论文、期刊、公文、Word 模板和 Markdown 转 DOCX 等任务中自动发现它。

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
| `@openthesis/template-parser` | `.docx` 模板转带继承关系的 JSON 样式 DSL |
| `@openthesis/docx-renderer` | 模板驱动的 DOCX 渲染器 |
| `@openthesis/equation-engine` | LaTeX 数学 AST、原生 OMML 与兼容回退 |
| `@openthesis/cli` | `parse`、`import`、`init` 和 `build` 命令行工具 |

## 主要能力

- ✅ 解析并递归合并 DOCX 样式继承关系
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
- Node.js ≥ 22

## 路线图

| 阶段 | 目标 | 状态 |
| --- | --- | --- |
| V1 | 核心解析、渲染与 CLI | ✅ 完成 |
| V2 | 图片、双栏输出与 Markdown 解析器 | ✅ 完成 |
| V3 | 原生 LaTeX → OMML | 🚧 进行中（核心与常用高级子集） |
| V4 | 基于机器学习的模板布局理解 | 📋 未来计划 |
| V5 | 自动生成论文内容的 AI Agent | 📋 未来计划 |
| SaaS | 社区贡献的模板市场 | 📋 未来计划 |

## Star History

<p align="center">
  <a href="https://www.star-history.com/#Gluxggg/OpenThesis&Date">
    <img src="https://api.star-history.com/svg?repos=Gluxggg/OpenThesis&type=Date" alt="OpenThesis Star 历史趋势图" width="700">
  </a>
</p>

## 参与贡献

如果你希望 OpenThesis 支持某个高校或期刊的 `.docx` 模板，欢迎提交 issue 并附上模板或下载链接。

欢迎提交 PR。开发流程参见 [CONTRIBUTING.md](CONTRIBUTING.md)，架构说明参见 [CLAUDE.md](CLAUDE.md)。

## 开发检查

```bash
pnpm check
pnpm audit --prod
```

## 许可证

MIT © 2026 OpenThesis
