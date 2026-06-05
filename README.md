# Office MCP Server

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10+-green.svg)
[![MCP](https://img.shields.io/badge/MCP-1.0.0-orange.svg)](https://modelcontextprotocol.io)
[![Token Savings](https://img.shields.io/badge/Token%20Savings-99%25-brightgreen.svg)](https://github.com/chunleiding/office-mcp-server)

通用 Office 文档 MCP Server —— 基于 AI 助手**创建、读取、克隆 Word/Excel/PPT 文档**，格式100%保留。

> 🎯 **核心能力：通用模板克隆** —— 扔进任意 `.docx` 模板，AI 自动识别结构、替换内容，**所有格式原样保留**。适用于会议纪要、教案、周报、报告等一切重复性文档场景。

> 💰 **极致 Token 效率** —— 本地分析文档结构，只返回摘要给 AI，**节省 99% Token 消耗**，大幅降低 API 成本。

---

## ✨ 特性

### 💰 极致 Token 效率
- ✅ **本地分析** —— 文档结构在本地解析，只返回摘要给 AI（500 tokens vs 50,000+ tokens）
- ✅ **本地替换** —— 文本替换在本地执行，不消耗 AI Token
- ✅ **节省 99%** —— 相比传统方案（AI 读取整个 XML），Token 消耗降低 99%
- ✅ **批量友好** —— 处理 100 个文档仅需 70K tokens（传统方案需要 6.5M tokens）

### 📝 Word 文档（已实现）
- ✅ **通用分析** —— `word_analyze` 自动识别任意 Word 文档的段落+表格结构
- ✅ **通用替换** —— `word_replace` 接受替换字典，3 层寻址覆盖所有文档类型
- ✅ **格式 100% 保留** —— 只替换 `<w:t>` 文本，不碰 `<w:p>` 段落元素
- ✅ **Run 级格式** —— 支持下划线、加粗等局部格式保留（如标题"中一"下划线）
- ✅ **单元格批量替换** —— 用 `\n` 分隔多行，自动映射到单元格段落
- ✅ **兼容旧 API** —— `word_clone_template` 仍可使用

### 📊 Excel 文档（计划中）
- 🔜 读取单元格、公式
- 🔜 克隆 Excel 模板

### 📽️ PowerPoint（计划中）
- 🔜 创建演示文稿
- 🔜 克隆 PPT 模板

---

## 🚀 快速开始

### 安装

```bash
git clone https://github.com/chunleiding/office-mcp-server.git
cd office-mcp-server
uv sync
```

### 配置 MCP

在 `~/.workbuddy/mcp.json` 或 Claude Desktop 配置中添加：

```json
{
  "mcpServers": {
    "office-mcp-server": {
      "command": "uv",
      "args": [
        "--directory", "/path/to/office-mcp-server",
        "run", "python", "-m", "office_mcp.server"
      ]
    }
  }
}
```

---

## 🛠️ 工具参考

### 核心工具

#### `word_analyze` —— 分析文档结构

输入任意 Word 文档，返回文本索引（段落 + 单元格），供 AI 选择替换目标。

```json
{
  "path": "/path/to/template.docx"
}
```

返回示例：
```json
{
  "paras": [
    {"i": 0, "t": "健康活动：小毛毛虫去旅行（屈身爬）", "s": "标题 1"},
    {"i": 2, "t": "1.初步掌握身体拱起...", "s": "Normal"}
  ],
  "tables": [
    {
      "i": 0, "rows": 4, "cols": 10,
      "cells": [
        {"r": 0, "c": 1, "txt": "2026年5月28日", "pn": 1},
        {"r": 3, "c": 0, "txt": "一、本周回顾...", "pn": 37}
      ],
      "merge": [{"r": 3, "c": 0, "rs": 1, "cs": 10}]
    }
  ]
}
```

#### `word_replace` —— 通用文本替换

接受替换字典，支持 3 层寻址：

| 寻址格式 | 含义 | 示例 |
|----------|------|------|
| `p:{i}` | 正文段落 | `"p:0": "新标题"` |
| `c:{t}:{r}:{c}` | 整个单元格（自动按行映射） | `"c:0:3:0": "行1\n行2\n行3"` |
| `c:{t}:{r}:{c}:{p}` | 单元格内特定段落 | `"c:0:2:1:0": "新主题"` |

```json
{
  "path": "/path/to/template.docx",
  "output_path": "/path/to/output.docx",
  "replacements": "{\"p:0\": \"新标题\", \"c:0:3:0\": \"一、背景\\n二、分析\"}",
  "fmt_hints": "{\"_fmt:p:1\": {\"runs\": [{\"text\": \"中一\", \"underline\": true}]}}"
}
```

**`fmt_hints` 格式提示**（可选）：用于保留 Run 级格式（如下划线），AI 只在需要时才使用。

---

### 兼容工具

#### `word_clone_template` —— 表格式模板克隆（旧 API）

保留兼容，适用于"教研记录"类表格式模板：

```json
{
  "template_path": "/path/to/template.docx",
  "output_path": "/path/to/output.docx",
  "title": "中一班教研记录",
  "date": "2026年6月5日 星期五 下午",
  "main_topic": "幼儿吃饭不积极问题研讨",
  "process_record": "一、研讨背景\n..."
}
```

长文本可使用 `process_record_file` 参数指定文件路径。

#### `word_read_document` —— 读取文档内容

```json
{
  "document_path": "/path/to/document.docx"
}
```

---

## 💰 Token 效率对比

### 痛点：传统方案的成本黑洞

传统方式让 AI 直接操作 Word 文档，需要：
1. 读取整个 docx（XML 格式）→ ~50,000 tokens
2. AI 理解文档结构 → ~10,000 tokens
3. 生成修改方案 → ~5,000 tokens
4. **总计：~65,000 tokens/次**

按 GPT-4 价格（$3/1M input, $15/1M output）计算：
- 单次成本：$0.195
- 批量处理 100 个文档：$19.5

### 解决方案：office-mcp-server

使用 MCP 协议，文档处理在本地完成：
1. `word_analyze` 本地分析，只返回摘要 → ~500 tokens
2. `word_replace` 本地执行替换，不消耗 Token → ~200 tokens
3. **总计：~700 tokens/次**

- 单次成本：$0.0021
- 批量处理 100 个文档：$0.21
- **节省：98.9% 成本！**

### 详细对比表

| 操作 | 传统方式 Token | MCP 方式 Token | 节省比例 | 成本节省 |
|------|---------------|---------------|----------|----------|
| 读取文档 | 50,000 | 500 | 99.0% | $0.149 → $0.0015 |
| 替换变量 | 65,000 | 700 | 98.9% | $0.195 → $0.0021 |
| 批量处理 100 个 | 6,500,000 | 70,000 | 98.9% | $19.5 → $0.21 |
| 每日 1000 次操作 | 65,000,000 | 700,000 | 98.9% | $195 → $2.1 |

### 为什么能这么高效？

**MCP 设计原则：数据本地处理，上送 LLM 只摘要**

1. **本地分析**：`word_analyze` 在本地解析 docx XML，提取结构化摘要
2. **摘要上送**：只返回段落索引 + 前 50 字符，不发送原始 XML
3. **本地执行**：`word_replace` 在本地执行替换，AI 只发送替换字典
4. **零冗余**：AI 不需要读取/生成整个文档，只处理差异部分

---

## 🏗️ 架构

```
office-mcp-server/
├── src/office_mcp/
│   ├── server.py              # FastMCP 入口（2核心 + 2兼容工具）
│   └── word/
│       ├── engine.py          # 通用 Word 文档引擎（核心）
│       └── table_parser.py    # 表格格式解析
└── tests/
```

### 核心设计

**通用引擎 `engine.py`**：
- `analyze()` —— 提取文档文本索引
- `replace_text()` —— 通用替换，3 层寻址
- `clone_word_template()` —— 旧 API 兼容层

**格式保留原理**：
1. 只替换 `<w:t>` XML 文本节点
2. 不增删任何 `<w:p>` 段落元素
3. 段落间距由模板 `<w:pPr>` 控制，不受内容空行干扰
4. 单元格替换自动按行映射，超出部分追加新段落并复制末段 pPr

---

## 🗺️ 路线图

| 阶段 | 功能 | 状态 |
|------|------|------|
| 1 | Word 通用引擎 + 模板克隆 | ✅ 已完成 |
| 2 | Excel 支持 | 🔜 计划中 |
| 3 | PPT 支持 | 🔜 计划中 |
| 4 | 跨格式批量操作 | 💡 未来 |

---

## 📋 依赖

- Python 3.10+
- [FastMCP](https://github.com/modelcontextprotocol/python-sdk) >= 2.0.0
- `python-docx` >= 1.2.0
- `lxml` >= 5.0.0
- `openpyxl` >= 3.1.0（Excel，计划中）
- `python-pptx` >= 1.0.0（PPT，计划中）

---

## 📄 许可证

MIT License — 详见 [LICENSE](LICENSE)。
