---
name: generate-word
description: 生成 Word 文档 (.docx)。当用户提及写word文档、生成word、导出word、创建word、生成docx时调用。
user-invocable: true
allowed-tools: Read, Write, Edit, Bash
---

# Word 文档生成器

当用户说以下任意内容时启动：

- "写个word文档" / "写一个word" / "帮我写word"
- "生成word" / "生成docx" / "导出word"
- "创建word文档"
- "/generate-word"

## 工作流程

### Step 1：确认内容

如果用户没有说明文档内容，询问：

```
要写什么内容？比如：
  - 文档标题？
  - 需要哪些章节/段落？
  - 要不要表格？
  - 有没有特定的格式要求？
```

如果用户直接描述了内容（如"写一份项目周报"），就根据描述直接修改脚本。

### Step 2：修改 generate-word.js

根据用户需求，编辑项目根目录下的 [generate-word.js](generate-word.js)，使用 `Edit` 工具修改 `generateDoc()` 函数中的内容。

该脚本使用 `docx` npm 库，支持：
- 标题（Title / Heading 1-5）
- 正文段落（可设字体、字号、颜色、粗体、斜体、下划线）
- 项目符号列表
- 表格（带表头样式）
- 日期落款

直接参考现有模板的写法来改。

### Step 3：生成文档

```bash
node generate-word.js
```

文件名必须用文档标题命名（如标题为"周报"，则输出"周报.docx"），不要用 `output.docx`。在脚本中定义 `const title = "文档标题"` 和 `const filename = \`${title}.docx\`` 来控制。

### Step 4：确认

告知用户文档已生成，路径是 `{标题}.docx`。

## 注意事项

- 务必将 `"docx"` 导入保留在文件顶部
- 生成函数必须是 `async function generateDoc()` 并以 `generateDoc().catch(console.error)` 结尾
- 输出路径用 `writeFileSync` 写入
