MD to Word
将 Markdown 文件转换为 Word 文档 (.docx) 的 Node.js 工具集，专为 Claude Code 设计。

功能
md-to-word.js — 将 Markdown 文件转换为带格式的 Word 文档，支持标题、列表、代码块、引用、粗体/斜体等 Markdown 语法
generate-word.js — 用 docx 库直接生成 Word 文档（编程式创建，适合AI生成内容）
claude-md-to-word/ — 结合"去AI化"处理的工作流，先改写文风再导出 Word
快速开始
# 安装依赖
npm install

# Markdown 转 Word
node md-to-word.js 输入文件.md 输出文件.docx

# 直接生成 Word 文档（示例为新闻速览）
node generate-word.js
在 Claude Code 中使用
本工具集可作为 Claude Code 的 skill 使用：

/md-to-word path/to/file.md
或在对话中说"把这个md转成word"。

依赖
Node.js 18+
docx — 生成 Word 文档
协议
MIT
