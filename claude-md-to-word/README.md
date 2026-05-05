# Claude MD → Word (去AI化 + 导出)

一个 Claude Code skill，把 Markdown 文件自动**去AI味**再导出为 **Word 文档 (.docx)**。

## 工作流

```
.md 文件 → 去AI化改写 → 生成 .docx
           ↑                    ↑
        de-ai 算法         md-to-word.js
```

## 安装

```bash
# 1. 进入你的 Claude Code 项目
cd your-project

# 2. 复制整个文件夹内容到项目根目录
#    (或者直接在此目录安装依赖)
cd claude-md-to-word
npm install
```

## 在 Claude Code 中使用

```
/md-to-word path/to/file.md
```

或在对话中说：**"把这个md转成word"**。

## 手动使用（不通过 Claude）

```bash
node md-to-word.js input.md output.docx
```

## 依赖

- Node.js 18+
- [docx](https://github.com/dolanmedia/docx) — 生成 Word 文档
