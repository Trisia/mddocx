# mddocx — OpenCode 安装

## 安装

1. 在 OpenCode 配置 (`opencode.json`) 中添加插件：

```json
{
  "plugins": [
    "git:https://github.com/Trisia/mddocx"
  ]
}
```

2. 重启 OpenCode，插件自动加载。

3. 确保已安装 Python 依赖：

```bash
pip install python-docx Pillow requests mistune
```

## 使用

在 OpenCode 对话中输入：

```
/mddoc paper.md
```

或直接粘贴 Markdown 内容要求转换为 Word 文档。
