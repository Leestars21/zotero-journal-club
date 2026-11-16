# Zotero Journal Club

中文 | [English](README.en.md)

项目主页：[Leestars21/zotero-journal-club](https://github.com/Leestars21/zotero-journal-club)

将 Zotero 文献或本地 PDF 整理为中文组会汇报，围绕论文基本信息、问题方法与证据、具体收获组织内容。优先使用现有 Zotero 接口、MinerU 结果和用户 PPT 模板，输出可编辑 PPTX、逐页备注、论文分析、来源清单和质量报告。

## 安装和依赖

可以直接让 Codex 安装：

```text
请从 https://github.com/Leestars21/zotero-journal-club 安装 zotero-journal-club Skill。
```

也可以下载仓库，将其中的 `zotero-journal-club/` 子目录复制到 `~/.codex/skills/zotero-journal-club/`（Windows：`%USERPROFILE%\.codex\skills\zotero-journal-club\`）。在新会话中使用 `$zotero-journal-club`。

需要 Codex 和可用的 PPT 制作工具；附带的生成与结构检查脚本需要 Python 3、`python-pptx` 和 Pillow，PDF 读取辅助脚本需要 PyMuPDF。从 Zotero 直接读取需要可用的 Zotero MCP 或本地只读接口；不可用时可提供 PDF。逐页渲染可使用 PowerPoint 或现有 artifact-tool 运行时；两者显示效果可能有差异。无需额外付费模型 API。

## 使用

```text
$zotero-journal-club 帮我汇报 Zotero 中的《论文标题》，制作 15 分钟中文组会 PPT。
$zotero-journal-club 根据这份本地 PDF 制作汇报，使用我提供的 PPTX 模板。
```

也可以提供 DOI、Zotero 条目 key 或链接。默认模板可在任务中指定，也可在 `~/.codex/zotero-journal-club.json` 设置 `default_template`；配置格式见 [示例](zotero-journal-club/config.example.json)。设置环境变量 `ZOTERO_JOURNAL_CLUB_CONFIG` 可使用其他配置文件。任务中指定的模板优先；指定的文件不存在时会报错。报告内容由当前 Codex 会话分析与核验，脚本不会自动理解论文。模板生成辅助器可使用版式占位符；对无占位符的版式，需要先检查模板并指定内容区域。母版和主题保留，复杂模板需要人工调整和视觉检查。自动检查不能证明学术结论正确，也不能完整检测文字溢出。

## 验证与反馈

运行 `python -m unittest discover -s tests`，并使用 Codex 自带的 `quick_validate.py` 检查技能。遇到问题可在[项目议题](https://github.com/Leestars21/zotero-journal-club/issues)反馈输入类型、操作系统、错误与预期，勿公开私人笔记、完整论文或 Zotero 数据库。

## 许可证

见 [LICENSE](LICENSE)。
