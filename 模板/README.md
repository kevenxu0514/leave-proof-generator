# 模板文件说明

本仓库（GitHub 文本通道）**不直接存放**两份二进制 Word 模板（各约 90KB 的 `.docx`，无法经文本 API 保真传输）。模板原件由校站提供，存放位置：

- Skill 安装目录：`<skill>/模板/上课请假模板.docx`、`<skill>/模板/分寝室请假模板.docx`
- 开发/工作目录：`C:\Users\Keven\Desktop\校站\work\上课请假模板.docx`、`分寝室请假模板.docx`

`generate.py` 按 `模板/` 子目录（或同目录）查找模板，缺失时会明确报错。若需把模板纳入版本管理，请通过 GitHub 网页端「Add file」或桌面客户端直接上传这两个 `.docx` 到本目录。
