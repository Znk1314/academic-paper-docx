# 安全说明

## 这个项目会做什么

- 在**本地**读写文件：生成 `.docx`、导出 `.pdf`、写出预览 `.png`。
- 在 Windows 上通过 COM 调用本机已安装的 **Microsoft Word**（`postprocess.ps1`）。
- 在非 Windows 上调用本机的 **LibreOffice**（`postprocess.sh`）。
- **不联网**：没有 HTTP 请求、没有遥测、不下载字体或模板。

## 需要留意的点

| 场景 | 风险 | 建议 |
|---|---|---|
| 处理不受信任的 `.docx` | Word COM 打开恶意文档 | 只对你信任的文档跑 `postprocess` |
| `content.json` 里的图片路径 | 生成时会读取该路径 | 不要对来源不明的内容模型直接运行 |
| Word 宏 | 本项目生成的是 `.docx`，不含宏 | 不要把它另存为 `.docm` 再分发 |
| CI 中的 `git archive` / Action | 依赖第三方 Action | 已固定到主版本，可按需 pin 到 commit SHA |

## 报告漏洞

请使用 GitHub 的 [私密漏洞报告](https://github.com/Znk1314/academic-paper-docx/security/advisories/new)，
不要开公开 issue。请在报告里包含：影响版本（`scripts/_version.py`）、复现步骤、
以及你判断的影响面。一般会在 7 天内回复。

## 支持的版本

只有最新发布版本会得到安全修复；`main` 分支的问题会在下一个补丁版本中修复。
