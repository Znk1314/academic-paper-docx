## 这个 PR 做了什么

<!-- 一句话说清。如果是格式变更，说明它属于规范内的实现修正还是规范外的扩展。 -->

## 相关 issue

<!-- Fixes #123 / Refs #456 -->

## 变更类型

- [ ] 修 bug（不改变对外行为，只修不对的地方）
- [ ] 新能力（内容模型 / 格式库 / 公式 DSL / 质检规则）
- [ ] 破坏性变更（默认格式或内容模型字段发生变化）
- [ ] 文档 / 示例 / CI

## 自测

- [ ] `python -m ruff check .` 通过
- [ ] `python -m pytest -q` 通过
- [ ] `python scripts/apd.py build examples/demo/content_model.json -o build/demo.docx` 正常
- [ ] `python scripts/apd.py qa build/demo.docx` 无 FAIL
- [ ] 涉及版式变化时，附上了修改前后的 PDF 页面截图
- [ ] 涉及 Word 保存/视图/域的行为时，说明在哪个环境验证过（Windows + Word / WPS / LibreOffice）

## 检查清单

- [ ] 新格式参数先写进 `scripts/docx_builder.py` 的 `SPEC`，没有散落在函数里
- [ ] 格式规则变化时同步更新了 `scripts/qa_check.py` 与 `tests/`
- [ ] 用户可见的变化已写入 `CHANGELOG.md` 的 `[Unreleased]`
- [ ] 没有把示例内容（学校名、姓名、数据、文献）带进代码或测试
