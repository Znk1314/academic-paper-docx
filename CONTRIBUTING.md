# 贡献指南

感谢愿意帮忙。这个项目的核心价值是「**格式可验证**」，所以贡献方式也围绕这一点。

## 环境

```bash
python -m pip install -r requirements-dev.txt   # 运行时依赖 + pytest + ruff
python -m ruff check .
python -m pytest -q
python scripts/apd.py doctor                    # 环境自检
python scripts/apd.py demo -o build/demo        # 端到端跑一遍
```

Windows + 安装了 Microsoft Word 才能验证「更新目录 / 导出 PDF / 页面视图」这几项；
没有 Word 时请至少跑 `bash scripts/postprocess.sh` 或说明你验证过的环境。

## 两条硬规则

1. **格式参数只能放在 `SPEC` 里**（`scripts/docx_builder.py` 顶部）。
   不要在某个函数里写死 `Pt(16)` 之类的魔数——换主题、写测试、做质检都依赖 `SPEC`。
2. **改格式规则必须同时改质检和测试**：
   `scripts/qa_check.py` 加断言 → `tests/` 加用例 → `CHANGELOG.md` 记一笔。
   只改生成代码、不改质检，等于让规范悄悄失效。

## 目录职责

| 路径 | 职责 | 改动注意 |
|---|---|---|
| `scripts/docx_builder.py` | 格式库（页面/样式/元素） | 新参数进 `SPEC`；OOXML 子元素顺序用 `insert_ordered` |
| `scripts/omml_latex.py` | LaTeX 子集 → OMML | 新命令要加测试；解析失败不能抛异常 |
| `scripts/qa_check.py` | 规范断言 | 阈值来自 `EXPECT`；支持 `--spec` 校准 |
| `scripts/make_chart.py` | 图表 PNG | 新图型要复用 `PALETTE`，字号按 300 dpi 反算 |
| `scripts/postprocess.*` | Word/LibreOffice 后处理 | `.ps1` 必须 UTF-8 **with BOM** |
| `references/format-spec.md` | 规范原文 | 一般不改；确需改请单独说明理由 |
| `examples/` | 样例 | 不得包含真实个人信息与真实实验数据 |

## 新增一种内容块

1. 在 `PaperBuilder.render()` 里加分支（或新增 `add_xxx` 方法）。
2. 在 `tests/test_docx_builder.py` 加断言（结构 + 关键属性）。
3. 在 `qa_check.py` 增加对应检查项（若新块有格式约束）。
4. 更新 `SKILL.md` 的块类型表、`examples/minimal/content.json` 与 `CHANGELOG.md`。

## 提交信息

用 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)：

```
feat(builder): 支持图组与子图编号
fix(omml): 修复 \left[ ... \right] 缺少右定界符
test(qa): 补充表头重复的断言
docs: 补充 LibreOffice 的目录更新说明
chore(ci): 增加 Python 3.13
```

破坏性变更在正文里写 `BREAKING CHANGE:` 并说明影响。

## PR 检查清单

- [ ] `ruff check .` 与 `pytest -q` 全绿
- [ ] `apd.py build` + `apd.py qa` 对 `examples/demo` 与 `examples/minimal` 都通过
- [ ] 版式变化附修改前后的 PDF 页面截图
- [ ] 用户可见变化已写入 `CHANGELOG.md` 的 `[Unreleased]`
- [ ] 没有把示例内容（学校、姓名、数据、文献）带进代码或测试

## 发布（维护者）

见 [docs/publishing.md](docs/publishing.md)。
