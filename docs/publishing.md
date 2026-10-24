# 发布清单

仓库已经搭好，剩下的是把它推到你自己的 GitHub 账号上。以下步骤按顺序做一遍即可。

## 1. 替换 OWNER

README、CHANGELOG、SECURITY、`.github/ISSUE_TEMPLATE/config.yml` 里的 `OWNER` 需要换成你的
GitHub 用户名（或组织名）。二选一：

```bash
# 方式 A：全局替换（Linux / macOS / Git Bash）
grep -rl "OWNER" --include="*.md" --include="*.yml" --include="*.toml" . \
  | xargs sed -i 's|OWNER/academic-paper-docx|your-name/academic-paper-docx|g'
```

```powershell
# 方式 B：PowerShell
Get-ChildItem -Recurse -File -Include *.md,*.yml,*.toml |
  ForEach-Object {
    $c = [System.IO.File]::ReadAllText($_.FullName, [System.Text.UTF8Encoding]::new($false))
    if ($c -match 'OWNER') {
      [System.IO.File]::WriteAllText($_.FullName,
        ($c -replace 'OWNER/academic-paper-docx', 'your-name/academic-paper-docx'),
        [System.Text.UTF8Encoding]::new($false))
      "patched: $($_.Name)"
    }
  }
```

> 注意：`.ps1` 文件必须保持 **UTF-8 with BOM**，上面的写法不要用于 `*.ps1`。

## 2. 初始化并推送

```bash
git init -b main
git add .
git commit -m "chore: initial commit"
git remote add origin https://github.com/<your-name>/academic-paper-docx.git
git push -u origin main
```

## 3. 仓库设置

- **Description**：`学术论文/研究报告 Word 排版与生成：三线表、OMML 公式、自动目录、GB/T 7714，可作为 agent skill 使用`
- **Topics**：`word` `docx` `python-docx` `academic-writing` `thesis` `agent-skill` `omml` `chinese`
- **Features**：勾选 Issues、Discussions（`ISSUE_TEMPLATE/config.yml` 里有 Discussions 链接）
- **Pages**：Settings → Pages → Source 选 **GitHub Actions**（工作流已就绪，
  推送到 `main` 后 `docs/` 会自动部署）

## 4. 打第一个版本

```bash
git tag -a v1.0.0 -m "v1.0.0"
git push origin v1.0.0
```

`Release` 工作流会先校验标签与 `scripts/_version.py` 是否一致、跑测试，
然后打包 `academic-paper-docx-v1.0.0.zip`（skill 目录结构，解压即可用）并创建 Release。

## 5. 发布前自检

```bash
python -m ruff check .
python -m pytest -q
python scripts/apd.py build examples/demo/content_model.json -o build/demo.docx
python scripts/apd.py qa build/demo.docx
```

在 Windows + Word 上再确认一次：打开 `build/demo.docx` → 是页面视图 100% →
目录已展开 → 公式可双击编辑 → 另存 PDF 后页码正确。

## 6. 升级版本

1. 改 `scripts/_version.py` 里的版本号。
2. 在 `CHANGELOG.md` 顶部加 `## [x.y.z] - 日期`（`tests/test_cli.py` 会校验 CHANGELOG 含版本号）。
3. 打 tag 推送，Release 工作流自动出包。
