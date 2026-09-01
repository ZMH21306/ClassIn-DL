# AI 执行契约 - ClassIn-DL

> 本文档是 AI（包括 Trae/GitHub Copilot 等）执行项目操作时的**唯一权威指令来源**。
> 自然语言指令按本表映射，不得自行推断执行路径。

---

## 指令映射表

| 用户自然语言 | AI 执行操作 |
|------------|------------|
| **提交代码** / **commit** / **保存改动** | 按「Commit 规范」生成消息并执行 `git add` + `git commit` |
| **发布新版本** / **release** / **打标签** | 运行 `.\scriptselease.ps1`（若工作区干净且有功能/修复变更） |
| **发布 v1.2.3** / **release v1.2.3** | 运行 `.\scriptselease.ps1 -Version 1.2.3` |
| **更新 CHANGELOG** / **更新日志** | 运行 `.\scripts\generate-changelog.ps1 -Version x.y.z` |
| **本地构建** / **build** / **编译** | 运行 `dotnet build` 或 `dotnet run` |
| **运行测试** / **test** | 运行 `dotnet test` |

---

## Commit 规范（Conventional Commits）

```
<type>(<scope>): <subject>
```

### Type 与版本影响

| Type | 含义 | 自动版本递增 |
|------|------|-------------|
| `feat` | 新功能 | MINOR +1 |
| `fix` | Bug 修复 | PATCH +1 |
| `perf` | 性能优化 | PATCH +1 |
| `refactor` | 重构（不增功能、不修 bug） | 无 |
| `docs` | 文档变更 | 无 |
| `style` | 代码格式 | 无 |
| `test` | 测试相关 | 无 |
| `build` | 构建/依赖 | 无 |
| `ci` | CI/CD 配置 | 无 |
| `chore` | 杂项维护 | 无 |

### Subject 规则
- 使用中文，祈使语气
- 不超过 50 个字符，句尾不加句号
- 描述做了什么（而非"改了什么"）

### Scope（常用）
`parse`、`download`、`ui`、`config`、`build`、`deps`、`*`（无 scope）

### 示例

```text
feat(parse): 支持解析 EEO 视频课程页面
fix(download): 修复断网后重试次数计算错误
perf(parser): 优化正则匹配减少内存分配
refactor(download): 将并发控制抽取为独立服务
```

---

## 发布前检查清单（必须全部通过，否则停止）

执行 `release` 前，AI 必须依次验证：

1. **工作区干净**
   ```powershell
   git status --porcelain
   ```
   输出为空才继续；否则先提示用户提交未完成的改动。

2. **无未提交 CHANGELOG 条目**
   检查 `CHANGELOG.md` 是否存在 `## [Unreleased]`，不存在则先运行 `generate-changelog.ps1` 生成条目。

3. **目标 tag 不存在**
   ```powershell
   git tag -l "v${Version}"
   ```
   若已有此 tag 则报错停止，不可覆盖远端 tag。

4. **默认分支可写**
   ```powershell
   git symbolic-ref refs/remotes/origin/HEAD
   ```
   确认结果为 `refs/remotes/origin/main`（或当前项目的实际默认分支）。

---

## 版本号管理

- **单一版本源**：`<ProjectName>.csproj` 中的 `<Version>` 字段
- `AssemblyVersion` 和 `FileVersion` 由 `<Version>` 自动派生，**无需单独修改**
- `appsettings.json` 中**不包含**版本号，不在此文件维护版本

---

## 发布流程（完整步骤）

```text
1. 检查工作区干净（git status）
2. 运行 generate-changelog.ps1 更新 CHANGELOG
3. 运行 release.ps1 自动：
   a. 读取当前版本（从 .csproj）
   b. 分析 commit 确定语义化版本（major/minor/patch）
   c. 更新 .csproj 版本号
   d. 提交版本更新到 main
   e. 打 v*x.y.z tag 并推送到远端
4. GitHub Actions (release.yml) 自动触发
   -> 7 平台并行构建（dotnet publish --self-contained）
   -> 收集产物（.exe + .msi + .zip + .deb + .AppImage + .tar.gz + .dmg）
   -> 生成 SHA256SUMS.txt
   -> 创建 GitHub Release
```

---

## 产物规范（统一发行标准）

### 资产命名
```
ClassInDL_<ver>_<OS>_<arch>.<ext>
```
例：`ClassInDL_1.0.2_Windows_x64.exe`、`ClassInDL_1.0.2_Linux_x64.tar.gz`

### 产物矩阵（7 平台）

| 平台 | 架构 | 产物 |
|------|------|------|
| Windows | x64 | `.exe`（NSIS）+ `.msi`（WiX）+ `.zip`（便携） |
| Windows | x86 | `.exe` + `.msi` + `.zip` |
| Windows | arm64 | `.exe` + `.msi` + `.zip` |
| Linux | x64 | `.deb` + `.AppImage`+ `.tar.gz` |
| Linux | arm64 | `.deb` + `.AppImage` + `.tar.gz` |
| macOS | x64 | `.dmg` + `.tar.gz` |
| macOS | arm64 | `.dmg` + `.tar.gz` |

### Release Notes 结构（标准化模板）
```markdown
# ClassIn视频下载工具 v<tag>
发布日期：<YYYY-MM-DD>

## 新增功能
- ...

## 改进
- ...

## 修复
- ...

## 依赖与安全
- ...

## 下载

| 平台 | 架构 | 文件 | 说明 |
|------|------|------|------|

## 哈希校验

<details>
<summary>展开查看 SHA256 校验和</summary>
```
SHA256SUMS.txt 内容
```
</details>

完整变更日志：CHANGELOG.md
```

### CHANGELOG 模板
```markdown
## [<ver>] - <YYYY-MM-DD>

### 新增
- ...

### 改进
- ...

### 修复
- ...

### 依赖与安全
- ...
```

---

## 项目特有说明

| 项目 | 技术栈 | 构建命令 | CI 产物 |
|------|--------|---------|---------|
| ClassIn-DL | C# .NET 8 + Avalonia UI | `dotnet publish -c Release -r <rid> --self-contained` | .exe + .msi + .zip + .deb + .AppImage + .tar.gz + .dmg（共 22 项） |
| LlamaUI | Rust + Tauri 2 | `cargo tauri build --target <rust_target>` | .exe/.msi/.zip + .deb/.AppImage/.tar.gz + .dmg（共 22 项） |

---

## 常见错误与处理

| 错误现象 | 原因 | 处理方式 |
|---------|------|---------|
| `release.ps1` 报"工作区有未提交变更" | 有未 commit 的文件 | 先 commit，再执行 release |
| tag 已存在报错 | 同一版本被发布过两次 | 删除本地 tag 后重试，或手动 bump 版本号 |
| CHANGELOG 没有 `[Unreleased]` 段落 | 从未发布过或手动删除了 | `generate-changelog.ps1` 会 prepend，结果正常 |
| `git push origin main` 失败 | 分支名不是 main / 无权限 | 检查 `git remote show origin`，手动推送 |

---

## CI 工作流参考

- 触发条件：推送 `v*` tag 或手动 `workflow_dispatch`
- 构建矩阵：7 平台（Windows x64/x86/arm64、Linux x64/arm64、macOS x64/arm64）
- 发布地址：https://github.com/ZMH21306/ClassIn-DL/releases
- 产物命名：`ClassInDL_{VERSION}_{OS}_{ARCH}.<ext>`
- SHA256SUMS：每个 release 自动附带产物校验文件
