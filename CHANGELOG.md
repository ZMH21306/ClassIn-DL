# 更新日志

本文件记录 ClassIn 视频下载器项目的所有重要变更。格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

## [1.0.1] - 2026-08-08

### 改进
- 清理仓库垃圾文件（移除历史版本快照与第三方激活工具）
- 添加 7 平台 CI 发布工作流（Windows x64/x86/arm64、Linux x64/arm64、macOS x64/arm64）
- 添加 CHANGELOG、.gitignore、.editorconfig 等开源规范文件
- 添加统一构建与发布脚本（build/test/scripts/resources）

## [1.0.0] - 2026-08-08

### 新增
- Avalonia UI 图形化版本正式发布（C# .NET 8）
- 支持 Windows / Linux / macOS 三平台运行
- 批量视频解析与多线程并发下载
- 实时下载进度、速度与剩余时间显示
- 自定义下载路径与并发数限制
- 视频播放列表管理功能

### 改进
- 优化解析算法：增强对不同格式请求头的支持
- 改进 UI 设计：更直观的操作流程与视觉布局
- 增强稳定性：修复已知解析与下载问题

### 修复
- 解析逻辑优化：确保每次解析只保留最新视频信息
- 视频项管理：使用固定索引避免索引错误

---

## 版本说明

### 语义化版本

- **MAJOR**：破坏性变更（配置格式不兼容、核心解析逻辑重构）
- **MINOR**：新功能（向后兼容）
- **PATCH**：Bug 修复（向后兼容）
