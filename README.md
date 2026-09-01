<!-- PROJECT LOGO -->
<div align="center">
  <br />
  
  <h1 style="font-size: 3.5rem; margin-bottom: 0.5rem;">
    <span style="color: #2E86C1;">ClassIn</span> 视频下载器
  </h1>

[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![Release][release-shield]][release-url]
[![Downloads][downloads-shield]][release-url]

<h4>
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">报告问题</a>
  ·    
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=feature_request.yml">功能请求</a>
</h4>
</div>

<!-- ABOUT THE PROJECT -->
## 关于项目

一款功能强大的 ClassIn 视频下载器，提供图形界面和命令行两种版本，旨在帮助用户轻松下载 ClassIn 课程视频 📹
具备高速下载、批量处理功能，并提供基于 C# 构建的友好 Avalonia 图形界面 🚀
支持解析 ClassIn 平台的视频链接，高效管理下载任务 ⚡
需要使用 HTTP Debugger Pro 捕获视频请求，为获取视频资源提供可靠方式 🔍

*项目仍处于测试阶段，预计在1.5版本完成所有问题修复。由于学业任务繁忙开发进度较慢

> **⚠️ 重要声明**
>
> 本工具为技术学习项目。根据 ClassIn 平台规则，学生账户通常仅拥有课程回放的**观看权**，**无下载权**。未经授权下载课程内容可能**侵犯知识产权并违反平台用户协议**。请仅在拥有合法权利的课程内容上使用本工具，并自行承担相关风险。

> **🔄 开发通知**
>
> 为了实现更加便捷的解析流程，此软件正在开发自助抓包功能。最新版本已发布，欢迎前往 <a href="https://github.com/ZMH21306/ClassIn-DL/releases">Releases</a> 下载，或前往 <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">报告问题</a>

<!-- COMPATIBILITY -->
## 兼容性

|    平台    |   最低要求    |      架构      |   兼容性   |
|:----------:|:-------------:|:--------------:|:----------:|
| 🪟 **Windows** |   `7 SP1+`    | `x86_64`/`x86`/`arm64` |     ✅      |
|  🐧 **Linux**  | `glibc 2.35+` | `x86_64`/`arm64` |     ✅      |
|  🍎 **macOS**  |   `11.0+`     | `x86_64`/`arm64` |     ✅      |

<!-- ENCODING ISSUE -->
## 编码问题说明

> **🔧 HTTP Debugger Pro 乱码问题**
>
> 目前已知，使用 HTTP Debugger Pro 捕获请求时，部分用户会遇到**输出内容乱码**问题。这是由于 ClassIn 服务器响应使用了不同字符编码所致。
>
> **开发计划**：我们已开发简易的**自动编码修复功能**，计划在后续版本中集成更完善的编码修复，以解决此问题。该功能目前正在积极开发中。
>
> 若遇到此问题，可尝试在抓包时另存为 JSON 文件手动复制内容，或关注我们的更新。

<!-- ROADMAP -->
## 开发路线

### ✅ 已完成功能
- ✅ 图形用户界面 (WPF)
- ✅ 命令行界面支持
- ✅ 基本的 ClassIn 视频下载功能
- ✅ 批量视频下载
- ✅ 从 HTTP 请求中解析视频链接
- ✅ 可配置线程数的多线程下载
- ✅ 实时下载速度显示
- ✅ 下载进度跟踪
- ✅ 错误处理和日志记录
- ✅ 可配置的下载目录
- ✅ 可调节的并发下载限制

### 🔄 计划中功能
- 🔄 自助抓包 (长期目标)
- 🔄 **自动编码修复 (GBK→UTF-8)**

访问 [GitHub Issues](https://github.com/ZMH21306/ClassIn-DL/issues) 查看所有功能请求（和已知问题）。

<!-- TUTORIAL VIDEO -->
## 教程视频

📹 **教程视频即将推出，敬请期待！**

<!-- 稍后添加视频链接 -->

<!-- DOWNLOAD LINKS -->
## 下载

> [!TIP]
> 为获得最佳兼容性，请使用最新版本的工具。

获取适用于各平台的 ClassIn 视频下载器最新版本：

| 平台 | 架构 | 文件 | 说明 |
|------|------|------|------|
| Linux | arm64 | `ClassInDL_1.0.2_Linux_arm64.AppImage` | 通用应用 |
| Linux | arm64 | `ClassInDL_1.0.2_Linux_arm64.deb` | Debian/Ubuntu |
| Linux | arm64 | `ClassInDL_1.0.2_Linux_arm64.tar.gz` | 便携版 |
| Linux | x64 | `ClassInDL_1.0.2_Linux_x64.AppImage` | 通用应用 |
| Linux | x64 | `ClassInDL_1.0.2_Linux_x64.deb` | Debian/Ubuntu |
| Linux | x64 | `ClassInDL_1.0.2_Linux_x64.tar.gz` | 便携版 |
| Windows | arm64 | `ClassInDL_1.0.2_Windows_arm64.exe` | NSIS 安装程序 |
| Windows | arm64 | `ClassInDL_1.0.2_Windows_arm64.msi` | MSI 安装程序 |
| Windows | arm64 | `ClassInDL_1.0.2_Windows_arm64.zip` | 便携版 |
| Windows | x64 | `ClassInDL_1.0.2_Windows_x64.exe` | NSIS 安装程序 |
| Windows | x64 | `ClassInDL_1.0.2_Windows_x64.msi` | MSI 安装程序 |
| Windows | x64 | `ClassInDL_1.0.2_Windows_x64.zip` | 便携版 |
| Windows | x86 | `ClassInDL_1.0.2_Windows_x86.exe` | NSIS 安装程序 |
| Windows | x86 | `ClassInDL_1.0.2_Windows_x86.msi` | MSI 安装程序 |
| Windows | x86 | `ClassInDL_1.0.2_Windows_x86.zip` | 便携版 |
| macOS | arm64 | `ClassInDL_1.0.2_macOS_arm64.dmg` | 磁盘映像 |
| macOS | arm64 | `ClassInDL_1.0.2_macOS_arm64.tar.gz` | 便携版 |
| macOS | x64 | `ClassInDL_1.0.2_macOS_x64.dmg` | 磁盘映像 |
| macOS | x64 | `ClassInDL_1.0.2_macOS_x64.tar.gz` | 便携版 |

> 完整下载链接请前往 [GitHub Releases](https://github.com/ZMH21306/ClassIn-DL/releases/latest) 查看。
>
> **SHA256 校验**：每个 Release 附带 `SHA256SUMS.txt`，请在下载后校验文件完整性。
> - Windows: `Get-FileHash <文件名> -Algorithm SHA256`
> - Linux/macOS: `sha256sum <文件名>`

<!-- CONTRIBUTING -->
## 贡献指南

贡献让开源社区成为学习、启发和创造的绝佳场所。您的任何贡献我们都**深表感谢**。

若有建议，可 Fork 仓库并创建 Pull Request，也可直接打开带有 "Enhancement" 标签的 Issue。别忘了给项目点个星星 ⭐！再次感谢！

1. Fork 本仓库
2. 创建功能分支 (git checkout -b feature/AmazingFeature)
3. 提交更改 (git commit -m 'Add some AmazingFeature')
4. 推送到分支 (git push origin feature/AmazingFeature)
5. 打开 Pull Request

<a href="https://github.com/ZMH21306/ClassIn-DL/graphs/contributors"><img src="http://contrib.nn.ci/api?repo=ZMH21306/ClassIn-DL" alt="贡献者" /></a>

<!-- LICENSE -->
## 许可证

根据 GPL v3.0 许可证分发。更多信息请参阅 `LICENSE` 文件。

版权所有 © 2026 ZMH。

<!-- CONTACT -->
## 联系方式

* [电子邮箱](mailto:zhounbdev@gmail.com) - zhounbdev@gmail.com
* [QQ 群](https://qm.qq.com/q/PlUBdzqZCm) - 2130606191

## Star 历史
## Star History

<a href="https://www.star-history.com/?repos=ZMH21306%2FClassIn-DL&type=date&legend=top-left">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/chart?repos=ZMH21306/ClassIn-DL&type=date&theme=dark&legend=top-left&sealed_token=FVBBc-3cPWNb4wJAM7xbbO8PjNMGeQxHJIrf1QzB1gaZGdRiNB9Y-4Bqwvlc8Gz9mpBbZgPo_FwBGtUAB_0k97KKQT6IBfd1k-ksi_7ZWly5nD1HIbtX" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/chart?repos=ZMH21306/ClassIn-DL&type=date&legend=top-left&sealed_token=FVBBc-3cPWNb4wJAM7xbbO8PjNMGeQxHJIrf1QzB1gaZGdRiNB9Y-4Bqwvlc8Gz9mpBbZgPo_FwBGtUAB_0k97KKQT6IBfd1k-ksi_7ZWly5nD1HIbtX" />
   <img alt="Star History Chart" src="https://api.star-history.com/chart?repos=ZMH21306/ClassIn-DL&type=date&legend=top-left&sealed_token=FVBBc-3cPWNb4wJAM7xbbO8PjNMGeQxHJIrf1QzB1gaZGdRiNB9Y-4Bqwvlc8Gz9mpBbZgPo_FwBGtUAB_0k97KKQT6IBfd1k-ksi_7ZWly5nD1HIbtX" />
 </picture>
</a>

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[forks-shield]: https://img.shields.io/github/forks/ZMH21306/ClassIn-DL.svg?style=for-the-badge
[forks-url]: https://github.com/ZMH21306/ClassIn-DL/network/members
[stars-shield]: https://img.shields.io/github/stars/ZMH21306/ClassIn-DL.svg?style=for-the-badge
[stars-url]: https://github.com/ZMH21306/ClassIn-DL/stargazers
[issues-shield]: https://img.shields.io/github/issues/ZMH21306/ClassIn-DL.svg?style=for-the-badge
[issues-url]: https://github.com/ZMH21306/ClassIn-DL/issues
[release-shield]: https://img.shields.io/github/v/release/ZMH21306/ClassIn-DL?style=for-the-badge
[release-url]: https://github.com/ZMH21306/ClassIn-DL/releases/latest
[downloads-shield]: https://img.shields.io/github/downloads/ZMH21306/ClassIn-DL/total?style=for-the-badge
[qqgroup-shield]: https://img.shields.io/badge/QQ_Group-2130606191-blue.svg?color=blue&style=for-the-badge
[qqgroup-url]: https://qm.qq.com/q/4w0AZhrAcU
