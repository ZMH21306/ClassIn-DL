<h4 align="right">
  English | <a href="README.md">简体中文</a>
</h4>

<!-- PROJECT LOGO -->
<div align="center">
  <br />
  
  <h1 style="font-size: 3.5rem; margin-bottom: 0.5rem;">
    <span style="color: #2E86C1;">ClassIn</span> Video Downloader
  </h1>

[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![Release][release-shield]][release-url]
[![Downloads][downloads-shield]][release-url]

<h4>
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">Report Bug</a>
 ·    
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=feature_request.yml">Request Feature</a>
</h4>

</div>

<!-- ABOUT THE PROJECT -->
## About The Project

A powerful ClassIn video downloader with both graphical and command-line interfaces, designed to help users easily download ClassIn course videos 📹

Features high-speed downloading, batch processing, and a friendly WPF graphical interface built with C# 🚀

Supports parsing video links from the ClassIn platform and efficiently managing download tasks ⚡

Requires HTTP Debugger Pro to capture video requests, providing a reliable way to obtain video resources 🔍

> **⚠️ Important Notice**
>
> This tool is a technical learning project. According to ClassIn platform rules, student accounts typically only have **viewing rights** for course replays, **not downloading rights**. Unauthorized downloading of course content may **infringe on intellectual property rights and violate platform user agreements**. Please only use this tool on course content for which you have legal rights, and assume all related risks yourself.

> **🔄 Migration Notice**
>
> This project has currently migrated from WPF to **Avalonia C#** to support cross-platform compatibility (Windows, Linux, macOS). The current version still has many issues, please feel free to <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">report problems</a>

<!-- COMPATIBILITY -->
## Compatibility

|    Platform    |   Minimum Requirement    |      Architecture      |   Compatibility   |
|:----------:|:-------------:|:--------------:|:----------:|
| 🪟 **Windows** |   `7 SP1+`    | `x86_64`/`x86`/`arm64` |     ✅      |
|  🐧 **Linux**  | `glibc 2.35+` | `x86_64`/`arm64` |     ❌      |
|  🍎 **macOS**  |   `11.0+`     | `x86_64`/`arm64` |     ❌      |

<!-- ENCODING ISSUE -->
## Encoding Issue Description

> **🔧 HTTP Debugger Pro Garbled Text Issue**
>
> Currently known, when using HTTP Debugger Pro to capture requests, some users may encounter **garbled output content** issues. This is caused by ClassIn server responses using different character encodings.
>
> **Development Plan**: We have developed a simple **automatic encoding repair feature** and plan to integrate more complete encoding repair in subsequent versions to solve this issue. This feature is currently under active development.
>
> If you encounter this issue, you can try saving as a JSON file during packet capture and manually copy the content, or follow our updates.

<!-- ROADMAP -->
## Development Roadmap

### ✅ Completed Features
- ✅ Graphical User Interface (WPF)
- ✅ Command-line interface support
- ✅ Basic ClassIn video download functionality
- ✅ Batch video downloading
- ✅ Parse video links from HTTP requests
- ✅ Multi-threaded download with configurable thread count
- ✅ Real-time download speed display
- ✅ Download progress tracking
- ✅ Error handling and logging
- ✅ Configurable download directory
- ✅ Adjustable concurrent download limit

### 🔄 Planned Features
- 🔄 Self-service packet capture (long-term goal)
- 🔄 **Automatic encoding repair (GBK→UTF-8)**

Visit [GitHub Issues](https://github.com/ZMH21306/ClassIn-DL/issues) to view all feature requests (and known issues).

<!-- TUTORIAL VIDEO -->
## Tutorial Video

📹 **Tutorial video coming soon, stay tuned!**

<!-- Add video link later -->

<!-- DOWNLOAD LINKS -->
## Download

> [!TIP]
> For best compatibility, please use the latest version of the tool.

Get the latest version of ClassIn Video Downloader for Windows:

| Platform | Architecture | Download Links |
|:----:|:----:|:--------:|
| Windows | x86_64 | [GitHub Direct Link](https://github.com/ZMH21306/ClassIn-DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-x64.exe) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/Classin-DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-x64.exe) |
| Windows | x86    | [GitHub Direct Link](https://github.com/ZMH21306/ClassIn-DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-x86.exe) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/Classin_DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-x86.exe) |
| Windows | arm64  | [GitHub Direct Link](https://github.com/ZMH21306/ClassIn-DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-arm64.exe) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/Classin_DL/releases/download/v0.8.0/Classin_DL-v0.8.0-Windows-arm64.exe) |

<!-- CONTRIBUTING -->
## Contribution Guide

Contributions make the open-source community a great place to learn, inspire, and create. Any contribution you make is **greatly appreciated**.

If you have suggestions, you can Fork the repository and create a Pull Request. You can also directly open an Issue with the "Enhancement" label. Don't forget to give the project a star⭐! Thank you again!

1. Fork this repository
2. Create a feature branch (git checkout -b feature/AmazingFeature)
3. Commit your changes (git commit -m 'Add some AmazingFeature')
4. Push to the branch (git push origin feature/AmazingFeature)
5. Open a Pull Request

Thank you to all contributors who participated in this project!

<a href="https://github.com/ZMH21306/ClassIn-DL/graphs/contributors"><img src="http://contrib.nn.ci/api?repo=ZMH21306/ClassIn-DL" alt="Contributors" /></a>

<!-- LICENSE -->
## License

Distributed under the GPL v3.0 License. For more information, please refer to the `LICENSE` file.

Copyright © 2025 ZMH.

<!-- CONTACT -->
## Contact

* [Email](mailto:2130606191@qq.com) - 2130606191@qq.com
* [QQ Group](https://qm.qq.com/q/PlUBdzqZCm) - 2130606191

## Acknowledgments

* Special thanks to all open-source projects that made this tool possible!

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=ZMH21306/ClassIn-DL&type=timeline&legend=top-left)](https://www.star-history.com/#ZMH21306/ClassIn-DL&type=timeline&legend=top-left)

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