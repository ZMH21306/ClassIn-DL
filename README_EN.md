<h4 align="right">
  <a href="README_EN.md">English</a> | 简体中文
</h4>

> **⚠️ Important Notice**
>
> To facilitate the development of packet capture functionality, the current project is temporarily suspended. Please check the latest progress in the packet capture software <a href="https://github.com/ZMH21306/FlowReveal">prototype repository</a>

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
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">Report Issue</a>
  ·    
  <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=feature_request.yml">Feature Request</a>
</h4>
</div>

<!-- ABOUT THE PROJECT -->
## About The Project

A powerful ClassIn video downloader providing both GUI and command-line versions, designed to help users easily download ClassIn course videos 📹
Features high-speed downloading, batch processing, and a friendly Avalonia-based GUI built with C# 🚀
Supports parsing video links from the ClassIn platform and efficiently managing download tasks ⚡
Requires HTTP Debugger Pro to capture video requests, providing a reliable way to obtain video resources 🔍

*The project is still in the testing phase, with all issues expected to be fixed in version 1.5. Development progress is slow due to heavy academic workload.

> **⚠️ Important Disclaimer**
>
> This tool is a technical learning project. According to ClassIn platform rules, student accounts typically only have **viewing rights** for course recordings, **not downloading rights**. Unauthorized downloading of course content may **infringe intellectual property rights and violate platform user agreements**. Please use this tool only for course content you are legally entitled to, and bear any associated risks yourself.

> **🔄 Development Notice**
>
> To achieve a more convenient parsing process, this software is planning to introduce packet capture functionality, so release versions are temporarily on hold. Welcome to <a href="https://github.com/ZMH21306/ClassIn-DL/issues/new?template=bug_report.yml">Report Issue</a>

<!-- COMPATIBILITY -->
## Compatibility

|    Platform    | Minimum Requirements |    Architecture   | Compatibility |
|:--------------:|:--------------------:|:-----------------:|:-------------:|
| 🪟 **Windows** |      `7 SP1+`        | `x86_64`/`x86`/`arm64` |      ✅       |
|  🐧 **Linux**  |     `glibc 2.35+`    | `x86_64`/`arm64`      |      ✅       |
|  🍎 **macOS**  |      `11.0+`         | `x86_64`/`arm64`      |      ✅       |

<!-- ENCODING ISSUE -->
## Encoding Issue Explanation

> **🔧 HTTP Debugger Pro Garbled Text Issue**
>
> It is known that some users experience **garbled output** when capturing requests with HTTP Debugger Pro. This is caused by different character encodings used by the ClassIn server response.
>
> **Development Plan**: We have developed a simple **automatic encoding repair feature** and plan to integrate more comprehensive encoding fixes in future versions to resolve this issue. This feature is currently under active development.
>
> If you encounter this issue, you can try saving the capture as a JSON file and manually copying the content, or follow our updates.

<!-- ROADMAP -->
## Roadmap

### ✅ Completed Features
- ✅ Graphical User Interface (WPF)
- ✅ Command-Line Interface support
- ✅ Basic ClassIn video download functionality
- ✅ Batch video download
- ✅ Parse video links from HTTP requests
- ✅ Configurable multi-threaded downloading
- ✅ Real-time download speed display
- ✅ Download progress tracking
- ✅ Error handling and logging
- ✅ Configurable download directory
- ✅ Adjustable concurrent download limits

### 🔄 Planned Features
- 🔄 Self-service packet capture (long-term goal)
- 🔄 **Automatic encoding repair (GBK→UTF-8)**

Visit [GitHub Issues](https://github.com/ZMH21306/ClassIn-DL/issues) to see all feature requests (and known issues).

<!-- TUTORIAL VIDEO -->
## Tutorial Video

📹 **Tutorial video coming soon, stay tuned!**

<!-- Uncomment when video link is available -->
<!-- [![Tutorial Video](https://img.youtube.com/vi/VIDEO_ID/0.jpg)](https://www.youtube.com/watch?v=VIDEO_ID) -->

<!-- DOWNLOAD LINKS -->
## Download

> [!TIP]
> For best compatibility, please use the latest version of the tool.

Get the latest version of ClassIn Video Downloader for each platform:

| Platform | Architecture | Download Links |
|:--------:|:------------:|:---------------:|
| Windows  | x86_64       | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Windows-x64.exe) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Windows-x64.exe) |
| Windows  | arm64        | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Windows-arm64.exe) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Windows-arm64.exe) |
| Linux    | x86_64       | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Linux-x64) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Linux-x64) |
| Linux    | arm64        | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Linux-arm64) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-Linux-arm64) |
| macOS    | x86_64       | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-macOS-x64) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-macOS-x64) |
| macOS    | arm64        | [GitHub Direct](https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-macOS-arm64) <br> [CDN Mirror](https://gh-proxy.org/https://github.com/ZMH21306/ClassIn-DL/releases/download/v1.0.0-Avalonia_UI/Classin_DL-v1.0.0-macOS-arm64) |

<!-- CONTRIBUTING -->
## Contributing Guidelines

Contributions make the open source community a wonderful place for learning, inspiration, and creativity. Any contribution you make is **greatly appreciated**.

If you have suggestions, you can fork the repository and create a pull request, or directly open an issue with the "Enhancement" tag. Don't forget to give the project a star ⭐! Thanks again!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<a href="https://github.com/ZMH21306/ClassIn-DL/graphs/contributors"><img src="http://contrib.nn.ci/api?repo=ZMH21306/ClassIn-DL" alt="Contributors" /></a>

<!-- LICENSE -->
## License

Distributed under the GPL v3.0 License. See `LICENSE` for more information.

Copyright © 2026 ZMH.

<!-- CONTACT -->
## Contact

* [Email](mailto:zhounbdev@gmail.com) - zhounbdev@gmail.com
* [QQ Group](https://qm.qq.com/q/PlUBdzqZCm) - 2130606191

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
