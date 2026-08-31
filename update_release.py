#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Update release.yml with proper naming convention"""

import re
import sys

filepath = ".github/workflows/release.yml"

# Read with UTF-8 first, fall back to GBK
try:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
except UnicodeDecodeError:
    with open(filepath, 'r', encoding='gbk') as f:
        content = f.read()

original = content

# 1. Update env section
content = re.sub(
    r"permissions:\n  contents: write\n",
    """permissions:
  contents: write

env:
  PROJECT_NAME: 'Classin视频解析下载工具'
  FILE_VERSION: '0.7.0'
  DOTNET_VERSION: '8.0.x'
""",
    content
)

# 2. Update Collect artifacts - change NAME to use env vars
content = re.sub(
    r'VERSION="\$\{\{ steps\.version\.outputs\.version \}\}"\n\s*NAME="ClassIn-DL_\$\{\{ VERSION \}\}_\$\{\{ matrix\.os_name \}\}_\$\{\{ matrix\.arch \}\}"',
    'NAME="${{ env.PROJECT_NAME }}_${{ env.FILE_VERSION }}_${{ matrix.os_name }}_${{ matrix.arch }}"',
    content
)

# 3. Update portable zip name
content = re.sub(
    r'\$\{GITHUB_WORKSPACE\}/artifacts/\$\{NAME\}\.zip',
    '${GITHUB_WORKSPACE}/artifacts/${NAME}_portable.zip',
    content
)

# 4. Update Flatten artifacts to accept all files
content = re.sub(
    r"find artifacts -type f \\\\( -name '\*\.exe' -o -name '\*\.zip' -o -name '\*\.tar\.gz' \\\\) \| while read f; do",
    "find artifacts -type f | while read f; do",
    content
)

# 5. Update release notes section
old_notes_start = r"      - name: Generate release notes\n        id: notes\n        run: \|"
old_notes_end = r'          } >> "\$GITHUB_OUTPUT"'

# Find the release notes section
start_match = re.search(old_notes_start, content)
end_match = re.search(old_notes_end, content)

if start_match and end_match:
    new_notes = """      - name: Generate release notes
        id: notes
        shell: bash
        run: |
          TAG="${{ steps.version.outputs.tag }}"
          FILE_VERSION="${{ env.FILE_VERSION }}"
          DATE=$(date -u +%Y-%m-%d)
          REPO="${{ github.repository }}"
          BASE_URL="https://github.com/${REPO}/releases/download/${TAG}"
          PROJECT_NAME="${{ env.PROJECT_NAME }}"
          
          cat > notes.md << 'NOTES_EOF'
          ## ${PROJECT_NAME} ${TAG}

          **发布日期**: ${DATE}

          ---

          ### 新增功能

          - Avalonia UI 图形化版本（.NET 8）
          - 启动画面（Splash Screen）优化
          - URL 安全验证（SSRF 防护、域名白名单）
          - 异步日志服务（高性能文件写入）
          - 异步确认对话框
          - 配置保存重试机制
          - 内存监控服务优化
          - 进度变化检测 BUG 修复
          - 课程名称乱码问题修复

          ### 下载说明

          | 平台 | 架构 | 文件 | 说明 |
          |------|------|------|------|
          | Windows | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_x64.exe`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_x64.exe) | 单文件便携 |
          | Windows | x86 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_x86.exe`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_x86.exe) | 单文件便携 |
          | Windows | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_arm64.exe`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_arm64.exe) | 单文件便携 |
          | Windows | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_x64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_x64_portable.zip) | 便携版 ZIP |
          | Windows | x86 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_x86_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_x86_portable.zip) | 便携版 ZIP |
          | Windows | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_Windows_arm64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Windows_arm64_portable.zip) | 便携版 ZIP |
          | Linux | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_Linux_x64`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Linux_x64) | 单文件便携 |
          | Linux | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_Linux_arm64`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Linux_arm64) | 单文件便携 |
          | Linux | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_Linux_x64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Linux_x64_portable.zip) | 便携版 |
          | Linux | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_Linux_arm64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_Linux_arm64_portable.zip) | 便携版 |
          | macOS | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_macOS_x64`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_macOS_x64) | 单文件便携 |
          | macOS | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_macOS_arm64`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_macOS_arm64) | 单文件便携 |
          | macOS | x64 | [`${PROJECT_NAME}_${FILE_VERSION}_macOS_x64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_macOS_x64_portable.zip) | 便携版 |
          | macOS | arm64 | [`${PROJECT_NAME}_${FILE_VERSION}_macOS_arm64_portable.zip`](${BASE_URL}/${PROJECT_NAME}_${FILE_VERSION}_macOS_arm64_portable.zip) | 便携版 |

          ### 哈希校验

          <details>
          <summary>展开查看 SHA256 校验和</summary>

          ```
          ${SHA256_SUMS}
          ```

          **校验方法**：
          - Windows: `Get-FileHash <文件名> -Algorithm SHA256`
          - Linux/macOS: `sha256sum <文件名>`

          </details>

          ---

          **完整变更日志**: [CHANGELOG.md](https://github.com/${REPO}/blob/main/CHANGELOG.md)
          NOTES_EOF
          
          SHA256_SUMS=$(cat SHA256SUMS.txt)
          sed -i "s|SHA256_SUMS|${SHA256_SUMS}|g" notes.md
          
          echo "body<<EOF"
          cat notes.md
          echo "EOF"
"""
    
    content = content[:start_match.start()] + new_notes + content[end_match.end():]

# Write back with UTF-8
with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"Updated {filepath}")
print(f"Changes: {len(content) - len(original)} characters")