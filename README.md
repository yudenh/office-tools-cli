# Office Tools

本项目提供一个本地 Office/PDF 文件处理命令行工具，支持 DOCX 作者信息清理、DOCX/PDF 转换、PDF 合并、DOCX 内容查找和图片水印。
附带有SKILL，便于AI调用。

资源推荐：[专为 AI 智能体设计的 Office 套件。零依赖。全平台。](https://github.com/iOfficeAI/OfficeCLI)

## 功能

| 命令 | 功能 |
| --- | --- |
| `remove-author` | 清除 DOCX 作者信息 |
| `docx-to-pdf` | DOCX 转 PDF |
| `pdf-to-docx` | PDF 转 DOCX |
| `pdf-to-images` | PDF 每页转 PNG 图片 |
| `merge-pdfs` | 合并目录下的 PDF |
| `find-docx` | 在 DOCX 中查找字符串 |
| `watermark-docx` | 给 DOCX 添加图片水印 |
| `watermark-pdf` | 给 PDF 添加图片水印 |

## 环境要求

- Python 3.10+
- Windows 环境
- `docx-to-pdf` 和 `watermark-docx` 需要本机安装 Microsoft Word 或兼容的 WPS/Word COM 环境

## 安装

在项目根目录执行：

```powershell
python -m pip install --user .
```

安装后可在任意目录使用全局命令：

```powershell
office-tools
```

如果提示找不到 `office-tools` 命令，请确认 Python 用户脚本目录已加入 `PATH`。

Windows 常见目录：

```text
%APPDATA%\Python\Python313\Scripts
```

其中 `Python313` 需要按实际 Python 版本调整。

## 用法

无参数运行会进入交互式菜单：

```powershell
office-tools
```

非交互命令格式：

```powershell
office-tools <command> <path> [options]
```

适合 AI 或脚本调用时，建议加上 `--json`：

```powershell
office-tools <command> <path> [options] --json
```

## 示例

清除单个 DOCX 的作者信息：

```powershell
office-tools remove-author "D:\docs\a.docx" --json
```

批量将目录中的 DOCX 转为 PDF：

```powershell
office-tools docx-to-pdf "D:\docs" --mode dir --json
```

将 PDF 转为 DOCX：

```powershell
office-tools pdf-to-docx "D:\docs\a.pdf" --json
```

将 PDF 每页导出为 PNG 图片：

```powershell
office-tools pdf-to-images "D:\docs\a.pdf" --json
```

合并目录下的 PDF：

```powershell
office-tools merge-pdfs "D:\pdfs" --json
```

在 DOCX 中查找关键词：

```powershell
office-tools find-docx "D:\docs" --query "合同 金额" --json
```

给 DOCX 添加图片水印：

```powershell
office-tools watermark-docx "D:\docs\a.docx" --image "D:\logo.png" --json
```

给 PDF 添加图片水印：

```powershell
office-tools watermark-pdf "D:\docs\a.pdf" --image "D:\logo.png" --json
```

## 参数

| 参数 | 说明 |
| --- | --- |
| `command` | 要执行的命令 |
| `path` | 输入文件或目录路径 |
| `--mode auto\|file\|dir` | 输入路径类型，默认 `auto` |
| `--output <path>` | 单文件操作的输出文件或目录 |
| `--query <text>` | `find-docx` 要查找的字符串，多个关键词用空格分隔 |
| `--image <path>` | `watermark-docx` 和 `watermark-pdf` 使用的水印图片路径 |
| `--watermark-left <cm>` | 水印图片水平位置，`watermark-docx` 和 `watermark-pdf` 默认 `4` |
| `--watermark-top <cm>` | 水印图片垂直位置，`watermark-docx` 和 `watermark-pdf` 默认 `6` |
| `--watermark-width <cm>` | 水印图片宽度，`watermark-docx` 和 `watermark-pdf` 默认 `4` |
| `--watermark-height <cm>` | 水印图片高度，未指定时按图片比例自适应 |
| `--json` | 输出 JSON，便于脚本或 AI 解析 |

## 路径模式

- `--mode auto` 会根据输入路径是否为目录自动判断处理方式。
- `--mode file` 强制按单个文件处理。
- `--mode dir` 递归处理目录中匹配后缀的文件。
- 目录处理会跳过以 `~$` 开头的 Office 临时文件。
- `merge-pdfs` 只接受目录输入。

## JSON 输出

使用 `--json` 时会输出单个 JSON 对象：

```json
{
  "ok": true,
  "command": "find-docx",
  "mode": "dir",
  "input": "D:\\docs",
  "summary": {
    "total": 2,
    "ok": 2,
    "failed": 0,
    "elapsed_seconds": 0.12
  },
  "items": [
    {
      "path": "D:\\docs\\a.docx",
      "ok": true,
      "status": "matched",
      "matches": ["合同"]
    }
  ]
}
```

进程退出码为 `0` 表示成功，非 `0` 表示失败。目录批处理时也应检查 `summary.failed` 和每个 `items[].ok`。

## 卸载

```powershell
python -m pip uninstall office-tools-cli
```

## 注意事项

- `docx-to-pdf` 和 `watermark-docx` 依赖 Word/WPS COM 自动化。
- 处理前请确认目标文件没有被 Word 或 WPS 打开。
- `pdf-to-docx` 的转换质量取决于源 PDF 的结构。
- `pdf-to-images` 如果输出目录已存在，会删除后重新生成。
- `watermark-pdf` 默认覆盖原 PDF；单文件模式可用 `--output <path>` 输出到新文件。
