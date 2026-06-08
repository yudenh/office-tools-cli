---
name: office-tools
description: Use when operating local Office files with this repository's Python CLI, including docx/pdf conversion, PDF merge, DOCX search, author cleanup, and watermark tasks.
---

# Office Tools

Use this skill when the user asks to process local Office files with the `office_tools` project.

## Invocation

Use the globally installed `office-tools` command. Prefer non-interactive CLI calls and always add `--json` when another agent or script will consume the result.

```powershell
office-tools <command> <path> [options] --json
```

If `office-tools` is not installed, install it from the repository root first:

```powershell
python -m pip install --user .
```

## Commands

| Task | Command |
| --- | --- |
| Clear DOCX author metadata | `remove-author <file-or-dir>` |
| Convert DOCX to PDF | `docx-to-pdf <file-or-dir>` |
| Convert PDF to DOCX | `pdf-to-docx <file-or-dir>` |
| Render PDF pages to PNG images | `pdf-to-images <file-or-dir>` |
| Merge all PDFs in one directory | `merge-pdfs <dir>` |
| Search text in DOCX files | `find-docx <file-or-dir> --query "keyword1 keyword2"` |
| Add image watermark to DOCX | `watermark-docx <file-or-dir> --image <image-path> [--watermark-left <point>] [--watermark-top <point>] [--watermark-width <point>] [--watermark-height <point>]` |

Watermark position and size options use Word point units. Defaults match the implementation: `--watermark-left 100`, `--watermark-top 200`, `--watermark-width 100`. If `--watermark-height` is not passed, height scales proportionally from the image aspect ratio.

## Path Handling

- Default `--mode auto` treats existing directories as directory input and everything else as file input.
- Use `--mode file` or `--mode dir` when the intent must be explicit.
- Directory mode recursively processes matching files and skips temporary Office files whose names start with `~$`.
- `merge-pdfs` only accepts a directory.

## Output

With `--json`, the command prints one JSON object:

```json
{
  "ok": true,
  "command": "find-docx",
  "mode": "dir",
  "input": "D:\\docs",
  "summary": { "total": 2, "ok": 2, "failed": 0, "elapsed_seconds": 0.12 },
  "items": [
    { "path": "D:\\docs\\a.docx", "ok": true, "status": "matched", "matches": ["合同"] }
  ]
}
```

Treat process exit code `0` as success and non-zero as failure. Also inspect `summary.failed` and per-item `ok` because directory operations may partly fail.

## Examples

```powershell
office-tools remove-author "D:\docs\a.docx" --json
office-tools docx-to-pdf "D:\docs" --mode dir --json
office-tools pdf-to-images "D:\docs\a.pdf" --json
office-tools find-docx "D:\docs" --query "合同 金额" --json
office-tools watermark-docx "D:\docs\a.docx" --image "D:\logo.png" --watermark-left 100 --watermark-top 200 --watermark-width 100 --json
office-tools merge-pdfs "D:\pdfs" --json
```

## Caveats

- `docx-to-pdf` and `watermark-docx` use Windows COM automation and require Microsoft Word or a compatible WPS/Word COM environment.
- Make sure target files are not open in Word/WPS before modifying or converting them.
- `pdf-to-docx` quality depends on the source PDF structure.
- `pdf-to-images` overwrites the derived output directory if it already exists.
- **Error handling**: If a command fails or raises an error, stop immediately and report the error to the user. Do NOT attempt to diagnose or fix the issue yourself.
