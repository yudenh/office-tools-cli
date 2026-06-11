---
name: office-tools
description: Use when operating local Office/PDF files with the globally installed `office-tools` CLI, including TXT/DOCX/PDF conversion, PDF merge, and watermark tasks.
---

# Office Tools

Use this skill when the user asks to process local Office/PDF files with the globally installed `office-tools` command.

## Invocation

Use the globally installed `office-tools` command. Prefer non-interactive CLI calls and always add `--json` when another agent or script will consume the result.

```powershell
office-tools <command> <path> [options] --json
```

If `office-tools` is not available, ask the user to install it first and confirm the command is on `PATH`:

```powershell
office-tools --help
```

## Commands

| Task | Command |
| --- | --- |
| Convert TXT to DOCX | `txt-to-docx <file-or-dir>` |
| Convert DOCX to PDF | `docx-to-pdf <file-or-dir>` |
| Convert PDF to DOCX | `pdf-to-docx <file-or-dir>` |
| Render PDF pages to PNG images | `pdf-to-images <file-or-dir>` |
| Merge all PDFs in one directory | `merge-pdfs <dir>` |
| Add image watermark to DOCX | `watermark-docx <file-or-dir> --image <image-path> [--watermark-left <cm>] [--watermark-top <cm>] [--watermark-width <cm>] [--watermark-height <cm>]` |
| Add image watermark to PDF | `watermark-pdf <file-or-dir> --image <image-path> [--watermark-left <cm>] [--watermark-top <cm>] [--watermark-width <cm>] [--watermark-height <cm>]` |

Watermark position and size options use centimeter units. DOCX and PDF defaults are `--watermark-left 4`, `--watermark-top 6`, `--watermark-width 4`. If `--watermark-height` is not passed, height scales proportionally from the image aspect ratio.

## Path Handling

- Default `--mode auto` treats existing directories as directory input and everything else as file input.
- Use `--mode file` or `--mode dir` when the intent must be explicit.
- Directory mode recursively processes matching files and skips temporary Office files whose names start with `~$`.
- `merge-pdfs` only accepts a directory.
- `txt-to-docx` processes `.txt` files and writes same-name `.docx` files.

## Output

With `--json`, the command prints one JSON object:

```json
{
  "ok": true,
  "command": "docx-to-pdf",
  "mode": "dir",
  "input": "D:\\docs",
  "summary": { "total": 2, "ok": 2, "failed": 0, "elapsed_seconds": 0.12 },
  "items": [
    { "path": "D:\\docs\\a.docx", "ok": true, "status": "created", "output": "D:\\docs\\a.pdf" }
  ]
}
```

Treat process exit code `0` as success and non-zero as failure. Also inspect `summary.failed` and per-item `ok` because directory operations may partly fail.

## Examples

```powershell
office-tools docx-to-pdf "D:\docs" --mode dir --json
office-tools txt-to-docx "D:\docs" --mode dir --json
office-tools pdf-to-images "D:\docs\a.pdf" --json
office-tools watermark-docx "D:\docs\a.docx" --image "D:\logo.png" --json
office-tools watermark-pdf "D:\docs\a.pdf" --image "D:\logo.png" --json
office-tools merge-pdfs "D:\pdfs" --json
```

## Caveats

- `docx-to-pdf` and `watermark-docx` use Windows COM automation and require Microsoft Word or a compatible WPS/Word COM environment.
- Make sure target files are not open in Word/WPS before modifying or converting them.
- `pdf-to-docx` quality depends on the source PDF structure.
- `pdf-to-images` overwrites the derived output directory if it already exists.
- `txt-to-docx` reads UTF-8 TXT, removes empty lines, and converts obvious tab, pipe, or multi-space-delimited blocks into tables.
- **Error handling**: If a command fails or raises an error, stop immediately and report the error to the user. Do NOT attempt to diagnose or fix the issue yourself.
