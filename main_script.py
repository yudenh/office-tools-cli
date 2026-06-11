import argparse
import json
import os
import sys
from modules import docx_utils, pdf_utils, file_utils
from datetime import datetime
from modules.constants import HandleResult

COMMANDS = {
    "txt-to-docx": {"index": 0, "suffix": ".txt", "target": "file_or_dir"},
    "docx-to-pdf": {"index": 1, "suffix": ".docx", "target": "file_or_dir"},
    "pdf-to-docx": {"index": 2, "suffix": ".pdf", "target": "file_or_dir"},
    "pdf-to-images": {"index": 3, "suffix": ".pdf", "target": "file_or_dir"},
    "merge-pdfs": {"index": 4, "suffix": ".pdf", "target": "dir"},
    "watermark-docx": {"index": 6, "suffix": ".docx", "target": "file_or_dir"},
    "watermark-pdf": {"index": 7, "suffix": ".pdf", "target": "file_or_dir"},
}

RESULT_NAMES = {
    HandleResult.Silent: "silent",
    HandleResult.Created: "created",
    HandleResult.Overrided: "overrided",
    HandleResult.Skiped: "skipped",
    HandleResult.Processed: "processed",
}

def main():
    return cli_main()


def build_parser():
    parser = argparse.ArgumentParser(
        description="Office 文件处理命令行工具，面向 AI 或脚本的非交互式调用。"
    )
    parser.add_argument(
        "command",
        choices=sorted(COMMANDS.keys()),
        help="要执行的操作。",
    )
    parser.add_argument(
        "path",
        help="输入文件或目录路径。目录模式会递归处理匹配后缀的文件。",
    )
    parser.add_argument(
        "--mode",
        choices=["auto", "file", "dir"],
        default="auto",
        help="输入路径类型。默认 auto 根据路径自动判断。",
    )
    parser.add_argument(
        "--output",
        help="单文件操作的输出文件或目录。未指定时使用默认派生路径。",
    )
    parser.add_argument(
        "--image",
        help="watermark-docx/watermark-pdf 使用的水印图片路径。",
    )
    parser.add_argument(
        "--watermark-left",
        type=float,
        help="水印图片水平位置，单位 cm。watermark-docx/watermark-pdf 默认 4。",
    )
    parser.add_argument(
        "--watermark-top",
        type=float,
        help="水印图片垂直位置，单位 cm。watermark-docx/watermark-pdf 默认 6。",
    )
    parser.add_argument(
        "--watermark-width",
        type=float,
        help="水印图片宽度，单位 cm。watermark-docx/watermark-pdf 默认 4。",
    )
    parser.add_argument(
        "--watermark-height",
        type=float,
        help="水印图片高度，单位 cm。未指定时按图片比例自适应。",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="输出单个 JSON 对象，便于 AI 或脚本解析。",
    )
    return parser


def default_output_path(command, file_path):
    if command == "txt-to-docx":
        return os.path.splitext(file_path)[0] + ".docx"
    if command == "docx-to-pdf":
        return os.path.splitext(file_path)[0] + ".pdf"
    if command == "pdf-to-docx":
        return os.path.splitext(file_path)[0] + ".docx"
    if command == "pdf-to-images":
        return os.path.splitext(file_path)[0] + ".pdf_imgs"
    return ""


def make_item(path, ok=True, status="", output="", error=""):
    item = {
        "path": path,
        "ok": ok,
        "status": status,
    }
    if output:
        item["output"] = output
    if error:
        item["error"] = error
    return item


def make_watermark_options(args):
    options = {}
    if args.watermark_left is not None:
        options["left"] = args.watermark_left
    if args.watermark_top is not None:
        options["top"] = args.watermark_top
    if args.watermark_width is not None:
        options["width"] = args.watermark_width
    if args.watermark_height is not None:
        options["height"] = args.watermark_height
    return options


def run_file_command(command, file_path, output=None, image=None, watermark_options=None):
    try:
        if command == "txt-to-docx":
            out_file = output or default_output_path(command, file_path)
            result = docx_utils.txt_to_docx(file_path, out_file)
            return make_item(file_path, status=RESULT_NAMES[result], output=out_file)
        if command == "docx-to-pdf":
            out_file = output or default_output_path(command, file_path)
            result = docx_utils.docx_to_pdf(file_path, out_file)
            return make_item(file_path, status=RESULT_NAMES[result], output=out_file)
        if command == "pdf-to-docx":
            out_file = output or default_output_path(command, file_path)
            result = pdf_utils.pdf_to_docx(file_path, out_file)
            return make_item(file_path, status=RESULT_NAMES[result], output=out_file)
        if command == "pdf-to-images":
            out_dir = output or default_output_path(command, file_path)
            result = pdf_utils.pdf_to_images(file_path, out_dir)
            return make_item(file_path, status=RESULT_NAMES[result], output=out_dir)
        if command == "watermark-docx":
            result = docx_utils.add_image_watermark(file_path, image, **(watermark_options or {}))
            return make_item(file_path, status=RESULT_NAMES[result])
        if command == "watermark-pdf":
            result = pdf_utils.add_image_watermark(
                file_path,
                image,
                output_path=output,
                **(watermark_options or {}),
            )
            return make_item(file_path, status=RESULT_NAMES[result], output=output or "")
        return make_item(file_path, ok=False, status="unsupported", error="Unsupported file command")
    except Exception as e:
        return make_item(file_path, ok=False, status="failed", error=str(e))


def run_command(args):
    spec = COMMANDS[args.command]
    mode = args.mode
    if mode == "auto":
        mode = "dir" if os.path.isdir(args.path) else "file"

    if spec["target"] == "dir" and mode != "dir":
        raise ValueError(f"{args.command} only accepts directory input")
    if args.command in ["watermark-docx", "watermark-pdf"] and not args.image:
        raise ValueError(f"{args.command} requires --image")

    start_time = datetime.now()
    items = []
    watermark_options = make_watermark_options(args)
    if args.command == "merge-pdfs":
        try:
            out_file = pdf_utils.merge_pdfs(args.path)
            items.append(make_item(args.path, status="created", output=out_file))
        except Exception as e:
            items.append(make_item(args.path, ok=False, status="failed", error=str(e)))
    elif mode == "dir":
        files = file_utils.filter_files(args.path, spec["suffix"], exclude_prefix="~$")
        for _, file_path in files:
            items.append(
                run_file_command(
                    args.command,
                    file_path,
                    image=args.image,
                    watermark_options=watermark_options,
                )
            )
    else:
        items.append(
            run_file_command(
                args.command,
                args.path,
                output=args.output,
                image=args.image,
                watermark_options=watermark_options,
            )
        )

    elapsed = round((datetime.now() - start_time).total_seconds(), 2)
    ok_count = len([item for item in items if item["ok"]])
    fail_count = len(items) - ok_count
    return {
        "ok": fail_count == 0,
        "command": args.command,
        "mode": mode,
        "input": args.path,
        "summary": {
            "total": len(items),
            "ok": ok_count,
            "failed": fail_count,
            "elapsed_seconds": elapsed,
        },
        "items": items,
    }


def print_text_result(result):
    print(f"command: {result['command']}")
    print(f"mode: {result['mode']}")
    print(
        f"summary: total={result['summary']['total']} ok={result['summary']['ok']} "
        f"failed={result['summary']['failed']} elapsed={result['summary']['elapsed_seconds']}s"
    )
    for item in result["items"]:
        line = f"{item['status']}: {item['path']}"
        if "output" in item:
            line += f" -> {item['output']}"
        if "error" in item:
            line += f" error={item['error']}"
        print(line)


def cli_main():
    parser = build_parser()
    args = parser.parse_args()
    try:
        result = run_command(args)
    except Exception as e:
        result = {
            "ok": False,
            "command": getattr(args, "command", ""),
            "mode": getattr(args, "mode", ""),
            "input": getattr(args, "path", ""),
            "summary": {"total": 0, "ok": 0, "failed": 1, "elapsed_seconds": 0},
            "items": [make_item(getattr(args, "path", ""), ok=False, status="failed", error=str(e))],
        }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_text_result(result)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
