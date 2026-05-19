import argparse
import json
import os
import sys
from modules import docx_utils, pdf_utils, file_utils
from datetime import datetime
from time import sleep
from modules.constants import HandleResult

suffix_list = [".docx", ".docx", ".pdf", ".pdf", ".pdf", ".docx", ".docx"]
func_names = [
    "清除word作者",
    "word转pdf",
    "pdf转word",
    "pdf转图片",
    "合并pdf",
    "在docx中查找",
    "word加水印",
]
func_codes = [x + 1 for x in range(len(func_names))]
handle_names = ["单个文件", "目录"]
handle_codes = [x + 1 for x in range(len(handle_names))]
split_line = "--------------------------------------------------"

COMMANDS = {
    "remove-author": {"index": 0, "suffix": ".docx", "target": "file_or_dir"},
    "docx-to-pdf": {"index": 1, "suffix": ".docx", "target": "file_or_dir"},
    "pdf-to-docx": {"index": 2, "suffix": ".pdf", "target": "file_or_dir"},
    "pdf-to-images": {"index": 3, "suffix": ".pdf", "target": "file_or_dir"},
    "merge-pdfs": {"index": 4, "suffix": ".pdf", "target": "dir"},
    "find-docx": {"index": 5, "suffix": ".docx", "target": "file_or_dir"},
    "watermark-docx": {"index": 6, "suffix": ".docx", "target": "file_or_dir"},
}

RESULT_NAMES = {
    HandleResult.Silent: "silent",
    HandleResult.Created: "created",
    HandleResult.Overrided: "overrided",
    HandleResult.Skiped: "skipped",
    HandleResult.Processed: "processed",
}


def make_menu_items():
    menu_items = []
    menu_items.append(f"序号\t{"功能".ljust(16)}\t{"处理方式".ljust(16)}")
    menu_items.append(split_line)
    for i in range(len(func_names)):
        # 合并功能只处理目录
        if func_names[i].startswith("合并"):
            menu_items.append(
                f"{func_codes[i]}{handle_codes[1]}\t{func_names[i].ljust(16)}\t{handle_names[1].ljust(16)}"
            )
            menu_items.append("")
            continue
        for j in range(len(handle_names)):
            menu_items.append(
                f"{func_codes[i]}{handle_codes[j]}\t{func_names[i].ljust(16)}\t{handle_names[j].ljust(16)}"
            )
        menu_items.append("")
    if menu_items[len(menu_items) - 1] == "":
        menu_items.pop()
    return menu_items


def make_all_codes():
    all_codes = []
    for i in range(len(func_codes)):
        for j in range(len(handle_codes)):
            all_codes.append(func_codes[i] * 10 + handle_codes[j])
    return all_codes


def show_menu(items):
    print("")
    print(split_line)
    for item in items:
        print(item)
    print(split_line)


def show_input_file_path():
    file_path = input("请输入文件(目录)路径：").strip()
    return file_path


def process_file(i, file_path):
    if i == 4:
        return

    search_string = ""
    image_path = ""
    if i == 5:
        search_string = input("请输入要查找的字符串(多个用空格分开): ").strip()
    if i == 6:
        image_path = input("请输入水印图片路径: ").strip()

    result = HandleResult.Silent
    out_file = ""
    out_dir = ""
    try:
        if i == 0:
            result = docx_utils.remove_author_info_from_docx(file_path)
        elif i == 1:
            out_file = file_path.replace(".docx", ".pdf")
            result = docx_utils.docx_to_pdf(file_path, out_file)
        elif i == 2:
            out_file = file_path.replace(".pdf", ".docx")
            result = pdf_utils.pdf_to_docx(file_path, out_file)
        elif i == 3:
            out_dir = file_path.replace(".pdf", ".pdf_imgs")
            result = pdf_utils.pdf_to_images(file_path, out_dir)
        elif i == 5:
            docx_utils.find_string_in_docx(file_path, search_string)
        elif i == 6:
            result = docx_utils.add_image_watermark(file_path, image_path)

        if result == HandleResult.Skiped:
            print("已跳过：" + file_path)
        if result == HandleResult.Overrided:
            print("已覆盖：" + out_file)
        if result == HandleResult.Created:
            print("已生成：" + out_file if out_file != "" else out_dir)
        if result == HandleResult.Processed:
            print("已处理：" + file_path)
    except Exception as e:
        print("!失败：" + file_path)
        print("请检查文件是否存在, 或者是否已经被word或wps打开")
        print(e)


def process_merge(i, dir_path):
    if i != 4:
        return
    start_time = datetime.now()
    try:
        out_file = pdf_utils.merge_pdfs(dir_path)
        print("已合并为：" + out_file)
    except Exception as e:
        print("!合并失败：" + dir_path)
        print(e)
    print("完成")
    end_ime = datetime.now()
    elapsed_time = end_ime - start_time
    elapsed_time.total_seconds()
    print(f"任务耗时：{round(elapsed_time.total_seconds(), 2)}秒")


def process_directory(i, suffix, dir_path):
    if i == 4:
        process_merge(i, dir_path)
        return
    file_list = file_utils.filter_files(dir_path, suffix, exclude_prefix="~$")
    if len(file_list) == 0:
        print("没有符合条件的文件")
        return

    search_string = ""
    image_path = ""
    if i == 5:
        search_string = input("请输入要查找的字符串(多个用空格分开): ").strip()
    if i == 6:
        image_path = input("请输入水印图片路径: ").strip()

    start_time = datetime.now()
    print("待处理文件个数：" + str(len(file_list)))
    print(split_line)
    success_cnt = 0
    fail_cnt = 0

    result = HandleResult.Silent
    out_file = ""
    out_dir = ""
    for file in file_list:
        file_path = file[1]
        try:
            if i == 0:
                result = docx_utils.remove_author_info_from_docx(file_path)
            elif i == 1:
                out_file = file_path.replace(".docx", ".pdf")
                result = docx_utils.docx_to_pdf(file_path, out_file)
            elif i == 2:
                out_file = file_path.replace(".pdf", ".docx")
                result = pdf_utils.pdf_to_docx(file_path, out_file)
            elif i == 3:
                out_dir = file_path.replace(".pdf", ".pdf_imgs")
                result = pdf_utils.pdf_to_images(file_path, out_dir)
            # i==4的情况已经在上面处理过了
            elif i == 5:
                docx_utils.find_string_in_docx(file_path, search_string)
            elif i == 6:
                result = docx_utils.add_image_watermark(file_path, image_path)

            success_cnt += 1

            if result == HandleResult.Skiped:
                print("已跳过：" + file_path)
            if result == HandleResult.Overrided:
                print("已覆盖：" + out_file)
            if result == HandleResult.Created:
                print("已生成：" + out_file if out_file != "" else out_dir)
            if result == HandleResult.Processed:
                print("已处理：" + file_path)
        except Exception as e:
            fail_cnt += 1
            print("!失败：" + file_path)
            print("请检查文件是否存在, 或者是否已经被word或wps打开")
            print(e)
    print(split_line)
    if i in [0, 1, 2]:
        print("成功个数：" + str(success_cnt))
        print("失败个数：" + str(fail_cnt))
    print("完成")
    end_ime = datetime.now()
    elapsed_time = end_ime - start_time
    elapsed_time.total_seconds()
    print(f"任务耗时：{round(elapsed_time.total_seconds(), 2)}秒")


def main():
    if len(sys.argv) > 1:
        return cli_main()
    menu_items = make_menu_items()
    all_codes = make_all_codes()
    while True:
        show_menu(menu_items)
        try:
            choice = int(input("请输入要执行的功能序号 (输入 0 退出): "))
            if choice == 0:
                print("退出程序。")
                break
            elif all_codes.index(choice) >= 0:
                file_path = show_input_file_path()
                func_index = choice // 10 - 1
                suffix_index = func_index
                handle_type = choice % 10
                if handle_type == 1:
                    process_file(func_index, file_path)
                elif handle_type == 2:
                    process_directory(func_index, suffix_list[suffix_index], file_path)
                sleep(1)
                input("按任意键继续...")
            else:
                print("输入的序号无效，请重新输入。")
        except ValueError:
            print("输入无效，请输入一个有效的整数序号。")
        except EOFError:
            print("退出程序。")
            break


def build_parser():
    parser = argparse.ArgumentParser(
        description="Office 文件处理命令行工具。无参数运行时进入交互式菜单。"
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
        "--query",
        help="find-docx 要查找的字符串。多个关键词用空格分隔。",
    )
    parser.add_argument(
        "--image",
        help="watermark-docx 使用的水印图片路径。",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="输出单个 JSON 对象，便于 AI 或脚本解析。",
    )
    return parser


def default_output_path(command, file_path):
    if command == "docx-to-pdf":
        return os.path.splitext(file_path)[0] + ".pdf"
    if command == "pdf-to-docx":
        return os.path.splitext(file_path)[0] + ".docx"
    if command == "pdf-to-images":
        return os.path.splitext(file_path)[0] + ".pdf_imgs"
    return ""


def make_item(path, ok=True, status="", output="", error="", matches=None):
    item = {
        "path": path,
        "ok": ok,
        "status": status,
    }
    if output:
        item["output"] = output
    if error:
        item["error"] = error
    if matches is not None:
        item["matches"] = matches
    return item


def run_file_command(command, file_path, output=None, query=None, image=None):
    try:
        if command == "remove-author":
            result = docx_utils.remove_author_info_from_docx(file_path)
            return make_item(file_path, status=RESULT_NAMES[result])
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
        if command == "find-docx":
            matches = docx_utils.find_string_in_docx(file_path, query or "", quiet=True)
            return make_item(file_path, status="matched" if matches else "not_found", matches=matches)
        if command == "watermark-docx":
            result = docx_utils.add_image_watermark(file_path, image)
            return make_item(file_path, status=RESULT_NAMES[result])
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
    if args.command == "find-docx" and not args.query:
        raise ValueError("find-docx requires --query")
    if args.command == "watermark-docx" and not args.image:
        raise ValueError("watermark-docx requires --image")

    start_time = datetime.now()
    items = []
    if args.command == "merge-pdfs":
        try:
            out_file = pdf_utils.merge_pdfs(args.path)
            items.append(make_item(args.path, status="created", output=out_file))
        except Exception as e:
            items.append(make_item(args.path, ok=False, status="failed", error=str(e)))
    elif mode == "dir":
        files = file_utils.filter_files(args.path, spec["suffix"], exclude_prefix="~$")
        for _, file_path in files:
            items.append(run_file_command(args.command, file_path, query=args.query, image=args.image))
    else:
        items.append(
            run_file_command(
                args.command,
                args.path,
                output=args.output,
                query=args.query,
                image=args.image,
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
        if "matches" in item:
            line += f" matches={len(item['matches'])}"
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
