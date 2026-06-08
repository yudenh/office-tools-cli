import os
import fitz  # PyMuPDF 库
from pdf2docx import Converter
from modules import file_utils
from modules.constants import HandleResult
import shutil
import tempfile

# pip install PyMuPDF
# pip install pdf2docx

POINTS_PER_CM = 72 / 2.54
DEFAULT_WATERMARK_LEFT = 4 * POINTS_PER_CM
DEFAULT_WATERMARK_TOP = 6 * POINTS_PER_CM
DEFAULT_WATERMARK_WIDTH = 4 * POINTS_PER_CM


def pdf_to_images(pdf_path, output_dir):
    """
    将 PDF 的每一页转换为图片并保存到指定文件夹。

    :param pdf_path: 输入的 PDF 文件路径
    :param output_dir: 输出图片的文件夹路径
    :return: 输出图片的文件夹路径
    """
    # 如果文件名不以.pdf 结尾，则跳过
    if not pdf_path.endswith(".pdf"):
        return HandleResult.Skiped
    # 如果目录已存在，则强制删除
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    # 创建输出文件夹
    os.makedirs(output_dir)
    # 打开 PDF 文件
    pdf_document = fitz.open(pdf_path)
    # 遍历 PDF 的每一页
    for page_number in range(pdf_document.page_count):
        # 获取当前页
        page = pdf_document.load_page(page_number)
        # 设置图片分辨率为 300 DPI
        zoom = 300 / 72  # 72 是 PDF 的默认 DPI
        # 计算缩放后的页面大小
        mat = fitz.Matrix(zoom, zoom)
        # 将页面渲染为图像
        pix = page.get_pixmap(matrix=mat)
        # 生成图片保存路径
        image_path = f"{output_dir}/page_{page_number + 1}.png"
        # 保存图像
        pix.save(image_path)
    # 关闭 PDF 文件
    pdf_document.close()
    return HandleResult.Created


def pdf_to_docx(pdf_file, docx_file):
    """
    将 PDF 文件转换为 DOCX 文件。

    :param pdf_file: 输入的 PDF 文件路径
    :param docx_file: 输出的 DOCX 文件路径
    """
    result = HandleResult.Created
    if os.path.exists(docx_file):
        if file_utils.compare_file_modify_time(pdf_file, docx_file) == 1:
            # 源文件比目标文件新，则删除目标文件
            os.remove(docx_file)
            result = HandleResult.Overrided
        else:
            # 源文件比目标文件旧，则跳过
            result = HandleResult.Skiped
            return result
    # 创建 Converter 对象
    cv = Converter(pdf_file)
    # 执行转换
    cv.convert(docx_file)
    # 关闭 Converter 对象
    cv.close()
    return result


def add_image_watermark(
    pdf_path,
    image_path,
    output_path=None,
    left=DEFAULT_WATERMARK_LEFT,
    top=DEFAULT_WATERMARK_TOP,
    width=DEFAULT_WATERMARK_WIDTH,
    height=None,
):
    """
    为指定 PDF 文件的每一页添加图片水印。

    :param pdf_path: 要处理的 PDF 文件路径
    :param image_path: 水印图片路径
    :param output_path: 输出 PDF 路径。为 None 时覆盖原文件
    :param left: 水印图片距离页面左侧的位置，单位为 point
    :param top: 水印图片距离页面顶部的位置，单位为 point
    :param width: 水印图片宽度，单位为 point
    :param height: 水印图片高度，单位为 point。为 None 时按宽度等比自适应
    """
    if not pdf_path.endswith(".pdf"):
        return HandleResult.Skiped

    result = HandleResult.Processed
    target_path = output_path or pdf_path
    if output_path:
        result = HandleResult.Created
        if os.path.exists(output_path):
            if file_utils.compare_file_modify_time(pdf_path, output_path) == 1:
                os.remove(output_path)
                result = HandleResult.Overrided
            else:
                return HandleResult.Skiped

    image_document = fitz.open(image_path)
    try:
        image_rect = image_document[0].rect
        if height is None:
            height = width * image_rect.height / image_rect.width
    finally:
        image_document.close()

    pdf_document = fitz.open(pdf_path)
    temp_path = None
    try:
        watermark_rect = fitz.Rect(left, top, left + width, top + height)
        for page in pdf_document:
            page.insert_image(watermark_rect, filename=image_path, overlay=True)

        if output_path:
            pdf_document.save(target_path)
        else:
            temp_file = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".pdf",
                dir=os.path.dirname(os.path.abspath(pdf_path)),
            )
            temp_path = temp_file.name
            temp_file.close()
            pdf_document.save(temp_path)
    finally:
        pdf_document.close()

    if temp_path:
        os.replace(temp_path, pdf_path)

    return result


def merge_pdfs(directory):
    """
    将指定目录下的所有PDF文件合并为一个PDF文件。

    :param directory: 包含PDF文件的目录路径
    :return: 合并后的PDF文件路径
    """
    dir_name = os.path.basename(directory)
    # 输出目录名
    output_dir_name = dir_name + "_merged"
    # 输出目录路径
    output_dir = os.path.join(directory, output_dir_name)
    # 检查输出目录是否存在，如果存在则删除
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    # 创建输出目录
    os.makedirs(output_dir)

    # 创建一个新的PDF文档对象
    doc = fitz.open()

    # 获取目录下所有PDF文件
    pdf_files = [
        os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(".pdf")
    ]

    # 按文件名排序
    pdf_files.sort()

    # 遍历所有PDF文件并将它们合并到新的PDF文档中
    for pdf_file in pdf_files:
        # 打开每个PDF文件
        pdf = fitz.open(pdf_file)
        doc.insert_pdf(pdf, from_page=0, to_page=-1)  # 将整个页插入到文档末尾
        pdf.close()  # 关闭已读取的PDF文件以节省资源

    # 输出文件名
    fname = os.path.basename(directory) + ".pdf"
    # 输出文件路径
    out_filepath = os.path.join(output_dir, fname)

    # 保存合并后的文档
    doc.save(out_filepath)
    doc.close()

    return out_filepath
