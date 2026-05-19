import os
import fitz  # PyMuPDF 库
from pdf2docx import Converter
from modules import file_utils
from modules.constants import HandleResult
import shutil

# pip install PyMuPDF
# pip install pdf2docx


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
