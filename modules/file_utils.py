import os


def filter_files(dir_path, surfix=None, exclude_prefix=None):
    """
    递归遍历当前工作目录下的所有文件，根据指定的文件后缀和要排除的前缀筛选文件。

    :param dir_path: 要遍历的目录路径。
    :param file_extension: 要筛选的文件后缀，如 '.txt'。如果为 None，则不进行后缀筛选。
    :param exclude_prefix: 要排除的文件名前缀，如 'temp_'。如果为 None，则不进行前缀排除。
    :return: 包含符合条件的文件名和文件完整路径的元组列表。
    """
    result = []
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            # 检查文件后缀
            if surfix and not file.endswith(surfix):
                continue
            # 检查文件前缀
            if exclude_prefix and file.startswith(exclude_prefix):
                continue
            file_path = os.path.join(root, file)
            result.append((file, file_path))
    return result


def compare_file_modify_time(file1, file2):
    """
    比较两个文件的修改时间。

    :param file1: 第一个文件的路径。
    :param file2: 第二个文件的路径。
    :return: 如果 file1 比 file2 新，返回 1；如果 file1 比 file2 旧，返回 -1；如果两者相同，返回 0。
    """
    # 获取文件的修改时间
    mtime1 = os.path.getmtime(file1)
    mtime2 = os.path.getmtime(file2)

    # 比较修改时间
    if mtime1 > mtime2:
        return 1
    elif mtime1 < mtime2:
        return -1
    else:
        return 0
