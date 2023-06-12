import argparse
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter


# 转成报错信息
def str_to_error(message):
    return f"\033[0;31m[ERROR]{message}\033[0m"


# 软报错
def soft_raise(message):
    print(str_to_error(message))
    input(str_to_error("请输入回车键退出..."))
    exit(1)


# 检查参数是否为None
def check_arg(arg, message):
    if arg is None:
        soft_raise(message)

    return arg


# 检查页数
def check_page(page_num, page_begin, page_end=None):
    # 判断起始页
    if page_begin < 1 or page_begin > page_num:
        return False

    if page_end is not None:
        if page_end < 1 or page_end > page_num or page_begin >= page_end:
            return False
        else:
            return True
    else:
        return True


# 检查PDF路径
def check_pdf_path(pdf_path):
    # 输入PDF文件路径
    pdf_path = Path(pdf_path)

    # 判断
    if pdf_path.is_file() and pdf_path.suffix.lower() in [".pdf"]:
        return pdf_path

    soft_raise("填写的路径并非PDF文件")


# 输入PDF文件路径
def input_pdf_path(message):
    # 对输入的路径信息进行预处理
    inputs = input(message)
    inputs = inputs.strip("\u202a").strip("\n")

    # 返回
    return check_pdf_path(inputs)


# 输入整数
def input_integer(message):
    # 从键盘获取值
    value = input(message)

    try:
        # 强制类型转换
        value = int(value)
        return value
    except ValueError:
        soft_raise("输入的数字非整数")


# 裁切
def cut(args=None):
    # 输入PDF文件路径
    if args is None:
        pdf_path = input_pdf_path("请输入PDF文件路径: ")
    else:
        pdf_path = check_pdf_path(check_arg(args.cut_pdf_path, "PDF路径参数不存在"))

    # 读取PDF文件
    pdf_file = pdf_path.open("rb")
    pdf_reader = PdfReader(pdf_file)

    # 输入分割页码
    page_begin = input_integer("请输入需要裁切的起始页码: ") if args is None else check_arg(args.page_begin, "起始页码参数不存在")
    page_end = input_integer("请输入需要裁切的结尾页码(包括): ") if args is None else check_arg(args.page_end, "结束页码参数不存在")

    # 检查
    if not check_page(len(pdf_reader.pages), page_begin, page_end):
        soft_raise("页码输入不规范")

    # 写入
    pdf_writer = PdfWriter()

    # 获取对应页内容
    for i in range(page_begin - 1, page_end):
        pdf_writer.add_page(pdf_reader.pages[i])

    try:
        # 导出文件名
        out_path = pdf_path.parent.joinpath(f"{pdf_path.name[:-4]}_new.pdf")

        # 写入
        with out_path.open("wb") as file:
            pdf_writer.write(file)

        print(f"裁切结果保存至: {str(out_path)}")
    except Exception as ax:
        soft_raise(ax)
    finally:
        pdf_file.close()


# 合并
def merge(args=None):
    if args is None:
        pdf_path1 = input_pdf_path("请输入前置PDF文件路径: ")
        pdf_path2 = input_pdf_path("请输入后置PDF文件路径: ")
        out_path = Path(input(r"请输入合成的新PDF文件路径(例如: D:\new.pdf): "))
    else:
        pdf_path1 = check_pdf_path(check_arg(args.merge_pdf_path1, "合并PDF路径1参数不存在"))
        pdf_path2 = check_pdf_path(check_arg(args.merge_pdf_path2, "合并PDF路径2参数不存在"))
        out_path = check_arg(args.merge_out_path, "合并导出路径参数不存在")

    # 读取
    pdf_file1 = pdf_path1.open("rb")
    pdf_file2 = pdf_path2.open("rb")
    pdf_reader1 = PdfReader(pdf_file1)
    pdf_reader2 = PdfReader(pdf_file2)

    # 写入
    pdf_writer = PdfWriter()

    # 写入
    try:
        for page in pdf_reader1.pages:
            pdf_writer.add_page(page)

        for page in pdf_reader2.pages:
            pdf_writer.add_page(page)

        with out_path.open("wb") as file:
            pdf_writer.write(file)

        print(f"合并结果保存至: {str(out_path)}")
    except Exception as ax:
        soft_raise(ax)
    finally:
        pdf_file1.close()
        pdf_file2.close()


# 从命令行获取的参数
def get_args():
    # 创建命令行参数对象
    parse = argparse.ArgumentParser(description="输入参数以便程序快速运行")

    # 参数列表
    arguments_dict = {
        "function": [int, None],
        "cut_pdf_path": [str, None],
        "page_begin": [int, None],
        "page_end": [int, None],
        "merge_pdf_path1": [str, None],
        "merge_pdf_path2": [str, None],
        "merge_out_path": [str, None]
    }

    # 添加参数
    for key, value in arguments_dict.items():
        parse.add_argument(f"--{key}", nargs="?", type=value[0], default=value[1])

    # 获取参数
    parse_argument = parse.parse_args()

    # 遍历属性
    exist_arg = True
    attrs = arguments_dict.keys()
    for attr in attrs:
        if getattr(parse_argument, attr) is not None:
            break
    else:
        exist_arg = False

    return parse_argument, exist_arg


def _main():
    parse_argument, exist_arg = get_args()

    while True:
        # 判断是否存在参数
        if not exist_arg:
            print("-------------------")
            print("版本: v1.1")
            print("1、裁切(cut)")
            print("2、合并(merge)")
            print("3、退出")
            num = input_integer("输入序号(1-3): ")
        else:
            num = check_arg(parse_argument.function, "功能参数不存在")

        # 根据功能序号选择功能
        if num < 1 or num > 3:
            soft_raise("序号不正确")
        elif num == 1:
            cut() if not exist_arg else cut(parse_argument)
        elif num == 2:
            merge() if not exist_arg else merge(parse_argument)
        else:
            break

        # 如果参数存在则不继续循环
        if exist_arg:
            break


if __name__ == '__main__':
    _main()
