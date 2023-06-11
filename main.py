from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter


# 转成报错信息
def str_to_error(message):
    return f"\033[0;31m[ERROR]{message}\033[0m"


# 软报错
def sort_raise(message):
    print(str_to_error(message))
    exit(1)


# 输入PDF文件路径
def input_pdf_path(message):
    # 输入PDF文件路径
    pdf_path = Path(input(message))

    # 判断
    if pdf_path.is_file() and pdf_path.suffix.lower() in [".pdf"]:
        return pdf_path

    sort_raise("填写的路径并非PDF文件")


# 输入整数
def input_integer(message):
    # 从键盘获取值
    value = input(message)

    try:
        # 强制类型转换
        value = int(value)
        return value
    except ValueError:
        sort_raise("输入的数字非整数")


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


def cut():
    # 输入PDF文件路径
    pdf_path = input_pdf_path("请输入PDF文件路径: ")

    # 读取PDF文件
    pdf_file = pdf_path.open("rb")
    pdf_reader = PdfReader(pdf_file)

    # 输入分割页码
    page_begin = input_integer("请输入需要裁切的起始页码: ")
    page_end = input_integer("请输入需要裁切的结尾页码(包括): ")

    # 检查
    if not check_page(len(pdf_reader.pages), page_begin, page_end):
        sort_raise("页码输入不规范")

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
        sort_raise(ax)
    finally:
        pdf_file.close()


# 合并
def merge():
    pdf_path1 = input_pdf_path("请输入前置PDF文件路径: ")
    pdf_path2 = input_pdf_path("请输入后置PDF文件路径: ")
    out_path = Path(input(r"请输入合成的新PDF文件路径(例如: D:\new.pdf): "))

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
        sort_raise(ax)
    finally:
        pdf_file1.close()
        pdf_file2.close()


def _main():
    while True:
        print("-------------------")
        print("1、裁切(cut)")
        print("2、合并(merge)")
        print("3、退出")
        num = input_integer("输入序号(1-3): ")
        if num < 1 or num > 3:
            sort_raise("序号不正确")
        elif num == 1:
            cut()
        elif num == 2:
            merge()
        else:
            break


if __name__ == '__main__':
    _main()
