from PyPDF2 import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def select_pdf_file():
    """选择PDF文件"""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="选择要分割的PDF文件",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )


def get_page_range(total_pages):
    """获取用户输入的页码范围（从1开始计数）"""
    root = tk.Tk()
    root.withdraw()

    # 输入起始页
    start = simpledialog.askinteger(
        "输入起始页",
        f"请输入起始页码（1-{total_pages}）:",
        minvalue=1,
        maxvalue=total_pages
    )
    if start is None: return None, None  # 用户取消输入

    # 输入结束页
    end = simpledialog.askinteger(
        "输入结束页",
        f"请输入结束页码（{start}-{total_pages}）:",
        minvalue=start,
        maxvalue=total_pages
    )
    return start, end


def split_pdf(input_path, output_path, start_page, end_page):
    """将指定页码保存为单个PDF"""
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()

        # 注意：PyPDF2的页码从0开始，用户输入从1开始
        for page_num in range(start_page - 1, end_page):
            writer.add_page(reader.pages[page_num])

        with open(output_path, "wb") as f:
            writer.write(f)
        return True
    except Exception as e:
        messagebox.showerror("错误", f"处理失败：{str(e)}")
        return False


if __name__ == "__main__":
    # 选择文件
    input_pdf = select_pdf_file()
    if not input_pdf:
        messagebox.showwarning("取消", "未选择文件")
        exit()

    # 读取总页数
    try:
        total_pages = len(PdfReader(input_pdf).pages)
    except Exception as e:
        messagebox.showerror("错误", f"无法读取PDF：{str(e)}")
        exit()

    # 获取页码范围
    start, end = get_page_range(total_pages)
    if not start or not end:
        messagebox.showwarning("取消", "已取消操作")
        exit()

    # 选择保存路径
    output_pdf = filedialog.asksaveasfilename(
        title="保存分割后的PDF",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not output_pdf:
        messagebox.showwarning("取消", "已取消保存")
        exit()

    # 执行分割
    if split_pdf(input_pdf, output_pdf, start, end):
        messagebox.showinfo("成功", f"已保存：{output_pdf}\n页码范围：{start}-{end}")