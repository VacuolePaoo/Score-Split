import os
import warnings
from datetime import datetime
from openpyxl import load_workbook, Workbook
from prompt_toolkit import prompt
from prompt_toolkit.application import Application, get_app
from prompt_toolkit.key_binding import KeyBindings
from prompt_toolkit.layout import Layout
from prompt_toolkit.widgets import CheckboxList, Button, Label, Box, Frame, RadioList, TextArea
from prompt_toolkit.layout.containers import HSplit, VSplit, Window
from prompt_toolkit.styles import Style
# 移除了WordCompleter的导入，因为我们不再需要自动补全功能

# 导入拆分后的工具模块
from utils.directory_utils import choose_working_directory, list_excel_files
from utils.file_selection_utils import check_output_dir, choose_files
from utils.sheet_utils import list_all_sheets, choose_sheet
from utils.user_input_utils import ask_number, choose_class_column
from utils.split_utils import split_and_save

# 忽略所有警告
warnings.filterwarnings("ignore")


def main():
    # 步骤1: 选择工作目录
    working_dir, mode = choose_working_directory()
    if mode == "exit":
        print("程序已退出。")
        return
    if not working_dir:
        print("未选择有效的工作目录，程序退出。")
        return
    
    # 清屏并显示工作目录
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"工作目录: {os.path.abspath(working_dir)}")
    
    # 步骤2: 检查输出目录
    if not check_output_dir(working_dir):
        return

    # 步骤3: 扫描指定目录，获取所有xlsx文件
    # 清屏并显示文件列表
    os.system('cls' if os.name == 'nt' else 'clear')
    files = list_excel_files(working_dir)
    print("找到以下Excel文件:")
    for f in files:
        print(f"  - {f}")
    
    
    # 步骤4: 让用户选择要处理的文件
    selected = choose_files(files)
    if selected == "exit":
        print("程序已退出。")
        return
    if not selected:
        print("未选择文件，退出。")
        return

    # 步骤5: 获取第一个文件的sheet列表并让用户选择
    # 清屏并显示sheet列表
    os.system('cls' if os.name == 'nt' else 'clear')
    first_file = os.path.join(working_dir, selected[0])
    sheets = list_all_sheets(first_file)
    sheet_index, sheet_name = choose_sheet(sheets)
    
    if sheet_index == "exit":
        print("程序已退出。")
        return
    if sheet_index is None:
        print("未选择sheet，退出。")
        return

    # 步骤6: 询问用户表头所在行数
    # 清屏并询问表头行数
    os.system('cls' if os.name == 'nt' else 'clear')
    header_row = ask_number("表头所在行号: ")
    if header_row == 'exit':
        print("程序已退出。")
        return
    
    # 步骤7: 检索表头行内容并询问用户班级列所在列数
    # 清屏并显示表头内容
    os.system('cls' if os.name == 'nt' else 'clear')
    # 显示第一个文件的表头作为示例
    wb = load_workbook(first_file, read_only=True)
    ws = wb[sheet_name]
    header_row_data = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    wb.close()
    
    print(f"表头行（第{header_row}行）的前10个单元格内容:")
    for i, cell in enumerate(header_row_data[:10]):
        print(f"  列{i+1}: {cell}")
    
    # 使用单选方式选择班级列
    class_col = choose_class_column(header_row_data)
    if class_col == "exit":
        print("程序已退出。")
        return
    if class_col is None:
        print("未选择班级列，退出。")
        return

    # 步骤8-10: 拆分并保存文件
    # 清屏并开始处理文件
    os.system('cls' if os.name == 'nt' else 'clear')
    print("开始拆分文件...")
    split_and_save(selected, sheet_index, sheet_name, header_row, class_col, working_dir)
    print("拆分完成，结果保存在 '拆分' 文件夹中。")


if __name__ == "__main__":
    main()

