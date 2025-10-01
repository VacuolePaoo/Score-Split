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


def ask_student_id_column(header_row_data):
    """
    询问用户学号所在列，如果用户输入数字则忽略该列
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print("学号列忽略选项:")
    print("表头行内容:")
    for i, cell in enumerate(header_row_data[:10]):
        print(f"  列{i+1}: {cell}")
    
    while True:
        user_input = prompt("请输入学号所在列号 (直接按回车表示不忽略任何列): ").strip()
        if user_input.lower() == 'exit':
            return 'exit'
        if user_input == "":
            return None
        try:
            col_num = int(user_input)
            if 1 <= col_num <= len(header_row_data):
                return col_num
            else:
                print(f"列号必须在1到{len(header_row_data)}之间，请重新输入！")
        except ValueError:
            print("输入无效，请输入一个有效的数字或直接按回车！")


def ask_ignore_class_column(header_row_data, class_col):
    """
    询问用户是否忽略班级列，使用按钮交互方式
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    def on_yes():
        get_app().exit(result=True)
    
    def on_no():
        get_app().exit(result=False)
    
    def on_exit():
        get_app().exit(result="exit")
    
    # 创建按钮
    btn_yes = Button(text="是，忽略班级列", handler=on_yes)
    btn_no = Button(text="否，保留班级列", handler=on_no)
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建界面布局
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
    body = HSplit([
        Label("班级列忽略选项:", dont_extend_height=True),
        Window(height=1, char="-"),
        Label(f"已选择的班级列为: 列{class_col}: {header_row_data[class_col-1]}", dont_extend_height=True),
        Window(height=1, char=" "),
        Label("是否忽略班级列?", dont_extend_height=True),
        Window(height=1, char="-"),
        VSplit([btn_yes, btn_no], padding=3),
        Window(height=1, char="-"),
        btn_exit,
    ])
    
    application = Application(
        layout=Layout(body),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    return application.run()


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

    # 新增步骤: 询问是否忽略学号列
    student_id_col = ask_student_id_column(header_row_data)
    if student_id_col == "exit":
        print("程序已退出。")
        return

    # 新增步骤: 询问是否忽略班级列
    ignore_class_col = ask_ignore_class_column(header_row_data, class_col)
    if ignore_class_col == "exit":
        print("程序已退出。")
        return

    # 步骤8-10: 拆分并保存文件
    # 清屏并开始处理文件
    os.system('cls' if os.name == 'nt' else 'clear')
    print("开始拆分文件...")
    split_and_save(selected, sheet_index, sheet_name, header_row, class_col, working_dir, student_id_col, ignore_class_col)
    print("拆分完成，结果保存在 '拆分' 文件夹中。")


if __name__ == "__main__":
    main()