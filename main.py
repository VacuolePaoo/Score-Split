import os
import warnings
import platform
import json
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


def load_config():
    """
    加载配置文件
    """
    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"配置文件读取失败: {e}")
            return {"configs": []}
    else:
        # 如果配置文件不存在，创建一个默认的
        default_config = {
            "configs": [
                {
                    "name": "默认配置",
                    "sheet_index": 0,
                    "class_column": 2,
                    "header_row": 1,
                    "student_id_column": None,
                    "ignore_class_column": False
                }
            ]
        }
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"创建默认配置文件失败: {e}")
        return default_config


def choose_config(config_data):
    """
    让用户选择使用已有配置还是自定义配置
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    configs = config_data.get("configs", [])
    
    if not configs:
        return None
    
    def on_custom():
        get_app().exit(result="custom")
    
    def on_exit():
        get_app().exit(result="exit")
    
    # 创建配置选项
    config_buttons = []
    for i, config in enumerate(configs):
        def make_handler(index):
            def handler():
                get_app().exit(result=index)
            return handler
        
        btn = Button(text=config["name"], handler=make_handler(i))
        config_buttons.append(btn)
    
    # 添加自定义和退出按钮
    btn_custom = Button(text="自定义配置", handler=on_custom)
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建界面布局
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
    body_elements = [
        Label("选择配置:", dont_extend_height=True),
        Window(height=1, char="-"),
        Label("请选择要使用的配置:", dont_extend_height=True),
    ]
    
    # 添加配置按钮（每行一个）
    for btn in config_buttons:
        body_elements.append(btn)
    
    body_elements.extend([
        Window(height=1, char="-"),
        btn_custom,
        btn_exit,
    ])
    
    body = HSplit(body_elements)
    
    application = Application(
        layout=Layout(body),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    return application.run()


def ask_student_id_column(header_row_data):
    """
    询问用户学号所在列，如果用户选择则忽略该列，使用按钮交互方式
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    # 创建列选项
    values = []
    for i, cell in enumerate(header_row_data[:10]):
        values.append((i+1, f"列{i+1}: {cell}"))
    
    # 创建单选列表
    radio_list = RadioList(values=values)
    radio_list.current_value = 1  # 默认选择第一列
    
    def on_confirm():
        get_app().exit(result=radio_list.current_value)
    
    def on_skip():
        get_app().exit(result=None)
    
    def on_exit():
        get_app().exit(result="exit")
    
    # 创建按钮
    btn_confirm = Button(text="确认选择此列为学号列", handler=on_confirm)
    btn_skip = Button(text="不忽略任何列", handler=on_skip)
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建界面布局
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
    body = HSplit([
        Label("学号列忽略选项:", dont_extend_height=True),
        Window(height=1, char="-"),
        Label("表头行内容:", dont_extend_height=True),
        Box(body=radio_list, padding=1),
        Window(height=1, char="-"),
        VSplit([btn_confirm, btn_skip], padding=3),
        Window(height=1, char="-"),
        btn_exit,
    ])
    
    application = Application(
        layout=Layout(body, focused_element=radio_list),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    return application.run()


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


def show_completion_options(working_dir, stats):
    """
    显示处理完成后的操作选项和统计信息
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    output_dir = os.path.join(working_dir, "拆分")
    
    def on_open_folder():
        # 打开输出文件夹
        system = platform.system()
        if system == "Windows":
            os.startfile(output_dir)
        elif system == "Darwin":  # macOS
            os.system(f"open '{output_dir}'")
        else:  # Linux
            os.system(f"xdg-open '{output_dir}'")
        get_app().exit(result="open")
    
    def on_exit():
        get_app().exit(result="exit")
    
    # 创建按钮
    btn_open = Button(text="打开输出文件夹", handler=on_open_folder)
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建界面布局
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
    # 统计信息显示
    stats_labels = [
        Label("处理完成!", dont_extend_height=True),
        Window(height=1, char="-"),
        Label(f"结果已保存在: {output_dir}", dont_extend_height=True),
        Window(height=1, char="="),
        Label("处理统计信息:", dont_extend_height=True),
        Label(f"  处理文件数: {stats['processed_files']}/{stats['total_files']}", dont_extend_height=True),
        Label(f"  跳过文件数: {stats['skipped_files']}", dont_extend_height=True),
        Label(f"  生成班级数: {stats['generated_classes']}", dont_extend_height=True),
        Label(f"  处理数据行数: {stats['total_rows']}", dont_extend_height=True),
        Window(height=1, char="="),
        Label("请选择操作:", dont_extend_height=True),
        Window(height=1, char="-"),
        VSplit([btn_open, btn_exit], padding=3),
    ]
    
    body = HSplit(stats_labels)
    
    application = Application(
        layout=Layout(body),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    return application.run()


def main():
    # 加载配置文件
    config_data = load_config()
    
    # 选择配置
    config_choice = choose_config(config_data)
    preset_config = None
    if config_choice == "exit":
        print("程序已退出。")
        return
    elif config_choice == "custom":
        # 使用自定义配置流程（原有流程）
        pass
    elif isinstance(config_choice, int):
        # 使用预配置
        configs = config_data.get("configs", [])
        if 0 <= config_choice < len(configs):
            preset_config = configs[config_choice]
            print(f"使用预配置: {preset_config['name']}")
        else:
            print("配置选择无效，使用自定义配置。")
    
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
    
    # 如果使用预配置，直接使用预配置的sheet索引
    if preset_config:
        sheet_index = preset_config["sheet_index"]  # sheet索引（从0开始）
        sheet_name = sheets[sheet_index] if sheet_index < len(sheets) else sheets[0]  # 对应的sheet名称
        print(f"使用预配置的sheet: {sheet_name}")
    else:
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
    # 如果使用预配置，直接使用预配置的表头行
    if preset_config:
        header_row = preset_config["header_row"]
        print(f"使用预配置的表头行: {header_row}")
    else:
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
    # 如果使用预配置，直接使用预配置的班级列
    if preset_config:
        class_col = preset_config["class_column"]
        print(f"使用预配置的班级列: {class_col}")
    else:
        class_col = choose_class_column(header_row_data)
        if class_col == "exit":
            print("程序已退出。")
            return
        if class_col is None:
            print("未选择班级列，退出。")
            return

    # 新增步骤: 询问是否忽略学号列
    # 如果使用预配置，直接使用预配置的学号列设置
    if preset_config:
        student_id_col = preset_config["student_id_column"]
        print(f"使用预配置的学号列设置: {student_id_col}")
    else:
        student_id_col = ask_student_id_column(header_row_data)
        if student_id_col == "exit":
            print("程序已退出。")
            return

    # 新增步骤: 询问是否忽略班级列
    # 如果使用预配置，直接使用预配置的忽略班级列设置
    if preset_config:
        ignore_class_col = preset_config["ignore_class_column"]
        print(f"使用预配置的忽略班级列设置: {ignore_class_col}")
    else:
        ignore_class_col = ask_ignore_class_column(header_row_data, class_col)
        if ignore_class_col == "exit":
            print("程序已退出。")
            return

    # 步骤8-10: 拆分并保存文件
    # 清屏并开始处理文件
    os.system('cls' if os.name == 'nt' else 'clear')
    print("开始拆分文件...")
    stats = split_and_save(selected, sheet_index, sheet_name, header_row, class_col, working_dir, student_id_col, ignore_class_col)
    
    # 显示处理完成后的选项和统计信息
    result = show_completion_options(working_dir, stats)
    if result == "open":
        print("正在打开输出文件夹...")
    elif result == "exit":
        print("程序已退出。")


if __name__ == "__main__":
    main()