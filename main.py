# -*- coding: utf-8 -*-

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

from utils.directory_utils import choose_working_directory, list_excel_files
from utils.file_selection_utils import check_output_dir, choose_files
from utils.sheet_utils import list_all_sheets, choose_sheet
from utils.user_input_utils import ask_number, choose_class_column
from utils.split_utils import split_and_save

warnings.filterwarnings("ignore")


def load_config():
    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"配置文件读取失败: {e}")
            return {"configs": []}
    else:
        default_config = {
            "configs": [
                {
                    "name": "默认配置（需要修改）",
                    "sheet_index": 0,
                    "class_column": 3,
                    "header_row": 2,
                    "student_id_column": None,
                    "ignore_class_column": False,
                    "existing_files_action": None,
                    "file_selection_mode": None,
                    "auto_detect_directory": False
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
    os.system('cls' if os.name == 'nt' else 'clear')
    
    configs = config_data.get("configs", [])
    
    if not configs:
        return None
    
    def on_custom():
        get_app().exit(result="custom")
    
    def on_exit():
        get_app().exit(result="exit")
    
    config_buttons = []
    for i, config in enumerate(configs):
        def make_handler(index):
            def handler():
                get_app().exit(result=index)
            return handler
        
        btn = Button(text=config["name"], handler=make_handler(i))
        config_buttons.append(btn)
    
    btn_custom = Button(text="自定义配置", handler=on_custom)
    btn_exit = Button(text="退出", handler=on_exit)
    
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
    body_elements = [
        Label("选择配置:", dont_extend_height=True),
        Window(height=1, char="-"),
        Label("请选择要使用的配置:", dont_extend_height=True),
    ]
    
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
    os.system('cls' if os.name == 'nt' else 'clear')
    
    values = []
    for i, cell in enumerate(header_row_data[:10]):
        values.append((i+1, f"列{i+1}: {cell}"))
    
    radio_list = RadioList(values=values)
    radio_list.current_value = 1
    
    def on_confirm():
        get_app().exit(result=radio_list.current_value)
    
    def on_skip():
        get_app().exit(result=None)
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_confirm = Button(text="确认选择此列为学号列", handler=on_confirm)
    btn_skip = Button(text="不忽略任何列", handler=on_skip)
    btn_exit = Button(text="退出", handler=on_exit)
    
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
    os.system('cls' if os.name == 'nt' else 'clear')
    
    def on_yes():
        get_app().exit(result=True)
    
    def on_no():
        get_app().exit(result=False)
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_yes = Button(text="是，忽略班级列", handler=on_yes)
    btn_no = Button(text="否，保留班级列", handler=on_no)
    btn_exit = Button(text="退出", handler=on_exit)
    
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
    os.system('cls' if os.name == 'nt' else 'clear')
    
    output_dir = os.path.join(working_dir, "拆分")
    
    def on_open_folder():
        system = platform.system()
        if system == "Windows":
            os.startfile(output_dir)
        elif system == "Darwin":
            os.system(f"open '{output_dir}'")
        else:
            os.system(f"xdg-open '{output_dir}'")
        get_app().exit(result="open")
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_open = Button(text="打开输出文件夹", handler=on_open_folder)
    btn_exit = Button(text="退出", handler=on_exit)
    
    style = Style.from_dict({
        "button.focused": "fg:ansiblue bg:ansiwhite",
    })
    
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
    config_data = load_config()
    
    config_choice = choose_config(config_data)
    preset_config = None
    if config_choice == "exit":
        print("程序已退出。")
        return
    elif config_choice == "custom":
        pass
    elif isinstance(config_choice, int):
        configs = config_data.get("configs", [])
        if 0 <= config_choice < len(configs):
            preset_config = configs[config_choice]
            print(f"使用预配置: {preset_config['name']}")
        else:
            print("配置选择无效，使用自定义配置。")
    
    # 获取预配置中的自动检测目录设置
    auto_detect_directory = False
    if preset_config and "auto_detect_directory" in preset_config:
        auto_detect_directory = preset_config["auto_detect_directory"]
    
    # 如果配置了自动检测目录，则直接使用当前目录
    if auto_detect_directory:
        working_dir = "."
        mode = "current"
        print(f"自动检测到工作目录: {os.path.abspath(working_dir)}")
    else:
        working_dir, mode = choose_working_directory()
        if mode == "exit":
            print("程序已退出。")
            return
        if not working_dir:
            print("未选择有效的工作目录，程序退出。")
            return
        
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"工作目录: {os.path.abspath(working_dir)}")
    
    # 获取预配置中的现有文件处理方式
    existing_files_action = None
    if preset_config and "existing_files_action" in preset_config:
        existing_files_action = preset_config["existing_files_action"]
    
    if not check_output_dir(working_dir, existing_files_action):
        return

    os.system('cls' if os.name == 'nt' else 'clear')
    files = list_excel_files(working_dir)
    print("找到以下Excel文件:")
    for f in files:
        print(f"  - {f}")
    
    # 获取预配置中的文件选择模式
    file_selection_mode = None
    if preset_config and "file_selection_mode" in preset_config:
        file_selection_mode = preset_config["file_selection_mode"]
    
    selected = choose_files(files, file_selection_mode)
    if selected == "exit":
        print("程序已退出。")
        return
    if not selected:
        print("未选择文件，退出。")
        return

    os.system('cls' if os.name == 'nt' else 'clear')
    first_file = os.path.join(working_dir, selected[0])
    sheets = list_all_sheets(first_file)
    
    if preset_config:
        sheet_index = preset_config["sheet_index"]
        sheet_name = sheets[sheet_index] if sheet_index < len(sheets) else sheets[0]
        print(f"使用预配置的sheet: {sheet_name}")
    else:
        sheet_index, sheet_name = choose_sheet(sheets)
        
        if sheet_index == "exit":
            print("程序已退出。")
            return
        if sheet_index is None:
            print("未选择sheet，退出。")
            return

    os.system('cls' if os.name == 'nt' else 'clear')
    if preset_config:
        header_row = preset_config["header_row"]
        print(f"使用预配置的表头行: {header_row}")
    else:
        header_row = ask_number("表头所在行号: ")
        if header_row == 'exit':
            print("程序已退出。")
            return
    
    os.system('cls' if os.name == 'nt' else 'clear')
    wb = load_workbook(first_file, read_only=True)
    ws = wb[sheet_name]
    header_row_data = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    wb.close()
    
    print(f"表头行（第{header_row}行）的前10个单元格内容:")
    for i, cell in enumerate(header_row_data[:10]):
        print(f"  列{i+1}: {cell}")
    
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

    if preset_config:
        student_id_col = preset_config["student_id_column"]
        print(f"使用预配置的学号列设置: {student_id_col}")
    else:
        student_id_col = ask_student_id_column(header_row_data)
        if student_id_col == "exit":
            print("程序已退出。")
            return

    if preset_config:
        ignore_class_col = preset_config["ignore_class_column"]
        print(f"使用预配置的忽略班级列设置: {ignore_class_col}")
    else:
        ignore_class_col = ask_ignore_class_column(header_row_data, class_col)
        if ignore_class_col == "exit":
            print("程序已退出。")
            return

    os.system('cls' if os.name == 'nt' else 'clear')
    print("开始拆分文件...")
    stats = split_and_save(selected, sheet_index, sheet_name, header_row, class_col, working_dir, student_id_col, ignore_class_col)
    
    result = show_completion_options(working_dir, stats)
    if result == "open":
        print("正在打开输出文件夹...")
    elif result == "exit":
        print("程序已退出。")


if __name__ == "__main__":
    main()