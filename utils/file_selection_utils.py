# -*- coding: utf-8 -*-
"""
文件选择工具模块
处理文件选择和输出目录检查功能
"""

import os
from prompt_toolkit.application import get_app, Application
from prompt_toolkit.layout import Layout
from prompt_toolkit.widgets import CheckboxList, RadioList, Button, Label, Box
from prompt_toolkit.layout.containers import HSplit, VSplit, Window
from prompt_toolkit.styles import Style


def check_output_dir(working_dir="."):
    """
    检查输出目录是否存在文件，并根据情况提供用户选项
    """
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    output_dir = os.path.join(working_dir, "拆分")
    os.makedirs(output_dir, exist_ok=True)
    
    # 检查目录中是否存在任何文件
    existing_files = [f for f in os.listdir(output_dir) if os.path.isfile(os.path.join(output_dir, f))]
    
    # 如果目录为空，直接返回继续执行
    if not existing_files:
        return True
    
    # 如果目录不为空，显示提示并提供选项
    print(f"警告: 输出目录 '{output_dir}' 中已存在以下文件:")
    for f in existing_files[:5]:  # 只显示前5个文件
        print(f"  - {f}")
    if len(existing_files) > 5:
        print(f"  ... 还有 {len(existing_files) - 5} 个文件")
    
    # 创建选项对话框
    values = [
        ("exit", "退出程序，自行处理文件"),
        ("delete", "删除所有现有文件"),
        ("overwrite", "直接覆盖现有文件")
    ]
    radio_list = RadioList(values=values)
    radio_list.current_value = "exit"  # 默认选择退出
    
    def on_confirm():
        get_app().exit(result=radio_list.current_value)
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_confirm = Button(text="确认", handler=on_confirm)
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建布局
    style = Style.from_dict({
        "button.focused": "reverse",
    })
    
    body = HSplit([
        Label(f"输出目录 '{output_dir}' 中已存在 {len(existing_files)} 个文件:", dont_extend_height=True),
        Window(height=1, char="-"),
        Box(body=radio_list, padding=1),
        Window(height=1, char="-"),
        VSplit([btn_confirm, btn_exit], padding=3),
    ])
    
    application = Application(
        layout=Layout(body, focused_element=radio_list),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    choice = application.run()
    
    if choice == "exit":
        print("用户选择退出程序，请自行处理输出目录中的文件。")
        return False
    elif choice == "delete":
        print("正在删除现有文件...")
        for f in existing_files:
            try:
                os.remove(os.path.join(output_dir, f))
                print(f"  已删除: {f}")
            except Exception as e:
                print(f"  删除 {f} 失败: {e}")
        print("所有现有文件已删除。")
        return True
    elif choice == "overwrite":
        print("用户选择直接覆盖现有文件。")
        return True
    else:
        # 默认退出
        print("操作被取消。")
        return False


def choose_files(files):
    # 清屏
    os.system('cls' if os.name == 'nt' else 'clear')
    
    if not files:
        print("未找到任何xlsx文件")
        return []
    
    # 创建自定义的checkboxlist_dialog
    values = [(f, f) for f in files]
    checkbox = CheckboxList(values=values)
    
    # 创建按钮
    def on_next():
        get_app().exit(result=list(checkbox.current_values or []))
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_next = Button(text="下一步", handler=on_next)
    btn_select_all = Button(text="全选", handler=lambda: setattr(checkbox, 'current_values', [v for _, v in values]))
    btn_exit = Button(text="退出", handler=on_exit)
    
    # 创建布局
    style = Style.from_dict({
        "button.focused": "reverse",
    })
    
    body = HSplit([
        Label("请选择要处理的学科 xlsx 文件：", dont_extend_height=True),
        Window(height=1, char="-"),
        Box(body=checkbox, padding=1),
        Window(height=1, char="-"),
        VSplit([btn_select_all, btn_next, btn_exit], padding=3),
    ])
    
    application = Application(
        layout=Layout(body, focused_element=checkbox),
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    selected = application.run()
    
    if selected == "exit":
        return "exit"
    return selected or []

