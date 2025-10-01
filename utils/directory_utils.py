# -*- coding: utf-8 -*-

import os
from prompt_toolkit import prompt
from prompt_toolkit.application import get_app, Application
from prompt_toolkit.layout import Layout
from prompt_toolkit.widgets import RadioList, Button, Label, Box
from prompt_toolkit.layout.containers import HSplit, VSplit, Window
from prompt_toolkit.styles import Style


def choose_working_directory():
    os.system('cls' if os.name == 'nt' else 'clear')
    
    values = [
        ("current", "扫描本文件夹下工作簿"),
        ("manual", "手动输入文件夹路径")
    ]
    radio_list = RadioList(values=values)
    radio_list.current_value = "current"
    
    def on_confirm():
        get_app().exit(result=radio_list.current_value)
    
    def on_exit():
        get_app().exit(result="exit")
    
    btn_confirm = Button(text="确认", handler=on_confirm)
    btn_exit = Button(text="退出", handler=on_exit)
    
    style = Style.from_dict({
        "button.focused": "reverse",
    })
    
    body = HSplit([
        Label("请选择工作目录:", dont_extend_height=True),
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
        return "exit", "exit"
    elif not choice:
        return None, None
    
    if choice == "current":
        return ".", "current"
    else:
        path = prompt("请输入文件夹路径 (支持拖入文件夹获取路径): ").strip().strip('"\'')
        if not os.path.exists(path):
            print(f"路径不存在: {path}")
            return None, None
        if not os.path.isdir(path):
            print(f"路径不是文件夹: {path}")
            return None, None
        return path, "manual"


def list_excel_files(directory="."):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    return files