# -*- coding: utf-8 -*-
"""
用户输入工具模块
处理用户输入相关功能
"""

import os
from prompt_toolkit import prompt
from prompt_toolkit.application import get_app, Application
from prompt_toolkit.layout import Layout
from prompt_toolkit.widgets import RadioList, Button, Label, Box
from prompt_toolkit.layout.containers import HSplit, VSplit, Window
from prompt_toolkit.styles import Style


def ask_number(prompt_text):
    while True:
        try:
            user_input = prompt("输入" + prompt_text)
            if user_input.lower() == 'exit':
                return 'exit'
            return int(user_input)
        except ValueError:
            print("输入无效，请输入一个有效的数字！")


def choose_class_column(header_row_data):
    os.system('cls' if os.name == 'nt' else 'clear')
    
    values = []
    for i, cell in enumerate(header_row_data[:10]):
        values.append((i+1, f"列{i+1}: {cell}"))
    
    radio_list = RadioList(values=values)
    radio_list.current_value = 1
    
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
        Label("请选择班级所在列:", dont_extend_height=True),
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
        return "exit"
    elif choice is not None:
        return choice
    else:
        return None