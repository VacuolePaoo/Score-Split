# -*- coding: utf-8 -*-
"""
模块名称: exam_splitter
职责范围: 将多个学科成绩单(Excel)按班级拆分，并生成每个班级的成绩汇总表
期望实现计划:
    1. 扫描目录获取所有 xlsx 文件 -> 用户确认
    2. 提取每个文件内的 "学生得分" sheet -> 用户确认
    3. 用户输入表头所在行号
    4. 用户确认班级列所在列号
    5. 按班级排序
    6. 拆分保存到"拆分"文件夹，每个班级一个 Excel，包含各科成绩
已实现功能:
    - 支持交互式选择 (上下键+回车)
    - 自动创建输出文件夹
    - 按班级拆分保存
使用依赖:
    - openpyxl (Excel 处理)
    - prompt_toolkit (交互式菜单)
主要接口:
    - main(): 程序入口
注意事项:
    - 假设所有 Excel 都在当前目录
    - 文件名就是学科名
    - sheet 名需包含 "学生得分"
"""

import os
import warnings
from datetime import datetime
from openpyxl import load_workbook, Workbook
from prompt_toolkit import prompt
from prompt_toolkit.application import Application, get_app
from prompt_toolkit.key_binding import KeyBindings
from prompt_toolkit.layout import Layout
from prompt_toolkit.widgets import CheckboxList, Button, Label, Box, Frame, RadioList
from prompt_toolkit.layout.containers import HSplit, VSplit, Window
from prompt_toolkit.styles import Style
# 移除了WordCompleter的导入，因为我们不再需要自动补全功能

# 忽略所有警告
warnings.filterwarnings("ignore")


def list_excel_files():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    return files


def check_output_dir():
    """
    检查输出目录是否存在文件，并根据情况提供用户选项
    """
    output_dir = "拆分"
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
    
    def on_cancel():
        get_app().exit(result="exit")  # 取消操作等同于选择退出
    
    btn_confirm = Button(text="确认", handler=on_confirm)
    btn_cancel = Button(text="取消", handler=on_cancel)
    
    # 创建按键绑定
    kb = KeyBindings()
    
    @kb.add('enter')
    def _(event):
        # 回车：确认选择
        get_app().exit(result=radio_list.current_value)
    
    # 创建布局
    style = Style.from_dict({
        "button.focused": "reverse",
    })
    
    body = HSplit([
        Label(f"输出目录 '{output_dir}' 中已存在 {len(existing_files)} 个文件:", dont_extend_height=True),
        Window(height=1, char="-"),
        Box(body=radio_list, padding=1),
        Window(height=1, char="-"),
        VSplit([btn_confirm, btn_cancel], padding=3),
    ])
    
    application = Application(
        layout=Layout(body, focused_element=radio_list),
        key_bindings=kb,
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
    if not files:
        print("未找到任何xlsx文件")
        return []
    
    # 创建自定义的checkboxlist_dialog
    values = [(f, f) for f in files]
    checkbox = CheckboxList(values=values)
    
    # 创建按钮
    def on_next():
        get_app().exit(result=list(checkbox.current_values or []))
    
    def on_cancel():
        get_app().exit(result=[])
    
    btn_next = Button(text="下一步", handler=on_next)
    btn_cancel = Button(text="取消", handler=on_cancel)
    btn_select_all = Button(text="全选", handler=lambda: setattr(checkbox, 'current_values', [v for _, v in values]))
    
    # 创建按键绑定
    kb = KeyBindings()
    
    @kb.add('a')
    def _(event):
        # A键：切换全选 / 取消全选
        current = set(checkbox.current_values or [])
        all_items = [v for _, v in values]
        if len(current) == len(all_items):
            checkbox.current_values = []
        else:
            checkbox.current_values = all_items
    
    @kb.add('enter')
    def _(event):
        # 回车：确认并下一步
        get_app().exit(result=list(checkbox.current_values or []))
    
    # 创建布局
    style = Style.from_dict({
        "button.focused": "reverse",
    })
    
    body = HSplit([
        Label("请选择要处理的学科 xlsx 文件：", dont_extend_height=True),
        Window(height=1, char="-"),
        Box(body=checkbox, padding=1),
        Window(height=1, char="-"),
        VSplit([btn_select_all, btn_next, btn_cancel], padding=3),
    ])
    
    application = Application(
        layout=Layout(body, focused_element=checkbox),
        key_bindings=kb,
        mouse_support=True,
        full_screen=False,
        style=style,
    )
    
    selected = application.run()
    return selected or []


def list_all_sheets(file):
    wb = load_workbook(file, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    return sheets


def ask_sheet_index(sheets):
    print("所有学科的sheet结构相同，以下为第一个文件的sheet列表:")
    for i, sheet in enumerate(sheets):
        print(f"  {i+1}. {sheet}")
    
    while True:
        try:
            index = int(prompt("请输入学生得分sheet的序号: "))  # 移除了completer参数
            if 1 <= index <= len(sheets):
                return index-1, sheets[index-1]  # 返回索引和sheet名称
            else:
                print(f"请输入1到{len(sheets)}之间的数字")
        except ValueError:
            print("请输入有效的数字")


def ask_number(prompt_text):
    return int(prompt(prompt_text))  # 移除了completer参数


def split_and_save(selected_files, sheet_index, sheet_name, header_row, class_col):
    output_dir = "拆分"
    os.makedirs(output_dir, exist_ok=True)

    class_data = {}  # {class_name: {subject: [rows]}}
    subject_headers = {}  # {subject: header}

    for file in selected_files:
        subject = os.path.splitext(file)[0]  # 文件名 = 学科
        wb = load_workbook(file, read_only=False)
        
        # 使用指定索引获取sheet
        sheets = wb.sheetnames
        if sheet_index < len(sheets):
            ws = wb[sheets[sheet_index]]
        else:
            print(f"文件 {file} 没有足够多的sheet，跳过该文件")
            wb.close()
            continue

        # 提取表头（每个学科的表头可能不同）
        header = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        subject_headers[subject] = header

        # 提取数据
        rows = list(ws.iter_rows(values_only=True))

        for row in rows[header_row:]:
            if not row or not row[class_col - 1]:
                continue
            class_name = str(row[class_col - 1])
            if class_name not in class_data:
                class_data[class_name] = {}
            if subject not in class_data[class_name]:
                class_data[class_name][subject] = []
            class_data[class_name][subject].append(row)

        wb.close()

    # 按班级排序
    sorted_classes = sorted(class_data.keys(), key=lambda x: int(x) if x.isdigit() else x)
    
    # 获取当前日期用于制表日期
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # 保存每个班的文件
    for cls in sorted_classes:
        subjects = class_data[cls]
        out_file = os.path.join(output_dir, f"{cls}.xlsx")
        out_wb = Workbook()
        out_wb.remove(out_wb.active)  # 删除默认sheet
        for subject, rows in subjects.items():
            ws = out_wb.create_sheet(title=subject)
            # 添加标题行，分别放在四个单元格中
            title_row = [f"{subject}", f"{current_date}"]
            # 根据表头长度调整标题行的长度
            if len(subject_headers[subject]) > len(title_row):
                title_row.extend([""] * (len(subject_headers[subject]) - len(title_row)))
            ws.append(title_row)
            # 使用每个学科自己的表头
            ws.append(subject_headers[subject])
            for row in rows:
                ws.append(row)
        out_wb.save(out_file)


def main():
    # 步骤1: 检查输出目录
    if not check_output_dir():
        return

    # 步骤2: 扫描当前目录，获取所有xlsx文件
    files = list_excel_files()
    print("找到以下Excel文件:")
    for f in files:
        print(f"  - {f}")
    
    
    # 步骤3: 让用户选择要处理的文件
    selected = choose_files(files)
    if not selected:
        print("未选择文件，退出。")
        return

    # 步骤4: 获取第一个文件的sheet列表并让用户输入序号
    first_file = selected[0]
    sheets = list_all_sheets(first_file)
    sheet_index, sheet_name = ask_sheet_index(sheets)

    # 步骤5: 询问用户表头所在行数
    header_row = ask_number("请输入表头所在行号: ")
    
    # 步骤6: 检索表头行内容并询问用户班级列所在列数
    # 显示第一个文件的表头作为示例
    wb = load_workbook(first_file, read_only=True)
    ws = wb[sheet_name]
    header_row_data = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    wb.close()
    
    print(f"表头行（第{header_row}行）的前10个单元格内容:")
    for i, cell in enumerate(header_row_data[:10]):
        print(f"  列{i+1}: {cell}")
    
    class_col = ask_number("请输入班级所在列号: ")

    # 步骤7-9: 拆分并保存文件
    split_and_save(selected, sheet_index, sheet_name, header_row, class_col)
    print("拆分完成，结果保存在 '拆分' 文件夹中。")


if __name__ == "__main__":
    main()

