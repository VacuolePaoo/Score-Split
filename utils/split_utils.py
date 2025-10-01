# -*- coding: utf-8 -*-
"""
拆分工具模块
处理Excel文件的拆分和保存功能
"""

import os
from datetime import datetime
from openpyxl import load_workbook, Workbook


def split_and_save(selected_files, sheet_index, sheet_name, header_row, class_col, working_dir=".", student_id_col=None, ignore_class_col=False):
    output_dir = os.path.join(working_dir, "拆分")
    os.makedirs(output_dir, exist_ok=True)

    class_data = {}  # {class_name: {subject: [rows]}}
    subject_headers = {}  # {subject: header}
    
    # 统计信息
    stats = {
        "processed_files": 0,
        "total_files": len(selected_files),
        "generated_classes": 0,
        "total_rows": 0,
        "skipped_files": 0
    }
    
    total_files = len(selected_files)
    print(f"开始处理 {total_files} 个文件...")
    
    for idx, file in enumerate(selected_files, 1):
        print(f"\r正在处理文件 ({idx}/{total_files}): {file}", end="", flush=True)
        
        # 构造完整的文件路径
        full_file_path = os.path.join(working_dir, file)
        subject = os.path.splitext(file)[0]  # 文件名 = 学科
        wb = load_workbook(full_file_path, read_only=False)
        
        # 使用指定索引获取sheet
        sheets = wb.sheetnames
        if sheet_index < len(sheets):
            ws = wb[sheets[sheet_index]]
        else:
            print(f"\n文件 {file} 没有足够多的sheet，跳过该文件")
            stats["skipped_files"] += 1
            wb.close()
            continue

        # 提取表头（每个学科的表头可能不同）
        header = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        
        # 如果指定了学号列或需要忽略班级列，则从表头中移除相应列
        if student_id_col is not None or ignore_class_col:
            modified_header = []
            for i, cell in enumerate(header):
                # 如果指定了学号列且当前列是学号列，则跳过
                if student_id_col is not None and i + 1 == student_id_col:
                    continue
                # 如果需要忽略班级列且当前列是班级列，则跳过
                if ignore_class_col and i + 1 == class_col:
                    continue
                modified_header.append(cell)
            subject_headers[subject] = modified_header
        else:
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
            
            # 如果指定了学号列或需要忽略班级列，则从数据行中移除相应列
            if student_id_col is not None or ignore_class_col:
                modified_row = []
                for i, cell in enumerate(row):
                    # 如果指定了学号列且当前列是学号列，则跳过
                    if student_id_col is not None and i + 1 == student_id_col:
                        continue
                    # 如果需要忽略班级列且当前列是班级列，则跳过
                    if ignore_class_col and i + 1 == class_col:
                        continue
                    modified_row.append(cell)
                class_data[class_name][subject].append(tuple(modified_row))
            else:
                class_data[class_name][subject].append(row)
            
            stats["total_rows"] += 1

        stats["processed_files"] += 1
        wb.close()

    print("\n数据提取完成，正在生成班级文件...")

    # 按班级排序
    sorted_classes = sorted(class_data.keys(), key=lambda x: int(x) if x.isdigit() else x)
    stats["generated_classes"] = len(sorted_classes)
    
    # 获取当前日期用于制表日期
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # 保存每个班的文件
    total_classes = len(sorted_classes)
    for idx, cls in enumerate(sorted_classes, 1):
        print(f"\r正在保存班级文件 ({idx}/{total_classes}): {cls}.xlsx", end="", flush=True)
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
    
    print("\n所有班级文件保存完成!")
    return stats