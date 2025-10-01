# -*- coding: utf-8 -*-

import os
from datetime import datetime
from openpyxl import load_workbook, Workbook


def split_and_save(selected_files, sheet_index, sheet_name, header_row, class_col, working_dir=".", student_id_col=None, ignore_class_col=False):
    output_dir = os.path.join(working_dir, "拆分")
    os.makedirs(output_dir, exist_ok=True)

    class_data = {}
    subject_headers = {}
    
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
        
        full_file_path = os.path.join(working_dir, file)
        subject = os.path.splitext(file)[0]
        wb = load_workbook(full_file_path, read_only=False)
        
        sheets = wb.sheetnames
        if sheet_index < len(sheets):
            ws = wb[sheets[sheet_index]]
        else:
            print(f"\n文件 {file} 没有足够多的sheet，跳过该文件")
            stats["skipped_files"] += 1
            wb.close()
            continue

        header = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        
        if student_id_col is not None or ignore_class_col:
            modified_header = []
            for i, cell in enumerate(header):
                if student_id_col is not None and i + 1 == student_id_col:
                    continue
                if ignore_class_col and i + 1 == class_col:
                    continue
                modified_header.append(cell)
            subject_headers[subject] = modified_header
        else:
            subject_headers[subject] = header

        rows = list(ws.iter_rows(values_only=True))

        for row in rows[header_row:]:
            if not row or not row[class_col - 1]:
                continue
            class_name = str(row[class_col - 1])
            if class_name not in class_data:
                class_data[class_name] = {}
            if subject not in class_data[class_name]:
                class_data[class_name][subject] = []
            
            if student_id_col is not None or ignore_class_col:
                modified_row = []
                for i, cell in enumerate(row):
                    if student_id_col is not None and i + 1 == student_id_col:
                        continue
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

    sorted_classes = sorted(class_data.keys(), key=lambda x: int(x) if x.isdigit() else x)
    stats["generated_classes"] = len(sorted_classes)
    
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    total_classes = len(sorted_classes)
    for idx, cls in enumerate(sorted_classes, 1):
        print(f"\r正在保存班级文件 ({idx}/{total_classes}): {cls}.xlsx", end="", flush=True)
        subjects = class_data[cls]
        out_file = os.path.join(output_dir, f"{cls}.xlsx")
        out_wb = Workbook()
        out_wb.remove(out_wb.active)
        for subject, rows in subjects.items():
            ws = out_wb.create_sheet(title=subject)
            title_row = [f"{subject}", f"{current_date}"]
            if len(subject_headers[subject]) > len(title_row):
                title_row.extend([""] * (len(subject_headers[subject]) - len(title_row)))
            ws.append(title_row)
            ws.append(subject_headers[subject])
            for row in rows:
                ws.append(row)
        out_wb.save(out_file)
    
    print("\n所有班级文件保存完成!")
    return stats