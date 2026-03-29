#!/usr/bin/env python3
"""
差旅发票整理 - 行程单处理补丁
"""

def add_trip_sheet_support():
    """添加行程单处理功能"""
    
    main_py = '/Users/lichengbao/.openclaw/workspace/skills/invoice-organizer/scripts/main.py'
    
    with open(main_py, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 1. 找到 download_attachments 函数中的附件处理部分
    # 2. 添加行程单处理逻辑
    
    new_lines = []
    in_attachment_loop = False
    
    for i, line in enumerate(lines):
        # 找到 PDF 处理部分
        if "if filename.endswith('.pdf') and '通行费' not in filename and '行程' not in filename:" in line:
            # 替换为不跳过行程单
            new_lines.append("                if filename.endswith('.pdf') and '通行费' not in filename:\n")
        elif "process_pdf(tmp_path, pdf_dir, all_records, inv_type, em, filename)" in line and not in_attachment_loop:
            # 第一个出现是在 download_attachments 中，需要添加行程单判断
            indent = "                    "
            new_lines.append(f"{indent}if '行程' in filename or '报销单' in filename:\n")
            new_lines.append(f"{indent}    process_trip_sheet(tmp_path, all_records, em)\n")
            new_lines.append(f"{indent}else:\n")
            new_lines.append(f"{indent}    process_pdf(tmp_path, pdf_dir, all_records, inv_type, em, filename)\n")
            in_attachment_loop = True
        else:
            new_lines.append(line)
    
    # 添加 process_trip_sheet 函数（在 process_pdf 之前）
    final_lines = []
    for line in new_lines:
        if 'def process_pdf(' in line:
            # 在 process_pdf 之前插入 process_trip_sheet
            final_lines.append('''
def process_trip_sheet(trip_path, all_records, em):
    """从行程单 PDF 提取路线信息"""
    try:
        from pdfminer.high_level import extract_text
        import re
        
        raw_text = extract_text(trip_path)
        text = " ".join(raw_text.split())
        
        route_from = ''
        route_to = ''
        
        # Pattern 1: 从 XXX 到 XXX
        route_m = re.search(r'从 ([\\u4e00-\\u9fffA-Za-z]{2,20}?) 到 ([\\u4e00-\\u9fffA-Za-z]{2,20}?)', text)
        if route_m:
            route_from = route_m.group(1).strip()
            route_to = route_m.group(2).strip()
        
        # Pattern 2: XXX→XXX
        if not route_from or not route_to:
            arrow_m = re.search(r'([\\u4e00-\\u9fffA-Za-z]{2,20}?)→([\\u4e00-\\u9fffA-Za-z]{2,20}?)', text)
            if arrow_m:
                route_from = arrow_m.group(1).strip()
                route_to = arrow_m.group(2).strip()
        
        # Match with invoices by date
        if route_from and route_to:
            trip_date = em.get('date', '')
            for rec in all_records:
                if rec['type'] == '滴滴' and rec['trip_date'] == trip_date and not rec.get('from'):
                    rec['from'] = route_from
                    rec['to'] = route_to
                    print(f"    [行程单匹配] {rec['inv_no']} {route_from}→{route_to}")
    except Exception as e:
        print(f"    行程单处理失败：{e}")

''')
        final_lines.append(line)
    
    with open(main_py, 'w', encoding='utf-8') as f:
        f.writelines(final_lines)
    
    print('✅ 补丁应用完成')

if __name__ == '__main__':
    add_trip_sheet_support()
