#!/usr/bin/env python3
"""
差旅发票整理 - main.py
读取 config.env 配置，交互式输入日期范围，全自动整理发票。
"""
import os, sys, re, json, zipfile, shutil, imaplib, email, fitz
from email.header import decode_header
from email.utils import parsedate_to_datetime
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pdfminer.high_level import extract_text

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(os.path.dirname(SKILL_DIR), 'config.env')
OUTPUT_DIR = os.path.join(os.path.expanduser('~/Desktop'), '差旅发票')

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============ 1. 读取配置 ============
def load_config():
    # 尝试读取记忆文件 MEMORY.md 中的 IMAP 配置
    memory_path = os.path.join(os.path.dirname(os.path.dirname(SKILL_DIR)), '..', 'workspace', 'MEMORY.md')
    memory_path = os.path.normpath(os.path.expanduser(memory_path))
    memory_path = os.path.join(os.path.expanduser('~'), '.qclaw', 'workspace', 'MEMORY.md')
    
    cfg = {}
    if os.path.exists(memory_path):
        with open(memory_path, encoding='utf-8') as f:
            text = f.read()
        # 在记忆文件中提取 IMAP 配置
        import re
        for match in re.finditer(r'^([A-Z_]+)=(.*)$', text, re.MULTILINE):
            k, v = match.group(1), match.group(2).strip()
            if k in ('IMAP_HOST','IMAP_PORT','IMAP_USER','IMAP_PASS','IMAP_MAILBOX'):
                cfg[k] = v
        print(f"  从记忆文件读取IMAP配置: {cfg.get('IMAP_USER','')}")
    
    # 备选：读取 config.env
    if not cfg.get('IMAP_USER'):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if '=' in line and not line.startswith('#'):
                        k, v = line.split('=', 1)
                        if k.strip() in ('IMAP_HOST','IMAP_PORT','IMAP_USER','IMAP_PASS','IMAP_MAILBOX'):
                            cfg[k.strip()] = v.strip()
    
    # 检查必要项
    if not cfg.get('IMAP_USER') or not cfg.get('IMAP_PASS'):
        print(f"❌ 未配置IMAP信息。")
        print(f"请在 ~/.qclaw/workspace/MEMORY.md 中添加邮件配置（参考 SKILL.md）")
        sys.exit(1)
    
    return cfg

# ============ 2. 辅助函数 ============
def decode_str(s):
    if s is None: return ''
    parts = decode_header(s)
    result = []
    for b, cs in parts:
        if isinstance(b, bytes): result.append(b.decode(cs or 'utf-8', errors='replace'))
        else: result.append(b)
    return ''.join(result)

def get_pdf_path(inv_no, inv_type, didi_dir, yg_dir):
    """根据发票号和类型查找PDF路径"""
    candidates = [
        f"{didi_dir}/{inv_no}_发票.pdf",
        f"{didi_dir}/{inv_no}.pdf",
        f"{yg_dir}/{inv_no}_发票.pdf",
        f"{yg_dir}/{inv_no}.pdf",
    ]
    for p in candidates:
        if os.path.exists(p): return p
    return None

# ============ 3. 连接邮箱并搜索发票邮件 ============
def scan_invoices(cfg, start_date, end_date):
    print(f"\n📧 连接邮箱 {cfg['IMAP_USER']} ...")
    mail = imaplib.IMAP4_SSL(cfg.get('IMAP_HOST','imap.qq.com'),
                              int(cfg.get('IMAP_PORT') or '993'))
    mail.login(cfg['IMAP_USER'], cfg['IMAP_PASS'].encode('ascii', 'replace').decode('ascii'))
    mail.select(cfg.get('IMAP_MAILBOX', 'INBOX'))

    # 搜索时间范围内的所有邮件
    _, msgs = mail.search(None, 'SINCE', f'"{start_date}"', 'BEFORE', f'"{end_date}"')
    seq_nums = [b.decode() for b in msgs[0].split()]
    print(f"  找到 {len(seq_nums)} 封邮件，正在扫描发票类型...")

    invoice_emails = {}  # type -> list of {seq, subject, date, from}

    for seq in seq_nums:
        status, data = mail.fetch(seq.encode(), '(BODY[HEADER.FIELDS (DATE FROM SUBJECT)])')
        if status != 'OK' or not data or not data[0]: continue
        msg = email.message_from_bytes(data[0][1])
        sender = decode_str(msg.get('From',''))
        subject = decode_str(msg.get('Subject',''))
        try:
            email_date = parsedate_to_datetime(msg.get('Date','')).strftime('%Y-%m-%d')
        except:
            email_date = ''
        
        # 自动识别发票类型
        inv_type = None
        if 'xiaojukeji.com' in sender or 'didichuxing' in sender.lower():
            if '第三方' in subject or '阳光出行' in subject:
                inv_type = '阳光出行'
            else:
                inv_type = '滴滴'
        elif 'txffp.com' in sender or '通行费' in subject:
            inv_type = '通行费'
        elif '火车票' in subject or '高铁' in subject or '机票' in subject:
            inv_type = '其他'
        
        if inv_type:
            if inv_type not in invoice_emails:
                invoice_emails[inv_type] = []
            invoice_emails[inv_type].append({'seq': seq, 'subject': subject, 'date': email_date, 'sender': sender})
            print(f"  [{seq}] {inv_type}: {subject[:40]}")

    mail.logout()
    return invoice_emails

# ============ 4. 下载并处理附件 ============
def download_attachments(cfg, invoice_emails, start_date, end_date):
    print(f"\n📥 下载附件...")
    mail = imaplib.IMAP4_SSL(cfg.get('IMAP_HOST','imap.qq.com'),
                              int(cfg.get('IMAP_PORT') or '993'))
    mail.login(cfg['IMAP_USER'], cfg['IMAP_PASS'].encode('ascii', 'replace').decode('ascii'))
    mail.select(cfg.get('IMAP_MAILBOX', 'INBOX'))

    # 创建临时目录
    tmp_dir = f'{OUTPUT_DIR}/_tmp'
    os.makedirs(tmp_dir, exist_ok=True)
    pdf_dir = f'{OUTPUT_DIR}/发票PDF'
    os.makedirs(pdf_dir, exist_ok=True)

    all_records = []  # 所有发票记录
    invoice_pages = {}  # inv_no -> word_page

    for inv_type, emails in invoice_emails.items():
        print(f"\n  === {inv_type} ({len(emails)}封邮件) ===")
        
        for em in emails:
            status, data = mail.fetch(em['seq'].encode(), '(RFC822)')
            if status != 'OK' or not data or not data[0]: continue
            raw = data[0][1]
            msg = email.message_from_bytes(raw)
            
            counter = {}
            for part in msg.walk():
                disp = part.get_content_disposition()
                filename = decode_str(part.get_filename(''))
                if not filename: continue
                raw_data = part.get_payload(decode=True)
                if raw_data is None: continue
                
                # PDF附件（通行费汇总单从ZIP处理，跳过）
                if filename.endswith('.pdf') and '通行费' not in filename and '行程' not in filename:
                    key = filename
                    counter[key] = counter.get(key, 0) + 1
                    tmp_path = f"{tmp_dir}/{em['seq']}_{counter[key]}_{filename}"
                    with open(tmp_path, 'wb') as f: f.write(raw_data)
                    process_pdf(tmp_path, pdf_dir, all_records, inv_type, em, filename)
                
                # ZIP附件（通行费）
                elif filename.endswith('.zip') and '通行费' in filename:
                    zip_path = f"{tmp_dir}/{em['seq']}_{filename}"
                    with open(zip_path, 'wb') as f: f.write(raw_data)
                    process_toll_zip(zip_path, pdf_dir, all_records, em)

    mail.logout()

    # 清理临时目录
    shutil.rmtree(tmp_dir, ignore_errors=True)
    
    return all_records
def process_pdf(tmp_path, pdf_dir, all_records, inv_type, em, filename):
    """Process single PDF invoice, skip trip summary docs"""
    # Skip non-invoice attachments
    if '行示' in filename or '汇总e单' in filename:
        return

    try:
        raw_text = extract_text(tmp_path)
        text = " ".join(raw_text.split()).replace(" ", "")
    except:
        text = ''

    # Extract invoice number
    inv_no = ''
    if inv_type == '滴滴':
        m = re.search(r'发票号码[：:]+\s*([\d]+)', text)
        if m: inv_no = m.group(1)
    elif inv_type == '阳光出行':
        m = re.search(r'发票号码[：:]+\s*([\d]+)', text)
        if m: inv_no = m.group(1)

    if not inv_no:
        # 文件名格式: SEQ_序号_滴滴电子发票.pdf，跳过前两个字段
        parts = os.path.basename(tmp_path).replace('.pdf','').split('_')
        inv_no = parts[-1] if len(parts) > 2 else os.path.basename(tmp_path).replace('.pdf','')

    # Save PDF (dedup)
    final_path = f"{pdf_dir}/{inv_no}.pdf"
    if not os.path.exists(final_path):
        shutil.copy(tmp_path, final_path)

    # Extract amount
    amounts = re.findall(r'(\d+\.\d+)', text)
    amount = float(amounts[-1]) if amounts else 0

    # Extract trip date
    date_m = re.search(r'出行日期[：:]+\s*(\d{4}[-/.]\d{2}[-/.]\d{2})', text)
    trip_date = date_m.group(1) if date_m else em.get('date','')

    # Extract from/to
    from_m = re.search(r'出\s*发\s*地 [：:]*\s*([\u4e00-\u9fffA-Za-z0-9\(\)\-]{0,50}?)', text)
    to_m = re.search(r'到\s*达\s*地 [：:]*\s*([\u4e00-\u9fffA-Za-z0-9\(\)\-]{0,50}?)', text)
    route_from = from_m.group(1).strip() if from_m else ''
    route_to = to_m.group(1).strip() if to_m else ''

    record = {
        'type': inv_type, 'inv_no': inv_no, 'trip_date': trip_date,
        'from': route_from, 'to': route_to, 'amount': amount,
        'pdf_path': final_path
    }
    all_records.append(record)
    print(f"    {inv_no} ¥{amount:.2f} {trip_date}")


def process_toll_zip(zip_path, pdf_dir, all_records, em):
    """处理通行费ZIP包"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            xml_files = [n for n in zf.namelist() if n.startswith('xml/') and not n.endswith('/')]
            pdf_files = {n.replace('pdf/','').replace('.pdf',''): n for n in zf.namelist() if n.startswith('pdf/')}
            
            for xml_name in sorted(xml_files):
                # 只处理XML文件（真实发票数据）
                if not xml_name.endswith('.xml'):
                    continue
                content = zf.read(xml_name).decode('utf-8', errors='replace')
                
                # 真实发票号
                inv_m = re.search(r'<EIid>(\d{20,})</EIid>', content)
                if not inv_m:
                    inv_m = re.search(r'<InvoiceNumber>(\d{20,})</InvoiceNumber>', content)
                inv_no = inv_m.group(1).strip() if inv_m else ''
                
                # 金额
                amt_m = re.search(r'<TotaltaxIncludedAmount>(\d+\.?\d*)</TotaltaxIncludedAmount>', content)
                amount = float(amt_m.group(1)) if amt_m else 0
                
                # 日期
                date_m = re.search(r'<IssueTime>(\d{4}-\d{2}-\d{2})</IssueTime>', content)
                invoice_date = date_m.group(1) if date_m else em.get('date','')
                
                # 路线（从汇总单获取）
                route_m = re.search(r'出入口信息\s*[^\d]*?([\u4e00-\u9fff·A-Za-z0-9]+\s*站)', content)
                route_from = ''
                route_to = ''
                
                # 对应PDF
                xml_key = xml_name.replace('xml/', '').replace('.xml', '')
                pdf_name = pdf_files.get(xml_key, '')
                
                if inv_no:
                    final_path = f"{pdf_dir}/{inv_no}.pdf"
                    if not os.path.exists(final_path) and pdf_name:
                        pdf_data = zf.read(pdf_name)
                        with open(final_path, 'wb') as f:
                            f.write(pdf_data)
                    
                    record = {
                        'type': '通行费', 'inv_no': inv_no,
                        'trip_date': invoice_date, 'from': route_from, 'to': route_to,
                        'amount': amount, 'pdf_path': final_path
                    }
                    all_records.append(record)
                    print(f"    {inv_no} ¥{amount} {invoice_date}")
    except Exception as e:
        print(f"    ZIP处理错误: {e}")

# ============ 5. 生成Word（A5横向，每页一张发票+行程信息） ============
def build_word(all_records):
    print(f"\n📄 生成Word...")
    doc = Document()
    s0 = doc.sections[0]
    s0.page_width = Mm(210); s0.page_height = Mm(148)
    s0.left_margin = Mm(8); s0.right_margin = Mm(8)
    s0.top_margin = Mm(8); s0.bottom_margin = Mm(8)

    for idx, rec in enumerate(all_records, 1):
        if idx > 1:
            sec = doc.add_section()
            sec.page_width = Mm(210); sec.page_height = Mm(148)
            sec.left_margin = Mm(8); sec.right_margin = Mm(8)
            sec.top_margin = Mm(8); sec.bottom_margin = Mm(8)
        
        # PDF图片
        if os.path.exists(rec['pdf_path']):
            try:
                pdf_doc = fitz.open(rec['pdf_path'])
                page = pdf_doc[0]
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                img_path = f'/tmp/travel_inv_{idx:03d}.png'
                pix.save(img_path)
                pdf_doc.close()
                
                p2 = doc.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run2 = p2.add_run()
                run2.add_picture(img_path, width=Mm(194))
            except Exception as e:
                print(f"    PDF渲染失败 {rec['inv_no']}: {e}")
        
        # 右下角页码
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run3 = p3.add_run()
        run3.text = f"Page {idx}"
        run3.font.size = Pt(7)
        run3.font.name = "Arial"
        
        rec['word_page'] = idx
    
    out = f'{OUTPUT_DIR}/差旅发票汇总打印.docx'
    doc.save(out)
    print(f"  已保存: {out}")
    return out

# ============ 6. 生成Excel ============
def build_excel(all_records):
    print(f"\n📊 生成Excel...")
    wb = Workbook()
    ws = wb.active
    ws.title = "差旅发票汇总"
    
    thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    hf = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    dfont = Font(name="Arial", size=10)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    ws.merge_cells("A1:H1")
    ws["A1"] = "差旅发票汇总表"
    ws["A1"].font = Font(name="Arial", bold=True, size=14)
    ws["A1"].alignment = ctr
    
    headers = ["序号","类型","发票号码","行程日期","出发地/出发站","到达地/到达站","金额(元)","Word页码"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = hfont; c.fill = hf; c.alignment = ctr; c.border = thin
    
    row = 3
    # 按类型分组
    type_order = ['滴滴', '阳光出行', '通行费']
    type_counters = {t: 0 for t in type_order}
    
    for inv_type in type_order:
        recs = [r for r in all_records if r['type'] == inv_type]
        if not recs: continue
        
        if inv_type == '通行费':
            # 通行费：按行程分组
            trip_map = {}
            for r in recs:
                # 用相近日期近似分组
                key = r['trip_date']
                if key not in trip_map:
                    trip_map[key] = []
                trip_map[key].append(r)
            
            for trip_date in sorted(trip_map.keys()):
                type_counters[inv_type] += 1
                trip_recs = trip_map[trip_date]
                inv_list = '\n'.join(r['inv_no'] for r in sorted(trip_recs, key=lambda x: x['inv_no']))
                pages = ','.join(str(r['word_page']) for r in sorted(trip_recs, key=lambda x: x['word_page']))
                total_amt = sum(r['amount'] for r in trip_recs)
                route_from = trip_recs[0]['from'] or '详见发票'
                route_to = trip_recs[0]['to'] or '详见发票'
                
                for col, v in enumerate([str(type_counters[inv_type]), '通行费', inv_list, trip_date, route_from, route_to, round(total_amt,2), pages], 1):
                    c = ws.cell(row=row, column=col, value=v)
                    c.font = dfont; c.border = thin
                    c.alignment = ctr if col in (1,2,4,7,8) else lft
                    if col == 7: c.number_format = '#,##0.00'
                row += 1
        else:
            for r in recs:
                type_counters[inv_type] += 1
                for col, v in enumerate([str(type_counters[inv_type]), r['type'], r['inv_no'], r['trip_date'], r['from'], r['to'], r['amount'], r.get('word_page','')], 1):
                    c = ws.cell(row=row, column=col, value=v)
                    c.font = dfont; c.border = thin
                    c.alignment = ctr if col in (1,2,3,4,7,8) else lft
                    if col == 7: c.number_format = '#,##0.00'
                row += 1
    
    # 合计行
    ws.cell(row=row, column=1, value="合计").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws.cell(row=row, column=1).fill = hf; ws.cell(row=row, column=1).alignment = ctr; ws.cell(row=row, column=1).border = thin
    ws.merge_cells(f"A{row}:F{row}")
    total = round(sum(r['amount'] for r in all_records), 2)
    ws.cell(row=row, column=7, value=total)
    ws.cell(row=row, column=7).number_format = '#,##0.00'
    ws.cell(row=row, column=7).font = Font(name="Arial", bold=True, size=11)
    ws.cell(row=row, column=7).alignment = ctr
    ws.cell(row=row, column=7).fill = PatternFill("solid", fgColor="D6E4F0")
    ws.cell(row=row, column=7).border = thin
    
    for i, w in enumerate([6,10,30,22,26,26,13,16], 1):
        ws.column_dimensions[chr(64+i)].width = w
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 20
    for r in range(3, row+1):
        ws.row_dimensions[r].height = 36
    
    out = f'{OUTPUT_DIR}/差旅发票汇总.xlsx'
    wb.save(out)
    print(f"  已保存: {out}")
    return out

# ============ 主流程 ============
def main():
    print("=" * 50)
    print("  差旅发票整理工具")
    print("=" * 50)
    
    # 1. 读取配置
    cfg = load_config()
    
    # 2. 询问日期范围
    print(f"\n📅 请输入日期范围（格式: YYYY-MM-DD）")
    start = input(f"  开始日期（如 2026-04-01）: ").strip()
    end = input(f"  结束日期（如 2026-04-30）: ").strip()
    if not start or not end:
        print("日期不能为空，退出。")
        sys.exit(1)
    
    # 3. 扫描邮箱
    invoice_emails = scan_invoices(cfg, start, end)
    if not invoice_emails:
        print("❌ 未找到任何发票邮件。")
        sys.exit(0)
    
    # 4. 下载附件
    all_records = download_attachments(cfg, invoice_emails, start, end)
    if not all_records:
        print("❌ 未下载到任何发票。")
        sys.exit(1)
    
    # 5. 生成Word
    word_path = build_word(all_records)
    
    # 6. 生成Excel
    excel_path = build_excel(all_records)
    
    # 7. 完成
    print(f"\n✅ 完成！共整理 {len(all_records)} 张发票")
    print(f"  📂 输出目录: {OUTPUT_DIR}")
    print(f"  📊 Excel: {excel_path}")
    print(f"  📄 Word: {word_path}")
    print(f"  📁 PDF文件夹: {OUTPUT_DIR}/发票PDF/ ({len([f for f in os.listdir(f'{OUTPUT_DIR}/发票PDF') if f.endswith('.pdf')])}个)")

if __name__ == '__main__':
    main()
