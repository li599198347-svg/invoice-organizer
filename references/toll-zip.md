# 通行费 ZIP 处理逻辑

## ZIP 结构

通行费邮件的ZIP附件包含以下文件：

```
pdf/
  1_<invoice_id>.pdf          # 子票据PDF
  2_<invoice_id>.pdf
  3_<invoice_id>.pdf
xml/
  1_<invoice_id>.xml           # 发票元数据
  2_<invoice_id>.xml
  3_<invoice_id>.xml
```

## 关键字段（XML解析）

```python
import zipfile, re

with zipfile.ZipFile(zip_path) as zf:
    for xml_name in zf.namelist():
        if not xml_name.startswith('xml/'): continue
        content = zf.read(xml_name).decode('utf-8', errors='replace')
        
        # 真实发票号（20位数字）
        inv_no = re.search(r'<EIid>(\d{20,})</EIid>', content)
        if not inv_no:
            inv_no = re.search(r'<InvoiceNumber>(\d{20,})</InvoiceNumber>', content)
        
        # 含税金额
        amount = re.search(r'<TotaltaxIncludedAmount>(\d+\.?\d*)</TotaltaxIncludedAmount>', content)
        
        # 开票日期
        date = re.search(r'<IssueTime>(\d{4}-\d{2}-\d{2})</IssueTime>', content)
        
        # 通行日期（原始格式）
        passage_date = re.search(r'<StartDatesOfPassage>(\d{14})</StartDatesOfPassage>', content)
```

## 汇总单（行程PDF）

每封通行费邮件还包含两个汇总PDF：
- `通行费电子票据汇总单(票据).pdf` — 按票据索引，含所有子票据汇总
- `通行费电子票据汇总单(行程).pdf` — 按行程索引，含路线信息

## 行程分组

同一趟行程可能开多张发票（汇总单号相同）。判断方法：
- ZIP文件名SEQ对应一个汇总单号
- 同一SEQ下所有子票据属于同一趟行程
- 行程路线从汇总单PDF提取
