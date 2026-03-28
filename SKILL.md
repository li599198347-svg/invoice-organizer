---
name: invoice-organizer
description: 差旅发票整理技能。从QQ邮箱自动提取各类报销发票（滴滴/通行费/高铁/飞机等），生成汇总Excel和Word打印版。
---

# 差旅发票整理

从QQ邮箱自动提取发票，生成报销用Excel和Word。

## 首次使用

需要配置QQ邮箱IMAP访问（只读权限）：

1. 登录 QQ 邮箱 → 设置 → 账户 → POP3/IMAP/SMTP 服务
2. 开启 IMAP/SMTP
3. 生成授权码

然后把配置写入记忆文件：

```
MEMORY.md 中添加：

## 邮件配置
IMAP_HOST=imap.qq.com
IMAP_PORT=993
IMAP_USER=你的QQ邮箱
IMAP_PASS=IMAP授权码
IMAP_MAILBOX=INBOX
```

技能运行时会自动从记忆文件读取。

## 执行

```bash
python3 scripts/main.py
```

运行时交互式输入日期范围：

```
开始日期（如 2026-04-01）: 2026-04-01
结束日期（如 2026-04-30）: 2026-04-30
```

## 自动识别发票类型

脚本自动扫描邮件发件人和主题，无需预设。已验证类型：

| 类型 | 发件人 | 关键词 |
|------|--------|--------|
| 滴滴 | didifapiao@mailgate.xiaojukeji.com | - |
| 第三方/阳光出行 | xiaojukeji.com | 第三方 |
| 通行费 | service@invoice.txffp.com | 通行费 |

## 处理逻辑

- **滴滴/阳光出行**：下载邮件PDF附件，提取发票号、金额、行程日期
- **通行费**：下载ZIP附件，解压后从XML提取真实发票号（`<EIid>`字段）
- **自动过滤**：行程报销单、汇总单不计入发票

## 输出结构

```
~/Desktop/差旅发票/
├── 发票PDF/                 # 每个PDF以发票号命名
├── 差旅发票汇总.xlsx        # Excel清单
└── 差旅发票汇总打印.docx    # Word打印版（A5横向）
```

Excel列：序号、类型、发票号码、行程日期、出发地/出发站、到达地/到达站、金额(元)、Word页码

通行费：同一趟行程的多张子票据合并为一行，发票号码列换行列出所有子票据号，Word页码列对应所有页码。

## 扩展新的发票类型

在 `scripts/main.py` 的 `scan_invoices()` 函数中找到类型识别块添加新规则：

```python
elif '你的关键词' in subject:
    inv_type = '你的类型'
```
