#!/usr/bin/env python3
"""
英国合同查找器 — 招标数据爬取（通过下载csv文件）、上传与邮件通知

用法:
  python3 england_tender.py --sender 发件邮箱 --auth-code 授权码 --receiver 收件邮箱 [--attachments]
参数说明：
  --sender        发件邮箱（QQ邮箱）
  --auth-code     发件邮箱SMTP授权码
  --receiver      收件邮箱，多个邮箱用分号分隔（如：xxx@qq.com;yyy@live.com）
  --attachments   传此参数时，文件符合规则则作为邮件附件发送；不传则强制走 Gofile 链接方式
"""

import argparse
import os
import re
import smtplib
import time
import requests
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from urllib.parse import urljoin, urlparse

from playwright.sync_api import sync_playwright


# ==================== 固定配置 ====================

TARGET_URL = "https://www.contractsfinder.service.gov.uk/Search/Results"
DOWNLOAD_BUTTON_SELECTOR = "#get_csv_file"
GOFILE_API = "https://api.gofile.io"
VIEWPORT = {"width": 1280, "height": 800}
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

SAVE_DIR = os.path.join(os.getcwd(), "英国招标")

# QQ邮箱附件限制
QQ_MAX_SINGLE_FILE_SIZE = 20 * 1024 * 1024  # 20MB
QQ_MAX_TOTAL_SIZE = 50 * 1024 * 1024  # 50MB

# ==================== 工具函数 ====================

def extract_filename_from_response(resp):
    """从 HTTP 响应头 Content-Disposition 提取文件名"""
    cd = resp.headers.get("Content-Disposition", "")
    if cd:
        match = re.search(
            r'filename[^;=\n]*=(["\']?)(.+?)\1(;|$)', cd, re.IGNORECASE
        )
        if match:
            return match.group(2).strip()
    return None


def download_with_requests(url, save_path, timeout=120):
    """使用 requests 下载文件（优先方式）"""
    headers = {"User-Agent": USER_AGENT}
    resp = requests.get(
        url, stream=True, timeout=timeout, headers=headers, allow_redirects=True
    )
    resp.raise_for_status()

    filename = extract_filename_from_response(resp)
    if not filename:
        filename = os.path.basename(urlparse(resp.url).path)

    with open(save_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    return filename


def create_gofile_account():
    """创建 Gofile 匿名账户，返回 token 和 rootFolderId"""
    resp = requests.post(f"{GOFILE_API}/accounts")
    data = resp.json()["data"]
    return data["token"], data["rootFolder"]


def upload_to_gofile(filepath, token, folder_id, max_retries=2):
    """上传单个文件到 Gofile，支持重试"""
    filename = os.path.basename(filepath)
    filesize = os.path.getsize(filepath)
    if filesize >= 1024 * 1024:  # >= 1MB
        size_str = f"{filesize / 1024 / 1024:.1f} MB"
    else:
        size_str = f"{filesize / 1024:.1f} KB"
    print(f"  上传: {filename} ({size_str})")

    last_error = None
    for attempt in range(max_retries + 1):
        try:
            if attempt > 0:
                print(f"  第 {attempt + 1} 次重试上传...")

            # 获取最佳服务器
            resp = requests.get(f"{GOFILE_API}/servers", params={"token": token})
            servers = resp.json()["data"]["servers"]
            na_servers = [s["name"] for s in servers if s["zone"] == "na"]
            server = na_servers[0] if na_servers else servers[0]["name"]

            # 上传
            upload_url = f"https://{server}.gofile.io/contents/uploadfile"
            start_time = time.time()
            with open(filepath, "rb") as f:
                files = {"file": (filename, f)}
                data = {"token": token, "folderId": folder_id}
                resp = requests.post(upload_url, files=files, data=data, timeout=1800)

            elapsed = time.time() - start_time
            result = resp.json()
            if result.get("status") != "ok":
                raise Exception(f"上传失败: {result}")

            download_page = result["data"]["downloadPage"]
            print(f"  上传成功! ({elapsed:.1f}s) 下载页: {download_page}")

            return {
                "filename": filename,
                "downloadPage": download_page,
                "fileId": result["data"]["id"],
                "size": filesize,
            }
        except Exception as e:
            last_error = e
            print(f"  上传失败(尝试 {attempt + 1}/{max_retries + 1}): {e}")
            if attempt < max_retries:
                time.sleep(3)  # 重试前等待3秒

    raise Exception(f"上传失败，已重试 {max_retries} 次: {last_error}")


def set_gofile_folder_public(folder_id, token):
    """设置 Gofile 文件夹为公开访问（关键步骤！）"""
    resp = requests.put(
        f"{GOFILE_API}/contents/{folder_id}/update",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"attribute": "public", "attributeValue": True},
    )
    return resp.json()


def send_email_with_gofile_link(subject, sender, auth_code, recipient, download_link, upload_results):
    """
    发送包含Gofile下载链接的邮件（HTML格式）
    recipient: 单个邮箱字符串，或多个邮箱用分号分隔的字符串
    """
    # 解析多个收件人（用分号分隔）
    if ';' in recipient:
        recipient_list = [r.strip() for r in recipient.split(';') if r.strip()]
    else:
        recipient_list = [recipient.strip()]
    
    print(f"\n发送邮件到 {len(recipient_list)} 个收件人: {', '.join(recipient_list)}...")

    # 文件清单表格行
    file_rows_html = ""
    file_rows_text = ""
    for i, r in enumerate(upload_results, 1):
        bg = "#f2f2f2" if i % 2 == 1 else "#ffffff"
        size_bytes = r['size']
        if size_bytes >= 1024 * 1024:  # >= 1MB
            size_str = f"{size_bytes/1024/1024:.1f} MB"
        else:
            size_str = f"{size_bytes/1024:.1f} KB"
        file_rows_html += f"""<tr style="background-color: {bg};">
<td style="border: 1px solid #ddd; padding: 10px;">{i}</td>
<td style="border: 1px solid #ddd; padding: 10px;">{r['filename']}</td>
<td style="border: 1px solid #ddd; padding: 10px;">{size_str}</td></tr>
"""
        file_rows_text += f"{i}. {r['filename']} ({size_str})\n"

    tz_beijing = timezone(timedelta(hours=8))
    now_str = datetime.now(tz_beijing).strftime('%Y-%m-%d %H:%M:%S')

    html_body = f"""<html><body style="font-family: Microsoft YaHei, Arial, sans-serif; font-size: 15px; color: #333; line-height: 1.8; max-width: 700px; margin: 0 auto;">
<h2 style="color: #1a5276; border-bottom: 2px solid #2980b9; padding-bottom: 8px; font-size: 20px;">英国合同查找器 - 招标数据爬取结果</h2>
<div style="background-color: #eaf2f8; padding: 15px; border-radius: 5px; margin: 15px 0;">
<strong>爬取时间：</strong>{now_str}<br>
<strong>数据来源：</strong>https://www.contractsfinder.service.gov.uk<br>
<strong>总计：</strong>{len(upload_results)}个文件
</div>
<h3 style="font-size: 17px;">下载链接</h3>
<div style="background-color: #fff3cd; border: 1px solid #ffc107; padding: 20px; border-radius: 5px; text-align: center; margin: 15px 0;">
<a href="{download_link}" style="font-size: 18px; color: #2980b9; font-weight: bold; text-decoration: none;">点击下载全部文件</a>
<br><span style="color: #888; font-size: 13px;">{download_link}</span>
<br><span style="color: #666; font-size: 12px;">（{len(upload_results)}个文件均在同一目录下，可逐个下载）</span>
</div>
<h3>文件清单</h3>
<table style="border-collapse: collapse; width: 100%;">
<tr style="background-color: #2980b9; color: white;"><th style="padding: 10px; text-align: left;">序号</th><th style="padding: 10px; text-align: left;">文件名</th><th style="padding: 10px; text-align: left;">大小</th></tr>
{file_rows_html}
</table>
<div style="background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 5px; margin: 20px 0;">
<strong style="color: #721c24;">提醒：Gofile.io 链接有效期有限，建议尽快下载。</strong>
</div>
<p style="color:#888; font-size:12px; margin-top:20px;">此邮件由自动化爬取系统发送，请勿直接回复。</p>
</body></html>"""

    text_body = f"""英国合同查找器 - 招标数据爬取结果

爬取时间：{now_str}
数据来源：https://www.contractsfinder.service.gov.uk
总计：{len(upload_results)}个文件

下载链接（{len(upload_results)}个文件在同一目录下，可逐个下载）：
{download_link}

文件清单：
{file_rows_text}
提醒：Gofile.io 链接有效期有限，建议尽快下载。
"""

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = ", ".join(recipient_list)  # 多个收件人用逗号分隔
    msg["Subject"] = subject

    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        server = smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(sender, auth_code)
        server.sendmail(sender, recipient_list, msg.as_string())
        server.quit()
        print("  [OK] 邮件发送成功！")
        return True
    except Exception as e:
        # 尝试TLS方式
        try:
            server = smtplib.SMTP("smtp.qq.com", 587)
            server.starttls()
            server.login(sender, auth_code)
            server.sendmail(sender, recipient_list, msg.as_string())
            server.quit()
            print("  [OK] 邮件发送成功(TLS)！")
            return True
        except Exception as e2:
            print(f"  [FAIL] 邮件发送失败: {e2}")
            return False


def send_email_with_attachments(subject, sender, auth_code, recipient, files):
    """
    发送带附件的邮件（符合QQ邮箱规则）
    recipient: 单个邮箱字符串，或多个邮箱用分号分隔的字符串
    """
    # 解析多个收件人（用分号分隔）
    if ';' in recipient:
        recipient_list = [r.strip() for r in recipient.split(';') if r.strip()]
    else:
        recipient_list = [recipient.strip()]
    
    print(f"\n发送邮件到 {len(recipient_list)} 个收件人: {', '.join(recipient_list)}（带附件）...")

    # 文件清单
    file_rows_text = ""
    total_size = 0
    for i, (filepath, filename, filesize) in enumerate(files, 1):
        total_size += filesize
        if filesize >= 1024 * 1024:  # >= 1MB
            size_str = f"{filesize/1024/1024:.1f} MB"
        else:
            size_str = f"{filesize/1024:.1f} KB"
        file_rows_text += f"{i}. {filename} ({size_str})\n"

    tz_beijing = timezone(timedelta(hours=8))
    now_str = datetime.now(tz_beijing).strftime('%Y-%m-%d %H:%M:%S')

    html_body = f"""<html><body style="font-family: Microsoft YaHei, Arial, sans-serif; font-size: 15px; color: #333; line-height: 1.8; max-width: 700px; margin: 0 auto;">
<h2 style="color: #1a5276; border-bottom: 2px solid #2980b9; padding-bottom: 8px; font-size: 20px;">英国合同查找器 - 招标数据爬取结果</h2>
<div style="background-color: #eaf2f8; padding: 15px; border-radius: 5px; margin: 15px 0;">
<strong>爬取时间：</strong>{now_str}<br>
<strong>数据来源：</strong>https://www.contractsfinder.service.gov.uk<br>
<strong>总计：</strong>{len(files)}个文件（作为附件发送）
</div>
<h3 style="font-size: 17px;">文件清单</h3>
<pre style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; font-family: Consolas, Monaco, monospace;">{file_rows_text}</pre>
<div style="background-color: #d4edda; border: 1px solid #c3e6cb; padding: 15px; border-radius: 5px; margin: 20px 0;">
<strong style="color: #155724;">✓ 文件已作为附件发送，可直接下载。</strong>
</div>
<p style="color:#888; font-size:12px; margin-top:20px;">此邮件由自动化爬取系统发送，请勿直接回复。</p>
</body></html>"""

    text_body = f"""英国合同查找器 - 招标数据爬取结果

爬取时间：{now_str}
数据来源：https://www.contractsfinder.service.gov.uk
总计：{len(files)}个文件（作为附件发送）

文件清单：
{file_rows_text}
文件已作为附件发送，可直接下载。
"""

    msg = MIMEMultipart("mixed")
    msg["From"] = sender
    msg["To"] = ", ".join(recipient_list)  # 多个收件人用逗号分隔
    msg["Subject"] = subject

    # 正文部分（纯文本 + HTML）
    body = MIMEMultipart("alternative")
    body.attach(MIMEText(text_body, "plain", "utf-8"))
    body.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(body)

    # 添加附件
    for filepath, filename, filesize in files:
        try:
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                # 正确设置附件文件名
                part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', filename))
                msg.attach(part)
            print(f"  已添加附件: {filename}")
        except Exception as e:
            print(f"  添加附件失败: {filename} - {e}")

    try:
        server = smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(sender, auth_code)
        server.sendmail(sender, recipient_list, msg.as_string())
        server.quit()
        print("  [OK] 邮件发送成功（带附件）！")
        return True
    except Exception as e:
        # 尝试TLS方式
        try:
            server = smtplib.SMTP("smtp.qq.com", 587)
            server.starttls()
            server.login(sender, auth_code)
            server.sendmail(sender, recipient_list, msg.as_string())
            server.quit()
            print("  [OK] 邮件发送成功（带附件，TLS）！")
            return True
        except Exception as e2:
            print(f"  [FAIL] 邮件发送失败: {e2}")
            return False


# ==================== 主流程 ====================

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description="英国合同查找器 — 招标数据爬取、上传与邮件通知"
    )
    parser.add_argument(
        "--sender", required=True, help="发件邮箱（QQ 邮箱）"
    )
    parser.add_argument(
        "--auth-code", required=True, help="发件邮箱 SMTP 授权码"
    )
    parser.add_argument(
        "--receiver", required=True, help="收件邮箱"
    )
    parser.add_argument(
        "--attachments", action="store_true",
        help="传此参数时，文件符合规则则作为邮件附件发送；不传则强制走 Gofile 链接方式"
    )
    args = parser.parse_args()

    sender = args.sender
    auth_code = args.auth_code
    recipient = args.receiver

    os.makedirs(SAVE_DIR, exist_ok=True)
    print(f"保存目录: {SAVE_DIR}\n")

    # ====== 阶段一：浏览器爬取与下载 ======

    with sync_playwright() as p:
        # 步骤 1：初始化浏览器
        print("[步骤1] 初始化浏览器...")
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            viewport=VIEWPORT,
            user_agent=USER_AGENT,
            bypass_csp=True,
            accept_downloads=True,
        )
        page = context.new_page()
        page.set_default_timeout(120000)
        page.set_default_navigation_timeout(120000)

        # 步骤 2：访问目标页面（增加重试机制）
        print("[步骤2] 访问目标页面...")
        max_retries = 2
        page_loaded = False
        for attempt in range(max_retries + 1):
            try:
                if attempt > 0:
                    print(f"  第 {attempt + 1} 次重试访问...")
                page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=120000)
                page.wait_for_load_state("load", timeout=60000)
                print("  页面加载完成")
                page_loaded = True
                break
            except Exception as e:
                print(f"  页面加载失败(尝试 {attempt + 1}/{max_retries + 1}): {e}")
                if attempt < max_retries:
                    time.sleep(5)
                else:
                    print("  页面加载失败，程序退出")
                    browser.close()
                    return

        # 步骤 3：点击下载按钮下载 CSV
        print("[步骤3] 点击下载按钮下载 CSV...")
        
        tz_beijing = timezone(timedelta(hours=8))
        timestamp = datetime.now(tz_beijing).strftime("%Y%m%d")
        new_filename = f"英国-招标-合同数据-{timestamp}.csv"
        save_path = os.path.join(SAVE_DIR, new_filename)
        
        try:
            # 等待下载按钮出现
            download_button = page.wait_for_selector(DOWNLOAD_BUTTON_SELECTOR, timeout=30000)
            if download_button:
                print(f"  找到下载按钮，准备下载...")
                # 使用 Playwright 的下载功能
                with page.expect_download(timeout=120000) as download_info:
                    download_button.click()
                
                download = download_info.value
                download.save_as(save_path)
                print(f"  下载成功: {new_filename}")
            else:
                print("  未找到下载按钮")
        except Exception as e:
            print(f"  下载失败: {e}")
            print("  尝试使用 requests 直接下载...")
            try:
                # 尝试获取下载链接
                download_button = page.query_selector(DOWNLOAD_BUTTON_SELECTOR)
                if download_button:
                    href = download_button.get_attribute("href")
                    if href:
                        if not href.startswith("http"):
                            href = urljoin(TARGET_URL, href)
                        download_with_requests(href, save_path)
                        print(f"  下载成功: {new_filename}")
            except Exception as e2:
                print(f"  requests 下载也失败: {e2}")

        browser.close()

    # ====== 阶段二：判断发送方式 ======

    # 获取所有下载的文件及其大小
    files = []
    for f in os.listdir(SAVE_DIR):
        filepath = os.path.join(SAVE_DIR, f)
        if os.path.isfile(filepath):
            filesize = os.path.getsize(filepath)
            files.append((filepath, f, filesize))

    total_size = sum(f[2] for f in files)

    # 判断发送方式
    # - 不传 --attachments：强制走 Gofile 方式
    # - 传 --attachments：按现有规则判断
    if not args.attachments:
        # 强制走 Gofile 方式
        can_send_as_attachment = False
    else:
        # 检查是否符合QQ邮箱附件规则
        can_send_as_attachment = True
        
        # 检查单个文件大小
        for filepath, filename, filesize in files:
            if filesize > QQ_MAX_SINGLE_FILE_SIZE:
                can_send_as_attachment = False
                print(f"\n  ⚠️ 文件 {filename} 大小 {filesize/1024/1024:.1f}MB 超过20MB限制")
                break
        
        # 检查总大小
        if can_send_as_attachment and total_size > QQ_MAX_TOTAL_SIZE:
            can_send_as_attachment = False
            print(f"\n  ⚠️ 文件总大小 {total_size/1024/1024:.1f}MB 超过50MB限制")

    if can_send_as_attachment:
        # 方式1：作为邮件附件发送
        print("\n[步骤4] 发送邮件（带附件）...")
        print(f"  文件总大小: {total_size/1024/1024:.1f}MB，符合QQ邮箱附件规则")
        
        try:
            send_email_with_attachments(
                subject="英国合同查找器 - 招标数据",
                sender=sender,
                auth_code=auth_code,
                recipient=recipient,
                files=files,
            )
        except Exception as e:
            print(f"  邮件发送失败: {e}")

        # 汇总
        print(f"\n{'='*60}")
        print(f"完成! 共下载 {len(files)} 个文件")
        print(f"文件已作为附件发送")
        print(f"保存目录: {SAVE_DIR}")

    else:
        # 方式2：上传到Gofile，邮件发送下载链接
        print("\n[步骤4] 上传文件到 Gofile.io...")
        token, root_folder = create_gofile_account()
        print(f"  Gofile 账户已创建")

        # 按文件大小排序，小文件先上传
        files_sorted = sorted(files, key=lambda x: x[2])
        
        upload_results = []
        for filepath, filename, filesize in files_sorted:
            try:
                result = upload_to_gofile(filepath, token, root_folder)
                upload_results.append(result)
            except Exception as e:
                print(f"  上传失败: {filename} - {e}")

        # 关键：设置文件夹公开访问
        print("\n  设置文件夹为公开访问...")
        set_gofile_folder_public(root_folder, token)
        print("  文件夹已设为公开")

        # 发送邮件
        print("\n[步骤5] 发送邮件...")
        download_page = upload_results[0]["downloadPage"] if upload_results else "N/A"

        try:
            send_email_with_gofile_link(
                subject="英国合同查找器 - 招标数据 Gofile 下载链接",
                sender=sender,
                auth_code=auth_code,
                recipient=recipient,
                download_link=download_page,
                upload_results=upload_results,
            )
        except Exception as e:
            print(f"  邮件发送失败: {e}")

        # 汇总
        print(f"\n{'='*60}")
        print(f"完成! 共下载 {len(files)} 个文件，上传 {len(upload_results)} 个")
        print(f"Gofile 链接: {download_page}")
        print(f"保存目录: {SAVE_DIR}")


if __name__ == "__main__":
    main()
