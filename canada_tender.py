#!/usr/bin/env python3
"""
加拿大开放数据平台 — 招标数据爬取、上传与邮件通知

用法:
  python3 canada_tender.py --sender 发件邮箱 --auth-code 授权码 --receiver 收件邮箱 [--full]

  不传 --full 时，仅下载 "New tender notices" 和 "Open tender notices" 两份 CSV；
  传入 --full 时，爬取全量资源。
"""

import argparse
import os
import re
import smtplib
import sys
import time
import requests
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from urllib.parse import urljoin, urlparse

from playwright.sync_api import sync_playwright


# ==================== 固定配置 ====================

TARGET_URL = (
    "https://open.canada.ca/data/en/dataset/"
    "6abd20d4-7a1c-4b38-baa2-9525d0bb2fd2/"
    "resource/05b804dd-11ec-4271-8d69-d6044e1a5481"
)
EXCLUDE_KEYWORDS = ["supporting documentation", "data dictionary"]
GOFILE_API = "https://api.gofile.io"
VIEWPORT = {"width": 1280, "height": 800}
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

SAVE_DIR = os.path.join(os.getcwd(), "加拿大招标")

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
    print(f"  上传: {filename} ({filesize / 1024 / 1024:.1f} MB)")

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


def send_email(subject, sender, auth_code, recipient, download_link, upload_results):
    """发送包含下载链接的邮件（HTML格式）"""
    print(f"\n发送邮件到 {recipient}...")

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

    now_str = datetime.now().strftime('%Y年%m月%d日')

    html_body = f"""<html><body style="font-family: Microsoft YaHei, Arial, sans-serif; color: #333; line-height: 1.8; max-width: 700px; margin: 0 auto;">
<h2 style="color: #1a5276; border-bottom: 2px solid #2980b9; padding-bottom: 8px;">加拿大开放数据平台 - 招标数据爬取结果</h2>
<div style="background-color: #eaf2f8; padding: 15px; border-radius: 5px; margin: 15px 0;">
<strong>爬取时间：</strong>{now_str}<br>
<strong>数据来源：</strong>https://open.canada.ca<br>
<strong>总计：</strong>{len(upload_results)}个文件
</div>
<h3>下载链接</h3>
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

    text_body = f"""加拿大开放数据平台 - 招标数据爬取结果

爬取时间：{now_str}
数据来源：https://open.canada.ca
总计：{len(upload_results)}个文件

下载链接（{len(upload_results)}个文件在同一目录下，可逐个下载）：
{download_link}

文件清单：
{file_rows_text}
提醒：Gofile.io 链接有效期有限，建议尽快下载。
"""

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject

    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        server = smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(sender, auth_code)
        server.sendmail(sender, recipient, msg.as_string())
        server.quit()
        print("  [OK] 邮件发送成功！")
        return True
    except Exception as e:
        # 尝试TLS方式
        try:
            server = smtplib.SMTP("smtp.qq.com", 587)
            server.starttls()
            server.login(sender, auth_code)
            server.sendmail(sender, recipient, msg.as_string())
            server.quit()
            print("  [OK] 邮件发送成功(TLS)！")
            return True
        except Exception as e2:
            print(f"  [FAIL] 邮件发送失败: {e2}")
            return False


# ==================== 主流程 ====================

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description="加拿大开放数据平台 — 招标数据爬取、上传与邮件通知"
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
        "--full", action="store_true", default=False,
        help="爬取全量资源（默认仅下载 New tender notices 和 Open tender notices）"
    )
    args = parser.parse_args()

    sender = args.sender
    auth_code = args.auth_code
    recipient = args.receiver
    full_mode = args.full

    if full_mode:
        print("模式: 全量爬取")
    else:
        print("模式: 仅下载 New tender notices 和 Open tender notices")

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

        # 步骤 3：点击【Show more】展开资源列表
        print("[步骤3] 尝试点击【Show more】...")
        try:
            show_more = page.wait_for_selector(
                "#show-all-resources--label--more", timeout=10000
            )
            if show_more:
                show_more.click()
                print("  已点击 Show more 按钮")
                time.sleep(3)
        except Exception:
            print("  未找到 Show more 按钮，列表可能已展开，跳过")

        # 步骤 4：下载当前页面自身的资源
        print("[步骤4] 下载当前页面资源...")
        go_to_resource_links = page.query_selector_all("a.resource-url-analytics")
        print(f"  找到 {len(go_to_resource_links)} 个 Go to resource 链接")

        for link in go_to_resource_links:
            href = link.get_attribute("href")
            if not href:
                continue
            if not href.startswith("http"):
                href = urljoin(TARGET_URL, href)

            # CSV 过滤
            if not href.lower().endswith(".csv"):
                print(f"  跳过非 CSV 文件: {href}")
                continue

            save_path = os.path.join(SAVE_DIR, "加拿大-招标-New tender notices.csv")
            try:
                print(f"  requests 下载: {href}")
                original_name = download_with_requests(href, save_path)
                print(f"  下载成功! 原始文件名: {original_name}")
            except Exception as e:
                print(f"  requests 下载失败: {e}, 尝试浏览器下载...")
                try:
                    with page.expect_download(timeout=120000) as download_info:
                        link.click()
                    download_info.value.save_as(save_path)
                    print("  浏览器下载成功!")
                except Exception as e2:
                    print(f"  浏览器下载也失败: {e2}")

        # 步骤 5：获取 Resources 列表中的全部资源链接
        print("[步骤5] 获取 Resources 列表...")
        resources_panel_selector = (
            "#wb-auto-5 > aside > div > section > "
            "div.panel-body.resources-side-panel.resources-side-panel-no-edit"
        )
        resource_links = []

        try:
            panel = page.wait_for_selector(resources_panel_selector, timeout=10000)
            if panel:
                links = panel.query_selector_all("a[href]")
                for link in links:
                    href = link.get_attribute("href")
                    text = link.inner_text().strip()
                    if not href:
                        continue
                    if not href.startswith("http"):
                        href = urljoin(TARGET_URL, href)
                    if "/resource/" not in href:
                        continue
                    text_lower = text.lower()
                    skip = any(kw.lower() in text_lower for kw in EXCLUDE_KEYWORDS)
                    if skip:
                        print(f"  跳过: {text}")
                        continue
                    # 非全量模式下，只收集 "Open tender notices"
                    if not full_mode and "open tender notices" not in text_lower:
                        print(f"  跳过（非全量模式）: {text}")
                        continue
                    resource_links.append({"text": text, "href": href})
                    print(f"  收集: {text}")
        except Exception as e:
            print(f"  面板选择器未找到: {e}，尝试备选方案...")
            all_aside_links = page.query_selector_all("aside a[href*='/resource/']")
            for link in all_aside_links:
                href = link.get_attribute("href")
                text = link.inner_text().strip()
                if not href:
                    continue
                if not href.startswith("http"):
                    href = urljoin(TARGET_URL, href)
                skip = any(kw.lower() in text.lower() for kw in EXCLUDE_KEYWORDS)
                if skip:
                    continue
                # 非全量模式下，只收集 "Open tender notices"
                if not full_mode and "open tender notices" not in text.lower():
                    continue
                resource_links.append({"text": text, "href": href})

        print(f"  共收集 {len(resource_links)} 个资源链接")

        # 步骤 6：逐个访问资源页面并下载 CSV
        print("\n[步骤6] 逐个下载资源 CSV...")
        for i, res in enumerate(resource_links):
            print(f"\n  处理 {i+1}/{len(resource_links)}: {res['text']}")
            
            # 访问资源页面，增加重试机制
            max_retries = 2
            page_loaded = False
            for attempt in range(max_retries + 1):
                try:
                    if attempt > 0:
                        print(f"    第 {attempt + 1} 次重试访问页面...")
                    page.goto(
                        res["href"],
                        wait_until="domcontentloaded",
                        timeout=120000,
                    )
                    page.wait_for_load_state("load", timeout=60000)
                    time.sleep(2)
                    page_loaded = True
                    break
                except Exception as e:
                    print(f"    访问资源页面失败(尝试 {attempt + 1}/{max_retries + 1}): {e}")
                    if attempt < max_retries:
                        time.sleep(5)  # 重试前等待5秒
                    else:
                        print(f"    已重试 {max_retries} 次，跳过此资源")
            
            if not page_loaded:
                continue

            go_links = page.query_selector_all("a.resource-url-analytics")
            for link in go_links:
                href = link.get_attribute("href")
                if not href:
                    continue
                if not href.startswith("http"):
                    href = urljoin(res["href"], href)

                # 支持多种格式
                lower_href = href.lower()
                if lower_href.endswith(".csv"):
                    ext = "csv"
                elif lower_href.endswith(".xlsx"):
                    ext = "xlsx"
                elif lower_href.endswith(".xls"):
                    ext = "xls"
                else:
                    continue  # 跳过其他格式

                resource_name = re.sub(
                    r'[<>:"/\\|?*]', "_", res["text"].strip()
                )
                new_filename = f"加拿大-招标-{resource_name}.{ext}"
                save_path = os.path.join(SAVE_DIR, new_filename)

                try:
                    download_with_requests(href, save_path)
                    print(f"    下载成功: {new_filename}")
                except Exception as e:
                    print(f"    requests 失败: {e}")
                    try:
                        with page.expect_download(timeout=60000) as download_info:
                            link.click()
                        download_info.value.save_as(save_path)
                        print(f"    浏览器下载成功: {new_filename}")
                    except Exception as e2:
                        print(f"    浏览器下载也失败: {e2}")

        browser.close()

    # ====== 阶段二：上传到 Gofile ======

    print("\n[步骤7] 上传文件到 Gofile.io...")
    token, root_folder = create_gofile_account()
    print(f"  Gofile 账户已创建")

    files = sorted(
        [
            (f, os.path.getsize(os.path.join(SAVE_DIR, f)))
            for f in os.listdir(SAVE_DIR)
            if os.path.isfile(os.path.join(SAVE_DIR, f))
        ],
        key=lambda x: x[1],  # 小文件先上传
    )

    upload_results = []
    for filename, filesize in files:
        filepath = os.path.join(SAVE_DIR, filename)
        try:
            result = upload_to_gofile(filepath, token, root_folder)
            upload_results.append(result)
        except Exception as e:
            print(f"  上传失败: {filename} - {e}")

    # 关键：设置文件夹公开访问
    print("\n  设置文件夹为公开访问...")
    set_gofile_folder_public(root_folder, token)
    print("  文件夹已设为公开")

    # ====== 阶段三：发送邮件 ======

    print("\n[步骤8] 发送邮件...")
    download_page = upload_results[0]["downloadPage"] if upload_results else "N/A"

    try:
        send_email(
            subject="加拿大开放数据平台 - 招标数据 Gofile 下载链接",
            sender=sender,
            auth_code=auth_code,
            recipient=recipient,
            download_link=download_page,
            upload_results=upload_results,
        )
    except Exception as e:
        print(f"  邮件发送失败: {e}")

    # ====== 汇总 ======

    print(f"\n{'='*60}")
    print(f"完成! 共下载 {len(files)} 个文件，上传 {len(upload_results)} 个")
    print(f"Gofile 链接: {download_page}")
    print(f"保存目录: {SAVE_DIR}")


if __name__ == "__main__":
    main()
