"""
通过浏览器自动化登录 Outlook 读取验证码
"""

import time
import re
from patchright.sync_api import sync_playwright


def get_otp_from_outlook_browser(email, password, proxy=None, timeout=90):
    """
    使用浏览器登录 Outlook 并读取验证码邮件

    Args:
        email: Outlook 邮箱地址
        password: Outlook 密码
        proxy: 代理地址
        timeout: 超时时间（秒）

    Returns:
        验证码字符串，失败返回 None
    """
    playwright = None
    browser = None

    try:
        playwright = sync_playwright().start()

        proxy_settings = {
            "server": proxy,
            "bypass": "localhost",
        } if proxy else None

        browser = playwright.chromium.launch(
            headless=True,
            proxy=proxy_settings
        )

        context = browser.new_context()
        page = context.new_page()

        # 登录 Outlook
        print(f"  [Outlook] 登录 {email}...")
        page.goto("https://outlook.live.com/mail/0/", timeout=60000)
        time.sleep(3)

        # 输入邮箱
        try:
            email_input = page.locator('input[type="email"]').first
            email_input.fill(email, timeout=10000)
            time.sleep(1)

            # 点击下一步
            page.locator('input[type="submit"]').first.click(timeout=10000)
            time.sleep(3)
        except Exception as e:
            print(f"  [Outlook] 输入邮箱失败: {e}")
            return None

        # 输入密码
        try:
            password_input = page.locator('input[type="password"]').first
            password_input.fill(password, timeout=10000)
            time.sleep(1)

            # 点击登录
            page.locator('input[type="submit"]').first.click(timeout=10000)
            time.sleep(5)
        except Exception as e:
            print(f"  [Outlook] 输入密码失败: {e}")
            return None

        # 跳过"保持登录"提示
        try:
            page.locator('input[type="submit"]').first.click(timeout=5000)
            time.sleep(3)
        except:
            pass

        # 等待邮件列表加载
        print(f"  [Outlook] 等待验证码邮件...")
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 刷新邮件列表
                page.reload(timeout=30000)
                time.sleep(3)

                # 查找来自 OpenAI 的邮件
                email_items = page.locator('[role="listitem"]').all()

                for item in email_items[:10]:  # 只检查前10封
                    try:
                        text = item.inner_text(timeout=2000)

                        # 检查是否是 OpenAI 的验证码邮件
                        if "openai" in text.lower() or "chatgpt" in text.lower():
                            # 点击邮件
                            item.click(timeout=5000)
                            time.sleep(2)

                            # 读取邮件内容
                            email_body = page.locator('[role="document"]').first.inner_text(timeout=10000)

                            # 提取验证码
                            match = re.search(r'\b(\d{6})\b', email_body)
                            if match:
                                code = match.group(1)
                                print(f"  [Outlook] ✓ 找到验证码: {code}")
                                return code

                            match = re.search(r'\b(\d{8})\b', email_body)
                            if match:
                                code = match.group(1)
                                print(f"  [Outlook] ✓ 找到验证码: {code}")
                                return code
                    except:
                        continue

                # 等待一段时间再检查
                time.sleep(5)

            except Exception as e:
                print(f"  [Outlook] 检查邮件时出错: {e}")
                time.sleep(5)

        print(f"  [Outlook] ✗ 超时未找到验证码")
        return None

    except Exception as e:
        print(f"  [Outlook] 异常: {e}")
        return None

    finally:
        try:
            if browser:
                browser.close()
        except:
            pass
        try:
            if playwright:
                playwright.stop()
        except:
            pass
