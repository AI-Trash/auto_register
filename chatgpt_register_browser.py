"""
ChatGPT 批量自动注册工具 - 基于 Patchright 浏览器版
依赖: pip install patchright requests
功能: 使用 Outlook 邮箱 + Patchright 浏览器自动注册 ChatGPT 账号
"""

import os
import re
import json
import random
import time
import threading
from patchright.sync_api import sync_playwright
from outlook_mail_reader import OutlookAccountPool, OutlookMailReader


# ================= 加载配置 =================
def _load_config():
    """从 config.json 加载配置"""
    config = {
        "outlook_accounts_file": "../第二组100个邮箱_副本.txt",
        "proxy": "http://127.0.0.1:7897",
        "output_file": "openai_accounts.txt",
        "batch_size": 5,
        "headless": False,
        "wait_timeout": 30000,
        "manual_otp": False,
        "auto_continue": True,
    }

    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_browser.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                file_config = json.load(f)
                config.update(file_config)
        except Exception as e:
            print(f"⚠️ 加载 config_browser.json 失败: {e}")

    return config


_CONFIG = _load_config()
OUTLOOK_ACCOUNTS_FILE = _CONFIG["outlook_accounts_file"]
DEFAULT_PROXY = _CONFIG["proxy"]
DEFAULT_OUTPUT_FILE = _CONFIG["output_file"]
HEADLESS = _CONFIG["headless"]
WAIT_TIMEOUT = _CONFIG["wait_timeout"]
MANUAL_OTP = _CONFIG.get("manual_otp", False)
BATCH_SIZE = _CONFIG.get("batch_size", 5)
AUTO_CONTINUE = _CONFIG.get("auto_continue", True)

# 全局账号池
ACCOUNT_POOL = None
ACCOUNT_POOL_LOCK = threading.Lock()


def init_account_pool():
    """初始化账号池"""
    global ACCOUNT_POOL
    with ACCOUNT_POOL_LOCK:
        if ACCOUNT_POOL is None:
            ACCOUNT_POOL = OutlookAccountPool(OUTLOOK_ACCOUNTS_FILE)


def generate_password_from_outlook(outlook_password):
    """
    根据 Outlook 密码生成 ChatGPT 密码
    不足12位在后面加8
    """
    if len(outlook_password) >= 12:
        return outlook_password
    else:
        return outlook_password + "8" * (12 - len(outlook_password))


def generate_random_name():
    """生成随机英文姓名"""
    first_names = ["James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph",
                   "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Barbara", "Susan", "Jessica"]
    last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
                  "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson"]

    first = random.choice(first_names)
    last = random.choice(last_names)
    return f"{first} {last}"


def generate_random_birthday():
    """生成随机生日 (18-60岁之间)"""
    import datetime
    today = datetime.date.today()
    # 18到60岁之间
    min_age = 18
    max_age = 60

    birth_year = today.year - random.randint(min_age, max_age)
    birth_month = random.randint(1, 12)

    # 确保日期有效
    if birth_month in [1, 3, 5, 7, 8, 10, 12]:
        birth_day = random.randint(1, 31)
    elif birth_month in [4, 6, 9, 11]:
        birth_day = random.randint(1, 30)
    else:  # 2月
        birth_day = random.randint(1, 28)

    return birth_year, birth_month, birth_day


def register_chatgpt_with_browser(mail_reader: OutlookMailReader, proxy=None, headless=False):
    """
    使用浏览器注册 ChatGPT 账号

    Args:
        mail_reader: Outlook 邮件读取器
        proxy: 代理地址
        headless: 是否无头模式

    Returns:
        (success, email, password, error_msg)
    """
    email = mail_reader.email
    outlook_password = mail_reader.password
    chatgpt_password = generate_password_from_outlook(outlook_password)

    playwright = None
    browser = None
    context = None

    try:
        # 启动浏览器
        print(f"[{email}] 启动浏览器...")
        playwright = sync_playwright().start()

        proxy_settings = {
            "server": proxy,
            "bypass": "localhost",
        } if proxy else None

        browser = playwright.chromium.launch(
            headless=headless,
            args=['--lang=zh-CN'],
            proxy=proxy_settings
        )

        context = browser.new_context()
        page = context.new_page()

        # Step 1: 访问注册页面
        print(f"[{email}] Step 1: 访问注册页面...")
        page.goto("https://chatgpt.com/", timeout=60000)

        # 等待页面加载
        time.sleep(3)

        # 点击注册按钮（支持中英文）
        try:
            # 尝试中文
            page.get_by_text("免费注册").click(timeout=10000)
        except:
            try:
                # 尝试英文
                page.get_by_text("Sign up").click(timeout=10000)
            except:
                try:
                    page.locator('button:has-text("Sign up")').click(timeout=10000)
                except:
                    return False, email, chatgpt_password, "未找到注册按钮"

        time.sleep(2)

        # Step 2: 输入邮箱
        print(f"[{email}] Step 2: 输入邮箱...")
        try:
            # 查找邮箱输入框
            email_input = page.locator('input[type="email"]').first
            email_input.fill(email, timeout=10000)
            time.sleep(1)

            # 点击 Continue 按钮（支持中英文，使用 submit 类型）
            try:
                page.locator('button[type="submit"]:has-text("Continue")').click(timeout=10000)
            except:
                page.locator('button[type="submit"]:has-text("继续")').click(timeout=10000)
            time.sleep(3)

        except Exception as e:
            return False, email, chatgpt_password, f"输入邮箱失败: {str(e)}"

        # 检查页面状态：可能是验证码输入框或密码输入框
        time.sleep(5)

        # 检测当前是哪种流程
        # 流程1: 邮箱 → 验证码 → 姓名+生日 → 密码
        # 流程2: 邮箱 → 密码 → 验证码 → 姓名+年龄

        has_code_input = False
        has_password_input = False

        try:
            code_input = page.locator('input[name="code"]')
            if code_input.count() > 0 and code_input.first.is_visible(timeout=3000):
                has_code_input = True
                print(f"[{email}] 检测到流程1: 先验证邮箱")
        except:
            pass

        try:
            password_input = page.locator('input[type="password"]')
            if password_input.count() > 0 and password_input.first.is_visible(timeout=3000):
                has_password_input = True
                print(f"[{email}] 检测到流程2: 先设置密码")
        except:
            pass

        # ========== 流程2: 先设置密码 ==========
        if has_password_input and not has_code_input:
            print(f"[{email}] Step 3 (流程2): 设置密码...")
            try:
                password_input = page.locator('input[type="password"]').first
                password_input.fill(chatgpt_password, timeout=10000)
                time.sleep(2)

                # 点击继续
                try:
                    page.locator('button[type="submit"]').first.click(timeout=5000)
                except:
                    pass
                time.sleep(5)

            except Exception as e:
                return False, email, chatgpt_password, f"设置密码失败: {str(e)}"

            # Step 4 (流程2): 获取验证码
            print(f"[{email}] Step 4 (流程2): 获取验证码...")

            if MANUAL_OTP:
                print(f"[{email}] 请在浏览器中手动输入验证码...")
                time.sleep(120)
            else:
                otp_code = mail_reader.get_verification_code(
                    sender_pattern="noreply@tm.openai.com",
                    timeout=90,
                    check_interval=5
                )

                if not otp_code:
                    return False, email, chatgpt_password, "未收到验证码"

                # 输入验证码
                print(f"[{email}] 输入验证码: {otp_code}...")
                try:
                    code_input = page.locator('input[name="code"]').first
                    code_input.fill(otp_code, timeout=10000)
                    time.sleep(2)

                    # 点击继续
                    try:
                        page.locator('button[type="submit"]').first.click(timeout=5000)
                    except:
                        pass
                    time.sleep(5)

                except Exception as e:
                    return False, email, chatgpt_password, f"输入验证码失败: {str(e)}"

            # Step 5 (流程2): 填写姓名和年龄（单个数字）
            print(f"[{email}] Step 5 (流程2): 填写姓名和年龄...")
            time.sleep(3)

            try:
                # 填写姓名
                name_input = page.locator('input[name="name"]')
                if name_input.count() > 0 and name_input.first.is_visible(timeout=3000):
                    full_name = generate_random_name()
                    print(f"[{email}] 填写姓名: {full_name}")
                    name_input.first.fill(full_name, timeout=10000)
                    time.sleep(1)

                # 填写年龄（单个数字，18-65岁）
                age = random.randint(18, 65)
                print(f"[{email}] 填写年龄: {age}")

                # 查找年龄输入框 - 可能是 input[name="age"] 或其他
                age_input = None
                try:
                    # 尝试多种选择器
                    age_selectors = [
                        'input[name="age"]',
                        'input[placeholder*="年龄"]',
                        'input[placeholder*="Age"]',
                        'input[type="number"]'
                    ]

                    for selector in age_selectors:
                        try:
                            age_input = page.locator(selector).first
                            if age_input.count() > 0 and age_input.is_visible(timeout=2000):
                                age_input.fill(str(age), timeout=5000)
                                print(f"[{email}] 年龄填写完成")
                                break
                        except:
                            continue

                    if not age_input:
                        # 如果找不到，尝试找第二个input（第一个是姓名）
                        all_inputs = page.locator('input[type="text"], input:not([type])')
                        if all_inputs.count() >= 2:
                            age_input = all_inputs.nth(1)
                            age_input.fill(str(age), timeout=5000)
                            print(f"[{email}] 年龄填写完成（使用第二个输入框）")

                except Exception as e:
                    print(f"[{email}] 填写年龄失败: {e}")

                time.sleep(2)

                # 点击完成按钮
                try:
                    submit_btn = page.locator('button[type="submit"]')
                    if submit_btn.count() > 0:
                        submit_btn.first.click(timeout=5000)
                        time.sleep(8)
                        print(f"[{email}] 已点击完成按钮")
                except Exception as e:
                    print(f"[{email}] 点击完成按钮失败: {e}")

            except Exception as e:
                print(f"[{email}] 填写姓名年龄失败: {e}")

            # 检查是否成功
            time.sleep(5)
            current_url = page.url
            print(f"[{email}] 当前 URL: {current_url}")

            if "chatgpt.com" in current_url and "/auth/" not in current_url:
                print(f"[{email}] ✅ 注册成功! (流程2)")
                return True, email, chatgpt_password, None

        # ========== 流程1: 先验证邮箱 ==========
        elif has_code_input:
            # Step 3a: 获取邮箱验证码
            print(f"[{email}] Step 3a: 获取邮箱验证码...")

            if MANUAL_OTP:
                print(f"[{email}] 请在浏览器中手动输入验证码...")
                time.sleep(120)
                otp_code = "manual"
            else:
                otp_code = mail_reader.get_verification_code(
                    sender_pattern="noreply@tm.openai.com",
                    timeout=90,
                    check_interval=5
                )

                if not otp_code:
                    return False, email, chatgpt_password, "未收到邮箱验证码"

                # 输入验证码
                print(f"[{email}] Step 3b: 输入邮箱验证码: {otp_code}...")
                try:
                    code_input = page.locator('input[name="code"]').first
                    code_input.fill(otp_code, timeout=10000)
                    time.sleep(2)

                    # 点击继续按钮
                    try:
                        page.locator('button[type="submit"]').first.click(timeout=5000)
                    except:
                        pass

                    time.sleep(5)

                except Exception as e:
                    return False, email, chatgpt_password, f"输入邮箱验证码失败: {str(e)}"

            # Step 3c: 填写姓名和生日
            print(f"[{email}] Step 3c: 检查是否需要填写姓名和生日...")
            time.sleep(3)

            # 检查是否有姓名输入框
            try:
                name_input = page.locator('input[name="name"]')
                if name_input.count() > 0 and name_input.first.is_visible(timeout=3000):
                    full_name = generate_random_name()
                    print(f"[{email}] 填写姓名: {full_name}")
                    name_input.first.click(timeout=3000)
                    time.sleep(0.5)
                    name_input.first.fill(full_name, timeout=10000)
                    time.sleep(1)

                    # 填写完姓名后，按Tab或点击其他地方，让生日输入框显示
                    name_input.first.press('Tab', timeout=2000)
                    time.sleep(1)
            except Exception as e:
                print(f"[{email}] 未找到姓名输入框或已跳过: {e}")

            # 检查是否有生日输入框（连续输入组件：年4位→月2位→日2位）
            try:
                birth_year, birth_month, birth_day = generate_random_birthday()
                print(f"[{email}] 填写生日: {birth_year}-{birth_month:02d}-{birth_day:02d}")

                time.sleep(2)

                # 生日是hidden input，需要点击页面上的显示区域来激活
                # 查找包含"年"、"月"、"日"的可点击元素
                birthday_input = None
                try:
                    # 尝试多种方式找到生日输入区域
                    # 方式1: 查找包含"年"和"月"和"日"的div或button
                    clickable_areas = page.locator('div, button, span').filter(has_text='年')
                    print(f"[{email}] 找到 {clickable_areas.count()} 个包含'年'的元素")

                    for i in range(clickable_areas.count()):
                        try:
                            elem = clickable_areas.nth(i)
                            text = elem.text_content(timeout=500)
                            if '年' in text and '月' in text and '日' in text:
                                print(f"[{email}] 找到生日显示区域: {text}")
                                # 点击这个区域
                                elem.click(timeout=3000)
                                time.sleep(1)

                                # 点击后，查找3个contenteditable的div（年、月、日）
                                year_div = page.locator('div[contenteditable="true"][data-type="year"]').first
                                month_div = page.locator('div[contenteditable="true"][data-type="month"]').first
                                day_div = page.locator('div[contenteditable="true"][data-type="day"]').first

                                if year_div.count() > 0 and month_div.count() > 0 and day_div.count() > 0:
                                    birthday_input = {'year': year_div, 'month': month_div, 'day': day_div}
                                    print(f"[{email}] 找到年月日3个contenteditable div")
                                    break
                        except:
                            continue

                except Exception as e:
                    print(f"[{email}] 查找生日区域失败: {e}")

                # 如果还没找到，尝试直接操作hidden input（通过JavaScript）
                if not birthday_input:
                    try:
                        print(f"[{email}] 尝试通过JavaScript设置生日")
                        birthday_string = f"{birth_year}-{birth_month:02d}-{birth_day:02d}"
                        page.evaluate(f'''
                            const birthdayInput = document.querySelector('input[name="birthday"]');
                            if (birthdayInput) {{
                                birthdayInput.value = "{birthday_string}";
                                birthdayInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                birthdayInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            }}
                        ''')
                        time.sleep(1)
                        print(f"[{email}] 已通过JavaScript设置生日")
                    except Exception as e:
                        print(f"[{email}] JavaScript设置失败: {e}")
                else:
                    # 使用找到的3个contenteditable div
                    try:
                        print(f"[{email}] 开始填写年月日")

                        # 确保焦点在年份输入框上
                        birthday_input['year'].focus(timeout=2000)
                        time.sleep(0.5)

                        # 年份输入框已经聚焦，直接输入
                        year_str = str(birth_year)
                        for char in year_str:
                            page.keyboard.press(char)
                            time.sleep(0.15)

                        print(f"[{email}] 年份填写完成: {year_str}")
                        time.sleep(0.3)

                        # 输入完年份会自动跳到月份，直接输入
                        month_str = f"{birth_month:02d}"
                        for char in month_str:
                            page.keyboard.press(char)
                            time.sleep(0.15)

                        print(f"[{email}] 月份填写完成: {month_str}")
                        time.sleep(0.3)

                        # 输入完月份会自动跳到日期，直接输入
                        day_str = f"{birth_day:02d}"
                        for char in day_str:
                            page.keyboard.press(char)
                            time.sleep(0.15)

                        print(f"[{email}] 日期填写完成: {day_str}")
                        time.sleep(1)

                        print(f"[{email}] 生日填写完成: {birth_year}-{birth_month:02d}-{birth_day:02d}")

                    except Exception as e:
                        print(f"[{email}] 填写生日失败: {e}")

                # 点击继续按钮前，检查是否有错误提示
                time.sleep(3)
                try:
                    # 检查是否有错误提示
                    error_elements = page.locator('text=/请输入|Invalid|错误|error/i')
                    if error_elements.count() > 0:
                        for i in range(min(3, error_elements.count())):
                            try:
                                if error_elements.nth(i).is_visible(timeout=500):
                                    error_text = error_elements.nth(i).text_content()
                                    print(f"[{email}] 检测到错误提示: {error_text}")
                            except:
                                pass
                except:
                    pass

                # 检查hidden input的值
                try:
                    birthday_value = page.evaluate('document.querySelector(\'input[name="birthday"]\')?.value')
                    print(f"[{email}] 生日hidden input值: {birthday_value}")
                except:
                    pass

                # 点击继续按钮
                try:
                    submit_btn = page.locator('button[type="submit"]')
                    if submit_btn.count() > 0:
                        # 检查按钮是否可用
                        is_enabled = submit_btn.first.is_enabled(timeout=2000)
                        is_disabled = submit_btn.first.get_attribute('disabled', timeout=1000)
                        print(f"[{email}] 继续按钮状态: enabled={is_enabled}, disabled={is_disabled}")

                        if is_enabled:
                            submit_btn.first.click(timeout=5000)
                            time.sleep(8)  # 增加等待时间
                            print(f"[{email}] 已点击继续按钮")
                        else:
                            print(f"[{email}] 继续按钮不可用，可能生日格式有误")
                            # 尝试截图调试
                            try:
                                page.screenshot(path=f"/tmp/{email}_birthday_error.png")
                                print(f"[{email}] 已保存截图到 /tmp/{email}_birthday_error.png")
                            except:
                                pass
                except Exception as e:
                    print(f"[{email}] 点击继续按钮失败: {e}")

            except Exception as e:
                print(f"[{email}] 填写生日流程失败: {e}")

        # Step 4: 设置密码（如果需要）
        print(f"[{email}] Step 4: 检查当前状态...")
        time.sleep(5)

        current_url = page.url
        print(f"[{email}] 当前 URL: {current_url}")

        # 如果已经跳转到chatgpt.com首页，说明注册成功（可能不需要设置密码）
        if "chatgpt.com" in current_url and "/auth/" not in current_url:
            print(f"[{email}] ✅ 注册成功! (已跳转到首页)")
            return True, email, chatgpt_password, None

        # 否则，尝试设置密码
        print(f"[{email}] Step 4: 设置密码...")
        try:
            # 等待密码输入框出现
            time.sleep(3)
            password_input = page.locator('input[type="password"]')

            # 等待密码框可见
            try:
                password_input.first.wait_for(state="visible", timeout=15000)
            except:
                # 再次检查URL，可能已经跳转了
                current_url = page.url
                if "chatgpt.com" in current_url and "/auth/" not in current_url:
                    print(f"[{email}] ✅ 注册成功! (已跳转到首页)")
                    return True, email, chatgpt_password, None
                print(f"[{email}] 当前 URL: {page.url}")
                return False, email, chatgpt_password, "未找到密码输入框或超时"

            if password_input.count() == 0:
                return False, email, chatgpt_password, "未找到密码输入框"

            password_input.first.fill(chatgpt_password, timeout=10000)
            time.sleep(2)

            # 点击 Continue 按钮
            try:
                page.locator('button[type="submit"]:has-text("Continue")').click(timeout=10000)
            except:
                try:
                    page.locator('button[type="submit"]:has-text("继续")').click(timeout=10000)
                except:
                    page.locator('button[type="submit"]').first.click(timeout=10000)
            time.sleep(5)

        except Exception as e:
            return False, email, chatgpt_password, f"设置密码失败: {str(e)}"

        # Step 5: 如果之前没验证邮箱，现在可能需要验证（检查是否有验证码输入框）
        if not has_code_input:
            print(f"[{email}] Step 5: 检查是否需要验证码...")
            time.sleep(2)

            # 检查是否出现验证码输入框
            try:
                code_input = page.locator('input[name="code"]')
                if code_input.count() > 0 and code_input.first.is_visible(timeout=2000):
                    print(f"[{email}] 检测到需要验证码")

                    if MANUAL_OTP:
                        print(f"[{email}] 请在浏览器中手动输入验证码...")
                        time.sleep(120)
                    else:
                        otp_code = mail_reader.get_verification_code(
                            sender_pattern="noreply@tm.openai.com",
                            timeout=90,
                            check_interval=5
                        )

                        if not otp_code:
                            return False, email, chatgpt_password, "未收到验证码"

                        print(f"[{email}] 输入验证码: {otp_code}...")
                        try:
                            code_input.first.fill(otp_code, timeout=10000)
                            time.sleep(2)

                            # 点击继续按钮
                            try:
                                page.locator('button[type="submit"]').first.click(timeout=5000)
                            except:
                                pass

                            time.sleep(3)

                        except Exception as e:
                            return False, email, chatgpt_password, f"输入验证码失败: {str(e)}"
                else:
                    print(f"[{email}] 无需验证码，继续...")
            except:
                print(f"[{email}] 无需验证码，继续...")
        else:
            print(f"[{email}] 已完成邮箱验证，继续...")

        # Step 6: 检查是否注册成功
        print(f"[{email}] Step 6: 检查注册结果...")
        time.sleep(10)

        # 检查是否进入了主页面或需要填写更多信息
        current_url = page.url
        print(f"[{email}] 当前 URL: {current_url}")

        if "chatgpt.com" in current_url and "/auth/" not in current_url:
            print(f"[{email}] ✅ 注册成功!")
            return True, email, chatgpt_password, None

        # 可能需要填写更多信息，尝试跳过
        for _ in range(3):
            try:
                # 尝试多种跳过按钮
                skip_selectors = [
                    'button:has-text("Skip")',
                    'button:has-text("跳过")',
                    'button:has-text("Next")',
                    'button:has-text("下一步")',
                    'button:has-text("Continue")',
                    'button:has-text("继续")'
                ]

                for selector in skip_selectors:
                    try:
                        skip_button = page.locator(selector).first
                        if skip_button.is_visible(timeout=2000):
                            skip_button.click(timeout=5000)
                            print(f"[{email}] 点击了跳过/继续按钮")
                            time.sleep(3)
                            break
                    except:
                        continue

                # 检查URL是否变化
                current_url = page.url
                if "chatgpt.com" in current_url and "/auth/" not in current_url:
                    print(f"[{email}] ✅ 注册成功!")
                    return True, email, chatgpt_password, None

            except:
                pass

            time.sleep(2)

        # 最终检查
        current_url = page.url
        if "chatgpt.com" in current_url and "/auth/" not in current_url:
            print(f"[{email}] ✅ 注册成功!")
            return True, email, chatgpt_password, None
        else:
            # 即使URL不对，也可能是成功了，检查页面内容
            try:
                # 如果能看到聊天界面元素，说明成功了
                if page.locator('textarea').count() > 0:
                    print(f"[{email}] ✅ 注册成功! (检测到聊天界面)")
                    return True, email, chatgpt_password, None
            except:
                pass

            return False, email, chatgpt_password, f"注册流程未完成，当前 URL: {current_url}"

    except Exception as e:
        return False, email, chatgpt_password, f"异常: {str(e)}"

    finally:
        # 清理资源
        try:
            if context:
                context.close()
        except:
            pass
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


def register_single_account(proxy=None, headless=False):
    """注册单个账号"""
    global ACCOUNT_POOL

    # 获取下一个账号
    mail_reader = ACCOUNT_POOL.get_next_account(proxy=proxy)
    if not mail_reader:
        return False, None, None, "账号池已用完"

    # 执行注册
    return register_chatgpt_with_browser(mail_reader, proxy=proxy, headless=headless)


def load_progress(progress_file):
    """加载进度，返回已完成的邮箱索引"""
    if os.path.exists(progress_file):
        with open(progress_file, 'r', encoding='utf-8') as f:
            return int(f.read().strip())
    return 0


def save_progress(progress_file, index):
    """保存进度"""
    with open(progress_file, 'w', encoding='utf-8') as f:
        f.write(str(index))


def worker_thread(thread_id, proxy, headless, output_file, results, total):
    """工作线程函数"""
    success, email, password, error = register_single_account(proxy, headless)

    results.append({
        'thread_id': thread_id,
        'success': success,
        'email': email,
        'password': password,
        'error': error
    })

    if success:
        # 追加写入结果文件（加锁）
        with ACCOUNT_POOL_LOCK:
            with open(output_file, 'a', encoding='utf-8') as f:
                f.write(f"{email}----{password}\n")
        print(f"[线程{thread_id}] ✅ 成功: {email} | 密码: {password}")
    else:
        print(f"[线程{thread_id}] ❌ 失败: {email or '未知'} | 原因: {error}")


def run_all_batches(batch_size=5, output_file="openai_accounts.txt", proxy=None, headless=False):
    """
    分批注册所有账号，每批完成后暂停等待用户切换IP
    每批内使用多线程并行处理

    Args:
        batch_size: 每批处理数量（并行线程数）
        output_file: 输出文件（email----password 格式）
        proxy: 代理地址
        headless: 是否无头模式
    """
    init_account_pool()

    total = len(ACCOUNT_POOL.accounts)
    progress_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".progress")
    start_index = load_progress(progress_file)

    if start_index > 0:
        ACCOUNT_POOL.current_index = start_index
        print(f"[恢复] 从第 {start_index + 1} 个账号继续 (已完成 {start_index}/{total})")

    total_success = 0
    total_fail = 0
    batch_num = start_index // batch_size + 1

    while ACCOUNT_POOL.current_index < total:
        current_start = ACCOUNT_POOL.current_index
        current_end = min(current_start + batch_size, total)
        count = current_end - current_start

        print()
        print("=" * 60)
        print(f"第 {batch_num} 组 | 账号 {current_start + 1}-{current_end} / {total}")
        print(f"代理: {proxy or '无'} | 无头模式: {headless}")
        print(f"🚀 同时启动 {count} 个浏览器并行注册")
        print("=" * 60)

        # 创建线程列表
        threads = []
        results = []

        # 启动多个线程
        for i in range(count):
            thread_id = current_start + i + 1
            t = threading.Thread(
                target=worker_thread,
                args=(thread_id, proxy, headless, output_file, results, total)
            )
            threads.append(t)
            t.start()
            print(f"[启动] 线程 {thread_id}")
            time.sleep(0.5)  # 稍微错开启动时间

        # 等待所有线程完成
        print(f"\n⏳ 等待 {count} 个线程完成...")
        for t in threads:
            t.join()

        # 统计结果
        batch_success = sum(1 for r in results if r['success'])
        batch_fail = len(results) - batch_success
        total_success += batch_success
        total_fail += batch_fail

        # 保存进度
        save_progress(progress_file, ACCOUNT_POOL.current_index)

        print()
        print("=" * 60)
        print(f"第 {batch_num} 组完成!")
        print(f"本组: 成功 {batch_success} | 失败 {batch_fail}")
        print(f"累计: 成功 {total_success} | 失败 {total_fail} | 已处理 {ACCOUNT_POOL.current_index}/{total}")
        print(f"结果文件: {output_file}")
        print("=" * 60)

        # 如果还有剩余账号，暂停等待切换IP
        if ACCOUNT_POOL.current_index < total:
            print()
            if AUTO_CONTINUE:
                print("⚠️  自动继续下一组...")
                print(f"   剩余: {total - ACCOUNT_POOL.current_index} 个账号")
                time.sleep(3)
            else:
                print("⚠️  请切换IP后按回车继续下一组...")
                print(f"   剩余: {total - ACCOUNT_POOL.current_index} 个账号")
                input()
            batch_num += 1

    # 全部完成
    print()
    print("=" * 60)
    print("全部注册完成!")
    print(f"总计: 成功 {total_success} | 失败 {total_fail}")
    print(f"账号密码文件: {output_file}")
    print("=" * 60)

    # 清除进度文件
    if os.path.exists(progress_file):
        os.remove(progress_file)


def main():
    """主函数"""
    print("=" * 60)
    print("ChatGPT 批量注册工具 - 浏览器版")
    print(f"每组 {BATCH_SIZE} 个，完成后暂停切换IP")
    print("=" * 60)

    proxy = DEFAULT_PROXY
    headless = HEADLESS

    if proxy:
        print(f"[Info] 代理: {proxy}")
    print(f"[Info] 无头模式: {headless}")
    print(f"[Info] 账号文件: {OUTLOOK_ACCOUNTS_FILE}")
    print(f"[Info] 输出文件: {DEFAULT_OUTPUT_FILE}")

    run_all_batches(
        batch_size=BATCH_SIZE,
        output_file=DEFAULT_OUTPUT_FILE,
        proxy=proxy,
        headless=headless,
    )


if __name__ == "__main__":
    main()
