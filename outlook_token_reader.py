"""
通过 Microsoft Graph API 使用 token 读取 Outlook 邮件
"""

import time
import re
import requests


DEFAULT_CLIENT_ID = 'dbc8e03a-b00c-46bd-ae65-b683e7707cb0'
UUID_RE = re.compile(r'^[0-9a-fA-F-]{36}$')


def _looks_like_refresh_token(value):
    """判断字符串是否像 refresh_token。"""
    if not value:
        return False
    value = value.strip()
    return (
        value.startswith("M.")
        or value.startswith("0.A")
        or "!" in value
        or len(value) > 80
    )


def refresh_access_token(refresh_token, client_id=DEFAULT_CLIENT_ID):
    """使用 refresh_token 刷新 access_token - 用于 Graph API"""
    scopes = [
        'mail.read offline_access',
        'https://graph.microsoft.com/Mail.Read offline_access',
    ]

    for scope in scopes:
        try:
            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data={
                    'client_id': client_id,
                    'refresh_token': refresh_token,
                    'grant_type': 'refresh_token',
                    'scope': scope
                },
                headers={'Content-Type': 'application/x-www-form-urlencoded'},
                timeout=30
            )

            if response.status_code == 200:
                tokens = response.json()
                return tokens.get('access_token'), tokens.get('refresh_token', refresh_token)

            print(f"  [Token] 刷新失败({scope}): {response.status_code}")
            print(f"  [Token] 错误详情: {response.text}")
        except Exception as e:
            print(f"  [Token] 刷新异常({scope}): {e}")

    return None, None


def _is_openai_mail(sender, subject):
    """判断是否为验证码相关邮件。"""
    sender_l = (sender or "").lower()
    subject_l = (subject or "").lower()
    return (
        'openai' in sender_l
        or 'chatgpt' in sender_l
        or 'openai' in subject_l
        or 'chatgpt' in subject_l
        or 'verification' in subject_l
        or 'code' in subject_l
    )


def _extract_code(text):
    """从文本中提取验证码。"""
    content = text or ""
    for pattern in (r'\b(\d{6})\b', r'\b(\d{8})\b', r'\b(\d{4,8})\b'):
        match = re.search(pattern, content)
        if match:
            return match.group(1)
    return None


class OutlookMailReader:
    """Outlook 邮件读取器 - 使用 token"""
    def __init__(self, email, password, access_token=None, refresh_token=None, client_id=None):
        self.email = email
        self.password = password
        self.access_token = access_token
        self.refresh_token = refresh_token
        self.client_id = client_id or DEFAULT_CLIENT_ID

    def get_verification_code(self, sender_pattern=None, timeout=90, check_interval=5):
        """获取验证码"""
        # 如果 access_token 过期，尝试刷新
        if not self.access_token and self.refresh_token:
            print(f"  [Outlook] 尝试刷新 token...")
            new_access, new_refresh = refresh_access_token(self.refresh_token, self.client_id)
            if new_access:
                self.access_token = new_access
                self.refresh_token = new_refresh
                print(f"  [Outlook] ✓ Token 刷新成功")
            else:
                print(f"  [Outlook] ✗ Token 刷新失败（仅支持 Graph，不使用第三方 API）")
                return None

        if self.access_token:
            print(f"  [Outlook] 使用 Graph token 读取邮件...")
        else:
            print(f"  [Outlook] 错误: 没有可用 token")
            return None
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                headers = {
                    'Authorization': f'Bearer {self.access_token}',
                    'Content-Type': 'application/json'
                }
                url = (
                    'https://graph.microsoft.com/v1.0/me/messages'
                    '?$top=20'
                    '&$orderby=receivedDateTime%20DESC'
                    '&$select=subject,from,body,bodyPreview,receivedDateTime'
                )
                response = requests.get(url, headers=headers, timeout=30)

                if response.status_code == 401:
                    if self.refresh_token:
                        print(f"  [Outlook] Graph token 过期，尝试刷新...")
                        new_access, new_refresh = refresh_access_token(self.refresh_token, self.client_id)
                        if new_access:
                            self.access_token = new_access
                            self.refresh_token = new_refresh
                            print(f"  [Outlook] ✓ Graph Token 刷新成功，继续...")
                            continue
                    print(f"  [Outlook] Graph token 无法刷新（仅支持 Graph）")
                    return None

                if response.status_code != 200:
                    print(f"  [Outlook] Graph API 错误: {response.status_code}")
                    time.sleep(check_interval)
                    continue

                data = response.json()
                messages = data.get('value', [])
                for msg in messages:
                    sender = msg.get('from', {}).get('emailAddress', {}).get('address', '')
                    subject = msg.get('subject', '')
                    body_preview = msg.get('bodyPreview', '')
                    body = msg.get('body', {}).get('content', '')
                    if _is_openai_mail(sender, subject):
                        print(f"  [Outlook] 找到 OpenAI 邮件(Graph): {subject}")
                        code = _extract_code(f"{subject}\n{body_preview}\n{body}")
                        if code:
                            print(f"  [Outlook] ✓ 找到验证码(Graph): {code}")
                            return code

                print(f"  [Outlook] 未找到验证码，{check_interval} 秒后重试...")
                time.sleep(check_interval)

            except Exception as e:
                print(f"  [Outlook] 异常: {e}")
                time.sleep(check_interval)

        print(f"  [Outlook] ✗ 超时未找到验证码")
        return None


class OutlookAccountPool:
    """Outlook 账号池"""
    def __init__(self, accounts_file, default_client_id=None):
        self.accounts_file = accounts_file
        self.default_client_id = default_client_id or DEFAULT_CLIENT_ID
        self.accounts = []
        self.current_index = 0
        self.load_accounts()

    def load_accounts(self):
        """从文件加载账号"""
        try:
            with open(self.accounts_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue

                    # 支持格式：
                    # 1) email----password----client_id----refresh_token
                    # 2) email----password----refresh_token----client_id
                    # 3) email----password----refresh_token
                    # 4) email----password----access_token----refresh_token
                    if '----' in line:
                        parts = line.split('----')
                        if len(parts) >= 4:
                            email = parts[0].strip()
                            password = parts[1].strip()
                            p3 = parts[2].strip()
                            p4 = parts[3].strip()

                            client_id = None
                            refresh_token = None
                            access_token = None

                            # 兼容多种位置：优先识别 refresh_token
                            if _looks_like_refresh_token(p3):
                                refresh_token = p3
                                if UUID_RE.match(p4):
                                    client_id = p4
                                else:
                                    access_token = p4 if p4 else None
                            elif _looks_like_refresh_token(p4):
                                refresh_token = p4
                                if UUID_RE.match(p3):
                                    client_id = p3
                                else:
                                    access_token = p3 if p3 else None
                            else:
                                # 回退旧约定：第三个是 client_id，第四个是 refresh_token
                                client_id = p3 if p3 else None
                                refresh_token = p4 if p4 else None

                            self.accounts.append({
                                'email': email,
                                'password': password,
                                'client_id': client_id or self.default_client_id,
                                'refresh_token': refresh_token,
                                'access_token': access_token
                            })
                        elif len(parts) == 3:
                            email = parts[0].strip()
                            password = parts[1].strip()
                            p3 = parts[2].strip()
                            self.accounts.append({
                                'email': email,
                                'password': password,
                                'client_id': self.default_client_id,
                                'refresh_token': p3 if _looks_like_refresh_token(p3) else None,
                                'access_token': None if _looks_like_refresh_token(p3) else p3
                            })
                        elif len(parts) >= 2:
                            email = parts[0].strip()
                            password = parts[1].strip()
                            self.accounts.append({
                                'email': email,
                                'password': password,
                                'client_id': self.default_client_id,
                                'refresh_token': None,
                                'access_token': None
                            })
                    elif ':' in line:
                        email, password = line.split(':', 1)
                        self.accounts.append({
                            'email': email.strip(),
                            'password': password.strip(),
                            'client_id': self.default_client_id,
                            'refresh_token': None,
                            'access_token': None
                        })
        except FileNotFoundError:
            print(f"账号文件不存在: {self.accounts_file}")

    def get_next_account(self, proxy=None):
        """获取下一个邮件读取器"""
        if self.current_index >= len(self.accounts):
            return None
        account = self.accounts[self.current_index]
        self.current_index += 1
        return OutlookMailReader(
            account['email'],
            account['password'],
            account.get('access_token'),
            account.get('refresh_token'),
            account.get('client_id')
        )

    def get_reader(self, proxy=None):
        """获取一个邮件读取器（别名）"""
        return self.get_next_account(proxy)
