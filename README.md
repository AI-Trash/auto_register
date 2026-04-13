# ChatGPT 批量注册工具 - 浏览器自动化版

基于 Playwright 的 ChatGPT 账号批量注册工具，使用 Outlook 邮箱进行注册。

## 特性

- 🚀 使用浏览器自动化，模拟真实用户操作
- 📧 自动读取 Outlook 邮箱验证码
- 🔄 支持批量并行注册（每组 3 个账号）
- 💾 自动保存注册成功的账号信息
- 🔁 支持断点续传，失败后可继续
- 🌐 支持代理设置

## 环境要求

- Python 3.8+
- 可用的代理（推荐使用 TUN 模式）

## 安装

1. 创建虚拟环境并激活：
```bash
python3 -m venv venv
source venv/bin/activate  # macOS/Linux
# 或
venv\Scripts\activate  # Windows
```

2. 安装依赖：
```bash
pip install patchright
python -m patchright install chromium
```

## 配置

编辑 `config_browser.json`：

```json
{
    "outlook_accounts_file": "邮箱文件.txt",
    "output_file": "Results/registered_accounts.txt",
    "proxy": "http://127.0.0.1:7897",
    "headless": false,
    "batch_size": 3
}
```

### 配置说明

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| outlook_accounts_file | Outlook 邮箱账号文件路径 | 必填 |
| output_file | 注册成功的账号输出文件 | registered_accounts.txt |
| proxy | 代理地址（支持 http/https/socks5） | null |
| headless | 是否无头模式运行 | false |
| batch_size | 每组并行注册的账号数量 | 3 |

## 邮箱文件格式

创建 `邮箱文件.txt`，每行一个账号，格式：
```
邮箱----密码
```

示例（参考 `邮箱文件.txt`）：
```
example1@outlook.com----Password123!
example2@outlook.com----SecurePass456
example3@outlook.com----MyPassword789
```

注意：
- 使用 `----` 作为分隔符（4个短横线）
- 支持 `#` 开头的注释行
- 密码建议包含大小写字母、数字和特殊字符

## 使用方法

1. 准备好 Outlook 邮箱账号文件
2. 配置 `config_browser.json`
3. 运行脚本：

```bash
source venv/bin/activate
python chatgpt_register_browser.py
```

## 输出格式

注册成功的账号保存在 `Results/registered_accounts.txt`，格式：
```
邮箱----密码
```

示例：
```
example1@outlook.com----Password123!88888888
example2@outlook.com----SecurePass45688888888
```

注意：ChatGPT 密码会在原 Outlook 密码基础上补充 `8` 至 12 位。

## 工作流程

1. 从配置文件读取 Outlook 邮箱列表
2. 每组启动 3 个浏览器并行注册
3. 访问 ChatGPT 注册页面
4. 输入邮箱地址
5. 自动读取 Outlook 邮箱中的验证码
6. 填写验证码和生日信息
7. 设置密码（原密码 + "888"）
8. 保存注册成功的账号
9. 完成一组后暂停，等待手动切换 IP
10. 继续下一组

## 注意事项

- 建议使用 TUN 模式代理，确保浏览器流量走代理
- 每组注册完成后会暂停，需要手动按回车继续
- 建议每组 3 个账号，避免触发风控
- 首次运行建议设置 `headless: false` 观察运行情况
- 注册失败的账号可以重新运行脚本继续注册

## 故障排除

### 浏览器无法启动
```bash
python -m patchright install chromium
```

### 代理连接失败
- 检查代理地址是否正确
- 确认代理服务正在运行
- 尝试使用 TUN 模式代理

### 验证码读取失败
- 检查 Outlook 邮箱密码是否正确
- 确认邮箱可以正常登录
- 等待几秒后重试

## 文件说明

- `chatgpt_register_browser.py` - 主程序
- `outlook_browser_reader.py` - Outlook 邮箱验证码读取模块
- `config_browser.json` - 配置文件
- `邮箱文件.txt` - 邮箱账号列表示例
- `Results/` - 输出目录（注册成功的账号）
- `venv/` - Python 虚拟环境

## 致谢

本项目基于 https://github.com/adminlove520/chatgpt_register 改造而来。
