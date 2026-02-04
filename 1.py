# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
import telepot
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
fin=None
token = str(sys.argv[1]) if len(sys.argv) > 1 else ""
chat_id = str(sys.argv[2]) if len(sys.argv) > 2 else ""
bot = telepot.Bot(token) if token and chat_id else None

def send(message):
    if not bot:
        print(message)
        return
    try:
        bot.sendMessage(chat_id,message, parse_mode=None, disable_web_page_preview=None, disable_notification=None, reply_to_message_id=None, reply_markup=None)
    except Exception as exc:
        print(f"Telegram send failed: {exc}")

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    if 'access_token' not in jsontxt:
        raise RuntimeError(f"token refresh failed: {jsontxt}")
    refresh_token = jsontxt.get('refresh_token', refresh_token)
    access_token = jsontxt['access_token']
    return access_token, refresh_token

def run_calls(access_token):
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    endpoints = [
        ('1', r'https://graph.microsoft.com/v1.0/me/drive/root'),
        ('2', r'https://graph.microsoft.com/v1.0/me/drive'),
        ('3', r'https://graph.microsoft.com/v1.0/drive/root'),
        ('4', r'https://graph.microsoft.com/v1.0/users '),
        ('5', r'https://graph.microsoft.com/v1.0/me/messages'),
        ('6', r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules'),
        ('7', r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta'),
        ('8', r'https://graph.microsoft.com/v1.0/me/drive/root/children'),
        ('9', r'https://api.powerbi.com/v1.0/myorg/apps'),
        ('10', r'https://graph.microsoft.com/v1.0/me/mailFolders'),
        ('11', r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'),
    ]
    success = 0
    failed = []
    for name, url in endpoints:
        status = req.get(url, headers=headers).status_code
        if status == 200:
            success += 1
            print(f"{name}调用成功{success}次")
        else:
            failed.append((name, status))
    return success, failed

def run_account(refresh_token, rounds=8):
    total_success = 0
    total_failed = []
    current_token = refresh_token
    for _ in range(rounds):
        access_token, current_token = gettoken(current_token)
        success, failed = run_calls(access_token)
        total_success += success
        total_failed.extend(failed)
    return total_success, total_failed, current_token

def load_refresh_tokens():
    with open(path, "r+") as fo:
        tokens = [line.strip() for line in fo.read().splitlines() if line.strip()]
    if not tokens:
        raise RuntimeError("no refresh_token found in 1.txt")
    return tokens

def save_refresh_tokens(tokens):
    with open(path, 'w+') as f:
        f.write("\n".join(tokens))

def main():
    global fin
    fin = localtime = time.asctime(time.localtime(time.time()))
    tokens = load_refresh_tokens()
    updated_tokens = []
    summary = []
    for index, refresh_token in enumerate(tokens, start=1):
        try:
            success, failed, new_token = run_account(refresh_token)
            updated_tokens.append(new_token)
            summary_line = f"账号{index}: 成功{success}次, 失败{len(failed)}次"
            summary.append(summary_line)
            send(f"[AutoApiSecret] {summary_line}")
        except Exception as exc:
            updated_tokens.append(refresh_token)
            summary.append(f"账号{index}: 运行失败 ({exc})")
    save_refresh_tokens(updated_tokens)
    msg='[AutoApiSecret]运行结束时间为{}\\n{}'.format(localtime, "\\n".join(summary))
    send(msg)

main()
