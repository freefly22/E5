# -*- coding: UTF-8 -*-
import requests as req
import json, sys, time, os
import telepot

# ===== 从环境变量读取（CI 标准写法）=====
BOT_TOKEN = os.getenv("TELEBOT_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
CLIENT_ID = os.getenv("CONFIG_ID")
CLIENT_SECRET = os.getenv("CONFIG_KEY")

if not all([BOT_TOKEN, CHAT_ID, CLIENT_ID, CLIENT_SECRET]):
    raise RuntimeError("Missing required environment variables")

bot = telepot.Bot(BOT_TOKEN)

BASE_DIR = sys.path[0]
TOKEN_FILE = os.path.join(BASE_DIR, "1.txt")

num1 = 0
fin = None

def send(message):
    bot.sendMessage(CHAT_ID, message)

def gettoken(old_refresh_token):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': old_refresh_token,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'redirect_uri': 'http://localhost:53682/'
    }

    r = req.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        data=data,
        headers=headers
    )

    jsontxt = r.json()

    if 'access_token' not in jsontxt:
        raise RuntimeError(f"Token refresh failed: {jsontxt}")

    # 有就更新，没有就沿用旧的
    new_refresh_token = jsontxt.get('refresh_token', old_refresh_token)

    with open(TOKEN_FILE, 'w') as f:
        f.write(new_refresh_token)

    return jsontxt['access_token']

def main():
    global num1, fin

    fin = time.asctime(time.localtime())

    # 第一次 CI 运行时兜底
    if not os.path.exists(TOKEN_FILE):
        raise RuntimeError("refresh_token file 1.txt not found")

    with open(TOKEN_FILE, 'r') as f:
        refresh_token = f.read().strip()

    access_token = gettoken(refresh_token)

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    urls = [
        'https://graph.microsoft.com/v1.0/me/drive/root',
        'https://graph.microsoft.com/v1.0/me/drive',
        'https://graph.microsoft.com/v1.0/drive/root',
        'https://graph.microsoft.com/v1.0/users',
        'https://graph.microsoft.com/v1.0/me/messages',
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
        'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
        'https://graph.microsoft.com/v1.0/me/drive/root/children',
        'https://api.powerbi.com/v1.0/myorg/apps',
        'https://graph.microsoft.com/v1.0/me/mailFolders',
        'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'
    ]

    for i, url in enumerate(urls, 1):
        r = req.get(url, headers=headers)
        if r.status_code == 200:
            num1 += 1
            print(f"{i} 调用成功 {num1} 次")

for _ in range(8):
    main()

msg = f"[AutoApiSecret] 已成功调用 {num1} 次，结束时间为 {fin}"
send(msg)
