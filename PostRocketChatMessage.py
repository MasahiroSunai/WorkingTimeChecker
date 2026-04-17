#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import os
from pathlib import Path
from rocketchat_API.rocketchat import RocketChat
from config_loader import load_config

def post_message_to_rocketchat(message, url: str, userid: str, authtoken: str, cert_dir: str):
    try:
        if not os.path.exists(cert_dir):
            raise FileNotFoundError(f"証明書ディレクトリが見つかりません: {cert_dir}")
        if not os.path.exists(os.path.join(cert_dir, 'client.cert')) or not os.path.exists(os.path.join(cert_dir, 'client.key')):
            raise FileNotFoundError("証明書ファイルが見つかりません。client.certとclient.keyが必要です。")
        cert_files = (
            os.path.join(cert_dir, 'client.cert'),
            os.path.join(cert_dir, 'client.key')
        )
        # RocketChatクライアントの初期化
        rocket = RocketChat(
            user_id=userid,
            auth_token=authtoken,
            server_url=url,
            client_certs=cert_files,
            ssl_verify=False,
            proxies=None,
            timeout=1000
        )

        # メッセージの投稿
        response = rocket.chat_post_message(message, channel=roomname)

        # レスポンスの確認
        if response.status_code == 200:
            print("Message posted successfully.")
        else:
            print(f"Failed to post message. Status code: {response.status_code}")
            print(f"Response: {response.json()}")

    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    parser = argparse.ArgumentParser(description="Rocket.Chatにメッセージを投稿するスクリプト")
    parser.add_argument('--message', '-m', type=str, help='投稿するメッセージを指定', required=True)
    parser.add_argument("--config", default="config/config.yaml", help="config.yaml のパス")
    args = parser.parse_args()
    config = load_config(args.config)

    cert_dir = config["confluence"]["dndev"].get("cert_dir")
    rocket_conf = config["rocketchat"]
    url = rocket_conf["url"]
    roomname = rocket_conf["room"]
    userid = os.environ["ROCKETCHAT_USER_ID"]
    authtoken = os.environ["ROCKETCHAT_TOKEN"]

    if not userid or not authtoken:
        raise RuntimeError("Rocket.Chat 認証情報が環境変数にありません")

    post_message_to_rocketchat(args.message, url, userid, authtoken, cert_dir)

if __name__ == "__main__":
    main()