#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import os
from pathlib import Path
from rocketchat_API.rocketchat import RocketChat
import urllib3
from utils.config_loader import load_config
from utils.logger_utils import setup_logger
from utils.month_utils import resolve_target_month

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = setup_logger("PostRocketChatMessage")

def build_message_from_config(config: dict) -> str:
    dndev = config["confluence"]["dndev"]
    geniie = config["confluence"]["geniie"]

    
    vars = resolve_target_month(config)

    dndev_url = (
        f'{dndev["url"]}/pages/viewpage.action?pageId={dndev["upload_page_id"]}'
    )
    geniie_url = (
        f'{geniie["url"]}/pages/viewpage.action?pageId={geniie["upload_page_id"]}'
    )

    message = (
        "@all\n"
        f"{vars['YYYY']}/{vars['MM']} 工数表⇔Web勤怠相違チェック処理が完了しました。\n"
        "Confluenceの該当ページをご確認ください。\n\n"
        f"dndevWikiURL: {dndev_url}\n"
        f"geniieWikiURL: {geniie_url}"
    )

    return message

def post_message_to_rocketchat(message, url: str, userid: str, authtoken: str, cert_dir: str, roomname: str):
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
        if isinstance(response, dict) and response.get("success"):
            logger.info("Message posted successfully.")
        else:
            logger.error(f"Failed to post message. Response: {response}")
            raise RuntimeError(f"Rocket.Chat post failed: {response}")

    except Exception as e:
        logger.exception("Rocket.Chat 投稿エラー")
        raise

def main():
    logger.info("PostRocketChatMessage start")
    try:
        parser = argparse.ArgumentParser(description="Rocket.Chatにメッセージを投稿するスクリプト")
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

        dry_run = config.get("system", {}).get("dry_run", False)

        message = build_message_from_config(config)

        if dry_run:
            logger.info("Dry run mode. Message would be posted to Rocket.Chat:")
            logger.info(message)
        else:
            post_message_to_rocketchat(message, url, userid, authtoken, cert_dir, roomname)
            logger.info("Message posted to Rocket.Chat successfully")

    except Exception as e:
        logger.exception(f"エラーが発生しました: {e}")
        raise

if __name__ == "__main__":
    main()