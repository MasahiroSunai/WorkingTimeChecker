# -*- coding: utf-8 -*-

import os
from atlassian import Confluence

def create_confluence(conf: dict, username_env: str, password_env: str) -> Confluence:
    username = os.environ.get(username_env)
    password = os.environ.get(password_env)
    if not username or not password:
        raise RuntimeError(f"環境変数 {username_env}/{password_env} が設定されていません")

    cert_dir = conf.get("cert_dir")
    if cert_dir:
        cert = (
            os.path.join(cert_dir, "client.cert"),
            os.path.join(cert_dir, "client.key"),
        )
        return Confluence(
            url=conf["url"],
            username=username,
            password=password,
            verify_ssl=True,
            cert=cert,
        )
    else:
        return Confluence(
            url=conf["url"],
            username=username,
            password=password,
            verify_ssl=False,
        )