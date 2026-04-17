# -*- coding: utf-8 -*-

import argparse
from config_loader import load_config
from confluence_utils import create_confluence

def main():
    parser = argparse.ArgumentParser(description="Confluenceページコピー")
    parser.add_argument("--config", default="config/config.yaml")
    args = parser.parse_args()

    config = load_config(args.config)

    dndev_conf = config["confluence"]["dndev"]
    geniie_conf = config["confluence"]["geniie"]

    src = create_confluence(
        dndev_conf,
        username_env="DNDEV_AWS_ID",
        password_env="DNDEV_AWS_PASSWORD",
    )
    dst = create_confluence(
        geniie_conf,
        username_env="GENIIE_ID",
        password_env="GENIIE_PASSWORD",
    )

    src_page_id = dndev_conf["upload_page_id"]
    dst_page_id = geniie_conf["upload_page_id"]

    src_page = src.get_page_by_id(src_page_id, expand="body.storage")
    if not src_page:
        raise RuntimeError(f"コピー元ページ {src_page_id} が見つかりません")

    body = src_page["body"]["storage"]["value"]
    title = src_page["title"]

    dst_page = dst.get_page_by_id(dst_page_id, expand="body.storage")
    if not dst_page:
        raise RuntimeError(f"コピー先ページ {dst_page_id} が見つかりません")

    if dst_page["body"]["storage"]["value"] == body:
        print("ページ内容は同一のためコピー不要")
        return

    dst.update_page(
        page_id=dst_page_id,
        title=src_page["title"],
        body=body,
        representation="storage",
        minor_edit=True,
        version_comment=f"Copied from page {src_page_id}",
    )

    print(f"{src_page_id} → {dst_page_id} コピー完了")

if __name__ == "__main__":
    main()