# -*- coding: utf-8 -*-

import argparse
import os
from config_loader import load_config
from confluence_utils import create_confluence
from month_utils import resolve_target_month
from path_utils import expand_path

def main():
    parser = argparse.ArgumentParser(description="Confluenceの添付ファイルをダウンロード")
    parser.add_argument("--config", default="config/config.yaml")
    args = parser.parse_args()

    config = load_config(args.config)
    vars = resolve_target_month(config)

    conf = config["confluence"]["dndev"]
    download_dir = expand_path(config["paths"]["workload_download_dir"], vars)
    os.makedirs(download_dir, exist_ok=True)

    confluence = create_confluence(
        conf,
        username_env="DNDEV_AWS_ID",
        password_env="DNDEV_AWS_PASSWORD",
    )

    page_id = conf["download_page_id"]
    if confluence.download_attachments_from_page(page_id=page_id, path=download_dir) is None:
        raise RuntimeError(f"ページ {page_id} のダウンロードに失敗しました")

    print(f"Confluence(dndev) ページ {page_id} の添付を {download_dir} に保存しました")

if __name__ == "__main__":
    main()