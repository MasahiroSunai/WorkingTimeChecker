#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import pandas as pd
import win32com.client
import traceback
import os
import sys
import argparse
from atlassian import Confluence
from typing import Dict, Set, Optional
from utils.config_loader import load_config
from utils.month_utils import resolve_target_month
from utils.path_utils import expand_path
from utils.logger_utils import setup_logger

logger = setup_logger("WorkingTimeChecker")

# 各種辞書型およびセットを初期化
dicWorkLoad2WebAttendance: Dict[str, str] = {}
dicWebAttendance2WorkLoad: Dict[str, str] = {}
dicWorkLoad: Dict[str, Dict[str, Dict[int, float]]] = {}
dicWebAttendance: Dict[str, Dict[str, Dict[int, float]]] = {}
dicCheckMember: Dict[str, bool] = {}
setLocalTaskJob: Set[str] = set()

# 出力用データフレームを初期化
dfOutWorkLoad = pd.DataFrame(columns=['社員番号', '社員名', 'プロジェクト-タスク名', '日付', '工数表工数', 'Web勤怠工数'])
dfOutWebAttendance = pd.DataFrame(columns=['社員番号', '社員名', 'ワーク名', '日付', 'Web勤怠工数', '工数表工数'])
dfOutLastDay = pd.DataFrame(columns=['社員番号', '社員名', '最終勤務日'])
dfLocalTaskAndNoNote = pd.DataFrame(columns=['社員番号', '社員名', 'ワーク名', '日付'])

def read_check_member(file_path: str):
    """比較対象メンバーをCSVファイルから読み込む"""
    try:
        with open(file_path, mode='r', encoding='cp932') as file:
            reader = csv.DictReader(file)
            for row in reader:
                isCheck = row['比較対象']
                if isCheck == 'TRUE':
                    employee_id_name = f"{row['社員番号']}-{row['名前']}"
                    isLocalTaskCheck = row['社内工数対象']
                    dicCheckMember[employee_id_name] = (isLocalTaskCheck == 'TRUE')
        logger.info(f"Successfully read check member from {file_path}")
    except FileNotFoundError:
        logger.error(f"Error: The file {file_path} was not found.")
    except Exception as e:
        logger.exception(f"An error occurred while reading the file: {e}")

def read_working_map(file_path: str):
    """ワークの対応付けをCSVファイルから読み込む"""
    try:
        with open(file_path, mode='r', encoding='cp932') as file:
            reader = csv.DictReader(file)
            for row in reader:
                prj_name = row['工数集計 プロジェクト名']
                task_name = row['工数集計 タスク名']
                work_name = row['Web勤怠 ワーク名'].replace('社内PJT-', '')

                if prj_name == '社内工数':
                    setLocalTaskJob.add(work_name)
                
                key = f"{prj_name}-{task_name}" if task_name else prj_name
                dicWorkLoad2WebAttendance[key] = work_name
                dicWebAttendance2WorkLoad[work_name] = key
        print(f"Successfully read working map from {file_path}")
    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")

def read_working_load(file_path: str, sheet_name: str = '工数一覧表'):
    """工数表をExcelファイルから読み込む"""
    try:
        ws = pd.read_excel(file_path, sheet_name=sheet_name, header=5)
        for row in ws.itertuples(index=False):
            if pd.isna(row[1]) or pd.isna(row[2]):
                continue
            employee_id_name = f"{row[1]}-{row[2]}"
            project_name = str(row[8]).strip()
            task_name = str(row[9]).strip() if pd.notna(row[9]) else None

            # プロジェクト名が「休憩」で始まる場合はスキップ
            if project_name.startswith('休憩'):
                continue

            key = f"{project_name}-{task_name}" if pd.notna(task_name) else project_name

            dicWorkLoad.setdefault(employee_id_name, {}).setdefault(key, {})

            for col in range(12, 43):
                day = col - 11
                if pd.notna(row[col]):
                    worktime = row[col]
                    dicWorkLoad[employee_id_name][key].setdefault(day, 0)
                    dicWorkLoad[employee_id_name][key][day] += worktime

        print(f"Successfully read working load from {file_path}")

    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")

def read_web_attendance(file_path: str):
    """Web勤怠をcsvファイルから読み込む"""
    try:
        dicTempLastDay = {}
        with open(file_path, mode='r', encoding='cp932') as file:
            reader = csv.DictReader(file)
            for row in reader:
                employee_id_name = f"{row['社員番号']}-{row['社員名称: 社員名']}"
                work_name = row['ワーク番号: ワーク名'].strip()
                day = int(row['勤務日'].split('/')[2])
                work_time = float(row['時間・分']) / 60

                dicTempLastDay[employee_id_name] = max(dicTempLastDay.get(employee_id_name, 0), day)
                
                if work_name.startswith('準）'):
                    continue

                # 社内工数で、備考の記載がない場合、dfLocalTaskAndNoNoteに追加
                if work_name in setLocalTaskJob and employee_id_name in dicCheckMember and dicCheckMember[employee_id_name]:
                    if pd.isna(row.get('備考', None)) or row['備考'].strip() == '':
                        dfLocalTaskAndNoNote.loc[len(dfLocalTaskAndNoNote)] = [row['社員番号'], row['社員名称: 社員名'], work_name, day]

                dicWebAttendance.setdefault(employee_id_name, {}).setdefault(work_name, {}).setdefault(day, 0)
                dicWebAttendance[employee_id_name][work_name][day] += work_time

        for employee_id_name, last_day in sorted(dicTempLastDay.items(), key=lambda x: x[1]):
            dfOutLastDay.loc[len(dfOutLastDay)] = [employee_id_name.split('-')[0], employee_id_name.split('-')[1], last_day]

        print(f"Successfully read web attendance from {file_path}")

    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")

def check_web_attendance_work_day(file_path: str) -> Optional[pd.DataFrame]:
    """Web勤怠の勤務日とワーク開始・完了予定日の不整合をチェックする"""
    try:
        df = pd.read_csv(file_path, encoding='cp932', usecols=['社員番号', '社員名称: 社員名', '勤務日', 'ワーク番号: ワーク名', 'ワーク番号: ワーク開始予定日', 'ワーク番号: ワーク完了予定日'])
        dfOutWorkDayMismatch = df.query("`ワーク番号: ワーク開始予定日` > `勤務日` or `ワーク番号: ワーク完了予定日` < `勤務日`")
        # rename
        dfOutWorkDayMismatch = dfOutWorkDayMismatch.rename(columns={
            '社員番号': '社員番号',
            '社員名称: 社員名': '社員名',
            '勤務日': '勤務日',
            'ワーク番号: ワーク名': 'ワーク名',
            'ワーク番号: ワーク開始予定日': 'ワーク開始予定日',
            'ワーク番号: ワーク完了予定日': 'ワーク完了予定日'
        })
        # dfOutWorkDayMismatchからdicCheckMemberにいないメンバーを除外
        dfOutWorkDayMismatch = dfOutWorkDayMismatch[
            dfOutWorkDayMismatch.apply(
                lambda row: f"{row['社員番号']}-{row['社員名']}" in dicCheckMember, axis=1
            )
        ]
        return dfOutWorkDayMismatch if not dfOutWorkDayMismatch.empty else None
    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
        return None

def check_work_load_to_web_attendance():
    """工数表とWeb勤怠の不整合をチェックする"""
    for employee_id_name, projectTasks in dicWorkLoad.items():
        if employee_id_name not in dicCheckMember:
            continue
        for projectTask, days in projectTasks.items():
            if '社内工数' in projectTask and not dicCheckMember[employee_id_name]:
                continue
            for day, workloadtime in days.items():
                if workloadtime == 0:
                    continue
                workwebtime = 0
                if employee_id_name in dicWebAttendance:
                    for work_name, web_days in dicWebAttendance[employee_id_name].items():
                        if projectTask in dicWorkLoad2WebAttendance and work_name == dicWorkLoad2WebAttendance[projectTask]:
                            if day in web_days:
                                workwebtime += web_days[day]
                if workloadtime != workwebtime:
                    dfOutWorkLoad.loc[len(dfOutWorkLoad)] = [employee_id_name.split('-')[0], employee_id_name.split('-')[1], projectTask, int(day), round(workloadtime, 2), round(workwebtime, 2)]

def check_web_attendance_to_work_load():
    """Web勤怠と工数表の不整合をチェックする"""
    for employee_id_name, work_days in dicWebAttendance.items():
        if employee_id_name not in dicCheckMember:
            continue
        for work_name, days in work_days.items():
            if work_name in setLocalTaskJob and not dicCheckMember[employee_id_name]:
                continue
            for day, workwebtime in days.items():
                if workwebtime == 0:
                    continue
                workloadtime = 0
                if employee_id_name in dicWorkLoad:
                    for projectTask, projectTasks in dicWorkLoad[employee_id_name].items():
                        if work_name in dicWebAttendance2WorkLoad and projectTask == dicWebAttendance2WorkLoad[work_name]:
                            if day in projectTasks:
                                workloadtime += projectTasks[day]
                if workwebtime != workloadtime:
                    dfOutWebAttendance.loc[len(dfOutWebAttendance)] = [employee_id_name.split('-')[0], employee_id_name.split('-')[1], work_name, int(day), round(workwebtime, 2), round(workloadtime, 2)]

def update_confluence_page(confluence, page_id, body, title, always_update=False):
    """Confluenceページを更新する"""
    try:
        # ページの現在の内容を取得
        page = confluence.get_page_by_id(page_id, expand='body.storage')
        current_body = page['body']['storage']['value']
        
        # 更新が必要かどうかをチェック
        if not always_update and current_body == body:
            print("ページは最新です。更新は行いません。")
            return False
        
        # ページを更新
        confluence.update_page(
            page_id=page_id,
            title=title,
            body=body,
            representation='storage'
        )
        print("ページが更新されました。")
        return True
    except Exception as e:
        print(f"ページの更新中にエラーが発生しました: {e}")
        return False

def send_mail(to_address: str, subject: str, html_body: str) -> None:
    """メールを送信する"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_address
        mail.Subject = subject
        mail.BodyFormat = 2
        mail.HTMLBody = html_body
        mail.Send()
        print(f"メールを送信しました: {to_address}")
    except Exception as e:
        print(f"メール送信エラー: {e}", file=sys.stderr)
        traceback.print_exc()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Web勤怠と工数表の不整合チェックツール"
    )
    parser.add_argument(
        "--config",
        default="config/config.yaml",
        help="config.yaml のパス"
    )
    return parser.parse_args()

def build_attachment_links(files: list[str]) -> str:
    # 事前に本文へ添付リンクをまとめて埋め込む（append_page不要化）
    links = []
    for f in files:
        if not f:
            continue
        fname = os.path.basename(f)
        links.append(
            f"""<p>
                <ac:link>
                    <ri:attachment ri:filename="{fname}"/>
                </ac:link>
            </p>"""
        )
    return "\n".join(links)

def main():
    """メイン処理"""
    logger.info("WorkingTimeChecker start")

    try:
        # コマンドライン引数の解析
        args = parse_args()

        # ---------- config 読み込み ----------
        config = load_config(args.config)

        vars = resolve_target_month(config)
        paths = config["paths"]
        confl_conf = config["confluence"]
        dry_run = config.get("system", {}).get("dry_run", False)
        logger.info(f"dry_run={dry_run}")

        # ---------- ファイルパス ----------
        member_file = expand_path(paths["member_file"], vars)
        workmap_file = expand_path(paths["work_map_file"], vars)
        workload_file = expand_path(paths["workload_aggregate_file"], vars)
        web_attendance_file = expand_path(paths["web_attendance_file"], vars)

        # ---------- Confluence ----------
        confluence_url = confl_conf["dndev"]["url"]
        target_page_id = confl_conf["dndev"]["upload_page_id"]
        cert_dir = confl_conf["dndev"]["cert_dir"]

        # ---------- Mail ----------
        to_mail = config["system"].get("toAddress")

        # ---------- 認証情報 (.env) ----------
        aws_id = os.environ.get("DNDEV_AWS_ID")
        aws_password = os.environ.get("DNDEV_AWS_PASSWORD")

        if not aws_id or not aws_password:
            raise RuntimeError("DNDEV_AWS_ID / DNDEV_AWS_PASSWORD が環境変数にありません")

        # 各ファイルの読み込み
        read_check_member(member_file)
        read_working_map(workmap_file)
        read_working_load(workload_file)    
        read_web_attendance(web_attendance_file)
        
        # 不整合チェック
        check_work_load_to_web_attendance()
        check_web_attendance_to_work_load()
        dfOutWorkDayMismatch = check_web_attendance_work_day(web_attendance_file)

        # ここで添付予定ファイル名のリンクブロックを作る
        files_to_attach = [workload_file, web_attendance_file, workmap_file, member_file]
        attachment_links_block = build_attachment_links(files_to_attach)

        # 結果をHTMLに変換
        html_body = """
        <style>
        table {border-collapse: collapse;}
        th, td {border: 1px solid #ccc; padding: 8px; text-align: left;}
        th {background-color: #f2f2f2;}
        </style>
        """
        html_body += "<h3>※Web勤怠はAppsから取得のため前日又は前々日入力データです</h3>"
        html_body += f"<h2>工数表⇒Web勤怠の不整合チェック結果</h2>"
        html_body += dfOutWorkLoad.sort_values(['社員番号', '日付']).to_html(index=False, escape=False)
        html_body += "<h2>Web勤怠⇒工数表の不整合チェック結果</h2>"
        html_body += dfOutWebAttendance.sort_values(['社員番号', '日付']).to_html(index=False, escape=False)
        html_body += "<h2>Web勤怠の勤務日とワーク開始・完了予定日の不整合チェック結果</h2>"
        if dfOutWorkDayMismatch is not None:
            html_body += dfOutWorkDayMismatch.sort_values(['社員番号', '勤務日']).to_html(index=False, escape=False)
        else:
            html_body += "<p>不整合無し</p>"
        html_body += "<h2>Web勤怠：社内工数で備考未入力の一覧(社内工数チェック対象者のみ)</h2>"
        if dfLocalTaskAndNoNote is None or dfLocalTaskAndNoNote.empty:
            html_body += "<p>該当無し</p>"
        else:
            html_body += dfLocalTaskAndNoNote.sort_values(['社員番号', '日付']).to_html(index=False, escape=False)
        html_body += "<h2>最終勤務入力日</h2>"
        html_body += dfOutLastDay.sort_values(['社員番号']).to_html(index=False, escape=False)
        html_body += "<h2>チェック対象メンバー</h2>"
        html_body += pd.DataFrame(list(dicCheckMember.items()), columns=['社員番号-名前', '社内工数対象']).sort_values(['社員番号-名前']).to_html(index=False, escape=False)
        html_body += "<h2>添付ファイル</h2>"
        html_body += attachment_links_block

        # メール送信
        if to_mail:
            if dry_run:
                logger.info("Dry run mode. Email would be sent to:")
                logger.info(to_mail)
                logger.info("Email subject:")
                logger.info("【自動送信】【PXT連絡】工数表⇔Web勤怠相違")
                logger.info("Email body:")
                logger.info(html_body)
            else:
                send_mail(
                    to_address=to_mail,
                    subject="【自動送信】【PXT連絡】工数表⇔Web勤怠相違",
                    html_body=html_body
                )
                logger.info(f"Email sent to {to_mail}")

        certkeypath = os.path.join(cert_dir, 'client.key')
        certcertpath = os.path.join(cert_dir, 'client.cert')

        # Confluenceページの更新
        if not os.path.exists(certkeypath) or not os.path.exists(certcertpath):
            raise FileNotFoundError(f"証明書ファイルが見つかりません: {certkeypath} または {certcertpath}")

        if target_page_id:
            confluence = Confluence(
                url=confluence_url,
                username=aws_id,
                password=aws_password,
                verify_ssl=True,
                cert=(certcertpath, certkeypath)
            )
            page = confluence.get_page_by_id(page_id=target_page_id)

            if dry_run:
                logger.info("Dry run mode. Confluence page would be updated with the following content:")
                logger.info(f"Page ID: {target_page_id}")
                logger.info(f"Title: 工数表とWeb勤怠の不整合チェック結果 - {vars['YYYY']}/{vars['MM']}")
                logger.info("Body:")
                logger.info(html_body)
                logger.info("Attachments:")
                for file in files_to_attach:
                    if file:
                        logger.info(f"- {file}")
            else:
                if not dry_run:
                    update_confluence_page(
                        confluence=confluence,
                        page_id=target_page_id,
                        body=html_body,
                        title=f"工数表とWeb勤怠の不整合チェック結果 - {vars['YYYY']}/{vars['MM']}",
                        always_update=True
                )

                for file in files_to_attach:
                    if file:
                        confluence.attach_file(
                            page_id=target_page_id,
                            filename=file,
                            name=os.path.basename(file)
                        )

        logger.info("WorkingTimeChecker finished successfully")
    except Exception as e:
        logger.exception(f"エラーが発生しました: {e}")
        raise

if __name__ == "__main__":
    main()
