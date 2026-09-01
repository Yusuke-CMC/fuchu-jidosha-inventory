#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
府中自動車 在庫ページ 自動更新スクリプト

やっていること:
  1. サービスアカウントでGoogle Sheets APIにアクセスし、「車両一覧」シートから
     ステータスが「在庫」の行だけを取り出す（内部限定列は除外）
  2. Apps Script のウェブアプリ（別途デプロイ）を呼び出し、スプレッドシートの
     写真セル（Z列優先・なければE列）に貼られている画像をサムネイルとして取得する
  3. Google Drive の「【写真】在庫車両」フォルダを走査し、各車両（管理№）ごとの
     写真フォルダを見つけて画像をダウンロード・リサイズ・base64化する
     （carousel用の追加写真。サムネイルはこれより2の画像が優先される）
  4. リポジトリ内の index.html を読み込み、埋め込まれている
     FALLBACK_CARS と PHOTO_DATA だけを新しい内容で置き換えて書き戻す
     （デザインやロジック部分のHTML/JS/CSSはそのまま維持される）

このスクリプトは GitHub Actions から定期実行される想定です。
認証情報は環境変数 GOOGLE_SERVICE_ACCOUNT_KEY_FILE （JSONキーファイルのパス）から読み込みます。
Apps Script ウェブアプリのURLは環境変数 APPS_SCRIPT_WEB_APP_URL から読み込みます
（未設定の場合はこの手順をスキップし、Driveの写真だけを使う）。
"""

import base64
import io
import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from PIL import Image

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:  # noqa: BLE001
    pass

# ============ 設定 ============

# スプレッドシートID（URLの /d/ と /edit の間の文字列）
SPREADSHEET_ID = "1o9Flp2R2MwFzCMgfzHT1ImisTOSCy5IOgzxv0MpM8Qc"
SHEET_NAME = "車両一覧"

# Google Drive 上の「【写真】在庫車両」フォルダのID
PHOTO_ROOT_FOLDER_ID = "1rUxZ3BD_GUhRbcLaZi7aWcaPKu0jbx3B"

# 公開ページのファイル
INDEX_HTML_PATH = os.path.join(os.path.dirname(__file__), "..", "index.html")

# 掲載しない内部限定の列（見出し文字列に部分一致すれば除外）
INTERNAL_ONLY_COLUMNS = ["商談担当者", "仕入れ担当者", "期日メモ", "保険", "車台番号下3桁"]

# 写真のリサイズ設定（サイト掲載用：軽量寄り）
PHOTO_MAX_SIDE = 900
PHOTO_QUALITY = 62

# 除外するフォルダ名（管理№のフォルダではないもの）
NON_NUMBER_FOLDER_NAMES = {"売却済み", "0 例", "【写真】在庫車"}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def get_credentials():
    key_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_KEY_FILE")
    if not key_path or not os.path.exists(key_path):
        print("ERROR: GOOGLE_SERVICE_ACCOUNT_KEY_FILE が設定されていません", file=sys.stderr)
        sys.exit(1)
    return service_account.Credentials.from_service_account_file(key_path, scopes=SCOPES)


def fetch_sheet_rows(sheets_service):
    """車両一覧シートを全部読み込み、ヘッダー行をキーにした辞書のリストで返す"""
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=SPREADSHEET_ID, range=f"{SHEET_NAME}!A:AZ")
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return []
    headers = values[0]
    rows = []
    for raw_row in values[1:]:
        row = {}
        for i, header in enumerate(headers):
            row[header] = raw_row[i] if i < len(raw_row) else ""
        rows.append(row)
    return rows


def find_header(headers, *keywords):
    """見出しの中から、いずれかのキーワードを含むものを探す（部分一致）"""
    for h in headers:
        for kw in keywords:
            if kw in h:
                return h
    return None


def build_car_records(rows):
    if not rows:
        return []
    headers = list(rows[0].keys())

    col_no = find_header(headers, "管理№", "管理No", "管理番号")
    col_status = find_header(headers, "ステータス")
    col_name = find_header(headers, "車名")
    col_grade = find_header(headers, "グレード")
    col_color = find_header(headers, "ボディカラー", "色")
    col_category = find_header(headers, "車格")
    col_regdate = find_header(headers, "初年度登録年月")
    col_mileage = find_header(headers, "走行距離(km)", "走行距離")
    col_shaken = find_header(headers, "車検")
    col_price = find_header(headers, "支払総額")
    col_equip = find_header(headers, "装備")
    col_repair = find_header(headers, "修復歴")
    col_inspection = find_header(headers, "整備記録簿", "点検")

    missing = [
        name
        for name, col in [
            ("管理№", col_no),
            ("ステータス", col_status),
            ("車名", col_name),
        ]
        if col is None
    ]
    if missing:
        print(f"ERROR: 必須列が見つかりません: {missing}", file=sys.stderr)
        print(f"実際の見出し: {headers}", file=sys.stderr)
        sys.exit(1)

    cars = []
    for row in rows:
        if (row.get(col_status) or "").strip() != "在庫":
            continue
        no = (row.get(col_no) or "").strip()
        if not no:
            continue
        cars.append(
            {
                "no": no,
                "status": "在庫",
                "name": (row.get(col_name) or "").strip(),
                "grade": (row.get(col_grade) or "").strip() if col_grade else "",
                "color": (row.get(col_color) or "").strip() if col_color else "",
                "category": (row.get(col_category) or "").strip() if col_category else "",
                "regDate": (row.get(col_regdate) or "").strip() if col_regdate else "",
                "mileage": (row.get(col_mileage) or "").strip() if col_mileage else "",
                "shaken": (row.get(col_shaken) or "").strip() if col_shaken else "",
                "price": (row.get(col_price) or "").strip() if col_price else "",
                "equip": (row.get(col_equip) or "").strip() if col_equip else "",
                "basePrice": "",
                "repair": (row.get(col_repair) or "").strip() if col_repair else "",
                "inspection": (row.get(col_inspection) or "").strip() if col_inspection else "",
            }
        )
    return cars


def list_children(drive_service, folder_id):
    files = []
    page_token = None
    while True:
        resp = (
            drive_service.files()
            .list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType)",
                pageToken=page_token,
                pageSize=200,
            )
            .execute()
        )
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def build_number_to_folder_map(drive_service):
    """管理№(文字列) -> 写真フォルダID のマップを作る"""
    mapping = {}
    range_folders = [
        f
        for f in list_children(drive_service, PHOTO_ROOT_FOLDER_ID)
        if f["mimeType"] == "application/vnd.google-apps.folder"
        and f["name"] not in NON_NUMBER_FOLDER_NAMES
    ]
    for rf in range_folders:
        children = list_children(drive_service, rf["id"])
        for child in children:
            if child["mimeType"] != "application/vnd.google-apps.folder":
                continue
            name = child["name"].strip()
            if re.fullmatch(r"\d+", name):
                mapping[name] = child["id"]
    return mapping


def download_and_resize(drive_service, file_id):
    request = drive_service.files().get_media(fileId=file_id)
    raw = request.execute()
    img = Image.open(io.BytesIO(raw)).convert("RGB")
    w, h = img.size
    scale = PHOTO_MAX_SIDE / max(w, h)
    if scale < 1:
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=PHOTO_QUALITY, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def fetch_thumbnail_data():
    """Apps Script ウェブアプリから、管理№ -> サムネイル画像(data URL) の辞書を取得する。
    Z列優先・なければE列、どちらもなければ null が返る（Apps Script側のロジック）。
    URLが未設定、または取得に失敗した場合は空の辞書を返し、Drive写真だけで動作を継続する。
    """
    url = os.environ.get("APPS_SCRIPT_WEB_APP_URL")
    if not url:
        print("INFO: APPS_SCRIPT_WEB_APP_URL 未設定のため、スプレッドシートのセル写真は取得しません")
        return {}
    try:
        resp = requests.get(url, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            print(f"WARN: Apps Script側でエラー: {data['error']}", file=sys.stderr)
            return {}
        return {k: v for k, v in data.items() if v}
    except Exception as e:  # noqa: BLE001
        print(f"WARN: Apps Scriptからのサムネイル取得に失敗: {e}", file=sys.stderr)
        return {}


def build_photo_data(drive_service, cars, number_to_folder, thumbnail_data):
    photo_data = {}
    for car in cars:
        no = car["no"]
        folder_id = number_to_folder.get(no)
        encoded = []
        if folder_id:
            image_files = [
                f
                for f in list_children(drive_service, folder_id)
                if f["mimeType"].startswith("image/")
            ]
            # 簡易的なカバー写真選定: ファイル名の昇順で先頭を採用
            # (人手でのアングル選定ほど精緻ではないが、無人実行のための現実的な妥協)
            image_files.sort(key=lambda f: f["name"])
            for f in image_files:
                try:
                    encoded.append(f"data:image/jpeg;base64,{download_and_resize(drive_service, f['id'])}")
                except Exception as e:  # noqa: BLE001
                    print(f"WARN: 写真取得失敗 {no}/{f['name']}: {e}", file=sys.stderr)

        thumb = thumbnail_data.get(no)
        if thumb:
            # スプレッドシートのセル写真（Z列優先・なければE列）を先頭（サムネイル）に採用
            encoded = [thumb] + encoded

        photo_data[no] = encoded
        source = "セル写真+Drive" if thumb and encoded[1:] else ("セル写真のみ" if thumb else "Driveのみ")
        print(f"  管理№{no}: 写真{len(encoded)}枚 ({source})")
    return photo_data


def extract_balanced(s, start_idx, open_ch, close_ch):
    depth = 0
    i = start_idx
    started = False
    while i < len(s):
        c = s[i]
        if c == open_ch:
            depth += 1
            started = True
        elif c == close_ch:
            depth -= 1
            if started and depth == 0:
                return i
        i += 1
    raise ValueError("unbalanced")


def update_html(cars, photo_data):
    with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    marker1 = "const FALLBACK_CARS = "
    idx1 = content.find(marker1)
    if idx1 == -1:
        print("ERROR: index.html 内に FALLBACK_CARS が見つかりません", file=sys.stderr)
        sys.exit(1)
    arr_start1 = idx1 + len(marker1)
    arr_end1 = extract_balanced(content, arr_start1, "[", "]")
    content = content[:arr_start1] + json.dumps(cars, ensure_ascii=False) + content[arr_end1 + 1:]

    marker2 = "const PHOTO_DATA = "
    idx2 = content.find(marker2)
    if idx2 == -1:
        print("ERROR: index.html 内に PHOTO_DATA が見つかりません", file=sys.stderr)
        sys.exit(1)
    obj_start2 = idx2 + len(marker2)
    obj_end2 = extract_balanced(content, obj_start2, "{", "}")
    content = content[:obj_start2] + json.dumps(photo_data, ensure_ascii=False) + content[obj_end2 + 1:]

    # フッターのスナップショット日付を更新（JST基準）
    jst = timezone(timedelta(hours=9))
    today = datetime.now(jst)
    date_str = f"{today.year}年{today.month}月{today.day}日"
    content = re.sub(r"\d{4}年\d{1,2}月\d{1,2}日 時点のスナップショットを表示", f"{date_str} 時点のスナップショットを表示", content)

    with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(content)


def main():
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    print("スプレッドシートを読み込み中...")
    rows = fetch_sheet_rows(sheets_service)
    cars = build_car_records(rows)
    print(f"在庫車両: {len(cars)}台")

    print("スプレッドシートのセル写真（サムネイル用）を取得中...")
    thumbnail_data = fetch_thumbnail_data()
    print(f"  セル写真が見つかった台数: {len(thumbnail_data)}")

    print("Driveの写真フォルダ構成を取得中...")
    number_to_folder = build_number_to_folder_map(drive_service)

    print("写真を取得中...")
    photo_data = build_photo_data(drive_service, cars, number_to_folder, thumbnail_data)

    print("index.html を更新中...")
    update_html(cars, photo_data)

    print("完了")


if __name__ == "__main__":
    main()
