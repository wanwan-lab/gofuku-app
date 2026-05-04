"""
商品在庫・販売管理 (Streamlit)

st.secrets に以下を設定してください（例は .streamlit/secrets.toml）。

必須キー:
  GEMINI_API_KEY
  GAS_UPLOAD_URL           … 画像を Google ドライブに保存する Web アプリ（GAS）の URL
  GAS_API_KEY                … GAS Web アプリ呼び出し用の共有キー（payload の apiKey に付与）
  GOOGLE_DRIVE_FOLDER_ID     … 保存先フォルダID（GAS に渡す）
  google_service_account     … サービスアカウントJSONの各フィールド（[google_service_account] セクション）
    または GOOGLE_SERVICE_ACCOUNT_JSON … JSON文字列1本

スプレッドシート運用時に追加で必須:
  GOOGLE_SPREADSHEET_ID      … 記録用スプレッドシートID（``INVENTORY_SOURCE=csv`` のときは不要で、共有 ``inventory.csv`` を使用）

任意:
  GEMINI_MODEL_NAME          … Gemini モデル ID（未設定時は下記 DEFAULT_GEMINI_MODEL）
  GEMINI_VOUCHER_MODEL_NAME  … 証憑（納品書等）画像解析専用モデル（未設定時は GEMINI_MODEL_NAME と同じ既定）
  GOOGLE_WORKSHEET_NAME      … ワークシート名（未設定時は DEFAULT_WORKSHEET_NAME）
  GAS_UPLOAD_TIMEOUT_SECONDS … GAS への POST タイムアウト秒（既定 300、1〜3600 にクランプ）
  FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED … GAS 未設定時に台帳の画像URL列へ入れるプレースホルダ URL
  APP_PASSWORD               … アプリ画面の簡易ログイン用（平文。GitHub には secrets.toml をコミットしないこと）
  INVENTORY_SOURCE           … ``csv`` | ``sheet`` 。未指定時は **GOOGLE_SPREADSHEET_ID があるとき sheet**、無いとき **csv**（リポジトリ直下 ``inventory.csv`` または環境変数 ``GOFUKU_INVENTORY_CSV``）。

※ 画像の Gemini 解析は **google-generativeai** を使用します。モデル名は ``GEMINI_MODEL_NAME``（既定は flash 系プレビュー）。
※ **登録（インプット）** ページの証憑取込で、納品書・請求書・領収書を画像・PDF・Excel・Word から解析し入庫（購入）として台帳に追記できます（確定前に表で編集可能）。
※ PDF/Excel/Word 取込には ``pypdf`` / ``pymupdf`` / ``openpyxl`` / ``python-docx`` を使用します（requirements.txt）。
※ アップロード画像は任意。商品写真は Pillow で長辺最大1280px・JPEG品質80に変換してから解析・ドライブ保存します。
※ 証憑ファイルを台帳に確定反映するとき、Drive 保存用には画像のみ長辺最大2000px・JPEG品質75に再圧縮します（PDF/xlsx/docx は原本のまま）。
※ 台帳日時・撮影日時未取得時の現在時刻は **pytz** の ``Asia/Tokyo``（JST）です。

サイドバーで **登録（インプット）** / **ギャラリー（カタログ）・在庫一覧** / **集計・分析（ダッシュボード）** を切り替えられます。
在庫データは **共有の inventory.csv**（ローカル）または **Google スプレッドシート**（``INVENTORY_SOURCE`` で選択）に読み書きします。
列定義・CSV 入出力は **app.py 内に内包**しています。

スプレッドシート1行目はヘッダーとして次の列順を想定:
  日時 | 商品名 | 仕入先・取引先 | 数量 | 仕入金額（税抜） | 仕入金額（税込）
  | 販売予定金額（税抜） | 販売予定金額（税込） | 実売金額（税抜） | 実売金額（税込） | 粗利 | ステータス（在庫中/販売済） | メモ（任意） | 画像URL | 管理ID | 最後に確認した日付（棚卸日） | 販売元管理ID | 証憑記録日時 | 証憑URL
  | 仕入日時 | 入庫種別 | 浮貸日時 | 販売日時 | 出庫種別
  ※在庫は **1点につき1行** で統一します。登録時の行数は **数量** と同じで、各行の数量は **1** です。
  ※写真は **1枚まで** アップロードできます。写真があるときは1回だけドライブに保存し、数量が **2以上** のときは **全行に同じ画像URL** を入れます（数量が1のときはその1行のみ）。
  ※「管理ID」列は自動採番（例: G00000001）のシリアルです。既存行の末尾に列を追加しても列位置はずれません。
  ※「**日時**」列（A列）は **その行が最後に台帳へ保存された時点の JST 時刻**（登録・販売反映・一覧からの保存など）です。**仕入日時** は仕入の暦（EXIF 等を ``record_datetime`` に渡した値）、**入庫種別** は登録画面の区分（入庫（購入）・入庫（返品）・入庫（浮貸）等）です。**販売日時**・**出庫種別** は販売確定時に記録します。
  ※旧シートの「入出庫種別」列は読み込み時に **入庫種別** へ移して無視します（ヘッダーは新列順に更新されます）。
  ※「仕入金額（税抜）」「仕入金額（税込）」は **1点あたりの行合計**（台帳の各行は数量1）です。
  ※旧シートに「仕入単価（税抜）」列が残っている場合は、読み込み時にその列を除いて新しい列構成に揃えます。
  ※新規登録画面では仕入金額（税込）の計算に使う消費税を **10% / 8% / 非課税** から選べます（既定は10%）。
  ※「販売予定金額（税抜）」「実売金額（税抜）」には **1点あたりの税抜金額** を保存し（数量と掛けて行計にします）、税込総額列はその税抜行合計に、仕入行と同じ税率で四捨五入します。
  ※金額列（仕入〜粗利まで）は書き込み時に表示形式 **#,##0** を適用します。
  ※粗利は税抜ベースで「販売済」なら（実売金額（税抜）×数量）−原価、「在庫中」なら（販売予定金額（税抜）×数量）−原価。台帳保存時に再計算します。
  ※「最後に確認した日付（棚卸日）」は棚卸作業用の任意列です（YYYY-MM-DD 推奨）。1人棚卸しの進捗把握に使います。
  ※「販売元管理ID」は **販売管理** タブで **在庫中の行の管理ID（G########）** を指定します。出庫（販売）または出庫（浮貸）で **販売済** にするときは実売が必須で、各IDの行を順に更新します（新規行は追加しません）。出庫（浮貸）で **在庫中** のままにするときは **浮貸日時** 列に確定時の JST（または手入力）を記録します。
  ※「証憑記録日時」は証憑取込の **確定ボタンを押した JST 時刻**（recorded_at に相当）。「証憑URL」はその証憑を GAS 経由で Drive に保存したときの表示 URL（evidence_url）です。
  ※台帳一覧から手動で在庫行を販売済に編集する場合は、**販売日時**・**出庫種別**・実売・ステータスを整合させてください（保存時に変更があった行の「日時」は自動で更新されます）。
"""

from __future__ import annotations

import base64
import difflib
import io
import json
import math
import os
import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

import altair as alt
import numpy as np
import google.generativeai as genai
import gspread
import pandas as pd
import pytz
import requests
import streamlit as st
from google.oauth2 import service_account
from PIL import Image, ImageOps

# --- スプレッドシート列・税率（app 単体で完結：サブモジュール未コミットでも Cloud で動く） ---
COL_DATETIME = "日時"
# 旧シート互換（読み込み時のみ。列は台帳から廃止）
LEGACY_COL_MOVEMENT_TYPE = "入出庫種別"
COL_NAME = "商品名"
COL_SUPPLIER = "仕入先・取引先"
COL_QTY = "数量"
COL_PRICE_EXCL = "仕入金額（税抜）"
COL_PRICE_INCL = "仕入金額（税込）"
LEGACY_COL_UNIT_PRICE = "仕入単価（税抜）"
COL_PLANNED_SALE = "販売予定金額（税抜）"
COL_PLANNED_SALE_INCL = "販売予定金額（税込）"
COL_ACTUAL_SALE = "実売金額（税抜）"
COL_ACTUAL_SALE_INCL = "実売金額（税込）"
COL_GROSS_PROFIT = "粗利"
COL_STOCK_STATUS = "ステータス（在庫中/販売済）"
COL_IMAGE_URL = "画像URL"
COL_MEMO = "メモ"
COL_MANAGEMENT_ID = "管理ID"
COL_LAST_STOCKTAKE = "最後に確認した日付（棚卸日）"
COL_SALE_SOURCE_MGMT_ID = "販売元管理ID"
COL_LOAN_DATETIME = "浮貸日時"
# 証憑取込（recorded_at / evidence_url に相当）
COL_VOUCHER_RECORDED_AT = "証憑記録日時"
COL_VOUCHER_EVIDENCE_URL = "証憑URL"
# 仕入と販売を同一行で分離（「日時」は行の最終更新時刻。仕入の暦は下記の仕入日時・入庫種別）
COL_PURCHASE_DATETIME = "仕入日時"
COL_PURCHASE_MOVEMENT = "入庫種別"
COL_SALE_DATETIME = "販売日時"
COL_SALE_OUTBOUND_TYPE = "出庫種別"

STATUS_IN_STOCK = "在庫中"
STATUS_SOLD = "販売済"
STOCK_STATUS_OPTIONS: tuple[str, ...] = (STATUS_IN_STOCK, STATUS_SOLD)

# 登録画面「数量」の number_input 専用キー。
# 旧実装で max_value=1・disabled が付いていた端末では、同一キーのままだと数量が増やせないことがあるため分離する。
REGISTRATION_QTY_WIDGET_KEY = "registration_qty_input"
SALES_TAB_QTY_WIDGET_KEY = "sales_tab_registration_qty"


def _movement_is_outbound(mv: str) -> bool:
    """入出庫種別が出庫（販売・浮貸など）かどうか。"""
    return (mv or "").strip().startswith("出庫")


CONSUMPTION_TAX_RATE = 0.10
CONSUMPTION_TAX_CHOICE_TO_RATE: dict[str, float] = {
    "10%": 0.10,
    "8%": 0.08,
    "非課税": 0.0,
}

EXPECTED_HEADERS: list[str] = [
    COL_DATETIME,
    COL_NAME,
    COL_SUPPLIER,
    COL_QTY,
    COL_PRICE_EXCL,
    COL_PRICE_INCL,
    COL_PLANNED_SALE,
    COL_PLANNED_SALE_INCL,
    COL_ACTUAL_SALE,
    COL_ACTUAL_SALE_INCL,
    COL_GROSS_PROFIT,
    COL_STOCK_STATUS,
    COL_MEMO,
    COL_IMAGE_URL,
    COL_MANAGEMENT_ID,
    COL_LAST_STOCKTAKE,
    COL_SALE_SOURCE_MGMT_ID,
    COL_VOUCHER_RECORDED_AT,
    COL_VOUCHER_EVIDENCE_URL,
    COL_PURCHASE_DATETIME,
    COL_PURCHASE_MOVEMENT,
    COL_LOAN_DATETIME,
    COL_SALE_DATETIME,
    COL_SALE_OUTBOUND_TYPE,
]

_INVENTORY_CSV_DEFAULT_NAME = "inventory.csv"


def _inventory_csv_path() -> Path:
    raw = (os.environ.get("GOFUKU_INVENTORY_CSV") or "").strip()
    if raw:
        return Path(raw).expanduser().resolve()
    return Path(__file__).resolve().parent / _INVENTORY_CSV_DEFAULT_NAME


@st.cache_data(show_spinner=False)
def _inventory_csv_read_df_cached(mtime_ns: int, path_str: str) -> pd.DataFrame:
    """CSV の内容を mtime 付きでキャッシュ（同一内容の連続再描画を軽くする）。"""
    _ = mtime_ns
    path = Path(path_str)
    if not path.exists():
        return pd.DataFrame(columns=EXPECTED_HEADERS)
    df = pd.read_csv(path, encoding="utf-8-sig")
    if LEGACY_COL_MOVEMENT_TYPE in df.columns:
        if COL_PURCHASE_MOVEMENT not in df.columns:
            df[COL_PURCHASE_MOVEMENT] = ""
        m = df[COL_PURCHASE_MOVEMENT].fillna("").astype(str).str.strip() == ""
        df.loc[m, COL_PURCHASE_MOVEMENT] = (
            df.loc[m, LEGACY_COL_MOVEMENT_TYPE].fillna("").astype(str)
        )
        df = df.drop(columns=[LEGACY_COL_MOVEMENT_TYPE], errors="ignore")
    for c in EXPECTED_HEADERS:
        if c not in df.columns:
            df[c] = ""
    return df[EXPECTED_HEADERS].copy()


def _inventory_csv_read_df() -> pd.DataFrame:
    path = _inventory_csv_path()
    if not path.exists():
        return pd.DataFrame(columns=EXPECTED_HEADERS)
    try:
        mt = int(path.stat().st_mtime_ns)
    except OSError:
        mt = 0
    return _inventory_csv_read_df_cached(mt, str(path.resolve()))


def _inventory_csv_write_df(df: pd.DataFrame) -> None:
    path = _inventory_csv_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    out = df.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    out.to_csv(path, index=False, encoding="utf-8-sig")


# --- st.secrets のキー名（文字列リテラルの散在を避ける） ---
SECRET_GEMINI_API_KEY = "GEMINI_API_KEY"
SECRET_GEMINI_MODEL_NAME = "GEMINI_MODEL_NAME"
SECRET_GEMINI_VOUCHER_MODEL_NAME = "GEMINI_VOUCHER_MODEL_NAME"
SECRET_GAS_UPLOAD_URL = "GAS_UPLOAD_URL"
SECRET_GAS_API_KEY = "GAS_API_KEY"
SECRET_GAS_UPLOAD_TIMEOUT_SECONDS = "GAS_UPLOAD_TIMEOUT_SECONDS"
SECRET_GOOGLE_DRIVE_FOLDER_ID = "GOOGLE_DRIVE_FOLDER_ID"
SECRET_GOOGLE_SPREADSHEET_ID = "GOOGLE_SPREADSHEET_ID"
SECRET_GOOGLE_WORKSHEET_NAME = "GOOGLE_WORKSHEET_NAME"
SECRET_GOOGLE_SERVICE_ACCOUNT_JSON = "GOOGLE_SERVICE_ACCOUNT_JSON"
SECRET_GOOGLE_SERVICE_ACCOUNT_SECTION = "google_service_account"
SECRET_APP_PASSWORD = "APP_PASSWORD"
SECRET_INVENTORY_SOURCE = "INVENTORY_SOURCE"
SECRET_FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED = "FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED"

# --- secrets に無いときの既定（非機密のデフォルトのみ） ---
DEFAULT_GEMINI_MODEL = "gemini-3-flash-preview"
DEFAULT_WORKSHEET_NAME = "在庫履歴"
DEFAULT_GAS_FALLBACK_IMAGE_URL = "https://example.com/?gofuku-app=skipped-no-gas-secrets"
DEFAULT_GAS_UPLOAD_TIMEOUT_SECONDS = 300

# --- 画像アップロード前処理 ---
UPLOAD_JPEG_MAX_LONG_EDGE = 1280
UPLOAD_JPEG_QUALITY = 80
# 仕入れ登録タブから Drive へ保存する商品写真（長辺・品質）
PURCHASE_DRIVE_JPEG_MAX_LONG_EDGE = 2000
PURCHASE_DRIVE_JPEG_QUALITY = 75

# --- 証憑ファイルを Google ドライブへ送る直前の軽量化（画像のみ） ---
VOUCHER_DRIVE_JPEG_MAX_LONG_EDGE = 2000
VOUCHER_DRIVE_JPEG_QUALITY = 75

SHEET_AMOUNT_NUMBER_PATTERN = "#,##0"
TZ_JP = pytz.timezone("Asia/Tokyo")
LEDGER_DATA_EDITOR_KEY = "inventory_ledger_data_editor"
INV_GALLERY_PAGE_SIZE = 30
SESSION_KEY_INV_SHEET_CACHE_BUST = "_inv_sheet_cache_bust"
VOUCHER_DATA_EDITOR_KEY = "voucher_inventory_preview_editor"
LEDGER_PICK_PLACEHOLDER = "（選ばない）"

# 証憑取込（Gemini への共通指示・JSON 仕様）
VOUCHER_EXTRACTION_RULES = """注意点:
- 商品名が曖昧な場合（例：着物一式）は、入力内の他の情報から「振袖」や「訪問着」などのカテゴリーを推測し、items[].category と name の両方に反映できるよう整理してください。
- 手書きの修正や合計金額がある場合は、最新の数値を優先してください。
- 出力は、プログラムで解析可能な純粋な JSON オブジェクトのみとしてください（説明文・Markdown・コードフェンスは禁止）。"""

VOUCHER_JSON_SPEC = """JSON の構造（キーは次のみ。値の型を守ること）:
{
  "supplier_name": "仕入先名（文字列。読めなければ空文字）",
  "purchase_date": "YYYY-MM-DD（暦日のみ。読めなければ空文字）",
  "items": [
    {
      "name": "商品名",
      "quantity": 整数（1以上。読めなければ1）,
      "unit_price": 1点または当該明細行の税抜単価・金額（円の整数。読めなければ null）,
      "category": "推定区分（振袖・訪問着・帯など。不明なら空文字）"
    }
  ]
}
unit_price は原則として税抜の円整数。明細が税込のみのときは税抜に換算して整数で入れてください。"""


def check_password() -> bool:
    """認証済みになるまで在庫アプリ本体を起動しない。未認証時は認証UIのみ表示し st.stop() する。"""
    if st.session_state.get("authenticated"):
        return True

    st.set_page_config(page_title="認証", layout="centered")
    st.header("認証")

    raw_pw = st.secrets.get(SECRET_APP_PASSWORD)
    if raw_pw is None:
        st.error(
            f"`.streamlit/secrets.toml` に **{SECRET_APP_PASSWORD}** を設定してください。"
            "（ローカルで `secrets.toml` が無い場合は作成してください）"
        )
        st.stop()
    expected = str(raw_pw).strip()

    if not expected:
        st.error(f"{SECRET_APP_PASSWORD} が空です。secrets.toml を確認してください。")
        st.stop()

    with st.form("auth_screen_form"):
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button(
            "ログイン（パスワード入力後は Enter でも送信できます）"
        )

    if submitted:
        if (password or "").strip() == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("パスワードが正しくありません。")

    st.stop()


def _apply_inventory_amount_number_formats(ws) -> None:
    """金額系の列（仕入金額（税抜）〜粗利まで）に、2行目以降で #,##0 を適用する。"""
    idx_start = EXPECTED_HEADERS.index(COL_PRICE_EXCL)
    idx_end = EXPECTED_HEADERS.index(COL_GROSS_PROFIT)
    end_row = max(int(ws.row_count), 2)
    ws.spreadsheet.batch_update(
        {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": idx_start,
                            "endColumnIndex": idx_end + 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "NUMBER",
                                    "pattern": SHEET_AMOUNT_NUMBER_PATTERN,
                                }
                            }
                        },
                        "fields": "userEnteredFormat.numberFormat",
                    }
                }
            ]
        }
    )


def jst_now() -> datetime:
    """現在の日本時間（JST・timezone-aware）。"""
    return datetime.now(TZ_JP)


def jst_now_str() -> str:
    """スプレッドシート用の日時文字列（JST・秒まで）。"""
    return jst_now().strftime("%Y-%m-%d %H:%M:%S")


def capture_datetime_jst_from_bytes(raw: bytes) -> str | None:
    """画像バイナリの EXIF から撮影日時を読み、JST の壁時計として解釈して文字列化する。

    リサイズ前の元データに対して呼ぶこと（EXIF 失効前に取得する）。
    EXIF にタイムゾーンが無いため、取得値は **日本のローカル時刻** として ``Asia/Tokyo`` に固定する。
    失敗時は ``None``（呼び出し側で ``jst_now_str()`` をデフォルトにする）。
    """
    try:
        with Image.open(io.BytesIO(raw)) as img:
            exif = img.getexif()
        if not exif:
            return None
        dt_s: str | None = None
        try:
            from PIL.ExifTags import IFD

            sub = exif.get_ifd(IFD.Exif)
            if sub:
                dt_s = sub.get(36867) or sub.get(36868)  # DateTimeOriginal, DateTimeDigitized
            if not dt_s:
                sub0 = exif.get_ifd(IFD.IFD0)
                if sub0:
                    dt_s = sub0.get(306)  # DateTime
        except Exception:
            pass
        if not dt_s:
            dt_s = exif.get(36867) or exif.get(306)
        if not dt_s:
            return None
        if isinstance(dt_s, bytes):
            dt_s = dt_s.decode("utf-8", errors="ignore")
        dt_s = str(dt_s).strip()
        naive: datetime | None = None
        for fmt in ("%Y:%m:%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                naive = datetime.strptime(dt_s, fmt)
                break
            except ValueError:
                continue
        if naive is None:
            return None
        aware = TZ_JP.localize(naive)
        return aware.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def capture_datetime_jst_from_upload(uploaded) -> str | None:
    """UploadedFile から EXIF 日時を取得（内部は :func:`capture_datetime_jst_from_bytes`）。"""
    try:
        return capture_datetime_jst_from_bytes(uploaded.getvalue())
    except Exception:
        return None


def _resize_long_edge_max(img: Image.Image, max_edge: int) -> Image.Image:
    w, h = img.size
    long_edge = max(w, h)
    if long_edge <= max_edge:
        return img
    scale = max_edge / float(long_edge)
    nw = max(1, int(round(w * scale)))
    nh = max(1, int(round(h * scale)))
    return img.resize((nw, nh), Image.Resampling.LANCZOS)


def prepare_upload_image_jpeg(
    raw: bytes,
    *,
    max_long_edge: int | None = None,
    quality: int | None = None,
) -> tuple[bytes, str]:
    """Gemini 送信用・GAS 保存用の共通前処理。

    EXIF 向き補正のうえ、長辺リサイズと JPEG 再エンコードを行う。
    ``max_long_edge`` / ``quality`` を省略したときは :data:`UPLOAD_JPEG_MAX_LONG_EDGE` と
    :data:`UPLOAD_JPEG_QUALITY` を使う（Gemini 向けの既定）。

    Returns:
        (jpeg_bytes, mime_type)  mime_type は常に ``image/jpeg`` 。
    """
    mx = int(max_long_edge) if max_long_edge is not None else UPLOAD_JPEG_MAX_LONG_EDGE
    q = int(quality) if quality is not None else UPLOAD_JPEG_QUALITY
    mx = max(256, min(mx, 8192))
    q = max(40, min(q, 95))
    img = Image.open(io.BytesIO(raw))
    img = ImageOps.exif_transpose(img)
    rgba = img.convert("RGBA")
    bg = Image.new("RGB", rgba.size, (255, 255, 255))
    bg.paste(rgba, mask=rgba.getchannel("A"))
    img = bg
    img = _resize_long_edge_max(img, mx)
    buf = io.BytesIO()
    img.save(
        buf,
        format="JPEG",
        quality=q,
        optimize=True,
        progressive=True,
    )
    return buf.getvalue(), "image/jpeg"


def _secret_str(key: str, default: str = "") -> str:
    """設定文字列を ``st.secrets.get`` で取得。欠損・空は default。"""
    v = st.secrets.get(key, default)
    if v is None:
        return default
    s = str(v).strip()
    return s if s else default


def _uses_local_inventory_csv() -> bool:
    """``INVENTORY_SOURCE=csv`` なら True。``sheet`` 明示なら False。未指定時はスプレッドシートIDが無ければ CSV。"""
    v = _secret_str(SECRET_INVENTORY_SOURCE, "").lower()
    if v == "csv":
        return True
    if v == "sheet":
        return False
    return not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID)


def _secret_int(
    key: str, default: int, *, min_value: int = 1, max_value: int = 10_000
) -> int:
    raw = st.secrets.get(key)
    if raw is None:
        return default
    s = str(raw).strip()
    if not s:
        return default
    try:
        return max(min_value, min(max_value, int(s)))
    except ValueError:
        return default


def _gas_upload_timeout_seconds() -> int:
    return _secret_int(
        SECRET_GAS_UPLOAD_TIMEOUT_SECONDS,
        DEFAULT_GAS_UPLOAD_TIMEOUT_SECONDS,
        min_value=30,
        max_value=3600,
    )


def _fallback_image_url_when_gas_unconfigured() -> str:
    return _secret_str(
        SECRET_FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED,
        DEFAULT_GAS_FALLBACK_IMAGE_URL,
    )


def _gemini_model_name() -> str:
    return _secret_str(SECRET_GEMINI_MODEL_NAME, DEFAULT_GEMINI_MODEL)


def _gemini_voucher_model_name() -> str:
    """証憑画像解析用モデル。専用キーが空なら通常の Gemini モデル名にフォールバック。"""
    v = _secret_str(SECRET_GEMINI_VOUCHER_MODEL_NAME, "")
    return v if v else _gemini_model_name()


def _load_service_account_info() -> dict[str, Any]:
    raw_json = st.secrets.get(SECRET_GOOGLE_SERVICE_ACCOUNT_JSON)
    if raw_json is not None and str(raw_json).strip():
        if isinstance(raw_json, str):
            return json.loads(raw_json)
        raise ValueError(
            f"{SECRET_GOOGLE_SERVICE_ACCOUNT_JSON} は文字列である必要があります。"
        )
    ga = st.secrets.get(SECRET_GOOGLE_SERVICE_ACCOUNT_SECTION)
    if ga is not None:
        if hasattr(ga, "to_dict"):
            return dict(ga.to_dict())
        return dict(ga)
    raise ValueError(
        "サービスアカウントが見つかりません。"
        f" [{SECRET_GOOGLE_SERVICE_ACCOUNT_SECTION}] セクションか "
        f"{SECRET_GOOGLE_SERVICE_ACCOUNT_JSON} を設定してください。"
    )


def _credentials():
    info = _load_service_account_info()
    scopes = ("https://www.googleapis.com/auth/spreadsheets",)
    return service_account.Credentials.from_service_account_info(info, scopes=scopes)


@st.cache_resource
def _gspread_client():
    return gspread.authorize(_credentials())


def _parse_json_from_model(text: str) -> dict[str, Any]:
    """モデル出力から JSON オブジェクトを抽出する（コードフェンスや前後の説明文を許容）。"""
    t = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", t)
    if fence:
        t = fence.group(1).strip()
    try:
        obj = json.loads(t)
    except json.JSONDecodeError:
        i0 = t.find("{")
        i1 = t.rfind("}")
        if i0 == -1 or i1 <= i0:
            raise
        obj = json.loads(t[i0 : i1 + 1])
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        obj = obj[0]
    if not isinstance(obj, dict):
        raise ValueError("JSON がオブジェクト形式ではありません。")
    return obj


def _coerce_positive_int(val: Any, default: int = 1) -> int:
    try:
        n = int(float(str(val).strip()))
        return max(1, n)
    except Exception:
        return default


def _coerce_unit_price_yen(val: Any) -> int | None:
    """税抜単価（円）を整数にする。不明・null・空なら None（既存入力を上書きしない）。"""
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ("null", "none", "-", "不明"):
        return None
    s = re.sub(r"[,\s円￥¥]", "", s, flags=re.UNICODE)
    s = re.sub(r"(?i)yen", "", s)
    if not s or not re.search(r"\d", s):
        return None
    try:
        n = int(round(float(s)))
        return max(1, n)
    except Exception:
        return None


def _apply_gemini_json_to_session(
    result: dict[str, Any],
    df_ledger: pd.DataFrame | None = None,
) -> None:
    """Gemini の JSON をフォーム用 session_state に反映する（英日キー両対応）。"""
    r = result
    st.session_state.pop("_gemini_match_management_id", None)
    _qty_g = _coerce_positive_int(
        r.get("quantity") or r.get("数量") or r.get("qty") or 1,
        default=1,
    )
    st.session_state.field_qty = _qty_g
    st.session_state[REGISTRATION_QTY_WIDGET_KEY] = _qty_g

    m = r.get("match")
    row_hit: pd.Series | None = None
    if isinstance(m, dict) and df_ledger is not None and not df_ledger.empty:
        mid_hit = str(m.get("management_id") or m.get("管理ID") or "").strip()
        if mid_hit and COL_MANAGEMENT_ID in df_ledger.columns:
            mask_id = df_ledger[COL_MANAGEMENT_ID].astype(str).str.strip() == mid_hit
            mask_in = _mask_ledger_in_stock(df_ledger)
            sub_hit = df_ledger.loc[mask_id & mask_in]
            if len(sub_hit) == 1:
                row_hit = sub_hit.iloc[0]
                st.session_state["_gemini_match_management_id"] = mid_hit

    pn = str(
        r.get("product_name")
        or r.get("商品名")
        or r.get("name")
        or ""
    ).strip()
    su = str(
        r.get("supplier")
        or r.get("仕入先・取引先")
        or r.get("仕入先")
        or r.get("取引先")
        or r.get("vendor")
        or ""
    ).strip()
    if row_hit is not None:
        rpn = str(row_hit.get(COL_NAME, "") or "").strip()
        rsu = str(row_hit.get(COL_SUPPLIER, "") or "").strip()
        if rpn:
            pn = rpn
        if rsu:
            su = rsu
    match_conf_ok = isinstance(m, dict) and float(m.get("confidence") or 0) >= 0.75
    if match_conf_ok:
        if not pn:
            pn = str(m.get("product_name") or "").strip()
        if not su:
            su = str(m.get("supplier") or "").strip()
    if pn:
        st.session_state.field_product_name = pn
    if su:
        st.session_state.field_supplier = su

    line_yen: int | None = None
    if row_hit is not None:
        ly0 = _finite_int(row_hit.get(COL_PRICE_EXCL), 0)
        if ly0 > 0:
            line_yen = ly0
    if line_yen is None:
        line_yen = _coerce_unit_price_yen(
            r.get("line_price_excl")
            or r.get("line_excl_yen")
            or r.get("仕入金額（税抜）")
            or r.get("unit_price_excl")
            or r.get("unit_price")
            or r.get("product_unit_price_excl")
            or r.get("単価")
            or r.get("単価（税抜）")
        )
    if line_yen is None and match_conf_ok and isinstance(m, dict):
        line_yen = _coerce_unit_price_yen(
            m.get("line_price_excl")
            or m.get("unit_price_excl")
            or m.get("unit_price")
        )
    if line_yen is not None:
        st.session_state.field_line_excl_yen = line_yen

    kind = str(
        r.get("product_kind")
        or r.get("種類")
        or r.get("type")
        or r.get("商品カテゴリ")
        or ""
    ).strip()
    if not kind and pn:
        kind = pn
    st.session_state.ai_kind = kind

    vf = r.get("visual_features")
    if not vf:
        parts = [
            r.get(k)
            for k in (
                "色",
                "柄",
                "素材",
                "状態",
                "色柄",
                "備考",
                "color",
                "pattern",
                "material",
                "condition",
            )
            if r.get(k)
        ]
        vf = " / ".join(str(p) for p in parts)
    st.session_state.ai_features = str(vf or "")
    st.session_state.ai_parse_ran = True


def _apply_gemini_sale_link_to_session(
    result: dict[str, Any],
    df_ledger: pd.DataFrame | None,
    *,
    fill_product_preview_fields: bool = True,
) -> None:
    """販売元管理ID（入庫時の管理ID）の写真照合結果を session_state に反映する。"""
    st.session_state.pop("_sale_link_management_id", None)
    st.session_state.pop("_sale_link_warn", None)
    r = result
    m = r.get("match")
    if not isinstance(m, dict):
        m = {}
    mid = str(
        m.get("management_id")
        or m.get("管理ID")
        or r.get("management_id")
        or ""
    ).strip()
    conf = float(m.get("confidence") or r.get("confidence") or 0)
    row_hit: pd.Series | None = None
    if (
        mid
        and df_ledger is not None
        and not df_ledger.empty
        and COL_MANAGEMENT_ID in df_ledger.columns
    ):
        mask_mid = df_ledger[COL_MANAGEMENT_ID].astype(str).str.strip() == mid
        hits_in = df_ledger.loc[mask_mid & _mask_ledger_in_stock(df_ledger)]
        if len(hits_in) == 1:
            row_hit = hits_in.iloc[0]
    if not mid:
        return
    if row_hit is None:
        if (
            df_ledger is not None
            and not df_ledger.empty
            and COL_MANAGEMENT_ID in df_ledger.columns
        ):
            hits_any = df_ledger.loc[
                df_ledger[COL_MANAGEMENT_ID].astype(str).str.strip() == mid
            ]
            if len(hits_any) == 1:
                stt_bad = _normalize_stock_status(
                    str(hits_any.iloc[0].get(COL_STOCK_STATUS, ""))
                )
                st.session_state["_sale_link_warn"] = (
                    f"管理ID {mid} は「{stt_bad}」のため販売元に使えません。"
                    "照合対象は **在庫中** の行のみです。"
                )
                return
        st.session_state.field_sale_source_mgmt_id = mid
        st.session_state["_sale_link_warn"] = (
            f"管理ID {mid} が台帳に見つかりません。表記を確認するか手入力してください。"
        )
        return
    st.session_state.field_sale_source_mgmt_id = mid
    st.session_state["_sale_link_management_id"] = mid
    if fill_product_preview_fields and conf >= 0.72:
        rpn = str(row_hit.get(COL_NAME, "") or "").strip()
        rsu = str(row_hit.get(COL_SUPPLIER, "") or "").strip()
        if rpn:
            st.session_state.field_product_name = rpn
        if rsu:
            st.session_state.field_supplier = rsu
        ly = _finite_int(row_hit.get(COL_PRICE_EXCL), 0)
        if ly > 0:
            st.session_state.field_line_excl_yen = ly


def _gemini_input_image_from_upload(uploaded) -> Image.Image:
    """解析直前に ``prepare_upload_image_jpeg`` と同じ圧縮・リサイズを適用した PIL 画像を返す。"""
    jpeg_bytes, _ = prepare_upload_image_jpeg(uploaded.getvalue())
    return Image.open(io.BytesIO(jpeg_bytes)).convert("RGB")


def _consumption_tax_rate_from_choice_label(label: str) -> float:
    return CONSUMPTION_TAX_CHOICE_TO_RATE.get(label, CONSUMPTION_TAX_RATE)


def _finite_int(val: Any, default: int = 0) -> int:
    """NaN / inf / 非数値を default に落とし、有限な int にする（金額・数量用）。"""
    try:
        if val is None:
            return default
        if isinstance(val, (float, np.floating)):
            if not math.isfinite(float(val)) or pd.isna(val):
                return default
        if isinstance(val, str):
            t = val.strip()
            if not t or t.lower() in ("nan", "none", "<na>", "nat"):
                return default
            t = (
                t.replace(",", "")
                .replace("，", "")
                .replace("¥", "")
                .replace("\u00a5", "")
                .strip()
            )
            if not t:
                return default
            val = t
        n = float(pd.to_numeric(val, errors="coerce"))
        if not math.isfinite(n) or pd.isna(n):
            return default
        return int(round(n))
    except (OverflowError, ValueError, TypeError):
        return default


def _series_to_numeric_loose(s: pd.Series) -> pd.Series:
    """カンマ区切り・円記号付きのセルでも数値化する（スプレッドシート取り込み用）。"""
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    if s.dtype == object or pd.api.types.is_string_dtype(s):
        st = s.where(pd.notna(s), "").astype(str)
        st = st.str.replace(",", "", regex=False)
        st = st.str.replace("，", "", regex=False)
        st = st.str.replace("¥", "", regex=False)
        st = st.str.replace("\u00a5", "", regex=False)
        st = st.str.strip()
        st = st.replace({"nan": "", "None": "", "<NA>": ""})
        return pd.to_numeric(st, errors="coerce")
    return pd.to_numeric(s, errors="coerce")


def price_incl_tax(price_excl_yen: int, tax_rate: float | None = None) -> int:
    """税抜き行金額から税込円金額（四捨五入）。

    Args:
        price_excl_yen: 税抜き金額（円）
        tax_rate: 消費税率（例: 0.1）。None のときは :data:`CONSUMPTION_TAX_RATE`（10%）
    """
    ex = _finite_int(price_excl_yen, 0)
    r_raw = CONSUMPTION_TAX_RATE if tax_rate is None else tax_rate
    r = float(pd.to_numeric(r_raw, errors="coerce"))
    if not math.isfinite(r) or pd.isna(r):
        r = float(CONSUMPTION_TAX_RATE)
    return int(round(ex * (1 + r)))


def _infer_tax_rate_from_main_line(line_excl_yen: int, line_incl_yen: int) -> float:
    """仕入金額（税抜）と仕入金額（税込）から、登録時と同じ消費税区分を推定する。"""
    excl = _finite_int(line_excl_yen, 0)
    incl = _finite_int(line_incl_yen, 0)
    if excl <= 0:
        return float(CONSUMPTION_TAX_RATE)
    if incl <= excl:
        return 0.0
    for _label, rate in CONSUMPTION_TAX_CHOICE_TO_RATE.items():
        if price_incl_tax(excl, float(rate)) == incl:
            return float(rate)
    return float(CONSUMPTION_TAX_RATE)


def _estimate_excl_yen_from_incl_yen(incl_yen: int) -> int:
    """税込行金額から税抜行金額を推定（登録時の税率候補に税込が一致する組を採用）。"""
    incl_i = _finite_int(incl_yen, 0)
    if incl_i <= 0:
        return 0
    seen: set[float] = set()
    for rate in CONSUMPTION_TAX_CHOICE_TO_RATE.values():
        rr = float(rate)
        if rr in seen:
            continue
        seen.add(rr)
        if rr == 0.0:
            return incl_i
        ex = int(round(incl_i / (1 + rr)))
        if price_incl_tax(ex, rr) == incl_i:
            return ex
    return int(round(incl_i / (1 + float(CONSUMPTION_TAX_RATE))))


def _planned_actual_line_amounts(
    qty: int,
    planned_unit_excl: int,
    actual_unit_excl: int,
    status: str,
    tax_rate: float,
) -> tuple[int, int, int, int]:
    """販売予定・実売の税抜行計と税込行計（税抜列×数量を行合計にしてから税込）。"""
    q = max(1, _finite_int(qty, 1))
    pu, au = _finite_int(planned_unit_excl, 0), _finite_int(actual_unit_excl, 0)
    st = _normalize_stock_status(status)
    tr = float(pd.to_numeric(tax_rate, errors="coerce"))
    if not math.isfinite(tr) or pd.isna(tr):
        tr = float(CONSUMPTION_TAX_RATE)
    plex = pu * q if pu > 0 else 0
    pincl = price_incl_tax(plex, tr) if plex > 0 else 0
    aex = (au * q) if (st == STATUS_SOLD and au > 0) else 0
    aincl = price_incl_tax(aex, tr) if aex > 0 else 0
    return plex, pincl, aex, aincl


def analyze_image_with_gemini(
    image_data,
    *,
    inventory_context: str | None = None,
    prompt_mode: str = "full",
) -> str:
    api_key = _secret_str(SECRET_GEMINI_API_KEY)
    if not api_key:
        raise RuntimeError(
            f"{SECRET_GEMINI_API_KEY} が設定されていません。`.streamlit/secrets.toml` を確認してください。"
        )
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(_gemini_model_name())
    inv_block = ""
    if inventory_context and inventory_context.strip():
        inv_block = f"""
次のリストは、すでに台帳にある **在庫中** の行の抜粋です（**販売済は含みません**。最大約90件。写真と同一・類似の商品がありそうなら必ず照合してください）。
{inventory_context.strip()}

照合するときは必ず "match" オブジェクトを付け、少なくとも次を含めてください:
- "management_id" (string): 上のリストにある **在庫中** 行の **管理ID** と完全一致する値（該当がなければ ""）
- "product_name" (string): 台帳の商品名に合わせた確定案（推測でも可）
- "supplier" (string): 台帳の仕入先に合わせた確定案（推測でも可）
- "line_price_excl" (integer or null): 台帳の仕入金額（税抜）と一致する整数。不明なら null
- "confidence" (number): 0.0〜1.0 で、写真と台帳行が同一在庫である確信度

同一行が見つからない場合は management_id を "" にし、confidence は 0.4 未満にしてください。
"""
    if prompt_mode == "stocktake_match":
        if not inventory_context or not inventory_context.strip():
            raise ValueError(
                "棚卸しの照合には台帳に在庫中の行が必要です（スプレッドシートを確認してください）。"
            )
        prompt = f"""この写真は **店舗で棚卸しのために撮影した現物1点** です（呉服・和装の在庫）。
次のリストは台帳の **在庫中** の行だけです（販売済は含みません）。
写真と **同一の在庫1行** を特定し、JSON だけを返してください（説明文・Markdown のコードフェンス禁止）。

{inventory_context.strip()}

返却形式（キーは次のみ）:
- "match" (object): 必須。フィールド:
  - "management_id" (string): 選んだ行の管理ID（G########）。該当なしなら ""
  - "confidence" (number): 0.0〜1.0
  - "product_name" (string): その行の商品名（参考）
  - "supplier" (string): その行の仕入先（参考）

該当がなければ management_id を ""、confidence は 0.35 以下にしてください。"""
        response = model.generate_content([prompt, image_data])
        return response.text or ""

    if prompt_mode == "sale_link":
        if not inventory_context or not inventory_context.strip():
            raise ValueError(
                "販売元の写真照合には、台帳に在庫中の行が必要です（スプレッドシートを確認してください）。"
            )
        prompt = f"""この画像は、**販売時にどの在庫行に対応するか** を特定するための商品写真です（呉服店の在庫）。
次のリストは台帳の **在庫中** の行だけです（**販売済の行は含めていません**）。必ずこのリストの中からだけ management_id を選べ。
写真と **同一の在庫1行** を選び、JSON だけを返してください（説明文・コードフェンス禁止）。

{inventory_context.strip()}

返却形式（キーは次のみ）:
- "match" (object): 必須。フィールド:
  - "management_id" (string): 選んだ行の管理ID（G########）。該当なしなら ""
  - "confidence" (number): 0.0〜1.0
  - "product_name" (string): その行の商品名（参考）
  - "supplier" (string): その行の仕入先（参考）
  - "line_price_excl" (integer or null): その行の仕入金額（税抜）

該当がなければ management_id を ""、confidence は 0.25 以下にする。"""
        response = model.generate_content([prompt, image_data])
        return response.text or ""

    prompt = f"""この画像は呉服店の在庫・売買用の商品写真です。次のキーだけを持つ JSON オブジェクトを 1 つだけ返してください。
説明文や Markdown のコードフェンスは付けず、JSON のみを出力してください。

必須キー（値の型を守ること）:
- "product_name" (string): 商品名として適切な短い名称。不明なら ""
- "supplier" (string): 仕入先・取引先として推測できる名称。不明なら ""
- "quantity" (integer): 写っている点数・束の本数などの推定。最低 1
- "product_kind" (string): 種類の推定（例: 振袖、訪問着、帯、長襦袢）。不明なら ""
- "color" (string): 色の推定。不明なら ""
- "pattern" (string): 柄の推定。不明なら ""
- "material" (string): 素材の推定。不明なら ""
- "condition" (string): 状態の推定。不明なら ""
- "unit_price_excl" (integer or null): 1点あたりの税抜の仕入金額（円）の推定。相場・品質から読めない場合は null（勝手に 1 にしない）
{inv_block}
任意: 台帳照合結果を "match" にまとめる（上記リストがあるときはできる限り付与）
  例: {{"management_id": "G00000001", "product_name": "…", "supplier": "…", "line_price_excl": 12345, "confidence": 0.85}}
  不要・該当なしのときは "match" キー自体を省略してもよい。"""
    response = model.generate_content([prompt, image_data])
    return response.text or ""


def _voucher_configure_model():
    api_key = _secret_str(SECRET_GEMINI_API_KEY)
    if not api_key:
        raise RuntimeError(
            f"{SECRET_GEMINI_API_KEY} が設定されていません。`.streamlit/secrets.toml` を確認してください。"
        )
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(_gemini_voucher_model_name())


def analyze_voucher_document_with_gemini_images(images: list[Any]) -> str:
    """証憑のページ画像（1枚以上）を Gemini に渡し、supplier_name / purchase_date / items の JSON 文字列を返す。"""
    if not images:
        raise ValueError("画像がありません。")
    model = _voucher_configure_model()
    n = len(images)
    multi = (
        "複数のページ画像が順に渡されています。全体を通して一つの証憑として読み取ってください。"
        if n > 1
        else ""
    )
    prompt = f"""内部システムプロンプト:
あなたはプロの会計士です。提供された呉服店向け証憑（納品書・請求書・領収書）のページ画像を解析し、在庫管理に必要なデータを抽出してください。
{multi}

{VOUCHER_EXTRACTION_RULES}

{VOUCHER_JSON_SPEC}"""
    response = model.generate_content([prompt, *images])
    return response.text or ""


def analyze_voucher_document_with_gemini(image_data) -> str:
    """単一の証憑画像を Gemini に渡す（後方互換・内部は複数画像 API と共通）。"""
    return analyze_voucher_document_with_gemini_images([image_data])


def analyze_voucher_document_with_gemini_text(
    *, source_instruction: str, document_body: str
) -> str:
    """テキスト（PDF抽出・Markdown表・Word全文など）から同一 JSON 形式で抽出する。"""
    model = _voucher_configure_model()
    body = (document_body or "").strip()
    max_chars = 600_000
    if len(body) > max_chars:
        body = body[:max_chars] + "\n\n（長文のため先頭のみ。末尾を省略しました）"
    prompt = f"""内部システムプロンプト:
あなたはプロの会計士です。呉服店向け証憑（納品書・請求書・領収書）に関する次の入力から、在庫管理に必要なデータを抽出してください。

{source_instruction}

{VOUCHER_EXTRACTION_RULES}

{VOUCHER_JSON_SPEC}

--- 入力ここから ---
{body}
--- 入力ここまで ---"""
    response = model.generate_content(prompt)
    return response.text or ""


def _voucher_upload_suffix(filename: str) -> str:
    parts = (filename or "").rsplit(".", 1)
    return parts[-1].lower() if len(parts) == 2 else ""


def prepare_voucher_file_for_drive_storage(
    raw: bytes, original_filename: str
) -> tuple[bytes, str, str]:
    """証憑を Drive 保存用に整形する。JPG/PNG 等は長辺最大2000px・JPEG quality=75。PDF/xlsx/docx は原本のまま。

    Returns:
        (保存バイナリ, mime_type, 拡張子に使う文字列 ``.jpg`` 等)
    """
    suf = _voucher_upload_suffix(original_filename)
    if suf in ("jpg", "jpeg", "png", "webp"):
        img = Image.open(io.BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.getchannel("A"))
        img = bg
        img = _resize_long_edge_max(img, VOUCHER_DRIVE_JPEG_MAX_LONG_EDGE)
        buf = io.BytesIO()
        img.save(
            buf,
            format="JPEG",
            quality=VOUCHER_DRIVE_JPEG_QUALITY,
            optimize=True,
            progressive=True,
        )
        return buf.getvalue(), "image/jpeg", ".jpg"
    mime_map = {
        "pdf": "application/pdf",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    mime = mime_map.get(suf, "application/octet-stream")
    ext = f".{suf}" if suf else ""
    return raw, mime, ext


def _voucher_drive_safe_filename(
    purchase_date: str, supplier: str, original_filename: str, ext_with_dot: str
) -> str:
    """``[仕入日]_[仕入先]_[元ファイル名].拡子`` 形式（ファイル名向けに禁則文字を除去）。"""
    d = (purchase_date or "").strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
        d = jst_now().strftime("%Y-%m-%d")
    d_part = d.replace("-", "")
    sup = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (supplier or "").strip())[:50].strip(
        "_"
    ) or "unknown"
    base = Path(str(original_filename)).name
    stem = Path(base).stem
    safe_stem = re.sub(
        r'[^\w\-_.\u3040-\u30ff\u4e00-\u9fff]', "_", stem
    ).strip("._")[:100] or "voucher"
    ext = (
        ext_with_dot
        if ext_with_dot.startswith(".")
        else (f".{ext_with_dot}" if ext_with_dot else "")
    )
    return f"{d_part}_{sup}_{safe_stem}{ext}"


def _gas_evidence_upload_ready() -> bool:
    """証憑を Google ドライブ（GAS 経由）に送るのに必要な secrets が揃っているか。"""
    return bool(
        _secret_str(SECRET_GAS_UPLOAD_URL)
        and _secret_str(SECRET_GAS_API_KEY)
        and _secret_str(SECRET_GOOGLE_DRIVE_FOLDER_ID)
    )


def _voucher_extract_pdf_text(pdf_bytes: bytes) -> str:
    """pypdf で PDF からプレーンテキストを抽出する。"""
    from pypdf import PdfReader

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception as e:
        raise ValueError(f"PDF を開けませんでした: {e}") from e
    if getattr(reader, "is_encrypted", False):
        try:
            auth = reader.decrypt("")
        except Exception as e:
            raise ValueError("パスワード付きPDFには未対応です。") from e
        if auth == 0:
            raise ValueError("パスワード付きPDFには未対応です。")
    chunks: list[str] = []
    for page in reader.pages:
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        chunks.append(t)
    return "\n".join(chunks).strip()


def _voucher_pdf_pages_to_images(
    pdf_bytes: bytes, *, max_pages: int = 15, zoom: float = 2.0
) -> list[Image.Image]:
    """スキャン PDF 等をページ画像にレンダリングする（pymupdf）。"""
    import fitz

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        n = min(doc.page_count, max_pages)
        if n <= 0:
            raise ValueError("PDF にページがありません。")
        out: list[Image.Image] = []
        for i in range(n):
            page = doc.load_page(i)
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out.append(
                Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            )
        return out
    finally:
        doc.close()


def _dataframe_to_markdown_table(df: pd.DataFrame, *, max_rows: int = 500) -> str:
    """pandas DataFrame を簡易 Markdown 表にする（外部パッケージ不要）。"""
    df2 = df.fillna("").head(max_rows).copy()
    cols = [str(c).replace("|", "/") for c in df2.columns]
    header = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join("---" for _ in cols) + " |"
    lines = [header, sep]
    for _, row in df2.iterrows():
        cells = [
            str(v).replace("|", "\\|").replace("\n", " ").strip()
            for v in row.tolist()
        ]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _voucher_excel_bytes_to_markdown(xlsx_bytes: bytes) -> str:
    """Excel を全シート読み、Markdown 形式のテキストにまとめる。"""
    bio = io.BytesIO(xlsx_bytes)
    try:
        xl = pd.ExcelFile(bio, engine="openpyxl")
    except Exception as e:
        raise ValueError(
            f"Excel（.xlsx）の読み込みに失敗しました。openpyxl 対応形式か確認してください: {e}"
        ) from e
    parts: list[str] = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=0, engine="openpyxl")
        if df.empty:
            parts.append(f"## シート: {sheet}\n\n（データ行なし）\n")
        else:
            parts.append(
                f"## シート: {sheet}\n\n{_dataframe_to_markdown_table(df)}\n"
            )
    text = "\n".join(parts).strip()
    if not text:
        raise ValueError("Excel に読み取れるシートがありません。")
    return text


def _voucher_docx_bytes_to_text(docx_bytes: bytes) -> str:
    """Word（.docx）から段落および表セルのテキストを抽出する。"""
    from docx import Document

    doc = Document(io.BytesIO(docx_bytes))
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    table_lines: list[str] = []
    for table in doc.tables:
        for row in table.rows:
            cells = [
                cell.text.strip().replace("\n", " ") for cell in row.cells
            ]
            table_lines.append("\t".join(cells))
    blocks = paras + table_lines
    return "\n".join(blocks).strip()


class _VoucherRawBytesUpload:
    """``_gemini_input_image_from_upload`` 用の最小インターフェース。"""

    __slots__ = ("_data",)

    def __init__(self, data: bytes) -> None:
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def analyze_voucher_upload_bytes(raw: bytes, filename: str) -> str:
    """アップロード種別に応じて分岐し、Gemini から JSON 文字列を返す。"""
    suf = _voucher_upload_suffix(filename)
    if suf in ("jpg", "jpeg", "png", "webp"):
        img = _gemini_input_image_from_upload(_VoucherRawBytesUpload(raw))
        return analyze_voucher_document_with_gemini(img)
    if suf == "pdf":
        text = _voucher_extract_pdf_text(raw)
        if text:
            return analyze_voucher_document_with_gemini_text(
                source_instruction=(
                    "以下は PDF からテキスト抽出した内容です。"
                    "この請求・納品・領収に相当する情報から在庫データを抽出してください。"
                ),
                document_body=text,
            )
        imgs = _voucher_pdf_pages_to_images(raw)
        return analyze_voucher_document_with_gemini_images(imgs)
    if suf == "xlsx":
        md = _voucher_excel_bytes_to_markdown(raw)
        return analyze_voucher_document_with_gemini_text(
            source_instruction=(
                "以下は Excel（.xlsx）を Markdown 表に変換したものです。"
                "この表から仕入・納品に相当する在庫情報を抽出してください。"
            ),
            document_body=md,
        )
    if suf == "docx":
        t = _voucher_docx_bytes_to_text(raw)
        if not t:
            raise ValueError("Word 文書からテキストを抽出できませんでした。")
        return analyze_voucher_document_with_gemini_text(
            source_instruction=(
                "以下は Word（.docx）から抽出した全文です。"
                "この文書内の請求・納品・領収に相当する情報から在庫データを抽出してください。"
            ),
            document_body=t,
        )
    raise ValueError(
        f"未対応の拡張子です（.jpg / .png / .pdf / .xlsx / .docx のみ）: .{suf}"
    )


def _merge_voucher_items_for_preview(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """同一商品名（大文字小文字・前後空白無視）の数量を合算し、後から現れる単価で上書き（手書き修正の優先に近づける）。"""
    order: dict[str, dict[str, Any]] = {}
    key_order: list[str] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        name = str(it.get("name") or "").strip()
        if not name:
            continue
        key = name.casefold()
        q = _coerce_positive_int(it.get("quantity"), 1)
        up = _coerce_unit_price_yen(it.get("unit_price"))
        if up is None:
            up = _coerce_unit_price_yen(it.get("unit_price_excl"))
        cat = str(it.get("category") or "").strip()
        if key not in order:
            order[key] = {"商品名": name, "数量": q, "単価（税抜）": up, "カテゴリ": cat}
            key_order.append(key)
        else:
            row = order[key]
            row["数量"] = int(row["数量"]) + q
            if up is not None:
                row["単価（税抜）"] = int(up)
            if cat:
                row["カテゴリ"] = cat
    out: list[dict[str, Any]] = []
    for k in key_order:
        r = order[k]
        up = r["単価（税抜）"]
        if up is None:
            up = 1
        out.append(
            {
                "商品名": r["商品名"],
                "数量": max(1, int(r["数量"])),
                "単価（税抜）": max(1, int(up)),
                "カテゴリ": str(r.get("カテゴリ") or ""),
            }
        )
    return out


def _voucher_record_datetime_jst(purchase_date_str: str) -> str:
    """証憑の仕入日を台帳の日時列用に整形。YYYY-MM-DD のときはその日の JST 正午、それ以外は現在 JST。"""
    s = (purchase_date_str or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return f"{s} 12:00:00"
    return jst_now_str()


def _init_voucher_sidebar_state() -> None:
    if "voucher_supplier_edit" not in st.session_state:
        st.session_state.voucher_supplier_edit = ""
    if "voucher_date_edit" not in st.session_state:
        st.session_state.voucher_date_edit = ""
    if "voucher_consumption_tax_choice" not in st.session_state:
        st.session_state.voucher_consumption_tax_choice = "10%"


def _confirm_voucher_import(
    df: pd.DataFrame | None,
    supplier: str,
    purchase_date_str: str,
    tax_choice_label: str,
) -> None:
    """証憑プレビュー表の内容を 1点1行で台帳に追記する。"""
    if df is None or df.empty:
        st.error("反映する行がありません。")
        return
    rows_spec: list[tuple[str, int, int, str]] = []
    for _, row in df.iterrows():
        name = str(row.get("商品名", "") or "").strip()
        if not name:
            continue
        q = max(1, _finite_int(row.get("数量"), 1))
        up = max(1, _finite_int(row.get("単価（税抜）"), 1))
        cat = str(row.get("カテゴリ", "") or "").strip()
        rows_spec.append((name, q, up, cat))
    if not rows_spec:
        st.error("商品名が入っている行がありません。")
        return
    total_q = sum(q for _, q, _, _ in rows_spec)
    ws: Any = None
    if not _uses_local_inventory_csv():
        ws = ensure_worksheet_header()
        if ws is None:
            st.warning("スプレッドシート未設定のため保存できません。")
            return
    tax_r = _consumption_tax_rate_from_choice_label(str(tax_choice_label))
    rec_dt = _voucher_record_datetime_jst(purchase_date_str)
    sup = (supplier or "").strip()
    recorded_at = jst_now_str()
    evidence_url = ""
    stash_b = st.session_state.get("voucher_stash_bytes")
    stash_name = str(st.session_state.get("voucher_stash_name") or "voucher")
    if stash_b is not None:
        if _gas_evidence_upload_ready():
            try:
                blob, mime, ext_dot = prepare_voucher_file_for_drive_storage(
                    bytes(stash_b), stash_name
                )
                fname = _voucher_drive_safe_filename(
                    recorded_at[:10], sup, stash_name, ext_dot
                )
                evidence_url = upload_image_to_drive(fname, mime, blob)
                fb = _fallback_image_url_when_gas_unconfigured()
                if not (evidence_url or "").strip() or (evidence_url or "").strip() == (
                    fb or ""
                ).strip():
                    evidence_url = ""
                    st.warning(
                        "ドライブの表示 URL を取得できませんでした。台帳には証憑URLなしで記録します。"
                    )
            except Exception as e:
                st.error(f"証憑の Google ドライブ保存に失敗しました: {e}")
                return
        else:
            st.warning(
                "GAS_UPLOAD_URL / GAS_API_KEY / GOOGLE_DRIVE_FOLDER_ID が未設定のため、"
                "証憑ファイルは Drive に保存されません（台帳のみ反映）。"
            )
    try:
        ids = allocate_management_ids(ws, total_q)
    except Exception as e:
        st.error(f"管理IDの採番に失敗しました: {e}")
        return
    idx = 0
    try:
        with st.spinner("台帳に書き込んでいます…"):
            _v_dt = jst_now_str()
            batch_rows: list[list[Any]] = []
            for name, q, up, cat in rows_spec:
                memo = f"証憑取込 category={cat}" if cat else "証憑取込"
                incl = price_incl_tax(up, tax_r)
                for _ in range(q):
                    batch_rows.append(
                        _inventory_row_values_for_append(
                            _v_dt,
                            rec_dt,
                            "入庫（購入）",
                            name,
                            sup,
                            up,
                            incl,
                            "",
                            ids[idx],
                            memo,
                            consumption_tax_rate=tax_r,
                            voucher_recorded_at=recorded_at,
                            voucher_evidence_url=evidence_url,
                        )
                    )
                    idx += 1
            _append_inventory_data_rows(batch_rows)
    except Exception as e:
        st.error(f"台帳の更新に失敗しました: {e}")
        return
    st.session_state.pop("voucher_preview_df", None)
    st.session_state.pop(VOUCHER_DATA_EDITOR_KEY, None)
    st.session_state.pop("voucher_stash_bytes", None)
    st.session_state.pop("voucher_stash_name", None)
    st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
    _ev = (evidence_url or "").strip()
    _msg = f"証憑取込を記録しました（{total_q} 行・1点1行）。証憑記録日時: {recorded_at}"
    if _ev:
        _msg += " 証憑を Drive に保存済みです。"
    st.session_state["_voucher_import_flash"] = _msg
    st.rerun()


def _render_voucher_inventory_panel() -> None:
    """登録ページ内: 証憑ファイルのアップロード・解析・プレビュー編集・確定反映。"""
    _vflash = st.session_state.pop("_voucher_import_flash", None)
    if _vflash:
        st.success(_vflash)
    st.markdown("### 証憑から在庫反映")
    st.caption(
        "納品書・請求書・領収書を画像・PDF・Excel・Word から読み取り、入庫（購入）として台帳に反映します。"
        "PDF はテキスト優先、空ならページ画像として解析します。解析後は表で修正してから確定してください。"
    )
    if not _secret_str(SECRET_GEMINI_API_KEY):
        st.info(f"{SECRET_GEMINI_API_KEY} が未設定のため使えません。")
        return
    if not _uses_local_inventory_csv() and not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        st.info(
            "在庫の保存先が未設定です。`GOOGLE_SPREADSHEET_ID` を設定するか、"
            "`INVENTORY_SOURCE = \"csv\"` で共有の inventory.csv を使ってください。"
        )
        return

    voucher_up = st.file_uploader(
        "証憑ファイル（画像 / PDF / Excel / Word）",
        type=["jpg", "jpeg", "png", "pdf", "xlsx", "docx"],
        key="voucher_file_uploader",
    )
    if st.button(
        "証憑を解析",
        key="voucher_analyze_btn",
        disabled=voucher_up is None,
    ):
        with st.spinner("証憑を解析しています…"):
            try:
                raw = analyze_voucher_upload_bytes(
                    voucher_up.getvalue(), voucher_up.name
                )
                d = _parse_json_from_model(raw)
                items = d.get("items")
                if not isinstance(items, list) or not items:
                    st.error(
                        "items が見つかりませんでした。ファイル内容を確認してください。"
                    )
                else:
                    merged = _merge_voucher_items_for_preview(
                        [x for x in items if isinstance(x, dict)]
                    )
                    if not merged:
                        st.error("有効な商品行がありません。")
                    else:
                        st.session_state.voucher_supplier_edit = str(
                            d.get("supplier_name") or ""
                        ).strip()
                        st.session_state.voucher_date_edit = str(
                            d.get("purchase_date") or ""
                        ).strip()
                        st.session_state.voucher_preview_df = pd.DataFrame(merged)
                        st.session_state.pop(VOUCHER_DATA_EDITOR_KEY, None)
                        st.session_state["voucher_stash_bytes"] = voucher_up.getvalue()
                        st.session_state["voucher_stash_name"] = voucher_up.name
                        st.success(
                            f"解析しました（{len(merged)} 商品行）。内容を確認して確定してください。"
                        )
            except Exception as e:
                st.warning(str(e))

    st.text_input(
        "仕入先（証憑・上書き可）",
        key="voucher_supplier_edit",
        placeholder="仕入先・取引先",
    )
    st.text_input(
        "仕入日（YYYY-MM-DD・上書き可）",
        key="voucher_date_edit",
        placeholder="例: 2026-04-15",
    )
    st.radio(
        "証憑取込の消費税（税込計算）",
        options=list(CONSUMPTION_TAX_CHOICE_TO_RATE.keys()),
        horizontal=True,
        key="voucher_consumption_tax_choice",
    )

    base_df = st.session_state.get("voucher_preview_df")
    if base_df is not None and not base_df.empty:
        edited_df = st.data_editor(
            base_df,
            num_rows="dynamic",
            key=VOUCHER_DATA_EDITOR_KEY,
            use_container_width=True,
        )
        if st.button(
            "この内容で台帳に確定反映",
            type="primary",
            key="voucher_confirm_btn",
        ):
            _confirm_voucher_import(
                edited_df,
                str(st.session_state.get("voucher_supplier_edit", "") or ""),
                str(st.session_state.get("voucher_date_edit", "") or ""),
                str(
                    st.session_state.get("voucher_consumption_tax_choice", "10%")
                    or "10%"
                ),
            )
        if st.button("プレビューをクリア", key="voucher_clear_preview_btn"):
            st.session_state.pop("voucher_preview_df", None)
            st.session_state.pop(VOUCHER_DATA_EDITOR_KEY, None)
            st.session_state.pop("voucher_stash_bytes", None)
            st.session_state.pop("voucher_stash_name", None)
            st.rerun()


def _open_inventory_workbook():
    sid = _secret_str(SECRET_GOOGLE_SPREADSHEET_ID)
    if not sid:
        return None
    try:
        return _gspread_client().open_by_key(str(sid))
    except Exception:
        return None


def _get_or_create_inventory_worksheet():
    """在庫ワークシートを開く。無ければ十分な行・列で作成する。失敗時は None。"""
    sh = _open_inventory_workbook()
    if sh is None:
        return None
    wname = _secret_str(SECRET_GOOGLE_WORKSHEET_NAME, DEFAULT_WORKSHEET_NAME)
    try:
        try:
            return sh.worksheet(str(wname))
        except gspread.WorksheetNotFound:
            return sh.add_worksheet(
                title=str(wname),
                rows=2000,
                cols=max(20, len(EXPECTED_HEADERS) + 2),
            )
    except Exception:
        return None


def _bump_inventory_sheet_cache_bust() -> None:
    """スプレッドシート由来の get_all_values キャッシュを無効化する（追記・全置換・手動再読込後）。"""
    try:
        st.session_state[SESSION_KEY_INV_SHEET_CACHE_BUST] = int(
            st.session_state.get(SESSION_KEY_INV_SHEET_CACHE_BUST, 0)
        ) + 1
    except Exception:
        pass


@st.cache_data(show_spinner=False)
def _inventory_sheet_get_all_values_cached(
    sheet_id: str, worksheet_title: str, bust: int
) -> list[list[str]] | None:
    """同一 bust の間は get_all_values の結果を再利用する（bust は書き込み・再読込で進める）。"""
    _ = bust
    try:
        sh = _gspread_client().open_by_key(str(sheet_id))
    except Exception:
        return None
    try:
        try:
            ws = sh.worksheet(str(worksheet_title))
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(
                title=str(worksheet_title),
                rows=2000,
                cols=max(20, len(EXPECTED_HEADERS) + 2),
            )
    except Exception:
        return None
    try:
        return ws.get_all_values()
    except Exception:
        return None


def ensure_worksheet_header():
    """1行目がヘッダーでなければ作成（初回のみ想定）。secrets 未設定時は None。"""
    ws = _get_or_create_inventory_worksheet()
    if ws is None:
        return None
    try:
        first = ws.row_values(1)
        if not first or first[: len(EXPECTED_HEADERS)] != EXPECTED_HEADERS:
            ws.update("A1", [EXPECTED_HEADERS], value_input_option="USER_ENTERED")
            try:
                _apply_inventory_amount_number_formats(ws)
            except Exception:
                pass
        return ws
    except Exception:
        return None


def upload_image_to_drive(filename: str, mime: str, data: bytes) -> str:
    """Google Apps Script 経由で Google ドライブに保存し、閲覧用 URL を返す。

    GAS 側は JSON で
    { folderId, fileName, mimeType, base64Data, apiKey } を受け取り、
    { status: \"success\", url: \"...\" } 形式で返す想定。
    """
    gas_url = _secret_str(SECRET_GAS_UPLOAD_URL)
    gas_api_key = _secret_str(SECRET_GAS_API_KEY)
    folder_id = _secret_str(SECRET_GOOGLE_DRIVE_FOLDER_ID)
    if not gas_url or not gas_api_key or not folder_id:
        return _fallback_image_url_when_gas_unconfigured()

    base64_data = base64.b64encode(data).decode("ascii")
    payload = {
        "folderId": folder_id,
        "fileName": filename,
        "mimeType": mime,
        "base64Data": base64_data,
        "apiKey": gas_api_key,
    }

    resp = requests.post(
        gas_url,
        json=payload,
        headers={"Content-Type": "application/json"},
        timeout=_gas_upload_timeout_seconds(),
    )
    resp.raise_for_status()

    try:
        body = resp.json()
    except json.JSONDecodeError as e:
        raise RuntimeError(f"GAS の応答が JSON ではありません: {resp.text[:500]}") from e

    if body.get("status") != "success":
        raise RuntimeError(
            body.get("message") or body.get("error") or f"GAS アップロード失敗: {body!r}"
        )

    url = body.get("url") or body.get("webViewLink") or body.get("fileUrl")
    if not url:
        raise RuntimeError(f"GAS 応答に URL がありません: {body!r}")
    return str(url)


def _optional_amount_cell(yen: int) -> int:
    """0 以下は 0（数値列・スプレッドシートともに空欄相当）。"""
    v = _finite_int(yen, 0)
    return max(0, v)


def _int_from_cell(v: Any) -> int:
    """セル値を有限な int に（計算前の正規化用）。"""
    return _finite_int(v, 0)


def _coerce_money_columns_for_recalc(df: pd.DataFrame) -> pd.DataFrame:
    """数値列を ``pd.to_numeric`` で揃え、inf/NaN を 0 にして int 化する。"""
    out = df.copy()
    money_cols = (
        COL_PRICE_EXCL,
        COL_PRICE_INCL,
        COL_PLANNED_SALE,
        COL_PLANNED_SALE_INCL,
        COL_ACTUAL_SALE,
        COL_ACTUAL_SALE_INCL,
        COL_GROSS_PROFIT,
        COL_QTY,
    )
    for c in money_cols:
        if c not in out.columns:
            continue
        s = (
            _series_to_numeric_loose(out[c])
            .replace([np.inf, -np.inf], np.nan)
            .fillna(0)
        )
        x = pd.to_numeric(s, errors="coerce").to_numpy(dtype=np.float64, copy=False)
        x = np.where(np.isfinite(x), np.rint(x), 0.0)
        out[c] = x.astype(np.int64)
    return out


def _normalize_stock_status(status: str) -> str:
    s = (status or "").strip()
    return s if s in STOCK_STATUS_OPTIONS else STATUS_IN_STOCK


def _ledger_stocktake_dates_parsed(df: pd.DataFrame) -> pd.Series:
    if COL_LAST_STOCKTAKE not in df.columns:
        return pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    return pd.to_datetime(df[COL_LAST_STOCKTAKE], errors="coerce")


def _today_jst_date() -> date:
    return datetime.now(TZ_JP).date()


def _mask_ledger_in_stock(df: pd.DataFrame) -> pd.Series:
    if COL_STOCK_STATUS not in df.columns:
        return pd.Series(False, index=df.index)
    return (
        df[COL_STOCK_STATUS].astype(str).str.strip().map(_normalize_stock_status)
        == STATUS_IN_STOCK
    )


def _mask_ledger_stocktake_unverified(df: pd.DataFrame) -> pd.Series:
    """在庫中かつ棚卸日が空または解釈不能な行。"""
    m_in = _mask_ledger_in_stock(df)
    dt = _ledger_stocktake_dates_parsed(df)
    return m_in & dt.isna()


def _mask_ledger_stocktake_today_jst(df: pd.DataFrame) -> pd.Series:
    """在庫中かつ棚卸日が今日（JST）の行。"""
    m_in = _mask_ledger_in_stock(df)
    dt = _ledger_stocktake_dates_parsed(df)
    today = _today_jst_date()
    parts = (
        dt.dt.year.eq(today.year)
        & dt.dt.month.eq(today.month)
        & dt.dt.day.eq(today.day)
    )
    return m_in & dt.notna() & parts


def _ledger_in_stock_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or COL_STOCK_STATUS not in df.columns:
        return df.iloc[:0].copy()
    return df.loc[_mask_ledger_in_stock(df)].copy()


def _build_gemini_inventory_context(df: pd.DataFrame, *, max_lines: int = 90) -> str:
    """在庫中の行だけを短い箇条書きにし、画像照合用プロンプトへ埋め込む。"""
    sub = _ledger_in_stock_rows(df)
    if sub.empty:
        return ""
    lines: list[str] = []
    for _, row in sub.iterrows():
        if len(lines) >= max_lines:
            break
        mid = str(row.get(COL_MANAGEMENT_ID, "") or "").strip()
        pn = str(row.get(COL_NAME, "") or "").strip().replace("\n", " ")
        su = str(row.get(COL_SUPPLIER, "") or "").strip().replace("\n", " ")
        cogs = _int_from_cell(row.get(COL_PRICE_EXCL))
        lines.append(
            f"- 管理ID={json.dumps(mid, ensure_ascii=False)} "
            f"商品名={json.dumps(pn, ensure_ascii=False)} "
            f"仕入先={json.dumps(su, ensure_ascii=False)} "
            f"仕入金額税抜={cogs}"
        )
    return "\n".join(lines)


def _fuzzy_ledger_match_rows(
    df: pd.DataFrame,
    product_name: str,
    supplier: str,
    *,
    limit: int = 8,
) -> pd.DataFrame:
    """在庫中の行から、商品名・仕入先の近い候補を返す（写真解析後の補助）。"""
    sub = _ledger_in_stock_rows(df)
    if sub.empty:
        return sub.iloc[:0]
    pn = (product_name or "").strip().casefold()
    su = (supplier or "").strip().casefold()
    if not pn and not su:
        return sub.iloc[:0]
    scores: list[tuple[float, Any]] = []
    for i, row in sub.iterrows():
        rpn = str(row.get(COL_NAME, "") or "").strip().casefold()
        rsu = str(row.get(COL_SUPPLIER, "") or "").strip().casefold()
        a = f"{rpn} {rsu}".strip()
        b = f"{pn} {su}".strip()
        if not a:
            continue
        r0 = difflib.SequenceMatcher(None, a, b).ratio() if b else 0.0
        r1 = difflib.SequenceMatcher(None, rpn, pn).ratio() if pn else 0.0
        r2 = difflib.SequenceMatcher(None, rsu, su).ratio() if su else 0.0
        bonus = 0.0
        if pn and pn in rpn:
            bonus += 0.12
        if su and su in rsu:
            bonus += 0.12
        sc = max(r0, 0.55 * r1 + 0.45 * r2) + bonus
        scores.append((sc, i))
    scores.sort(key=lambda x: -x[0])
    picked = [i for _, i in scores[:limit]]
    if not picked:
        return sub.iloc[:0]
    return sub.loc[picked]


def _infer_tax_rate_from_lines_vectorized(
    excl: np.ndarray, incl: np.ndarray
) -> np.ndarray:
    """各行の仕入税抜・税込から税率を推定（:func:`_infer_tax_rate_from_main_line` と同じ優先）。"""
    excl = np.asarray(excl, dtype=np.int64)
    incl = np.asarray(incl, dtype=np.int64)
    default_r = float(CONSUMPTION_TAX_RATE)
    out = np.full(excl.shape[0], default_r, dtype=np.float64)
    excl_le0 = excl <= 0
    out[excl_le0] = default_r
    incl_le = (~excl_le0) & (incl <= excl)
    out[incl_le] = 0.0
    cand = (~excl_le0) & (~incl_le)
    if not np.any(cand):
        return out
    ec = excl[cand].astype(np.float64)
    ic = incl[cand]
    idx = np.flatnonzero(cand)
    rates_found = np.full(len(idx), default_r, dtype=np.float64)
    matched = np.zeros(len(idx), dtype=bool)
    for _label, rate in CONSUMPTION_TAX_CHOICE_TO_RATE.items():
        rr = float(rate)
        pred = np.rint(ec * (1.0 + rr)).astype(np.int64)
        m_new = (~matched) & (pred == ic)
        rates_found[m_new] = rr
        matched |= m_new
    out[idx] = rates_found
    return out


def _compute_gross_profit_row(
    cogs_line_excl: int,
    planned_line_excl: int,
    actual_line_excl: int,
    status: str,
) -> int | None:
    """税抜ベース（行計）。販売済は実売行計−原価、在庫中は販売予定行計−原価。算出不可時は None。"""
    st = _normalize_stock_status(status)
    cg = _finite_int(cogs_line_excl, 0)
    plex = _finite_int(planned_line_excl, 0)
    aex = _finite_int(actual_line_excl, 0)
    if st == STATUS_SOLD:
        if aex > 0:
            return int(aex - cg)
        return None
    if st == STATUS_IN_STOCK:
        if plex > 0:
            return int(plex - cg)
        return None
    return None


def _recalc_gross_profit_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """販売予定/実売の税込総額列と粗利を再計算する（数値列は int で統一）。

    行ループではなく NumPy ベクトル演算で処理し、大行数でも応答を速くする。
    """
    need = (
        COL_GROSS_PROFIT,
        COL_STOCK_STATUS,
        COL_PRICE_EXCL,
        COL_PRICE_INCL,
        COL_QTY,
        COL_PLANNED_SALE,
        COL_ACTUAL_SALE,
        COL_PLANNED_SALE_INCL,
        COL_ACTUAL_SALE_INCL,
    )
    if not all(c in df.columns for c in need):
        return df.copy()
    out = _coerce_money_columns_for_recalc(df)
    n = len(out.index)
    if n == 0:
        return out

    def _i64_col(c: str) -> np.ndarray:
        x = pd.to_numeric(out[c], errors="coerce").fillna(0).to_numpy(
            dtype=np.float64, copy=False
        )
        return np.where(np.isfinite(x), np.rint(x), 0.0).astype(np.int64, copy=False)

    cogs = _i64_col(COL_PRICE_EXCL)
    line_in = _i64_col(COL_PRICE_INCL)
    qty = np.maximum(_i64_col(COL_QTY), 1)
    pu = _i64_col(COL_PLANNED_SALE)
    au = _i64_col(COL_ACTUAL_SALE)

    st_raw = out[COL_STOCK_STATUS].astype(str).str.strip().to_numpy()
    st = np.where(np.isin(st_raw, list(STOCK_STATUS_OPTIONS)), st_raw, STATUS_IN_STOCK)
    out[COL_STOCK_STATUS] = st

    tax_r = _infer_tax_rate_from_lines_vectorized(cogs, line_in)

    plex = np.where(pu > 0, pu * qty, 0).astype(np.int64, copy=False)
    fplex = plex.astype(np.float64, copy=False)
    pincl = np.zeros(n, dtype=np.int64)
    m_plex = plex > 0
    pincl[m_plex] = np.rint(
        fplex[m_plex] * (1.0 + tax_r[m_plex].astype(np.float64))
    ).astype(np.int64)

    is_sold = st == STATUS_SOLD
    aex = np.where(is_sold & (au > 0), au * qty, 0).astype(np.int64, copy=False)
    faex = aex.astype(np.float64, copy=False)
    aincl = np.zeros(n, dtype=np.int64)
    m_aex = aex > 0
    aincl[m_aex] = np.rint(
        faex[m_aex] * (1.0 + tax_r[m_aex].astype(np.float64))
    ).astype(np.int64)

    gp = np.zeros(n, dtype=np.int64)
    m_sold = is_sold & (aex > 0)
    gp[m_sold] = aex[m_sold] - cogs[m_sold]
    is_in = st == STATUS_IN_STOCK
    m_in = is_in & (plex > 0)
    gp[m_in] = plex[m_in] - cogs[m_in]

    out[COL_PLANNED_SALE_INCL] = pincl
    out[COL_ACTUAL_SALE_INCL] = aincl
    out[COL_GROSS_PROFIT] = gp
    return out


def _max_management_serial_from_dataframe(df: pd.DataFrame) -> int:
    """DataFrame の管理ID列から最大シリアルを返す。"""
    mx = 0
    if df is None or df.empty or COL_MANAGEMENT_ID not in df.columns:
        return mx
    for v in df[COL_MANAGEMENT_ID].astype(str):
        s = v.strip()
        if not s:
            continue
        m = re.fullmatch(r"(?i)G(\d+)", s)
        if m:
            mx = max(mx, int(m.group(1)))
            continue
        if s.isdigit():
            mx = max(mx, int(s))
    return mx


def allocate_management_ids(ws: Any, count: int) -> list[str]:
    """管理ID（G########）を count 件、CSV またはシート現状から連番で採番する。"""
    if count <= 0:
        return []
    if _uses_local_inventory_csv():
        df = _inventory_csv_read_df()
        mx = _max_management_serial_from_dataframe(df)
        return [f"G{mx + i + 1:08d}" for i in range(count)]
    if ws is None:
        return []
    try:
        df_exist = load_inventory_dataframe()
    except Exception:
        df_exist = None
    mx = _max_management_serial_from_dataframe(df_exist) if df_exist is not None else 0
    return [f"G{mx + i + 1:08d}" for i in range(count)]


def _inventory_row_values_for_append(
    dt_a: str,
    dt_purchase: str,
    purchase_movement: str,
    product_name: str,
    supplier: str,
    line_price_excl_yen: int,
    line_price_incl_yen: int,
    image_url: str,
    management_id: str,
    memo: str,
    *,
    planned_sale_unit_excl_yen: int = 0,
    actual_sale_unit_excl_yen: int = 0,
    stock_status: str = STATUS_IN_STOCK,
    consumption_tax_rate: float | None = None,
    sale_source_management_id: str = "",
    loan_datetime: str = "",
    voucher_recorded_at: str = "",
    voucher_evidence_url: str = "",
) -> list[Any]:
    """台帳 EXPECTED_HEADERS 順の1行分セル値を組み立てる（追記用）。"""
    cogs = _finite_int(line_price_excl_yen, 0)
    qty_i = 1
    pl_u = _finite_int(planned_sale_unit_excl_yen, 0)
    ac_u = _finite_int(actual_sale_unit_excl_yen, 0)
    stt = _normalize_stock_status(str(stock_status))
    tax_r = (
        float(consumption_tax_rate)
        if consumption_tax_rate is not None
        and math.isfinite(float(consumption_tax_rate))
        else _infer_tax_rate_from_main_line(
            _finite_int(line_price_excl_yen, 0),
            _finite_int(line_price_incl_yen, 0),
        )
    )
    plex, pincl, aex, aincl = _planned_actual_line_amounts(
        qty_i, pl_u, ac_u, stt, tax_r
    )
    planned_unit_cell = _optional_amount_cell(pl_u)
    planned_incl_cell = _optional_amount_cell(pincl)
    actual_unit_cell = (
        _optional_amount_cell(ac_u) if stt == STATUS_SOLD else 0
    )
    actual_incl_cell = (
        _optional_amount_cell(aincl) if stt == STATUS_SOLD else 0
    )
    gp = _compute_gross_profit_row(
        cogs,
        plex,
        aex if stt == STATUS_SOLD else 0,
        stt,
    )
    gross_cell = 0 if gp is None else _finite_int(gp, 0)
    pm = (purchase_movement or "").strip()
    return [
        dt_a,
        product_name,
        supplier,
        1,
        line_price_excl_yen,
        line_price_incl_yen,
        planned_unit_cell,
        planned_incl_cell,
        actual_unit_cell,
        actual_incl_cell,
        gross_cell,
        stt,
        memo,
        image_url,
        management_id,
        "",
        (sale_source_management_id or "").strip(),
        (voucher_recorded_at or "").strip(),
        (voucher_evidence_url or "").strip(),
        dt_purchase,
        pm,
        (loan_datetime or "").strip(),
        (dt_a if stt == STATUS_SOLD else ""),
        (
            pm
            if stt == STATUS_SOLD and _movement_is_outbound(pm)
            else ("出庫（販売）" if stt == STATUS_SOLD else "")
        ),
    ]


def _append_inventory_data_rows(rows: list[list[Any]]) -> None:
    """台帳に複数行をまとめて追記する（CSV は1回の再計算・書き込み、シートは append_rows 1回）。

    連続で数十行追記するとき、行ごとの append + 書式 API を繰り返すと失敗や取りこぼしの原因になるため一括にする。
    """
    if not rows:
        return
    if _uses_local_inventory_csv():
        df = _inventory_csv_read_df()
        for rv in rows:
            new_row = dict(zip(EXPECTED_HEADERS, rv))
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df = _recalc_gross_profit_dataframe(df)
        _inventory_csv_write_df(df.reindex(columns=EXPECTED_HEADERS))
        return
    ws = ensure_worksheet_header()
    if ws is None:
        raise RuntimeError(
            "スプレッドシートに接続できません。"
            f"{SECRET_GOOGLE_SPREADSHEET_ID} とサービスアカウント権限を確認してください。"
        )
    try:
        try:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
        except Exception:
            try:
                ws.resize(rows=int(ws.row_count) + len(rows) + 200)
            except Exception:
                pass
            ws.append_rows(rows, value_input_option="USER_ENTERED")
        try:
            _apply_inventory_amount_number_formats(ws)
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(f"スプレッドシート追記に失敗しました: {e}") from e
    _bump_inventory_sheet_cache_bust()


def append_sheet_row(
    purchase_movement: str,
    product_name: str,
    supplier: str,
    line_price_excl_yen: int,
    line_price_incl_yen: int,
    image_url: str,
    management_id: str,
    memo: str = "",
    record_datetime: str | None = None,
    *,
    planned_sale_unit_excl_yen: int = 0,
    actual_sale_unit_excl_yen: int = 0,
    stock_status: str = STATUS_IN_STOCK,
    consumption_tax_rate: float | None = None,
    sale_source_management_id: str = "",
    loan_datetime: str = "",
    voucher_recorded_at: str = "",
    voucher_evidence_url: str = "",
):
    """1点1行で台帳に追記する（数量列は常に 1。仕入単価列は持たない）。

    A列「日時」は **追記実行の JST**。仕入の暦は「仕入日時」に ``record_datetime``（EXIF 等）を渡す。
    """
    dt_a = jst_now_str()
    dt_purchase = ((record_datetime or "").strip() or dt_a)
    row_vals = _inventory_row_values_for_append(
        dt_a,
        dt_purchase,
        purchase_movement,
        product_name,
        supplier,
        line_price_excl_yen,
        line_price_incl_yen,
        image_url,
        management_id,
        memo,
        planned_sale_unit_excl_yen=planned_sale_unit_excl_yen,
        actual_sale_unit_excl_yen=actual_sale_unit_excl_yen,
        stock_status=stock_status,
        consumption_tax_rate=consumption_tax_rate,
        sale_source_management_id=sale_source_management_id,
        loan_datetime=loan_datetime,
        voucher_recorded_at=voucher_recorded_at,
        voucher_evidence_url=voucher_evidence_url,
    )
    _append_inventory_data_rows([row_vals])


def _sheet_header_row_to_expected_list(header: list[str], row: list[Any]) -> list[str]:
    """ヘッダー名でセルを対応付け、EXPECTED_HEADERS 順の1行にする。旧「入出庫種別」は入庫種別へ移す。"""
    h = [("" if x is None else str(x)).strip() for x in header]
    rlist = [("" if x is None else str(x)) for x in list(row)]
    h2: list[str] = []
    r2: list[str] = []
    for i, nm in enumerate(h):
        if nm == LEGACY_COL_UNIT_PRICE:
            continue
        h2.append(nm)
        r2.append(rlist[i] if i < len(rlist) else "")
    dd = dict(zip(h2, r2))
    out: dict[str, str] = {c: str(dd.get(c, "") or "") for c in EXPECTED_HEADERS}
    if LEGACY_COL_MOVEMENT_TYPE in dd and not (out.get(COL_PURCHASE_MOVEMENT) or "").strip():
        out[COL_PURCHASE_MOVEMENT] = str(dd.get(LEGACY_COL_MOVEMENT_TYPE, "") or "")
    return [out[c] for c in EXPECTED_HEADERS]


def load_inventory_dataframe() -> pd.DataFrame | None:
    """1行目をヘッダー、2行目以降をデータとして読み込み、列は EXPECTED_HEADERS に揃える。"""
    if _uses_local_inventory_csv():
        return _inventory_csv_read_df()
    sid = _secret_str(SECRET_GOOGLE_SPREADSHEET_ID)
    if not sid:
        return None
    wname = _secret_str(SECRET_GOOGLE_WORKSHEET_NAME, DEFAULT_WORKSHEET_NAME)
    bust = int(st.session_state.get(SESSION_KEY_INV_SHEET_CACHE_BUST, 0))
    raw = _inventory_sheet_get_all_values_cached(str(sid), str(wname), bust)
    if raw is None:
        return None
    if not raw:
        return pd.DataFrame(columns=EXPECTED_HEADERS)
    header0 = [("" if c is None else str(c)).strip() for c in raw[0]]
    rows = raw[1:]
    data_rows = [_sheet_header_row_to_expected_list(header0, list(r)) for r in rows]
    return pd.DataFrame(data_rows, columns=EXPECTED_HEADERS)


def _ledger_hint_dataframe() -> pd.DataFrame | None:
    """登録画面の台帳照合用に在庫を読む（失敗時は None）。"""
    if not _uses_local_inventory_csv() and not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        return None
    try:
        df = load_inventory_dataframe()
        return df
    except Exception:
        return None


def _refresh_ledger_quick_search_candidates(df_ledger: pd.DataFrame | None) -> None:
    """写真解析・手入力の商品名／仕入先から、在庫中の近い行を session_state に格納する。"""
    if df_ledger is None or df_ledger.empty:
        st.session_state.pop("ledger_quick_candidates", None)
        return
    pn = str(st.session_state.get("field_product_name", "") or "").strip()
    su = str(st.session_state.get("field_supplier", "") or "").strip()
    cand = _fuzzy_ledger_match_rows(df_ledger, pn, su, limit=8)
    if cand.empty:
        st.session_state.pop("ledger_quick_candidates", None)
    else:
        st.session_state["ledger_quick_candidates"] = cand


def _ledger_unique_col_values(df: pd.DataFrame, col: str, *, max_n: int = 800) -> list[str]:
    """台帳 DataFrame から列のユニーク値（空除く）を昇順で返す。"""
    if df is None or df.empty or col not in df.columns:
        return []
    s = df[col].astype(str).str.strip()
    s = s[s != ""]
    return sorted(set(s.tolist()), key=lambda x: (x.casefold(), x))[:max_n]


def _ledger_in_stock_management_ids(df: pd.DataFrame, *, max_n: int = 600) -> list[str]:
    """在庫中の行の管理ID一覧（販売元のプルダウン用）。"""
    if df is None or df.empty or COL_MANAGEMENT_ID not in df.columns:
        return []
    sub = df.loc[_mask_ledger_in_stock(df)]
    s = sub[COL_MANAGEMENT_ID].astype(str).str.strip()
    s = s[s != ""]
    return sorted(set(s.tolist()), key=lambda x: (x.casefold(), x))[:max_n]


def _sync_planned_sale_from_ledger_in_stock_match() -> None:
    """在庫中の台帳行で、入力中の商品名（＋仕入先が入っていれば一致）に合わせ販売予定（税抜・1点）を反映する。"""
    try:
        df = load_inventory_dataframe()
    except Exception:
        return
    if df is None or df.empty or COL_PLANNED_SALE not in df.columns:
        return
    pn = str(st.session_state.get("field_product_name", "") or "").strip()
    if not pn:
        return
    sub = df.loc[_mask_ledger_in_stock(df)]
    if sub.empty:
        return
    m = sub[COL_NAME].astype(str).str.strip() == pn
    su = str(st.session_state.get("field_supplier", "") or "").strip()
    if su and COL_SUPPLIER in sub.columns:
        m = m & (sub[COL_SUPPLIER].astype(str).str.strip() == su)
    hit = sub.loc[m]
    if hit.empty:
        return
    if COL_MANAGEMENT_ID in hit.columns:
        hit = hit.sort_values(
            COL_MANAGEMENT_ID,
            key=lambda s: s.astype(str).str.strip(),
            na_position="last",
        )
    pl = _finite_int(hit.iloc[0].get(COL_PLANNED_SALE), 0)
    st.session_state.field_planned_sale_excl = max(0, int(pl))


def _on_ledger_pick_product_name() -> None:
    v = st.session_state.get("ledger_pick_product_name", "")
    if v and v != LEDGER_PICK_PLACEHOLDER:
        st.session_state.field_product_name = v
        st.session_state.ledger_pick_product_name = LEDGER_PICK_PLACEHOLDER
        _sync_planned_sale_from_ledger_in_stock_match()


def _on_ledger_pick_supplier() -> None:
    v = st.session_state.get("ledger_pick_supplier", "")
    if v and v != LEDGER_PICK_PLACEHOLDER:
        st.session_state.field_supplier = v
        st.session_state.ledger_pick_supplier = LEDGER_PICK_PLACEHOLDER
        _sync_planned_sale_from_ledger_in_stock_match()


def _on_sale_pick_source_id() -> None:
    v = st.session_state.get("sale_pick_source_id", "")
    if v and v != LEDGER_PICK_PLACEHOLDER:
        st.session_state.field_sale_source_mgmt_id = v
        st.session_state.sale_pick_source_id = LEDGER_PICK_PLACEHOLDER


def _ledger_row_content_differs(
    a: pd.Series, b: pd.Series, *, skip_cols: frozenset[str]
) -> bool:
    """同一行の表示内容に差があるか（日時列などは比較から除外）。"""
    for c in EXPECTED_HEADERS:
        if c in skip_cols:
            continue
        if c not in a.index or c not in b.index:
            continue
        if str(a.get(c, "")).strip() != str(b.get(c, "")).strip():
            return True
    return False


def _stamp_row_datetime_on_changes(
    before: pd.DataFrame, after: pd.DataFrame
) -> pd.DataFrame:
    """行順・行数が一致する前提で、内容が変わった行の「日時」を JST 現在にする。"""
    out = after.copy()
    if COL_DATETIME not in out.columns or before is None or before.empty or after.empty:
        return out
    ts = jst_now_str()
    skip = frozenset({COL_DATETIME})
    n = min(len(before), len(after))
    loc_dt = out.columns.get_loc(COL_DATETIME)
    for i in range(n):
        if _ledger_row_content_differs(before.iloc[i], after.iloc[i], skip_cols=skip):
            out.iloc[i, loc_dt] = ts
    for i in range(n, len(after)):
        out.iloc[i, loc_dt] = ts
    return out


def _cell_value_for_sheet(v: Any) -> Any:
    """スプレッドシート1セル向けに欠損を正規化する。数値の NaN は 0、それ以外の欠損は空文字。"""
    try:
        if pd.api.types.is_scalar(v) and pd.isna(v):
            if isinstance(v, (float, np.floating)):
                return 0
            return ""
        if isinstance(v, (float, np.floating)) and (
            not math.isfinite(float(v)) or pd.isna(v)
        ):
            return 0
    except Exception:
        pass
    return v


def overwrite_inventory_worksheet_from_dataframe(
    df: pd.DataFrame, *, previous_df: pd.DataFrame | None = None
) -> None:
    """編集後の DataFrame で inventory.csv またはワークシートを全置換する。

    ``previous_df`` を渡したとき、同一行位置で内容が変わった行の **日時** を保存実行の JST に更新する。
    """
    work = df
    if previous_df is not None and COL_DATETIME in df.columns:
        work = _stamp_row_datetime_on_changes(
            previous_df.reset_index(drop=True),
            df.reset_index(drop=True),
        )
    out = _recalc_gross_profit_dataframe(
        work.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    )
    if _uses_local_inventory_csv():
        _inventory_csv_write_df(out)
        return
    ws = _get_or_create_inventory_worksheet()
    if ws is None:
        raise RuntimeError(
            f"スプレッドシートに接続できません。{SECRET_GOOGLE_SPREADSHEET_ID} とサービスアカウントを確認してください。"
        )
    values: list[list[Any]] = [EXPECTED_HEADERS]
    if not out.empty:
        for row in out[EXPECTED_HEADERS].to_numpy(dtype=object):
            values.append([_cell_value_for_sheet(x) for x in row.tolist()])
    try:
        ws.clear()
        ws.update("A1", values, value_input_option="USER_ENTERED")
        try:
            _apply_inventory_amount_number_formats(ws)
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(f"スプレッドシートの上書きに失敗しました: {e}") from e
    _bump_inventory_sheet_cache_bust()


def _ledger_df_loosen_numeric_columns_for_assignment(df: pd.DataFrame) -> None:
    """StringDtype 等の厳格な列に int を代入すると失敗するため、金額・数量列を object に揃える（原地変更）。"""
    for _c in (
        COL_QTY,
        COL_PRICE_EXCL,
        COL_PRICE_INCL,
        COL_PLANNED_SALE,
        COL_PLANNED_SALE_INCL,
        COL_ACTUAL_SALE,
        COL_ACTUAL_SALE_INCL,
        COL_GROSS_PROFIT,
    ):
        if _c in df.columns:
            df[_c] = df[_c].astype(object)


def lookup_ledger_row_by_management_id(
    df: pd.DataFrame | None, management_id: str
) -> pd.Series | None:
    """管理IDに一致する台帳行を1件返す。無ければ None。"""
    if df is None or df.empty or COL_MANAGEMENT_ID not in df.columns:
        return None
    sid = (management_id or "").strip()
    if not sid:
        return None
    m = df[COL_MANAGEMENT_ID].astype(str).str.strip() == sid
    if not m.any():
        return None
    return df.loc[m].iloc[0]


def _split_management_ids_from_field(raw: str) -> list[str]:
    """販売元管理ID欄をカンマ・読点・区切り文字・空白・改行で分割し、空を除いたリストを返す。"""
    parts = re.split(r"[,、;；\s\n]+", (raw or "").strip())
    return [p.strip() for p in parts if p.strip()]


def apply_outbound_sale_to_ledger_by_management_id(
    source_management_id: str,
    *,
    actual_sale_unit_excl_yen: int,
    new_image_url: str = "",
    memo_suffix: str = "",
    update_sale_voucher: bool = False,
    sale_voucher_recorded_at: str = "",
    sale_voucher_evidence_url: str = "",
    sale_outbound_type: str = "出庫（販売）",
) -> None:
    """在庫中の1行を販売済に更新（新規行なし）。A列「日時」は確定実行の JST。出庫種別は ``sale_outbound_type``（例: 出庫（販売）／出庫（浮貸））。"""
    sid = (source_management_id or "").strip()
    if not sid:
        raise ValueError("販売元管理ID（管理ID）が空です。")
    df_src = load_inventory_dataframe()
    if df_src is None or df_src.empty:
        raise RuntimeError("台帳を読み込めませんでした。")
    df_src = df_src.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    _ledger_df_loosen_numeric_columns_for_assignment(df_src)
    msk = df_src[COL_MANAGEMENT_ID].astype(str).str.strip() == sid
    if not msk.any():
        raise RuntimeError(f"管理ID {sid} の行が台帳に見つかりません。")
    if int(msk.sum()) != 1:
        raise RuntimeError(f"管理ID {sid} が複数行に重複しています。")
    cur_st = _normalize_stock_status(
        str(df_src.loc[msk, COL_STOCK_STATUS].iloc[0])
    )
    if cur_st != STATUS_IN_STOCK:
        raise RuntimeError(
            f"管理ID {sid} は「{cur_st}」のため、販売反映の対象外です（在庫中の行のみ更新します）。"
        )
    av = _finite_int(actual_sale_unit_excl_yen, 0)
    if av < 1:
        raise RuntimeError("実売金額（税抜）は1円以上にしてください。")

    now_exec = jst_now_str()
    _prev_row_dt = str(df_src.loc[msk, COL_DATETIME].iloc[0] or "")
    if COL_PURCHASE_DATETIME in df_src.columns:
        cp = str(df_src.loc[msk, COL_PURCHASE_DATETIME].iloc[0] or "").strip()
        if not cp:
            df_src.loc[msk, COL_PURCHASE_DATETIME] = _prev_row_dt
    if COL_PURCHASE_MOVEMENT in df_src.columns:
        cm = str(df_src.loc[msk, COL_PURCHASE_MOVEMENT].iloc[0] or "").strip()
        if not cm:
            df_src.loc[msk, COL_PURCHASE_MOVEMENT] = "入庫（購入）"
    if COL_SALE_DATETIME in df_src.columns:
        df_src.loc[msk, COL_SALE_DATETIME] = now_exec
    if COL_SALE_OUTBOUND_TYPE in df_src.columns:
        _ot = (sale_outbound_type or "").strip() or "出庫（販売）"
        df_src.loc[msk, COL_SALE_OUTBOUND_TYPE] = _ot
    df_src.loc[msk, COL_DATETIME] = now_exec
    df_src.loc[msk, COL_STOCK_STATUS] = STATUS_SOLD
    df_src.loc[msk, COL_ACTUAL_SALE] = av
    df_src.loc[msk, COL_SALE_SOURCE_MGMT_ID] = ""
    cur_img = str(df_src.loc[msk, COL_IMAGE_URL].iloc[0] or "").strip()
    nu = (new_image_url or "").strip()
    if not cur_img and nu:
        df_src.loc[msk, COL_IMAGE_URL] = nu
    if (memo_suffix or "").strip():
        old_memo = str(df_src.loc[msk, COL_MEMO].iloc[0] or "").strip()
        tag = (memo_suffix or "").strip()
        df_src.loc[msk, COL_MEMO] = (old_memo + "\n" if old_memo else "") + tag
    if update_sale_voucher and (sale_voucher_evidence_url or "").strip():
        df_src.loc[msk, COL_VOUCHER_RECORDED_AT] = (
            sale_voucher_recorded_at or ""
        ).strip()
        df_src.loc[msk, COL_VOUCHER_EVIDENCE_URL] = sale_voucher_evidence_url.strip()
    df_src = _recalc_gross_profit_dataframe(df_src)
    overwrite_inventory_worksheet_from_dataframe(df_src)


def apply_outbound_loan_in_stock_datetime_by_management_id(
    source_management_id: str,
    *,
    loan_datetime_jst: str | None = None,
    new_image_url: str = "",
    memo_suffix: str = "",
) -> None:
    """出庫（浮貸）かつ在庫中のまま: 該当行の **浮貸日時** を記録し、出庫種別を記録。新規行は追加しない。"""
    sid = (source_management_id or "").strip()
    if not sid:
        raise ValueError("管理IDが空です。")
    df_src = load_inventory_dataframe()
    if df_src is None or df_src.empty:
        raise RuntimeError("台帳を読み込めませんでした。")
    df_src = df_src.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    _ledger_df_loosen_numeric_columns_for_assignment(df_src)
    msk = df_src[COL_MANAGEMENT_ID].astype(str).str.strip() == sid
    if not msk.any():
        raise RuntimeError(f"管理ID {sid} の行が台帳に見つかりません。")
    if int(msk.sum()) != 1:
        raise RuntimeError(f"管理ID {sid} が複数行に重複しています。")
    cur_st = _normalize_stock_status(
        str(df_src.loc[msk, COL_STOCK_STATUS].iloc[0])
    )
    if cur_st != STATUS_IN_STOCK:
        raise RuntimeError(
            f"管理ID {sid} は「{cur_st}」のため、浮貸（在庫中）の記録対象外です（在庫中の行のみ）。"
        )
    loan_dt = (loan_datetime_jst or "").strip() or jst_now_str()
    now_exec = jst_now_str()
    if COL_LOAN_DATETIME in df_src.columns:
        df_src.loc[msk, COL_LOAN_DATETIME] = loan_dt
    if COL_SALE_OUTBOUND_TYPE in df_src.columns:
        df_src.loc[msk, COL_SALE_OUTBOUND_TYPE] = "出庫（浮貸）"
    df_src.loc[msk, COL_DATETIME] = now_exec
    cur_img = str(df_src.loc[msk, COL_IMAGE_URL].iloc[0] or "").strip()
    nu = (new_image_url or "").strip()
    if not cur_img and nu:
        df_src.loc[msk, COL_IMAGE_URL] = nu
    if (memo_suffix or "").strip():
        old_memo = str(df_src.loc[msk, COL_MEMO].iloc[0] or "").strip()
        tag = (memo_suffix or "").strip()
        df_src.loc[msk, COL_MEMO] = (old_memo + "\n" if old_memo else "") + tag
    df_src = _recalc_gross_profit_dataframe(df_src)
    overwrite_inventory_worksheet_from_dataframe(df_src)


def apply_last_stocktake_jst_for_management_id(management_id: str) -> None:
    """在庫中の1行について「最後に確認した日付（棚卸日）」を本日（JST）にし、日時を更新して保存する。"""
    sid = (management_id or "").strip()
    if not sid:
        raise ValueError("管理IDが空です。")
    df_src = load_inventory_dataframe()
    if df_src is None or df_src.empty:
        raise RuntimeError("台帳を読み込めませんでした。")
    df_src = df_src.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    _ledger_df_loosen_numeric_columns_for_assignment(df_src)
    msk = df_src[COL_MANAGEMENT_ID].astype(str).str.strip() == sid
    if not msk.any():
        raise RuntimeError(f"管理ID {sid} の行が台帳に見つかりません。")
    if int(msk.sum()) != 1:
        raise RuntimeError(f"管理ID {sid} が複数行に重複しています。")
    cur_st = _normalize_stock_status(
        str(df_src.loc[msk, COL_STOCK_STATUS].iloc[0])
    )
    if cur_st != STATUS_IN_STOCK:
        raise RuntimeError(
            f"管理ID {sid} は在庫中ではないため棚卸確定できません（現在: {cur_st}）。"
        )
    today_s = _today_jst_date().isoformat()
    now_exec = jst_now_str()
    if COL_LAST_STOCKTAKE in df_src.columns:
        df_src.loc[msk, COL_LAST_STOCKTAKE] = today_s
    df_src.loc[msk, COL_DATETIME] = now_exec
    df_src = _recalc_gross_profit_dataframe(df_src)
    overwrite_inventory_worksheet_from_dataframe(df_src)


def _apply_ledger_sort(
    df: pd.DataFrame,
    primary: str,
    primary_asc: bool,
    secondary: str,
    secondary_asc: bool,
    tertiary: str,
    tertiary_asc: bool,
) -> pd.DataFrame:
    """在庫一覧の表示用ソート（最大3キー、コピーを返す）。"""
    if df.empty:
        return df
    col_map = {
        "日時": COL_DATETIME,
        "仕入先・取引先": COL_SUPPLIER,
        "管理ID": COL_MANAGEMENT_ID,
        "仕入日時": COL_PURCHASE_DATETIME,
        "販売日時": COL_SALE_DATETIME,
    }
    pairs: list[tuple[str, bool]] = []
    for label, asc in (
        (primary, primary_asc),
        (secondary, secondary_asc),
        (tertiary, tertiary_asc),
    ):
        if label == "なし" or label not in col_map:
            continue
        col = col_map[label]
        if any(p[0] == col for p in pairs):
            continue
        pairs.append((col, asc))
    if not pairs:
        return df.copy()

    out = df.copy()
    sort_cols: list[str] = []
    ascending: list[bool] = []
    for col, asc in pairs:
        if col not in out.columns:
            continue
        if col in (COL_DATETIME, COL_PURCHASE_DATETIME, COL_SALE_DATETIME):
            tmp = "_sort_dt_internal"
            if col != COL_DATETIME:
                tmp = f"_sort_dt_internal_{col}"
            out[tmp] = pd.to_datetime(out[col], errors="coerce")
            sort_cols.append(tmp)
        else:
            sort_cols.append(col)
        ascending.append(asc)
    if not sort_cols:
        return out
    out = out.sort_values(by=sort_cols, ascending=ascending, na_position="last")
    return out.drop(columns=[c for c in out.columns if c.startswith("_sort_dt")], errors="ignore")


def _evidence_url_column_all_http_or_blank(s: pd.Series) -> bool:
    """証憑URL列を LinkColumn にできるか（空・欠損・http(s) のみ）。"""
    t = s.fillna("").astype(str).str.strip().str.lower()
    t = t.replace({"nan": "", "none": "", "<na>": "", "nat": ""}, regex=False)
    nonempty = t[t != ""]
    if nonempty.empty:
        return True
    ok = nonempty.str.startswith("http://") | nonempty.str.startswith("https://")
    return bool(ok.all())


def _normalize_evidence_urls_for_link_editor(df: pd.DataFrame, col: str) -> bool:
    """証憑URL列を LinkColumn 向けに正規化（空相当→None）。LinkColumn を使うなら True。"""
    if col not in df.columns:
        return False
    if not _evidence_url_column_all_http_or_blank(df[col]):
        return False
    v = df[col]
    strv = v.fillna("").astype(str).str.strip()
    low = strv.str.lower()
    isblank = v.isna() | low.isin(("", "nan", "none", "<na>", "nat"))
    nb = isblank.fillna(False).to_numpy(dtype=bool, copy=False)
    arr = np.empty(len(strv), dtype=object)
    arr[nb] = None
    arr[~nb] = strv[~isblank].to_numpy(dtype=object, copy=False)
    df[col] = arr
    return True


def _ledger_dashboard_axis_datetime(df: pd.DataFrame) -> pd.Series:
    """ダッシュボードの期間・月次軸に使う日時。販売済は **販売日時**（空なら日時）を優先。"""
    base = pd.to_datetime(df[COL_DATETIME], errors="coerce")
    if COL_SALE_DATETIME not in df.columns or COL_STOCK_STATUS not in df.columns:
        return base
    st = df[COL_STOCK_STATUS].astype(str).map(_normalize_stock_status)
    sold = st == STATUS_SOLD
    sd = pd.to_datetime(df[COL_SALE_DATETIME], errors="coerce")
    alt = sd.combine_first(base)
    return base.where(~sold, alt)


def _prepare_ledger_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """入出庫判定・行金額・年月などの派生列を付与したコピーを返す。

    1点1行ライフサイクル前提: **在庫中** かつ **入庫** だけを仕入（入庫）側に計上し、
    **販売済** かつ実売がある行は **実売行計（税抜・税込）** を売上（出庫）側に計上する
    （「入庫種別」が空の在庫中は入庫扱い。販売済＋実売で売上を計上し、仕入金額を二重に出庫しない）。
    **在庫中** のまま **出庫**（浮貸など別レコード）の行は、従来どおり仕入列ベースで出庫に含める。
    """
    d = df.copy()
    if d.empty:
        return d
    d[COL_DATETIME] = pd.to_datetime(d[COL_DATETIME], errors="coerce")
    qty = _series_to_numeric_loose(d[COL_QTY]).fillna(0)
    line_stored_ex = _series_to_numeric_loose(d[COL_PRICE_EXCL]).fillna(0)
    line_stored_in = _series_to_numeric_loose(d[COL_PRICE_INCL]).fillna(0)
    if COL_STOCK_STATUS in d.columns:
        st_col = d[COL_STOCK_STATUS].astype(str).map(_normalize_stock_status)
    else:
        st_col = pd.Series(STATUS_IN_STOCK, index=d.index)
    pur_mv = (
        d[COL_PURCHASE_MOVEMENT].astype(str).str.strip()
        if COL_PURCHASE_MOVEMENT in d.columns
        else pd.Series("", index=d.index, dtype=str)
    )
    is_in_mv = pur_mv.str.startswith("入庫") | (
        (pur_mv == "") & (st_col == STATUS_IN_STOCK)
    )
    is_out_pur = pur_mv.str.startswith("出庫")
    # 税抜・税込はいずれも行合計として解釈（仕入単価列は廃止）
    line_ex = line_stored_ex.astype(float)
    mask_ex_from_incl = (
        (line_ex.fillna(0) <= 0)
        & (qty.fillna(0) > 0)
        & (line_stored_in.fillna(0) > 0)
    )
    if bool(mask_ex_from_incl.any()):
        fill_ex = line_stored_in.map(
            lambda v: float(_estimate_excl_yen_from_incl_yen(_finite_int(v, 0)))
        )
        line_ex = line_ex.where(~mask_ex_from_incl, fill_ex)
    line_in = line_stored_in.astype(float)
    line_needs_derive_in = line_stored_in.fillna(0) <= 0
    _ex_int = line_ex.fillna(0).round().clip(lower=0).astype(int)
    line_in = line_in.mask(line_needs_derive_in, _ex_int.map(price_incl_tax).astype(float))
    line_ex = line_ex.fillna(0).replace([np.inf, -np.inf], 0)
    line_in = line_in.fillna(0).replace([np.inf, -np.inf], 0)

    ac_u = _series_to_numeric_loose(
        d[COL_ACTUAL_SALE] if COL_ACTUAL_SALE in d.columns else 0
    ).fillna(0)
    ac_incl_line = _series_to_numeric_loose(
        d[COL_ACTUAL_SALE_INCL] if COL_ACTUAL_SALE_INCL in d.columns else 0
    ).fillna(0)
    rev_ex = (ac_u * qty).clip(lower=0)
    m_stock_in = is_in_mv & (st_col == STATUS_IN_STOCK)
    # 販売済は実売ベースで売上計上（区分がまだ入庫の旧行も、実売があればここに含める）
    m_sold_rev = (st_col == STATUS_SOLD) & (rev_ex > 0)
    m_float_out = is_out_pur & (st_col == STATUS_IN_STOCK)

    d["_qty_in"] = qty.where(m_stock_in, 0.0).fillna(0).astype(float)
    d["_amt_ex_in"] = line_ex.where(m_stock_in, 0.0).fillna(0).astype(float)
    d["_amt_in_in"] = line_in.where(m_stock_in, 0.0).fillna(0).astype(float)

    d["_qty_out"] = qty.where(m_sold_rev | m_float_out, 0.0).fillna(0).astype(float)
    d["_amt_ex_out"] = (
        rev_ex.where(m_sold_rev, 0.0).fillna(0)
        + line_ex.where(m_float_out, 0.0).fillna(0)
    ).astype(float)
    d["_amt_in_out"] = (
        ac_incl_line.where(m_sold_rev, 0.0).fillna(0)
        + line_in.where(m_float_out, 0.0).fillna(0)
    ).astype(float)

    axis_dt = _ledger_dashboard_axis_datetime(d)
    d["_ym"] = axis_dt.dt.to_period("M").astype(str)
    d["_year"] = axis_dt.dt.year
    d["_month"] = axis_dt.dt.month
    return d


def _ledger_dashboard_date_bounds(df: pd.DataFrame) -> tuple[date, date]:
    """ダッシュボード用の日付範囲既定値（販売済は販売日時軸を含む）。"""
    s = _ledger_dashboard_axis_datetime(df).dropna()
    if s.empty:
        t = jst_now().date()
        return t, t
    return s.min().date(), s.max().date()


def _altair_y_scale_positive(s: pd.Series) -> alt.Scale:
    """金額がすべて 0 のときでも棒グラフが潰れないよう Y 軸上限を確保する。"""
    v = pd.to_numeric(s, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    hi = float(v.max())
    if not math.isfinite(hi):
        hi = 0.0
    top = max(hi * 1.08, 1.0)
    return alt.Scale(domain=[0.0, top])


def render_ledger_dashboard(df: pd.DataFrame) -> None:
    """在庫台帳 DataFrame から入出庫集計・仕入先・取引先別・グラフを表示する。"""
    st.subheader("集計")
    st.caption(
        "上の表の現在の内容（未保存の編集を含む）を集計します。"
        f"金額はシートの「{COL_PRICE_EXCL}」「{COL_PRICE_INCL}」列を行合計として集計します。"
        f"仕入先・取引先別の粗利は「{COL_GROSS_PROFIT}」列を合算しています（税抜・台帳保存時の値）。"
        "出庫（販売）は **在庫行の更新** のみのため、入庫／出庫の数量・金額は **在庫中＝仕入**、**販売済＝実売** を二重計上しないよう派生列で計上しています。"
        "期間フィルタと月次の軸は **販売済は販売日時**（未入力の旧行は日時にフォールバック）、在庫中は **日時** です。"
        "（税抜の仕入金額が空で税込だけある行は、10%/8%/非課税のいずれかに税込が一致する税抜を逆算します。"
        "カンマ区切り・円記号付きの数値も読み取ります。）"
    )
    if df.empty:
        st.info("集計する行がありません。")
        return

    df_in = df.copy()
    for _col in (
        COL_QTY,
        COL_PRICE_EXCL,
        COL_PRICE_INCL,
        COL_GROSS_PROFIT,
    ):
        if _col in df_in.columns:
            df_in[_col] = (
                _series_to_numeric_loose(df_in[_col])
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
            )
    ad = _prepare_ledger_analysis(df_in)
    ad_f = ad.dropna(subset=[COL_DATETIME], how="all")

    d_lo, d_hi = _ledger_dashboard_date_bounds(ad_f)
    p1, p2, p3 = st.columns([1, 1, 2])
    with p1:
        date_from = st.date_input(
            "開始日（From）",
            value=d_lo,
            min_value=date(1970, 1, 1),
            max_value=date(2100, 12, 31),
            key="dash_date_from",
        )
    with p2:
        date_to = st.date_input(
            "終了日（To）",
            value=d_hi,
            min_value=date(1970, 1, 1),
            max_value=date(2100, 12, 31),
            key="dash_date_to",
        )
    with p3:
        supplier_filter = st.multiselect(
            "仕入先・取引先で絞り込み（未選択は全件）",
            options=sorted(ad_f[COL_SUPPLIER].fillna("").astype(str).unique().tolist()),
            key="dash_supplier_filter",
        )

    dfb = date_from
    dtb = date_to
    if dfb > dtb:
        st.warning("開始日が終了日より後です。入れ替えて集計します。")
        dfb, dtb = dtb, dfb

    row_ts = _ledger_dashboard_axis_datetime(ad_f)
    row_day = row_ts.dt.normalize()
    from_ts = pd.Timestamp(datetime.combine(dfb, datetime.min.time()))
    to_ts = pd.Timestamp(datetime.combine(dtb, datetime.min.time()))
    flt = ad_f[(row_day >= from_ts) & (row_day <= to_ts)]
    if supplier_filter:
        flt = flt[flt[COL_SUPPLIER].astype(str).isin(supplier_filter)]

    ad_sup = ad_f
    if supplier_filter:
        ad_sup = ad_sup[ad_sup[COL_SUPPLIER].astype(str).isin(supplier_filter)]
    stock_cogs_total = 0
    if COL_STOCK_STATUS in ad_sup.columns:
        _mstk = (
            ad_sup[COL_STOCK_STATUS]
            .astype(str)
            .map(_normalize_stock_status)
            == STATUS_IN_STOCK
        )
        stock_cogs_total = _finite_int(
            _series_to_numeric_loose(ad_sup.loc[_mstk, COL_PRICE_EXCL])
            .fillna(0)
            .sum(),
            0,
        )
    gp_sold_period = 0
    if not flt.empty and COL_STOCK_STATUS in flt.columns:
        _msf = (
            flt[COL_STOCK_STATUS].astype(str).map(_normalize_stock_status)
            == STATUS_SOLD
        )
        gp_sold_period = _finite_int(
            _series_to_numeric_loose(flt.loc[_msf, COL_GROSS_PROFIT])
            .fillna(0)
            .sum(),
            0,
        )
    st.markdown("##### ライフサイクル指標")
    st.caption(
        "在庫総額は **日付は問わず**（仕入先フィルタのみ適用）で在庫中の原価税抜を合算します。"
        "確定粗利は **From〜To と仕入先フィルタ内** の販売済行の粗利列の合計です。"
    )
    _lc1, _lc2 = st.columns(2)
    with _lc1:
        st.metric("在庫総額（在庫中・税抜原価）", f"¥{stock_cogs_total:,}")
    with _lc2:
        st.metric("確定粗利（期間内・販売済・税抜）", f"¥{gp_sold_period:,}")

    if flt.empty:
        st.warning(
            "条件に一致するデータがありません。"
            "From〜To の日付範囲または仕入先・取引先の絞り込みを見直してください。"
        )
        return

    q_in = _finite_int(flt["_qty_in"].sum(), 0)
    q_out = _finite_int(flt["_qty_out"].sum(), 0)
    q_net = q_in - q_out
    ex_in = _finite_int(flt["_amt_ex_in"].sum(), 0)
    ex_out = _finite_int(flt["_amt_ex_out"].sum(), 0)
    ex_net = ex_in - ex_out
    in_in = _finite_int(flt["_amt_in_in"].sum(), 0)
    in_out = _finite_int(flt["_amt_in_out"].sum(), 0)
    in_net = in_in - in_out

    m1, m2, m3 = st.columns(3)
    m1.metric("入庫 合計数量", f"{q_in:,}")
    m2.metric("出庫 合計数量", f"{q_out:,}")
    m3.metric("差し引き 数量（入−出）", f"{q_net:,}")
    m5, m6, m7, m8, m9 = st.columns(5)
    m5.metric("入庫 合計金額（税抜）", f"¥{ex_in:,}")
    m6.metric("出庫 合計金額（税抜）", f"¥{ex_out:,}")
    m7.metric("差し引き 税抜（入−出）", f"¥{ex_net:,}")
    m8.metric("差し引き 税込（入−出）", f"¥{in_net:,}")
    with m9:
        if COL_GROSS_PROFIT in flt.columns:
            gp_tot = _finite_int(
                _series_to_numeric_loose(flt[COL_GROSS_PROFIT]).fillna(0).sum(), 0
            )
            st.metric("粗利合計（税抜）", f"¥{gp_tot:,}")
        else:
            st.metric("粗利合計（税抜）", "—")

    st.markdown("##### 仕入先・取引先別サマリー（税抜金額・数量・粗利）")
    sup_col = "仕入先・取引先"

    def _supplier_grp(part: pd.DataFrame) -> pd.DataFrame:
        if part.empty:
            cols = [
                sup_col,
                "入庫数量",
                "出庫数量",
                "入庫金額税抜",
                "出庫金額税抜",
            ]
            if COL_GROSS_PROFIT in flt.columns:
                cols.append("粗利合計")
            return pd.DataFrame(columns=cols)
        g = part.assign(**{sup_col: part[COL_SUPPLIER].fillna("(未設定)").astype(str)})
        _agg_sup: dict[str, tuple[str, str]] = {
            "入庫数量": ("_qty_in", "sum"),
            "出庫数量": ("_qty_out", "sum"),
            "入庫金額税抜": ("_amt_ex_in", "sum"),
            "出庫金額税抜": ("_amt_ex_out", "sum"),
        }
        if COL_GROSS_PROFIT in g.columns:
            _agg_sup["粗利合計"] = (COL_GROSS_PROFIT, "sum")
        out = g.groupby(sup_col, dropna=False).agg(**_agg_sup).reset_index()
        for _gc in ("入庫数量", "出庫数量", "入庫金額税抜", "出庫金額税抜", "粗利合計"):
            if _gc in out.columns:
                out[_gc] = (
                    pd.to_numeric(out[_gc], errors="coerce")
                    .replace([np.inf, -np.inf], np.nan)
                    .fillna(0.0)
                )
        return out

    grp = _supplier_grp(flt)
    if COL_STOCK_STATUS not in flt.columns:
        st.caption(
            f"「{COL_STOCK_STATUS}」列がないため、在庫中／販売済に分けず期間内の全体を表示します。"
        )
        st.dataframe(
            grp.sort_values(sup_col),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.caption(
            f"「{COL_STOCK_STATUS}」に従い、在庫中と販売済の行をそれぞれ集計した表です。"
            "下の「仕入先・取引先別 税抜金額」「粗利」グラフは従来どおり期間内の全体です。"
        )
        sn = flt[COL_STOCK_STATUS].astype(str).map(_normalize_stock_status)
        flt_in = flt.loc[sn == STATUS_IN_STOCK]
        flt_so = flt.loc[sn == STATUS_SOLD]
        st.markdown("###### 在庫中")
        st.dataframe(
            _supplier_grp(flt_in).sort_values(sup_col),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("###### 販売済")
        st.dataframe(
            _supplier_grp(flt_so).sort_values(sup_col),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("##### 月次推移（数量）")
    st.caption("各月で入庫・出庫を並べた棒グラフ（積み上げではありません）。")
    monthly = (
        flt.dropna(subset=[COL_DATETIME])
        .groupby("_ym", as_index=False)
        .agg(
            入庫数量=("_qty_in", "sum"),
            出庫数量=("_qty_out", "sum"),
            入庫金額税抜=("_amt_ex_in", "sum"),
            出庫金額税抜=("_amt_ex_out", "sum"),
        )
        .sort_values("_ym")
    )
    if not monthly.empty:
        for _mc in (
            "入庫数量",
            "出庫数量",
            "入庫金額税抜",
            "出庫金額税抜",
        ):
            if _mc in monthly.columns:
                monthly[_mc] = (
                    pd.to_numeric(monthly[_mc], errors="coerce")
                    .replace([np.inf, -np.inf], np.nan)
                    .fillna(0.0)
                )
        month_order = monthly["_ym"].astype(str).tolist()
        mdf_qty = monthly.rename(
            columns={"_ym": "月", "入庫数量": "入庫", "出庫数量": "出庫"}
        )
        qty_long = pd.melt(
            mdf_qty,
            id_vars=["月"],
            value_vars=["入庫", "出庫"],
            var_name="区分",
            value_name="数量",
        )
        qty_long["数量"] = pd.to_numeric(
            qty_long["数量"], errors="coerce"
        ).fillna(0.0)
        chart_qty = (
            alt.Chart(qty_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y(
                    "数量:Q",
                    title="数量",
                    axis=alt.Axis(format=",.0f"),
                    scale=_altair_y_scale_positive(qty_long["数量"]),
                ),
                xOffset=alt.XOffset("区分:N"),
                color=alt.Color(
                    "区分:N",
                    scale=alt.Scale(domain=["入庫", "出庫"], range=["#1f77b4", "#ff7f0e"]),
                    legend=alt.Legend(title=None),
                ),
                tooltip=[
                    alt.Tooltip("月:N", title="月"),
                    "区分",
                    alt.Tooltip("数量:Q", title="数量", format=",.0f"),
                ],
            )
            .properties(height=340)
        )
        st.altair_chart(chart_qty, use_container_width=True)
    else:
        st.caption("月次グラフを表示できる日付がありません。")

    st.markdown("##### 月次推移（金額・税抜）")
    st.caption(
        "数量グラフと同じ月次集計で、入庫・出庫の税抜金額を並べた棒グラフです（積み上げではありません）。"
    )
    if not monthly.empty:
        month_order = monthly["_ym"].astype(str).tolist()
        mdf_amt = monthly.rename(
            columns={"_ym": "月", "入庫金額税抜": "入庫", "出庫金額税抜": "出庫"}
        )
        amt_bar_long = pd.melt(
            mdf_amt,
            id_vars=["月"],
            value_vars=["入庫", "出庫"],
            var_name="区分",
            value_name="金額",
        )
        amt_bar_long["金額"] = pd.to_numeric(
            amt_bar_long["金額"], errors="coerce"
        ).fillna(0.0)
        chart_amt_bar = (
            alt.Chart(amt_bar_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y(
                    "金額:Q",
                    title="金額（税抜・円）",
                    axis=alt.Axis(format=",.0f"),
                    scale=_altair_y_scale_positive(amt_bar_long["金額"]),
                ),
                xOffset=alt.XOffset("区分:N"),
                color=alt.Color(
                    "区分:N",
                    scale=alt.Scale(domain=["入庫", "出庫"], range=["#1f77b4", "#ff7f0e"]),
                    legend=alt.Legend(title=None),
                ),
                tooltip=[
                    alt.Tooltip("月:N", title="月"),
                    "区分",
                    alt.Tooltip("金額:Q", title="金額（円）", format=",.0f"),
                ],
            )
            .properties(height=340)
        )
        st.altair_chart(chart_amt_bar, use_container_width=True)
    else:
        st.caption("月次の金額グラフを表示できる日付がありません。")

    st.markdown("##### 金額推移（税抜）")
    st.caption(
        "その月までの入庫・出庫それぞれの税抜金額の**累計**（いわゆる累積）を、"
        "各月で入庫・出庫の2本の棒として並べたグラフです。"
        "上の From〜To・仕入先・取引先の絞り込みに従います。"
    )
    if not monthly.empty:
        month_order_c = monthly["_ym"].astype(str).tolist()
        mc = monthly.sort_values("_ym").copy()
        mc["入庫累積"] = mc["入庫金額税抜"].cumsum()
        mc["出庫累積"] = mc["出庫金額税抜"].cumsum()
        mdf_cum = mc.rename(columns={"_ym": "月", "入庫累積": "入庫", "出庫累積": "出庫"})
        cum_long = pd.melt(
            mdf_cum,
            id_vars=["月"],
            value_vars=["入庫", "出庫"],
            var_name="区分",
            value_name="累積金額",
        )
        cum_long["累積金額"] = pd.to_numeric(
            cum_long["累積金額"], errors="coerce"
        ).fillna(0.0)
        chart_cum_bar = (
            alt.Chart(cum_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order_c, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y(
                    "累積金額:Q",
                    title="累積金額（税抜・円）",
                    axis=alt.Axis(format=",.0f"),
                    scale=_altair_y_scale_positive(cum_long["累積金額"]),
                ),
                xOffset=alt.XOffset("区分:N"),
                color=alt.Color(
                    "区分:N",
                    scale=alt.Scale(domain=["入庫", "出庫"], range=["#1f77b4", "#ff7f0e"]),
                    legend=alt.Legend(title=None),
                ),
                tooltip=[
                    alt.Tooltip("月:N", title="月"),
                    "区分",
                    alt.Tooltip("累積金額:Q", title="累積（円）", format=",.0f"),
                ],
            )
            .properties(height=340)
        )
        st.altair_chart(chart_cum_bar, use_container_width=True)
    else:
        st.caption("累積金額グラフを表示できる月次データがありません。")

    st.markdown("##### 仕入先・取引先別 税抜金額（変動幅の大きい順・上位15件）")
    st.caption("各仕入先・取引先で入庫・出庫の税抜金額を並べた棒グラフです。")
    chart_src = (
        grp.set_index(sup_col)[["入庫金額税抜", "出庫金額税抜"]]
        .assign(_abs=lambda x: (x["入庫金額税抜"] - x["出庫金額税抜"]).abs())
        .sort_values("_abs", ascending=False)
        .drop(columns=["_abs"])
        .head(15)
    )
    if not chart_src.empty:
        top_src = chart_src.reset_index()
        top_order = top_src[sup_col].astype(str).tolist()
        top_long = pd.melt(
            top_src.rename(columns={"入庫金額税抜": "入庫", "出庫金額税抜": "出庫"}),
            id_vars=[sup_col],
            value_vars=["入庫", "出庫"],
            var_name="区分",
            value_name="金額",
        )
        top_long["金額"] = pd.to_numeric(
            top_long["金額"], errors="coerce"
        ).fillna(0.0)
        sup_chart = (
            alt.Chart(top_long)
            .mark_bar()
            .encode(
                x=alt.X(
                    f"{sup_col}:N",
                    sort=top_order,
                    axis=alt.Axis(title=sup_col, labelAngle=-45),
                ),
                y=alt.Y(
                    "金額:Q",
                    title="金額（税抜・円）",
                    axis=alt.Axis(format=",.0f"),
                    scale=_altair_y_scale_positive(top_long["金額"]),
                ),
                xOffset=alt.XOffset("区分:N"),
                color=alt.Color(
                    "区分:N",
                    scale=alt.Scale(domain=["入庫", "出庫"], range=["#1f77b4", "#ff7f0e"]),
                    legend=alt.Legend(title=None),
                ),
                tooltip=[
                    alt.Tooltip(f"{sup_col}:N", title=sup_col),
                    "区分",
                    alt.Tooltip("金額:Q", title="金額（円）", format=",.0f"),
                ],
            )
            .properties(height=380)
        )
        st.altair_chart(sup_chart, use_container_width=True)

    if "粗利合計" in grp.columns:
        st.markdown("##### 仕入先・取引先別 粗利（税抜・上位15件）")
        st.caption(
            f"台帳の「{COL_GROSS_PROFIT}」列を仕入先・取引先ごとに合算しています。"
            "未販売は「在庫中」、販売済は「販売済」の行に限定します。"
            "各グラフの並びは粗利の絶対値が大きい順です（マイナスも含みます）。"
        )

        def _gp_supplier_top15_bar(sub: pd.DataFrame) -> None:
            if sub.empty:
                st.caption("該当する行がありません。")
                return
            gsub = sub.assign(**{sup_col: sub[COL_SUPPLIER].fillna("(未設定)").astype(str)})
            gg = (
                gsub.groupby(sup_col, dropna=False)
                .agg(粗利合計=(COL_GROSS_PROFIT, "sum"))
                .reset_index()
            )
            gg["粗利合計"] = (
                pd.to_numeric(gg["粗利合計"], errors="coerce")
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0.0)
            )
            gp_ch = (
                gg.assign(_abs=lambda x: x["粗利合計"].abs())
                .sort_values("_abs", ascending=False)
                .drop(columns=["_abs"])
                .head(15)
            )
            if gp_ch.empty:
                st.caption("該当するデータがありません。")
                return
            g_order = gp_ch[sup_col].astype(str).tolist()
            gp_bar = (
                alt.Chart(gp_ch)
                .mark_bar()
                .encode(
                    x=alt.X(
                        f"{sup_col}:N",
                        sort=g_order,
                        axis=alt.Axis(title=sup_col, labelAngle=-45),
                    ),
                    y=alt.Y(
                        "粗利合計:Q",
                        title="粗利（税抜・円）",
                        axis=alt.Axis(format=",.0f"),
                    ),
                    tooltip=[
                        alt.Tooltip(f"{sup_col}:N", title=sup_col),
                        alt.Tooltip("粗利合計:Q", title="粗利（円）", format=",.0f"),
                    ],
                )
                .properties(height=380)
            )
            st.altair_chart(gp_bar, use_container_width=True)

        if COL_STOCK_STATUS not in flt.columns:
            st.caption(
                f"「{COL_STOCK_STATUS}」列がないため、未販売／販売済に分けず期間内の全体を表示します。"
            )
            st.markdown("###### 全体")
            _gp_supplier_top15_bar(flt)
        else:
            sn = flt[COL_STOCK_STATUS].astype(str).map(_normalize_stock_status)
            flt_unsold = flt.loc[sn == STATUS_IN_STOCK]
            flt_sold = flt.loc[sn == STATUS_SOLD]
            st.markdown("###### 未販売（在庫中）")
            _gp_supplier_top15_bar(flt_unsold)
            st.markdown("###### 販売済")
            _gp_supplier_top15_bar(flt_sold)


def _render_inventory_price_summary(df: pd.DataFrame) -> None:
    """在庫中の行について、合計原価・販売予定（税抜行計・税込総額）・想定粗利を表示する。"""
    st.markdown("##### 価格管理サマリー（在庫中）")
    st.caption(
        "上の表のうち、ステータスが「在庫中」の行だけを合算しています（未保存の編集を含みます）。"
        "「販売予定金額（税抜）」列の値×数量の税抜行計、税込列は仕入行と同じ税率で算出した値の合計です。"
    )
    if df is None or df.empty:
        return
    if COL_STOCK_STATUS not in df.columns:
        return
    calc = _recalc_gross_profit_dataframe(df.copy())
    mask = calc[COL_STOCK_STATUS].astype(str).str.strip() == STATUS_IN_STOCK
    sub = calc.loc[mask]
    if sub.empty:
        st.info("「在庫中」の行がありません。")
        return
    cg = sub[COL_PRICE_EXCL].map(_int_from_cell)
    pl_u = sub[COL_PLANNED_SALE].map(_int_from_cell)
    qv = sub[COL_QTY].map(lambda x: max(1, _int_from_cell(x)))
    pl_line_ex = pl_u * qv
    pl_in = sub[COL_PLANNED_SALE_INCL].map(_int_from_cell)
    total_cogs = int(cg.sum())
    total_planned_excl = int(pl_line_ex.sum())
    total_planned_incl = int(pl_in.sum())
    m = pl_line_ex > 0
    total_margin = int((pl_line_ex.loc[m] - cg.loc[m]).sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("合計原価（税抜）", f"¥{total_cogs:,}")
    m2.metric("合計販売予定（税抜・行計）", f"¥{total_planned_excl:,}")
    m3.metric("合計販売予定（税込）", f"¥{total_planned_incl:,}")
    m4.metric("想定粗利（税抜・合計）", f"¥{total_margin:,}")


def _product_keyword_category(name: str) -> str:
    """ダッシュボード用の簡易カテゴリ（商品名のキーワードから推定）。"""
    s = str(name or "")
    if re.search("振袖", s):
        return "振袖"
    if re.search("訪問着", s):
        return "訪問着"
    if re.search("帯", s):
        return "帯"
    if re.search("長襦袢|襦袢", s):
        return "長襦袢・襦袢"
    return "その他"


def _render_inventory_category_pie_altair(pie_df: pd.DataFrame) -> None:
    """Plotly 未導入時のカテゴリ構成比（円グラフ相当）。"""
    chart = (
        alt.Chart(pie_df)
        .mark_arc(innerRadius=55)
        .encode(
            theta=alt.Theta("金額税抜:Q", stack=True),
            color=alt.Color("カテゴリー:N"),
            tooltip=[
                alt.Tooltip("カテゴリー:N"),
                alt.Tooltip("金額税抜:Q", title="金額（税抜）", format=","),
            ],
        )
        .properties(
            title="在庫中の原価シェア（簡易カテゴリ・Altair）",
            height=400,
        )
    )
    st.altair_chart(chart, use_container_width=True)


def _render_inventory_category_pie(pie_df: pd.DataFrame) -> None:
    """plotly が入っていれば ``st.plotly_chart``、なければ Altair にフォールバック。"""
    if float(pie_df["金額税抜"].sum()) <= 0:
        st.caption("在庫中で原価が入っている行がありません。")
        return
    try:
        import plotly.express as px
    except ImportError:
        _render_inventory_category_pie_altair(pie_df)
        return
    fig = px.pie(
        pie_df,
        names="カテゴリー",
        values="金額税抜",
        hole=0.35,
        title="在庫中の原価シェア（簡易カテゴリ）",
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig, use_container_width=True)


def render_analytics_dashboard_page() -> None:
    """集計・分析: メトリクス・Plotly・既存の月次ダッシュボード。"""
    st.subheader("分析")
    st.caption(
        "共有の **inventory.csv**（`INVENTORY_SOURCE=csv`）または **Google スプレッドシート**から読み込んだ最新データを集計します。"
    )
    if not _uses_local_inventory_csv() and not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        st.info(
            "`GOOGLE_SPREADSHEET_ID` を設定するか、`INVENTORY_SOURCE = \"csv\"` で CSV を使ってください。"
        )
        return
    try:
        df_raw = load_inventory_dataframe()
    except Exception as e:
        st.error(f"読み込みに失敗しました: {e}")
        return
    if df_raw is None:
        st.warning("台帳を開けませんでした。サービスアカウントと共有設定を確認してください。")
        return
    if df_raw.empty:
        st.info("集計する行がありません。")
        return

    calc = _recalc_gross_profit_dataframe(df_raw.copy())
    if COL_STOCK_STATUS not in calc.columns:
        st.warning("ステータス列がないため在庫中の集計ができません。")
        render_ledger_dashboard(calc)
        return

    mask_stock = calc[COL_STOCK_STATUS].astype(str).str.strip() == STATUS_IN_STOCK
    sub = calc.loc[mask_stock].copy()
    cg = _series_to_numeric_loose(sub[COL_PRICE_EXCL]).fillna(0).clip(lower=0)
    total_inv = int(cg.sum())
    n_lines = int(len(sub))
    gp_s = _series_to_numeric_loose(sub[COL_GROSS_PROFIT]).fillna(0)
    ratios: list[float] = []
    for i in range(len(sub)):
        c = float(cg.iloc[i])
        g = float(gp_s.iloc[i])
        if c > 0:
            ratios.append(100.0 * g / c)
    avg_ratio = float(np.mean(ratios)) if ratios else None

    k1, k2, k3 = st.columns(3)
    k1.metric("在庫中 総額（仕入・税抜）", f"¥{total_inv:,}")
    k2.metric("在庫中 行数（点数）", f"{n_lines:,}")
    k3.metric(
        "在庫中 平均粗利率（粗利÷原価）",
        f"{avg_ratio:.1f} %" if avg_ratio is not None else "—",
    )

    st.markdown("##### カテゴリー別 在庫原価（税抜）の構成比")
    st.caption("商品名に含まれるキーワードで振り分けたうえで、在庫中の仕入金額（税抜）を合算しています。")
    sub["_category"] = sub[COL_NAME].astype(str).map(_product_keyword_category)
    sub["_px"] = _series_to_numeric_loose(sub[COL_PRICE_EXCL]).fillna(0)
    pie_df = sub.groupby("_category", dropna=False)["_px"].sum().reset_index()
    pie_df.columns = ["カテゴリー", "金額税抜"]
    _render_inventory_category_pie(pie_df)

    st.divider()
    render_ledger_dashboard(calc)


def _inject_prominent_main_tabs_style() -> None:
    """メインエリアの ``st.tabs`` ラベルを大きく太字にする（登録・在庫一覧で共通）。"""
    st.markdown(
        """
        <style>
        section.main [data-testid="stTabs"] button {
            font-size: clamp(1.1rem, 2.2vw, 1.45rem) !important;
            font-weight: 700 !important;
            line-height: 1.35 !important;
            padding-top: 0.65rem !important;
            padding-bottom: 0.65rem !important;
            letter-spacing: 0.02em !important;
        }
        section.main [data-testid="stTabs"] [role="tablist"] {
            gap: 0.35rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _google_drive_file_id_from_url(url: str) -> str | None:
    u = (url or "").strip()
    if not u:
        return None
    m = re.search(r"/file/d/([a-zA-Z0-9_-]{10,})", u, re.I)
    if m:
        return m.group(1)
    m2 = re.search(r"[?&]id=([a-zA-Z0-9_-]{10,})", u, re.I)
    if m2:
        return m2.group(1)
    m3 = re.search(r"googleusercontent\.com/d/([a-zA-Z0-9_-]{10,})", u, re.I)
    if m3:
        return m3.group(1)
    return None


def _render_inventory_gallery_thumbnail(image_url: str, *, width: int, sold: bool) -> None:
    """ギャラリー用。Drive 直リンクは ``st.image(URL)`` が効かないことが多いため、取得して JPEG 化して表示する。"""
    iu = (image_url or "").strip()
    if not (iu.startswith("http://") or iu.startswith("https://")):
        st.caption("（画像なし）")
        return
    _w = max(120, min(480, int(width)))
    sz = min(1200, _w * 4)
    fid = _google_drive_file_id_from_url(iu)
    candidates: list[str] = []
    if fid:
        candidates.append(f"https://drive.google.com/thumbnail?id={fid}&sz=w{sz}")
        candidates.append(
            f"https://drive.usercontent.google.com/download?id={fid}&export=view"
        )
    candidates.append(iu)

    ua = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
    }
    timeout = 20.0
    for cand in candidates:
        try:
            r = requests.get(cand, timeout=timeout, headers=ua, allow_redirects=True)
        except Exception:
            continue
        if r.status_code != 200 or not r.content or len(r.content) < 32:
            continue
        low512 = r.content[:512].lower()
        if b"<html" in low512 or b"<!doctype" in low512:
            continue
        try:
            im = Image.open(io.BytesIO(r.content))
            im = ImageOps.exif_transpose(im)
            im.thumbnail((_w * 2, _w * 2), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            im.convert("RGB").save(buf, format="JPEG", quality=85, optimize=True)
            st.image(buf.getvalue(), width=_w, use_container_width=False)
            return
        except Exception:
            try:
                st.image(r.content, width=_w, use_container_width=False)
                return
            except Exception:
                continue
    st.caption("（ブラウザで開くと表示できる画像です）")
    st.link_button("画像を開く", iu, use_container_width=True)


def _render_inventory_ledger_data_editor_section(df_sorted: pd.DataFrame) -> None:
    """在庫一覧の表形式: data_editor と保存ボタン。"""
    _ledger_base_for_save = df_sorted
    df_sorted = df_sorted.copy()
    _ledger_evidence_link_mode = _normalize_evidence_urls_for_link_editor(
        df_sorted, COL_VOUCHER_EVIDENCE_URL
    )

    _ledger_col_cfg: dict[str, Any] = {}
    if COL_MANAGEMENT_ID in df_sorted.columns:
        _ledger_col_cfg[COL_MANAGEMENT_ID] = st.column_config.TextColumn(
            COL_MANAGEMENT_ID,
            disabled=True,
            help="1点1行の自動採番（シリアル）。通常は手入力しません。",
        )
    if COL_STOCK_STATUS in df_sorted.columns:
        _ledger_col_cfg[COL_STOCK_STATUS] = st.column_config.SelectboxColumn(
            COL_STOCK_STATUS,
            options=list(STOCK_STATUS_OPTIONS),
            help="在庫中＝未販売想定、販売済＝実売価格で粗利を計算します。",
        )
    if COL_LAST_STOCKTAKE in df_sorted.columns:
        _ledger_col_cfg[COL_LAST_STOCKTAKE] = st.column_config.TextColumn(
            COL_LAST_STOCKTAKE,
            help=f"棚卸・実地確認した日（日本時間の暦日推奨）。例: {_today_jst_date().isoformat()}",
        )
    if COL_SALE_SOURCE_MGMT_ID in df_sorted.columns:
        _ledger_col_cfg[COL_SALE_SOURCE_MGMT_ID] = st.column_config.TextColumn(
            COL_SALE_SOURCE_MGMT_ID,
            help="出庫（販売）などで売れた在庫行の入庫時管理ID（G########）。登録画面の販売管理からも入力できます。",
        )
    if COL_VOUCHER_RECORDED_AT in df_sorted.columns:
        _ledger_col_cfg[COL_VOUCHER_RECORDED_AT] = st.column_config.TextColumn(
            COL_VOUCHER_RECORDED_AT,
            help="証憑を台帳に反映した日時（JST・recorded_at に相当）。",
        )
    if COL_VOUCHER_EVIDENCE_URL in df_sorted.columns:
        if _ledger_evidence_link_mode:
            _ledger_col_cfg[COL_VOUCHER_EVIDENCE_URL] = st.column_config.LinkColumn(
                COL_VOUCHER_EVIDENCE_URL,
                help="Google ドライブの証憑 URL。アイコンをクリックするとブラウザで開きます。",
                disabled=True,
                display_text=":material/open_in_new:",
            )
        else:
            _ledger_col_cfg[COL_VOUCHER_EVIDENCE_URL] = st.column_config.TextColumn(
                COL_VOUCHER_EVIDENCE_URL,
                help="http(s) 以外の文字が混ざる行があるためテキスト表示です。URL をコピーしてブラウザで開いてください。",
            )

    _editor_kw: dict[str, Any] = {
        "num_rows": "dynamic",
        "key": LEDGER_DATA_EDITOR_KEY,
        "use_container_width": True,
        "hide_index": True,
        "height": 720,
    }
    if _ledger_col_cfg:
        _editor_kw["column_config"] = _ledger_col_cfg
    edited = st.data_editor(df_sorted, **_editor_kw)

    if st.button(
        "在庫中かつ棚卸日が未入力の行に、今日の日付（JST）を一括入力してすぐ保存",
        key="ledger_bulk_stocktake_today_save",
    ):
        ed_bulk = edited.copy()
        if COL_LAST_STOCKTAKE not in ed_bulk.columns:
            ed_bulk[COL_LAST_STOCKTAKE] = ""
        m_b = _mask_ledger_stocktake_unverified(ed_bulk)
        ed_bulk.loc[m_b, COL_LAST_STOCKTAKE] = _today_jst_date().isoformat()
        with st.spinner("台帳を保存しています…"):
            try:
                overwrite_inventory_worksheet_from_dataframe(
                    ed_bulk.reset_index(drop=True),
                    previous_df=edited.reset_index(drop=True),
                )
            except Exception as e:
                st.error(str(e))
            else:
                st.session_state["_ledger_saved_flash"] = (
                    "棚卸日（今日・JST）を未確認の在庫中に一括入力し、台帳を保存しました。"
                )
                st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                st.rerun()

    _render_inventory_price_summary(edited)

    if st.button("台帳を更新する", type="primary", key="ledger_save_overwrite"):
        with st.spinner("台帳を保存しています…"):
            try:
                overwrite_inventory_worksheet_from_dataframe(
                    edited.reset_index(drop=True),
                    previous_df=_ledger_base_for_save.reset_index(drop=True),
                )
            except Exception as e:
                st.error(str(e))
                return
        st.session_state["_ledger_saved_flash"] = "台帳を更新しました。"
        st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
        st.rerun()


def render_inventory_list_page() -> None:
    st.markdown("## 在庫一覧")
    st.caption(
        "共有の **inventory.csv** または **スプレッドシート**の全データを編集できます。行の追加・削除は表から操作し、"
        "「台帳を更新する」で保存します。"
        "「日時」は **保存時にセル内容が変わった行**（および表で追加した新規行）で **JST の現在時刻** に自動更新されます。"
        "「証憑記録日時」は証憑取込の確定時刻、「証憑URL」は Drive 上の証憑です（"
        "台帳内の値がすべて空または http(s) のときはリンク列として表示されクリックで開けます。"
        "http 以外の文字が混ざる行がある場合はテキスト列のままです）。"
        "棚卸し用の「最後に確認した日付（棚卸日）」は **YYYY-MM-DD** 推奨です（例: 今日なら "
        f"{_today_jst_date().isoformat()}）。"
        "既定は **ギャラリー（カタログ）** タブです。**在庫一覧** タブの表は **全行** を表示します（未確認だけの一覧は展開パネルで参照できます）。"
    )

    if msg := st.session_state.pop("_ledger_saved_flash", None):
        st.success(msg)

    if not _uses_local_inventory_csv() and not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        st.info(
            f"{SECRET_GOOGLE_SPREADSHEET_ID} を設定するか、`INVENTORY_SOURCE = \"csv\"` で "
            "共有の inventory.csv を有効にしてください。"
        )
        return

    r1, _ = st.columns([1, 2])
    with r1:
        _reload_label = (
            "inventory.csv から再読込"
            if _uses_local_inventory_csv()
            else "スプレッドシートから再読込"
        )
        if st.button(_reload_label, key="ledger_reload_from_sheet"):
            try:
                _inventory_csv_read_df_cached.clear()
            except Exception:
                pass
            try:
                _inventory_sheet_get_all_values_cached.clear()
            except Exception:
                pass
            _bump_inventory_sheet_cache_bust()
            st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
            st.rerun()

    try:
        df_sheet = load_inventory_dataframe()
    except Exception as e:
        st.error(f"読み込みに失敗しました: {e}")
        return

    if df_sheet is None:
        st.warning("台帳を開けませんでした。サービスアカウントと共有設定を確認してください。")
        return

    n_in_stock = int(_mask_ledger_in_stock(df_sheet).sum())
    n_unverified = int(_mask_ledger_stocktake_unverified(df_sheet).sum())
    n_today_done = int(_mask_ledger_stocktake_today_jst(df_sheet).sum())
    sk1, sk2, sk3, sk4 = st.columns(4)
    sk1.metric("在庫中（件数）", f"{n_in_stock:,}")
    sk2.metric("棚卸・未確認（在庫中）", f"{n_unverified:,}")
    sk3.metric("今日確認済（在庫中・JST）", f"{n_today_done:,}")
    with sk4:
        if n_in_stock > 0:
            st.caption("残り（未確認 / 在庫中）")
            st.markdown(f"**{n_unverified} / {n_in_stock}**")
        else:
            st.caption("在庫中の行がないため比率は出ません。")

    with st.expander("棚卸し: 在庫中かつ未確認の一覧（参照のみ）", expanded=False):
        st.caption(
            "「最後に確認した日付（棚卸日）」が空、または日付として解釈できない **在庫中** だけを表示します。"
            "1人作業時の残件数の把握用です。日付の入力・保存は下の表で行ってください。"
        )
        unv = df_sheet.loc[_mask_ledger_stocktake_unverified(df_sheet)].copy()
        if unv.empty:
            st.success("在庫中で、かつ棚卸日が未入力の行はありません。")
        else:
            st.metric("この一覧の件数（＝未確認の在庫中）", f"{len(unv):,}")
            _ucols = [
                c
                for c in (
                    COL_MANAGEMENT_ID,
                    COL_NAME,
                    COL_SUPPLIER,
                    COL_DATETIME,
                    COL_LAST_STOCKTAKE,
                )
                if c in unv.columns
            ]
            st.dataframe(unv[_ucols], use_container_width=True, hide_index=True)

    if n_today_done > 0 and COL_MANAGEMENT_ID in df_sheet.columns:
        _td_rows = df_sheet.loc[_mask_ledger_stocktake_today_jst(df_sheet)]
        _ids_show = (
            _td_rows[COL_MANAGEMENT_ID].astype(str).str.strip().head(18).tolist()
        )
        tail = " …" if len(_td_rows) > len(_ids_show) else ""
        st.caption(
            f"今日（JST {_today_jst_date().isoformat()}）の棚卸日が入っている在庫中: **{n_today_done}** 件。"
            f"管理IDの例: {', '.join(_ids_show)}{tail}"
        )

    st.markdown("##### 表示の並び順（台帳表に反映・保存時もこの順で書き込みます）")
    s1, s2, s3, s4, s5, s6 = st.columns([2, 1, 2, 1, 2, 1])
    sort_choices = ["日時", "仕入先・取引先", "管理ID", "仕入日時", "販売日時", "なし"]
    with s1:
        prim = st.selectbox("第1ソート", sort_choices, index=5, key="ledger_sort_p")
    with s2:
        prim_ord = st.radio("第1の順序", ["昇順", "降順"], horizontal=True, key="ledger_sort_p_ord")
    with s3:
        sec = st.selectbox("第2ソート", sort_choices, index=5, key="ledger_sort_s")
    with s4:
        sec_ord = st.radio("第2の順序", ["昇順", "降順"], horizontal=True, key="ledger_sort_s_ord")
    with s5:
        ter = st.selectbox("第3ソート", sort_choices, index=5, key="ledger_sort_t")
    with s6:
        ter_ord = st.radio("第3の順序", ["昇順", "降順"], horizontal=True, key="ledger_sort_t_ord")

    df_sorted = _apply_ledger_sort(
        df_sheet,
        prim,
        prim_ord == "昇順",
        sec,
        sec_ord == "昇順",
        ter,
        ter_ord == "昇順",
    )

    df_sorted_calc = _recalc_gross_profit_dataframe(df_sorted.copy())

    _inject_prominent_main_tabs_style()
    st.markdown("## 表示形式")
    st.caption(
        "**上の大きなタブ** で切り替えます（先頭タブが既定表示）。**ギャラリー（カタログ）** は接客用プレビュー（"
        f"1ページ **{INV_GALLERY_PAGE_SIZE}** 件・ページ切替で全件）、**在庫一覧** は全行の表で編集・保存します。"
    )
    tab_ledger_gallery, tab_ledger_table = st.tabs(
        ("ギャラリー（カタログ）", "在庫一覧"),
    )

    with tab_ledger_gallery:
        st.markdown("### ギャラリー（カタログ）")
        st.caption(
            "Google ドライブの画像 URL はサーバー側で取得して表示します（表示できない場合は **画像を開く** からブラウザで確認できます）。"
        )
        g1, g2, g3 = st.columns([2, 2, 1])
        with g1:
            st.text_input(
                "フリーワード（商品名・管理ID・メモ）",
                key="inv_gallery_search_text",
                placeholder="部分一致で検索",
            )
        with g2:
            sup_opts: list[str] = []
            if COL_SUPPLIER in df_sorted_calc.columns:
                sup_opts = sorted(
                    {
                        x
                        for x in df_sorted_calc[COL_SUPPLIER]
                        .astype(str)
                        .str.strip()
                        .tolist()
                        if x
                    }
                )
            st.multiselect(
                "仕入先で絞り込み（複数可）",
                options=sup_opts,
                key="inv_gallery_suppliers_filter",
            )
        with g3:
            st.selectbox(
                "ステータス",
                ("すべて", "在庫中", "販売済"),
                key="inv_gallery_status_filter",
            )

        _fw = str(st.session_state.get("inv_gallery_search_text", "") or "")
        _sup_f = list(st.session_state.get("inv_gallery_suppliers_filter") or [])
        _st_f = str(st.session_state.get("inv_gallery_status_filter", "すべて") or "すべて")
        df_view = _filter_inventory_df_for_view(
            df_sorted_calc,
            q=_fw,
            suppliers=_sup_f,
            status_mode=_st_f,
        )
        n_total = len(df_view)
        st.caption(
            f"該当 **{n_total:,}** 行（台帳全体 {len(df_sorted_calc):,} 行・粗利は再計算済み）。"
            f"表示は **{INV_GALLERY_PAGE_SIZE}** 件ずつです。"
        )

        if "inv_gallery_page" not in st.session_state:
            st.session_state.inv_gallery_page = 0
        _fp_gal = f"{_fw!r}|{repr(_sup_f)}|{_st_f!r}"
        if st.session_state.get("_inv_gallery_filter_fp") != _fp_gal:
            st.session_state._inv_gallery_filter_fp = _fp_gal
            st.session_state.inv_gallery_page = 0

        n_pages = max(1, (n_total + INV_GALLERY_PAGE_SIZE - 1) // INV_GALLERY_PAGE_SIZE)
        page_idx = int(st.session_state.inv_gallery_page)
        if page_idx >= n_pages:
            page_idx = n_pages - 1
            st.session_state.inv_gallery_page = page_idx
        if page_idx < 0:
            page_idx = 0
            st.session_state.inv_gallery_page = 0

        start_idx = page_idx * INV_GALLERY_PAGE_SIZE
        end_idx = min(n_total, start_idx + INV_GALLERY_PAGE_SIZE)
        df_tiles = df_view.iloc[start_idx:end_idx].reset_index(drop=True)

        if n_pages > 1:
            p1, p2, p3 = st.columns([1, 3, 1])
            with p1:
                if st.button(
                    "◀ 前のページ",
                    disabled=page_idx <= 0,
                    key="inv_gallery_prev_page",
                ):
                    st.session_state.inv_gallery_page = max(0, page_idx - 1)
                    st.rerun()
            with p2:
                st.markdown(
                    f"**ページ {page_idx + 1} / {n_pages}**　"
                    f"（{n_total:,} 件中 **{start_idx + 1}〜{end_idx}** 件を表示）"
                )
            with p3:
                if st.button(
                    "次のページ ▶",
                    disabled=page_idx >= n_pages - 1,
                    key="inv_gallery_next_page",
                ):
                    st.session_state.inv_gallery_page = min(n_pages - 1, page_idx + 1)
                    st.rerun()

        ncols = 4
        for i in range(0, len(df_tiles), ncols):
            gc = st.columns(ncols)
            for j in range(ncols):
                ridx = i + j
                if ridx >= len(df_tiles):
                    break
                row = df_tiles.iloc[ridx]
                sold = (
                    _normalize_stock_status(str(row.get(COL_STOCK_STATUS, "")))
                    == STATUS_SOLD
                )
                mid = str(row.get(COL_MANAGEMENT_ID, "") or "").strip() or f"_{ridx}"
                safe_key = re.sub(r"[^\w\-]", "_", mid)[:48]
                with gc[j]:
                    with st.container(border=True):
                        if sold:
                            st.caption("販売済")
                        iu = str(row.get(COL_IMAGE_URL, "") or "").strip()
                        _img_w = 200 if sold else 240
                        _render_inventory_gallery_thumbnail(
                            iu, width=_img_w, sold=sold
                        )
                        st.markdown(
                            f'<p style="opacity:{"0.55" if sold else "1"};margin:0.2rem 0 0 0;font-size:1.05rem;">'
                            f"<b>{mid}</b></p>",
                            unsafe_allow_html=True,
                        )
                        nm = str(row.get(COL_NAME, "") or "").strip() or "—"
                        st.markdown(
                            f'<p style="opacity:{"0.55" if sold else "1"};margin:0;font-size:0.98rem;">'
                            f"{(nm if len(nm) <= 96 else nm[:93] + '…')}</p>",
                            unsafe_allow_html=True,
                        )
                        ps_raw = row.get(COL_PLANNED_SALE, "")
                        try:
                            psv = int(float(ps_raw)) if str(ps_raw).strip() != "" else 0
                        except (TypeError, ValueError):
                            psv = 0
                        _pl_lbl = (
                            f"販売予定（税抜） ¥{psv:,}"
                            if psv > 0
                            else "販売予定（税抜） —"
                        )
                        st.markdown(
                            f'<p style="opacity:{"0.55" if sold else "1"};margin:0;font-size:0.95rem;">'
                            f"{_pl_lbl}</p>",
                            unsafe_allow_html=True,
                        )
                        rd = {str(c): row.get(c) for c in EXPECTED_HEADERS if c in row.index}
                        if st.button(
                            "詳細",
                            key=f"inv_gal_dlg_{ridx}_{safe_key}",
                            use_container_width=True,
                        ):
                            _inventory_gallery_detail_dialog(rd)

    with tab_ledger_table:
        st.markdown("### 在庫一覧")
        st.caption(
            "全列・全行を表示します。行数が多いときは表の **縦スクロール** で移動してください（保存はこのタブから）。"
        )
        _render_inventory_ledger_data_editor_section(df_sorted_calc)


def _init_registration_form_session_state() -> None:
    """登録フォーム用の session_state 初期値（キーはウィジェットと連動）。"""
    if "field_product_name" not in st.session_state:
        st.session_state.field_product_name = ""
    if "field_supplier" not in st.session_state:
        st.session_state.field_supplier = ""
    if "field_qty" not in st.session_state:
        st.session_state.field_qty = 1
    if REGISTRATION_QTY_WIDGET_KEY not in st.session_state:
        st.session_state[REGISTRATION_QTY_WIDGET_KEY] = int(
            st.session_state.get("field_qty", 1)
        )
    if "ai_kind" not in st.session_state:
        st.session_state.ai_kind = ""
    if "ai_features" not in st.session_state:
        st.session_state.ai_features = ""
    if "ai_parse_ran" not in st.session_state:
        st.session_state.ai_parse_ran = False
    if "field_memo" not in st.session_state:
        st.session_state.field_memo = ""
    if "field_line_excl_yen" not in st.session_state:
        st.session_state.field_line_excl_yen = 1
    st.session_state.pop("field_unit_price_excl", None)
    if "field_consumption_tax_choice" not in st.session_state:
        st.session_state.field_consumption_tax_choice = "10%"
    if "field_planned_sale_excl" not in st.session_state:
        st.session_state.field_planned_sale_excl = 0
    if "field_actual_sale_excl" not in st.session_state:
        st.session_state.field_actual_sale_excl = 0
    if "field_stock_status" not in st.session_state:
        st.session_state.field_stock_status = STATUS_IN_STOCK
    if "hint_filter_product_name" not in st.session_state:
        st.session_state.hint_filter_product_name = ""
    if "hint_filter_supplier" not in st.session_state:
        st.session_state.hint_filter_supplier = ""
    if "ledger_pick_product_name" not in st.session_state:
        st.session_state.ledger_pick_product_name = LEDGER_PICK_PLACEHOLDER
    if "ledger_pick_supplier" not in st.session_state:
        st.session_state.ledger_pick_supplier = LEDGER_PICK_PLACEHOLDER
    if "field_sale_source_mgmt_id" not in st.session_state:
        st.session_state.field_sale_source_mgmt_id = ""
    if "sale_pick_source_id" not in st.session_state:
        st.session_state.sale_pick_source_id = LEDGER_PICK_PLACEHOLDER
    if "s_reg_qty" not in st.session_state:
        st.session_state.s_reg_qty = 1
    if "s_field_sale_source_mgmt_id" not in st.session_state:
        st.session_state.s_field_sale_source_mgmt_id = ""
    if "s_field_actual_sale_excl" not in st.session_state:
        st.session_state.s_field_actual_sale_excl = 0
    if "s_field_memo" not in st.session_state:
        st.session_state.s_field_memo = ""
    if SALES_TAB_QTY_WIDGET_KEY not in st.session_state:
        st.session_state[SALES_TAB_QTY_WIDGET_KEY] = 1
    if "sales_tab_memo" not in st.session_state:
        st.session_state.sales_tab_memo = ""
    if "sales_tab_loan_datetime_manual" not in st.session_state:
        st.session_state.sales_tab_loan_datetime_manual = ""
    st.session_state.pop("field_price_excl", None)


def render_stocktake_scan_tab(df_ledger_hint: pd.DataFrame | None) -> None:
    """棚卸しスキャン: カメラ撮影 → AI 照合 → 棚卸日の確定更新のみ。"""
    st.markdown("##### 棚卸しスキャン（AI 照合）")
    st.caption(
        "現物を撮影し、在庫中の台帳行と照合します。候補が表示されたら内容を確認し、"
        "**棚卸を確定** でその行の「最後に確認した日付（棚卸日）」だけを **本日（JST）** に更新します（新規行は追加しません）。"
    )
    cam = st.camera_input("現物を撮影", key="stocktake_camera_input")
    if st.button("AIで台帳と照合", type="primary", key="stocktake_ai_match_btn"):
        st.session_state.pop("_stocktake_scan_result", None)
        st.session_state.pop("_stocktake_scan_warn", None)
        if cam is None:
            st.session_state["_stocktake_scan_warn"] = "先にカメラで撮影してください。"
        elif df_ledger_hint is None or df_ledger_hint.empty:
            st.session_state["_stocktake_scan_warn"] = "台帳を読み込めないため照合できません。"
        else:
            with st.spinner("画像を解析して台帳と照合しています…"):
                try:
                    inv_ctx = _build_gemini_inventory_context(df_ledger_hint)
                    img_b = cam.getvalue()

                    class _CamBytes:
                        __slots__ = ("_b",)

                        def __init__(self, b: bytes) -> None:
                            self._b = b

                        def getvalue(self) -> bytes:
                            return self._b

                    img_pil = _gemini_input_image_from_upload(_CamBytes(img_b))
                    raw = analyze_image_with_gemini(
                        img_pil,
                        inventory_context=inv_ctx or None,
                        prompt_mode="stocktake_match",
                    )
                    res = _parse_json_from_model(raw or "")
                    m = res.get("match") if isinstance(res, dict) else None
                    mid = ""
                    if isinstance(m, dict):
                        mid = str(m.get("management_id") or "").strip()
                    conf = float(m.get("confidence") or 0) if isinstance(m, dict) else 0.0
                    if not mid or conf < 0.36:
                        st.session_state["_stocktake_scan_warn"] = (
                            "在庫中の行と確実に一致する候補が得られませんでした。明るさ・距離を変えて再撮影するか、在庫一覧で管理IDを確認してください。"
                        )
                    else:
                        tr = lookup_ledger_row_by_management_id(df_ledger_hint, mid)
                        if tr is None:
                            st.session_state["_stocktake_scan_warn"] = (
                                f"管理ID **{mid}** が台帳に見つかりません。"
                            )
                        elif (
                            _normalize_stock_status(str(tr.get(COL_STOCK_STATUS, "")))
                            != STATUS_IN_STOCK
                        ):
                            st.session_state["_stocktake_scan_warn"] = (
                                f"管理ID **{mid}** は在庫中ではありません。"
                            )
                        else:
                            st.session_state["_stocktake_scan_result"] = {
                                "management_id": mid,
                                "product_name": str(tr.get(COL_NAME, "") or "").strip(),
                                "supplier": str(tr.get(COL_SUPPLIER, "") or "").strip(),
                                "last_stocktake": str(
                                    tr.get(COL_LAST_STOCKTAKE, "") or ""
                                ).strip(),
                                "image_url": str(tr.get(COL_IMAGE_URL, "") or "").strip(),
                                "confidence": conf,
                            }
                except Exception as e:
                    st.session_state["_stocktake_scan_warn"] = str(e)
        st.rerun()

    wn = st.session_state.pop("_stocktake_scan_warn", None)
    if wn:
        st.warning(wn)
    hit = st.session_state.get("_stocktake_scan_result")
    if isinstance(hit, dict) and hit.get("management_id"):
        mid = str(hit["management_id"])
        with st.container(border=True):
            st.markdown(f"### 照合結果: **{mid}**")
            c1, c2 = st.columns([1, 2])
            with c1:
                iu = str(hit.get("image_url") or "").strip()
                if iu.startswith("http://") or iu.startswith("https://"):
                    st.image(iu, use_container_width=True)
                else:
                    st.caption("（台帳に画像URLがありません）")
            with c2:
                st.write(f"**商品名:** {hit.get('product_name') or '—'}")
                st.write(f"**仕入先:** {hit.get('supplier') or '—'}")
                st.write(
                    f"**前回の棚卸日:** {hit.get('last_stocktake') or '—（未入力）'}"
                )
                st.caption(
                    f"AI 確信度: {float(hit.get('confidence') or 0):.2f}（参考）"
                )
            if st.button(
                "棚卸を確定（棚卸日を本日・JST に更新）",
                type="primary",
                key=f"stocktake_confirm_{mid}",
            ):
                try:
                    with st.spinner("台帳を更新しています…"):
                        apply_last_stocktake_jst_for_management_id(mid)
                except Exception as e:
                    st.error(str(e))
                else:
                    st.session_state.pop("_stocktake_scan_result", None)
                    st.success(f"管理ID **{mid}** の棚卸日を更新しました。")
                    st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                    st.rerun()


def _filter_inventory_df_for_view(
    df: pd.DataFrame,
    *,
    q: str,
    suppliers: list[str],
    status_mode: str,
) -> pd.DataFrame:
    """在庫一覧の検索・フィルタ（ギャラリー／表の共通ビュー用）。"""
    out = df.copy()
    if status_mode == "在庫中":
        out = out.loc[_mask_ledger_in_stock(out)]
    elif status_mode == "販売済":
        if COL_STOCK_STATUS in out.columns:
            out = out.loc[
                out[COL_STOCK_STATUS].astype(str).str.strip().map(_normalize_stock_status)
                == STATUS_SOLD
            ]
    if suppliers and COL_SUPPLIER in out.columns:
        sup_m = out[COL_SUPPLIER].astype(str).str.strip().isin(set(suppliers))
        out = out.loc[sup_m]
    qt = (q or "").strip()
    if qt and not out.empty:
        qf = qt.casefold()
        m = pd.Series(False, index=out.index)
        for col in (COL_NAME, COL_MANAGEMENT_ID, COL_MEMO):
            if col in out.columns:
                m = m | out[col].astype(str).str.casefold().str.contains(
                    qf, na=False, regex=False
                )
        out = out.loc[m]
    return out.reset_index(drop=True)


@st.dialog("在庫の詳細")
def _inventory_gallery_detail_dialog(row_dict: dict[str, Any]) -> None:
    """ギャラリーから開く行の全列表示。"""
    for k in EXPECTED_HEADERS:
        if k not in row_dict:
            continue
        v = row_dict.get(k)
        if k == COL_IMAGE_URL and str(v or "").strip().startswith("http"):
            st.markdown(f"**{k}**")
            st.image(str(v).strip(), use_container_width=True)
        else:
            st.markdown(f"**{k}**")
            st.write(str(v) if v is not None and str(v).strip() != "" else "—")


def _render_sales_management_tab(
    uploaded,
    df_ledger_hint: pd.DataFrame | None,
) -> None:
    """販売管理タブ: 出庫（販売）／出庫（浮貸）と管理ID・実売または浮貸日時の更新。"""
    st.markdown("##### 販売管理")
    st.caption(
        "**出庫（販売）** … 在庫行を **販売済** にし、実売と販売日時（確定の JST）を記録します（新規行なし）。"
        "**出庫（浮貸）** … **在庫中** のままなら **浮貸日時** 列へ日時を記録し、**販売済** を選ぶ場合は **出庫（販売）と同様** に在庫行を販売済へ更新します（出庫種別はいずれも記録）。"
        "写真は任意（上の共通アップローダ）。"
    )
    outbound_kind = st.radio(
        "出庫区分",
        ("出庫（販売）", "出庫（浮貸）"),
        horizontal=True,
        key="sales_tab_outbound_kind",
    )
    loan_target_status: str | None = None
    if outbound_kind == "出庫（浮貸）":
        loan_target_status = st.radio(
            "出庫（浮貸）の結果ステータス",
            (STATUS_IN_STOCK, STATUS_SOLD),
            horizontal=True,
            key="sales_tab_loan_stock_status",
        )
    _loan_keep_stock = (
        outbound_kind == "出庫（浮貸）" and loan_target_status == STATUS_IN_STOCK
    )
    _loan_as_sale = (
        outbound_kind == "出庫（浮貸）" and loan_target_status == STATUS_SOLD
    )
    _plain_sale = outbound_kind == "出庫（販売）"
    c1, c2 = st.columns(2)
    with c1:
        do_match = st.button(
            "販売元を写真で照合",
            disabled=uploaded is None,
            key="sales_tab_photo_match_btn",
        )
    with c2:
        if st.button("入力をクリア", key="sales_tab_clear_fields_btn"):
            st.session_state.field_sale_source_mgmt_id = ""
            st.session_state[SALES_TAB_QTY_WIDGET_KEY] = 1
            st.session_state.field_actual_sale_excl = 0
            st.session_state.sales_tab_memo = ""
            st.session_state.sales_tab_loan_datetime_manual = ""
            st.session_state.sale_pick_source_id = LEDGER_PICK_PLACEHOLDER
            st.session_state.pop("_sale_link_management_id", None)
            st.session_state.pop("_sale_link_warn", None)
            st.rerun()

    if do_match and uploaded is not None:
        with st.spinner("画像を解析して販売元を照合しています…"):
            try:
                img = _gemini_input_image_from_upload(uploaded)
                inv_ctx = ""
                if df_ledger_hint is not None and not df_ledger_hint.empty:
                    inv_ctx = _build_gemini_inventory_context(df_ledger_hint)
                raw_text = analyze_image_with_gemini(
                    img,
                    inventory_context=inv_ctx or None,
                    prompt_mode="sale_link",
                )
                result = _parse_json_from_model(raw_text or "")
                _apply_gemini_sale_link_to_session(
                    result,
                    df_ledger_hint,
                    fill_product_preview_fields=False,
                )
                st.success("照合が完了しました。管理IDを確認してください。")
            except Exception as e:
                st.warning(
                    "現在混み合っているか、無料枠の上限に達している可能性があります。"
                    "1分ほど待ってから再試行してください。"
                )
                st.caption(f"詳細: {e}")

    _swarn = st.session_state.pop("_sale_link_warn", None)
    if _swarn:
        st.warning(_swarn)
    _sale_link_flash = st.session_state.pop("_sale_link_management_id", None)
    if _sale_link_flash:
        st.info(
            f"販売元として **{_sale_link_flash}** をセットしました。"
            "内容を確認してから確定してください。"
        )

    if df_ledger_hint is not None and not df_ledger_hint.empty:
        _sale_id_opts = _ledger_in_stock_management_ids(df_ledger_hint)
        if _sale_id_opts:
            st.selectbox(
                "在庫中の管理IDから選ぶ",
                options=[LEDGER_PICK_PLACEHOLDER] + _sale_id_opts,
                key="sale_pick_source_id",
                on_change=_on_sale_pick_source_id,
            )

    st.text_input(
        "販売元管理ID（手入力・複数はカンマ等で区切り）",
        key="field_sale_source_mgmt_id",
        placeholder="例: G00000001 または G00000001, G00000002",
    )
    sales_qty = st.number_input(
        "数量（対象点数・管理IDの件数と一致）",
        min_value=1,
        step=1,
        key=SALES_TAB_QTY_WIDGET_KEY,
    )
    if _loan_keep_stock:
        st.text_input(
            "浮貸日時（空欄＝確定ボタン押下の JST）",
            key="sales_tab_loan_datetime_manual",
            help="在庫中のまま出庫（浮貸）を記録するとき、**浮貸日時** 列に入る値です。未入力なら確定実行の JST を記録します。",
        )
    st.number_input(
        "実売金額（税抜・1点あたり）",
        min_value=0,
        step=1,
        key="field_actual_sale_excl",
        disabled=_loan_keep_stock,
        help=(
            "出庫（販売）または出庫（浮貸）で **販売済** のとき必須（1円以上）。"
            "出庫（浮貸）で **在庫中** のときは不要です。"
        ),
    )
    memo_sales = st.text_area(
        "販売メモ（任意・台帳のメモに追記）",
        key="sales_tab_memo",
        height=80,
    )

    _q = int(st.session_state.get(SALES_TAB_QTY_WIDGET_KEY, 1))
    _ids_pv = _split_management_ids_from_field(
        str(st.session_state.get("field_sale_source_mgmt_id", "") or "")
    )
    _act_u = int(st.session_state.get("field_actual_sale_excl", 0))
    _pv_ok_rows: list[pd.Series] = []
    if _ids_pv and len(_ids_pv) != _q:
        st.warning(
            f"販売元管理IDが **{len(_ids_pv)}** 件ですが、数量は **{_q}** です。同じ件数にしてください。"
        )
    if _ids_pv and len(set(_ids_pv)) != len(_ids_pv):
        st.warning("販売元管理IDに **重複** があります。")
    if _ids_pv and df_ledger_hint is not None:
        _pv_msgs: list[str] = []
        for _mid_one in _ids_pv:
            _tr_pv = lookup_ledger_row_by_management_id(df_ledger_hint, _mid_one)
            if _tr_pv is None:
                st.warning(f"管理ID **{_mid_one}** が台帳に見つかりません。")
                continue
            _row_st = _normalize_stock_status(str(_tr_pv.get(COL_STOCK_STATUS, "")))
            if _row_st != STATUS_IN_STOCK:
                st.warning(
                    f"管理ID **{_mid_one}** は在庫中ではありません（現在: {_row_st}）。"
                )
                continue
            _pv_ok_rows.append(_tr_pv)
            _cg1 = _finite_int(_tr_pv.get(COL_PRICE_EXCL), 0)
            _pnv = str(_tr_pv.get(COL_NAME, "") or "").strip()
            _suv = str(_tr_pv.get(COL_SUPPLIER, "") or "").strip()
            _pv_msgs.append(
                f"**{_mid_one}** … {_pnv or '—'} ／ {_suv or '—'} ／ 原価税抜 ¥{_cg1:,}"
            )
        if _pv_msgs:
            st.info("紐付け元（在庫中）:\n" + "\n".join(_pv_msgs))
    elif _ids_pv:
        st.warning("台帳を読み込めないため、紐付け元の原価を表示できません。")

    _sale_pv_agg = (
        bool(_ids_pv)
        and len(_ids_pv) == _q
        and len(set(_ids_pv)) == len(_ids_pv)
        and len(_pv_ok_rows) == len(_ids_pv)
        and len(_pv_ok_rows) > 0
    )
    _cogs_preview = 0
    _pl_u_gp = 0
    _tax_preview = float(CONSUMPTION_TAX_RATE)
    _plex = _pin = _aex = _ain = 0
    _gp_preview: int | None = None
    if _loan_keep_stock and _sale_pv_agg:
        _cogs_preview = sum(_finite_int(x.get(COL_PRICE_EXCL), 0) for x in _pv_ok_rows)
        _tr0 = _pv_ok_rows[0]
        _tax_preview = _infer_tax_rate_from_main_line(
            _finite_int(_tr0.get(COL_PRICE_EXCL), 0),
            _finite_int(_tr0.get(COL_PRICE_INCL), 0),
        )
        _plex = sum(_finite_int(x.get(COL_PLANNED_SALE), 0) for x in _pv_ok_rows)
        _pin = price_incl_tax(_plex, _tax_preview) if _plex > 0 else 0
        _aex = _ain = 0
        _gp_acc = 0
        _gp_any = False
        for _xr in _pv_ok_rows:
            _cgx = _finite_int(_xr.get(COL_PRICE_EXCL), 0)
            _plx_u = _finite_int(_xr.get(COL_PLANNED_SALE), 0)
            _plex1, _, _, _ = _planned_actual_line_amounts(
                1, _plx_u, 0, STATUS_IN_STOCK, _tax_preview
            )
            _gpx = _compute_gross_profit_row(_cgx, _plex1, 0, STATUS_IN_STOCK)
            if _gpx is not None:
                _gp_acc += int(_gpx)
                _gp_any = True
        _gp_preview = _gp_acc if _gp_any else None
    elif _sale_pv_agg and not _loan_keep_stock:
        _cogs_preview = sum(_finite_int(x.get(COL_PRICE_EXCL), 0) for x in _pv_ok_rows)
        _tr0 = _pv_ok_rows[0]
        _tax_preview = _infer_tax_rate_from_main_line(
            _finite_int(_tr0.get(COL_PRICE_EXCL), 0),
            _finite_int(_tr0.get(COL_PRICE_INCL), 0),
        )
        _pl_u_gp = _finite_int(_tr0.get(COL_PLANNED_SALE), 0)
        _plex = sum(_finite_int(x.get(COL_PLANNED_SALE), 0) for x in _pv_ok_rows)
        _pin = price_incl_tax(_plex, _tax_preview) if _plex > 0 else 0
        _aex = _act_u * _q if _act_u > 0 else 0
        _ain = price_incl_tax(_aex, _tax_preview) if _aex > 0 else 0
        _gp_acc = 0
        _gp_any = False
        for _xr in _pv_ok_rows:
            _cgx = _finite_int(_xr.get(COL_PRICE_EXCL), 0)
            _plx = _finite_int(_xr.get(COL_PLANNED_SALE), 0)
            _gpx = _compute_gross_profit_row(_cgx, _plx, _act_u, STATUS_SOLD)
            if _gpx is not None:
                _gp_acc += int(_gpx)
                _gp_any = True
        _gp_preview = _gp_acc if _gp_any else None
    elif _loan_keep_stock and len(_pv_ok_rows) == 1:
        _tr_pv = _pv_ok_rows[0]
        _cogs_preview = _finite_int(_tr_pv.get(COL_PRICE_EXCL), 0)
        _pl_u_gp = _finite_int(_tr_pv.get(COL_PLANNED_SALE), 0)
        _tax_preview = _infer_tax_rate_from_main_line(
            _finite_int(_tr_pv.get(COL_PRICE_EXCL), 0),
            _finite_int(_tr_pv.get(COL_PRICE_INCL), 0),
        )
        _plex, _pin, _aex, _ain = _planned_actual_line_amounts(
            1, _pl_u_gp, 0, STATUS_IN_STOCK, _tax_preview
        )
        _gp_preview = _compute_gross_profit_row(
            _cogs_preview,
            _plex,
            0,
            STATUS_IN_STOCK,
        )
    elif len(_pv_ok_rows) == 1 and not _loan_keep_stock:
        _tr_pv = _pv_ok_rows[0]
        _cogs_preview = _finite_int(_tr_pv.get(COL_PRICE_EXCL), 0)
        _pl_u_gp = _finite_int(_tr_pv.get(COL_PLANNED_SALE), 0)
        _tax_preview = _infer_tax_rate_from_main_line(
            _finite_int(_tr_pv.get(COL_PRICE_EXCL), 0),
            _finite_int(_tr_pv.get(COL_PRICE_INCL), 0),
        )
        _plex, _pin, _aex, _ain = _planned_actual_line_amounts(
            1, _pl_u_gp, _act_u, STATUS_SOLD, _tax_preview
        )
        _gp_preview = _compute_gross_profit_row(
            _cogs_preview,
            _plex,
            _aex,
            STATUS_SOLD,
        )

    pm1, pm2, pm3, pm4, pm5 = st.columns(5)
    with pm1:
        _cogs_lbl = (
            "原価（税抜・合計）" if _sale_pv_agg and _q > 1 else "原価（税抜・参考）"
        )
        if _pv_ok_rows:
            st.metric(_cogs_lbl, f"¥{_cogs_preview:,}")
        else:
            st.metric(_cogs_lbl, "—")
    with pm2:
        st.metric(
            "販売予定（税抜・行計）",
            "—" if _plex <= 0 else f"¥{_plex:,}",
        )
    with pm3:
        st.metric(
            "販売予定（税込・総額）",
            "—" if _pin <= 0 else f"¥{_pin:,}",
        )
    with pm4:
        st.metric(
            "実売（税抜・行計）",
            "—" if _aex <= 0 else f"¥{_aex:,}",
        )
    with pm5:
        st.metric(
            "実売（税込・総額）",
            "—" if _ain <= 0 else f"¥{_ain:,}",
        )
    pm6, _, _ = st.columns([1, 1, 3])
    with pm6:
        st.metric(
            "粗利（税抜・プレビュー）",
            "—" if _gp_preview is None else f"¥{int(_gp_preview):,}",
        )
    if _loan_keep_stock:
        st.caption(
            f"在庫中のまま出庫（浮貸）を確定すると、各行の **{COL_LOAN_DATETIME}** に日時が入ります（手入力が空なら確定の JST）。"
        )

    _confirm_lbl = (
        "浮貸を確定（在庫中のまま・浮貸日時のみ）"
        if _loan_keep_stock
        else "販売を確定（在庫行のみ更新・新規行なし）"
    )
    confirm_sale = st.button(_confirm_lbl, type="primary", key="sales_tab_confirm_btn")

    if confirm_sale:
        _sale_src_save = str(
            st.session_state.get("field_sale_source_mgmt_id", "") or ""
        ).strip()
        _act_ex2 = int(st.session_state.get("field_actual_sale_excl", 0))
        _q_sv = int(st.session_state.get(SALES_TAB_QTY_WIDGET_KEY, 1))
        _ids_sale_val = _split_management_ids_from_field(_sale_src_save)
        memo_s = (memo_sales or "").strip()
        validation_ok = True
        _need_actual = _plain_sale or _loan_as_sale
        if not _sale_src_save:
            st.error("**管理ID**（販売元管理ID欄）の入力が必須です。")
            validation_ok = False
        elif _need_actual and _act_ex2 < 1:
            st.error("**実売金額（税抜）** を1円以上で入力してください。")
            validation_ok = False
        elif df_ledger_hint is None:
            st.error("台帳を読み込めないため、反映できません。")
            validation_ok = False
        elif len(_ids_sale_val) != _q_sv:
            st.error(
                "**管理ID** を **数量と同じ件数** で入力してください。"
                f"（数量 **{_q_sv}** に対し **{len(_ids_sale_val)}** 件と読み取りました。）"
            )
            validation_ok = False
        elif len(set(_ids_sale_val)) != len(_ids_sale_val):
            st.error("管理IDに **重複** があります。")
            validation_ok = False
        else:
            for _sid_v in _ids_sale_val:
                trv = lookup_ledger_row_by_management_id(df_ledger_hint, _sid_v)
                if trv is None:
                    st.error(f"管理ID {_sid_v} が台帳に見つかりません。")
                    validation_ok = False
                    break
                if (
                    _normalize_stock_status(str(trv.get(COL_STOCK_STATUS, "")))
                    != STATUS_IN_STOCK
                ):
                    st.error(
                        f"管理ID **{_sid_v}** は在庫中ではありません。既に販売済の可能性があります。"
                    )
                    validation_ok = False
                    break

        if validation_ok:
            urls: list[str] = [""]
            if uploaded is not None:
                with st.spinner("画像をリサイズしてドライブに保存しています…"):
                    raw0 = uploaded.getvalue()
                    try:
                        data0, mime0 = prepare_upload_image_jpeg(raw0)
                    except Exception as e:
                        st.error(f"画像の処理に失敗しました: {e}")
                        st.warning("画像なしで台帳の反映のみ続行します。")
                    else:
                        safe_base = re.sub(
                            r"[^\w\-_.]", "_", uploaded.name.rsplit(".", 1)[0]
                        )[:80]
                        fname0 = f"{jst_now().strftime('%Y%m%d_%H%M%S')}_{safe_base}_{uuid.uuid4().hex[:8]}.jpg"
                        try:
                            urls[0] = upload_image_to_drive(fname0, mime0, data0)
                        except Exception as e:
                            st.error(f"ドライブ保存に失敗しました: {e}")
                            st.warning("画像URLは付けずに反映のみ続行します。")
            if not _uses_local_inventory_csv() and not _secret_str(
                SECRET_GOOGLE_SPREADSHEET_ID
            ):
                st.warning("台帳の保存先が未設定のため、反映をスキップしました。")
            else:
                try:
                    _n_ids = len(_ids_sale_val)
                    if _loan_keep_stock:
                        _loan_manual = str(
                            st.session_state.get("sales_tab_loan_datetime_manual", "")
                            or ""
                        ).strip()
                        _spin = (
                            "浮貸日時を記録しています…"
                            if _n_ids <= 1
                            else f"在庫行 **{_n_ids} 件** に浮貸日時を記録しています…"
                        )
                        with st.spinner(_spin):
                            for _sid_save in _ids_sale_val:
                                apply_outbound_loan_in_stock_datetime_by_management_id(
                                    _sid_save,
                                    loan_datetime_jst=_loan_manual or None,
                                    new_image_url=(urls[0] if urls else "") or "",
                                    memo_suffix=memo_s,
                                )
                    else:
                        _ot = "出庫（販売）" if _plain_sale else "出庫（浮貸）"
                        _spin_sale = (
                            "該当の在庫行を販売済に更新しています…"
                            if _n_ids <= 1
                            else f"在庫行 **{_n_ids} 件** を順に販売済に更新しています…"
                        )
                        with st.spinner(_spin_sale):
                            for _sid_save in _ids_sale_val:
                                apply_outbound_sale_to_ledger_by_management_id(
                                    _sid_save,
                                    actual_sale_unit_excl_yen=_act_ex2,
                                    new_image_url=(urls[0] if urls else "") or "",
                                    memo_suffix=memo_s,
                                    sale_outbound_type=_ot,
                                )
                except Exception as e:
                    st.error(f"台帳の更新に失敗しました: {e}")
                else:
                    st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                    if _loan_keep_stock:
                        if len(_ids_sale_val) <= 1:
                            st.success(
                                f"管理ID **{_ids_sale_val[0]}** に **{COL_LOAN_DATETIME}** を記録しました（在庫中のまま）。"
                            )
                        else:
                            st.success(
                                f"**{len(_ids_sale_val)} 件** に {COL_LOAN_DATETIME} を記録しました（管理ID: {'、'.join(_ids_sale_val)}）。"
                            )
                    elif len(_ids_sale_val) <= 1:
                        _sid_one = _ids_sale_val[0]
                        st.success(
                            f"管理ID **{_sid_one}** の行を販売済に更新しました（実売 ¥{_act_ex2:,}・販売日時は確定実行の JST）。"
                        )
                    else:
                        _ids_join = "、".join(_ids_sale_val)
                        st.success(
                            f"**{len(_ids_sale_val)} 件** の在庫行を販売済に更新しました（管理ID: {_ids_join}）。"
                            f"実売（税抜・1点あたり）は各行 ¥{_act_ex2:,}、販売日時は確定実行の JST を記録しています。"
                        )
                    if urls[0]:
                        st.markdown(f"[保存した画像を開く]({urls[0]})")
                    st.balloons()


def main():
    st.set_page_config(page_title="商品在庫・販売", layout="wide")
    _nav_opts = ("登録（インプット）", "ギャラリー（カタログ）・在庫一覧", "集計・分析（ダッシュボード）")
    if "nav_page" not in st.session_state:
        st.session_state.nav_page = _nav_opts[0]
    with st.sidebar:
        st.markdown("### メニュー")
        page = st.radio("ページ", _nav_opts, key="nav_page")
    st.title("商品在庫・販売管理")
    st.caption(
        "写真は任意。台帳の必須項目のみの記録、または写真＋AI解析・ドライブ保存・"
        "**inventory.csv** またはスプレッドシートへの記録ができます。"
    )
    if page == "ギャラリー（カタログ）・在庫一覧":
        render_inventory_list_page()
        return
    if page == "集計・分析（ダッシュボード）":
        render_analytics_dashboard_page()
        return

    _init_registration_form_session_state()
    _init_voucher_sidebar_state()
    df_ledger_hint = _ledger_hint_dataframe()

    st.markdown("## 台帳登録")
    st.caption(
        "仕入れ・販売・棚卸しは **下の大きなタブ** で切り替えます。"
        "下の **1枚の写真** は全タブ共通です（AI 解析は長辺最大"
        f"{UPLOAD_JPEG_MAX_LONG_EDGE}px・品質{UPLOAD_JPEG_QUALITY}％、"
        f"仕入れ確定で Drive 保存するときは長辺{PURCHASE_DRIVE_JPEG_MAX_LONG_EDGE}px・品質{PURCHASE_DRIVE_JPEG_QUALITY}％に変換します）。"
    )
    uploaded = st.file_uploader(
        "商品写真（任意・1枚まで・カメラやギャラリーから）",
        type=["jpg", "jpeg", "png", "webp"],
        key="shared_reg_photo_uploader",
    )
    st.caption(
        "写真は **1枚まで** です。数量が **2以上** のときは、その1枚をドライブに保存し、"
        "作成する **全行に同じ画像URL** を入れます。"
        "台帳の日時は写真の EXIF 撮影日時を優先し、写真がないときは日本時間（JST）の現在時刻です。"
    )

    _inject_prominent_main_tabs_style()
    st.markdown("## 入力モード")
    st.caption(
        "**上の大きなタブ** で切り替えます。**仕入れ登録** で証憑・入庫と新規行、**販売管理** で販売・浮貸の反映、"
        "**棚卸しスキャン** で棚卸日の確定（いずれも **上の1枚の写真** を共通で使えます）。"
    )
    tab_purchase, tab_sales, tab_stock = st.tabs(
        ("仕入れ登録", "販売管理", "棚卸しスキャン")
    )

    with tab_purchase:
        _render_voucher_inventory_panel()
        st.divider()
        st.markdown("##### クイック検索（写真から検索）")
        st.caption(
            "**AIで画像を解析** で商品名・柄色などを推定しつつ在庫中と照合します。"
            "解析後は下の「在庫中の近い候補」も併せて確認してください。"
        )

        movement = st.radio(
            "区分（仕入れ・在庫の増減）",
            ("入庫（購入）", "入庫（返品）", "入庫（浮貸）"),
            horizontal=True,
            key="tab_purchase_movement",
        )
    
        col_a, col_c = st.columns([1, 1])
        with col_a:
            analyze = st.button(
                "AIで画像を解析",
                type="primary",
                disabled=uploaded is None,
            )
        with col_c:
            if st.button("候補の自動入力をクリア"):
                st.session_state.field_product_name = ""
                st.session_state.field_supplier = ""
                st.session_state.field_qty = 1
                st.session_state[REGISTRATION_QTY_WIDGET_KEY] = 1
                st.session_state.ai_kind = ""
                st.session_state.ai_features = ""
                st.session_state.ai_parse_ran = False
                st.session_state.field_memo = ""
                st.session_state.field_line_excl_yen = 1
                st.session_state.field_planned_sale_excl = 0
                st.session_state.field_actual_sale_excl = 0
                st.session_state.field_stock_status = STATUS_IN_STOCK
                st.session_state.hint_filter_product_name = ""
                st.session_state.hint_filter_supplier = ""
                st.session_state.ledger_pick_product_name = LEDGER_PICK_PLACEHOLDER
                st.session_state.ledger_pick_supplier = LEDGER_PICK_PLACEHOLDER
                st.session_state.field_sale_source_mgmt_id = ""
                st.session_state.sale_pick_source_id = LEDGER_PICK_PLACEHOLDER
                st.session_state.pop("ledger_quick_candidates", None)
                st.session_state.pop("_gemini_match_management_id", None)
                st.session_state.pop("_sale_link_management_id", None)
                st.session_state.pop("_sale_link_warn", None)
                st.rerun()
    
        if analyze and uploaded is not None:
            with st.spinner("画像を解析しています…"):
                try:
                    img = _gemini_input_image_from_upload(uploaded)
                    inv_ctx = ""
                    if df_ledger_hint is not None and not df_ledger_hint.empty:
                        inv_ctx = _build_gemini_inventory_context(df_ledger_hint)
                    raw_text = analyze_image_with_gemini(
                        img,
                        inventory_context=inv_ctx or None,
                    )
                    result = _parse_json_from_model(raw_text or "")
                    _apply_gemini_json_to_session(result, df_ledger_hint)
                    _refresh_ledger_quick_search_candidates(df_ledger_hint)
                    st.success(
                        "解析が完了しました。必要に応じて商品名・仕入先・取引先・数量・仕入金額（税抜）を修正してください。"
                    )
                except Exception as e:
                    st.warning(
                        "現在混み合っているか、無料枠の上限に達している可能性があります。"
                        "1分ほど待ってから再試行してください。"
                    )
                    st.caption(f"詳細: {e}")
    
        if st.session_state.get("ai_parse_ran"):
            st.subheader("AI解析結果（参考）")
            st.write(f"**推定種類:** {st.session_state.ai_kind or '—'}")
            st.write(f"**推定数量:** {int(st.session_state.field_qty)}")
            st.write(
                f"**推定仕入金額（税抜・1点）:** ¥{int(st.session_state.field_line_excl_yen):,}"
            )
            st.caption(f"マッチング用特徴: {st.session_state.ai_features or '—'}")
            mid_hit = st.session_state.get("_gemini_match_management_id")
            if mid_hit:
                st.info(f"台帳照合: 管理ID **{mid_hit}** の在庫行に合わせて、商品名・仕入先・仕入金額（税抜）を反映しました。")
    
        if df_ledger_hint is not None and not df_ledger_hint.empty:
            st.markdown("##### 台帳から入力補助（任意）")
            st.caption(
                "絞り込み欄に文字を入れると候補が絞られます。プルダウンで選ぶと下の入力欄に反映されます（あとから手修正も可能です）。"
                "在庫中の行に一致したときは **販売予定金額（税抜・任意）** にも、台帳の1点あたりの値を入れます（仕入先まで一致する行を優先）。"
            )
            hc1, hc2 = st.columns(2)
            with hc1:
                st.text_input(
                    "商品名の絞り込み（部分一致）",
                    key="hint_filter_product_name",
                    placeholder="例: 帯",
                )
                fp = st.session_state.get("hint_filter_product_name", "")
                if st.session_state.get("_hint_fp_seen", "") != fp:
                    st.session_state["_hint_fp_seen"] = fp
                    st.session_state.ledger_pick_product_name = LEDGER_PICK_PLACEHOLDER
                opts_p = _ledger_unique_col_values(df_ledger_hint, COL_NAME)
                if fp.strip():
                    q = fp.strip().casefold()
                    opts_p = [x for x in opts_p if q in x.casefold()][:400]
                st.selectbox(
                    "台帳に登録済みの商品名から選ぶ",
                    options=[LEDGER_PICK_PLACEHOLDER] + opts_p,
                    key="ledger_pick_product_name",
                    on_change=_on_ledger_pick_product_name,
                )
            with hc2:
                st.text_input(
                    "仕入先・取引先の絞り込み（部分一致）",
                    key="hint_filter_supplier",
                    placeholder="例: ⚫︎⚫︎会社",
                )
                fs = st.session_state.get("hint_filter_supplier", "")
                if st.session_state.get("_hint_fs_seen", "") != fs:
                    st.session_state["_hint_fs_seen"] = fs
                    st.session_state.ledger_pick_supplier = LEDGER_PICK_PLACEHOLDER
                opts_s = _ledger_unique_col_values(df_ledger_hint, COL_SUPPLIER)
                if fs.strip():
                    q = fs.strip().casefold()
                    opts_s = [x for x in opts_s if q in x.casefold()][:400]
                st.selectbox(
                    "台帳に登録済みの仕入先・取引先から選ぶ",
                    options=[LEDGER_PICK_PLACEHOLDER] + opts_s,
                    key="ledger_pick_supplier",
                    on_change=_on_ledger_pick_supplier,
                )
        elif _uses_local_inventory_csv() or _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
            st.caption("台帳が空か読み込めないため、入力補助の候補は表示できません。")
    
        st.markdown("##### 必須入力項目")
        st.caption(
            "このタブの確定は **在庫中** の新規行のみを追加します（入庫（購入）／入庫（返品）／入庫（浮貸））。"
            "**出庫（浮貸）・出庫（販売）** は **販売管理** タブで行ってください。"
        )
        product_name = st.text_input("商品名（必須）", key="field_product_name")
        supplier = st.text_input("仕入先・取引先（必須）", key="field_supplier")
        _refresh_ledger_quick_search_candidates(df_ledger_hint)
        _cand = st.session_state.get("ledger_quick_candidates")
        if (
            isinstance(_cand, pd.DataFrame)
            and not _cand.empty
            and df_ledger_hint is not None
        ):
            with st.expander("在庫中の近い候補（写真解析・入力文字から照合）", expanded=False):
                st.caption(
                    "商品名・仕入先の表記が近い **在庫中** の行を最大8件表示しています。"
                    "上の「台帳から入力補助」で同じ文言を選ぶか、管理IDを手元で確認して台帳一覧と突き合わせてください。"
                )
                _show_cols = [
                    c
                    for c in (
                        COL_MANAGEMENT_ID,
                        COL_NAME,
                        COL_SUPPLIER,
                        COL_PRICE_EXCL,
                        COL_PLANNED_SALE,
                        COL_LAST_STOCKTAKE,
                        COL_SALE_SOURCE_MGMT_ID,
                    )
                    if c in _cand.columns
                ]
                st.dataframe(
                    _cand[_show_cols],
                    use_container_width=True,
                    hide_index=True,
                )
    
        quantity = st.number_input(
            "数量（点数）",
            min_value=1,
            step=1,
            key=REGISTRATION_QTY_WIDGET_KEY,
            help=(
                "台帳は **1点1行** で保存します。行数は常にこの数量と同じです。"
                "写真は1枚まで・複数点のときは **同じ画像URL** を各行に入れます。"
            ),
        )
        st.session_state.field_qty = int(quantity)

        line_excl_yen = st.number_input(
            "仕入金額（税抜・必須）",
            min_value=1,
            step=1,
            key="field_line_excl_yen",
            help="1点あたりの税抜の仕入金額（円）。台帳の各行は数量1で、この金額が税抜行計になります。",
        )
    
        st.radio(
            "消費税（仕入金額（税込）の計算）",
            options=list(CONSUMPTION_TAX_CHOICE_TO_RATE.keys()),
            horizontal=True,
            key="field_consumption_tax_choice",
            help="仕入金額（税抜）の税込行計に使用します。非課税のときは税込＝税抜です。",
        )
        _tax_r = _consumption_tax_rate_from_choice_label(
            str(st.session_state.get("field_consumption_tax_choice", "10%"))
        )
    
        _q = int(quantity)
        _lex_inp = int(line_excl_yen)
        _n_save = _q
        _line_ex_one = _lex_inp
        _line_in_one = price_incl_tax(_line_ex_one, _tax_r)
    
        price_row = st.columns([1, 1, 1])
        with price_row[0]:
            st.metric("仕入金額（税抜・1点）", f"¥{_line_ex_one:,}")
            _cap_rows = (
                f"確定時は **{_n_save} 行**（各行 数量1）。税抜合計（参考） ¥{_line_ex_one * _n_save:,}。"
            )
            if _n_save > 1:
                _cap_rows += (
                    "写真があるとき、数量が2以上なら **同じ画像URLを全行** に記録します。"
                )
            st.caption(_cap_rows)
        with price_row[1]:
            st.metric("仕入金額（税込・1点・自動）", f"¥{_line_in_one:,}")
            _tl = st.session_state.get("field_consumption_tax_choice", "10%")
            if _tl == "非課税":
                st.caption("非課税のため税込＝税抜行合計")
            else:
                st.caption(f"消費税{_tl}を行合計に四捨五入")
        with price_row[2]:
            st.caption(
                "原価は各行の仕入金額（税抜）です。販売予定・実売・販売元の詳細は下の **価格管理／販売管理** で入力します。"
            )
    
        st.markdown("##### 価格管理（任意）")
        st.caption(
            "「販売予定金額（税抜）」は **1点あたりの税抜金額（円）** です。"
            "「在庫中」のときは販売予定行計−原価で粗利の参考になります（下のプレビュー）。"
        )
        planned_sale_excl = st.number_input(
            "販売予定金額（税抜・任意）",
            min_value=0,
            step=1,
            key="field_planned_sale_excl",
            help="1点あたり。0 のとき台帳では空欄。税抜行計・税込総額は各行数量1として自動計算します。",
        )
    
        st.markdown("##### 販売・実売について")
        st.caption(
            "このタブでは **新規行の追加のみ** です（常に **在庫中**）。"
            "**販売元管理ID・実売・販売済更新** は **販売管理** タブを使用してください。"
        )
        _pl_u = int(planned_sale_excl)
        _act_u = 0
        _cogs_preview = _lex_inp * _q
        _pl_u_gp = _pl_u
        _tax_preview = _tax_r
        _st_gp = STATUS_IN_STOCK
        _plex, _pin, _aex, _ain = _planned_actual_line_amounts(
            _q, _pl_u_gp, _act_u, _st_gp, _tax_preview
        )
        _gp_preview = _compute_gross_profit_row(
            _cogs_preview,
            _plex,
            0,
            _st_gp,
        )
        pm1, pm2, pm3, pm4, pm5 = st.columns(5)
        with pm1:
            _cogs_lbl = "原価（税抜・行合計）"
            st.metric(_cogs_lbl, f"¥{_cogs_preview:,}")
        with pm2:
            st.metric(
                "販売予定（税抜・行計）",
                "—" if _plex <= 0 else f"¥{_plex:,}",
            )
        with pm3:
            st.metric(
                "販売予定（税込・総額）",
                "—" if _pin <= 0 else f"¥{_pin:,}",
            )
        with pm4:
            st.metric(
                "実売（税抜・行計）",
                "—" if _aex <= 0 else f"¥{_aex:,}",
            )
        with pm5:
            st.metric(
                "実売（税込・総額）",
                "—" if _ain <= 0 else f"¥{_ain:,}",
            )
        pm6, _, _ = st.columns([1, 1, 3])
        with pm6:
            st.metric(
                "粗利（税抜・プレビュー）",
                "—" if _gp_preview is None else f"¥{int(_gp_preview):,}",
            )
    
        st.markdown("##### 補足情報（任意）")
        memo = st.text_area(
            "メモ（任意）",
            key="field_memo",
            height=100,
            placeholder="備考・社内メモなどがあれば入力してください",
        )
    
        confirm = st.button(
            "確定（台帳に記録・写真は任意でドライブ保存）",
            type="primary",
        )
    
        if confirm:
            validation_ok = True
            _sale_src_save = ""
            _act_ex2 = 0
            if not (product_name or "").strip():
                st.error("商品名を入力してください。")
                validation_ok = False
            elif not (supplier or "").strip():
                st.error("仕入先・取引先を入力してください。")
                validation_ok = False
            elif int(line_excl_yen) < 1:
                st.error("仕入金額（税抜）を1円以上で入力してください。")
                validation_ok = False

            if validation_ok:
                _lex_one = int(line_excl_yen)
                _tax_r2 = _consumption_tax_rate_from_choice_label(
                    str(st.session_state.get("field_consumption_tax_choice", "10%"))
                )
                _lin_one = price_incl_tax(_lex_one, _tax_r2)
                _plan2 = int(st.session_state.get("field_planned_sale_excl", 0))
                _stat2 = STATUS_IN_STOCK
                memo_s = (memo or "").strip()
    
                _q2 = int(quantity)
                n_save = _q2
                urls: list[str] = [""] * n_save
                _record_dt = jst_now_str()
                ready_for_sheet = True
    
                if uploaded is not None:
                    with st.spinner("画像をリサイズ・圧縮してドライブに保存しています…"):
                        raw0 = uploaded.getvalue()
                        _record_dt = (
                            capture_datetime_jst_from_bytes(raw0) or _record_dt
                        )
                        try:
                            data0, mime0 = prepare_upload_image_jpeg(
                                raw0,
                                max_long_edge=PURCHASE_DRIVE_JPEG_MAX_LONG_EDGE,
                                quality=PURCHASE_DRIVE_JPEG_QUALITY,
                            )
                        except Exception as e:
                            st.error(f"画像の処理に失敗しました: {e}")
                            ready_for_sheet = False
                        else:
                            safe_base = re.sub(
                                r"[^\w\-_.]", "_", uploaded.name.rsplit(".", 1)[0]
                            )[:80]
                            fname0 = f"{jst_now().strftime('%Y%m%d_%H%M%S')}_{safe_base}_{uuid.uuid4().hex[:8]}.jpg"
                            try:
                                shared_url = upload_image_to_drive(fname0, mime0, data0)
                            except Exception as e:
                                st.error(f"ドライブ保存に失敗しました: {e}")
                                ready_for_sheet = False
                            else:
                                urls = [shared_url] * n_save
    
                if ready_for_sheet:
                    ws0 = (
                        None
                        if _uses_local_inventory_csv()
                        else ensure_worksheet_header()
                    )
                    if not _uses_local_inventory_csv() and ws0 is None:
                        st.warning(
                            "スプレッドシート未設定のため、行の追記をスキップしました。"
                        )
                    else:
                        try:
                            ids = allocate_management_ids(ws0, n_save)
                            _spin_msg = (
                                "inventory.csv に記録しています…"
                                if _uses_local_inventory_csv()
                                else "スプレッドシートに記録しています…"
                            )
                            with st.spinner(_spin_msg):
                                _reg_row_dt = jst_now_str()
                                _reg_batch: list[list[Any]] = []
                                for i in range(n_save):
                                    _reg_batch.append(
                                        _inventory_row_values_for_append(
                                            _reg_row_dt,
                                            _record_dt,
                                            movement,
                                            product_name.strip(),
                                            supplier.strip(),
                                            _lex_one,
                                            _lin_one,
                                            urls[i],
                                            ids[i],
                                            memo_s,
                                            planned_sale_unit_excl_yen=_plan2,
                                            actual_sale_unit_excl_yen=_act_ex2,
                                            stock_status=_stat2,
                                            consumption_tax_rate=_tax_r2,
                                            sale_source_management_id=_sale_src_save,
                                        )
                                    )
                                _append_inventory_data_rows(_reg_batch)
                        except Exception as e:
                            st.error(f"台帳の更新に失敗しました: {e}")
                            if any(urls):
                                st.warning(
                                    "一部の画像はドライブに保存済みの可能性があります。台帳の内容を確認してください。"
                                )
                        else:
                            st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                            _msg_ok = (
                                f"記録しました（{n_save} 行・1点1行）。管理IDを自動付与しています。"
                            )
                            st.success(_msg_ok)
                            _link_urls = list(dict.fromkeys(u for u in urls if u))
                            for _uurl in _link_urls[:8]:
                                st.markdown(f"[保存した画像を開く]({_uurl})")
                            if len(_link_urls) > 8:
                                st.caption(
                                    f"ほか {len(_link_urls) - 8} 件の画像URLは台帳の「{COL_IMAGE_URL}」列を参照してください。"
                                )
                            st.balloons()

    with tab_sales:
        _render_sales_management_tab(uploaded, df_ledger_hint)

    with tab_stock:
        render_stocktake_scan_tab(df_ledger_hint)


if __name__ == "__main__":
    if check_password():
        main()
