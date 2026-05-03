"""
商品在庫・販売管理 (Streamlit)

st.secrets に以下を設定してください（例は .streamlit/secrets.toml）。

必須キー:
  GEMINI_API_KEY
  GAS_UPLOAD_URL           … 画像を Google ドライブに保存する Web アプリ（GAS）の URL
  GAS_API_KEY                … GAS Web アプリ呼び出し用の共有キー（payload の apiKey に付与）
  GOOGLE_DRIVE_FOLDER_ID     … 保存先フォルダID（GAS に渡す）
  GOOGLE_SPREADSHEET_ID      … 記録用スプレッドシートID
  google_service_account     … サービスアカウントJSONの各フィールド（[google_service_account] セクション）
    または GOOGLE_SERVICE_ACCOUNT_JSON … JSON文字列1本

任意:
  GEMINI_MODEL_NAME          … Gemini モデル ID（未設定時は下記 DEFAULT_GEMINI_MODEL）
  GOOGLE_WORKSHEET_NAME      … ワークシート名（未設定時は DEFAULT_WORKSHEET_NAME）
  GAS_UPLOAD_TIMEOUT_SECONDS … GAS への POST タイムアウト秒（既定 300、1〜3600 にクランプ）
  FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED … GAS 未設定時に台帳の画像URL列へ入れるプレースホルダ URL
  APP_PASSWORD               … アプリ画面の簡易ログイン用（平文。GitHub には secrets.toml をコミットしないこと）

※ 画像の Gemini 解析は **google-generativeai** を使用します。モデル名は ``GEMINI_MODEL_NAME``（既定は flash 系プレビュー）。
※ アップロード画像は任意。ある場合のみ Pillow で長辺最大1280px・JPEG品質80に変換してから解析・ドライブ保存します。
※ 台帳日時・撮影日時未取得時の現在時刻は **pytz** の ``Asia/Tokyo``（JST）です。

画面下部の「在庫一覧」で、同一スプレッドシートを表形式で読み書きし、
入出庫の集計・仕入先・取引先別サマリー・月次グラフを表示できます。

スプレッドシート1行目はヘッダーとして次の列順を想定:
  日時 | 入出庫種別 | 商品名 | 仕入先・取引先 | 数量 | 仕入金額（税抜） | 仕入金額（税込）
  | 販売予定単価（税抜） | 販売予定金額（税込） | 実売単価（税抜） | 実売金額（税込） | 粗利 | ステータス（在庫中/販売済） | メモ（任意） | 画像URL | 管理ID
  ※在庫は **1点につき1行** で統一します。登録時の行数は **数量** と同じで、各行の数量は **1** です。
  ※写真は **1枚まで** アップロードできます。写真があるときは1回だけドライブに保存し、数量が **2以上** のときは **全行に同じ画像URL** を入れます（数量が1のときはその1行のみ）。
  ※「管理ID」列は自動採番（例: G00000001）のシリアルです。既存行の末尾に列を追加しても列位置はずれません。
  ※「日時」列への新規記入は **日本時間（JST / Asia/Tokyo）** で行い、画像に EXIF 撮影日時があればそれを JST として解釈して優先します。
  ※「仕入金額（税抜）」「仕入金額（税込）」は **1点あたりの行合計**（台帳の各行は数量1）です。
  ※旧シートに「仕入単価（税抜）」列が残っている場合は、読み込み時にその列を除いて新しい列構成に揃えます。
  ※新規登録画面では仕入金額（税込）の計算に使う消費税を **10% / 8% / 非課税** から選べます（既定は10%）。
  ※販売予定・実売は **1点あたり税抜単価** を保存し、税込総額列は「単価×数量」を税抜行合計にしたうえで仕入行と同じ税率で四捨五入します。
  ※金額列（単価〜粗利まで）は書き込み時に表示形式 **#,##0** を適用します。
  ※粗利は税抜ベースで「販売済」なら（実売単価×数量）−原価、「在庫中」なら（販売予定単価×数量）−原価。台帳保存時に再計算します。
"""

from __future__ import annotations

import base64
import io
import json
import math
import re
import uuid
from datetime import date, datetime
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

# --- st.secrets のキー名（文字列リテラルの散在を避ける） ---
SECRET_GEMINI_API_KEY = "GEMINI_API_KEY"
SECRET_GEMINI_MODEL_NAME = "GEMINI_MODEL_NAME"
SECRET_GAS_UPLOAD_URL = "GAS_UPLOAD_URL"
SECRET_GAS_API_KEY = "GAS_API_KEY"
SECRET_GAS_UPLOAD_TIMEOUT_SECONDS = "GAS_UPLOAD_TIMEOUT_SECONDS"
SECRET_GOOGLE_DRIVE_FOLDER_ID = "GOOGLE_DRIVE_FOLDER_ID"
SECRET_GOOGLE_SPREADSHEET_ID = "GOOGLE_SPREADSHEET_ID"
SECRET_GOOGLE_WORKSHEET_NAME = "GOOGLE_WORKSHEET_NAME"
SECRET_GOOGLE_SERVICE_ACCOUNT_JSON = "GOOGLE_SERVICE_ACCOUNT_JSON"
SECRET_GOOGLE_SERVICE_ACCOUNT_SECTION = "google_service_account"
SECRET_APP_PASSWORD = "APP_PASSWORD"
SECRET_FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED = "FALLBACK_IMAGE_URL_WHEN_GAS_UNCONFIGURED"

# --- secrets に無いときの既定（非機密のデフォルトのみ） ---
DEFAULT_GEMINI_MODEL = "gemini-3-flash-preview"
DEFAULT_WORKSHEET_NAME = "在庫履歴"
DEFAULT_GAS_FALLBACK_IMAGE_URL = "https://example.com/?gofuku-app=skipped-no-gas-secrets"
DEFAULT_GAS_UPLOAD_TIMEOUT_SECONDS = 300

# --- 画像アップロード前処理 ---
UPLOAD_JPEG_MAX_LONG_EDGE = 1280
UPLOAD_JPEG_QUALITY = 80

# --- スプレッドシート列 ---
COL_DATETIME = "日時"
COL_TYPE = "入出庫種別"
COL_NAME = "商品名"
COL_SUPPLIER = "仕入先・取引先"
COL_QTY = "数量"
COL_PRICE_EXCL = "仕入金額（税抜）"
COL_PRICE_INCL = "仕入金額（税込）"
# 旧スプレッドシート1行目に残る列名（load 時に列を落として新 EXPECTED に合わせる）
LEGACY_COL_UNIT_PRICE = "仕入単価（税抜）"
COL_PLANNED_SALE = "販売予定単価（税抜）"
COL_PLANNED_SALE_INCL = "販売予定金額（税込）"
COL_ACTUAL_SALE = "実売単価（税抜）"
COL_ACTUAL_SALE_INCL = "実売金額（税込）"
COL_GROSS_PROFIT = "粗利"
COL_STOCK_STATUS = "ステータス（在庫中/販売済）"
COL_IMAGE_URL = "画像URL"
COL_MEMO = "メモ"
COL_MANAGEMENT_ID = "管理ID"

STATUS_IN_STOCK = "在庫中"
STATUS_SOLD = "販売済"
STOCK_STATUS_OPTIONS: tuple[str, ...] = (STATUS_IN_STOCK, STATUS_SOLD)

CONSUMPTION_TAX_RATE = 0.10
CONSUMPTION_TAX_CHOICE_TO_RATE: dict[str, float] = {
    "10%": 0.10,
    "8%": 0.08,
    "非課税": 0.0,
}

EXPECTED_HEADERS: list[str] = [
    COL_DATETIME,
    COL_TYPE,
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
]

SHEET_AMOUNT_NUMBER_PATTERN = "#,##0"
TZ_JP = pytz.timezone("Asia/Tokyo")
LEDGER_DATA_EDITOR_KEY = "inventory_ledger_data_editor"
LEDGER_PICK_PLACEHOLDER = "（選ばない）"


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
    """金額系の列（単価〜粗利まで）に、2行目以降で #,##0 を適用する。"""
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


def prepare_upload_image_jpeg(raw: bytes) -> tuple[bytes, str]:
    """Gemini 送信用・GAS 保存用の共通前処理。

    EXIF 向き補正のうえ、:data:`UPLOAD_JPEG_MAX_LONG_EDGE` と :data:`UPLOAD_JPEG_QUALITY` に従い
    長辺リサイズと JPEG 再エンコードを行う。

    Returns:
        (jpeg_bytes, mime_type)  mime_type は常に ``image/jpeg`` 。
    """
    img = Image.open(io.BytesIO(raw))
    img = ImageOps.exif_transpose(img)
    rgba = img.convert("RGBA")
    bg = Image.new("RGB", rgba.size, (255, 255, 255))
    bg.paste(rgba, mask=rgba.getchannel("A"))
    img = bg
    img = _resize_long_edge_max(img, UPLOAD_JPEG_MAX_LONG_EDGE)
    buf = io.BytesIO()
    img.save(
        buf,
        format="JPEG",
        quality=UPLOAD_JPEG_QUALITY,
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


def _apply_gemini_json_to_session(result: dict[str, Any]) -> None:
    """Gemini の JSON をフォーム用 session_state に反映する（英日キー両対応）。"""
    r = result
    st.session_state.field_qty = _coerce_positive_int(
        r.get("quantity") or r.get("数量") or r.get("qty") or 1,
        default=1,
    )

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
    m = r.get("match")
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
    if line_yen is None and match_conf_ok:
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
    """販売予定・実売の税抜行計と税込行計（単価×数量を税抜合計にしてから税込）。"""
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


def analyze_image_with_gemini(image_data):
    api_key = _secret_str(SECRET_GEMINI_API_KEY)
    if not api_key:
        raise RuntimeError(
            f"{SECRET_GEMINI_API_KEY} が設定されていません。`.streamlit/secrets.toml` を確認してください。"
        )
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(_gemini_model_name())
    prompt = """この画像は呉服店の在庫・売買用の商品写真です。次のキーだけを持つ JSON オブジェクトを 1 つだけ返してください。
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

任意: 既存在庫と照合する場合のみ "match" を付けてもよい
{"product_name": "...", "supplier": "...", "unit_price_excl": 整数またはnull, "confidence": 0.0〜1.0} の形。不要なら省略。"""
    response = model.generate_content([prompt, image_data])
    return response.text or ""


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
        out[c] = s.map(lambda x: _finite_int(x, 0))
    return out


def _normalize_stock_status(status: str) -> str:
    s = (status or "").strip()
    return s if s in STOCK_STATUS_OPTIONS else STATUS_IN_STOCK


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
    """販売予定/実売の税込総額列と粗利を再計算する（数値列は int で統一）。"""
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
    for i in out.index:
        cogs = _finite_int(out.at[i, COL_PRICE_EXCL], 0)
        line_in = _finite_int(out.at[i, COL_PRICE_INCL], 0)
        qty = max(1, _finite_int(out.at[i, COL_QTY], 1))
        pl_u = _finite_int(out.at[i, COL_PLANNED_SALE], 0)
        ac_u = _finite_int(out.at[i, COL_ACTUAL_SALE], 0)
        stt = _normalize_stock_status(str(out.at[i, COL_STOCK_STATUS]))
        out.at[i, COL_STOCK_STATUS] = stt
        tax_r = _infer_tax_rate_from_main_line(cogs, line_in)
        plex, pincl, aex, aincl = _planned_actual_line_amounts(
            qty, pl_u, ac_u, stt, tax_r
        )
        pincl_i = _finite_int(pincl, 0)
        aincl_i = _finite_int(aincl, 0)
        out.at[i, COL_PLANNED_SALE_INCL] = pincl_i if pincl_i > 0 else 0
        out.at[i, COL_ACTUAL_SALE_INCL] = aincl_i if aincl_i > 0 else 0
        gp = _compute_gross_profit_row(cogs, plex, aex, stt)
        out.at[i, COL_GROSS_PROFIT] = 0 if gp is None else _finite_int(gp, 0)
    return out


def _parse_max_management_serial(rows: list[Any], col_idx: int) -> int:
    """管理ID列の最大連番（G######## または数字のみ）を返す。ヘッダー行は含めない。"""
    mx = 0
    for r in rows[1:]:
        row = [("" if c is None else str(c)) for c in list(r)]
        if col_idx >= len(row):
            continue
        s = str(row[col_idx]).strip()
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
    """管理ID（G########）を count 件、シート現状から連番で採番する。"""
    if count <= 0:
        return []
    try:
        raw = ws.get_all_values()
    except Exception:
        raw = []
    if not raw:
        raw = [EXPECTED_HEADERS]
    idx = EXPECTED_HEADERS.index(COL_MANAGEMENT_ID)
    mx = _parse_max_management_serial(raw, idx)
    return [f"G{mx + i + 1:08d}" for i in range(count)]


def append_sheet_row(
    movement: str,
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
):
    """1点1行で台帳に追記する（数量列は常に 1。仕入単価列は持たない）。"""
    ws = ensure_worksheet_header()
    if ws is None:
        st.warning("スプレッドシート未設定のため、行の追記をスキップしました。")
        return
    now = (record_datetime or "").strip() or jst_now_str()
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
    try:
        ws.append_row(
            [
                now,
                movement,
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
            ],
            value_input_option="USER_ENTERED",
        )
        try:
            _apply_inventory_amount_number_formats(ws)
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(f"スプレッドシート追記に失敗しました: {e}") from e


def load_inventory_dataframe() -> pd.DataFrame | None:
    """1行目をヘッダー、2行目以降をデータとして読み込み、列は EXPECTED_HEADERS に揃える。"""
    ws = _get_or_create_inventory_worksheet()
    if ws is None:
        return None
    try:
        raw = ws.get_all_values()
    except Exception:
        return None
    if not raw:
        return pd.DataFrame(columns=EXPECTED_HEADERS)
    header0 = [("" if c is None else str(c)).strip() for c in raw[0]]
    rows = raw[1:]
    n = len(EXPECTED_HEADERS)
    drop_unit_j: int | None = None
    if LEGACY_COL_UNIT_PRICE in header0:
        drop_unit_j = header0.index(LEGACY_COL_UNIT_PRICE)

    def pad(row: list[Any]) -> list[str]:
        r = [("" if c is None else str(c)) for c in list(row)]
        if drop_unit_j is not None and drop_unit_j < len(r):
            r = r[:drop_unit_j] + r[drop_unit_j + 1 :]
        if len(r) < n:
            r.extend([""] * (n - len(r)))
        return r[:n]

    data_rows = [pad(r) for r in rows]
    return pd.DataFrame(data_rows, columns=EXPECTED_HEADERS)


def _ledger_unique_col_values(df: pd.DataFrame, col: str, *, max_n: int = 800) -> list[str]:
    """台帳 DataFrame から列のユニーク値（空除く）を昇順で返す。"""
    if df is None or df.empty or col not in df.columns:
        return []
    s = df[col].astype(str).str.strip()
    s = s[s != ""]
    return sorted(set(s.tolist()), key=lambda x: (x.casefold(), x))[:max_n]


def _on_ledger_pick_product_name() -> None:
    v = st.session_state.get("ledger_pick_product_name", "")
    if v and v != LEDGER_PICK_PLACEHOLDER:
        st.session_state.field_product_name = v
        st.session_state.ledger_pick_product_name = LEDGER_PICK_PLACEHOLDER


def _on_ledger_pick_supplier() -> None:
    v = st.session_state.get("ledger_pick_supplier", "")
    if v and v != LEDGER_PICK_PLACEHOLDER:
        st.session_state.field_supplier = v
        st.session_state.ledger_pick_supplier = LEDGER_PICK_PLACEHOLDER


def _cell_value_for_sheet(v: Any) -> Any:
    try:
        if pd.api.types.is_scalar(v) and pd.isna(v):
            return 0
        if isinstance(v, (float, np.floating)) and (
            not math.isfinite(float(v)) or pd.isna(v)
        ):
            return 0
    except Exception:
        pass
    return v


def overwrite_inventory_worksheet_from_dataframe(df: pd.DataFrame) -> None:
    """編集後の DataFrame の内容でワークシートをクリアし、ヘッダー＋全行を書き直す。"""
    ws = _get_or_create_inventory_worksheet()
    if ws is None:
        raise RuntimeError(
            f"スプレッドシートに接続できません。{SECRET_GOOGLE_SPREADSHEET_ID} とサービスアカウントを確認してください。"
        )
    out = _recalc_gross_profit_dataframe(df.reindex(columns=EXPECTED_HEADERS, fill_value="").copy())
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


def _apply_ledger_sort(
    df: pd.DataFrame,
    primary: str,
    primary_asc: bool,
    secondary: str,
    secondary_asc: bool,
) -> pd.DataFrame:
    """日時・仕入先・取引先による表示用ソート（コピーを返す）。"""
    if df.empty:
        return df
    col_map = {"日時": COL_DATETIME, "仕入先・取引先": COL_SUPPLIER}
    pairs: list[tuple[str, bool]] = []
    if primary != "なし" and primary in col_map:
        pairs.append((col_map[primary], primary_asc))
    if secondary != "なし" and secondary in col_map:
        c2 = col_map[secondary]
        if not pairs or pairs[0][0] != c2:
            pairs.append((c2, secondary_asc))
    if not pairs:
        return df.copy()

    out = df.copy()
    sort_cols: list[str] = []
    ascending: list[bool] = []
    for col, asc in pairs:
        if col not in out.columns:
            continue
        if col == COL_DATETIME:
            tmp = "_sort_dt_internal"
            out[tmp] = pd.to_datetime(out[col], errors="coerce")
            sort_cols.append(tmp)
        else:
            sort_cols.append(col)
        ascending.append(asc)
    if not sort_cols:
        return out
    out = out.sort_values(by=sort_cols, ascending=ascending, na_position="last")
    return out.drop(columns=[c for c in out.columns if c.startswith("_sort_dt")], errors="ignore")


def _prepare_ledger_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """入出庫判定・行金額・年月などの派生列を付与したコピーを返す。"""
    d = df.copy()
    if d.empty:
        return d
    d[COL_DATETIME] = pd.to_datetime(d[COL_DATETIME], errors="coerce")
    qty = _series_to_numeric_loose(d[COL_QTY]).fillna(0)
    line_stored_ex = _series_to_numeric_loose(d[COL_PRICE_EXCL]).fillna(0)
    line_stored_in = _series_to_numeric_loose(d[COL_PRICE_INCL]).fillna(0)
    movement = d[COL_TYPE].astype(str).str.strip()
    is_in = movement.str.startswith("入庫")
    is_out = movement.str.startswith("出庫")
    # 税抜・税込はいずれも行合計として解釈（仕入単価列は廃止）
    line_ex = line_stored_ex.astype(float)
    # 税抜が取り込めず 0 のままだが税込と数量がある行は、税込から税抜を逆算する
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
    d["_qty_in"] = qty.where(is_in, 0.0).fillna(0).astype(float)
    d["_qty_out"] = qty.where(is_out, 0.0).fillna(0).astype(float)
    d["_amt_ex_in"] = line_ex.where(is_in, 0.0).fillna(0).astype(float)
    d["_amt_ex_out"] = line_ex.where(is_out, 0.0).fillna(0).astype(float)
    d["_amt_in_in"] = line_in.where(is_in, 0.0).fillna(0).astype(float)
    d["_amt_in_out"] = line_in.where(is_out, 0.0).fillna(0).astype(float)
    d["_ym"] = d[COL_DATETIME].dt.to_period("M").astype(str)
    d["_year"] = d[COL_DATETIME].dt.year
    d["_month"] = d[COL_DATETIME].dt.month
    return d


def _ledger_dashboard_date_bounds(df: pd.DataFrame) -> tuple[date, date]:
    """台帳の日時列から From/To の既定値（JST 日付）。有効な日付が無いときは今日（JST）。"""
    s = pd.to_datetime(df[COL_DATETIME], errors="coerce").dropna()
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
    st.subheader("集計・ダッシュボード")
    st.caption(
        "上の表の現在の内容（未保存の編集を含む）を集計します。"
        f"金額はシートの「{COL_PRICE_EXCL}」「{COL_PRICE_INCL}」列を行合計として集計します。"
        f"仕入先・取引先別の粗利は「{COL_GROSS_PROFIT}」列を合算しています（税抜・台帳保存時の値）。"
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

    row_ts = pd.to_datetime(ad_f[COL_DATETIME], errors="coerce")
    row_day = row_ts.dt.normalize()
    from_ts = pd.Timestamp(datetime.combine(dfb, datetime.min.time()))
    to_ts = pd.Timestamp(datetime.combine(dtb, datetime.min.time()))
    flt = ad_f[(row_day >= from_ts) & (row_day <= to_ts)]
    if supplier_filter:
        flt = flt[flt[COL_SUPPLIER].astype(str).isin(supplier_filter)]

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
    g = flt.assign(**{sup_col: flt[COL_SUPPLIER].fillna("(未設定)").astype(str)})
    _agg_sup: dict[str, tuple[str, str]] = {
        "入庫数量": ("_qty_in", "sum"),
        "出庫数量": ("_qty_out", "sum"),
        "入庫金額税抜": ("_amt_ex_in", "sum"),
        "出庫金額税抜": ("_amt_ex_out", "sum"),
    }
    if COL_GROSS_PROFIT in g.columns:
        _agg_sup["粗利合計"] = (COL_GROSS_PROFIT, "sum")
    grp = g.groupby(sup_col, dropna=False).agg(**_agg_sup).reset_index()
    for _gc in ("入庫数量", "出庫数量", "入庫金額税抜", "出庫金額税抜", "粗利合計"):
        if _gc in grp.columns:
            grp[_gc] = (
                pd.to_numeric(grp[_gc], errors="coerce")
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0.0)
            )
    grp["差し引き数量"] = (grp["入庫数量"] - grp["出庫数量"]).astype(int)
    grp["差し引き税抜"] = (grp["入庫金額税抜"] - grp["出庫金額税抜"]).round(0).astype(int)
    st.dataframe(
        grp.sort_values(sup_col),
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

    st.markdown("##### 金額推移（税抜・折れ線）")
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
            "並びは粗利の絶対値が大きい順です（マイナスも含みます）。"
        )
        gp_ch = (
            grp[[sup_col, "粗利合計"]]
            .assign(
                粗利合計=lambda x: pd.to_numeric(
                    x["粗利合計"], errors="coerce"
                ).fillna(0.0)
            )
            .assign(_abs=lambda x: x["粗利合計"].abs())
            .sort_values("_abs", ascending=False)
            .drop(columns=["_abs"])
            .head(15)
        )
        if not gp_ch.empty:
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


def _render_inventory_price_summary(df: pd.DataFrame) -> None:
    """在庫中の行について、合計原価・販売予定（税抜行計・税込総額）・想定粗利を表示する。"""
    st.markdown("##### 価格管理サマリー（在庫中）")
    st.caption(
        "上の表のうち、ステータスが「在庫中」の行だけを合算しています（未保存の編集を含みます）。"
        "販売予定は単価×数量の税抜行計、税込列は仕入行と同じ税率で算出した値の合計です。"
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


def render_inventory_manager() -> None:
    st.divider()
    st.markdown("##### 在庫一覧")
    st.caption(
        "スプレッドシートの全データを編集できます。行の追加・削除は表から操作し、"
        "表の直下の「台帳を更新する」でシートを上書き保存します。"
    )

    if msg := st.session_state.pop("_ledger_saved_flash", None):
        st.success(msg)

    if not _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        st.info(
            f"{SECRET_GOOGLE_SPREADSHEET_ID} を設定すると、台帳の表示・編集ができます。"
        )
        return

    r1, _ = st.columns([1, 2])
    with r1:
        if st.button("スプレッドシートから再読込", key="ledger_reload_from_sheet"):
            st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
            st.rerun()

    try:
        df_sheet = load_inventory_dataframe()
    except Exception as e:
        st.error(f"読み込みに失敗しました: {e}")
        return

    if df_sheet is None:
        st.warning("スプレッドシートを開けませんでした。サービスアカウントと共有設定を確認してください。")
        return

    st.markdown("##### 表示の並び順（台帳表に反映・保存時もこの順で書き込みます）")
    s1, s2, s3, s4 = st.columns([2, 1, 2, 1])
    sort_choices = ["日時", "仕入先・取引先", "なし"]
    with s1:
        prim = st.selectbox("第1ソート", sort_choices, index=0, key="ledger_sort_p")
    with s2:
        prim_ord = st.radio("第1の順序", ["昇順", "降順"], horizontal=True, key="ledger_sort_p_ord")
    with s3:
        sec = st.selectbox("第2ソート", sort_choices, index=2, key="ledger_sort_s")
    with s4:
        sec_ord = st.radio("第2の順序", ["昇順", "降順"], horizontal=True, key="ledger_sort_s_ord")

    df_sorted = _apply_ledger_sort(
        df_sheet,
        prim,
        prim_ord == "昇順",
        sec,
        sec_ord == "昇順",
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

    _editor_kw: dict[str, Any] = {
        "num_rows": "dynamic",
        "key": LEDGER_DATA_EDITOR_KEY,
        "use_container_width": True,
        "hide_index": True,
    }
    if _ledger_col_cfg:
        _editor_kw["column_config"] = _ledger_col_cfg
    edited = st.data_editor(df_sorted, **_editor_kw)

    _render_inventory_price_summary(edited)

    if st.button("台帳を更新する", type="primary", key="ledger_save_overwrite"):
        with st.spinner("スプレッドシートに書き込んでいます…"):
            try:
                overwrite_inventory_worksheet_from_dataframe(edited)
            except Exception as e:
                st.error(str(e))
                return
        st.session_state["_ledger_saved_flash"] = "台帳を更新しました。"
        st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
        st.rerun()

    render_ledger_dashboard(edited)


def _init_registration_form_session_state() -> None:
    """登録フォーム用の session_state 初期値（キーはウィジェットと連動）。"""
    if "field_product_name" not in st.session_state:
        st.session_state.field_product_name = ""
    if "field_supplier" not in st.session_state:
        st.session_state.field_supplier = ""
    if "field_qty" not in st.session_state:
        st.session_state.field_qty = 1
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
    st.session_state.pop("field_price_excl", None)


def main():
    st.set_page_config(page_title="商品在庫・販売", layout="wide")
    st.title("商品在庫・販売管理")
    st.caption("写真は任意。台帳の必須項目のみの記録、または写真＋AI解析・ドライブ保存・スプレッドシート記録ができます。")
    st.subheader("台帳登録")

    _init_registration_form_session_state()

    uploaded = st.file_uploader(
        "商品写真（任意・1枚まで・カメラやギャラリーから）",
        type=["jpg", "jpeg", "png", "webp"],
    )
    st.caption(
        "写真は **1枚まで** です。数量が **2以上** のときは、その1枚をドライブに保存し、"
        "作成する **全行に同じ画像URL** を入れます。"
        f"写真がある場合のみ、EXIF向き補正のうえ長辺最大{UPLOAD_JPEG_MAX_LONG_EDGE}px・"
        f"JPEG品質{UPLOAD_JPEG_QUALITY}％へ変換してから AI 解析・ドライブ保存します。"
        "台帳の日時は写真の EXIF 撮影日時を優先し、写真がないときは日本時間（JST）の現在時刻です。"
        "必須項目だけでも確定して台帳記録できます（画像URLは空欄になります）。"
    )

    movement = st.radio(
        "区分",
        ("入庫（購入）", "入庫（返品）", "出庫（販売）", "出庫（浮貸）"),
        horizontal=True,
    )

    col_a, col_b = st.columns(2)
    with col_a:
        analyze = st.button(
            "AIで画像を解析",
            type="primary",
            disabled=uploaded is None,
        )
    with col_b:
        if st.button("候補の自動入力をクリア"):
            st.session_state.field_product_name = ""
            st.session_state.field_supplier = ""
            st.session_state.field_qty = 1
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
            st.rerun()

    if analyze and uploaded is not None:
        with st.spinner("画像を解析しています…"):
            try:
                img = _gemini_input_image_from_upload(uploaded)
                raw_text = analyze_image_with_gemini(img)
                result = _parse_json_from_model(raw_text or "")
                _apply_gemini_json_to_session(result)
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

    df_ledger_hint: pd.DataFrame | None = None
    if _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        try:
            df_ledger_hint = load_inventory_dataframe()
        except Exception:
            df_ledger_hint = None

    if df_ledger_hint is not None and not df_ledger_hint.empty:
        st.markdown("##### 台帳から入力補助（任意）")
        st.caption(
            "絞り込み欄に文字を入れると候補が絞られます。プルダウンで選ぶと下の入力欄に反映されます（あとから手修正も可能です）。"
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
    elif _secret_str(SECRET_GOOGLE_SPREADSHEET_ID):
        st.caption("台帳が空か読み込めないため、入力補助の候補は表示できません。")

    st.markdown("##### 必須入力項目")
    product_name = st.text_input("商品名（必須）", key="field_product_name")
    supplier = st.text_input("仕入先・取引先（必須）", key="field_supplier")
    quantity = st.number_input(
        "数量（点数）",
        min_value=1,
        step=1,
        key="field_qty",
        help="台帳は **1点1行** で保存します。行数は常にこの数量と同じです（写真は1枚まで・複数点のときは同じ画像URLを各行に入れます）。",
    )

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
        st.caption(
            f"確定時は **{_n_save} 行**（各行 数量1）。税抜合計（参考） ¥{_line_ex_one * _n_save:,}。"
            f"写真があるとき、数量が2以上なら **同じ画像URLを全行** に記録します。"
        )
    with price_row[1]:
        st.metric("仕入金額（税込・1点・自動）", f"¥{_line_in_one:,}")
        _tl = st.session_state.get("field_consumption_tax_choice", "10%")
        if _tl == "非課税":
            st.caption("非課税のため税込＝税抜行合計")
        else:
            st.caption(f"消費税{_tl}を行合計に四捨五入")
    with price_row[2]:
        st.caption(
            "原価は各行の仕入金額（税抜）です。販売予定・実売は **1点あたり税抜単価** を入力し、台帳では各行数量1として税抜行計と税込総額を記録します。"
        )

    st.markdown("##### 価格管理（任意）")
    st.caption(
        "販売予定・実売は **1点あたりの税抜単価（円）** です。税込総額は「単価×数量」した税抜合計に、上の消費税と同じ税率を掛けて四捨五入します。"
        "ステータスが「販売済」のときのみ実売単価・実売金額（税込）を記録し、粗利は税抜で「実売行計−原価」です。"
        "「在庫中」のときは「販売予定行計−原価」で粗利を計算します。"
    )
    planned_sale_excl = st.number_input(
        "販売予定単価（税抜・任意）",
        min_value=0,
        step=1,
        key="field_planned_sale_excl",
        help="1点あたり。0 のとき台帳では空欄。税抜行計・税込総額は各行数量1として自動計算します。",
    )
    st.selectbox(
        "ステータス（在庫中／販売済）",
        options=list(STOCK_STATUS_OPTIONS),
        key="field_stock_status",
    )
    _st = str(st.session_state.get("field_stock_status", STATUS_IN_STOCK)).strip()
    actual_sale_excl = st.number_input(
        "実売単価（税抜・任意）",
        min_value=0,
        step=1,
        key="field_actual_sale_excl",
        disabled=(_st != STATUS_SOLD),
        help="1点あたり。ステータスが「販売済」のときのみ入力・記録されます。",
    )
    _pl_u = int(planned_sale_excl)
    _act_u = int(actual_sale_excl)
    _cogs_preview = _lex_inp
    _plex, _pin, _aex, _ain = _planned_actual_line_amounts(
        1, _pl_u, _act_u, _st, _tax_r
    )
    _gp_preview = _compute_gross_profit_row(
        _cogs_preview,
        _plex,
        _aex if _st == STATUS_SOLD else 0,
        _st,
    )
    pm1, pm2, pm3, pm4, pm5 = st.columns(5)
    with pm1:
        st.metric("原価（税抜・1点）", f"¥{_cogs_preview:,}")
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
        "確定（スプレッドシート記録・写真は任意でドライブ保存）",
        type="primary",
    )

    if confirm:
        validation_ok = True
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
            _act_ex2 = int(st.session_state.get("field_actual_sale_excl", 0))
            _stat2 = str(
                st.session_state.get("field_stock_status", STATUS_IN_STOCK)
            ).strip()
            if _stat2 not in STOCK_STATUS_OPTIONS:
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
                        data0, mime0 = prepare_upload_image_jpeg(raw0)
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
                ws0 = ensure_worksheet_header()
                if ws0 is None:
                    st.warning("スプレッドシート未設定のため、行の追記をスキップしました。")
                else:
                    try:
                        ids = allocate_management_ids(ws0, n_save)
                        with st.spinner("スプレッドシートに記録しています…"):
                            for i in range(n_save):
                                append_sheet_row(
                                    movement,
                                    product_name.strip(),
                                    supplier.strip(),
                                    _lex_one,
                                    _lin_one,
                                    urls[i],
                                    ids[i],
                                    memo_s,
                                    record_datetime=_record_dt,
                                    planned_sale_unit_excl_yen=_plan2,
                                    actual_sale_unit_excl_yen=_act_ex2,
                                    stock_status=_stat2,
                                    consumption_tax_rate=_tax_r2,
                                )
                    except Exception as e:
                        st.error(f"スプレッドシート更新に失敗しました: {e}")
                        if any(urls):
                            st.warning(
                                "一部の画像はドライブに保存済みの可能性があります。台帳の内容を確認してください。"
                            )
                    else:
                        st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                        st.success(f"記録しました（{n_save} 行・1点1行）。管理IDを自動付与しています。")
                        _link_urls = list(dict.fromkeys(u for u in urls if u))
                        for _uurl in _link_urls[:8]:
                            st.markdown(f"[保存した画像を開く]({_uurl})")
                        if len(_link_urls) > 8:
                            st.caption(f"ほか {len(_link_urls) - 8} 件の画像URLは台帳の「{COL_IMAGE_URL}」列を参照してください。")
                        st.balloons()

    render_inventory_manager()


if __name__ == "__main__":
    if check_password():
        main()
