"""
呉服在庫・販売管理 (Streamlit)

st.secrets に以下を設定してください（例は .streamlit/secrets.toml）。

必須キー:
  GEMINI_API_KEY
  GAS_UPLOAD_URL           … 画像を Google ドライブに保存する Web アプリ（GAS）の URL
  GAS_API_KEY              … GAS Web アプリ呼び出し用の共有キー（payload の apiKey に付与）
  GOOGLE_DRIVE_FOLDER_ID   … 保存先フォルダID（GAS に渡す）
  GOOGLE_SPREADSHEET_ID    … 記録用スプレッドシートID
  google_service_account   … サービスアカウントJSONの各フィールド（[google_service_account] セクション）
    または GOOGLE_SERVICE_ACCOUNT_JSON … JSON文字列1本

任意:
  GOOGLE_WORKSHEET_NAME    … ワークシート名（既定: 在庫履歴）
  APP_PASSWORD             … アプリ画面の簡易ログイン用（平文。GitHub には secrets.toml をコミットしないこと）

※ 画像の Gemini 解析は **google-generativeai** で、まず ``genai.GenerativeModel('gemini-1.5-flash-8b')`` を試し、
  404 等でモデルが無い場合のみ ``gemini-1.5-pro`` に切り替えます（``models/``・api_version は指定しません）。
※ アップロード画像は Pillow で長辺最大1280px・JPEG品質80に変換したうえで解析・ドライブ保存します。
※ 台帳日時・撮影日時未取得時の現在時刻は **pytz** の ``Asia/Tokyo``（JST）です。

画面下部の「在庫一覧マネージャー」で、同一スプレッドシートを表形式で読み書きし、
入出庫の集計・仕入先・取引先別サマリー・月次グラフを表示できます。

スプレッドシート1行目はヘッダーとして次の列順を想定:
  日時 | 入出庫種別 | 商品名 | 仕入先・取引先 | 数量 | 商品単価（税抜） | 商品金額（税抜） | 税込金額 | 画像URL | メモ（任意）
  ※「日時」列への新規記入は **日本時間（JST / Asia/Tokyo）** で行い、画像に EXIF 撮影日時があればそれを JST として解釈して優先します。
  ※商品金額（税抜）は「数量×商品単価」の行合計。旧データは単価列が空のとき従来どおり数量×金額列で集計します。
"""

from __future__ import annotations

import streamlit as st


def check_password() -> bool:
    """認証済みになるまで在庫アプリ本体を起動しない。未認証時は認証UIのみ表示し st.stop() する。"""
    if st.session_state.get("authenticated"):
        return True

    st.set_page_config(page_title="認証", layout="centered")
    st.header("認証")

    try:
        expected = str(st.secrets["APP_PASSWORD"]).strip()
    except Exception:
        st.error(
            "`.streamlit/secrets.toml` に **APP_PASSWORD** を設定してください。"
            "（ローカルで `secrets.toml` が無い場合は作成してください）"
        )
        st.stop()

    if not expected:
        st.error("APP_PASSWORD が空です。secrets.toml を確認してください。")
        st.stop()

    with st.form("auth_screen_form"):
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン（パスワード入力後は Enter でも送信できます）")

    if submitted:
        if (password or "").strip() == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("パスワードが正しくありません。")

    st.stop()


import base64
import io
import json
import re
import uuid
from datetime import datetime
from typing import Any

import pytz

import altair as alt
import pandas as pd
import google.api_core.exceptions as google_api_exceptions
import google.generativeai as genai
import gspread
import requests
from google.oauth2 import service_account
from PIL import Image, ImageOps


# --- スプレッドシート列（ヘッダーと一致させる） ---
COL_DATETIME = "日時"
COL_TYPE = "入出庫種別"
COL_NAME = "商品名"
COL_SUPPLIER = "仕入先・取引先"
COL_QTY = "数量"
COL_PRICE_UNIT = "商品単価（税抜）"
COL_PRICE_EXCL = "商品金額（税抜）"
COL_PRICE_INCL = "税込金額"
COL_IMAGE_URL = "画像URL"
COL_MEMO = "メモ"

# 消費税の自動計算（標準税率）。軽減税率の品目は手入力・メモで補足してください。
CONSUMPTION_TAX_RATE = 0.10

# Gemini 画像解析（プレフィックスなし。404 時はフォールバック）
_GEMINI_VISION_PRIMARY = "gemini-1.5-flash-8b"
_GEMINI_VISION_FALLBACK = "gemini-1.5-pro"

EXPECTED_HEADERS = [
    COL_DATETIME,
    COL_TYPE,
    COL_NAME,
    COL_SUPPLIER,
    COL_QTY,
    COL_PRICE_UNIT,
    COL_PRICE_EXCL,
    COL_PRICE_INCL,
    COL_IMAGE_URL,
    COL_MEMO,
]

# 一時的: secrets.toml が無い／空でもアプリを落とさない（AI 解析テスト用）
_PLACEHOLDER_DRIVE_URL = "https://example.com/?gofuku-app=skipped-no-gas-secrets"

# アプリ全体の基準タイムゾーン（台帳の「日時」・ファイル名・EXIF 未取得時のデフォルト）
TZ_JP = pytz.timezone("Asia/Tokyo")


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


# 解析・ドライブ保存共通: Pillow で長辺リサイズ + JPEG 再エンコード（データ量削減）
_UPLOAD_JPEG_MAX_LONG_EDGE = 1280
_UPLOAD_JPEG_QUALITY = 80


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

    EXIF 向き補正のうえ長辺を最大 1280px に収め、JPEG 品質 80% で再エンコードする。

    Returns:
        (jpeg_bytes, mime_type)  mime_type は常に ``image/jpeg`` 。
    """
    img = Image.open(io.BytesIO(raw))
    img = ImageOps.exif_transpose(img)
    rgba = img.convert("RGBA")
    bg = Image.new("RGB", rgba.size, (255, 255, 255))
    bg.paste(rgba, mask=rgba.getchannel("A"))
    img = bg
    img = _resize_long_edge_max(img, _UPLOAD_JPEG_MAX_LONG_EDGE)
    buf = io.BytesIO()
    img.save(
        buf,
        format="JPEG",
        quality=_UPLOAD_JPEG_QUALITY,
        optimize=True,
        progressive=True,
    )
    return buf.getvalue(), "image/jpeg"


def _safe_secret(key: str, default: str = "") -> str:
    try:
        v = st.secrets[key]
        if v is None:
            return default
        s = str(v).strip()
        return s if s else default
    except Exception:
        return default


def _get_secrets() -> dict[str, Any]:
    return dict(st.secrets)


def _load_service_account_info() -> dict[str, Any]:
    s = _get_secrets()
    if "GOOGLE_SERVICE_ACCOUNT_JSON" in s:
        raw = s["GOOGLE_SERVICE_ACCOUNT_JSON"]
        if isinstance(raw, str):
            return json.loads(raw)
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON は文字列である必要があります。")
    if "google_service_account" in s:
        ga = s["google_service_account"]
        if hasattr(ga, "to_dict"):
            return dict(ga.to_dict())
        return dict(ga)
    raise ValueError(
        "サービスアカウントが見つかりません。"
        " [google_service_account] セクションか GOOGLE_SERVICE_ACCOUNT_JSON を設定してください。"
    )


def _credentials():
    info = _load_service_account_info()
    scopes = ("https://www.googleapis.com/auth/spreadsheets",)
    return service_account.Credentials.from_service_account_info(info, scopes=scopes)


@st.cache_resource
def _gspread_client():
    return gspread.authorize(_credentials())


def _parse_json_from_model(text: str) -> dict[str, Any]:
    """モデル出力から JSON オブジェクトを抽出する。"""
    t = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", t)
    if fence:
        t = fence.group(1).strip()
    return json.loads(t)


def _gemini_input_image_from_upload(uploaded) -> Image.Image:
    """解析直前に ``prepare_upload_image_jpeg`` と同じ圧縮・リサイズを適用した PIL 画像を返す。"""
    jpeg_bytes, _ = prepare_upload_image_jpeg(uploaded.getvalue())
    return Image.open(io.BytesIO(jpeg_bytes)).convert("RGB")


def price_incl_tax(price_excl_yen: int) -> int:
    """税抜き円金額から、標準税率を乗じた税込円金額（四捨五入）。"""
    return int(round(int(price_excl_yen) * (1 + CONSUMPTION_TAX_RATE)))


def _gemini_model_unavailable(exc: BaseException) -> bool:
    """404 / モデル廃止など、別モデルへの切り替えが妥当なエラーか。"""
    if isinstance(exc, google_api_exceptions.NotFound):
        return True
    msg = str(exc).lower()
    return any(
        x in msg
        for x in ("404", "not found", "no longer available", "not_found", "is not found")
    )


def analyze_image_with_gemini(image_data):
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    prompt = "この呉服の画像を解析し、商品名、色、柄、素材、状態を推定してJSON形式で返してください。"
    contents = [prompt, image_data]
    models = (_GEMINI_VISION_PRIMARY, _GEMINI_VISION_FALLBACK)
    for i, model_id in enumerate(models):
        try:
            model = genai.GenerativeModel(model_id)
            response = model.generate_content(contents)
            return response.text or ""
        except Exception as e:
            if i < len(models) - 1 and _gemini_model_unavailable(e):
                continue
            raise


def ensure_worksheet_header():
    """1行目がヘッダーでなければ作成（初回のみ想定）。secrets 未設定時は None。"""
    sid = _safe_secret("GOOGLE_SPREADSHEET_ID")
    if not sid:
        return None
    wname = _safe_secret("GOOGLE_WORKSHEET_NAME", "在庫履歴")
    try:
        sh = _gspread_client().open_by_key(str(sid))
        try:
            ws = sh.worksheet(str(wname))
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=str(wname), rows=1000, cols=10)

        first = ws.row_values(1)
        if not first or first[: len(EXPECTED_HEADERS)] != EXPECTED_HEADERS:
            ws.update("A1", [EXPECTED_HEADERS], value_input_option="USER_ENTERED")
        return ws
    except Exception:
        return None


def upload_image_to_drive(filename: str, mime: str, data: bytes) -> str:
    """Google Apps Script 経由で Google ドライブに保存し、閲覧用 URL を返す。

    GAS 側は JSON で
    { folderId, fileName, mimeType, base64Data, apiKey } を受け取り、
    { status: \"success\", url: \"...\" } 形式で返す想定。
    """
    gas_url = _safe_secret("GAS_UPLOAD_URL")
    gas_api_key = _safe_secret("GAS_API_KEY")
    folder_id = _safe_secret("GOOGLE_DRIVE_FOLDER_ID")
    if not gas_url or not gas_api_key or not folder_id:
        return _PLACEHOLDER_DRIVE_URL

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
        timeout=300,
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


def append_sheet_row(
    movement: str,
    product_name: str,
    supplier: str,
    quantity: int,
    unit_price_excl_yen: int,
    line_price_excl_yen: int,
    line_price_incl_yen: int,
    image_url: str,
    memo: str = "",
    record_datetime: str | None = None,
):
    ws = ensure_worksheet_header()
    if ws is None:
        st.warning("スプレッドシート未設定のため、行の追記をスキップしました。")
        return
    now = (record_datetime or "").strip() or jst_now_str()
    try:
        ws.append_row(
            [
                now,
                movement,
                product_name,
                supplier,
                quantity,
                unit_price_excl_yen,
                line_price_excl_yen,
                line_price_incl_yen,
                image_url,
                memo,
            ],
            value_input_option="USER_ENTERED",
        )
    except Exception as e:
        st.warning(f"スプレッドシート更新をスキップしました: {e}")


LEDGER_DATA_EDITOR_KEY = "inventory_ledger_data_editor"


def _open_inventory_worksheet():
    """在庫ワークシートを開く。存在しなければ作成する。失敗時は None。"""
    sid = _safe_secret("GOOGLE_SPREADSHEET_ID")
    if not sid:
        return None
    wname = _safe_secret("GOOGLE_WORKSHEET_NAME", "在庫履歴")
    try:
        sh = _gspread_client().open_by_key(str(sid))
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


def load_inventory_dataframe() -> pd.DataFrame | None:
    """1行目をヘッダー、2行目以降をデータとして読み込み、列は EXPECTED_HEADERS に揃える。"""
    ws = _open_inventory_worksheet()
    if ws is None:
        return None
    try:
        raw = ws.get_all_values()
    except Exception:
        return None
    if not raw:
        return pd.DataFrame(columns=EXPECTED_HEADERS)
    rows = raw[1:]
    n = len(EXPECTED_HEADERS)

    def pad(row: list[Any]) -> list[str]:
        r = [("" if c is None else str(c)) for c in list(row)]
        if len(r) < n:
            r.extend([""] * (n - len(r)))
        return r[:n]

    data_rows = [pad(r) for r in rows]
    return pd.DataFrame(data_rows, columns=EXPECTED_HEADERS)


def _cell_value_for_sheet(v: Any) -> Any:
    try:
        if pd.api.types.is_scalar(v) and pd.isna(v):
            return ""
    except Exception:
        pass
    return v


def overwrite_inventory_worksheet_from_dataframe(df: pd.DataFrame) -> None:
    """編集後の DataFrame の内容でワークシートをクリアし、ヘッダー＋全行を書き直す。"""
    ws = _open_inventory_worksheet()
    if ws is None:
        raise RuntimeError(
            "スプレッドシートに接続できません。GOOGLE_SPREADSHEET_ID とサービスアカウントを確認してください。"
        )
    out = df.reindex(columns=EXPECTED_HEADERS, fill_value="").copy()
    values: list[list[Any]] = [EXPECTED_HEADERS]
    if not out.empty:
        for row in out[EXPECTED_HEADERS].to_numpy(dtype=object):
            values.append([_cell_value_for_sheet(x) for x in row.tolist()])
    try:
        ws.clear()
        ws.update("A1", values, value_input_option="USER_ENTERED")
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
    qty = pd.to_numeric(d[COL_QTY], errors="coerce").fillna(0)
    unit_ex = (
        pd.to_numeric(d[COL_PRICE_UNIT], errors="coerce").fillna(0)
        if COL_PRICE_UNIT in d.columns
        else pd.Series(0.0, index=d.index)
    )
    line_stored_ex = pd.to_numeric(d[COL_PRICE_EXCL], errors="coerce").fillna(0)
    line_stored_in = pd.to_numeric(d[COL_PRICE_INCL], errors="coerce").fillna(0)
    movement = d[COL_TYPE].astype(str).str.strip()
    is_in = movement.str.startswith("入庫")
    is_out = movement.str.startswith("出庫")
    # 単価列あり: 行の金額列を優先（台帳で修正した値を尊重）。単価列なし: 旧形式（金額列=単価）として数量倍
    line_ex = line_stored_ex.where(unit_ex > 0, qty * line_stored_ex)
    line_in = line_stored_in.where(unit_ex > 0, qty * line_stored_in)
    d["_qty_in"] = qty.where(is_in, 0.0)
    d["_qty_out"] = qty.where(is_out, 0.0)
    d["_amt_ex_in"] = line_ex.where(is_in, 0.0)
    d["_amt_ex_out"] = line_ex.where(is_out, 0.0)
    d["_amt_in_in"] = line_in.where(is_in, 0.0)
    d["_amt_in_out"] = line_in.where(is_out, 0.0)
    d["_ym"] = d[COL_DATETIME].dt.to_period("M").astype(str)
    d["_year"] = d[COL_DATETIME].dt.year
    d["_month"] = d[COL_DATETIME].dt.month
    return d


def render_ledger_dashboard(df: pd.DataFrame) -> None:
    """在庫台帳 DataFrame から入出庫集計・仕入先・取引先別・グラフを表示する。"""
    st.subheader("集計・ダッシュボード")
    st.caption(
        "上の表の現在の内容（未保存の編集を含む）を集計します。"
        "金額はシートの「商品金額（税抜）」「税込金額」列を行合計として集計します。"
        "（単価列が空の旧行は、金額列を単価として数量倍します。）"
    )
    if df.empty:
        st.info("集計する行がありません。")
        return

    ad = _prepare_ledger_analysis(df)
    ad_f = ad.dropna(subset=[COL_DATETIME], how="all")

    p1, p2, p3, p4 = st.columns(4)
    with p1:
        period = st.radio(
            "期間",
            ("全期間", "年を指定", "月を指定"),
            horizontal=True,
            key="dash_period_mode",
        )
    years = sorted(ad_f["_year"].dropna().unique().astype(int).tolist(), reverse=True)
    months = list(range(1, 13))
    with p2:
        year_sel = st.selectbox(
            "年",
            options=years if years else [jst_now().year],
            key="dash_year_sel",
            disabled=period == "全期間",
        )
    with p3:
        month_sel = st.selectbox(
            "月",
            options=months,
            format_func=lambda m: f"{m}月",
            key="dash_month_sel",
            disabled=period != "月を指定",
        )
    with p4:
        supplier_filter = st.multiselect(
            "仕入先・取引先で絞り込み（未選択は全件）",
            options=sorted(ad_f[COL_SUPPLIER].fillna("").astype(str).unique().tolist()),
            key="dash_supplier_filter",
        )

    flt = ad_f
    if period == "年を指定":
        flt = flt[flt["_year"] == int(year_sel)]
    elif period == "月を指定":
        flt = flt[(flt["_year"] == int(year_sel)) & (flt["_month"] == int(month_sel))]
    if supplier_filter:
        flt = flt[flt[COL_SUPPLIER].astype(str).isin(supplier_filter)]

    if flt.empty:
        st.warning("条件に一致するデータがありません。")
        return

    q_in = int(flt["_qty_in"].sum())
    q_out = int(flt["_qty_out"].sum())
    q_net = q_in - q_out
    ex_in = int(round(flt["_amt_ex_in"].sum()))
    ex_out = int(round(flt["_amt_ex_out"].sum()))
    ex_net = ex_in - ex_out
    in_in = int(round(flt["_amt_in_in"].sum()))
    in_out = int(round(flt["_amt_in_out"].sum()))
    in_net = in_in - in_out

    m1, m2, m3 = st.columns(3)
    m1.metric("入庫 合計数量", f"{q_in:,}")
    m2.metric("出庫 合計数量", f"{q_out:,}")
    m3.metric("差し引き 数量（入−出）", f"{q_net:,}")
    m5, m6, m7, m8 = st.columns(4)
    m5.metric("入庫 合計金額（税抜）", f"¥{ex_in:,}")
    m6.metric("出庫 合計金額（税抜）", f"¥{ex_out:,}")
    m7.metric("差し引き 税抜（入−出）", f"¥{ex_net:,}")
    m8.metric("差し引き 税込（入−出）", f"¥{in_net:,}")

    st.markdown("##### 仕入先・取引先別サマリー（税抜金額・数量）")
    sup_col = "仕入先・取引先"
    g = flt.assign(**{sup_col: flt[COL_SUPPLIER].fillna("(未設定)").astype(str)})
    grp = (
        g.groupby(sup_col, dropna=False)
        .agg(
            入庫数量=("_qty_in", "sum"),
            出庫数量=("_qty_out", "sum"),
            入庫金額税抜=("_amt_ex_in", "sum"),
            出庫金額税抜=("_amt_ex_out", "sum"),
        )
        .reset_index()
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
        chart_qty = (
            alt.Chart(qty_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y("数量:Q", title="数量", axis=alt.Axis(format=",.0f")),
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
        chart_amt_bar = (
            alt.Chart(amt_bar_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y("金額:Q", title="金額（税抜・円）", axis=alt.Axis(format=",.0f")),
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
        "各月で入庫・出庫の2本の棒として並べたグラフです（積み上げ棒ではありません）。"
        "上の期間・仕入先・取引先の絞り込みに従います。"
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
        chart_cum_bar = (
            alt.Chart(cum_long)
            .mark_bar()
            .encode(
                x=alt.X("月:N", sort=month_order_c, axis=alt.Axis(title="月", labelAngle=-45)),
                y=alt.Y("累積金額:Q", title="累積金額（税抜・円）", axis=alt.Axis(format=",.0f")),
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
    st.caption("各仕入先・取引先で入庫・出庫の税抜金額を並べた棒グラフ（積み上げではありません）。")
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
        sup_chart = (
            alt.Chart(top_long)
            .mark_bar()
            .encode(
                x=alt.X(
                    f"{sup_col}:N",
                    sort=top_order,
                    axis=alt.Axis(title=sup_col, labelAngle=-45),
                ),
                y=alt.Y("金額:Q", title="金額（税抜・円）", axis=alt.Axis(format=",.0f")),
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


def render_inventory_manager() -> None:
    st.divider()
    st.subheader("在庫一覧マネージャー")
    st.caption(
        "スプレッドシートの全データを編集できます。行の追加・削除は表から操作し、"
        "表の直下の「台帳を更新する」でシートを上書き保存します。"
    )

    if msg := st.session_state.pop("_ledger_saved_flash", None):
        st.success(msg)

    if not _safe_secret("GOOGLE_SPREADSHEET_ID"):
        st.info("GOOGLE_SPREADSHEET_ID を設定すると、台帳の表示・編集ができます。")
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

    edited = st.data_editor(
        df_sorted,
        num_rows="dynamic",
        key=LEDGER_DATA_EDITOR_KEY,
        use_container_width=True,
        hide_index=True,
    )

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


def main():
    st.set_page_config(page_title="呉服 在庫・販売", layout="wide")
    st.title("呉服 在庫・販売管理")
    st.caption("写真アップロード・AI解析・Googleドライブ保存・スプレッドシート記録")

    # --- session 初期化（フォームは key で状態管理） ---
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
    if "field_memo" not in st.session_state:
        st.session_state.field_memo = ""
    if "field_unit_price_excl" not in st.session_state:
        st.session_state.field_unit_price_excl = 1
    st.session_state.pop("field_price_excl", None)

    # --- 一時緩和: secrets 一括チェックを無効化（No secrets found 回避・Gemini 動作確認優先） ---
    # try:
    #     _ = st.secrets["GEMINI_API_KEY"]
    #     _ = st.secrets["GAS_UPLOAD_URL"]
    #     _ = st.secrets["GAS_API_KEY"]
    #     _ = st.secrets["GOOGLE_DRIVE_FOLDER_ID"]
    #     _ = st.secrets["GOOGLE_SPREADSHEET_ID"]
    #     _load_service_account_info()
    # except Exception as e:
    #     st.error(f"設定エラー: {e}")
    #     st.info(
    #         "`.streamlit/secrets.toml` に GEMINI_API_KEY・GAS_UPLOAD_URL・GAS_API_KEY・"
    #         "GOOGLE_DRIVE_FOLDER_ID・GOOGLE_SPREADSHEET_ID・サービスアカウントを設定してください。"
    #     )
    #     return
    st.caption("開発モード: secrets の起動時チェックをスキップしています（AI解析の確認用）。")

    uploaded = st.file_uploader(
        "商品写真（カメラで撮影した画像をアップロード）",
        type=["jpg", "jpeg", "png", "webp"],
    )
    st.caption(
        "AI解析・ドライブ保存のいずれも、EXIF向き補正のうえ長辺最大1280px・JPEG品質80％へ変換してから行います（軽量化）。"
        "台帳の日時は EXIF の撮影日時が使えない場合は日本時間（JST）の現在時刻になります。"
    )

    movement = st.radio(
        "区分",
        ("入庫（購入）", "入庫（返品）", "出庫（販売）", "出庫（浮貸）"),
        horizontal=True,
    )

    col_a, col_b = st.columns(2)
    with col_a:
        analyze = st.button("AIで画像を解析", type="primary", disabled=uploaded is None)
    with col_b:
        if st.button("候補の自動入力をクリア"):
            st.session_state.field_product_name = ""
            st.session_state.field_supplier = ""
            st.session_state.field_qty = 1
            st.session_state.ai_kind = ""
            st.session_state.ai_features = ""
            st.session_state.field_memo = ""
            st.session_state.field_unit_price_excl = 1
            st.rerun()

    if analyze and uploaded is not None:
        with st.spinner("画像を解析しています…"):
            try:
                img = _gemini_input_image_from_upload(uploaded)
                raw_text = analyze_image_with_gemini(img)
                result = _parse_json_from_model(raw_text or "")
                st.session_state.field_qty = int(
                    result.get("quantity") or result.get("数量") or 1
                )
                st.session_state.ai_kind = str(
                    result.get("product_kind")
                    or result.get("商品名")
                    or result.get("種類")
                    or ""
                )
                vf = result.get("visual_features")
                if not vf:
                    parts = [
                        result.get(k)
                        for k in ("色", "柄", "素材", "状態", "色柄", "備考")
                        if result.get(k)
                    ]
                    vf = " / ".join(str(p) for p in parts)
                st.session_state.ai_features = str(vf or "")
                m = result.get("match")
                if isinstance(m, dict):
                    conf = float(m.get("confidence") or 0)
                    if conf >= 0.75:
                        st.session_state.field_product_name = str(m.get("product_name") or "").strip()
                        st.session_state.field_supplier = str(m.get("supplier") or "").strip()
                st.success("解析が完了しました。必要に応じて商品名・仕入先・取引先を修正してください。")
            except Exception as e:
                st.warning(
                    "現在混み合っているか、無料枠の上限に達している可能性があります。"
                    "1分ほど待ってから再試行してください。"
                )
                st.caption(f"詳細: {e}")

    if st.session_state.ai_kind or st.session_state.ai_features:
        st.subheader("AI解析結果（参考）")
        st.write(f"**推定種類:** {st.session_state.ai_kind or '—'}")
        st.write(f"**推定数量:** {int(st.session_state.field_qty)}")
        st.caption(f"マッチング用特徴: {st.session_state.ai_features or '—'}")

    st.subheader("必須入力")
    product_name = st.text_input("商品名（必須）", key="field_product_name")
    supplier = st.text_input("仕入先・取引先（必須）", key="field_supplier")
    quantity = st.number_input("数量", min_value=1, step=1, key="field_qty")

    unit_price_excl = st.number_input(
        "商品単価（税抜・必須）",
        min_value=1,
        step=1,
        key="field_unit_price_excl",
        help="1点あたりの税抜単価（円）。商品金額は数量×単価で自動計算します。",
    )

    _q = int(quantity)
    _u = int(unit_price_excl)
    _line_ex = _q * _u
    _line_in = price_incl_tax(_line_ex)

    price_row = st.columns([1, 1, 1])
    with price_row[0]:
        st.metric("商品金額（税抜・自動）", f"¥{_line_ex:,}")
        st.caption(f"数量 {_q} × 単価 ¥{_u:,}")
    with price_row[1]:
        st.metric("税込金額（行・自動）", f"¥{_line_in:,}")
        st.caption("標準税率10%を行合計に四捨五入")
    with price_row[2]:
        st.caption("スプレッドシートには単価・税抜行計・税込行計の3値を記録します。")

    st.subheader("任意入力")
    memo = st.text_area(
        "メモ（任意）",
        key="field_memo",
        height=100,
        placeholder="備考・社内メモなどがあれば入力してください",
    )

    confirm = st.button("確定（ドライブ保存・スプレッドシート記録）", type="primary", disabled=uploaded is None)

    if confirm:
        validation_ok = True
        if not (product_name or "").strip():
            st.error("商品名を入力してください。")
            validation_ok = False
        elif not (supplier or "").strip():
            st.error("仕入先・取引先を入力してください。")
            validation_ok = False
        elif int(unit_price_excl) < 1:
            st.error("商品単価（税抜）を1円以上で入力してください。")
            validation_ok = False
        elif uploaded is None:
            st.error("画像をアップロードしてください。")
            validation_ok = False

        if validation_ok:
            safe_base = re.sub(r"[^\w\-_.]", "_", uploaded.name.rsplit(".", 1)[0])[:80]
            raw_bytes = uploaded.getvalue()
            _record_dt = capture_datetime_jst_from_bytes(raw_bytes) or jst_now_str()

            try:
                with st.spinner("画像をリサイズ・圧縮しています…"):
                    data, mime = prepare_upload_image_jpeg(raw_bytes)
            except Exception as e:
                st.error(f"画像の処理に失敗しました: {e}")
            else:
                fname = f"{jst_now().strftime('%Y%m%d_%H%M%S')}_{safe_base}_{uuid.uuid4().hex[:8]}.jpg"

                with st.spinner("Googleドライブに保存しています…"):
                    try:
                        url = upload_image_to_drive(fname, mime, data)
                    except Exception as e:
                        st.error(f"ドライブ保存に失敗しました: {e}")
                    else:
                        _q2 = int(quantity)
                        _u2 = int(unit_price_excl)
                        _lex = _q2 * _u2
                        _lin = price_incl_tax(_lex)

                        with st.spinner("スプレッドシートに記録しています…"):
                            try:
                                append_sheet_row(
                                    movement,
                                    product_name.strip(),
                                    supplier.strip(),
                                    _q2,
                                    _u2,
                                    _lex,
                                    _lin,
                                    url,
                                    (memo or "").strip(),
                                    record_datetime=_record_dt,
                                )
                            except Exception as e:
                                st.error(f"スプレッドシート更新に失敗しました: {e}")
                                st.warning(f"画像は保存済みです: {url}")
                            else:
                                st.session_state.pop(LEDGER_DATA_EDITOR_KEY, None)
                                st.success("記録しました。")
                                st.markdown(f"[保存した画像を開く]({url})")
                                st.balloons()

    render_inventory_manager()


if __name__ == "__main__":
    if check_password():
        main()
