import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image as PILImage
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
# from concurrent.futures import ThreadPoolExecutor, as_completed # 並列処理を無効化

# ====== 固定設定 ======
SHOP_ID = "lilirena"     # 楽天ショップID
MAX_WORKERS = 12         # 並列取得 固定（スライダー廃止） -> 同期処理になったため、この変数は未使用
IMG_MAX_H = 120          # 画像高さ(px)固定
IMG_COL_WIDTH = 18       # B列の幅(文字数)固定
# =====================

st.set_page_config(page_title="Order Maker", page_icon="🧾", layout="centered")

# --- NE風スタイル（青ボタン＆青タイトル） ---
st.markdown("""
<style>
:root { --ne-blue:#2a6df4; }
.block-container { max-width: 880px; }
.titlebar { font-size:22px; font-weight:800; display:flex; gap:10px; align-items:center; color:var(--ne-blue); }
.titlebar:before{content:"📄";}
.subtle{color:#667085; font-size:13px; margin-bottom:8px;}
.card{border:1px solid #e6e9ef; border-radius:14px; padding:22px; margin:14px 0; background:#fff; box-shadow:0 2px 8px rgba(16,24,40,.04);}
.drop{border:2px dashed #d5d9e3; border-radius:12px; padding:28px; text-align:center; color:#6b7280; background:#fafbff;}
/* Streamlitのボタン強制ブルー化 */
.stButton > button {
  width: 100%; height: 52px; font-weight: 700; border-radius: 10px; font-size: 16px;
  background: var(--ne-blue) !important; color: #fff !important; border: none !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="titlebar">Order Maker</div><div class="subtle">発注書自動作成</div>', unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown('<div class="drop">受注データ</div>', unsafe_allow_html=True)
        up_orders = st.file_uploader("ファイルを選択", type=["csv"], label_visibility="collapsed", key="orders")
    with c2:
        st.markdown('<div class="drop">商品マスタ</div>', unsafe_allow_html=True)
        up_master = st.file_uploader("ファイルを選択", type=["xlsx","xls"], label_visibility="collapsed", key="master")
    st.markdown('</div>', unsafe_allow_html=True)

# ------- HTTP client -------
DEFAULT_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36"),
    "Accept-Language": "ja,en;q=0.9",
}
session = requests.Session()
session.headers.update(DEFAULT_HEADERS)

# ------- helpers -------
def normalize_sku(s):
    if pd.isna(s): return None
    return str(s).strip().replace("　"," ").lower()

def base_code_from_sku(sku: str) -> str:
    if not sku: return ""
    return str(sku).split("-")[0].strip()

def build_rakuten_url(sku: str, shop: str = SHOP_ID) -> str:
    code = base_code_from_sku(sku)
    if not code or not shop: return ""
    return f"https://item.rakuten.co.jp/{shop}/{code}/"

def pick_rakuten_image(html: str) -> str | None:
    from bs4 import BeautifulSoup
    if not html: return None
    soup = BeautifulSoup(html, "html.parser")
    for attrs in ({"property":"og:image"},{"name":"og:image"},
                  {"property":"twitter:image"},{"name":"twitter:image"},
                  {"property":"og:image:url"}):
        tag = soup.find("meta", attrs=attrs)
        if tag and tag.get("content"):
            u = tag["content"].strip()
            if u.startswith("//"): u = "https:"+u
            return u
    for sel in ["#rakutenLimitedId_ImageMain img", "#productMainImage img", "#page-body img", "img"]:
        el = soup.select_one(sel)
        if el:
            src = el.get("src") or el.get("data-src") or ""
            if src:
                if src.startswith("//"): src = "https:"+src
                return src
    return None

def download_image(image_url: str, referer: str | None = None) -> BytesIO | None:
    if not image_url: return None
    headers = DEFAULT_HEADERS.copy()
    if referer: headers["Referer"] = referer
    try:
        resp = session.get(image_url, headers=headers, stream=True, timeout=12)
        resp.raise_for_status()
        img = PILImage.open(BytesIO(resp.content))
        img.thumbnail((IMG_COL_WIDTH*5, IMG_MAX_H))
        bio = BytesIO()
        img.save(bio, format="PNG")
        bio.seek(0)
        return bio
    except Exception:
        return None

def read_orders(file) -> pd.DataFrame:
    for enc in ["cp932","utf-8-sig","utf-8"]:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding=enc)
            if {"sku","購入数"}.issubset(df.columns):
                df["sku"] = df["sku"].apply(normalize_sku)
                df["購入数"] = pd.to_numeric(df["購入数"], errors="coerce").fillna(0).astype(int)
                return df
        except Exception:
            pass
    raise ValueError("受注CSVの列名は 'sku, 購入数' を想定しています。")

# --- 画像URL取得関数（同期処理で使用） ---
def fetch_image(idx: int, sku: str):
    rak_url = build_rakuten_url(sku)
    if not rak_url: return idx, "URLなし"
    try:
        resp = session.get(rak_url, timeout=12)
        resp.raise_for_status()
        img_url = pick_rakuten_image(resp.text)
        return idx, img_url if img_url else "画像なし"
    except Exception:
        return idx, "取得失敗"

# ====== メイン ======
go = st.button("発注書作成")

if go:
    if not (up_orders and up_master):
        st.error("受注データ と 商品マスタ を選択してください。")
        st.stop()

    try:
        orders = read_orders(up_orders)
        master = pd.read_excel(up_master)
        need = {"sku","原価","商品URL","商品名称","特記事項"}
        if not need.issubset(master.columns):
            st.error(f"DBの列名不足: 必要 {need} / 実際 {set(master.columns)}")
            st.stop()

        master["sku"] = master["sku"].apply(normalize_sku)
        orders_sum = orders.groupby("sku", as_index=False)["購入数"].sum().query("購入数>0")
        merged = orders_sum.merge(master, on="sku", how="left")

        # 進捗バー：画像URL取得
        prog = st.progress(0, text="画像URL取得中…")
        decided = [None]*len(merged)

        # ★★★ 修正箇所：ThreadPoolExecutor を削除し、同期ループに置き換え ★★★
        total = len(merged)
        for i, (_, row) in enumerate(merged.iterrows()):
            idx, url = fetch_image(i, row.get("sku"))
            decided[idx] = url
            prog.progress(int((i + 1) * 100 / total), text=f"画像URL取得 {i + 1}/{total}")
        # ★★★ 修正箇所終了 ★★★

        merged["画像URL(決定)"] = decided
        prog.progress(100, text="画像URL取得 完了")

        # Excel出力（画像埋め込み）
        wb = Workbook(); ws = wb.active; ws.title = "発注書"
        headers = ["", "写真", "sku", "購入数", "単価",
                   "特記事項", "商品名称", "商品URL", "変更後URL",
                   "サイズ", "色", "中国内送料",
                   "単価", "合計", "発注日", ""]
        ws.append(headers)
        ws.column_dimensions["B"].width = float(IMG_COL_WIDTH)

        d = datetime.now(); date_str = f"{d.year}/{d.month}/{d.day}"
        ok = fail = 0

        prog2 = st.progress(0, text="Excelに画像を埋め込み中…")
        total_rows = len(merged)

        # ★ ここを iterrows に変更（列名そのまま使える）
        for i, (_, row) in enumerate(merged.iterrows(), start=1):
            genka = pd.to_numeric(row.get("原価"), errors="coerce")
            qty = int(row.get("購入数") or 0)
            gokei = (genka * qty) if pd.notna(genka) else None

            excel_row = ["", "", row.get("sku"), qty, genka,
                         row.get("特記事項"), row.get("商品名称"),
                         row.get("商品URL"), "", "", "", "",
                         genka, gokei, date_str, ""]
            ws.append(excel_row)

            r_i = ws.max_row
            img_url = row.get("画像URL(決定)")  # ← そのまま参照OK
            referer = build_rakuten_url(row.get("sku"))

            bin_io = download_image(img_url, referer=referer) if (img_url and "http" in str(img_url)) else None
            if bin_io:
                try:
                    xlimg = XLImage(bin_io)
                    xlimg.anchor = f"B{r_i}"
                    ws.add_image(xlimg)
                    ws.row_dimensions[r_i].height = int(IMG_MAX_H * 0.75)
                    ok += 1
                except Exception:
                    fail += 1
            else:
                fail += 1

            prog2.progress(int(i*100/total_rows), text=f"Excel埋め込み {i}/{total_rows}")

        bio = BytesIO(); wb.save(bio); bio.seek(0)
        filename = f"発注書_{datetime.now().strftime('%Y%m%d')}_画像埋め込み.xlsx"
        st.download_button("📥 発注書（画像埋め込み）をダウンロード",
                           data=bio.getvalue(),
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success(f"画像埋め込み 成功: {ok} 件 / 失敗: {fail} 件")

    except Exception as e:
        st.error(f"エラー: {e}")
