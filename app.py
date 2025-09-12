import io
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import json

import pandas as pd
import requests
from bs4 import BeautifulSoup
from PIL import Image as PILImage
from io import BytesIO

import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="発注書 自動生成（楽天画像専用）", page_icon="🧾", layout="wide")
st.title("🧾 発注書 自動生成｜画像は楽天のみから取得")

# ------- UI -------
col1, col2 = st.columns(2)
up_orders = col1.file_uploader("受注CSV（列: sku, 購入数）", type=["csv"])
up_master = col2.file_uploader("DB（Excel: sku, 原価, 商品URL, 商品名称, 特記事項）", type=["xlsx","xls"])

order_date = st.date_input("発注日", value=datetime.now().date())
img_max_h = st.number_input("画像の最大高さ(px)", min_value=60, max_value=240, value=120)
img_col_width = st.number_input("B列の幅(文字数)", min_value=10, max_value=40, value=18)

st.markdown("---")
left, right = st.columns([2,1])
rakuten_shop = left.text_input("楽天のショップID（例: lilirena）", value="lilirena")
max_workers = right.slider("並列取得数", 2, 12, 6)

# ------- HTTP client -------
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.9",
}

session = requests.Session()
session.headers.update(DEFAULT_HEADERS)

# ------- helpers -------

def normalize_sku(s):
    if pd.isna(s):
        return None
    return str(s).strip().replace("　"," ").lower()

def base_code_from_sku(sku: str) -> str:
    if not sku:
        return ""
    code = str(sku).split("-")[0].strip()
    return code

def build_rakuten_url(sku: str, shop: str) -> str:
    code = base_code_from_sku(sku)
    if not code or not shop:
        return ""
    return f"https://item.rakuten.co.jp/{shop}/{code}/"

def pick_rakuten_image(html: str) -> str | None:
    if not html:
        return None
    soup = BeautifulSoup(html, "html.parser")

    for attrs in (
        {"property": "og:image"},
        {"name": "og:image"},
        {"property": "twitter:image"},
        {"name": "twitter:image"},
        {"property": "og:image:url"},
    ):
        tag = soup.find("meta", attrs=attrs)
        if tag and tag.get("content"):
            u = tag["content"].strip()
            if u.startswith("//"): u = "https:" + u
            return u

    for s in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            data = json.loads(s.string or "{}")
        except Exception:
            continue
        def extract_img(obj):
            if isinstance(obj, dict):
                img = obj.get("image")
                if isinstance(img, str):
                    return img
                if isinstance(img, list) and img:
                    return img[0]
            return None
        if isinstance(data, dict) and data.get("@type") == "Product":
            u = extract_img(data)
            if u:
                return u
        if isinstance(data, list):
            for node in data:
                if isinstance(node, dict) and node.get("@type") == "Product":
                    u = extract_img(node)
                    if u:
                        return u

    for sel in [
        "#rakutenLimitedId_ImageMain img",
        "#productMainImage img",
        "#page-body img",
        "img",
    ]:
        el = soup.select_one(sel)
        if el:
            src = el.get("src") or el.get("data-src") or ""
            if src:
                if src.startswith("//"): src = "https:" + src
                return src

    return None

def download_image(image_url: str, referer: str | None = None) -> BytesIO | None:
    if not image_url:
        return None
    headers = DEFAULT_HEADERS.copy()
    if referer:
        headers["Referer"] = referer
    try:
        resp = session.get(image_url, headers=headers, stream=True, timeout=5)
        resp.raise_for_status()
        ctype = (resp.headers.get("Content-Type") or "").lower()
        if not (ctype.startswith("image/") or any(x in ctype for x in ["webp","jpeg","png","jpg"])):
            return None
        data = BytesIO(resp.content)
        data.seek(0)
        return data
    except Exception:
        return None

def resize_keep_ratio(bin_io: BytesIO, max_h: int) -> BytesIO | None:
    try:
        img = PILImage.open(bin_io)
        if img.mode not in ("RGB","RGBA"):
            img = img.convert("RGB")
        w, h = img.size
        if h > max_h:
            scale = max_h / h
            img = img.resize((int(w*scale), int(h*scale)))
        out = BytesIO()
        img.save(out, format="PNG", optimize=True)
        out.seek(0)
        return out
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

if up_orders and up_master:
    try:
        orders = read_orders(up_orders)
        master = pd.read_excel(up_master)
        need = {"sku","原価","商品URL","商品名称","特記事項"}
        if not need.issubset(master.columns):
            st.error(f"DBの列名不足: 必要 {need} / 実際 {set(master.columns)}")
            st.stop()

        master["sku"] = master["sku"].apply(normalize_sku)
        orders_sum = orders.groupby("sku",as_index=False)["購入数"].sum().query("購入数>0")
        merged = orders_sum.merge(master, on="sku", how="left")

        st.info("画像URLを楽天から取得中…（見つからなければ『画像なし』）")

        decided_urls = [None] * len(merged)
        to_fetch = []

        for i, r in merged.iterrows():
            rak_url = build_rakuten_url(r.get("sku"), rakuten_shop)
            if rak_url:
                to_fetch.append((i, rak_url, r.get("sku")))

        def fetch_one(idx: int, url: str, sku: str):
            try:
                resp = session.get(url, timeout=5)
                resp.raise_for_status()
                html = resp.text
                return idx, (pick_rakuten_image(html) or None), url
            except Exception:
                return idx, None, url

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = [ex.submit(fetch_one, i, u, s) for (i,u,s) in to_fetch]
            for fut in as_completed(futures):
                i, img_u, ref = fut.result()
                decided_urls[i] = img_u

        merged["画像URL(決定)"] = [u if u else "画像なし" for u in decided_urls]

        st.subheader("プレビュー")
        st.dataframe(merged[["sku","購入数","商品名称","商品URL","画像URL(決定)"]], use_container_width=True, height=320)

        st.info("Excelを生成中…（画像を埋め込み）")
        wb = Workbook()
        ws = wb.active
        ws.title = "発注書"

        headers = ["", "写真", "sku", "購入数", "単価",
                   "特記事項", "商品名称", "商品URL", "変更後URL",
                   "サイズ", "色", "中国内送料",
                   "単価", "合計", "発注日",
                   ""]
        ws.append(headers)
        ws.column_dimensions["B"].width = float(img_col_width)

        failed_rows = []
        embed_ok = 0
        fallback_image_formula = 0
        embed_fail = 0

        d = order_date
        date_str = f"{d.year}/{d.month}/{d.day}"

        for _, row in merged.iterrows():
            genka = pd.to_numeric(row.get("原価"), errors="coerce")
            qty = int(row.get("購入数") or 0)
            gokei = (genka * qty) if pd.notna(genka) else None

            excel_row = ["", "", row.get("sku"), qty, "",
                         row.get("特記事項"), row.get("商品名称"),
                         row.get("商品URL"), "", "", "", "",
                         genka, gokei, date_str, ""]
            ws.append(excel_row)

            r_i = ws.max_row
            img_url = row.get("画像URL(決定)")
            referer = build_rakuten_url(row.get("sku"), rakuten_shop)

            bin_data = download_image(img_url, referer=referer) if (img_url and img_url != "画像なし") else None
            if bin_data:
                bin_resized = resize_keep_ratio(bin_data, max_h=int(img_max_h))
                if bin_resized:
                    try:
                        xlimg = XLImage(bin_resized)
                        ws.add_image(xlimg, f"B{r_i}")
                        ws.row_dimensions[r_i].height = int(int(img_max_h) / 1.33)
                        embed_ok += 1
                        continue
                    except Exception:
                        pass

            if img_url and img_url != "画像なし":
                try:
                    ws.cell(row=r_i, column=2).value = f'=IMAGE("{img_url}")'
                    ws.row_dimensions[r_i].height = int(int(img_max_h) / 1.33)
                    fallback_image_formula += 1
                except Exception:
                    embed_fail += 1
                    failed_rows.append({"sku": row.get("sku"), "商品URL": row.get("商品URL"), "画像URL": img_url, "理由": "IMAGE関数保険も失敗"})
            else:
                embed_fail += 1
                failed_rows.append({"sku": row.get("sku"), "商品URL": row.get("商品URL"), "画像URL": img_url, "理由": "URLなし"})

        if failed_rows:
            ws2 = wb.create_sheet("画像取得失敗")
            ws2.append(["sku","商品URL","画像URL","理由"])
            for fr in failed_rows:
                ws2.append([fr.get("sku"), fr.get("商品URL"), fr.get("画像URL"), fr.get("理由")])

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        filename = f"発注書_{datetime.now().strftime('%Y%m%d')}_画像埋め込み.xlsx"
        st.download_button("📥 発注書（画像埋め込み）をダウンロード",
                           data=bio.getvalue(),
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.success(f"画像埋め込み 成功: {embed_ok} 件 / 保険(IMAGE関数): {fallback_image_formula} 件 / 失敗: {embed_fail} 件")

    except Exception as e:
        st.error(f"エラー: {e}")
else:
    st.info("受注CSVとDBを選ぶと処理できます。")