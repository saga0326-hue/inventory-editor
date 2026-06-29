# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

st.set_page_config(
    page_title="盤點資料編輯器",
    page_icon="📋",
    layout="wide",
)

# ── 全域樣式 ──────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background: #2e1e1e; }
[data-testid="stSidebar"] h2 { color: #5e9bfa; }
.dup-card  { background:#5a1a1a; border-radius:6px; padding:8px 12px;
             display:inline-block; margin:4px; }
.dup-id    { color:#ff5555; font-weight:bold; font-size:13px; }
.dup-dates { color:#f39c12; font-size:12px; }
.cut-card  { background:#3d2020; border-radius:6px; padding:8px 10px; margin:4px 0; }
.cut-id    { color:#5e9bfa; font-weight:bold; }
.cut-tag   { color:#f39c12; font-size:12px; }
.cut-note  { color:#6c7086; font-size:11px; }
</style>
""", unsafe_allow_html=True)

FIELD_MAP = {
    "午別(1、2)":     "午別",
    "日期(yyyymmdd)": "日期",
    "店號(6碼)":      "店號",
    "店名":           "店名",
    "型態(4碼)":      "型態",
    "預定盤點者":      "預定盤點者",
    "備註":           "備註",
    "課別":           "課別",
}

COL_CONFIG = {
    "午別": st.column_config.SelectboxColumn("午別", options=["上午", "下午"], width="small"),
    "日期": st.column_config.TextColumn("日期", width="small"),
    "店號": st.column_config.TextColumn("店號", width="small"),
    "店名": st.column_config.TextColumn("店名", width="medium"),
    "型態": st.column_config.SelectboxColumn("型態",
            options=["閉", "轉", "解", "續", "FC", "RC"], width="small"),
    "預定盤點者": st.column_config.TextColumn("預定盤點者", width="small"),
    "備註": st.column_config.TextColumn("備註", width="medium"),
    "課別": st.column_config.TextColumn("課別", width="small"),
}


def fmt_date(v):
    try:
        dt = pd.to_datetime(str(v).strip())
        return f"{dt.month}/{dt.day}"
    except Exception:
        return str(v).strip()


def init_state():
    if "tab_data"  not in st.session_state: st.session_state.tab_data  = {}
    if "cut_list"  not in st.session_state: st.session_state.cut_list  = []
    if "date_order" not in st.session_state: st.session_state.date_order = []


# ── 匯入盤點表 ────────────────────────────────────────
def load_inventory(file):
    raw = pd.read_excel(file, dtype=str).fillna("")
    available = {k: v for k, v in FIELD_MAP.items() if k in raw.columns}
    df = raw[list(available.keys())].rename(columns=available).copy()

    if "午別" in df.columns:
        df["午別"] = df["午別"].map(
            lambda v: "上午" if str(v).strip() == "1"
            else ("下午" if str(v).strip() == "2" else v))
    if "日期" in df.columns:
        df["日期"] = df["日期"].map(fmt_date)

    tab_data   = {}
    date_order = []
    if "日期" in df.columns:
        df["_sort"] = pd.to_datetime(
            df["日期"].map(lambda v: f"2026/{v}"), errors="coerce")
        ordered = (df.drop_duplicates("日期")
                     .sort_values("_sort")["日期"].tolist())
        df = df.drop(columns=["_sort"])
        for d in ordered:
            tab_data[d] = df[df["日期"] == d].reset_index(drop=True)
            date_order.append(d)
    else:
        tab_data["全部"] = df
        date_order = ["全部"]

    return tab_data, date_order


# ── 匯入異動名單 ──────────────────────────────────────
def load_change_list(file, tree_cols):
    df = pd.read_excel(file, dtype=str).fillna("")
    col_map = {"日期": "日期", "午別": "午別", "店號": "店號",
               "店名": "店名", "型態": "型態", "備註": "備註"}
    entries = []
    for _, row in df.iterrows():
        store_id   = str(row.get("店號", "")).strip()
        store_name = str(row.get("店名", "")).strip()
        if not store_id and not store_name:
            continue
        vals = {col: "" for col in tree_cols}
        for src, dst in col_map.items():
            if src in df.columns and dst in tree_cols:
                v = str(row.get(src, "")).strip()
                if src == "午別":
                    v = "上午" if v == "1" else ("下午" if v == "2" else v)
                elif src == "日期":
                    v = fmt_date(v)
                vals[dst] = v
        entries.append({"store_id": store_id, "store_name": store_name,
                        "vals": vals})
    return entries


# ── 重複店號偵測 ──────────────────────────────────────
def find_duplicates(tab_data):
    store_map = defaultdict(list)
    for date, df in tab_data.items():
        if "店號" not in df.columns:
            continue
        for _, row in df.iterrows():
            sid   = str(row.get("店號", "")).strip()
            sname = str(row.get("店名", "")).strip()
            if sid:
                store_map[sid].append((date, sname))
    return {sid: ents for sid, ents in store_map.items() if len(ents) > 1}


# ── 下載 Excel ────────────────────────────────────────
def to_excel_bytes(tab_data, date_order):
    frames = [tab_data[d] for d in date_order if d in tab_data and len(tab_data[d])]
    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ════════════════════════════════════════════════════
def main():
    init_state()
    st.title("📋 盤點資料編輯器")

    # ── 頂部工具列 ────────────────────────────────────
    col1, col2, col3 = st.columns([2, 2, 6])
    with col1:
        inv_file = st.file_uploader("📂 匯入盤點表", type=["xlsx", "xls"],
                                    key="inv_upload", label_visibility="collapsed")
        if inv_file:
            tab_data, date_order = load_inventory(inv_file)
            st.session_state.tab_data   = tab_data
            st.session_state.date_order = date_order
            st.success(f"已載入 {len(date_order)} 個日期分頁")

    with col2:
        change_file = st.file_uploader("📋 匯入異動名單", type=["xlsx", "xls"],
                                       key="change_upload", label_visibility="collapsed")
        if change_file and st.session_state.tab_data:
            tree_cols = list(list(st.session_state.tab_data.values())[0].columns)
            entries   = load_change_list(change_file, tree_cols)
            st.session_state.cut_list.extend(entries)
            st.success(f"異動名單已載入 {len(entries)} 筆至左側剪貼區")

    with col3:
        if st.session_state.tab_data:
            excel_bytes = to_excel_bytes(
                st.session_state.tab_data, st.session_state.date_order)
            st.download_button("💾 另存新檔", data=excel_bytes,
                               file_name="盤點表_已編輯.xlsx",
                               mime="application/vnd.openxmlformats-officedocument"
                                    ".spreadsheetml.sheet")

    st.divider()

    if not st.session_state.tab_data:
        st.info("請先匯入盤點表 Excel 檔案。")
        return

    # ── 重複店號警示 ──────────────────────────────────
    duplicates = find_duplicates(st.session_state.tab_data)
    if duplicates:
        st.error(f"⚠ 重複店號：共 {len(duplicates)} 間店")
        cards_html = ""
        for sid, entries in sorted(duplicates.items()):
            sname = entries[0][1]
            dates = "、".join(e[0] for e in entries)
            cards_html += (f'<div class="dup-card">'
                           f'<div class="dup-id">{sid} {sname}</div>'
                           f'<div class="dup-dates">出現於：{dates}</div>'
                           f'</div>')
        st.markdown(cards_html, unsafe_allow_html=True)
        st.divider()

    # ── 搜尋列 ────────────────────────────────────────
    search_kw = st.text_input("🔍 搜尋（跨所有分頁）", placeholder="輸入店號、店名、日期…")

    # ── 左側剪貼區（sidebar）─────────────────────────
    with st.sidebar:
        st.markdown("## ✂ 剪下記錄")
        if not st.session_state.cut_list:
            st.caption("目前沒有剪下的資料")
        else:
            # 貼回表單
            date_order = st.session_state.date_order
            tree_cols  = list(list(st.session_state.tab_data.values())[0].columns)

            for i, entry in enumerate(st.session_state.cut_list):
                sid   = entry["store_id"]
                sname = entry["store_name"]
                vals  = entry["vals"]
                tag   = "  ".join(filter(None, [
                    vals.get("日期",""), vals.get("午別",""), vals.get("型態","")]))
                note  = vals.get("備註", "")

                with st.expander(f"**{sid}** {sname}"):
                    if tag:  st.markdown(f'<span class="cut-tag">{tag}</span>',
                                         unsafe_allow_html=True)
                    if note: st.caption(f"備註：{note}")

                    st.markdown("**編輯後插入：**")
                    am_val   = st.selectbox("午別", ["上午", "下午"],
                                            index=0 if vals.get("午別") != "下午" else 1,
                                            key=f"am_{i}")
                    date_val = st.text_input("日期", value=vals.get("日期",""),
                                             key=f"date_{i}")
                    type_val = st.selectbox("型態", ["", "閉", "轉", "解", "續", "FC", "RC"],
                                            index=["","閉","轉","解","續","FC","RC"].index(
                                                vals.get("型態","")) if vals.get("型態","")
                                                in ["","閉","轉","解","續","FC","RC"] else 0,
                                            key=f"type_{i}")

                    target_date = st.selectbox("插入到分頁",
                                               date_order, key=f"tdate_{i}")
                    insert_pos  = st.selectbox(
                        "插入位置",
                        ["插入在最前面"] + [
                            f"第 {j+1} 列之後（{row.get('店號','')} {row.get('店名','')}）"
                            for j, row in st.session_state.tab_data.get(
                                target_date, pd.DataFrame()).iterrows()],
                        key=f"pos_{i}")

                    if st.button("✅ 插入", key=f"insert_{i}"):
                        new_row = dict(vals)
                        new_row["午別"] = am_val
                        new_row["日期"] = date_val
                        new_row["型態"] = type_val

                        df = st.session_state.tab_data.get(target_date,
                                                            pd.DataFrame(columns=tree_cols))
                        new_df_row = pd.DataFrame([new_row])
                        if insert_pos == "插入在最前面":
                            df = pd.concat([new_df_row, df], ignore_index=True)
                        else:
                            pos = int(insert_pos.split("第 ")[1].split(" 列")[0])
                            top = df.iloc[:pos]
                            bot = df.iloc[pos:]
                            df  = pd.concat([top, new_df_row, bot], ignore_index=True)

                        st.session_state.tab_data[target_date] = df
                        st.session_state.cut_list.pop(i)
                        st.rerun()

                    if st.button("🗑 移除", key=f"del_cut_{i}"):
                        st.session_state.cut_list.pop(i)
                        st.rerun()

    # ── 分頁資料表 ────────────────────────────────────
    date_order = st.session_state.date_order
    tabs = st.tabs([f" {d} " for d in date_order])

    for tab, date in zip(tabs, date_order):
        with tab:
            df = st.session_state.tab_data.get(date, pd.DataFrame())

            # 套用搜尋高亮（只顯示，不過濾）
            if search_kw:
                mask = df.apply(
                    lambda row: search_kw.lower() in " ".join(
                        str(v) for v in row).lower(), axis=1)
                match_count = mask.sum()
                if match_count:
                    st.info(f"🔍 此分頁找到 {match_count} 筆符合「{search_kw}」的資料（已標示）")
                else:
                    st.caption(f"此分頁無符合「{search_kw}」的資料")

            col_cfg = {k: v for k, v in COL_CONFIG.items() if k in df.columns}

            edited = st.data_editor(
                df,
                key=f"editor_{date}",
                column_config=col_cfg,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
            )
            st.session_state.tab_data[date] = edited

            # 剪下選取功能（透過勾選行）
            st.caption("💡 在表格最左側打勾選取列，再點「✂ 剪下選取列」")
            if st.button("✂ 剪下選取列", key=f"cut_{date}"):
                st.info("請在表格中選取列後，使用表格右上角的刪除功能，"
                        "或直接在表格中刪除列後點「加入剪貼區」")


if __name__ == "__main__":
    main()
