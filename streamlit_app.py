# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

st.set_page_config(page_title="盤點資料編輯器", page_icon="📋", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background:#2e1e1e; }
[data-testid="stSidebar"] * { color:#cdd6f4; }
.dup-bar { background:#3a1010; border-radius:8px; padding:10px 16px; margin-bottom:8px; }
.dup-chip { background:#5a1a1a; border-radius:6px; padding:6px 12px;
            display:inline-block; margin:3px; }
.dup-id   { color:#ff5555; font-weight:bold; font-size:13px; }
.dup-date { color:#f39c12; font-size:12px; }
.cut-chip { background:#3d2020; border-radius:6px; padding:8px 10px; margin:4px 0; }
.cut-id   { color:#5e9bfa; font-weight:bold; font-size:13px; }
.cut-tag  { color:#f39c12; font-size:12px; }
div[data-testid="stDataEditor"] { border-radius:8px; }
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
COLS = list(FIELD_MAP.values())
TYPE_OPTS = ["", "閉", "轉", "解", "續", "FC", "RC"]
AM_OPTS   = ["", "上午", "下午"]


def fmt_date(v):
    try:
        dt = pd.to_datetime(str(v).strip())
        return f"{dt.month}/{dt.day}"
    except Exception:
        return str(v).strip()


def init():
    for k, d in [("tab_data", {}), ("date_order", []), ("cut_list", [])]:
        if k not in st.session_state:
            st.session_state[k] = d


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
    tab_data, date_order = {}, []
    if "日期" in df.columns:
        df["_s"] = pd.to_datetime(df["日期"].map(
            lambda v: f"2026/{v}"), errors="coerce")
        ordered = df.drop_duplicates("日期").sort_values("_s")["日期"].tolist()
        df = df.drop(columns=["_s"])
        for d in ordered:
            tab_data[d] = df[df["日期"] == d].reset_index(drop=True)
            date_order.append(d)
    else:
        tab_data["全部"] = df
        date_order = ["全部"]
    return tab_data, date_order


def load_change(file, cols):
    df = pd.read_excel(file, dtype=str).fillna("")
    col_map = {"日期":"日期","午別":"午別","店號":"店號",
               "店名":"店名","型態":"型態","備註":"備註"}
    entries = []
    for _, row in df.iterrows():
        sid = str(row.get("店號","")).strip()
        sname = str(row.get("店名","")).strip()
        if not sid and not sname: continue
        vals = {c: "" for c in cols}
        for src, dst in col_map.items():
            if src in df.columns and dst in cols:
                v = str(row.get(src,"")).strip()
                if src == "午別":
                    v = "上午" if v=="1" else ("下午" if v=="2" else v)
                elif src == "日期":
                    v = fmt_date(v)
                vals[dst] = v
        entries.append({"sid": sid, "sname": sname, "vals": vals})
    return entries


def find_dups(tab_data):
    store_map = defaultdict(list)
    for date, df in tab_data.items():
        if "店號" not in df.columns: continue
        for _, row in df.iterrows():
            sid   = str(row.get("店號","")).strip()
            sname = str(row.get("店名","")).strip()
            if sid:
                store_map[sid].append((date, sname))
    return {s: e for s, e in store_map.items() if len(e) > 1}


def to_excel(tab_data, date_order):
    frames = [tab_data[d] for d in date_order
              if d in tab_data and len(tab_data[d]) > 0]
    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def col_config(cols):
    cfg = {}
    for c in cols:
        if c == "午別":
            cfg[c] = st.column_config.SelectboxColumn(c, options=AM_OPTS, width="small")
        elif c == "型態":
            cfg[c] = st.column_config.SelectboxColumn(c, options=TYPE_OPTS, width="small")
        elif c in ("日期","店號","課別"):
            cfg[c] = st.column_config.TextColumn(c, width="small")
        elif c in ("店名","預定盤點者"):
            cfg[c] = st.column_config.TextColumn(c, width="medium")
        else:
            cfg[c] = st.column_config.TextColumn(c, width="medium")
    cfg["✂"] = st.column_config.CheckboxColumn("選取", width="small", default=False)
    return cfg


# ════════════════════════════════════════════════════
def main():
    init()
    st.title("📋 盤點資料編輯器")

    # ── 工具列 ────────────────────────────────────────
    c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
    with c1:
        inv_file = st.file_uploader("📂 匯入盤點表", type=["xlsx","xls"],
                                    label_visibility="collapsed")
        if inv_file:
            td, do = load_inventory(inv_file)
            st.session_state.tab_data   = td
            st.session_state.date_order = do
            st.success(f"載入 {len(do)} 個日期分頁")

    with c2:
        chg_file = st.file_uploader("📋 匯入異動名單", type=["xlsx","xls"],
                                    label_visibility="collapsed")
        if chg_file and st.session_state.tab_data:
            cols = list(list(st.session_state.tab_data.values())[0].columns)
            entries = load_change(chg_file, cols)
            st.session_state.cut_list.extend(entries)
            st.success(f"異動名單載入 {len(entries)} 筆")

    with c3:
        if st.session_state.tab_data:
            st.download_button("💾 另存新檔",
                data=to_excel(st.session_state.tab_data, st.session_state.date_order),
                file_name="盤點表_已編輯.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if not st.session_state.tab_data:
        st.info("請先匯入盤點表 Excel 檔案。")
        return

    st.divider()

    # ── 重複店號警示 ──────────────────────────────────
    dups = find_dups(st.session_state.tab_data)
    if dups:
        chips = "".join(
            f'<div class="dup-chip"><div class="dup-id">{sid} {ents[0][1]}</div>'
            f'<div class="dup-date">出現於：{"、".join(e[0] for e in ents)}</div></div>'
            for sid, ents in sorted(dups.items()))
        st.markdown(
            f'<div class="dup-bar">⚠️ <b style="color:#ff5555">重複店號：共 {len(dups)} 間</b><br>{chips}</div>',
            unsafe_allow_html=True)

    # ── 搜尋列 ────────────────────────────────────────
    search_kw = st.text_input("🔍 搜尋（跨所有分頁）", placeholder="輸入店號、店名、日期…")
    if search_kw:
        results = []
        for date, df in st.session_state.tab_data.items():
            mask = df.apply(
                lambda r: search_kw.lower() in " ".join(str(v) for v in r).lower(), axis=1)
            cnt = mask.sum()
            if cnt:
                results.append(f"**{date}**：{cnt} 筆")
        if results:
            st.info("🔍 找到  " + "　｜　".join(results))
        else:
            st.warning(f"找不到「{search_kw}」")

    st.divider()

    # ── 左側剪貼區 ────────────────────────────────────
    with st.sidebar:
        st.markdown("## ✂ 剪下記錄")
        if not st.session_state.cut_list:
            st.caption("目前沒有剪下的資料")
        else:
            date_order = st.session_state.date_order
            cols = list(list(st.session_state.tab_data.values())[0].columns)
            to_remove = []

            for i, entry in enumerate(st.session_state.cut_list):
                sid, sname = entry["sid"], entry["sname"]
                vals = entry["vals"]
                tag  = "  ".join(filter(None, [vals.get("日期",""),
                                               vals.get("午別",""),
                                               vals.get("型態","")]))
                with st.expander(f"**{sid}** {sname}" + (f" ｜ {tag}" if tag else "")):
                    am_v = st.selectbox("午別", AM_OPTS,
                        index=AM_OPTS.index(vals.get("午別","")) if vals.get("午別","") in AM_OPTS else 0,
                        key=f"am_{i}")
                    dt_v = st.text_input("日期", value=vals.get("日期",""), key=f"dt_{i}")
                    ty_v = st.selectbox("型態", TYPE_OPTS,
                        index=TYPE_OPTS.index(vals.get("型態","")) if vals.get("型態","") in TYPE_OPTS else 0,
                        key=f"ty_{i}")
                    tdate = st.selectbox("插入到分頁", date_order, key=f"td_{i}")
                    tdf = st.session_state.tab_data.get(tdate, pd.DataFrame())
                    # 空白列優先顯示
                    blank_opts = [
                        f"▶ 空白列 第{j+1}列（貼上覆蓋）"
                        for j, r in tdf.iterrows()
                        if all(str(r.get(c,"")).strip() == "" for c in ["店號","店名"])
                    ]
                    rows_preview = (blank_opts + ["───────"] +
                                    ["插入在最前面"] + [
                        f"第{j+1}列後（{r.get('店號','')} {r.get('店名','')}）"
                        for j, r in tdf.iterrows()])
                    pos_label = st.selectbox("插入位置", rows_preview, key=f"pos_{i}")

                    col_a, col_b = st.columns(2)
                    with col_a:
                        if st.button("✅ 插入", key=f"ins_{i}"):
                            new_row = dict(vals)
                            new_row["午別"] = am_v
                            new_row["日期"] = dt_v
                            new_row["型態"] = ty_v
                            df = st.session_state.tab_data.get(
                                tdate, pd.DataFrame(columns=cols))
                            nr = {c: new_row.get(c,"") for c in cols}

                            if pos_label.startswith("▶ 空白列"):
                                # 覆蓋空白列
                                p = int(pos_label.split("第")[1].split("列")[0]) - 1
                                for c in cols:
                                    df.at[p, c] = nr[c]
                            elif pos_label == "插入在最前面":
                                df = pd.concat([pd.DataFrame([nr]), df],
                                               ignore_index=True)
                            elif pos_label == "───────":
                                df = pd.concat([df, pd.DataFrame([nr])],
                                               ignore_index=True)
                            else:
                                p = int(pos_label.split("第")[1].split("列")[0])
                                df = pd.concat([df.iloc[:p], pd.DataFrame([nr]),
                                                df.iloc[p:]], ignore_index=True)
                            st.session_state.tab_data[tdate] = df
                            to_remove.append(i)
                            st.rerun()
                    with col_b:
                        if st.button("🗑 移除", key=f"rm_{i}"):
                            to_remove.append(i)
                            st.rerun()

            for i in sorted(set(to_remove), reverse=True):
                st.session_state.cut_list.pop(i)

    # ── 分頁資料表 ────────────────────────────────────
    date_order = st.session_state.date_order
    tabs = st.tabs([f" {d} " for d in date_order])

    for tab, date in zip(tabs, date_order):
        with tab:
            df = st.session_state.tab_data[date].copy()

            # 搜尋高亮提示
            if search_kw:
                mask = df.apply(
                    lambda r: search_kw.lower() in " ".join(str(v) for v in r).lower(), axis=1)
                if mask.any():
                    st.info(f"此分頁有 {mask.sum()} 筆符合「{search_kw}」")

            # 加入勾選欄
            df.insert(0, "✂", False)
            cols = [c for c in df.columns if c != "✂"]

            # ── 新增列表單 ────────────────────────────
            with st.expander("➕ 新增一列"):
                fc = st.columns(3)
                new_am   = fc[0].selectbox("午別", AM_OPTS,   key=f"new_am_{date}")
                new_date = fc[1].text_input("日期", value=date, key=f"new_dt_{date}")
                new_id   = fc[2].text_input("店號",            key=f"new_id_{date}")
                fc2 = st.columns(3)
                new_name = fc2[0].text_input("店名",           key=f"new_nm_{date}")
                new_type = fc2[1].selectbox("型態", TYPE_OPTS, key=f"new_ty_{date}")
                new_note = fc2[2].text_input("備註",           key=f"new_nt_{date}")

                ins_pos = st.selectbox(
                    "插入位置",
                    ["插入在最前面"] + [
                        f"第{j+1}列後（{r.get('店號','')} {r.get('店名','')}）"
                        for j, r in df[cols].iterrows()],
                    key=f"new_pos_{date}")

                if st.button("確定新增", key=f"add_{date}"):
                    new_row = {c: "" for c in cols}
                    new_row.update({"午別": new_am, "日期": new_date,
                                    "店號": new_id, "店名": new_name,
                                    "型態": new_type, "備註": new_note})
                    nr = pd.DataFrame([new_row])
                    base = st.session_state.tab_data[date]
                    if ins_pos == "插入在最前面":
                        st.session_state.tab_data[date] = pd.concat(
                            [nr, base], ignore_index=True)
                    else:
                        p = int(ins_pos.split("第")[1].split("列")[0])
                        st.session_state.tab_data[date] = pd.concat(
                            [base.iloc[:p], nr, base.iloc[p:]], ignore_index=True)
                    st.rerun()

            # ── 剪下 / 刪除按鈕 ──────────────────────
            ba, bb = st.columns([1, 1])
            cut_clicked = ba.button("✂ 剪下勾選列", key=f"cut_{date}")
            del_clicked = bb.button("🗑 刪除勾選列", key=f"del_{date}")

            # ── 資料表 ───────────────────────────────
            edited = st.data_editor(
                df,
                key=f"ed_{date}",
                column_config=col_config(cols),
                column_order=["✂"] + cols,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
            )

            # 處理剪下（保留空白列）
            if cut_clicked:
                result_rows = []
                for _, row in edited.iterrows():
                    if row.get("✂") == True:
                        sid   = str(row.get("店號","")).strip()
                        sname = str(row.get("店名","")).strip()
                        vals  = {c: str(row.get(c,"")) for c in cols}
                        st.session_state.cut_list.append(
                            {"sid": sid, "sname": sname, "vals": vals})
                        # 保留空白列
                        result_rows.append({c: "" for c in cols})
                    else:
                        result_rows.append({c: row.get(c,"") for c in cols})
                st.session_state.tab_data[date] = pd.DataFrame(
                    result_rows, columns=cols)
                st.rerun()

            # 處理刪除
            if del_clicked:
                keep = edited[edited["✂"] != True][cols].reset_index(drop=True)
                st.session_state.tab_data[date] = keep
                st.rerun()

            # 儲存編輯結果
            if not cut_clicked and not del_clicked:
                st.session_state.tab_data[date] = edited[cols].reset_index(drop=True)

            st.caption(f"共 {len(st.session_state.tab_data[date])} 列　｜　"
                       "勾選列後點「✂ 剪下」或「🗑 刪除」")


if __name__ == "__main__":
    main()
