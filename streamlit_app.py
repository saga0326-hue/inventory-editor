# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

st.set_page_config(page_title="盤點資料編輯器", page_icon="📋", layout="wide")

st.markdown("""
<style>
/* 禁止欄位拖移（隱藏拖曳游標） */
[data-testid="column-header-cell"] { cursor: default !important; }
[class*="dvn-drag"] { display: none !important; }

[data-testid="stSidebar"] { background:#2e1e1e; }
[data-testid="stSidebar"] * { color:#cdd6f4; }
.dup-bar  { background:#3a1010; border-radius:8px; padding:10px 16px; margin-bottom:8px; }
.dup-chip { background:#5a1a1a; border-radius:6px; padding:6px 12px;
            display:inline-block; margin:3px; }
.dup-id   { color:#ff5555; font-weight:bold; font-size:13px; }
.dup-date { color:#f39c12; font-size:12px; }
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
COLS      = list(FIELD_MAP.values())
TYPE_OPTS = ["", "閉", "轉", "解", "續", "FC", "RC"]
AM_OPTS   = ["", "上午", "下午"]
def fmt_date(v):
    try:
        dt = pd.to_datetime(str(v).strip())
        return f"{dt.month}/{dt.day}"
    except Exception:
        return str(v).strip()


def init():
    defaults = {"tab_data": {}, "date_order": [], "cut_list": [],
                "inv_key": "", "chg_key": "", "na_open": False}
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


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
        df["_s"] = pd.to_datetime(
            df["日期"].map(lambda v: f"2026/{v}"), errors="coerce")
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
               "店名":"店名","型態":"型態","備註":"備註","課別":"課別"}
    entries = []
    for _, row in df.iterrows():
        sid   = str(row.get("店號","")).strip()
        sname = str(row.get("店名","")).strip()
        if not sid and not sname:
            continue
        vals = {c: "" for c in cols}
        for src, dst in col_map.items():
            if src in df.columns and dst in cols:
                v = str(row.get(src,"")).strip()
                if src == "午別":
                    v = "上午" if v=="1" else ("下午" if v=="2" else v)
                elif src == "日期":
                    v = fmt_date(v)
                vals[dst] = v
        entries.append({"sid": sid, "sname": sname,
                        "vals": vals, "note": ""})
    return entries


def find_dups(tab_data):
    store_map = defaultdict(list)
    for date, df in tab_data.items():
        if "店號" not in df.columns:
            continue
        for _, row in df.iterrows():
            sid   = str(row.get("店號","")).strip()
            sname = str(row.get("店名","")).strip()
            if sid and sid not in ("", "nan"):
                store_map[sid].append((date, sname))
    return {s: e for s, e in store_map.items() if len(e) > 1}


def to_excel(tab_data, date_order):
    frames = [tab_data[d] for d in date_order
              if d in tab_data and len(tab_data[d]) > 0]
    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def style_by_am(df):
    """上午藍色、下午橘色，空白列無色"""
    def row_color(row):
        am = str(row.get("午別", "")).strip()
        if am == "上午":
            return ["background-color:#0d2f4f; color:#89c4e1"] * len(row)
        elif am == "下午":
            return ["background-color:#3d1f00; color:#ffb347"] * len(row)
        return ["color:#8888aa"] * len(row)
    return df.style.apply(row_color, axis=1)


def get_col_config(cols):
    cfg = {"✂": st.column_config.CheckboxColumn("選取", width="small", default=False)}
    for c in cols:
        if c == "午別":
            cfg[c] = st.column_config.SelectboxColumn(c, options=AM_OPTS, width="small")
        elif c == "型態":
            cfg[c] = st.column_config.SelectboxColumn(c, options=TYPE_OPTS, width="small")
        elif c in ("日期","店號","課別"):
            cfg[c] = st.column_config.TextColumn(c, width="small")
        else:
            cfg[c] = st.column_config.TextColumn(c, width="medium")
    return cfg


def render_sidebar(cols):
    with st.sidebar:
        st.markdown("### ✂ 剪下記錄")

        # ── 新增資料到剪貼區 ──────────────────────────
        with st.expander("➕ 新增資料",
                         expanded=st.session_state.get("na_open", False)):
            na_am   = st.selectbox("午別", AM_OPTS, key="na_am")
            na_dt   = st.text_input("日期", key="na_dt")
            na_id   = st.text_input("店號", key="na_id")
            na_nm   = st.text_input("店名", key="na_nm")
            na_ty   = st.selectbox("型態", TYPE_OPTS, key="na_ty")
            na_nt   = st.text_input("備註", key="na_nt")
            na_cl   = st.text_input("課別", key="na_cl")
            na_note = st.text_input("說明（僅顯示在剪貼區）", key="na_note")
            if st.button("➕ 加入剪貼區", key="na_add"):
                vals = {c: "" for c in cols}
                for k, v in [("午別",na_am),("日期",na_dt),("店號",na_id),
                              ("店名",na_nm),("型態",na_ty),("備註",na_nt),
                              ("課別",na_cl)]:
                    if k in cols:
                        vals[k] = v
                st.session_state.cut_list.append({
                    "sid":   na_id,
                    "sname": na_nm,
                    "vals":  vals,
                    "note":  na_note,
                })
                st.session_state.na_open = False   # 收起展開面板
                st.rerun()

        st.divider()

        if not st.session_state.cut_list:
            st.caption("目前沒有剪下的資料")
            return

        date_order = st.session_state.date_order

        for i, entry in enumerate(list(st.session_state.cut_list)):
            sid, sname = entry["sid"], entry["sname"]
            vals = entry["vals"]
            note = entry.get("note", "")
            tag  = "  ".join(filter(None, [vals.get("日期",""),
                                           vals.get("午別",""),
                                           vals.get("型態","")]))
            label = f"**{sid}** {sname}" + (f" ｜ {tag}" if tag else "")
            if note:
                label += f"\n_{note}_"

            with st.expander(label):
                if note:
                    st.caption(f"📝 說明：{note}")
                am_v = st.selectbox("午別", AM_OPTS,
                    index=AM_OPTS.index(vals.get("午別",""))
                          if vals.get("午別","") in AM_OPTS else 0,
                    key=f"am_{i}")
                dt_v = st.text_input("日期",
                    value=vals.get("日期",""), key=f"dt_{i}")
                ty_v = st.selectbox("型態", TYPE_OPTS,
                    index=TYPE_OPTS.index(vals.get("型態",""))
                          if vals.get("型態","") in TYPE_OPTS else 0,
                    key=f"ty_{i}")

                tdate = st.selectbox("插入到分頁",
                    date_order, key=f"td_{i}")
                tdf   = st.session_state.tab_data.get(tdate, pd.DataFrame())

                blank_opts = [
                    f"▶ 空白列 第{j+1}列（覆蓋）"
                    for j, r in tdf.iterrows()
                    if all(str(r.get(c,"")).strip() in ("","nan")
                           for c in ["店號","店名"])
                ]
                row_opts = (blank_opts +
                            (["───────"] if blank_opts else []) +
                            ["插入在最前面"] +
                            [f"第{j+1}列後（{str(r.get('店號',''))} "
                             f"{str(r.get('店名',''))}）"
                             for j, r in tdf.iterrows()])
                pos_label = st.selectbox("插入位置",
                    row_opts, key=f"pos_{i}")

                # ── 日期不符提示 ──────────────────────
                date_ok = (not dt_v) or (dt_v == tdate)
                if not date_ok:
                    st.warning(
                        f"⚠️ 日期不符：資料日期 **{dt_v}**，"
                        f"分頁日期 **{tdate}**")

                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("✅ 插入", key=f"ins_{i}"):
                        new_vals = dict(vals)
                        new_vals.update({"午別": am_v,
                                         "日期": dt_v,
                                         "型態": ty_v})
                        df = st.session_state.tab_data.get(
                            tdate, pd.DataFrame(columns=cols)).copy()
                        nr = pd.DataFrame(
                            [{c: new_vals.get(c,"") for c in cols}])

                        if pos_label.startswith("▶ 空白列"):
                            p = int(pos_label.split("第")[1].split("列")[0]) - 1
                            for c in cols:
                                df.at[p, c] = new_vals.get(c,"")
                        elif pos_label in ("插入在最前面", "───────"):
                            df = pd.concat([nr, df], ignore_index=True)
                        else:
                            p = int(pos_label.split("第")[1].split("列")[0])
                            df = pd.concat(
                                [df.iloc[:p], nr, df.iloc[p:]],
                                ignore_index=True)

                        st.session_state.tab_data[tdate] = df
                        st.session_state.cut_list.pop(i)
                        st.rerun()

                with col_b:
                    if st.button("🗑 移除", key=f"rm_{i}"):
                        st.session_state.cut_list.pop(i)
                        st.rerun()


# ════════════════════════════════════════════════════
def main():
    init()

    # 縮小標題
    st.markdown("#### 📋 盤點資料編輯器")

    # ── 工具列 ────────────────────────────────────────
    c1, c2, c3 = st.columns([2, 2, 2])

    with c1:
        inv_file = st.file_uploader("📂 匯入盤點表", type=["xlsx","xls"],
                                    label_visibility="collapsed")
        if inv_file:
            file_key = f"{inv_file.name}_{inv_file.size}"
            if st.session_state.inv_key != file_key:
                st.session_state.inv_key = file_key
                td, do = load_inventory(inv_file)
                st.session_state.tab_data   = td
                st.session_state.date_order = do
                st.success(f"載入 {len(do)} 個日期分頁")

    with c2:
        chg_file = st.file_uploader("📋 匯入異動名單", type=["xlsx","xls"],
                                    label_visibility="collapsed")
        if chg_file and st.session_state.tab_data:
            file_key = f"{chg_file.name}_{chg_file.size}"
            if st.session_state.chg_key != file_key:
                st.session_state.chg_key = file_key
                cols    = list(list(st.session_state.tab_data.values())[0].columns)
                entries = load_change(chg_file, cols)
                st.session_state.cut_list.extend(entries)
                st.success(f"異動名單載入 {len(entries)} 筆至左側剪貼區")

    with c3:
        if st.session_state.tab_data:
            st.download_button("💾 另存新檔",
                data=to_excel(st.session_state.tab_data,
                              st.session_state.date_order),
                file_name="盤點表_已編輯.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet")

    if not st.session_state.tab_data:
        st.info("請先匯入盤點表 Excel 檔案。")
        # 還是要渲染側邊欄，讓使用者可以手動新增資料
        render_sidebar([])
        return

    st.divider()

    # ── 重複店號警示 ──────────────────────────────────
    dups = find_dups(st.session_state.tab_data)
    if dups:
        chips = "".join(
            f'<div class="dup-chip">'
            f'<div class="dup-id">{sid} {ents[0][1]}</div>'
            f'<div class="dup-date">出現於：{"、".join(e[0] for e in ents)}</div>'
            f'</div>'
            for sid, ents in sorted(dups.items()))
        st.markdown(
            f'<div class="dup-bar">⚠️ <b style="color:#ff5555">'
            f'重複店號：共 {len(dups)} 間</b><br>{chips}</div>',
            unsafe_allow_html=True)
    else:
        st.success("✅ 無重複店號")

    # ── 搜尋列 ────────────────────────────────────────
    search_kw = st.text_input("🔍 搜尋（跨所有分頁）",
                               placeholder="輸入店號、店名、日期…")
    if search_kw:
        results = []
        for date, df in st.session_state.tab_data.items():
            mask = df.apply(
                lambda r: search_kw.lower() in
                " ".join(str(v) for v in r).lower(), axis=1)
            if mask.any():
                results.append(f"**{date}**：{mask.sum()} 筆")
        st.info("🔍 " + "　｜　".join(results) if results
                else f"找不到「{search_kw}」")

    st.divider()

    # ── 側邊欄 ────────────────────────────────────────
    cols_global = list(list(st.session_state.tab_data.values())[0].columns)
    render_sidebar(cols_global)

    # ── 分頁資料表 ────────────────────────────────────
    date_order = st.session_state.date_order
    tabs = st.tabs([f" {d} " for d in date_order])

    for tab, date in zip(tabs, date_order):
        with tab:
            df   = st.session_state.tab_data[date].copy()
            cols = list(df.columns)

            if search_kw:
                mask = df.apply(
                    lambda r: search_kw.lower() in
                    " ".join(str(v) for v in r).lower(), axis=1)
                if mask.any():
                    st.info(f"此分頁有 {mask.sum()} 筆符合「{search_kw}」")

            # ── 彩色預覽（上午藍／下午橘）────────────
            st.markdown(
                '<span style="display:inline-block;width:14px;height:14px;'
                'background:#0d2f4f;border-radius:3px;margin-right:4px"></span>'
                '<span style="color:#89c4e1;font-size:12px">上午</span>'
                '&nbsp;&nbsp;'
                '<span style="display:inline-block;width:14px;height:14px;'
                'background:#3d1f00;border-radius:3px;margin-right:4px"></span>'
                '<span style="color:#ffb347;font-size:12px">下午</span>',
                unsafe_allow_html=True)
            st.dataframe(style_by_am(df), use_container_width=True,
                         hide_index=True)

            st.caption("⬆ 彩色預覽（唯讀）　｜　⬇ 勾選後可剪下／刪除")

            # ── 剪下 / 刪除按鈕 ──────────────────────
            ba, bb, _ = st.columns([1, 1, 4])
            cut_clicked = ba.button("✂ 剪下勾選列", key=f"cut_{date}")
            del_clicked = bb.button("🗑 刪除勾選列", key=f"del_{date}")

            # 加入勾選欄 + 隱藏順序欄（防止點欄位標題排序後順序亂掉）
            df_show = df.copy()
            df_show.insert(0, "✂", False)
            df_show.insert(1, "_idx", range(len(df_show)))

            edited = st.data_editor(
                df_show,
                key=f"ed_{date}",
                column_config={
                    **get_col_config(cols),
                    "_idx": st.column_config.Column(disabled=True,
                                                    width=None,
                                                    label="_idx"),
                },
                column_order=["✂"] + cols,   # 隱藏 _idx（不列入顯示順序）
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
            )

            # 無論是否排序，都依 _idx 還原原始順序
            if "_idx" in edited.columns:
                edited = edited.sort_values("_idx").reset_index(drop=True)

            # 處理剪下（保留空白列）
            if cut_clicked:
                result_rows = []
                for _, row in edited.iterrows():
                    if row.get("✂") == True:
                        sid   = str(row.get("店號","")).strip()
                        sname = str(row.get("店名","")).strip()
                        vals  = {c: str(row.get(c,"")) for c in cols}
                        st.session_state.cut_list.append(
                            {"sid": sid, "sname": sname,
                             "vals": vals, "note": ""})
                        result_rows.append({c: "" for c in cols})
                    else:
                        result_rows.append({c: row.get(c,"") for c in cols})
                st.session_state.tab_data[date] = pd.DataFrame(
                    result_rows, columns=cols)
                st.rerun()

            elif del_clicked:
                keep = edited[edited["✂"] != True][cols].reset_index(drop=True)
                st.session_state.tab_data[date] = keep
                st.rerun()

            else:
                st.session_state.tab_data[date] = \
                    edited[cols].reset_index(drop=True)


if __name__ == "__main__":
    main()
