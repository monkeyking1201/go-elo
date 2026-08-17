"""
圍棋隊伍等級分系統 — Desktop Dashboard
技術架構：Streamlit + Google Sheets (gspread)
本機：credentials.json / 雲端：st.secrets["gcp_service_account"]
"""

import os
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# ─── 頁面設定（wide 模式釋放橫向空間）───────────────────────────────────────
st.set_page_config(
    page_title="圍棋等級分系統",
    page_icon="⚫",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Dashboard CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ══ 主容器：最寬 1300px，左右留白 ══ */
.main .block-container {
    max-width: 1300px !important;
    padding: 2.5rem 3rem 3rem !important;
    margin: 0 auto !important;
}

/* ══ 標題 ══ */
h1 { font-size: 2.6rem !important; font-weight: 900 !important; letter-spacing: -0.5px; margin-bottom: 0.3rem !important; }
h3 { font-size: 1.45rem !important; font-weight: 800 !important; margin-bottom: 1.1rem !important; }

/* ══ 分隔線 ══ */
hr { border-color: #e5e7eb !important; margin: 2rem 0 !important; }

/* ══ 面板卡片（HTML class） ══ */
.panel {
    background: #f8f9fb;
    border: 1.5px solid #e4e6ea;
    border-radius: 18px;
    padding: 28px 28px 20px;
    margin-bottom: 8px;
}
.panel-title {
    font-size: 1.4rem;
    font-weight: 800;
    color: #111;
    margin-bottom: 20px;
}

/* ══ 下拉選單：26px 字體、高度 72px ══ */
div[data-baseweb="select"] > div:first-child {
    min-height: 72px !important;
    border-radius: 12px !important;
    padding: 0 20px !important;
    font-size: 24px !important;
    display: flex !important;
    align-items: center !important;
    cursor: pointer !important;
    border: 2px solid #d0d0d0 !important;
    background: #fff !important;
    transition: border-color 0.15s, box-shadow 0.15s !important;
}
div[data-baseweb="select"] > div:first-child:hover {
    border-color: #555 !important;
    box-shadow: 0 0 0 3px rgba(0,0,0,0.06) !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[class*="singleValue"],
div[data-baseweb="select"] div[class*="ValueContainer"] {
    font-size: 24px !important;
    line-height: 1.3 !important;
}
div[data-baseweb="select"] div[class*="placeholder"] {
    font-size: 22px !important;
    color: #bbb !important;
}
/* 下拉清單選項 */
ul[data-baseweb="menu"] li {
    min-height: 62px !important;
    font-size: 22px !important;
    padding: 0 22px !important;
    display: flex !important;
    align-items: center !important;
}
/* 選單 label */
.stSelectbox > label {
    font-size: 1.15rem !important;
    font-weight: 700 !important;
    color: #333 !important;
    margin-bottom: 6px !important;
}

/* ══ 按鈕：大氣感 ══ */
.stButton > button {
    height: 72px !important;
    font-size: 1.35rem !important;
    font-weight: 800 !important;
    border-radius: 12px !important;
    width: 100% !important;
    letter-spacing: 0.04em !important;
    transition: opacity 0.15s, transform 0.1s, box-shadow 0.15s !important;
}
.stButton > button:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.13) !important; }
.stButton > button:active { transform: scale(0.97) !important; }

/* ══ Tabs ══ */
.stTabs [data-baseweb="tab"] {
    font-size: 1.2rem !important;
    font-weight: 700 !important;
    padding: 12px 28px !important;
}

/* ══ Alert 訊息 ══ */
div[data-testid="stAlert"] {
    border-radius: 12px !important;
    font-size: 1.1rem !important;
    padding: 14px 20px !important;
}

/* ══ Caption ══ */
.stCaption { font-size: 1rem !important; }

/* ══ Progress bar ══ */
.stProgress > div > div { border-radius: 8px !important; }

/* ══ Spinner ══ */
.stSpinner > div { font-size: 1.1rem !important; }

/* ══ 選手名單 HTML 表格 ══ */
.elo-table-wrap {
    max-width: 820px;
    margin: 0 auto;
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(0,0,0,0.10);
}
.elo-table {
    width: 100%;
    border-collapse: collapse;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans TC', sans-serif;
}
.elo-table thead tr {
    background: #1e293b;
    color: #f1f5f9;
}
.elo-table th {
    padding: 18px 22px;
    font-size: 20px;
    font-weight: 800;
    letter-spacing: 0.04em;
    border: none;
}
.elo-table td {
    padding: 17px 22px;
    font-size: 24px;
    border-bottom: 1px solid #e8ecf0;
}
.elo-table tbody tr:nth-child(odd)  td { background: #ffffff; }
.elo-table tbody tr:nth-child(even) td { background: #f3f6fd; }
.elo-table tbody tr:hover td {
    background: #dbeafe !important;
    transition: background 0.12s ease;
}
.elo-col-num  { text-align: center; width: 18%; }
.elo-col-name { text-align: left;   width: 52%; font-weight: 500; }
.elo-col-elo  { text-align: center; width: 30%; font-weight: 700; color: #1d4ed8; }
.elo-col-rank { text-align: center; width: 10%; color: #64748b; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ─── Google Sheets ───────────────────────────────────────────────────────────
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
SHEET_PLAYERS = "Players"
SHEET_HISTORY = "History"


@st.cache_resource
def get_gspread_client():
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPE)
    else:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPE
        )
    return gspread.authorize(creds)


def open_spreadsheet():
    client = get_gspread_client()
    sid = st.secrets.get("spreadsheet_id", os.environ.get("SPREADSHEET_ID", ""))
    if not sid:
        st.error("❌ 請在 `.streamlit/secrets.toml` 中設定 `spreadsheet_id`")
        st.stop()
    return client.open_by_key(sid)


def load_players() -> pd.DataFrame:
    ws = open_spreadsheet().worksheet(SHEET_PLAYERS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=["號碼", "姓名", "等級分"])
    df = pd.DataFrame(records)
    df["號碼"] = df["號碼"].astype(int)
    df["等級分"] = df["等級分"].astype(int)
    return df


# ─── Elo ────────────────────────────────────────────────────────────────────
def calculate_elo(r_w: int, r_l: int, k: int = 16):
    e_w = 1 / (1 + 10 ** ((r_l - r_w) / 400))
    delta = round(k * (1 - e_w))
    return r_w + delta, r_l - delta, delta


# ─── 結算 ────────────────────────────────────────────────────────────────────
def update_match(winner_num: int, loser_num: int):
    ss = open_spreadsheet()
    players_ws = ss.worksheet(SHEET_PLAYERS)
    history_ws = ss.worksheet(SHEET_HISTORY)

    records = players_ws.get_all_records()
    df = pd.DataFrame(records)
    df["號碼"] = df["號碼"].astype(int)
    df["等級分"] = df["等級分"].astype(int)

    w_rows = df[df["號碼"] == winner_num]
    l_rows = df[df["號碼"] == loser_num]
    if w_rows.empty:
        return False, f"找不到號碼 {winner_num}"
    if l_rows.empty:
        return False, f"找不到號碼 {loser_num}"

    w, l = w_rows.iloc[0], l_rows.iloc[0]
    r_w, r_l = int(w["等級分"]), int(l["等級分"])
    new_r_w, new_r_l, delta = calculate_elo(r_w, r_l)

    history_ws.append_row([winner_num, str(w["姓名"]), r_w, loser_num, str(l["姓名"]), r_l])
    players_ws.update_cell(w_rows.index[0] + 2, 3, new_r_w)
    players_ws.update_cell(l_rows.index[0] + 2, 3, new_r_l)

    return True, (
        f"✅ **{w['姓名']}** {r_w} → **{new_r_w}** (+{delta})　｜　"
        f"**{l['姓名']}** {r_l} → **{new_r_l}** (-{delta})"
    )


# ─── 復原 ────────────────────────────────────────────────────────────────────
def undo_last():
    ss = open_spreadsheet()
    players_ws = ss.worksheet(SHEET_PLAYERS)
    history_ws = ss.worksheet(SHEET_HISTORY)

    all_vals = history_ws.get_all_values()
    if len(all_vals) < 2:
        return False, "沒有可以復原的紀錄"

    last = all_vals[-1]
    try:
        w_num, w_name, old_r_w = int(last[0]), last[1], int(last[2])
        l_num, l_name, old_r_l = int(last[3]), last[4], int(last[5])
    except (IndexError, ValueError):
        return False, "History 資料格式有誤"

    records = players_ws.get_all_records()
    df = pd.DataFrame(records)
    df["號碼"] = df["號碼"].astype(int)

    w_rows = df[df["號碼"] == w_num]
    l_rows = df[df["號碼"] == l_num]
    if w_rows.empty or l_rows.empty:
        return False, "找不到選手資料"

    players_ws.update_cell(w_rows.index[0] + 2, 3, old_r_w)
    players_ws.update_cell(l_rows.index[0] + 2, 3, old_r_l)
    history_ws.delete_rows(len(all_vals))

    return True, f"↩️ 已復原：**{w_name}** → {old_r_w}　｜　**{l_name}** → {old_r_l}"


# ════════════════════════════════════════════════════════════════════════════
#  主介面
# ════════════════════════════════════════════════════════════════════════════
st.title("⚫⚪ 圍棋隊伍等級分系統")
st.caption("戰情儀表板　｜　Elo K=16")

# 載入選手
df_all = load_players()
player_opts = [
    f"{int(row['號碼']):02d} - {row['姓名']}"
    for _, row in df_all.sort_values("號碼").iterrows()
]

winner_opts = [None] + player_opts
loser_opts  = [None] + player_opts

def fmt_w(x): return "── 請選擇勝者 ──" if x is None else x
def fmt_l(x): return "── 請選擇敗者 ──" if x is None else x
def parse_num(opt: str) -> int: return int(opt.split(" - ")[0])


# ══════════════════════════════════════════════════════════════
#  上半部：左右並排（單筆 ｜ 批次）
# ══════════════════════════════════════════════════════════════
col_left, col_right = st.columns([1, 1], gap="large")

# ── 左側：單筆結算 ─────────────────────────────────────────
with col_left:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown("### 📝 單筆結算")

    c1, c2 = st.columns(2)
    with c1:
        winner_sel = st.selectbox(
            "🏆 勝者", winner_opts, format_func=fmt_w, index=0, key="single_winner"
        )
    with c2:
        loser_sel = st.selectbox(
            "❌ 敗者", loser_opts, format_func=fmt_l, index=0, key="single_loser"
        )

    st.markdown("<br>", unsafe_allow_html=True)
    settle = st.button("⚡ 結算", type="primary", use_container_width=True, key="btn_settle")

    st.markdown("<br>", unsafe_allow_html=True)
    col_u, col_sp = st.columns([1, 2])
    with col_u:
        undo = st.button("↩️ 復原上一步", use_container_width=True, key="btn_undo")

    st.markdown('</div>', unsafe_allow_html=True)

    # 防呆驗證
    if settle:
        if winner_sel is None and loser_sel is None:
            st.error("⚠️ 請選擇勝者與敗者")
        elif winner_sel is None:
            st.error("⚠️ 請選擇勝者")
        elif loser_sel is None:
            st.error("⚠️ 請選擇敗者")
        elif winner_sel == loser_sel:
            st.error("⚠️ 勝者與敗者不能是同一人！")
        else:
            with st.spinner("更新中…"):
                ok, msg = update_match(parse_num(winner_sel), parse_num(loser_sel))
            (st.success if ok else st.error)(msg)
            if ok:
                st.rerun()

    if undo:
        with st.spinner("復原中…"):
            ok, msg = undo_last()
        (st.success if ok else st.warning)(msg)
        if ok:
            st.rerun()


# ── 右側：批次輸入 ─────────────────────────────────────────
with col_right:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown("### 📋 批次輸入")
    st.caption("空白行自動略過，依上到下順序依序結算。")

    if "batch_rows" not in st.session_state:
        st.session_state.batch_rows = 4

    ca, cb = st.columns(2)
    with ca:
        if st.button("➕ 新增一行", use_container_width=True):
            st.session_state.batch_rows += 1
            st.rerun()
    with cb:
        if st.button("🗑 清除所有行", use_container_width=True):
            st.session_state.batch_rows = 4
            for i in range(50):
                st.session_state.pop(f"bw_{i}", None)
                st.session_state.pop(f"bl_{i}", None)
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    batch_pairs = []
    for i in range(st.session_state.batch_rows):
        r1, r2 = st.columns(2)
        with r1:
            bw = st.selectbox(f"🏆 勝者 {i+1}", winner_opts, format_func=fmt_w, index=0, key=f"bw_{i}")
        with r2:
            bl = st.selectbox(f"❌ 敗者 {i+1}", loser_opts, format_func=fmt_l, index=0, key=f"bl_{i}")
        batch_pairs.append((bw, bl))

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("⚡ 批次結算", type="primary", use_container_width=True):
        valid = [(w, l) for w, l in batch_pairs if w is not None or l is not None]
        if not valid:
            st.warning("⚠️ 請至少填寫一行")
        else:
            ok_msgs, err_msgs = [], []
            prog = st.progress(0)
            for i, (ws, ls) in enumerate(valid):
                label = f"第 {i+1} 行"
                if ws is None:
                    err_msgs.append(f"{label}：未選擇勝者")
                elif ls is None:
                    err_msgs.append(f"{label}：未選擇敗者")
                elif ws == ls:
                    err_msgs.append(f"{label}：勝敗同一人（{ws}）")
                else:
                    ok, msg = update_match(parse_num(ws), parse_num(ls))
                    (ok_msgs if ok else err_msgs).append(f"{label}：{msg}")
                prog.progress((i + 1) / len(valid))
            prog.empty()
            if ok_msgs:
                st.success(f"✅ 完成 {len(ok_msgs)} 筆")
                for m in ok_msgs:
                    st.markdown(f"- {m}")
            if err_msgs:
                st.error(f"⚠️ {len(err_msgs)} 筆有問題")
                for m in err_msgs:
                    st.markdown(f"- {m}")
            if ok_msgs:
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  下半部：選手名單（HTML 表格，置中 820px，斑馬紋 + Hover）
# ══════════════════════════════════════════════════════════════
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown("### 📋 選手名單")


TABLE_CSS = """
<style>
.elo-table-wrap{max-width:820px;margin:0 auto;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.10);}
.elo-table{width:100%;border-collapse:collapse;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Noto Sans TC',sans-serif;}
.elo-table thead tr{background:#1e293b;color:#f1f5f9;}
.elo-table th{padding:18px 22px;font-size:20px;font-weight:800;letter-spacing:.04em;border:none;white-space:nowrap;}
.elo-table td{padding:17px 22px;font-size:24px;border-bottom:1px solid #e8ecf0;}
.elo-table tbody tr:nth-child(odd) td{background:#ffffff;}
.elo-table tbody tr:nth-child(even) td{background:#f3f6fd;}
.elo-table tbody tr:hover td{background:#dbeafe!important;transition:background .12s ease;}
.elo-col-num{text-align:center;width:18%;}
.elo-col-name{text-align:left;width:52%;font-weight:500;}
.elo-col-elo{text-align:center;width:30%;font-weight:700;color:#1d4ed8;}
.elo-col-rank{text-align:center;width:10%;color:#64748b;font-weight:600;}
</style>
"""

def build_table(df: pd.DataFrame, show_rank: bool = False) -> str:
    rank_th = '<th class="elo-col-rank">排名</th>' if show_rank else ""
    header = (
        "<thead><tr>"
        + rank_th
        + '<th class="elo-col-num">號碼</th>'
        + '<th class="elo-col-name">姓名</th>'
        + '<th class="elo-col-elo">等級分</th>'
        + "</tr></thead>"
    )
    rows = ""
    for _, row in df.iterrows():
        rank_td = f'<td class="elo-col-rank">{int(row["排名"])}</td>' if show_rank else ""
        rows += (
            "<tr>"
            + rank_td
            + f'<td class="elo-col-num">{int(row["號碼"])}</td>'
            + f'<td class="elo-col-name">{row["姓名"]}</td>'
            + f'<td class="elo-col-elo">{int(row["等級分"])}</td>'
            + "</tr>"
        )
    return (
        TABLE_CSS
        + '<div class="elo-table-wrap">'
        + '<table class="elo-table">'
        + header
        + "<tbody>" + rows + "</tbody>"
        + "</table></div>"
    )


if df_all.empty:
    st.info("📭 尚無選手資料")
else:
    tab_num, tab_rank, tab_new = st.tabs(["🔢 號碼排序", "🏅 英雄榜排名", "🌟 新銳隊英雄榜"])

    with tab_num:
        df_n = df_all.sort_values("號碼").reset_index(drop=True)
        st.html(build_table(df_n, show_rank=False))

    with tab_rank:
        df_r = df_all.sort_values("等級分", ascending=False).reset_index(drop=True)
        df_r.insert(0, "排名", range(1, len(df_r) + 1))
        st.html(build_table(df_r, show_rank=True))

    with tab_new:
        NEW_TEAM = [10, 11, 12, 13, 14, 15, 16, 17, 43]
        df_new = (
            df_all[df_all["號碼"].isin(NEW_TEAM)]
            .sort_values("等級分", ascending=False)
            .reset_index(drop=True)
        )
        df_new.insert(0, "排名", range(1, len(df_new) + 1))
        st.caption("新銳隊：10–17 號選手，依等級分排名")
        st.html(build_table(df_new, show_rank=True))
