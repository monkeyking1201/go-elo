"""
勝負預測系統 — Elo 勝率分析與對局追蹤（Google Sheets 版）
"""

import os, uuid, time
from datetime import date, datetime

import streamlit as st
import gspread
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials
import pandas as pd

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="勝負預測", page_icon="🎯",
    layout="wide", initial_sidebar_state="collapsed",
)
st.markdown("""
<style>
.main .block-container{max-width:1300px!important;padding:2rem 3rem!important;}
h1{font-size:2.2rem!important;font-weight:900!important;}
h3{font-size:1.3rem!important;font-weight:800!important;}
.stButton>button{height:52px!important;font-size:1rem!important;font-weight:700!important;border-radius:12px!important;width:100%!important;}
div[data-baseweb="select"]>div:first-child{min-height:58px!important;border-radius:12px!important;font-size:19px!important;padding:0 16px!important;}
div[data-baseweb="select"] span{font-size:19px!important;}
ul[data-baseweb="menu"] li{min-height:52px!important;font-size:19px!important;}
.stSelectbox>label,.stTextInput>label,.stNumberInput>label,.stCheckbox>label,.stRadio>label{font-size:1rem!important;font-weight:700!important;}
.stNumberInput input,.stTextInput input{font-size:1.15rem!important;height:58px!important;}
.prob-box{text-align:center;border-radius:16px;padding:22px 12px;font-size:3.4rem;font-weight:900;line-height:1.1;}
.prob-label{font-size:.95rem;font-weight:600;margin-bottom:4px;color:#555;}
</style>
""", unsafe_allow_html=True)

# ─── 常數 ─────────────────────────────────────────────────────────────────────
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# birth_year 加在最後，確保舊資料欄位位置不變
OPPONENT_HDRS = ["id","name","nationality","rating","is_estimate","created_at","birth_year"]
MATCH_HDRS    = [
    "id","date","event","taiwan_num","taiwan_name","taiwan_rating",
    "opp_id","opp_name","opp_nationality","opp_rating","opp_is_estimate",
    "time_control","base_win_prob","system_win_prob",
    "my_prediction","actual_result","created_at",
    "taiwan_birth_year","opp_birth_year",   # 新增，放最後
]
SETTINGS_HDRS    = ["key","value"]
DEFAULT_SETTINGS = {"compression_one": 0.85, "compression_both": 0.75}
AGE_BRACKETS     = ["25歲以下","25–29","30–34","35–39","40–44","45歲以上"]

# ─── Google Sheets 連線 ───────────────────────────────────────────────────────
def api_retry(func, *args, retries=4, **kwargs):
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except APIError:
            if attempt == retries - 1: raise
            time.sleep(2 ** attempt)

@st.cache_resource
def get_gs_client():
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPE)
    else:
        creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPE)
    return gspread.authorize(creds)

def open_ss():
    sid = st.secrets.get("spreadsheet_id", os.environ.get("SPREADSHEET_ID",""))
    if not sid:
        st.error("❌ 未設定 spreadsheet_id"); st.stop()
    return get_gs_client().open_by_key(sid)

def get_ws(name, headers):
    ss = open_ss()
    try:
        return ss.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=1000, cols=26)
        api_retry(ws.append_row, headers)
        return ws

def ensure_col(ws, col_name):
    """在工作表標題列加入缺少的欄位（只執行一次，之後自動跳過）"""
    hdrs = api_retry(ws.row_values, 1)
    if col_name not in hdrs:
        api_retry(ws.update_cell, 1, len(hdrs) + 1, col_name)

# ─── 資料讀取 ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def load_opponents():
    ws = get_ws("Opponents", OPPONENT_HDRS[:6])
    ensure_col(ws, "birth_year")
    rows = api_retry(ws.get_all_records)
    for r in rows:
        r["is_estimate"] = str(r.get("is_estimate","")).upper() == "TRUE"
        r["rating"]      = int(r.get("rating") or 0)
        r["birth_year"]  = int(r.get("birth_year") or 0)
    return rows

@st.cache_data(ttl=30)
def load_matches():
    ws = get_ws("Matches", MATCH_HDRS[:17])
    ensure_col(ws, "taiwan_birth_year")
    ensure_col(ws, "opp_birth_year")
    rows = api_retry(ws.get_all_records)
    for r in rows:
        r["opp_is_estimate"]   = str(r.get("opp_is_estimate","")).upper() == "TRUE"
        r["actual_result"]     = r["actual_result"] if r.get("actual_result") else None
        for k in ["taiwan_rating","opp_rating"]:
            r[k] = int(r.get(k) or 0)
        for k in ["base_win_prob","system_win_prob"]:
            r[k] = float(r.get(k) or 0)
        for k in ["taiwan_birth_year","opp_birth_year"]:
            r[k] = int(r.get(k) or 0)
    return rows

@st.cache_data(ttl=60)
def load_settings():
    ws   = get_ws("PredSettings", SETTINGS_HDRS)
    rows = api_retry(ws.get_all_records)
    cfg  = DEFAULT_SETTINGS.copy()
    for r in rows:
        if r.get("key") in cfg:
            cfg[r["key"]] = float(r["value"])
    return cfg

@st.cache_data(ttl=60)
def load_taiwan_players():
    sid = st.secrets.get("spreadsheet_id", os.environ.get("SPREADSHEET_ID",""))
    if not sid:
        return pd.DataFrame(columns=["號碼","姓名","等級分","birth_year"])
    try:
        ws = get_gs_client().open_by_key(sid).worksheet("Players")
        df = pd.DataFrame(api_retry(ws.get_all_records))
        df["號碼"]  = df["號碼"].astype(int)
        df["等級分"] = df["等級分"].astype(int)
        if "birth_year" in df.columns:
            df["birth_year"] = pd.to_numeric(df["birth_year"], errors="coerce").fillna(0).astype(int)
        else:
            df["birth_year"] = 0
        return df
    except Exception as e:
        st.warning(f"⚠️ 無法讀取選手資料：{e}")
        return pd.DataFrame(columns=["號碼","姓名","等級分","birth_year"])

# ─── 資料寫入 ─────────────────────────────────────────────────────────────────
def append_opponent(opp):
    ws = get_ws("Opponents", OPPONENT_HDRS[:6])
    ensure_col(ws, "birth_year")
    api_retry(ws.append_row, [
        opp["id"], opp["name"], opp["nationality"],
        opp["rating"], str(opp["is_estimate"]), opp["created_at"],
        opp.get("birth_year") or "",
    ])
    st.cache_data.clear()

def update_opponent_row(opp_id, rating, nationality, is_estimate, birth_year):
    ws      = get_ws("Opponents", OPPONENT_HDRS[:6])
    ensure_col(ws, "birth_year")
    headers = api_retry(ws.row_values, 1)
    ids     = api_retry(ws.col_values, 1)
    for i, val in enumerate(ids):
        if val == opp_id:
            row = i + 1
            api_retry(ws.update_cell, row, headers.index("rating")      + 1, rating)
            api_retry(ws.update_cell, row, headers.index("nationality") + 1, nationality)
            api_retry(ws.update_cell, row, headers.index("is_estimate") + 1, str(is_estimate))
            if "birth_year" in headers:
                api_retry(ws.update_cell, row, headers.index("birth_year") + 1, birth_year or "")
            st.cache_data.clear()
            return

def delete_opponent_row(opp_id):
    ws  = get_ws("Opponents", OPPONENT_HDRS[:6])
    ids = api_retry(ws.col_values, 1)
    for i, val in enumerate(ids):
        if val == opp_id:
            api_retry(ws.delete_rows, i + 1)
            st.cache_data.clear()
            return

def append_match(m):
    ws = get_ws("Matches", MATCH_HDRS[:17])
    ensure_col(ws, "taiwan_birth_year")
    ensure_col(ws, "opp_birth_year")
    api_retry(ws.append_row, [
        m["id"], m["date"], m["event"],
        m["taiwan_num"], m["taiwan_name"], m["taiwan_rating"],
        m["opp_id"], m["opp_name"], m["opp_nationality"],
        m["opp_rating"], str(m["opp_is_estimate"]),
        m["time_control"], m["base_win_prob"], m["system_win_prob"],
        m["my_prediction"], "", m["created_at"],
        m.get("taiwan_birth_year") or "", m.get("opp_birth_year") or "",
    ])
    st.cache_data.clear()

def fill_match_result(match_id, result):
    ws      = get_ws("Matches", MATCH_HDRS[:17])
    headers = api_retry(ws.row_values, 1)
    if "actual_result" not in headers: return
    col = headers.index("actual_result") + 1
    ids = api_retry(ws.col_values, 1)
    for i, val in enumerate(ids):
        if val == match_id:
            api_retry(ws.update_cell, i + 1, col, result)
            st.cache_data.clear()
            return

def save_settings(c1, c2):
    ws = get_ws("PredSettings", SETTINGS_HDRS)
    api_retry(ws.clear)
    api_retry(ws.append_row, SETTINGS_HDRS)
    api_retry(ws.append_row, ["compression_one",  c1])
    api_retry(ws.append_row, ["compression_both", c2])
    st.cache_data.clear()

# ─── Elo 計算 ─────────────────────────────────────────────────────────────────
def calc_win_prob(r_tw, r_opp, opp_is_est, tw_is_est=False, settings=None):
    if settings is None: settings = DEFAULT_SETTINGS
    base = 1 / (1 + 10 ** ((r_opp - r_tw) / 400))
    if opp_is_est and tw_is_est: coeff = settings["compression_both"]
    elif opp_is_est or tw_is_est: coeff = settings["compression_one"]
    else: coeff = 1.0
    adj = 0.5 + (base - 0.5) * coeff if coeff < 1.0 else base
    return round(base * 100, 1), round(adj * 100, 1)

def prob_color(p):
    if p >= 65: return "#16a34a"
    if p >= 55: return "#65a30d"
    if p >= 45: return "#ca8a04"
    if p >= 35: return "#ea580c"
    return "#dc2626"

def prob_bg(p):
    if p >= 65: return "#f0fdf4"
    if p >= 55: return "#f7fee7"
    if p >= 45: return "#fefce8"
    if p >= 35: return "#fff7ed"
    return "#fff1f2"

# ─── 年齡分析 helpers ─────────────────────────────────────────────────────────
def age_bracket(age):
    if age < 25: return "25歲以下"
    if age < 30: return "25–29"
    if age < 35: return "30–34"
    if age < 40: return "35–39"
    if age < 45: return "40–44"
    return "45歲以上"

def build_age_records(completed):
    """每場對局拆成兩筆（台灣選手 + 對手），回傳 list of dict"""
    records = []
    for m in completed:
        try:
            yr = int(m["date"][:4])
        except Exception:
            continue
        tw_wins = m["actual_result"] == "台灣選手勝"
        # base_win_prob = 未修正的台灣勝率
        base_p = float(m.get("base_win_prob") or m.get("system_win_prob") or 50) / 100

        tw_by = int(m.get("taiwan_birth_year") or 0)
        if tw_by > 0:
            tw_age = yr - tw_by
            if 10 <= tw_age <= 70:
                records.append({
                    "side":        "台灣選手",
                    "name":        m["taiwan_name"],
                    "age_bracket": age_bracket(tw_age),
                    "predicted":   base_p,
                    "actual":      1 if tw_wins else 0,
                })

        opp_by = int(m.get("opp_birth_year") or 0)
        if opp_by > 0:
            opp_age = yr - opp_by
            if 10 <= opp_age <= 70:
                records.append({
                    "side":        "對手",
                    "name":        m["opp_name"],
                    "age_bracket": age_bracket(opp_age),
                    "predicted":   1 - base_p,
                    "actual":      0 if tw_wins else 1,
                })
    return records

def build_diff_records(completed):
    """回傳有雙方年齡的對局，用於表二（年齡差分析）"""
    records = []
    for m in completed:
        try:
            yr = int(m["date"][:4])
        except Exception:
            continue
        tw_by  = int(m.get("taiwan_birth_year") or 0)
        opp_by = int(m.get("opp_birth_year") or 0)
        if tw_by <= 0 or opp_by <= 0:
            continue
        diff    = (yr - tw_by) - (yr - opp_by)   # 台灣年齡 − 對手年齡
        tw_wins = m["actual_result"] == "台灣選手勝"
        base_p  = float(m.get("base_win_prob") or m.get("system_win_prob") or 50) / 100
        records.append({"diff": diff, "tw_win": 1 if tw_wins else 0, "base_p": base_p})
    return records

def diff_bracket(d):
    if d < -10: return "< −10"
    if d <  -5: return "−10 ~ −6"
    if d <   0: return "−5 ~ −1"
    if d ==  0: return "同齡"
    if d <=  5: return "+1 ~ +5"
    if d <= 10: return "+6 ~ +10"
    return "> +10"

DIFF_BRACKETS = ["< −10","−10 ~ −6","−5 ~ −1","同齡","+1 ~ +5","+6 ~ +10","> +10"]

# ─── HTML Table CSS ───────────────────────────────────────────────────────────
TABLE_CSS = (
    "<style>"
    ".ptbl-wrap{max-width:960px;margin:0 auto;border-radius:14px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);}"
    ".ptbl{width:100%;border-collapse:collapse;font-size:17px;}"
    ".ptbl thead tr{background:#1e293b;color:#fff;}"
    ".ptbl thead th{padding:13px 16px;text-align:left;font-weight:700;}"
    ".ptbl thead th.c{text-align:center;}"
    ".ptbl tbody tr:nth-child(odd) td{background:#fff;}"
    ".ptbl tbody tr:nth-child(even) td{background:#f3f6fd;}"
    ".ptbl tbody tr:hover td{background:#dbeafe!important;}"
    ".ptbl td{padding:12px 16px;border-bottom:1px solid #e2e8f0;}"
    ".ptbl td.c{text-align:center;}"
    ".pos{color:#16a34a;font-weight:700;}"
    ".neg{color:#dc2626;font-weight:700;}"
    ".badge-win{background:#dcfce7;color:#166534;border-radius:8px;padding:3px 10px;font-weight:700;}"
    ".badge-lose{background:#fee2e2;color:#991b1b;border-radius:8px;padding:3px 10px;font-weight:700;}"
    "</style>"
)

# ═══════════════════════════════════════════════════════════════════════════════
#  載入資料
# ═══════════════════════════════════════════════════════════════════════════════
settings   = load_settings()
opponents  = load_opponents()
matches    = load_matches()
df_players = load_taiwan_players()

player_opts = (
    [f"{int(r['號碼']):02d} - {r['姓名']}" for _, r in df_players.sort_values("號碼").iterrows()]
    if not df_players.empty else []
)

st.title("🎯 勝負預測系統")

tab_new, tab_pending, tab_stats, tab_opps, tab_cfg = st.tabs(
    ["📝 新增對局","⏳ 待填結果","📊 統計","👥 對手庫","⚙️ 設定"]
)

# ═══════════════════════════════════════════════════════════════════════════════
#  Tab 1：新增對局
# ═══════════════════════════════════════════════════════════════════════════════
with tab_new:
    col_tw, col_opp = st.columns(2, gap="large")
    tw_num, tw_rating, tw_name, tw_by = None, None, None, 0

    with col_tw:
        st.markdown("### 🇹🇼 台灣選手")
        if not player_opts:
            st.error("無法讀取選手資料")
        else:
            tw_sel = st.selectbox(
                "選擇台灣選手", [None] + player_opts,
                format_func=lambda x: "── 請選擇 ──" if x is None else x, key="tw_sel"
            )
            if tw_sel:
                tw_num  = int(tw_sel.split(" - ")[0])
                tw_row  = df_players[df_players["號碼"] == tw_num].iloc[0]
                tw_rating = int(tw_row["等級分"])
                tw_name   = str(tw_row["姓名"])
                tw_by     = int(tw_row.get("birth_year", 0))
                age_str   = f"（{date.today().year - tw_by} 歲）" if tw_by > 0 else ""
                st.metric("目前等級分", tw_rating, age_str if age_str else None)

    opp_data = None
    with col_opp:
        st.markdown("### 🌏 對手")
        opp_mode = st.radio("", ["從對手庫選取","新建對手"], horizontal=True, key="opp_mode")

        if opp_mode == "從對手庫選取":
            if not opponents:
                st.info("對手庫是空的，請切換到「新建對手」")
            else:
                opp_labels = [
                    f"{o['name']} ({o['nationality']})  {o['rating']}{'*' if o['is_estimate'] else ''}"
                    for o in opponents
                ]
                idx = st.selectbox("選擇對手", range(len(opp_labels)),
                                   format_func=lambda i: opp_labels[i], key="opp_sel")
                opp_data = opponents[idx]
                age_info = f"（{date.today().year - opp_data['birth_year']} 歲）" if opp_data.get("birth_year", 0) > 0 else "（年齡未填）"
                est_tag  = "（估計值）" if opp_data["is_estimate"] else ""
                st.metric("對手等級分", f"{opp_data['rating']}{est_tag}", age_info)
        else:
            nc1, nc2 = st.columns(2)
            with nc1:
                new_name = st.text_input("姓名", key="new_name")
                new_nat  = st.selectbox("國籍", ["中國","韓國","日本","其他"], key="new_nat")
                new_by   = st.number_input("出生年份", 1960, 2015, 1995, 1, key="new_by",
                                           help="例：1990。不確定可留預設，之後在對手庫補填")
            with nc2:
                new_rating = st.number_input("等級分", 1000, 5000, 3000, 10, key="new_rating")
                new_est    = st.checkbox("等級分為估計值", key="new_est")

            if st.button("💾 儲存並使用此對手", key="btn_save_opp"):
                if not new_name.strip():
                    st.error("請填入對手姓名")
                else:
                    opp_data = {
                        "id": str(uuid.uuid4()),
                        "name": new_name.strip(),
                        "nationality": new_nat,
                        "rating": int(new_rating),
                        "is_estimate": new_est,
                        "birth_year": int(new_by),
                        "created_at": date.today().isoformat(),
                    }
                    try:
                        append_opponent(opp_data)
                        st.success(f"✅ 已儲存：{opp_data['name']}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"儲存失敗：{e}")

    # ── 勝率顯示 ──
    st.markdown("---")
    if tw_rating and opp_data:
        base_p, adj_p = calc_win_prob(tw_rating, opp_data["rating"], opp_data["is_estimate"], False, settings)
        color, bg = prob_color(adj_p), prob_bg(adj_p)
        pc1, pc2, pc3 = st.columns([1,1,2])
        with pc1:
            st.markdown('<p class="prob-label">📊 Elo 基準勝率</p>', unsafe_allow_html=True)
            st.markdown(f'<div class="prob-box" style="background:#f0f9ff;color:#0369a1">{base_p}%</div>', unsafe_allow_html=True)
        with pc2:
            lbl = "（含估計修正）" if opp_data["is_estimate"] else "（無修正）"
            st.markdown(f'<p class="prob-label">🎯 系統預測 {lbl}</p>', unsafe_allow_html=True)
            st.markdown(f'<div class="prob-box" style="background:{bg};color:{color}">{adj_p}%</div>', unsafe_allow_html=True)
        with pc3:
            diff = tw_rating - opp_data["rating"]
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown(f"分差：**{'+' if diff>=0 else ''}{diff}**（台灣 {tw_rating} vs 對手 {opp_data['rating']}）")
            if opp_data["is_estimate"]:
                st.caption(f"對手等級分為估計值，係數 ×{settings['compression_one']} 壓縮")

        st.markdown("### 📋 對局資料")
        mc1, mc2, mc3 = st.columns(3)
        with mc1: match_date  = st.date_input("日期", value=date.today(), key="match_date")
        with mc2: match_event = st.text_input("賽事名稱", placeholder="例：三星杯預選", key="match_event")
        with mc3: time_ctrl   = st.selectbox("用時制度", ["慢棋","快棋","超快棋"], key="time_ctrl")
        my_pred = st.radio("📌 我的預測", ["台灣選手勝","對手勝"], horizontal=True, key="my_pred")

        if st.button("💾 儲存對局預測", type="primary", key="btn_save_match"):
            if not match_event.strip():
                st.error("請填入賽事名稱")
            else:
                rec = {
                    "id": str(uuid.uuid4()),
                    "date": match_date.isoformat(),
                    "event": match_event.strip(),
                    "taiwan_num": tw_num,
                    "taiwan_name": tw_name,
                    "taiwan_rating": tw_rating,
                    "opp_id": opp_data.get("id",""),
                    "opp_name": opp_data["name"],
                    "opp_nationality": opp_data["nationality"],
                    "opp_rating": opp_data["rating"],
                    "opp_is_estimate": opp_data["is_estimate"],
                    "time_control": time_ctrl,
                    "base_win_prob": base_p,
                    "system_win_prob": adj_p,
                    "my_prediction": my_pred,
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "taiwan_birth_year": tw_by if tw_by > 0 else "",
                    "opp_birth_year": opp_data.get("birth_year") or "",
                }
                try:
                    append_match(rec)
                    st.success(f"✅ {tw_name} vs {opp_data['name']}｜系統預測 {adj_p}%｜我押：{my_pred}")
                except Exception as e:
                    st.error(f"儲存失敗：{e}")
    else:
        st.info("請先選擇台灣選手與對手，系統即會計算勝率")

# ═══════════════════════════════════════════════════════════════════════════════
#  Tab 2：待填結果
# ═══════════════════════════════════════════════════════════════════════════════
with tab_pending:
    pending = [m for m in matches if m["actual_result"] is None]
    if not pending:
        st.success("✅ 所有對局皆已填入結果！")
    else:
        st.markdown(f"#### ⏳ 尚有 **{len(pending)}** 筆待填結果")
        for m in sorted(pending, key=lambda x: x["date"], reverse=True):
            with st.expander(f"📅 {m['date']}　{m['event']}　{m['taiwan_name']} vs {m['opp_name']}"):
                c1, c2, c3 = st.columns(3)
                c1.markdown(f"**台灣：** {m['taiwan_name']} ({m['taiwan_rating']})")
                c2.markdown(f"**對手：** {m['opp_name']} ({m['opp_nationality']}) {m['opp_rating']}{'*' if m['opp_is_estimate'] else ''}")
                c3.markdown(f"**系統預測：** {m['system_win_prob']}%　｜　**我押：** {m['my_prediction']}")
                result_choice = st.radio("實際結果", ["台灣選手勝","對手勝"], horizontal=True, key=f"res_{m['id']}")
                if st.button("✅ 確認填入", key=f"confirm_{m['id']}"):
                    try:
                        fill_match_result(m["id"], result_choice)
                        st.success("已記錄！"); st.rerun()
                    except Exception as e:
                        st.error(f"更新失敗：{e}")

# ═══════════════════════════════════════════════════════════════════════════════
#  Tab 3：統計
# ═══════════════════════════════════════════════════════════════════════════════
with tab_stats:
    completed = [m for m in matches if m["actual_result"]]

    if not completed:
        st.info("尚無已完成對局")
    else:
        df_c = pd.DataFrame(completed)
        df_c["tw_win"]      = (df_c["actual_result"] == "台灣選手勝").astype(int)
        df_c["sys_pred_tw"] = df_c["system_win_prob"] > 50
        df_c["my_pred_tw"]  = df_c["my_prediction"] == "台灣選手勝"
        df_c["sys_correct"] = df_c["sys_pred_tw"] == df_c["tw_win"].astype(bool)
        df_c["my_correct"]  = df_c["my_pred_tw"]  == df_c["tw_win"].astype(bool)
        df_c["brier"]       = ((df_c["system_win_prob"] / 100) - df_c["tw_win"]) ** 2
        df_c["rating_diff"] = df_c["taiwan_rating"] - df_c["opp_rating"]
        n = len(df_c)

        # ── 基本指標 ──
        st.markdown("### 基本指標")
        if n < 5: st.caption(f"⚠️ 目前 {n} 筆，建議 10 筆以上才穩定")
        sc1,sc2,sc3,sc4,sc5 = st.columns(5)
        sc1.metric("總對局數", n)
        sc2.metric("台灣實際勝率", f"{df_c['tw_win'].mean()*100:.1f}%")
        sc3.metric("系統準確率", f"{df_c['sys_correct'].mean()*100:.1f}%")
        sc4.metric("我的準確率", f"{df_c['my_correct'].mean()*100:.1f}%")
        sc5.metric("Brier Score", f"{df_c['brier'].mean():.3f}", help="越低越準")

        # ── 分歧盤 ──
        diverge = df_c[df_c["sys_pred_tw"] != df_c["my_pred_tw"]]
        st.markdown("### 🔑 分歧盤分析")
        if len(diverge) == 0:
            st.info("目前無分歧盤")
        else:
            div_sys = diverge["sys_correct"].mean() * 100
            div_my  = diverge["my_correct"].mean() * 100
            dc1,dc2,dc3 = st.columns(3)
            dc1.metric("分歧盤數", len(diverge))
            dc2.metric("系統命中率", f"{div_sys:.1f}%")
            dc3.metric("我的命中率", f"{div_my:.1f}%")
            if div_sys > div_my + 5: st.info("📈 分歧盤中系統較準")
            elif div_my > div_sys + 5: st.success("🧠 分歧盤中你的直覺較準！")
            else: st.caption("差距不大，繼續累積資料")

        # ── 校準表 ──
        st.markdown("### 校準表")
        bins    = list(range(0,110,10))
        bin_lbl = [f"{b}–{b+10}%" for b in bins[:-1]]
        df_c["prob_bin"] = pd.cut(df_c["system_win_prob"], bins=bins, labels=bin_lbl, include_lowest=True)
        calib = df_c.groupby("prob_bin", observed=True).agg(
            盤數=("tw_win","count"), 預測均值=("system_win_prob","mean"), 實際勝率=("tw_win","mean"),
        ).reset_index().rename(columns={"prob_bin":"勝率區間"})
        c_html = TABLE_CSS + '<div class="ptbl-wrap"><table class="ptbl"><thead><tr>'
        for col in ["勝率區間","盤數","預測均值","實際勝率"]:
            c_html += f'<th{"" if col=="勝率區間" else " class=c"}>{col}</th>'
        c_html += "</tr></thead><tbody>"
        for i,row in calib.iterrows():
            if row["盤數"]==0: continue
            bg = "#fff" if i%2==0 else "#f3f6fd"
            c_html += (f'<tr style="background:{bg}"><td>{row["勝率區間"]}</td>'
                       f'<td class="c">{row["盤數"]}</td>'
                       f'<td class="c">{row["預測均值"]:.1f}%</td>'
                       f'<td class="c"><strong>{row["實際勝率"]*100:.1f}%</strong></td></tr>')
        c_html += "</tbody></table></div>"
        st.html(c_html)

        # ── 分差對照表 ──
        st.markdown("### 分差對照表")
        gap_breaks = [-9999,-300,-150,-75,0,75,150,300,9999]
        gap_labels = ["<−300","−300~−150","−150~−75","−75~0","0~75","75~150","150~300",">300"]
        midpoints  = [-350,-225,-112,-37,37,112,225,350]
        df_c["diff_bin"] = pd.cut(df_c["rating_diff"], bins=gap_breaks, labels=gap_labels, include_lowest=True)
        gap = df_c.groupby("diff_bin", observed=True).agg(
            盤數=("tw_win","count"), 實際勝率=("tw_win","mean"),
        ).reset_index().rename(columns={"diff_bin":"分差"})
        mid_map = {lab: mid for lab,mid in zip(gap_labels,midpoints)}
        g_html = TABLE_CSS + '<div class="ptbl-wrap"><table class="ptbl"><thead><tr>'
        for col in ["分差（台灣－對手）","盤數","實際勝率","Elo 理論值"]:
            g_html += f'<th{"" if "分差" in col else " class=c"}>{col}</th>'
        g_html += "</tr></thead><tbody>"
        for i,row in gap.iterrows():
            if row["盤數"]==0: continue
            bg   = "#fff" if i%2==0 else "#f3f6fd"
            theo = round(1/(1+10**(-mid_map.get(str(row["分差"]),0)/400))*100,1)
            g_html += (f'<tr style="background:{bg}"><td>{row["分差"]}</td>'
                       f'<td class="c">{row["盤數"]}</td>'
                       f'<td class="c"><strong>{row["實際勝率"]*100:.1f}%</strong></td>'
                       f'<td class="c" style="color:#64748b">{theo}%</td></tr>')
        g_html += "</tbody></table></div>"
        st.html(g_html)

        # ══ 年齡分析 ══════════════════════════════════════════════════════════
        st.markdown("### 年齡分析")
        age_records  = build_age_records(completed)
        diff_records = build_diff_records(completed)

        if not age_records:
            st.info(
                "尚無年齡資料。在「對手庫」編輯對手時填入出生年份，"
                "之後新增的對局即會自動記錄，年齡分析表才會啟用。"
            )
        else:
            # ── 表一：絕對年齡表現 ──
            st.markdown("#### 表一：絕對年齡表現")
            st.caption(
                "每場對局拆成兩筆（台灣選手＋對手），按當時年齡分組。"
                "Elo 預測勝率一律使用**未修正原始分**。"
                "偏離值 = 實際勝率 − 預測勝率；負值表示該年齡層實際弱於 Elo 預測。"
            )
            df_ar = pd.DataFrame(age_records)
            df_ar["age_bracket"] = pd.Categorical(df_ar["age_bracket"], categories=AGE_BRACKETS, ordered=True)
            t1 = df_ar.groupby("age_bracket", observed=False).agg(
                count=("actual","count"),
                pred=("predicted","mean"),
                actual=("actual","mean"),
            ).reset_index()
            t1["deviation"] = t1["actual"] - t1["pred"]
            t1 = t1[t1["count"] > 0]

            t1_html = TABLE_CSS + '<div class="ptbl-wrap"><table class="ptbl"><thead><tr>'
            for col in ["年齡層","對局筆數","Elo 預測勝率","實際勝率","偏離值"]:
                t1_html += f'<th{"" if col=="年齡層" else " class=c"}>{col}</th>'
            t1_html += "</tr></thead><tbody>"
            for i, row in t1.iterrows():
                bg  = "#fff" if i%2==0 else "#f3f6fd"
                dev = row["deviation"]
                dev_str = f"{dev*100:+.1f}%"
                dev_cls = "pos" if dev > 0.01 else ("neg" if dev < -0.01 else "")
                t1_html += (
                    f'<tr style="background:{bg}">'
                    f'<td>{row["age_bracket"]}</td>'
                    f'<td class="c">{row["count"]}</td>'
                    f'<td class="c">{row["pred"]*100:.1f}%</td>'
                    f'<td class="c"><strong>{row["actual"]*100:.1f}%</strong></td>'
                    f'<td class="c"><span class="{dev_cls}">{dev_str}</span></td>'
                    "</tr>"
                )
            t1_html += "</tbody></table></div>"
            st.html(t1_html)

            # ── 表二：年齡差分析 ──
            st.markdown("#### 表二：年齡差分析（次要）")
            st.caption(
                "按「台灣選手年齡 − 對手年齡」分組（需雙方都有出生年份）。"
                "正值代表台灣選手較年長。用途：驗證表一修正在實際對局中是否有效。"
            )
            if not diff_records:
                st.info("需要台灣選手和對手雙方都填入出生年份，才能計算年齡差。")
            else:
                df_dr = pd.DataFrame(diff_records)
                df_dr["bracket"] = df_dr["diff"].apply(diff_bracket)
                df_dr["bracket"] = pd.Categorical(df_dr["bracket"], categories=DIFF_BRACKETS, ordered=True)
                t2 = df_dr.groupby("bracket", observed=False).agg(
                    盤數=("tw_win","count"),
                    台灣實際勝率=("tw_win","mean"),
                    Elo預測勝率=("base_p","mean"),
                ).reset_index().rename(columns={"bracket":"年齡差（台灣－對手）"})
                t2 = t2[t2["盤數"] > 0]
                t2["偏離值"] = t2["台灣實際勝率"] - t2["Elo預測勝率"]

                t2_html = TABLE_CSS + '<div class="ptbl-wrap"><table class="ptbl"><thead><tr>'
                for col in ["年齡差（台灣－對手）","盤數","Elo 預測","實際勝率","偏離值"]:
                    t2_html += f'<th{"" if "年齡差" in col else " class=c"}>{col}</th>'
                t2_html += "</tr></thead><tbody>"
                for i, row in t2.iterrows():
                    bg  = "#fff" if i%2==0 else "#f3f6fd"
                    dev = row["偏離值"]
                    dev_str = f"{dev*100:+.1f}%"
                    dev_cls = "pos" if dev > 0.01 else ("neg" if dev < -0.01 else "")
                    t2_html += (
                        f'<tr style="background:{bg}">'
                        f'<td>{row["年齡差（台灣－對手）"]}</td>'
                        f'<td class="c">{row["盤數"]}</td>'
                        f'<td class="c">{row["Elo預測勝率"]*100:.1f}%</td>'
                        f'<td class="c"><strong>{row["台灣實際勝率"]*100:.1f}%</strong></td>'
                        f'<td class="c"><span class="{dev_cls}">{dev_str}</span></td>'
                        "</tr>"
                    )
                t2_html += "</tbody></table></div>"
                st.html(t2_html)

        # ── 近期明細 ──
        st.markdown("### 近期對局明細（最近 20 筆）")
        recent = sorted(completed, key=lambda x: x["date"], reverse=True)[:20]
        r_html = TABLE_CSS + '<div class="ptbl-wrap"><table class="ptbl"><thead><tr>'
        for col in ["日期","賽事","台灣選手","對手","系統勝率","我押","結果"]:
            r_html += f'<th{"" if col in ["日期","賽事","台灣選手","對手"] else " class=c"}>{col}</th>'
        r_html += "</tr></thead><tbody>"
        for i,m in enumerate(recent):
            bg     = "#fff" if i%2==0 else "#f3f6fd"
            res    = m["actual_result"]
            sys_ok = (m["system_win_prob"]>50) == (res=="台灣選手勝")
            my_ok  = (m["my_prediction"]=="台灣選手勝") == (res=="台灣選手勝")
            badge  = '<span class="badge-win">台灣勝</span>' if res=="台灣選手勝" else '<span class="badge-lose">對手勝</span>'
            r_html += (f'<tr style="background:{bg}"><td>{m["date"]}</td><td>{m["event"]}</td>'
                       f'<td>{m["taiwan_name"]}</td><td>{m["opp_name"]} ({m["opp_nationality"]})</td>'
                       f'<td class="c">{m["system_win_prob"]}% {"✅" if sys_ok else "❌"}</td>'
                       f'<td class="c">{"台灣" if m["my_prediction"]=="台灣選手勝" else "對手"} {"✅" if my_ok else "❌"}</td>'
                       f'<td class="c">{badge}</td></tr>')
        r_html += "</tbody></table></div>"
        st.html(r_html)

        st.markdown("---")
        export_cols = ["date","event","taiwan_name","taiwan_rating","opp_name","opp_nationality",
                       "opp_rating","opp_is_estimate","time_control","base_win_prob",
                       "system_win_prob","my_prediction","actual_result"]
        csv_bytes = df_c[export_cols].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("📥 匯出全部對局 CSV", data=csv_bytes,
                           file_name=f"match_predictions_{date.today()}.csv", mime="text/csv")

# ═══════════════════════════════════════════════════════════════════════════════
#  Tab 4：對手庫
# ═══════════════════════════════════════════════════════════════════════════════
with tab_opps:
    st.markdown("### 👥 對手庫管理")
    if not opponents:
        st.info("對手庫是空的，請在「新增對局」頁面建立第一位對手")
    else:
        for i, o in enumerate(opponents):
            by      = int(o.get("birth_year", 0))
            age_tag = f"（{date.today().year - by} 歲）" if by > 0 else "（年齡未填）"
            est_tag = "（估計值）" if o["is_estimate"] else ""
            with st.expander(f"{o['name']}　{o['nationality']}　{o['rating']}{est_tag}　{age_tag}"):
                ec1, ec2, ec3, ec4 = st.columns([1,1,1,1])
                with ec1:
                    new_r = st.number_input("等級分", value=int(o["rating"]), step=10, key=f"or_{i}")
                with ec2:
                    nat_opts = ["中國","韓國","日本","其他"]
                    new_nat = st.selectbox("國籍", nat_opts,
                                           index=nat_opts.index(o.get("nationality","其他")),
                                           key=f"onat_{i}")
                with ec3:
                    new_by2 = st.number_input("出生年份", 1960, 2015,
                                              value=int(o.get("birth_year") or 1990),
                                              step=1, key=f"oby_{i}")
                with ec4:
                    new_e = st.checkbox("估計值", value=o["is_estimate"], key=f"oe_{i}")
                if st.button("💾 更新", key=f"oupd_{i}"):
                    try:
                        update_opponent_row(o["id"], int(new_r), new_nat, new_e, int(new_by2))
                        st.success("✅ 已更新"); st.rerun()
                    except Exception as e:
                        st.error(f"更新失敗：{e}")
                if st.button("🗑 刪除此對手", key=f"odel_{i}"):
                    try:
                        delete_opponent_row(o["id"]); st.rerun()
                    except Exception as e:
                        st.error(f"刪除失敗：{e}")

# ═══════════════════════════════════════════════════════════════════════════════
#  Tab 5：設定
# ═══════════════════════════════════════════════════════════════════════════════
with tab_cfg:
    st.markdown("### ⚙️ 可調參數")
    st.caption("對手等級分為估計值時，系統將勝率向 50% 壓縮。1.0=不壓縮；0.5=折半壓縮。")
    s1, s2 = st.columns(2)
    with s1:
        c1_val = st.slider("單方估計值壓縮係數", 0.50, 1.00,
                           float(settings.get("compression_one", 0.85)), 0.05,
                           help="對手是估計值時。預設 0.85")
    with s2:
        c2_val = st.slider("雙方估計值壓縮係數", 0.50, 1.00,
                           float(settings.get("compression_both", 0.75)), 0.05,
                           help="雙方都是估計值時。預設 0.75")
    st.markdown("#### 預覽（基準勝率 70%）")
    ex1 = round(50 + (70-50)*c1_val, 1)
    ex2 = round(50 + (70-50)*c2_val, 1)
    pc1, pc2, pc3 = st.columns(3)
    pc1.metric("原始 Elo", "70.0%")
    pc2.metric(f"單方估計（×{c1_val:.2f}）", f"{ex1}%", f"−{70-ex1:.1f}%")
    pc3.metric(f"雙方估計（×{c2_val:.2f}）", f"{ex2}%", f"−{70-ex2:.1f}%")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 儲存設定", type="primary", key="btn_save_cfg"):
        try:
            save_settings(c1_val, c2_val)
            st.success("✅ 設定已儲存至 Google Sheets")
        except Exception as e:
            st.error(f"儲存失敗：{e}")
