import streamlit as st
import pandas as pd
import json
from pathlib import Path

st.set_page_config(page_title="Programmatic Assessment - Early Warning MVP v1.1", layout="wide")

DATA_DIR = Path("./data")
DATA_DIR.mkdir(exist_ok=True)
STUDENTS_CSV = DATA_DIR / "students.csv"
PROGRAMS_CSV = DATA_DIR / "programs.csv"
SCORES_MASTER_CSV = DATA_DIR / "scores_master.csv"
THRESHOLDS_JSON = DATA_DIR / "thresholds.json"
FEEDBACKS_CSV = DATA_DIR / "feedbacks.csv"
ANON_MAP_CSV = DATA_DIR / "anon_map.csv"

SUBJECTS = ["BIOCHEM","MOLBIO"]
ASSESS_TYPES = ["WEEKLY","MIDTERM","FINAL"]

def load_csv(path: Path, fallback_df: pd.DataFrame) -> pd.DataFrame:
    if path.exists():
        try:
            return pd.read_csv(path)
        except Exception as e:
            st.warning(f"⚠️ 無法讀取 {path.name}：{e}，改用暫存資料。")
            return fallback_df.copy()
    return fallback_df.copy()

def save_csv(df: pd.DataFrame, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)

def load_thresholds():
    if THRESHOLDS_JSON.exists():
        try:
            return json.loads(THRESHOLDS_JSON.read_text(encoding="utf-8"))
        except Exception as e:
            st.warning(f"⚠️ 無法讀取 thresholds：{e}，改用預設。")
    # default, including new advanced thresholds
    return {
        "global":{"red_max":40,"yellow_max":70},
        "by_program":{},
        "advanced":{"mid_low":60,"final_low":60,"cross_gap":20}
    }

def save_thresholds(t: dict):
    THRESHOLDS_JSON.write_text(json.dumps(t, indent=2), encoding="utf-8")

@st.cache_data
def init_defaults():
    # empty shells
    students = pd.DataFrame(columns=["student_id","name","program","enrolled_year"])
    programs = pd.DataFrame(columns=["program","name"])
    scores = pd.DataFrame(columns=["student_id","week","subject","type","raw_score"])
    feedbacks = pd.DataFrame(columns=["student_id","assessment_key","note","author"])
    return students, programs, scores, feedbacks

def color_light(score, red_max, yellow_max):
    if pd.isna(score):
        return "GRAY"
    if score <= red_max:
        return "RED"
    if score <= yellow_max:
        return "YELLOW"
    return "GREEN"

def apply_thresholds(scores_df, thresholds, students_df):
    # merge program to each score
    merged = scores_df.merge(students_df[["student_id","program"]], on="student_id", how="left")
    def get_pair(row):
        prog = row.get("program", None)
        if prog and prog in thresholds.get("by_program", {}):
            r = thresholds["by_program"][prog]["red_max"]
            y = thresholds["by_program"][prog]["yellow_max"]
        else:
            r = thresholds["global"]["red_max"]
            y = thresholds["global"]["yellow_max"]
        return pd.Series({"red_max": r, "yellow_max": y})
    if len(merged):
        merged = pd.concat([merged, merged.apply(get_pair, axis=1)], axis=1)
        merged["light"] = merged.apply(lambda r: color_light(r["raw_score"], r["red_max"], r["yellow_max"]), axis=1)
    else:
        merged["red_max"] = []
        merged["yellow_max"] = []
        merged["light"] = []
    # add an assessment key for feedback tying
    if len(merged):
        merged["assessment_key"] = merged.apply(lambda r: f"{r['week']:02d}-{r['subject']}-{r['type']}", axis=1)
    return merged

def weekly_stack(merged):
    if len(merged)==0:
        return pd.DataFrame()
    ct = merged.pivot_table(index="week", columns="light", values="raw_score", aggfunc="count").fillna(0).reset_index()
    # ensure column order
    for col in ["RED","YELLOW","GREEN"]:
        if col not in ct.columns:
            ct[col] = 0
    return ct.sort_values("week")

def mid_final_scatter(merged):
    if len(merged)==0:
        return pd.DataFrame()
    # 平均兩科
    mid = merged[merged["type"]=="MIDTERM"].groupby("student_id")["raw_score"].mean()
    fin = merged[merged["type"]=="FINAL"].groupby("student_id")["raw_score"].mean()
    scat = pd.DataFrame({"midterm": mid, "final": fin}).dropna().reset_index()
    return scat

def risk_snapshot_basic(merged):
    if len(merged)==0:
        return pd.DataFrame()
    # 連續紅 >=2 或連續黃 >=3（週考）
    wk = merged[merged["type"]=="WEEKLY"].copy()
    if len(wk)==0:
        return pd.DataFrame(columns=["student_id","risk_reason"])
    wk = wk.sort_values(["student_id","week"])
    def risk_for(student_df):
        s = student_df["light"].tolist()
        cons_red = cons_yellow = max_red = max_yellow = 0
        for val in s:
            if val=="RED":
                cons_red += 1; cons_yellow = 0
            elif val=="YELLOW":
                cons_yellow += 1; cons_red = 0
            else:
                cons_red = cons_yellow = 0
            max_red = max(max_red, cons_red)
            max_yellow = max(max_yellow, cons_yellow)
        reasons = []
        if max_red >= 2: reasons.append("連續≥2 週紅燈")
        if max_yellow >= 3: reasons.append("連續≥3 週黃燈")
        return "; ".join(reasons)
    risks = wk.groupby("student_id").apply(risk_for).reset_index(name="risk_reason")
    risks = risks[risks["risk_reason"]!=""]
    return risks

def risk_snapshot_advanced(merged, thresholds):
    """落差偵測：
       1) 週考平均 >= yellow_max 但 期中/期末 < mid_low/final_low
       2) 跨科差距：兩科週考平均差 >= cross_gap
    """
    if len(merged)==0:
        return pd.DataFrame(columns=["student_id","adv_reason"])
    g_yellow = thresholds["global"]["yellow_max"]
    mid_low = thresholds.get("advanced",{}).get("mid_low",60)
    final_low = thresholds.get("advanced",{}).get("final_low",60)
    cross_gap = thresholds.get("advanced",{}).get("cross_gap",20)
    # per-student weekly mean (兩科合併)
    wk = merged[merged["type"]=="WEEKLY"]
    weekly_mean = wk.groupby("student_id")["raw_score"].mean().rename("weekly_mean")
    # mid/final mean (兩科)
    mid = merged[merged["type"]=="MIDTERM"].groupby("student_id")["raw_score"].mean().rename("mid_mean")
    fin = merged[merged["type"]=="FINAL"].groupby("student_id")["raw_score"].mean().rename("final_mean")
    # cross-subject gap on weekly means
    by_subj = wk.groupby(["student_id","subject"])["raw_score"].mean().unstack(fill_value=float("nan"))
    if "BIOCHEM" not in by_subj.columns: by_subj["BIOCHEM"] = float("nan")
    if "MOLBIO" not in by_subj.columns: by_subj["MOLBIO"] = float("nan")
    by_subj["cross_gap"] = (by_subj["BIOCHEM"] - by_subj["MOLBIO"]).abs()

    df = pd.DataFrame(weekly_mean).join([mid, fin, by_subj["cross_gap"]])
    df = df.fillna({"mid_mean": float("nan"), "final_mean": float("nan"), "cross_gap": float("nan")})

    def reason(row):
        rs = []
        if pd.notna(row["weekly_mean"]) and pd.notna(row["mid_mean"]):
            if row["weekly_mean"] >= g_yellow and row["mid_mean"] < mid_low:
                rs.append(f"週考高分(≥{g_yellow})但期中偏低(<{mid_low})")
        if pd.notna(row["weekly_mean"]) and pd.notna(row["final_mean"]):
            if row["weekly_mean"] >= g_yellow and row["final_mean"] < final_low:
                rs.append(f"週考高分(≥{g_yellow})但期末偏低(<{final_low})")
        if pd.notna(row["cross_gap"]) and row["cross_gap"] >= cross_gap:
            rs.append(f"跨科落差≥{cross_gap}")
        return "; ".join(rs)
    df["adv_reason"] = df.apply(reason, axis=1)
    out = df[ df["adv_reason"]!="" ][["adv_reason"]].reset_index()
    return out

def ensure_feedbacks_csv():
    if not FEEDBACKS_CSV.exists():
        pd.DataFrame(columns=["student_id","assessment_key","note","author"]).to_csv(FEEDBACKS_CSV, index=False)

def load_feedbacks():
    ensure_feedbacks_csv()
    try:
        return pd.read_csv(FEEDBACKS_CSV)
    except Exception:
        return pd.DataFrame(columns=["student_id","assessment_key","note","author"])

def save_feedback(student_id, assessment_key, note, author):
    fb = load_feedbacks()
    fb.loc[len(fb)] = {"student_id":student_id, "assessment_key":assessment_key, "note":note, "author":author}
    save_csv(fb, FEEDBACKS_CSV)

def get_assessment_keys(merged):
    if len(merged)==0:
        return []
    keys = merged[["assessment_key","week","subject","type"]].drop_duplicates().sort_values(["week","subject","type"])
    keys["label"] = keys.apply(lambda r: f"W{int(r['week'])} {r['subject']} {r['type']} ({r['assessment_key']})", axis=1)
    return keys

def generate_anon_map(students_df):
    # create or extend anon_map.csv: student_id -> S0001...
    if ANON_MAP_CSV.exists():
        amap = pd.read_csv(ANON_MAP_CSV)
    else:
        amap = pd.DataFrame(columns=["student_id","anon_id"])
    existing = set(amap["student_id"]) if len(amap) else set()
    next_idx = len(amap) + 1
    rows = []
    for sid in students_df["student_id"].astype(str):
        if sid not in existing:
            rows.append({"student_id":sid, "anon_id": f"S{next_idx:04d}"})
            next_idx += 1
    if rows:
        amap = pd.concat([amap, pd.DataFrame(rows)], ignore_index=True)
        save_csv(amap, ANON_MAP_CSV)
    return amap

def anonymize_view(df, students_df):
    if len(df)==0:
        return pd.DataFrame()
    amap = generate_anon_map(students_df)
    out = df.merge(amap, on="student_id", how="left")
    cols = ["anon_id","week","subject","type","raw_score","light"]
    if "program" in out.columns:
        cols.insert(1, "program")
    return out[cols].sort_values(["anon_id","week","subject","type"])

st.title("學生成績預警機制（Programmatic Assessment）— MVP v1.1")
st.caption("18 週課程；兩科（Biochem/MolBio）週考（1–16 週），期中（第 9 週），期末（第 18 週）｜新增：落差偵測、敘述性回饋、匿名化匯出")

with st.sidebar:
    st.header("資料初始化 / 載入")
    st.write("請先匯入 **students.csv** 與每週 **scores CSV**；範例模板可在系統提供的下載連結取得。")

    # Upload students.csv
    up_students = st.file_uploader("上傳 students.csv", type=["csv"], key="stu_up")
    if up_students is not None:
        df = pd.read_csv(up_students)
        save_csv(df, STUDENTS_CSV)
        st.success(f"students.csv 已更新（{len(df)} 筆）")

    # Upload programs.csv（可選）
    up_programs = st.file_uploader("上傳 programs.csv（選填）", type=["csv"], key="prog_up")
    if up_programs is not None:
        dfp = pd.read_csv(up_programs)
        save_csv(dfp, PROGRAMS_CSV)
        st.success(f"programs.csv 已更新（{len(dfp)} 筆）")

    # Upload weekly scores (merge into master)
    st.divider()
    st.subheader("上傳本週成績（可多檔）")
    uploaded_scores = st.file_uploader("scores CSV（student_id, week, subject, type, raw_score）", type=["csv"], accept_multiple_files=True, key="scores_up")
    if uploaded_scores:
        new_rows = []
        for f in uploaded_scores:
            try:
                df = pd.read_csv(f)
                new_rows.append(df)
            except Exception as e:
                st.error(f"{f.name} 讀取失敗：{e}")
        if new_rows:
            new_scores = pd.concat(new_rows, ignore_index=True)
            # standardize columns
            new_scores.columns = [c.lower() for c in new_scores.columns]
            required = {"student_id","week","subject","type","raw_score"}
            if required - set(new_scores.columns):
                st.error("上傳 CSV 欄位需包含：student_id, week, subject, type, raw_score")
            else:
                new_scores["subject"] = new_scores["subject"].str.upper().str.strip()
                new_scores["type"] = new_scores["type"].str.upper().str.strip()
                new_scores["week"] = new_scores["week"].astype(int)
                new_scores["raw_score"] = pd.to_numeric(new_scores["raw_score"], errors="coerce")
                ok = new_scores["subject"].isin(SUBJECTS) & new_scores["type"].isin(ASSESS_TYPES) & new_scores["week"].between(1,18)
                if (~ok).any():
                    st.warning(f"有 {len(new_scores[~ok])} 筆資料不合法（已忽略）。")
                new_scores = new_scores[ok]
                # upsert
                if SCORES_MASTER_CSV.exists():
                    master = pd.read_csv(SCORES_MASTER_CSV)
                else:
                    master = pd.DataFrame(columns=["student_id","week","subject","type","raw_score"])
                new_scores = new_scores.drop_duplicates(subset=["student_id","week","subject","type"], keep="last")
                if len(master):
                    key_cols = ["student_id","week","subject","type"]
                    master = master[~master.set_index(key_cols).index.isin(new_scores.set_index(key_cols).index)]
                    master = pd.concat([master, new_scores], ignore_index=True)
                else:
                    master = new_scores.copy()
                save_csv(master, SCORES_MASTER_CSV)
                st.success(f"已合併寫入 master（目前共 {len(master)} 筆）")

    st.divider()
    st.header("燈號門檻")
    t = load_thresholds()
    # if old file, ensure advanced keys exist
    if "advanced" not in t: t["advanced"] = {"mid_low":60,"final_low":60,"cross_gap":20}
    red = st.number_input("全域紅燈上限（含）", min_value=0, max_value=100, value=int(t["global"]["red_max"]))
    yellow = st.number_input("全域黃燈上限（含）", min_value=0, max_value=100, value=int(t["global"]["yellow_max"]))
    mid_low = st.number_input("期中低分警戒（平均兩科）", min_value=0, max_value=100, value=int(t["advanced"]["mid_low"]))
    final_low = st.number_input("期末低分警戒（平均兩科）", min_value=0, max_value=100, value=int(t["advanced"]["final_low"]))
    cross_gap = st.number_input("跨科差距警戒（週考兩科平均分差）", min_value=0, max_value=100, value=int(t["advanced"]["cross_gap"]))
    t["global"]["red_max"] = red
    t["global"]["yellow_max"] = yellow
    t["advanced"]["mid_low"] = mid_low
    t["advanced"]["final_low"] = final_low
    t["advanced"]["cross_gap"] = cross_gap

    # program-specific overrides
    st.caption("（選填）依系所覆寫門檻")
    if PROGRAMS_CSV.exists():
        pg = pd.read_csv(PROGRAMS_CSV)
        for prog in sorted(pg["program"].astype(str).unique().tolist()):
            with st.expander(f"門檻覆寫：{prog}"):
                cur = t["by_program"].get(prog, {"red_max": red, "yellow_max": yellow})
                r = st.number_input(f"{prog} 紅燈上限（含）", 0, 100, value=int(cur.get("red_max", red)), key=f"r_{prog}")
                y = st.number_input(f"{prog} 黃燈上限（含）", 0, 100, value=int(cur.get("yellow_max", yellow)), key=f"y_{prog}")
                t["by_program"][prog] = {"red_max": r, "yellow_max": y}
    if st.button("儲存門檻設定"):
        save_thresholds(t)
        st.success("門檻已儲存")

# Main data load
students, programs, scores, feedbacks = init_defaults()
students = load_csv(STUDENTS_CSV, students)
programs = load_csv(PROGRAMS_CSV, programs)
if SCORES_MASTER_CSV.exists():
    scores = pd.read_csv(SCORES_MASTER_CSV)
feedbacks = load_feedbacks()
thresholds = load_thresholds()

# ---- Filters ----
st.subheader("篩選條件")
c1, c2, c3, c4 = st.columns(4)
with c1:
    prog_filter = st.multiselect("系所", sorted(students["program"].dropna().astype(str).unique().tolist()))
with c2:
    subj_filter = st.multiselect("科目", SUBJECTS)
with c3:
    week_filter = st.multiselect("週次", sorted(scores["week"].dropna().astype(int).unique().tolist()))
with c4:
    light_filter = st.multiselect("燈號", ["RED","YELLOW","GREEN"])

# Merge + compute light
merged = apply_thresholds(scores, thresholds, students)

# Apply filters
if len(merged):
    if prog_filter:
        merged = merged[merged["program"].isin(prog_filter)]
    if subj_filter:
        merged = merged[merged["subject"].isin(subj_filter)]
    if week_filter:
        merged = merged[merged["week"].isin(week_filter)]
    if light_filter:
        merged = merged[merged["light"].isin(light_filter)]

# ---- Charts ----
st.divider()
cA, cB = st.columns([1,1])
with cA:
    st.markdown("### 每週燈號比例（RED/YELLOW/GREEN）")
    stack = weekly_stack(merged)
    if len(stack):
        st.bar_chart(stack.set_index("week")[["RED","YELLOW","GREEN"]])
    else:
        st.info("目前沒有可視化資料（尚未上傳或篩選過嚴）。")

with cB:
    st.markdown("### 期中 vs 期末（平均兩科）")
    scat = mid_final_scatter(merged)
    if len(scat):
        st.scatter_chart(scat, x="midterm", y="final")
    else:
        st.info("尚無期中/期末資料可供比對。")

# ---- Risk snapshot ----
st.divider()
st.markdown("### 風險快照（連續紅/黃）")
risk_basic = risk_snapshot_basic(merged)
if len(risk_basic):
    risk_basic = risk_basic.merge(students[["student_id","name","program"]], on="student_id", how="left")
    st.dataframe(risk_basic[["student_id","name","program","risk_reason"]].sort_values(["program","student_id"]))
else:
    st.info("目前沒有連續紅/黃的風險個案。")

st.markdown("### 落差偵測（週考 vs 期中/期末、跨科差距）")
risk_adv = risk_snapshot_advanced(merged, thresholds)
if len(risk_adv):
    risk_adv = risk_adv.merge(students[["student_id","name","program"]], on="student_id", how="left")
    st.dataframe(risk_adv[["student_id","name","program","adv_reason"]].sort_values(["program","student_id"]))
else:
    st.info("尚無落差偵測個案。")

# ---- Feedback (敘述性回饋) ----
st.divider()
st.markdown("## 敘述性回饋")

if len(students)==0 or len(merged)==0:
    st.info("請先載入學生名單與成績，才能新增回饋。")
else:
    cfb1, cfb2 = st.columns([1,2])
    with cfb1:
        sid = st.selectbox("選擇學生", students["student_id"].astype(str).tolist())
        keys = get_assessment_keys(merged[merged["student_id"]==sid])
        key_options = keys["label"].tolist() if len(keys) else []
        label = st.selectbox("選擇評量", key_options)
        note = st.text_area("回饋內容（可含學習建議）", height=120, placeholder="例：建議補強脂質代謝章節；週內安排30分鐘題庫練習。")
        author = st.text_input("回饋撰寫者（顯示名）", value="Teacher")
        if st.button("新增回饋", type="primary", use_container_width=True, disabled=(not label or not note.strip())):
            # parse assessment_key from label tail "(xx)"
            if label and "(" in label and label.endswith(")"):
                akey = label.split("(")[-1][:-1]
            else:
                akey = ""
            save_feedback(sid, akey, note.strip(), author.strip() or "Teacher")
            st.success("已新增回饋。")
    with cfb2:
        st.markdown("### 該生回饋列表")
        fb = load_feedbacks()
        if len(fb):
            fb_view = fb[fb["student_id"]==sid].merge(
                merged[["student_id","assessment_key","week","subject","type"]].drop_duplicates(),
                on=["student_id","assessment_key"], how="left"
            ).sort_values(["week","subject","type"])
            if len(fb_view):
                st.dataframe(fb_view[["week","subject","type","note","author"]])
            else:
                st.info("此學生暫無回饋。")
        else:
            st.info("目前尚無任何回饋紀錄。")

# ---- Raw table & downloads ----
st.divider()
st.markdown("### 明細（套用門檻後）")
if len(merged):
    st.dataframe(merged.sort_values(["week","program","subject","student_id"]))
    st.download_button("下載目前視圖（CSV）", merged.to_csv(index=False).encode("utf-8"),
                       file_name="filtered_scores.csv", mime="text/csv")
    # anonymized export
    st.markdown("#### 匿名化匯出（研究/評鑑）")
    anon = anonymize_view(merged, students)
    if len(anon):
        st.download_button("下載匿名化視圖（CSV）", anon.to_csv(index=False).encode("utf-8"),
                           file_name="filtered_scores_anonymized.csv", mime="text/csv")
    else:
        st.info("沒有可匯出的匿名資料。")
else:
    st.info("沒有明細可顯示。")

st.caption("提示：請將 data/ 資料夾與此 app.py 放在同一層；先放 students.csv，再逐週上傳 scores CSV。")
