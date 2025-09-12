
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="成績預警儀表板", layout="wide")

st.title("成績預警儀表板（兩科按週）")

with st.sidebar:
    st.header("檔案與門檻與篩選")
    uploaded = st.file_uploader("上傳成績檔（Excel, 需包含 'score' 工作表，欄位：ID, Biochem, MolBio, Week）", type=["xlsx"])
    red_th = st.number_input("紅色門檻（未達）", min_value=0, max_value=100, value=40, step=1)
    yellow_th = st.number_input("黃色上限（含）", min_value=0, max_value=100, value=60, step=1)
    win_len = st.number_input("警示連續週數（任一科）", min_value=2, max_value=18, value=3, step=1)

    st.markdown("---")
    cohort_opt = st.selectbox("學籍（依 ID 是否以 413 開頭）", ["全部", "應屆", "非應屆"], index=0)
    dept_opt = st.selectbox("系所（依 ID 第4-5碼）", ["全部", "醫學系(01)", "牙醫學系(02)", "藥學系(03)"], index=0)

    if yellow_th < red_th:
        st.warning("⚠️ 黃色上限應大於等於紅色門檻（目前黃={yellow} < 紅={red}）。".format(yellow=yellow_th, red=red_th))

@st.cache_data
def load_score_df(file):
    if file is None:
        try:
            return pd.read_excel("2025-biochem_mobio_score.xlsx", sheet_name="score")
        except Exception as e:
            st.error("未上傳檔案，且找不到預設檔")
            st.stop()
    else:
        return pd.read_excel(file, sheet_name="score")

df = load_score_df(uploaded)

required_cols = {"ID", "Biochem", "MolBio", "Week"}
if not required_cols.issubset(df.columns):
    st.error(f"缺少必要欄位：{required_cols - set(df.columns)}，請確認 'score' 工作表欄位為 ID, Biochem, MolBio, Week。")
    st.stop()

# 清理
df["Week"] = pd.to_numeric(df["Week"], errors="coerce").astype("Int64")
df = df.dropna(subset=["Week"]).copy()
df["Week"] = df["Week"].astype(int)
df = df.sort_values(["ID", "Week"])

# ====== 衍生欄位：學籍與系所 ======
def to_str_id(x):
    # 將 ID 正規化為字串（保留前導 0，如有）
    try:
        s = str(x)
    except Exception:
        s = ""
    return s

df["ID_str"] = df["ID"].apply(to_str_id)
df["應屆"] = df["ID_str"].str.startswith("413")

def dept_from_id(s: str) -> str:
    s = to_str_id(s)
    # 取第 4-5 碼（1-based），對應 0-based 的 3:5
    code = s[3:5] if len(s) >= 5 else ""
    if code == "01":
        return "醫學系(01)"
    elif code == "02":
        return "牙醫學系(02)"
    elif code == "03":
        return "藥學系(03)"
    else:
        return "未知"

df["系所"] = df["ID_str"].apply(dept_from_id)

# 篩選學籍
if cohort_opt == "應屆":
    df = df[df["應屆"] == True]
elif cohort_opt == "非應屆":
    df = df[df["應屆"] == False]

# 篩選系所
if dept_opt != "全部":
    df = df[df["系所"] == dept_opt]

# 固定 18 週
WEEKS_FULL = list(range(1, 19))

# 透視表：整數顯示、缺值空白
def make_pivot(subject_col: str) -> pd.DataFrame:
    pivot = df.pivot(index="ID", columns="Week", values=subject_col)
    # 若過濾後無資料，提供空表骨架
    if pivot.empty:
        pivot = pd.DataFrame(index=[], columns=WEEKS_FULL)
    pivot = pivot.sort_index()
    pivot = pivot.reindex(columns=WEEKS_FULL)
    pivot = pivot.applymap(lambda x: int(x) if pd.notna(x) else "")
    return pivot

bio_pivot = make_pivot("Biochem")
mol_pivot = make_pivot("MolBio")

# 著色（空白不著色）
def color_cell(v):
    try:
        if v == "":
            return ""
        x = float(v)
    except Exception:
        return ""
    if x < red_th:
        return "background-color: #f8d7da;"
    elif x <= yellow_th:
        return "background-color: #fff3cd;"
    else:
        return "background-color: #d4edda;"

st.subheader("生物化學（Biochem）")
st.dataframe(bio_pivot.style.applymap(color_cell), use_container_width=True)

st.subheader("分子生物學（MolBio）")
st.dataframe(mol_pivot.style.applymap(color_cell), use_container_width=True)

# 任一科連續 win_len 週皆低於紅線（使用篩選後的 df）
def consecutive_any_subject_red(df_score: pd.DataFrame, red_threshold: float, window_len: int):
    out_rows = []
    if df_score.empty:
        return pd.DataFrame(columns=["ID","Weeks","Biochem_scores","MolBio_scores","科目","連續週數","學籍","系所"])
    for sid, g in df_score.groupby("ID"):
        g = g.set_index("Week").reindex(WEEKS_FULL)
        for i in range(len(WEEKS_FULL) - window_len + 1):
            win_weeks = WEEKS_FULL[i:i+window_len]
            sub = g.loc[win_weeks, ["Biochem", "MolBio"]]
            if sub.isna().any().any():
                continue
            cond_bio = (sub["Biochem"] < red_threshold).all()
            cond_mol = (sub["MolBio"] < red_threshold).all()
            if cond_bio or cond_mol:
                # 從原始 df 取學籍/系所標籤
                sid_meta = df_score[df_score["ID"] == sid].iloc[0]
                out_rows.append({
                    "ID": sid,
                    "Weeks": f"{win_weeks[0]}–{win_weeks[-1]}",
                    "Biochem_scores": tuple(int(x) for x in sub["Biochem"].values),
                    "MolBio_scores": tuple(int(x) for x in sub["MolBio"].values),
                    "科目": "Biochem" if cond_bio else "MolBio",
                    "連續週數": window_len,
                    "學籍": "應屆" if sid_meta["應屆"] else "非應屆",
                    "系所": sid_meta["系所"],
                })
    if not out_rows:
        return pd.DataFrame(columns=["ID","Weeks","Biochem_scores","MolBio_scores","科目","連續週數","學籍","系所"])
    df_out = pd.DataFrame(out_rows)
    df_out = df_out.drop_duplicates(subset=["ID","Weeks","科目","Biochem_scores","MolBio_scores","連續週數"])
    return df_out

alert_df = consecutive_any_subject_red(df, red_th, int(win_len))

st.subheader(f"⚠️ 任一科連續 {int(win_len)} 週紅色名單（套用目前篩選）")
if alert_df.empty:
    st.success(f"目前沒有符合『任一科連續 {int(win_len)} 週皆低於紅色門檻』的學生。")
else:
    show = alert_df.copy()
    show["Biochem_scores"] = show["Biochem_scores"].apply(lambda xs: "、".join(map(str, xs)))
    show["MolBio_scores"] = show["MolBio_scores"].apply(lambda xs: "、".join(map(str, xs)))
    st.dataframe(show, use_container_width=True)

# 匯出
with st.expander("⬇️ 下載目前結果（Excel）"):
    def to_excel_bytes():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            def unblank(df_in: pd.DataFrame) -> pd.DataFrame:
                return df_in.replace("", pd.NA)
            unblank(bio_pivot).to_excel(writer, sheet_name="Biochem_Pivot")
            unblank(mol_pivot).to_excel(writer, sheet_name="MolBio_Pivot")
            (alert_df if not alert_df.empty else pd.DataFrame(columns=["ID","Weeks","Biochem_scores","MolBio_scores","科目","連續週數","學籍","系所"]))\
                .to_excel(writer, sheet_name=f"Red_AnySubject_{int(win_len)}w", index=False)
        output.seek(0)
        return output

    st.download_button(
        label="下載 Excel",
        data=to_excel_bytes(),
        file_name="score_dashboard_outputs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("說明：固定顯示 18 週；空白代表未有成績。紅<紅色門檻；黃=紅色門檻~黃色上限；綠>黃色上限。警示邏輯：任一科連續 N 週（可調）皆低於紅線即列出。學籍依 ID 是否以 413 開頭；系所依 ID 第4-5碼（01醫、02牙、03藥）。")
