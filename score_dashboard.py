import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

st.set_page_config(page_title="成績預警儀表板", layout="wide")
st.title("成績預警儀表板")

with st.sidebar:
    st.header("篩選控制面板")

    # 學籍（依 ID 前三碼判斷：<413 重修, =413 應屆, >413 先修）
    cohort_opts = st.pills(
        "學籍（依 ID 開頭判斷）",
        options=["應屆", "重修", "先修"],
        selection_mode="multi"
    )

    # 系所（純名稱，不帶代碼）
    dept_opts = st.pills(
        "系所（依 ID 第4-5碼）",
        options=["醫學系", "牙醫學系", "藥學系"],
        selection_mode="multi"
    )

    st.markdown("---")
    red_th = st.number_input("紅色門檻（含）", min_value=0, max_value=100, value=40, step=1)
    yellow_th = st.number_input("黃色上限（含）", min_value=0, max_value=100, value=60, step=1)
    if yellow_th < red_th:
        st.warning("⚠️ 黃色上限應大於等於紅色門檻（目前黃={yellow} < 紅={red}）。".format(yellow=yellow_th, red=red_th))

    # === 新：AND 規則 ===
    st.markdown("#### 預警條件（AND 規則）")
    min_red = st.number_input(
        "紅燈最低數量（≥）", min_value=0, max_value=18, value=2, step=1,
        help="在滑動視窗內，紅燈（≤紅線）至少要達到這個數量"
    )
    min_total = st.number_input(
        "紅+黃合計最低數量（≥）", min_value=1, max_value=18, value=4, step=1,
        help="在滑動視窗內，紅燈+黃燈總數至少要達到這個數量"
    )

    with st.expander("進階設定"):
        # 預設把視窗長度 = 合計門檻（直覺做法）；必要時可放大視窗做更寬鬆偵測
        win_len = st.number_input(
            "滑動視窗長度（週）", min_value=2, max_value=18, value=int(min_total), step=1,
            help="預設等於「紅+黃合計最低數量」。你也可以設更大，表示在較長週數視窗內套用相同門檻。"
        )

    st.markdown("---")
    uploaded = st.file_uploader(
        "上傳成績檔（Excel, 需包含 'score' 工作表，欄位：ID, Name, Biochem, MolBio, Week）",
        type=["xlsx"]
    )

@st.cache_data(show_spinner=False)
def load_score_df(file):
    if file is not None:
        return pd.read_excel(file, sheet_name="score")

    # 預設本地檔
    candidates = [
        Path(__file__).resolve().parent / "2025-biochem_molbio_score.xlsx",
        Path.cwd() / "2025-biochem_molbio_score.xlsx",
        Path(__file__).resolve().parent / "data" / "2025-biochem_molbio_score.xlsx",
        Path.cwd() / "data" / "2025-biochem_molbio_score.xlsx",
    ]
    for p in candidates:
        if p.exists():
            return pd.read_excel(p, sheet_name="score")

    st.error("未上傳檔案，且找不到預設檔")
    st.stop()

# 讀檔
df = load_score_df(uploaded)

# 基本欄位檢查
required_cols = {"ID", "Biochem", "MolBio", "Week"}
if not required_cols.issubset(df.columns):
    st.error(f"缺少必要欄位：{required_cols - set(df.columns)}，請確認 'score' 工作表欄位至少包含 ID, Biochem, MolBio, Week。")
    st.stop()

# Name 欄可選；若沒有則補空白
if "Name" not in df.columns:
    df["Name"] = ""

# 轉型與排序
df["Week"] = pd.to_numeric(df["Week"], errors="coerce").astype("Int64")
df = df.dropna(subset=["Week"]).copy()
df["Week"] = df["Week"].astype(int)
df = df.sort_values(["ID", "Week"])

# 工具函式與衍生欄位
def to_str_id(x):
    try:
        return str(x)
    except Exception:
        return ""

df["ID_str"] = df["ID"].apply(to_str_id)

def cohort_from_id(s: str) -> str:
    try:
        prefix = int(str(s)[:3])
    except Exception:
        return "未知"
    if prefix < 413:
        return "重修"
    elif prefix == 413:
        return "應屆"
    else:
        return "先修"

def dept_from_id(s: str) -> str:
    s = to_str_id(s)
    code = s[3:5] if len(s) >= 5 else ""
    if code == "01":
        return "醫學系"
    elif code == "02":
        return "牙醫學系"
    elif code == "03":
        return "藥學系"
    else:
        return "未知"

df["學籍分類"] = df["ID_str"].apply(cohort_from_id)
df["系所"] = df["ID_str"].apply(dept_from_id)

# 多選篩選（pills）
if cohort_opts:  # 若有選任何項目才過濾
    df = df[df["學籍分類"].isin(cohort_opts)]
if dept_opts:
    df = df[df["系所"].isin(dept_opts)]

# 固定 18 週
WEEKS_FULL = list(range(1, 19))

# ID->Name 對照
id_name_map = (
    df.sort_values(["ID", "Week"])
      .groupby("ID")["Name"]
      .apply(lambda s: s.dropna().iloc[0] if len(s.dropna()) else "")
)

# 透視表（整數＋缺值 pd.NA；欄名轉字串；在 ID 後插入 Name）
def make_pivot(subject_col: str) -> pd.DataFrame:
    pivot = df.pivot(index="ID", columns="Week", values=subject_col).sort_index()
    pivot = pivot.reindex(columns=WEEKS_FULL)
    pivot = pivot.map(lambda x: int(x) if pd.notna(x) else pd.NA)
    pivot.columns = pivot.columns.map(str)
    pivot.insert(0, "Name", pivot.index.map(id_name_map).fillna(""))
    return pivot

bio_pivot = make_pivot("Biochem")
mol_pivot = make_pivot("MolBio")

# 著色（空白與字串不著色）
def color_cell(v):
    try:
        if v == "" or isinstance(v, str):
            return ""
        x = float(v)
    except Exception:
        return ""
    if x <= red_th:
        return "background-color: #f8d7da;"
    elif x <= yellow_th:
        return "background-color: #fff3cd;"
    else:
        return "background-color: #d4edda;"

st.subheader("生物化學（Biochem）")
st.dataframe(bio_pivot.style.map(color_cell), use_container_width=True)

st.subheader("分子生物學（MolBio）")
st.dataframe(mol_pivot.style.map(color_cell), use_container_width=True)

# === AND 規則預警（任一科別同時滿足：紅≥min_red 且 紅+黃≥min_total） ===
def window_any_subject_alert_AND(df_score: pd.DataFrame,
                                 red_threshold: float,
                                 yellow_threshold: float,
                                 window_len: int,
                                 min_red: int,
                                 min_total: int) -> pd.DataFrame:
    """
    在每位學生的連續視窗（長度 window_len）內，若任一科：
      - 紅燈數量(<= red_threshold) 累計 >= min_red           AND
      - 紅燈 + 黃燈數量(<= yellow_threshold) 累計 >= min_total
    則觸發預警。
    例：min_red=2, min_total=4 → 2紅2黃、3紅1黃、4紅0黃皆觸發；1紅3黃不觸發。
    """
    out_rows = []
    if df_score.empty:
        return pd.DataFrame(columns=["ID","Name","Weeks","Biochem_scores","MolBio_scores","觸發條件","視窗長度","學籍","系所"])

    for sid, g in df_score.groupby("ID"):
        g = g.set_index("Week").reindex(WEEKS_FULL)[["Biochem","MolBio"]]
        for i in range(len(WEEKS_FULL) - window_len + 1):
            win_weeks = WEEKS_FULL[i:i+window_len]
            sub = g.loc[win_weeks, ["Biochem", "MolBio"]].copy()

            # 視窗內如有缺值 → 略過（可改成允許缺值但只計有分數週）
            if sub.isna().any().any():
                continue

            def counts(series):
                reds = (series <= red_threshold).sum()
                yellows = ((series > red_threshold) & (series <= yellow_threshold)).sum()
                total = reds + yellows
                return reds, yellows, total

            bio_r, bio_y, bio_t = counts(sub["Biochem"])
            mol_r, mol_y, mol_t = counts(sub["MolBio"])

            triggers = []
            if (bio_r >= min_red) and (bio_t >= min_total):
                triggers.append(f"Biochem：紅≥{min_red} 且 紅+黃≥{min_total}（實得：紅{bio_r}、黃{bio_y}）")
            if (mol_r >= min_red) and (mol_t >= min_total):
                triggers.append(f"MolBio：紅≥{min_red} 且 紅+黃≥{min_total}（實得：紅{mol_r}、黃{mol_y}）")

            if triggers:
                sid_meta = df_score[df_score["ID"] == sid].iloc[0]
                out_rows.append({
                    "ID": sid,
                    "Name": sid_meta.get("Name", ""),
                    "Weeks": f"{win_weeks[0]}–{win_weeks[-1]}",
                    "Biochem_scores": tuple(int(x) for x in sub["Biochem"].values),
                    "MolBio_scores": tuple(int(x) for x in sub["MolBio"].values),
                    "觸發條件": "；".join(triggers),
                    "視窗長度": window_len,
                    "學籍": sid_meta["學籍分類"],
                    "系所": sid_meta["系所"],
                })

    if not out_rows:
        return pd.DataFrame(columns=["ID","Name","Weeks","Biochem_scores","MolBio_scores","觸發條件","視窗長度","學籍","系所"])

    df_out = pd.DataFrame(out_rows)
    df_out = df_out.drop_duplicates(subset=["ID","Weeks","觸發條件"])
    return df_out

# 呼叫（預設 win_len = min_total；可在進階設定改）
alert_df = window_any_subject_alert_AND(
    df_score=df,
    red_threshold=red_th,
    yellow_threshold=yellow_th,
    window_len=int(win_len),
    min_red=int(min_red),
    min_total=int(min_total),
)

# 顯示區塊
st.subheader(
    f"⚠️ 預警名單（視窗={int(win_len)} 週；條件：紅≥{int(min_red)} 且 紅+黃≥{int(min_total)}）"
)
if alert_df.empty:
    st.success("目前沒有符合預警條件的學生。")
else:
    show = alert_df.copy()
    if "Biochem_scores" in show.columns:
        show["Biochem_scores"] = show["Biochem_scores"].apply(lambda xs: "、".join(map(str, xs)))
    if "MolBio_scores" in show.columns:
        show["MolBio_scores"] = show["MolBio_scores"].apply(lambda xs: "、".join(map(str, xs)))

    cols = ["ID", "Name", "學籍", "系所", "Weeks", "Biochem_scores", "MolBio_scores"]
    cols = [c for c in cols if c in show.columns]
    st.dataframe(show[cols], use_container_width=True)

# 匯出
with st.expander("⬇️ 下載目前結果（Excel）"):
    def to_excel_bytes():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            def unblank(df_in: pd.DataFrame) -> pd.DataFrame:
                return df_in.replace("", pd.NA)
            unblank(bio_pivot).to_excel(writer, sheet_name="Biochem_Pivot")
            unblank(mol_pivot).to_excel(writer, sheet_name="MolBio_Pivot")
            (alert_df if not alert_df.empty else pd.DataFrame(columns=[
                "ID","Name","Weeks","Biochem_scores","MolBio_scores","觸發條件","視窗長度","學籍","系所"
            ])).to_excel(writer, sheet_name=f"Alerts_AND_win{int(win_len)}", index=False)
        output.seek(0)
        return output

    st.download_button(
        label="下載 Excel",
        data=to_excel_bytes(),
        file_name="score_dashboard_outputs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption(
    "說明：固定顯示 18 週；空白代表未有成績。紅≤紅色門檻；黃=紅色門檻~黃色上限；綠>黃色上限。"
    "AND 規則：在滑動視窗內，同時滿足「紅燈數量≥設定」且「紅+黃合計≥設定」即列入；預設視窗長度=合計門檻，可在進階設定調整。"
    "學籍依 ID 開頭（<413 重修、=413 應屆、>413 先修）；系所依 ID 第4-5碼。所有表皆在 ID 後顯示姓名。"
)
