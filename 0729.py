import streamlit as st
import pandas as pd
import math
import io
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="圖形碼排程系統", layout="centered")
st.title("工單排程系統")

# === 跨平台日期顯示工具（M/D，無前導 0） ===
def format_md(dt) -> str:
    """安全輸出 M/D（無前導 0），跨平台可用。"""
    ts = pd.to_datetime(dt)
    return f"{ts.month}/{ts.day}"

# === 安全串接工具 ===
def safe_concat(dfs):
    valid = [d for d in dfs if isinstance(d, pd.DataFrame) and not d.empty]
    return pd.concat(valid, ignore_index=True) if valid else pd.DataFrame()

# === 上傳區 ===
file_jobs = st.file_uploader("匯入工單 Excel 檔案：", type="xlsx")
file_machines = st.file_uploader("匯入機台 Excel 檔案：", type="xlsx")

shuffle_code = st.checkbox("打亂圖形碼")
group_priority = st.checkbox("按群組優先圖形碼")

# === 新增產能估算欄位 ===
custom_target_quantity = st.text_input("預計完成工單筆數：")
input_machines = st.text_input("今日開機機台數（模式一）：")
custom_target_days = st.text_input("希望完成天數（模式二）：")

# === 初始化 ===
df_jobs = pd.DataFrame()
df_machines_raw = pd.DataFrame()
priority_list = []
priority_text_display = ""

# === 工單處理 ===
if file_jobs:
    job_xls = pd.ExcelFile(file_jobs)
    job_sheet = st.selectbox("選擇工單工作表：", job_xls.sheet_names)
    if job_sheet:
        df_jobs = pd.read_excel(job_xls, sheet_name=job_sheet)
        if "排程日" not in df_jobs.columns:
            st.error("工單中缺少『排程日』欄位")
            st.stop()
        df_jobs["排程日"] = pd.to_datetime(df_jobs["排程日"], errors="coerce")

# === 機台處理 ===
if file_machines:
    machine_xls = pd.ExcelFile(file_machines)
    selected_sheets = st.multiselect("選擇要使用的機台工作表（可複選）：", machine_xls.sheet_names)
    if selected_sheets:
        priority_sheet_name = "優先圖形碼"
        normal_sheets = [s for s in selected_sheets if s != priority_sheet_name]
        if normal_sheets:
            df_machines_raw = pd.concat(
                [pd.read_excel(machine_xls, sheet_name=s) for s in normal_sheets],
                ignore_index=True
            )
            if "備註" not in df_machines_raw.columns:
                df_machines_raw["備註"] = ""

        if group_priority and priority_sheet_name in selected_sheets:
            try:
                df_priority = pd.read_excel(machine_xls, sheet_name=priority_sheet_name)
                if "優先圖形碼" in df_priority.columns:
                    priority_list = df_priority["優先圖形碼"].dropna().astype(str).tolist()
                    priority_text_display = "\n".join(priority_list)
                    st.success(f"已讀取 {len(priority_list)} 組優先圖形碼")
                else:
                    st.warning("優先圖形碼工作表中找不到『優先圖形碼』欄位")
            except Exception as e:
                st.warning(f"讀取優先圖形碼失敗：{e}")

priority_codes = st.text_area("優先圖形碼（可修改）：", value=priority_text_display, height=120)
priority_list = [code.strip() for code in priority_codes.replace("\n", ",").split(",") if code.strip()]

# === 廠區檢查與分離 ===
if not df_jobs.empty and not df_machines_raw.empty:
    has_2A = any(df_jobs.iloc[:, 0].astype(str).str.startswith("2A"))
    has_2B = any(df_jobs.iloc[:, 0].astype(str).str.startswith("2B"))
    if has_2A and has_2B:
        if not any(df_machines_raw["備註"].astype(str).str.contains("2A")) or not any(df_machines_raw["備註"].astype(str).str.contains("2B")):
            st.error("錯誤：工單中同時有 2A 與 2B，但機台未標註廠區（備註欄需含 2A 或 2B）")
            st.stop()
    df_machines_A = df_machines_raw[df_machines_raw["備註"].astype(str).str.contains("2A")].copy()
    df_machines_B = df_machines_raw[df_machines_raw["備註"].astype(str).str.contains("2B")].copy()
    if not has_2B:
        df_machines_A = df_machines_raw.copy()
    if not has_2A:
        df_machines_B = df_machines_raw.copy()

if st.button("開始排程"):
    if file_jobs is None or file_machines is None or not job_sheet or not selected_sheets:
        st.error("請上傳工單與機台檔案，並選擇工作表")
        st.stop()

    # 若勾選打亂圖形碼：依圖形碼分組交錯取出，避免相同連續
    if shuffle_code:
        code_col = next((col for col in df_jobs.columns if "圖形碼" in col), None)
        if code_col:
            grouped = df_jobs.groupby(code_col)
            shuffled_rows = []
            max_len = max(len(g) for _, g in grouped)
            for i in range(max_len):
                for _, group in grouped:
                    if i < len(group):
                        shuffled_rows.append(group.iloc[i])
            df_jobs = pd.DataFrame(shuffled_rows).reset_index(drop=True)

    # 標註廠區
    df_jobs["廠區"] = df_jobs.iloc[:, 0].astype(str).str[:2]

    # === 核心排程函式 ===
    def schedule_jobs(df_jobs_one_side, df_machines, job_date):
        # 機台清單與每日產能
        if df_machines.empty or "Machine_ID" not in df_machines.columns:
            return pd.DataFrame(), set(), pd.DataFrame()

        machine_ids = df_machines["Machine_ID"].dropna().tolist()
        machine_capacity = {mid: 86400 for mid in machine_ids}
        result = []
        unassigned_rows = []
        used_machines = set()

        for _, row in df_jobs_one_side.iterrows():
            # 這裡假設排隊數在第 5 欄（iloc[4]），可依實際欄位名稱調整
            count = row.iloc[4]
            try:
                count_val = float(count)
            except:
                count_val = 0.0
            duration = math.ceil((count_val / 25) * 30 + 900)

            assigned = False
            for m in machine_ids:
                if machine_capacity.get(m, 0) >= duration:
                    start_sec = 86400 - machine_capacity[m]
                    machine_capacity[m] -= duration
                    result.append({
                        **row.to_dict(),
                        "所需秒數": duration,
                        "機台編號": m,
                        "開始": pd.to_datetime(job_date) + timedelta(seconds=start_sec),
                        "結束": pd.to_datetime(job_date) + timedelta(seconds=start_sec + duration)
                    })
                    used_machines.add(m)
                    assigned = True
                    break

            if not assigned:
                unassigned_rows.append({
                    **row.to_dict(),
                    "所需秒數": None,
                    "機台編號": "未排入",
                    "開始": pd.to_datetime(job_date),
                    "結束": pd.to_datetime(job_date)
                })

        return pd.DataFrame(result), used_machines, pd.DataFrame(unassigned_rows)

    df_jobs_all = []
    df_unassigned_all = []
    daily_machine_dict = {}

    # === 逐日排程 ===
    for day in sorted(df_jobs["排程日"].dropna().unique()):
        daily_jobs = df_jobs[df_jobs["排程日"] == day].copy()

        # 當日優先圖形碼排序
        if group_priority and priority_list:
            code_col = next((c for c in daily_jobs.columns if "圖形碼" in c), None)
            if code_col:
                daily_jobs["__key"] = daily_jobs[code_col].astype(str).apply(
                    lambda x: priority_list.index(x) if x in priority_list else len(priority_list)
                )
                daily_jobs = daily_jobs.sort_values("__key").drop(columns="__key")

        # 分 2A / 2B
        daily_2A = daily_jobs[daily_jobs["廠區"] == "2A"]
        daily_2B = daily_jobs[daily_jobs["廠區"] == "2B"]

        result_2A, machines_2A, unassigned_2A = schedule_jobs(daily_2A, df_machines_A, day)
        result_2B, machines_2B, unassigned_2B = schedule_jobs(daily_2B, df_machines_B, day)

        df_jobs_all.extend([result_2A, result_2B])
        df_unassigned_all.extend([unassigned_2A, unassigned_2B])

        # 紀錄當日哪些機台有開機（用 M/D 字串）
        day_str = format_md(day)
        for m in sorted(machines_2A.union(machines_2B), key=lambda x: str(x)):
            daily_machine_dict.setdefault(m, {})[day_str] = True

    # === 整理每日開機機台表 ===
    all_machines = sorted(daily_machine_dict.keys(), key=lambda x: str(x))
    all_days = sorted(set(format_md(d) for d in df_jobs["排程日"].dropna()))
    machine_df = pd.DataFrame({"Machine_ID": all_machines})
    for day in all_days:
        machine_df[day] = machine_df["Machine_ID"].apply(lambda m: daily_machine_dict.get(m, {}).get(day, False))

    # === 合併結果（安全串接） ===
    df_result = safe_concat(df_jobs_all)
    df_unassigned = safe_concat(df_unassigned_all)

    # === 產能估算分析 ===
    machine_sec_per_day = 86400
    current_machine_count = len(df_machines_raw[df_machines_raw["Machine_ID"].notna()]) if not df_machines_raw.empty else 0
    mode1_df, mode2_df = pd.DataFrame(), pd.DataFrame()
    mode1_valid, mode2_valid = False, False

    if custom_target_quantity and input_machines:
        try:
            target_quantity = int(custom_target_quantity)
            opened_machines = int(input_machines)
            total_needed_seconds = int(math.ceil((target_quantity / 25) * 30 + 900))
            total_daily_capacity = opened_machines * machine_sec_per_day
            estimated_days = math.ceil(total_needed_seconds / total_daily_capacity) if total_daily_capacity > 0 else None
            mode1_df = pd.DataFrame({
                "目標排隊數": [target_quantity],
                "輸入開機機台數": [opened_machines],
                "每台每日產能(秒)": [machine_sec_per_day],
                "總所需秒數": [total_needed_seconds],
                "預估完成天數": [estimated_days],
            })
            mode1_valid = estimated_days is not None
        except:
            st.warning("請輸入有效的『預計完成工單筆數』與『開機機台數』")

    if custom_target_quantity and custom_target_days:
        try:
            target_quantity = int(custom_target_quantity)
            target_days = int(custom_target_days)
            total_needed_seconds = int(math.ceil((target_quantity / 25) * 30 + 900))
            sec_needed_per_day = total_needed_seconds / target_days if target_days > 0 else float("inf")
            required_machines = int(math.ceil(sec_needed_per_day / machine_sec_per_day)) if target_days > 0 else None
            machine_gap = max(0, (required_machines or 0) - current_machine_count) if required_machines is not None else None

            mode2_df = pd.DataFrame({
                "目標排隊數": [target_quantity],
                "希望完成天數": [target_days],
                "總所需秒數": [total_needed_seconds],
                "每天所需秒數": [int(math.ceil(sec_needed_per_day)) if target_days > 0 else None],
                "每台每日產能(秒)": [machine_sec_per_day],
                "所需機台數": [required_machines],
                "現有機台數": [current_machine_count],
                "還需新增機台數": [machine_gap],
            })
            mode2_valid = required_machines is not None
        except:
            st.warning("請輸入有效的『預計完成工單筆數』與『完成天數』")

    if not (mode1_valid or mode2_valid):
        st.error("請至少正確輸入以下其中之一：\n- 預計完成工單筆數 + 開機台數\n- 預計完成工單筆數 + 完成天數")
        st.stop()

    # === 輔助：清理 Excel 工作表名稱 ===
    def sanitize_sheet_name(name: str) -> str:
        # 移除 Excel 禁用字元：: \ / ? * [ ]
        name = re.sub(r'[:\\/*?\[\]]', '_', str(name))
        # Excel 工作表名最長 31 字元
        return name[:31]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="yyyy-mm-dd HH:mm:ss") as writer:
        # 分廠區輸出邏輯
        has_A = not df_result[df_result.get("廠區", "") == "2A"].empty if not df_result.empty else False
        has_B = not df_result[df_result.get("廠區", "") == "2B"].empty if not df_result.empty else False
        if has_A:
            df_result[df_result["廠區"] == "2A"].to_excel(writer, sheet_name="A廠區排程結果", index=False)
        if has_B:
            df_result[df_result["廠區"] == "2B"].to_excel(writer, sheet_name="B廠區排程結果", index=False)
        if not has_A and not has_B and not df_result.empty:
            df_result.to_excel(writer, sheet_name="排程結果", index=False)

        # 每日開機機台
        (machine_df if not machine_df.empty else pd.DataFrame({"Machine_ID": []})).to_excel(
            writer, sheet_name="每日開機機台", index=False
        )

        # 未排入工單（即使為空也輸出，避免後續流程依賴）
        df_unassigned.to_excel(writer, sheet_name="未排入工單", index=False)

        # 產能與機台估算分析（兩段式寫入）
        row_cursor = 0
        if mode1_valid:
            mode1_df.to_excel(writer, sheet_name="產能與機台估算分析", index=False, startrow=row_cursor)
            row_cursor += len(mode1_df) + 3
        if mode2_valid:
            mode2_df.to_excel(writer, sheet_name="產能與機台估算分析", index=False, startrow=row_cursor)

        # 🆕 每台機台一張工作表（僅輸出有被指派到工單的機台）
        if not df_result.empty and "機台編號" in df_result.columns:
            used_machines = (
                df_result.loc[
                    df_result["機台編號"].notna() & (df_result["機台編號"].astype(str) != "未排入"),
                    "機台編號"
                ]
                .astype(str)
                .unique()
                .tolist()
            )
            for mid in sorted(used_machines, key=lambda x: str(x)):
                df_m = df_result[df_result["機台編號"].astype(str) == str(mid)].copy()
                sort_cols = [c for c in ["排程日", "開始"] if c in df_m.columns]
                if sort_cols:
                    df_m = df_m.sort_values(by=sort_cols, kind="mergesort")
                sheet_name = sanitize_sheet_name(f"機台_{mid}")
                df_m.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    file_name_base = file_jobs.name.rsplit(".", 1)[0]
    date_suffix = datetime.now().strftime("%m%d")
    file_name_output = f"{file_name_base}_排程結果_{date_suffix}.xlsx"

    st.success("排程完成，請下載結果：")
    st.download_button(
        label="下載排程結果 Excel",
        data=output,
        file_name=file_name_output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
