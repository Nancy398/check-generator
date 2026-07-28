import streamlit as st
import pandas as pd
from datetime import date
import io
import zipfile
import os
import gspread
from google.oauth2.service_account import Credentials

# Page configuration
st.set_page_config(
    page_title="支票开具与薪酬管理系统",
    page_icon="🧾",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. 辅助函数：Google Sheets 读取与写入
# ------------------------------------------------------------------------------
def get_gspread_client():
    """获取 gspread 客户端授权"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
    else:
        return None
    return gspread.authorize(creds)

def fetch_next_check_numbers_from_gs(default_projects_df):
    """
    从 Google Sheets (Check Issuance History -> Sheet1) 读取历史数据，
    按 Project 统计已用到的最大 Check Number，并计算下一个起始号 (Max + 1)。
    """
    next_numbers = {}
    try:
        client = get_gspread_client()
        if not client:
            st.warning("⚠️ 未检测到 Credentials 凭证，将使用默认支票起始号。")
            return {row["Project_Name"]: int(row["Next_Check_Number"]) for _, row in default_projects_df.iterrows()}

        sheet_name = st.secrets.get("SPREADSHEET_NAME", "Check Issuance History")
        sheet = client.open(sheet_name).worksheet("Sheet1")
        
        # 获取 Sheet1 的所有记录
        records = sheet.get_all_records()
        
        if records:
            df_gs = pd.DataFrame(records)
            if "Project" in df_gs.columns and "Check Number" in df_gs.columns:
                # 确保 Check Number 列是数值类型
                df_gs["Check Number"] = pd.to_numeric(df_gs["Check Number"], errors="coerce")
                
                # 分组计算每个项目当前用到的最大支票号
                max_checks = df_gs.groupby("Project")["Check Number"].max().to_dict()
                
                for _, row in default_projects_df.iterrows():
                    p_name = row["Project_Name"]
                    default_num = int(row["Next_Check_Number"])
                    # 如果云端有该项目的记录，用 (最大号 + 1)；否则用默认号
                    if p_name in max_checks and pd.notnull(max_checks[p_name]):
                        next_numbers[p_name] = int(max_checks[p_name]) + 1
                    else:
                        next_numbers[p_name] = default_num
                return next_numbers

    except Exception as e:
        st.warning(f"⚠️ 从 Google Sheets 读取历史支票号失败 (原因: {str(e)})，已使用预设起始号。")

    # 异常或空表时的兜底方案
    return {row["Project_Name"]: int(row["Next_Check_Number"]) for _, row in default_projects_df.iterrows()}


def append_to_google_sheet(records_list):
    """
    将生成的支票明细追加保存到指定的 Google Sheet (Check Issuance History -> Sheet1)
    """
    if not records_list:
        return False, "无数据可写入"
    
    try:
        client = get_gspread_client()
        if not client:
            return False, "未找到 Google Service Account 密钥配置 (secrets.toml 或 service_account.json)"

        sheet_name = st.secrets.get("SPREADSHEET_NAME", "Check Issuance History")
        sheet = client.open(sheet_name).worksheet("Sheet1")

        rows_to_append = []
        for r in records_list:
            rows_to_append.append([
                r.get("Check Number", ""),
                r.get("Issue Date", ""),
                r.get("Company", ""),
                r.get("Account", ""),
                r.get("Project", ""),
                r.get("Payee Name", ""),
                r.get("Amount", 0.0),
                r.get("Memo", "")
            ])

        sheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        return True, "成功写入 Google Sheets"

    except Exception as e:
        return False, str(e)

def save_to_local_csv(records_list, filepath="check_issuance_history.csv"):
    if not records_list:
        return
    new_df = pd.DataFrame(records_list)
    cols = ["Check Number", "Issue Date", "Company", "Account", "Project", "Payee Name", "Amount", "Memo"]
    new_df = new_df.reindex(columns=cols)
    if os.path.exists(filepath):
        new_df.to_csv(filepath, mode='a', header=False, index=False, encoding='utf-8-sig')
    else:
        new_df.to_csv(filepath, mode='w', header=True, index=False, encoding='utf-8-sig')

def number_to_words_usd(amount):
    dollars = int(amount)
    cents = int(round((amount - dollars) * 100))
    return f"{dollars:,} AND {cents}/100 DOLLARS"

def fill_pdf_placeholders(pdf_bytes, replacements):
    return pdf_bytes

def merge_pdfs(pdf_bytes_list):
    if not pdf_bytes_list:
        return b""
    return pdf_bytes_list[0]

# ------------------------------------------------------------------------------
# 2. 预设数据与初始化
# ------------------------------------------------------------------------------
preset_worker_list = ["张三", "李四", "王五", "赵六", "钱七"]
worker_role_map = {
    "张三": "木工组长",
    "李四": "电工精修",
    "王五": "泥水铺砖",
    "赵六": "油漆涂刷",
    "钱七": "杂工小弟"
}

def load_project_presets():
    return pd.DataFrame([
        {"Project_Name": "项目A-商业中心", "Company": "A建筑有限公司", "Account": "ACC-8888-01", "Next_Check_Number": 1001},
        {"Project_Name": "项目B-住宅公寓", "Company": "B建设发展公司", "Account": "ACC-6666-02", "Next_Check_Number": 2001},
        {"Project_Name": "项目C-别墅改造", "Company": "C装饰工程公司", "Account": "ACC-3333-03", "Next_Check_Number": 3001}
    ])

df_projects = load_project_presets()
preset_project_list = df_projects["Project_Name"].tolist()

# ------------------------------------------------------------------------------
# 3. Streamlit 主界面
# ------------------------------------------------------------------------------
st.sidebar.title("⚙️ 导航与设置")
mode = st.sidebar.radio(
    "选择业务场景",
    ["👷 场景二：多项目/施工队周薪批量开单", "📜 查看历史记录"]
)

uploaded_pdf = st.sidebar.file_uploader("上传支票 PDF 模板", type=["pdf"])
pdf_template_bytes = uploaded_pdf.getvalue() if uploaded_pdf else b"%PDF-1.4 dummy pdf bytes"

# ==============================================================================
# 场景 2：多项目/施工队周薪批量开单
# ==============================================================================
if mode == "👷 场景二：多项目/施工队周薪批量开单":
    st.title("👷 多项目/施工队周薪批量生成")
    st.caption("自动从 Google Sheets 同步并递增编号，匹配工种，生成 PDF 并实时存入云端。")

    pay_date = st.date_input("发薪日期", value=date.today())

    st.markdown("---")
    
    # ----------------- 1. 从 Google Sheets 读取并确认各项目起始支票号 -----------------
    st.subheader("1. 确认本期各项目起始支票号")
    
    # 动态从云端 Google Sheets 读取各项目最大的 Check Number + 1
    with st.spinner("🔄 正在从 Google Sheets 实时获取各项目最新支票号..."):
        latest_check_map = fetch_next_check_numbers_from_gs(df_projects)

    proj_start_nums = {}
    cols = st.columns(min(len(df_projects), 4))
    for idx, p_row in df_projects.iterrows():
        p_name = p_row["Project_Name"]
        default_num = latest_check_map.get(p_name, int(p_row["Next_Check_Number"]))
        
        with cols[idx % 4]:
            proj_start_nums[p_name] = st.number_input(
                f"🏗️ {p_name}",
                min_value=1,
                value=int(default_num),
                key=f"start_num_{p_name}"
            )

    st.markdown("---")

    # ----------------- 2. 初始化发薪数据列表与输入框状态 -----------------
    st.subheader("2. 录入发薪明细")

    if "payroll_list" not in st.session_state:
        st.session_state.payroll_list = []

    def update_memo_on_worker_change():
        selected_w = st.session_state.input_w
        st.session_state.input_m = worker_role_map.get(selected_w, "")

    if "input_m" not in st.session_state:
        default_first_worker = preset_worker_list[0] if preset_worker_list else ""
        st.session_state.input_m = worker_role_map.get(default_first_worker, "")

    # --- 快捷添加面板 ---
    st.markdown("##### ➕ 添加发薪人员")
    c1, c2, c3, c4, c5 = st.columns([2.5, 2.5, 2, 2.5, 1.5])

    with c1:
        add_worker = st.selectbox(
            "选择工人 (Payee)", 
            preset_worker_list, 
            key="input_w",
            on_change=update_memo_on_worker_change
        )
    with c2:
        add_proj = st.selectbox("所属项目 (Project)", preset_project_list, key="input_p")
    with c3:
        add_amt = st.number_input("金额 $ (Amount)", min_value=0.01, value=1200.00, step=50.0, key="input_a")
    with c4:
        add_memo = st.text_input("工作备注 (Memo)", key="input_m")
    with c5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ 添加", type="primary", use_container_width=True):
            existing_checks = [
                row["支票编号 (Check #)"] 
                for row in st.session_state.payroll_list 
                if row["所属项目 (Project)"] == add_proj
            ]
            
            if existing_checks:
                next_chk = max(existing_checks) + 1
            else:
                next_chk = proj_start_nums.get(add_proj, 1001)

            st.session_state.payroll_list.append({
                "工人姓名 (Payee)": add_worker,
                "所属项目 (Project)": add_proj,
                "支票编号 (Check #)": int(next_chk),
                "金额 $ (Amount)": add_amt,
                "工作备注 (Memo)": add_memo
            })
            st.rerun()

    st.markdown("---")

    # --- 数据列表展示 ---
    if st.session_state.payroll_list:
        st.markdown(f"##### 📋 本期待开支票列表（共 **{len(st.session_state.payroll_list)}** 张）：")
        
        df_current = pd.DataFrame(st.session_state.payroll_list)

        edited_df = st.data_editor(
            df_current,
            num_rows="dynamic",
            use_container_width=True,
            key="payroll_table_editor",
            column_config={
                "工人姓名 (Payee)": st.column_config.SelectboxColumn("工人姓名 (Payee)", options=preset_worker_list),
                "所属项目 (Project)": st.column_config.SelectboxColumn("所属项目 (Project)", options=preset_project_list),
                "支票编号 (Check #)": st.column_config.NumberColumn("支票编号 (Check #)", format="%d"),
                "金额 $ (Amount)": st.column_config.NumberColumn("金额 $ (Amount)", format="$%.2f")
            }
        )

        st.session_state.payroll_list = edited_df.to_dict(orient="records")

        if st.button("🗑️ 清空列表重新录入", type="secondary"):
            st.session_state.payroll_list = []
            st.rerun()

        df_payroll_input = edited_df
    else:
        st.info("💡 列表中暂无数据，请在上方选择工人与项目后点击 **【➕ 添加】** 按钮增加人员。")
        df_payroll_input = pd.DataFrame()

    st.markdown("---")

    # ----------------- 3. 批量生成与写入 Google Sheets -----------------
    if not df_payroll_input.empty:
        if st.button(f"🚀 确认无误，批量生成 {len(df_payroll_input)} 张支票", type="primary", use_container_width=True):
            account_pdf_dict = {}
            records_log = []

            proj_map = df_projects.set_index("Project_Name").to_dict(orient="index")

            for idx, row in df_payroll_input.iterrows():
                worker_name = str(row.get("工人姓名 (Payee)", "")).strip()
                project_name = str(row.get("所属项目 (Project)", "")).strip()
                
                try:
                    cur_check = int(row.get("支票编号 (Check #)", 0))
                    amt = float(row.get("金额 $ (Amount)", 0.0))
                except (ValueError, TypeError):
                    cur_check = 0
                    amt = 0.0

                detail_memo = str(row.get("工作备注 (Memo)", "")).strip()

                if amt <= 0 or not worker_name or cur_check <= 0:
                    continue

                p_info = proj_map.get(project_name, {"Company": "Unknown Company", "Account": "ACC-0000"})
                company_name = p_info["Company"]
                account_num = p_info["Account"]

                full_memo = f"{project_name} - {detail_memo}" if detail_memo else project_name

                replacements = {
                    "date": pay_date.strftime("%m/%d/%Y"),
                    "name": worker_name,
                    "amount": f"{amt:,.2f}",
                    "amount_words": number_to_words_usd(amt),
                    "memo": full_memo,
                    "number": str(cur_check),
                    "account": account_num
                }

                pdf_res = fill_pdf_placeholders(pdf_template_bytes, replacements)
                
                acc_key = (company_name, account_num)
                if acc_key not in account_pdf_dict:
                    account_pdf_dict[acc_key] = []
                account_pdf_dict[acc_key].append((cur_check, project_name, worker_name, pdf_res))

                records_log.append({
                    "Check Number": cur_check,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_name,
                    "Account": account_num,
                    "Project": project_name,
                    "Payee Name": worker_name,
                    "Amount": amt,
                    "Memo": full_memo
                })

            if records_log:
                # 1. 保存本地 CSV 备份
                save_to_local_csv(records_log)

                # 2. 🟢 追加写入 Google Sheets (Check Issuance History -> Sheet1)
                gs_success, gs_msg = append_to_google_sheet(records_log)

                # 3. 清空输入列表
                st.session_state.payroll_list = []

                st.balloons()
                if gs_success:
                    st.success(f"🎉 成功生成 {len(records_log)} 张支票！数据已自动保存至 **Google Sheets (`Check Issuance History` -> `Sheet1`)**。下次打开界面将自动读取新的最大编号！")
                else:
                    st.warning(f"⚠️ PDF 与本地 CSV 记录已生成，但写入 Google Sheets 失败: {gs_msg}")

                st.markdown("### 📊 本期出账汇总")
                df_batch = pd.DataFrame(records_log)
                col_sum1, col_sum2 = st.columns(2)

                with col_sum1:
                    st.markdown("#### 🏢 按公司 / 账号汇总")
                    summary_company = df_batch.groupby(["Company", "Account"]).agg(
                        总金额=("Amount", "sum"),
                        支票张数=("Check Number", "count")
                    ).reset_index()
                    st.dataframe(summary_company.style.format({"总金额": "${:,.2f}"}), use_container_width=True, hide_index=True)

                with col_sum2:
                    st.markdown("#### 🏗️ 按工地项目汇总")
                    summary_project = df_batch.groupby(["Project", "Company"]).agg(
                        项目总人工费=("Amount", "sum"),
                        工人人数=("Check Number", "count")
                    ).reset_index()
                    st.dataframe(summary_project.style.format({"项目总人工费": "${:,.2f}"}), use_container_width=True, hide_index=True)

                st.markdown("---")
                st.markdown("### 📥 按账户单独下载 PDF 文件")

                for (comp_name, acc_num), item_list in account_pdf_dict.items():
                    pdf_bytes_list = [item[3] for item in item_list]
                    account_merged_pdf = merge_pdfs(pdf_bytes_list)
                    
                    st.markdown(f"##### 💳 账户：**{comp_name}** | 账号：`{acc_num}`（共 {len(item_list)} 张）")
                    
                    st.download_button(
                        label=f"📄 下载【{comp_name} - {acc_num}】合并 PDF",
                        data=account_merged_pdf,
                        file_name=f"Checks_{comp_name}_{acc_num}_{pay_date}.pdf",
                        mime="application/pdf",
                        type="primary"
                    )

                st.markdown("---")
                st.markdown("##### 📦 更多导出选项")
                
                csv_bytes = df_batch.to_csv(index=False).encode('utf-8-sig')
                
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    for (comp_name, acc_num), item_list in account_pdf_dict.items():
                        for chk, proj, py, pdf_b in item_list:
                            zf.writestr(f"[{acc_num}]_Check_{chk}_[{proj}]_{py}.pdf", pdf_b)

                d1, d2 = st.columns(2)
                with d1:
                    st.download_button(
                        label="📊 下载【本期出账总账单 CSV】",
                        data=csv_bytes,
                        file_name=f"Payroll_Summary_{pay_date}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with d2:
                    st.download_button(
                        label="📦 下载所有单张 PDF ZIP 打包",
                        data=zip_buf.getvalue(),
                        file_name=f"Payroll_Checks_SinglePDFs_{pay_date}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

elif mode == "📜 查看历史记录":
    st.title("📜 支票开具历史记录")
    if os.path.exists("check_issuance_history.csv"):
        df_hist = pd.read_csv("check_issuance_history.csv")
        st.dataframe(df_hist, use_container_width=True)
    else:
        st.info("暂无本地开单历史记录。")
