import base64
from datetime import date
import io
import os
import zipfile
import fitz
from num2words import num2words
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="Check Generator System", page_icon="🧾", layout="wide"
)

# ----------------- 配置文件与路径定义 -----------------
DEFAULT_TEMPLATE_PATH = "check_run.pdf"
PROJECTS_CSV = "projects_config.csv"
WORKERS_CSV = "workers_config.csv"

GS_SPREADSHEET_NAME = "Check Issuance History"  # Google 表格的名字
GS_WORKSHEET_NAME = "Sheet1"                  # 工作表的名字

# 预设常见施工/业务阶段 (Stage)
PRESET_STAGES = [
    "Stage 1: Demolition / Site Prep",
    "Stage 2: Foundation & Framing",
    "Stage 3: Rough-In (MEP)",
    "Stage 4: Drywall & Insulation",
    "Stage 5: Finishes & Painting",
    "Stage 6: Final Inspection / Cleanup",
    "Other / Custom Stage"
]

# ----------------- 1. 获取 Authorization 客户端 -----------------
def get_gc_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["GOOGLE_APPLICATION_CREDENTIALS"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(credentials)

# ----------------- 2. 读取 Google Sheets -----------------
def read_file(name, sheet):
    gc = get_gc_client()
    worksheet = gc.open(name).worksheet(sheet)
    rows = worksheet.get_all_values()
    if not rows or len(rows) <= 1:
        return pd.DataFrame()
    df = pd.DataFrame.from_records(rows)
    df = pd.DataFrame(df.values[1:], columns=df.iloc[0])
    return df

# ----------------- 3. 保存写入 Google Sheets -----------------
def save_to_history(records):
    """把生成的支票记录直接追加保存到 Google Sheets 中 (已增加 Stage 字段)"""
    try:
        gc = get_gc_client()
        worksheet = gc.open(GS_SPREADSHEET_NAME).worksheet(GS_WORKSHEET_NAME)
        
        # 将记录字典转换为对应的表格行数据列表
        rows_to_append = []
        for r in records:
            rows_to_append.append([
                r["Check Number"],
                r["Issue Date"],
                r["Company"],
                r["Account"],
                r["Project"],
                r.get("Stage", ""),  # 包含 Stage 字段
                r["Payee Name"],
                r["Amount"],
                r["Memo"]
            ])
        # 使用 gspread 的 append_rows 实现云端保存
        worksheet.append_rows(rows_to_append)
        return True
    except Exception as e:
        st.error(f"❌ 数据保存至 Google Sheets 失败: {e}")
        return False

# ----------------- 4. 读取云端表格并计算最新可用支票号 -----------------
def fetch_next_check_numbers_from_gs(df_p_list, default_start_number=1001):
    """计算各个项目最新的 Check Number (+1)"""
    next_numbers = {}
    try:
        df_gs = read_file(GS_SPREADSHEET_NAME, GS_WORKSHEET_NAME)
        if not df_gs.empty and "Project" in df_gs.columns and "Check Number" in df_gs.columns:
            df_gs["Check Number"] = pd.to_numeric(df_gs["Check Number"], errors="coerce")
            max_checks = df_gs.groupby("Project")["Check Number"].max().to_dict()
            
            for p_name in df_p_list:
                if p_name in max_checks and pd.notnull(max_checks[p_name]):
                    next_numbers[p_name] = int(max_checks[p_name]) + 1
    except Exception as e:
        st.sidebar.warning(f"⚠️ 读取云端历史推算支票号失败，将使用默认起始号。({e})")

    # 兜底默认值
    for p_name in df_p_list:
        if p_name not in next_numbers:
            next_numbers[p_name] = default_start_number
            
    return next_numbers

# ----------------- 初始化/读取基础预设文件 -----------------
def load_project_presets():
    """读取或创建公司与工地项目对应表"""
    if not os.path.exists(PROJECTS_CSV):
        df_default = pd.DataFrame(
            [
                {
                    "Project_Name": "123 Main St",
                    "Company": "Moo Construction Inc",
                    "Account": "ACC-8652",
                },
                {
                    "Project_Name": "456 Oak Ave",
                    "Company": "Moo Construction Inc",
                    "Account": "ACC-3738",
                },
                {
                    "Project_Name": "789 Pine Rd",
                    "Company": "Moo Housing Inc",
                    "Account": "ACC-8652",
                },
            ]
        )
        df_default.to_csv(PROJECTS_CSV, index=False)

    df_p = pd.read_csv(PROJECTS_CSV)
    if "Next_Check_Number" in df_p.columns:
        df_p = df_p.drop(columns=["Next_Check_Number"])
        df_p.to_csv(PROJECTS_CSV, index=False)
    return df_p

def load_worker_presets():
    """读取常用工人及其默认岗位 (Default_Role)"""
    if not os.path.exists(WORKERS_CSV):
        df_default = pd.DataFrame(
            [
                {"Worker_Name": "John Smith", "Default_Role": "Framing Lead"},
                {"Worker_Name": "Carlos Mendez", "Default_Role": "Drywaller"},
                {"Worker_Name": "David Lee", "Default_Role": "Electrician"},
                {"Worker_Name": "Jose Rodriguez", "Default_Role": "General Labor"},
            ]
        )
        df_default.to_csv(WORKERS_CSV, index=False)

    df_workers = pd.read_csv(WORKERS_CSV)
    if "Default_Role" not in df_workers.columns:
        df_workers["Default_Role"] = "Worker"
        df_workers.to_csv(WORKERS_CSV, index=False)
    return df_workers

df_projects = load_project_presets()
df_workers = load_worker_presets()

# 构建工人及角色的字典与列表
worker_role_map = dict(zip(df_workers["Worker_Name"], df_workers["Default_Role"]))
preset_worker_list = df_workers["Worker_Name"].dropna().tolist()
preset_project_list = df_projects["Project_Name"].dropna().tolist()

# 每次刷新页面实时从 Google Sheets 提取最新支票号
latest_check_map = fetch_next_check_numbers_from_gs(preset_project_list)

# ----------------- 核心工具函数 -----------------
def number_to_words_usd(amount):
    """金额转英文大写"""
    try:
        dollars = int(amount)
        cents = int(round((amount - dollars) * 100))
        words = num2words(dollars, lang="en").title()
        return f"{words} and {cents:02d}/100 Dollars"
    except Exception:
        return ""

def fill_pdf_placeholders(pdf_bytes, replacements):
    """填充 PDF 模板"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        for key, val in replacements.items():
            str_val = str(val) if val is not None else ""
            patterns = [
                f"{{{{ {key} }}}}",
                f"{{{{{key}}}}}",
                f"{{{{  {key}  }}}}",
            ]
            for pattern in patterns:
                rects = page.search_for(pattern)
                for rect in rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))
                    page.apply_redactions()
                    point = fitz.Point(rect.x0, rect.y1 - 2)
                    page.insert_text(
                        point, str_val, fontsize=10, color=(0, 0, 0)
                    )
    output_stream = io.BytesIO()
    doc.save(output_stream)
    doc.close()
    return output_stream.getvalue()

def merge_pdfs(pdf_bytes_list):
    """合并 PDF"""
    merged_doc = fitz.open()
    for b in pdf_bytes_list:
        doc = fitz.open(stream=b, filetype="pdf")
        merged_doc.insert_pdf(doc)
        doc.close()
    out_stream = io.BytesIO()
    merged_doc.save(out_stream)
    merged_doc.close()
    return out_stream.getvalue()

# ----------------- 模板检测 -----------------
pdf_template_bytes = None
if os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        pdf_template_bytes = f.read()

# ----------------- 页面架构 -----------------
st.sidebar.title("⚙️ 系统导航")
mode = st.sidebar.radio(
    "请选择业务场景：",
    [
        "📝 Single Mannul Check",
        "👷 Construction Bulk Checks",
    ],
)

if pdf_template_bytes is None:
    st.warning("Please Upload Template")
    uploaded_tpl = st.sidebar.file_uploader("Upload Template", type=["pdf"])
    if uploaded_tpl:
        pdf_template_bytes = uploaded_tpl.read()

# ==============================================================================
# 场景 1：单张手动生成支票
# ==============================================================================
if mode == "📝 Single Mannul Check":
    st.title("📝 Single Mannul Check")
    st.caption("Please enter the information")

    if not pdf_template_bytes:
        st.stop()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Information")

        biz_mode = st.radio(
            "Business Entity：",
            ["🏗️ Moo Construction", "🏠 Moo Housing Inc"],
            horizontal=True,
        )

        st.markdown("---")

        if "Construction" in biz_mode:
            company_name = "Moo Construction"
            
            project_options = preset_project_list + ["+ New Project"]
            selected_proj = st.selectbox("Project", project_options)

            if selected_proj != "+ New Project":
                p_info = df_projects[df_projects["Project_Name"] == selected_proj].iloc[0]
                default_account = p_info["Account"]
                project_site = selected_proj
                default_chk_val = latest_check_map.get(selected_proj, 1001)
            else:
                project_site = st.text_input("Project Name", value="New Site")
                default_account = "ACC-8652"
                default_chk_val = 1001

            account_num = st.text_input("Bank Account", value=default_account)
            
            # --- 增加 Stage 选择 ---
            stage_choice = st.selectbox("Stage", PRESET_STAGES)
            if stage_choice == "Other / Custom Stage":
                selected_stage = st.text_input("Custom Stage", value="Stage 1")
            else:
                selected_stage = stage_choice.split(":")[0].strip() # 简化显示为 Stage X
            user_input = st.text_input(
                "Memo", 
                value=""
            )
            
            # 2. 最终使用的完整 Memo 文本
            memo_text = f"{project_site} [{selected_stage}] - {user_input}" if user_input else f"{project_site} [{selected_stage}]"

        else:
            company_name = "Moo Housing Inc"
            
            project_options = preset_project_list + ["+ New Project"]
            selected_proj = st.selectbox("Project", project_options)

            if selected_proj != "+ New Project":
                project_site = selected_proj
                default_chk_val = latest_check_map.get(selected_proj, 1001)
            else:
                project_site = st.text_input("Enter Address", value="Moo Housing Property")
                default_chk_val = 1001

            account_choice = st.selectbox(
                "Bank Account",
                ["ACC-8652", "ACC-3738", "Other"]
            )

            if account_choice == "自定义账号":
                account_num = st.text_input("Enter Account", value="ACC-")
            else:
                account_num = account_choice

            selected_stage = st.text_input("Stage", value="Move-out")
            default_memo = "Deposit Refund"

        company_display = st.text_input("Company", value=company_name)

        st.markdown("---")

        payee_name = st.text_input(
            "Payee Name",
            value="John Smith",
        )
        pay_amount = st.number_input(
            "Amount", min_value=0.01, value=1500.00, step=100.0
        )

        c_a, c_b = st.columns(2)
        with c_a:
            pay_date = st.date_input("Date", value=date.today())
        with c_b:
            check_num = st.number_input(
                "Check Number",
                min_value=1,
                value=int(default_chk_val),
                step=1,
                help="Automatically generated by System"
            )

        user_input = st.text_input(
            "Memo", 
            value=""
        )
        
        # 2. 最终使用的完整 Memo 文本
        memo_text = f"{project_site} [{selected_stage}] - {user_input}" if user_input else f"{project_site} [{selected_stage}]"
        amount_words = number_to_words_usd(pay_amount)

    replacements = {
        "date": pay_date.strftime("%m/%d/%Y"),
        "name": payee_name,
        "amount": f"{pay_amount:,.2f}",
        "amount_words": amount_words,
        "memo": memo_text,
        "number": str(check_num),
        "account": account_num,
    }
    
    # 提前填充 PDF 用于预览与下载 (修复 original 代码中的 NameError)
    filled_pdf = fill_pdf_placeholders(pdf_template_bytes, replacements)

    with col2:
        st.subheader("👁️ Check Preview")
        st.markdown(f"""
        > **Company**: {company_name}  
        > **Bank Account**: `{account_num}`  
        > **Stage**: {project_site} | **`{selected_stage}`**  
        > **Check Number**: `#{check_num}`  
        > **Date**: {pay_date.strftime("%Y-%m-%d")}  
        > **Payee**: **{payee_name}**  
        > **Amount**: **${pay_amount:,.2f}**  
        > **Amount Words**: *{amount_words}*  
        > **Memo**: {memo_text}
        """)

        if st.button("🚀 Successfully transfer to Google Sheets", type="primary", use_container_width=True):
            record = [
                {
                    "Check Number": check_num,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_display,
                    "Account": account_num,
                    "Project": project_site,
                    "Stage": selected_stage,
                    "Payee Name": payee_name,
                    "Amount": pay_amount,
                    "Memo": memo_text,
                }
            ]
            if save_to_history(record):
                st.balloons()
                st.success(f"🎉 Check #{check_num} Successfully transfer to Google Sheets！")

        st.download_button(
            label=f"📥 Download PDF (#{check_num})",
            data=filled_pdf,
            file_name=f"Check_{check_num}_{payee_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

# ==============================================================================
# 场景 2：多项目/施工队周薪批量开单
# ==============================================================================
elif mode == "👷 Construction Bulk Checks":
    st.title("👷 Construction Bulk Checks")

    if not pdf_template_bytes:
        st.stop()

    pay_date = st.date_input("Date", value=date.today())

    st.markdown("---")
    
    # ----------------- 1. 确认各项目起始支票号 -----------------
    st.subheader("1. Confirm the start check number")

    proj_start_nums = {}
    cols = st.columns(min(len(df_projects), 4))
    for idx, p_row in df_projects.iterrows():
        p_name = p_row["Project_Name"]
        default_num = latest_check_map.get(p_name, 1001)
        with cols[idx % 4]:
            proj_start_nums[p_name] = st.number_input(
                f"🏗️ {p_name}",
                min_value=1,
                value=int(default_num),
                key=f"start_num_{p_name}"
            )

    st.markdown("---")

    # ----------------- 2. 初始化发薪数据列表 -----------------
    st.subheader("2.Information")

    if "payroll_list" not in st.session_state:
        st.session_state.payroll_list = []

    def update_memo_on_worker_change():
        selected_w = st.session_state.input_w
        st.session_state.input_m = worker_role_map.get(selected_w, "")

    if "input_m" not in st.session_state:
        default_first_worker = preset_worker_list[0] if preset_worker_list else ""
        st.session_state.input_m = worker_role_map.get(default_first_worker, "")

    # --- 快捷添加面板 ---
    st.markdown("##### ➕ New Check")
    c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 1.5, 2, 1.2])

    with c1:
        add_worker = st.selectbox(
            "Payee", 
            preset_worker_list, 
            key="input_w",
            on_change=update_memo_on_worker_change
        )
    with c2:
        add_proj = st.selectbox("Project", preset_project_list, key="input_p")
    with c3:
        add_stage = st.selectbox("Stage", PRESET_STAGES, key="input_s")
        stage_val = add_stage.split(":")[0].strip()
    with c4:
        add_amt = st.number_input("Amount", min_value=0.01, value=1200.00, step=50.0, key="input_a")
    with c5:
        add_memo = st.text_input("Memo", key="input_m")
    with c6:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Add", type="primary", use_container_width=True):
            existing_checks = [
                row["Check #"] 
                for row in st.session_state.payroll_list 
                if row["Project"] == add_proj
            ]
            
            if existing_checks:
                next_chk = max(existing_checks) + 1
            else:
                next_chk = proj_start_nums.get(add_proj, 1001)

            st.session_state.payroll_list.append({
                "Payee": add_worker,
                "Project": add_proj,
                "Stage": stage_val,
                "Check #": int(next_chk),
                "Amount": add_amt,
                "Memo": add_memo
            })
            st.rerun()

    st.markdown("---")

    # --- 数据列表展示 ---
    if st.session_state.payroll_list:
        st.markdown(f"##### 📋 List of Pending Checks ({len(st.session_state.payroll_list)} total):")
        
        df_current = pd.DataFrame(st.session_state.payroll_list)

        edited_df = st.data_editor(
            df_current,
            num_rows="dynamic",
            use_container_width=True,
            key="payroll_table_editor",
            column_config={
                "Payee": st.column_config.SelectboxColumn("Payee", options=preset_worker_list),
                "Project": st.column_config.SelectboxColumn("Project", options=preset_project_list),
                "Stage": st.column_config.TextColumn("Stage"),
                "Check #": st.column_config.NumberColumn("Check #", format="%d"),
                "Amount": st.column_config.NumberColumn("Amount", format="$%.2f")
            }
        )

        st.session_state.payroll_list = edited_df.to_dict(orient="records")

        if st.button("🗑️ Clear", type="secondary"):
            st.session_state.payroll_list = []
            st.rerun()

        df_payroll_input = edited_df
    else:
        st.info("💡 No data in the list. Please select a worker and project above, then click **[➕ Add]** to add team members.")
        df_payroll_input = pd.DataFrame()

    st.markdown("---")

    # ----------------- 3. 批量生成与导出 -----------------
    if not df_payroll_input.empty:
        if st.button(f"🚀 Confirm & Batch Generate {len(df_payroll_input)} Checks", type="primary", use_container_width=True):
            account_pdf_dict = {}
            records_log = []

            proj_map = df_projects.set_index("Project_Name").to_dict(orient="index")

            for idx, row in df_payroll_input.iterrows():
                worker_name = str(row.get("Payee", "")).strip()
                project_name = str(row.get("Project", "")).strip()
                stage_name = str(row.get("Stage", "")).strip()
                
                try:
                    cur_check = int(row.get("Check #", 0))
                    amt = float(row.get("Amount", 0.0))
                except (ValueError, TypeError):
                    cur_check = 0
                    amt = 0.0

                detail_memo = str(row.get("Memo", "")).strip()

                if amt <= 0 or not worker_name or cur_check <= 0:
                    continue

                p_info = proj_map.get(project_name, {"Company": "Unknown Company", "Account": "ACC-0000"})
                company_name = p_info["Company"]
                account_num = p_info["Account"]

                # 构造包含 Stage 的完整 Memo 文本
                stage_prefix = f"[{stage_name}] " if stage_name else ""
                full_memo = f"{project_name} {stage_prefix}- {detail_memo}" if detail_memo else f"{project_name} {stage_prefix}"

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
                    "Stage": stage_name,
                    "Payee Name": worker_name,
                    "Amount": amt,
                    "Memo": full_memo
                })

            if records_log:
                # 发送至云端 Google Sheets
                if save_to_history(records_log):
                    st.session_state.payroll_list = []

                    st.balloons()
                    st.success(f"🎉 Successfully generated {len(records_log)} check(s)! Data synced to Google Sheets.")

                    st.markdown("### 📊 Current Period Disbursement Summary")
                    df_batch = pd.DataFrame(records_log)
                    col_sum1, col_sum2 = st.columns(2)

                    with col_sum1:
                        st.markdown("#### 🏢 Summary by Company / Account")
                        summary_company = df_batch.groupby(["Company", "Account"]).agg(
                            **{
                                "Total Amount": ("Amount", "sum"),
                                "Total Number": ("Check Number", "count")
                            }
                        ).reset_index()
                        st.dataframe(summary_company.style.format({"Total Amount": "${:,.2f}"}), use_container_width=True, hide_index=True)

                    with col_sum2:
                        st.markdown("#### 🏗️ Summary by Project")
                        summary_project = df_batch.groupby(["Project", "Company"]).agg(
                            **{
                                "Total Labor Cost": ("Amount", "sum"),
                                "Worker Count": ("Check Number", "count")
                            }
                        ).reset_index()
                        st.dataframe(summary_project.style.format({"Total Labor Cost": "${:,.2f}"}), use_container_width=True, hide_index=True)

                    st.markdown("---")
                    st.markdown("### 📥 Download PDFs by Account")

                    for (comp_name, acc_num), item_list in account_pdf_dict.items():
                        pdf_bytes_list = [item[3] for item in item_list]
                        account_merged_pdf = merge_pdfs(pdf_bytes_list)
                        
                        st.markdown(f"##### 💳 Account: **{comp_name}** | No.: `{acc_num}` ({len(item_list)} check(s) total)")
                        
                        st.download_button(
                            label=f"📄 Download Merged PDF ({comp_name} - {acc_num})",
                            data=account_merged_pdf,
                            file_name=f"Checks_{comp_name}_{acc_num}_{pay_date}.pdf",
                            mime="application/pdf",
                            type="primary"
                        )

                    st.markdown("---")
                    st.markdown("##### 📦 More Export Options")
                    
                    csv_bytes = df_batch.to_csv(index=False).encode('utf-8-sig')
                    
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for (comp_name, acc_num), item_list in account_pdf_dict.items():
                            for chk, proj, py, pdf_b in item_list:
                                zf.writestr(f"[{acc_num}]_Check_{chk}_[{proj}]_{py}.pdf", pdf_b)

                    d1, d2 = st.columns(2)
                    with d1:
                        st.download_button(
                            label="📊 Download Payroll Summary (CSV)",
                            data=csv_bytes,
                            file_name=f"Payroll_Summary_{pay_date}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                    with d2:
                        st.download_button(
                            label="📦 Download All Individual PDFs (ZIP)",
                            data=zip_buf.getvalue(),
                            file_name=f"Payroll_Checks_SinglePDFs_{pay_date}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
