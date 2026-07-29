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

# ----------------- 页面配置 -----------------
st.set_page_config(
    page_title="Check Generator System", page_icon="🧾", layout="wide"
)

# ----------------- 配置文件与路径定义 -----------------
DEFAULT_TEMPLATE_PATH = "check_run.pdf"

GS_SPREADSHEET_NAME = "Check Issuance History"  # Google 表格的名字
GS_WORKSHEET_NAME = "Sheet1"                  # 历史记录工作表的名字

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
    try:
        gc = get_gc_client()
        worksheet = gc.open(name).worksheet(sheet)
        rows = worksheet.get_all_values()
        if not rows or len(rows) <= 1:
            return pd.DataFrame()
        df = pd.DataFrame.from_records(rows)
        df = pd.DataFrame(df.values[1:], columns=df.iloc[0])
        # 清除列名两端空格
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"❌ 读取 Google Sheets ({sheet}) 失败: {e}")
        return pd.DataFrame()

# ----------------- 3. 保存写入 Google Sheets -----------------
def save_to_history(records):
    """把生成的支票记录直接追加保存到 Google Sheets 中"""
    try:
        gc = get_gc_client()
        worksheet = gc.open(GS_SPREADSHEET_NAME).worksheet(GS_WORKSHEET_NAME)
        
        rows_to_append = []
        for r in records:
            rows_to_append.append([
                r["Check Number"],
                r["Issue Date"],
                r["Company"],
                r["Account"],
                r["Project"],
                r.get("Stage", ""),
                r["Payee Name"],
                r["Amount"],
                r["Memo"]
            ])
        worksheet.append_rows(rows_to_append)
        return True
    except Exception as e:
        st.error(f"❌ 数据保存至 Google Sheets 失败: {e}")
        return False

# ----------------- 4. 计算最新可用支票号 -----------------
def fetch_next_check_numbers_from_gs(df_p_list, default_start_number=1001):
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

    for p_name in df_p_list:
        if p_name not in next_numbers:
            next_numbers[p_name] = default_start_number
            
    return next_numbers

# ----------------- 5. 数据预设加载函数 -----------------
@st.cache_data(ttl=60)
def load_project_presets():
    df_p = read_file(GS_SPREADSHEET_NAME, "Project")
    if df_p.empty or "Project_Name" not in df_p.columns:
        st.warning("⚠️ Google Sheets 中 'Project' 表格为空或缺失 'Project_Name' 列！")
        return pd.DataFrame(columns=["Project_Name", "Company", "Account"])
    return df_p

@st.cache_data(ttl=60)
def load_worker_presets():
    """读取 Worker 表格，提取 Worker_Name, Stage, Stage_Name, Sub_Stage"""
    df_w = read_file(GS_SPREADSHEET_NAME, "Worker")
    required_cols = {"Worker_Name", "Stage", "Stage_Name", "Sub_Stage"}
    if df_w.empty or not required_cols.issubset(df_w.columns):
        st.warning("⚠️ Worker 表格为空或缺失 'Worker_Name', 'Stage', 'Stage_Name', 'Sub_Stage' 列！")
        return pd.DataFrame(columns=["Worker_Name", "Stage", "Stage_Name", "Sub_Stage"])
    return df_w

@st.cache_data(ttl=60)
def load_stage_presets():
    """读取 Stage 表格配置"""
    df_s = read_file(GS_SPREADSHEET_NAME, "Stage")
    required_cols = {"Stage", "Stage_Name", "Sub_Stage"}
    if df_s.empty or not required_cols.issubset(df_s.columns):
        st.warning("⚠️ Stage 表格为空或缺少 'Stage', 'Stage_Name', 'Sub_Stage' 列！")
        return pd.DataFrame(columns=["Stage", "Stage_Name", "Sub_Stage"])
    return df_s

# 加载云端预设数据
df_projects = load_project_presets()
df_workers = load_worker_presets()
df_stages = load_stage_presets()

preset_worker_list = df_workers["Worker_Name"].dropna().tolist() if not df_workers.empty else []
preset_project_list = df_projects["Project_Name"].dropna().tolist() if not df_projects.empty else []

latest_check_map = fetch_next_check_numbers_from_gs(preset_project_list)

# ----------------- 通用 Stage 联动选择组件 -----------------
def render_stage_selector(key_prefix="default", default_stage="", default_stage_name="", default_sub_stage=""):
    """
    嵌入式的 Stage 三级联动选择组件
    返回: (selected_stage_code, selected_stage_name, sub_stage_1, sub_stage_2)
    """
    if df_stages.empty:
        st.warning("Stage 配置数据为空")
        return "", "", "", None

    # 构造显示的选项：例 "Stage - 1: Pre-construction & Demo"
    df_stages_temp = df_stages.copy()
    df_stages_temp["Display_Stage"] = df_stages_temp["Stage"] + ": " + df_stages_temp["Stage_Name"]
    unique_stages = df_stages_temp["Display_Stage"].unique().tolist()

    # 寻找默认 Stage 的索引位置
    target_display = f"{default_stage}: {default_stage_name}"
    stage_idx = unique_stages.index(target_display) if target_display in unique_stages else 0

    col_s1, col_s2, col_s3 = st.columns([2, 2, 2])

    with col_s1:
        selected_stage_display = st.selectbox(
            "Stage (主阶段)",
            options=unique_stages,
            index=stage_idx,
            key=f"{key_prefix}_main_stage"
        )

    # 筛选子类别
    filtered_df = df_stages_temp[df_stages_temp["Display_Stage"] == selected_stage_display]
    sub_stage_options = filtered_df["Sub_Stage"].dropna().tolist()

    sub1_idx = sub_stage_options.index(default_sub_stage) if default_sub_stage in sub_stage_options else 0

    with col_s2:
        sub_stage_1 = st.selectbox(
            "Sub-Stage 1 (必选)",
            options=sub_stage_options,
            index=sub1_idx,
            key=f"{key_prefix}_sub_stage_1"
        )

    with col_s3:
        remaining_options = ["None (无)"] + [item for item in sub_stage_options if item != sub_stage_1]
        sub_stage_2 = st.selectbox(
            "Sub-Stage 2 (可选)",
            options=remaining_options,
            index=0,
            key=f"{key_prefix}_sub_stage_2"
        )

    selected_stage_code = filtered_df["Stage"].iloc[0]
    selected_stage_name = filtered_df["Stage_Name"].iloc[0]
    opt_sub_stage_2 = sub_stage_2 if sub_stage_2 != "None (无)" else None

    return selected_stage_code, selected_stage_name, sub_stage_1, opt_sub_stage_2

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
            account_num = st.text_input("Enter Account", value="ACC-") if account_choice == "Other" else account_choice

        company_display = st.text_input("Company", value=company_name)

        st.markdown("---")

        # ---- 收款人选择与 Worker 表属性自动匹配 ----
        payee_name = st.selectbox("Payee Name", options=preset_worker_list) if preset_worker_list else st.text_input("Payee Name", value="John Smith")

        # 根据选中的收款人查找 Worker 表预设
        def_stg, def_stg_name, def_sub_stg = "", "", ""
        if not df_workers.empty and payee_name in df_workers["Worker_Name"].values:
            w_info = df_workers[df_workers["Worker_Name"] == payee_name].iloc[0]
            def_stg = w_info.get("Stage", "")
            def_stg_name = w_info.get("Stage_Name", "")
            def_sub_stg = w_info.get("Sub_Stage", "")

        # 调整为：Stage_Name - Sub_Stage
        default_memo_text = f"{def_stg_name} - {def_sub_stg}"

        st.markdown("##### 🏗️ 工程阶段 (Stage Selection)")
        st_code, st_name, sub1, sub2 = render_stage_selector(
            key_prefix="single_check",
            default_stage=def_stg,
            default_stage_name=def_stg_name,
            default_sub_stage=def_sub_stg
        )

        user_memo = st.text_input("Memo Detail", value=default_memo_text, help="自动预填工人活计，可手动更改")

        pay_amount = st.number_input("Amount", min_value=0.01, value=1500.00, step=100.0)

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

        # 组合 Stage 标记文本
        sub_str = f"{sub1}, {sub2}" if sub2 else sub1
        selected_stage_str = f"{st_code} ({sub_str})" if st_code else ""

        # 统一清理并组合最终 Memo 文本
        memo_components = [project_site]
        if selected_stage_str:
            memo_components.append(f"[{selected_stage_str}]")
        if user_memo.strip():
            memo_components.append(f"- {user_memo.strip()}")
            
        memo_text = " ".join(memo_components)
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
    
    filled_pdf = fill_pdf_placeholders(pdf_template_bytes, replacements)

    with col2:
        st.subheader("👁️ Check Preview")
        st.markdown(f"""
        > **Company**: {company_display}  
        > **Bank Account**: `{account_num}`  
        > **Project / Stage**: {project_site} | **`{selected_stage_str}`**  
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
                    "Stage": selected_stage_str,
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
    cols = st.columns(min(max(len(df_projects), 1), 4))
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

    # ----------------- 2. 快捷添加面板 -----------------
    st.subheader("2. Information")

    if "payroll_list" not in st.session_state:
        st.session_state.payroll_list = []

    # 更换 Worker 时自动同步该工人的默认活计 Memo (调整为 Stage_Name - Sub_Stage)
    def update_memo_on_worker_change():
        selected_w = st.session_state.get("input_w", "")
        if not df_workers.empty and selected_w in df_workers["Worker_Name"].values:
            w_info = df_workers[df_workers["Worker_Name"] == selected_w].iloc[0]
            stg_name = w_info.get("Stage_Name", "")
            sub_stg = w_info.get("Sub_Stage", "")
            st.session_state.input_m = f"{stg_name} - {sub_stg}" if sub_stg else stg_name

    def calculate_amount_from_days():
        days = st.session_state.get("input_days", 0.0)
        rate = st.session_state.get("input_rate", 0.0)
        if days > 0 and rate > 0:
            st.session_state.input_a = round(days * rate, 2)

    # 初始化默认 Worker 与 Memo (调整为 Stage_Name - Sub_Stage)
    if "input_m" not in st.session_state and preset_worker_list:
        first_w = preset_worker_list[0]
        if not df_workers.empty and first_w in df_workers["Worker_Name"].values:
            w_info = df_workers[df_workers["Worker_Name"] == first_w].iloc[0]
            st.session_state.input_m = f"{w_info.get('Stage_Name', '')} - {w_info.get('Sub_Stage', '')}"

    st.markdown("##### ➕ New Check")
    
    # 第一行：项目与人员基本信息
    r1_c1, r1_c2 = st.columns([1, 1])
    with r1_c1:
        add_worker = st.selectbox(
            "Payee Name", 
            preset_worker_list, 
            key="input_w",
            on_change=update_memo_on_worker_change
        )
    with r1_c2:
        add_proj = st.selectbox("Project", preset_project_list, key="input_p")

    # 查找工人的默认 Stage 配置用于联动定位
    w_stg, w_stg_name, w_sub_stg = "", "", ""
    if not df_workers.empty and add_worker in df_workers["Worker_Name"].values:
        w_info = df_workers[df_workers["Worker_Name"] == add_worker].iloc[0]
        w_stg = w_info.get("Stage", "")
        w_stg_name = w_info.get("Stage_Name", "")
        w_sub_stg = w_info.get("Sub_Stage", "")

    # 嵌入式 Stage 联动选择
    st_code, st_name, sub1, sub2 = render_stage_selector(
        key_prefix="bulk_check",
        default_stage=w_stg,
        default_stage_name=w_stg_name,
        default_sub_stage=w_sub_stg
    )

    # 第二行：金额计算、Memo 与提交按钮
    r2_c1, r2_c2, r2_c3, r2_c4, r2_c5 = st.columns([1.5, 1.5, 2.0, 3.0, 1.2])
    with r2_c1:
        add_days = st.number_input(
            "Days (Opt)", 
            min_value=0.0, 
            value=0.0, 
            step=0.5, 
            key="input_days",
            on_change=calculate_amount_from_days
        )
    with r2_c2:
        add_rate = st.number_input(
            "Rate/Day (Opt)", 
            min_value=0.0, 
            value=0.0, 
            step=10.0, 
            key="input_rate",
            on_change=calculate_amount_from_days
        )
    with r2_c3:
        add_amt = st.number_input("Total Amount", min_value=0.01, value=1200.00, step=50.0, key="input_a")
    with r2_c4:
        add_memo = st.text_input("Memo Detail", key="input_m")
    with r2_c5:
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

            stage_val_str = f"{st_code} ({f'{sub1}, {sub2}' if sub2 else sub1})" if st_code else ""

            st.session_state.payroll_list.append({
                "Payee": add_worker,
                "Project": add_proj,
                "Stage": stage_val_str,
                "Days": add_days if add_days > 0 else None,
                "Rate": add_rate if add_rate > 0 else None,
                "Check #": int(next_chk),
                "Amount": add_amt,
                "Memo": add_memo
            })
            st.rerun()

    st.markdown("---")

    # --- 3. 数据列表展示 ---
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
                "Days": st.column_config.NumberColumn("Days", format="%.1f days"),
                "Rate": st.column_config.NumberColumn("Rate/Day", format="$%.2f"),
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
    
    # ----------------- 4. 批量生成与导出 -----------------
    if not df_payroll_input.empty:
        if st.button(f"🚀 Confirm & Batch Generate {len(df_payroll_input)} Checks", type="primary", use_container_width=True):
            account_pdf_dict = {}
            records_log = []
    
            proj_map = df_projects.set_index("Project_Name").to_dict(orient="index") if not df_projects.empty else {}
    
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
    
                p_info = proj_map.get(project_name, {"Company": "Moo Construction", "Account": "ACC-8652"})
                company_name = p_info["Company"]
                account_num = p_info["Account"]
    
                # 只组合 Project 和 Detail Memo (不加 Stage 标签)
                if project_name and detail_memo:
                    full_memo = f"{project_name} - {detail_memo}"
                elif project_name:
                    full_memo = project_name
                else:
                    full_memo = detail_memo
    
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
