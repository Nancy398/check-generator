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

# ----------------- 2. 读取 Google Sheets (你的原始写法) -----------------
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
    """把生成的支票记录直接追加保存到 Google Sheets 中"""
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
    """利用你的 read_file 读取数据，计算各个项目最新的 Check Number (+1)"""
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
        "📝 场景一：单张手动生成",
        "👷 场景二：多项目/施工队周薪批量开单",
    ],
)


if pdf_template_bytes is None:
    st.warning("请先在左侧上传 PDF 模板文件。")
    uploaded_tpl = st.sidebar.file_uploader("上传支票模板", type=["pdf"])
    if uploaded_tpl:
        pdf_template_bytes = uploaded_tpl.read()

# ==============================================================================
# 场景 1：单张手动生成支票
# ==============================================================================
if mode == "📝 场景一：单张手动生成":
    st.title("📝 场景一：单张手动生成支票")
    st.caption("选择业务主体，自行直接输入支票编号与相关信息。")

    if not pdf_template_bytes:
        st.stop()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 基础信息填报")

        biz_mode = st.radio(
            "选择业务主体 (Business Entity)：",
            ["🏗️ Moo Construction (施工/项目模式)", "🏠 Moo Housing Inc (房屋租赁/退押金模式)"],
            horizontal=True,
        )

        st.markdown("---")

        if "Construction" in biz_mode:
            company_name = "Moo Construction Inc"
            
            project_options = preset_project_list + ["+ 自定义新项目"]
            selected_proj = st.selectbox("选择工地/项目 (Project)", project_options)

            if selected_proj != "+ 自定义新项目":
                p_info = df_projects[df_projects["Project_Name"] == selected_proj].iloc[0]
                default_account = p_info["Account"]
                project_site = selected_proj
                default_chk_val = latest_check_map.get(selected_proj, 1001)
            else:
                project_site = st.text_input("输入新项目名称", value="New Site")
                default_account = "ACC-8652"
                default_chk_val = 1001

            account_num = st.text_input("付款账号", value=default_account)
            default_memo = f"{project_site} - Labor Fee"

        else:
            company_name = "Moo Housing Inc"
            
            project_options = preset_project_list + ["+ 自定义新项目"]
            selected_proj = st.selectbox("关联房产/项目 (Project)", project_options)

            if selected_proj != "+ 自定义新项目":
                project_site = selected_proj
                default_chk_val = latest_check_map.get(selected_proj, 1001)
            else:
                project_site = st.text_input("输入房产名称/地址", value="Moo Housing Property")
                default_chk_val = 1001

            account_choice = st.selectbox(
                "选择付款账号 (Account)",
                ["ACC-8652", "ACC-3738", "Other (自定义账号)"]
            )

            if account_choice == "Other (自定义账号)":
                account_num = st.text_input("输入自定义账号", value="ACC-")
            else:
                account_num = account_choice

            default_memo = "Deposit Refund"

        company_display = st.text_input("付款公司名称", value=company_name)

        st.markdown("---")

        payee_name = st.text_input(
            "收款人 (Payee Name)",
            value="John Smith",
        )
        pay_amount = st.number_input(
            "金额 $ (Amount)", min_value=0.01, value=1500.00, step=100.0
        )

        c_a, c_b = st.columns(2)
        with c_a:
            pay_date = st.date_input("开票日期", value=date.today())
        with c_b:
            check_num = st.number_input(
                "支票编号 (Check Number)",
                min_value=1,
                value=int(default_chk_val),
                step=1,
                help="系统已从 Google Sheets 读取最新可用编号，可自由更改"
            )

        memo_text = st.text_input(
            "备注 (Memo)", value=default_memo
        )
        amount_words = number_to_words_usd(pay_amount)

        # st.info(f"🔤 **英文金额大写预览：**\n\n`{amount_words}`")

    replacements = {
        "date": pay_date.strftime("%m/%d/%Y"),
        "name": payee_name,
        "amount": f"{pay_amount:,.2f}",
        "amount_words": amount_words,
        "memo": memo_text,
        "number": str(check_num),
        "account": account_num,
    }

    with col2:
        st.subheader("👁️ 支票信息预览")
        st.markdown(f"""
        > **公司名称**: {company_name}  
        > **银行账号**: `{account_num}`  
        > **支票编号**: `#{check_num}`  
        > **开单日期**: {pay_date.strftime("%Y-%m-%d")}  
        > **收款人**: **{payee_name}**  
        > **金额**: **${pay_amount:,.2f}**  
        > **金额大写**: *{amount_words}*  
        > **Memo**: {selected_proj} - {memo_text}
        """)

        if st.button("🚀 确认生成并推送至 Google Sheets", type="primary", use_container_width=True):
            record = [
                {
                    "Check Number": check_num,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_display,
                    "Account": account_num,
                    "Project": project_site,
                    "Payee Name": payee_name,
                    "Amount": pay_amount,
                    "Memo": memo_text,
                }
            ]
            if save_to_history(record):
                st.balloons()
                st.success(f"🎉 支票 #{check_num} 已成功生成并写入云端 Google Sheets！")

            st.download_button(
                label=f"📥 下载支票 PDF (#{check_num})",
                data=filled_pdf,
                file_name=f"Check_{check_num}_{payee_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

# ==============================================================================
# 场景 2：多项目/施工队周薪批量开单
# ==============================================================================
elif mode == "👷 场景二：多项目/施工队周薪批量开单":
    st.title("👷 多项目/施工队周薪批量生成")
    st.caption("从云端推算项目起始号，快捷录入并自动同步至 Google Sheets。")

    if not pdf_template_bytes:
        st.stop()

    pay_date = st.date_input("发薪日期", value=date.today())

    st.markdown("---")
    
    # ----------------- 1. 确认各项目起始支票号 -----------------
    st.subheader("1. 确认本期各项目起始支票号")

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

    # ----------------- 3. 批量生成与导出 -----------------
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
                # 仅发送至云端 Google Sheets
                if save_to_history(records_log):
                    st.session_state.payroll_list = []

                    st.balloons()
                    st.success(f"🎉 成功生成 {len(records_log)} 张支票！数据已同步至 Google Sheets。")

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
                        
                        st.markdown(f"##### 💳 账户：**{comp_name}** | 账号：`{acc_num}`（共 {len(item_list)} 张支票）")
                        
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
