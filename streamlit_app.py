import base64
from datetime import date
import io
import os
import zipfile
import fitz
from num2words import num2words
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Check Generator System", page_icon="🧾", layout="wide"
)

# ----------------- 配置文件与路径定义 -----------------
DEFAULT_TEMPLATE_PATH = "check_run.pdf"
LOG_FILE = "check_issuance_history.csv"
PROJECTS_CSV = "projects_config.csv"
WORKERS_CSV = "workers_config.csv"


# ----------------- 初始化/读取配置文件 -----------------
def load_project_presets():
    """读取或创建公司与工地项目对应表 (包含独立的支票起始号)"""
    if not os.path.exists(PROJECTS_CSV):
        df_default = pd.DataFrame(
            [
                {
                    "Project_Name": "123 Main St",
                    "Company": "Moo Construction Inc",
                    "Account": "ACC-8652",
                    "Next_Check_Number": 1001,
                },
                {
                    "Project_Name": "456 Oak Ave",
                    "Company": "Moo Construction Inc",
                    "Account": "ACC-3738",
                    "Next_Check_Number": 5001,
                },
                {
                    "Project_Name": "789 Pine Rd",
                    "Company": "Moo Housing Inc",
                    "Account": "ACC-8652",
                    "Next_Check_Number": 8001,
                },
            ]
        )
        df_default.to_csv(PROJECTS_CSV, index=False)

    df_p = pd.read_csv(PROJECTS_CSV)
    # 确保 Next_Check_Number 列存在
    if "Next_Check_Number" not in df_p.columns:
        df_p["Next_Check_Number"] = 1001
        df_p.to_csv(PROJECTS_CSV, index=False)
    return df_p


def save_project_presets(df_p):
    """更新保存项目的支票号状态"""
    df_p.to_csv(PROJECTS_CSV, index=False)


def load_worker_presets():
    """读取常用工人列表"""
    if not os.path.exists(WORKERS_CSV):
        df_default = pd.DataFrame(
            [
                {"Worker_Name": "John Smith"},
                {"Worker_Name": "Carlos Mendez"},
                {"Worker_Name": "David Lee"},
                {"Worker_Name": "Jose Rodriguez"},
            ]
        )
        df_default.to_csv(WORKERS_CSV, index=False)

    df_workers = pd.read_csv(WORKERS_CSV)
    return df_workers["Worker_Name"].dropna().tolist()


df_projects = load_project_presets()
preset_worker_list = load_worker_presets()
preset_project_list = df_projects["Project_Name"].dropna().tolist()


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


def save_to_history(records):
    """追加写入历史发票台账 CSV"""
    df_new = pd.DataFrame(records)
    if os.path.exists(LOG_FILE):
        try:
            df_new.to_csv(LOG_FILE, mode="a", index=False, header=False)
            return
        except Exception:
            pass
    df_new.to_csv(LOG_FILE, index=False)


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
# 场景 1：单张手动生成支票（带 Moo Housing / Construction 模式切换与独立支票号）
# ==============================================================================
if mode == "📝 场景一：单张手动生成":
    st.title("📝 场景一：单张手动生成支票")
    st.caption("选择业务主体，自动适配账号选项与默认备注，并自动带出选定项目独立的下一张支票号。")

    if not pdf_template_bytes:
        st.stop()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 基础信息填报")

        # ----------------- 模式选择 -----------------
        biz_mode = st.radio(
            "选择业务主体 (Business Entity)：",
            ["🏗️ Moo Construction (施工/项目模式)", "🏠 Moo Housing Inc (房屋租赁/退押金模式)"],
            horizontal=True,
        )

        st.markdown("---")

        # ----------------- 分支逻辑处理 -----------------
        if "Construction" in biz_mode:
            # === Construction 模式：原逻辑（按项目联动） ===
            company_name = "Moo Construction Inc"
            
            project_options = preset_project_list + ["+ 自定义新项目"]
            selected_proj = st.selectbox("选择工地/项目 (Project)", project_options)

            if selected_proj != "+ 自定义新项目":
                p_info = df_projects[df_projects["Project_Name"] == selected_proj].iloc[0]
                default_account = p_info["Account"]
                default_check_num = int(p_info["Next_Check_Number"])
                project_site = selected_proj
            else:
                project_site = st.text_input("输入新项目名称", value="New Site")
                default_account = "ACC-8652"
                default_check_num = 1001

            account_num = st.text_input("付款账号", value=default_account)
            default_memo = f"{project_site} - Material Fee"

        else:
            # === Moo Housing Inc 模式：退押金专享选项 ===
            company_name = "Moo Housing Inc"

            # 账号选择：8652 / 3738 / Other
            account_choice = st.selectbox(
                "选择付款账号 (Account)",
                ["ACC-8652", "ACC-3738", "Other (自定义账号)"]
            )

            if account_choice == "Other (自定义账号)":
                account_num = st.text_input("输入自定义账号", value="ACC-")
            else:
                account_num = account_choice

            # 默认 Memo 为 Deposit Refund
            default_memo = "Deposit Refund"

        # ----------------- 通用信息填写 -----------------
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
                value=1001,  # 默认起始值，可随意修改
                step=1,
                help="请在此处直接输入本次开票的支票号码"
            )

        memo_text = st.text_input(
            "备注 (Memo)", value=default_memo
        )
        amount_words = number_to_words_usd(pay_amount)

        st.info(f"🔤 **英文金额大写预览：**\n\n`{amount_words}`")

    # 替换占位符字典
    replacements = {
        "date": pay_date.strftime("%m/%d/%Y"),
        "name": payee_name,
        "amount": f"{pay_amount:,.2f}",
        "amount_words": amount_words,
        "memo": memo_text,
        "number": str(check_num),
        "account": account_num,
    }

    # 右侧预览与保存
    with col2:
        st.subheader("2. 实时生成与预览")
        filled_pdf = fill_pdf_placeholders(pdf_template_bytes, replacements)

        base64_pdf = base64.b64encode(filled_pdf).decode("utf-8")
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="480" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🚀 确认生成并记录台账", type="primary", use_container_width=True):
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
            save_to_history(record)

            # 更新选定项目的下一张支票号递增状态
            if selected_proj != "+ 自定义新项目":
                df_projects.loc[
                    df_projects["Project_Name"] == selected_proj,
                    "Next_Check_Number",
                ] = (check_num + 1)
                save_project_presets(df_projects)

            st.balloons()
            st.success(
                f"🎉 支票 #{check_num} 已成功生成并写入历史台账！"
            )

            st.download_button(
                label=f"📥 下载支票 PDF (#{check_num})",
                data=filled_pdf,
                file_name=f"Check_{check_num}_{payee_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

# ==============================================================================
# 场景 2：多项目混合周薪批量开单（支持核对起始号 + 表格自动推算/手动修改支票号）
# ==============================================================================
elif mode == "👷 场景二：多项目/施工队周薪批量开单":
    st.title("👷 多项目/施工队周薪批量生成")
    st.caption("确认各项目起始号后，表格将自动推算支票号；如有特殊跳号可直接在表格中修改。")

    if not pdf_template_bytes:
        st.stop()

    pay_date = st.date_input("发薪日期", value=date.today())

    st.markdown("---")
    
    # ----------------- 环节 1：确认各项目起始支票号 -----------------
    st.subheader("1. 确认本期各项目起始支票号")
    st.caption("系统已根据上次记录自动拉取，如有变更（如领了新支票本），请在此处调整：")

    proj_start_nums = {}
    cols = st.columns(min(len(df_projects), 4))  # 动态分列展示
    for idx, p_row in df_projects.iterrows():
        p_name = p_row["Project_Name"]
        default_num = int(p_row["Next_Check_Number"])
        with cols[idx % 4]:
            proj_start_nums[p_name] = st.number_input(
                f"🏗️ {p_name}",
                min_value=1,
                value=default_num,
                key=f"start_num_{p_name}"
            )

    st.markdown("---")

    # ----------------- 环节 2：录入发薪明细（自动分配并支持手动改号） -----------------
    st.subheader("2. 录入发薪明细（支票号已自动分配，支持手动修改）")

    # 预设首行数据并自动推算支票号
    default_p1 = preset_project_list[0] if preset_project_list else "123 Main St"
    default_p2 = preset_project_list[1] if len(preset_project_list) > 1 else default_p1

    # 根据顶部确认的起始号，实时初始化首行数据
    p1_start = proj_start_nums.get(default_p1, 1001)
    p2_start = proj_start_nums.get(default_p2, 5001) if default_p2 != default_p1 else p1_start + 1

    init_data = [
        {
            "工人姓名 (Payee)": preset_worker_list[0] if preset_worker_list else "",
            "所属项目 (Project)": default_p1,
            "支票编号 (Check #)": int(p1_start),
            "金额 $ (Amount)": 1200.00,
            "工作备注 (Memo)": "Weekly Work"
        },
        {
            "工人姓名 (Payee)": preset_worker_list[1] if len(preset_worker_list) > 1 else "",
            "所属项目 (Project)": default_p2,
            "支票编号 (Check #)": int(p2_start),
            "金额 $ (Amount)": 950.00,
            "工作备注 (Memo)": "Weekly Work"
        }
    ]
    df_init = pd.DataFrame(init_data)

    st.info("💡 提示：增加行或切换项目时，请检查/调整【支票编号】列。如果某张支票纸废掉需跳号，直接双击修改数字即可。")

    df_payroll_input = st.data_editor(
        df_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "工人姓名 (Payee)": st.column_config.SelectboxColumn(
                "工人姓名 (Payee)",
                options=preset_worker_list,
            ),
            "所属项目 (Project)": st.column_config.SelectboxColumn(
                "所属项目 (Project)",
                options=preset_project_list,
                required=True,
            ),
            "支票编号 (Check #)": st.column_config.NumberColumn(
                "支票编号 (Check #)",
                help="自动推算出的支票号，允许手动修改跳号",
                required=True,
                format="%d"
            ),
            "金额 $ (Amount)": st.column_config.NumberColumn(
                "金额 $ (Amount)",
                format="$%.2f"
            )
        },
    )

    st.markdown("---")

    # ----------------- 环节 3：批量生成与更新 -----------------
    if not df_payroll_input.empty:
        if st.button("🚀 确认无误，批量生成支票", type="primary", use_container_width=True):
            generated_pdfs = []
            records_log = []
            max_check_used = {} # 用于记录每个项目本次用到的最大支票号

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

                # 填充 PDF
                replacements = {
                    "date": pay_date.strftime("%m/%d/%Y"),
                    "name": worker_name,
                    "amount": f"{amt:,.2f}",
                    "amount_words": number_to_words_usd(amt),
                    "memo": full_memo,
                    "number": str(cur_check), # 直接取表格中最终确认/手动修改后的支票号
                    "account": account_num
                }

                pdf_res = fill_pdf_placeholders(pdf_template_bytes, replacements)
                generated_pdfs.append((cur_check, project_name, worker_name, pdf_res))

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

                # 追溯该项目使用的最大支票号，以便更新下一次的起始号
                if project_name not in max_check_used or cur_check > max_check_used[project_name]:
                    max_check_used[project_name] = cur_check

            if generated_pdfs:
                # 1. 保存流水台账
                save_to_history(records_log)

                # 2. 将每个项目用到的（最大支票号 + 1）更新回配置文件
                for p_name, max_num in max_check_used.items():
                    df_projects.loc[df_projects["Project_Name"] == p_name, "Next_Check_Number"] = max_num + 1
                save_project_presets(df_projects)

                st.balloons()
                st.success(f"🎉 成功生成 {len(generated_pdfs)} 张支票！各项目下次起始号已自动更新为最大使用号 + 1。")

                # 统计看板与导出区
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
                st.markdown("### 📥 导出与下载")

                csv_bytes = df_batch.to_csv(index=False).encode('utf-8-sig')
                merged_pdf_bytes = merge_pdfs([p[3] for p in generated_pdfs])

                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button(
                        label="📄 下载【全项目合并 PDF】",
                        data=merged_pdf_bytes,
                        file_name=f"Payroll_Checks_AllProjects_{pay_date}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )

                with d2:
                    st.download_button(
                        label="📊 下载【本期出账总账单 CSV】",
                        data=csv_bytes,
                        file_name=f"Payroll_Summary_{pay_date}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

                with d3:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for chk, proj, py, pdf_b in generated_pdfs:
                            zf.writestr(f"Check_{chk}_[{proj}]_{py}.pdf", pdf_b)

                    st.download_button(
                        label="📦 下载单张 ZIP 包",
                        data=zip_buf.getvalue(),
                        file_name=f"Payroll_Checks_{pay_date}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
