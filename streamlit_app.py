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

# ----------------- 配置文件与路径定义 (纯 CSV 版，免 openpyxl) -----------------
DEFAULT_TEMPLATE_PATH = "check_run.pdf"
LOG_FILE = "check_issuance_history.csv"
PROJECTS_CSV = "projects_config.csv"
WORKERS_CSV = "workers_config.csv"


# ----------------- 初始化/读取配置文件 -----------------
def load_project_presets():
    """读取或创建公司与工地项目对应表"""
    if not os.path.exists(PROJECTS_CSV):
        df_default = pd.DataFrame(
            [
                {
                    "Project_Name": "123 Main St",
                    "Company": "AAA Construction Inc",
                    "Account": "ACC-8652",
                },
                {
                    "Project_Name": "456 Oak Ave",
                    "Company": "BBB Development LLC",
                    "Account": "ACC-3738",
                },
                {
                    "Project_Name": "789 Pine Rd",
                    "Company": "CCC Management Group",
                    "Account": "ACC-9901",
                },
            ]
        )
        df_default.to_csv(PROJECTS_CSV, index=False)

    return pd.read_csv(PROJECTS_CSV)


def load_worker_presets():
    """读取或创建常用工人列表"""
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
    """把数字金额转换为支票的标准大写英文"""
    try:
        dollars = int(amount)
        cents = int(round((amount - dollars) * 100))
        words = num2words(dollars, lang="en").title()
        return f"{words} and {cents:02d}/100 Dollars"
    except Exception:
        return ""


def fill_pdf_placeholders(pdf_bytes, replacements):
    """查找 PDF 中的 {{ key }} 占位符并替换"""
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


def get_next_check_number():
    """读取 CSV 历史台账，自动推荐下一个支票号"""
    if os.path.exists(LOG_FILE):
        try:
            df_log = pd.read_csv(LOG_FILE)
            if not df_log.empty and "Check Number" in df_log.columns:
                valid_nums = pd.to_numeric(
                    df_log["Check Number"], errors="coerce"
                ).dropna()
                if not valid_nums.empty:
                    return int(valid_nums.max()) + 1
        except Exception:
            pass
    return 1001


def save_to_history(records):
    """追加写入历史发纸台账 CSV (无需 openpyxl)"""
    df_new = pd.DataFrame(records)
    if os.path.exists(LOG_FILE):
        try:
            df_new.to_csv(LOG_FILE, mode="a", index=False, header=False)
            return
        except Exception:
            pass
    df_new.to_csv(LOG_FILE, index=False)


def merge_pdfs(pdf_bytes_list):
    """将多个单页支票 PDF 合并为一个 PDF"""
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

# ----------------- 页面架构与全局导航 -----------------
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
    st.caption(
        "适合临时开单、个人报销或临时供应商付款，可自动联动选择项目与公司。"
    )

    if not pdf_template_bytes:
        st.error("请先上传 PDF 模板！")
        st.stop()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 基础信息填报")

        # 工地项目与公司账户联动选择
        project_options = preset_project_list + ["+ 自定义新项目"]
        selected_proj = st.selectbox("选择工地/项目 (Project)", project_options)

        if selected_proj != "+ 自定义新项目":
            p_info = df_projects[
                df_projects["Project_Name"] == selected_proj
            ].iloc[0]
            default_company = p_info["Company"]
            default_account = p_info["Account"]
            project_site = selected_proj
        else:
            project_site = st.text_input("输入新项目名称", value="New Site")
            default_company = "AAA Construction Inc"
            default_account = "ACC-8652"

        company_name = st.text_input("付款公司名称", value=default_company)
        account_num = st.text_input("付款账号", value=default_account)

        st.markdown("---")

        payee_name = st.text_input(
            "收款人 (Payee Name)",
            value="John Smith",
            help="输入个人或公司名称",
        )
        pay_amount = st.number_input(
            "金额 $ (Amount)", min_value=0.01, value=1500.00, step=100.0
        )

        c_a, c_b = st.columns(2)
        with c_a:
            pay_date = st.date_input("开票日期", value=date.today())
        with c_b:
            check_num = st.number_input(
                "支票编号", min_value=1, value=get_next_check_number()
            )

        memo_text = st.text_input(
            "备注 (Memo)", value=f"{project_site} - Material Fee"
        )

        # 动态大写金额实时计算展示
        amount_words = number_to_words_usd(pay_amount)
        st.info(f"🔤 **英文金额大写预览：**\n\n`{amount_words}`")

    # 组合占位符字典
    replacements = {
        "date": pay_date.strftime("%m/%d/%Y"),
        "name": payee_name,
        "amount": f"{pay_amount:,.2f}",
        "amount_words": amount_words,
        "memo": memo_text,
        "number": str(check_num),
        "account": account_num,
    }

    # 右侧：实时预览与生成下载
    with col2:
        st.subheader("2. 实时生成与预览")

        # 实时渲染填充后的 PDF
        filled_pdf = fill_pdf_placeholders(pdf_template_bytes, replacements)

        # 将 PDF 转为 Base64 在 Streamlit 中原生内嵌预览
        base64_pdf = base64.b64encode(filled_pdf).decode("utf-8")
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="450" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 生成与保存按钮
        if st.button("🚀 确认生成并记录台账", type="primary", use_container_width=True):
            # 记录台账
            record = [
                {
                    "Check Number": check_num,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_name,
                    "Account": account_num,
                    "Project": project_site,
                    "Payee Name": payee_name,
                    "Amount": pay_amount,
                    "Memo": memo_text,
                }
            ]
            save_to_history(record)
            st.balloons()
            st.success(
                f"🎉 支票 #{check_num} 已成功生成并已写入历史记录！"
            )

            # 单张下载按钮
            st.download_button(
                label=f"📥 下载支票 PDF (#{check_num})",
                data=filled_pdf,
                file_name=f"Check_{check_num}_{payee_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


# ==============================================================================
# 场景 2：多项目混合周薪批量开单
# ==============================================================================
elif mode == "👷 场景二：多项目/施工队周薪批量开单":
    st.title("👷 多项目/施工队周薪批量生成")
    st.caption(
        "支持一次性录入多个不同工地的工人薪资，自动匹配公司账号并输出分类统计。"
    )

    if not pdf_template_bytes:
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        pay_date = st.date_input("发薪日期", value=date.today())
    with c2:
        start_check = st.number_input(
            "起始支票编号", min_value=1, value=get_next_check_number()
        )

    st.markdown("---")
    st.subheader("1. 录入发薪明细（可混合选择不同项目）")

    input_type = st.radio(
        "数据录入模式：",
        ["在线表格快捷录入", "上传包含 Project 列的 Excel/CSV"],
        horizontal=True,
    )

    df_payroll_input = pd.DataFrame()

    if input_type == "在线表格快捷录入":
        default_proj = (
            preset_project_list[0] if preset_project_list else "123 Main St"
        )
        init_data = [
            {
                "工人姓名 (Payee)": preset_worker_list[0]
                if preset_worker_list
                else "",
                "所属项目 (Project)": default_proj,
                "金额 $ (Amount)": 1200.00,
                "工作备注 (Memo)": "Weekly Work",
            },
            {
                "工人姓名 (Payee)": preset_worker_list[1]
                if len(preset_worker_list) > 1
                else "",
                "所属项目 (Project)": preset_project_list[1]
                if len(preset_project_list) > 1
                else default_proj,
                "金额 $ (Amount)": 950.00,
                "工作备注 (Memo)": "Weekly Work",
            },
        ]
        df_init = pd.DataFrame(init_data)

        df_payroll_input = st.data_editor(
            df_init,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "工人姓名 (Payee)": st.column_config.SelectboxColumn(
                    "工人姓名 (Payee)",
                    options=preset_worker_list,
                    help="可直接选择或双击输入",
                ),
                "所属项目 (Project)": st.column_config.SelectboxColumn(
                    "所属项目 (Project)",
                    options=preset_project_list,
                    help="选择对应的工地项目（系统会自动关联出对应的公司账号）",
                    required=True,
                ),
            },
        )
    else:
        uploaded_file = st.file_uploader(
            "上传批量发薪表 (必须包含: Payee, Project, Amount 三列)",
            type=["csv", "xlsx"],
        )
        if uploaded_file:
            df_payroll_input = (
                pd.read_csv(uploaded_file)
                if uploaded_file.name.endswith(".csv")
                else pd.read_excel(uploaded_file)
            )
            st.dataframe(df_payroll_input, use_container_width=True)

    st.markdown("---")

    if not df_payroll_input.empty:
        if st.button(
            "🚀 批量生成所有项目的支票与统计表",
            type="primary",
            use_container_width=True,
        ):
            generated_pdfs = []
            records_log = []
            cur_check = start_check

            proj_map = df_projects.set_index("Project_Name").to_dict(
                orient="index"
            )

            for idx, row in df_payroll_input.iterrows():
                worker_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
                project_name = str(row.iloc[1]).strip() if len(row) > 1 else ""
                try:
                    amt = float(row.iloc[2])
                except (ValueError, TypeError):
                    amt = 0.0
                detail_memo = str(row.iloc[3]).strip() if len(row) > 3 else ""

                if amt <= 0 or not worker_name:
                    continue

                p_info = proj_map.get(
                    project_name,
                    {"Company": "Unknown Company", "Account": "ACC-0000"},
                )
                company_name = p_info["Company"]
                account_num = p_info["Account"]

                full_memo = (
                    f"{project_name} - {detail_memo}"
                    if detail_memo
                    else project_name
                )

                replacements = {
                    "date": pay_date.strftime("%m/%d/%Y"),
                    "name": worker_name,
                    "amount": f"{amt:,.2f}",
                    "amount_words": number_to_words_usd(amt),
                    "memo": full_memo,
                    "number": str(cur_check),
                    "account": account_num,
                }

                pdf_res = fill_pdf_placeholders(
                    pdf_template_bytes, replacements
                )
                generated_pdfs.append(
                    (cur_check, project_name, worker_name, pdf_res)
                )

                records_log.append(
                    {
                        "Check Number": cur_check,
                        "Issue Date": pay_date.strftime("%Y-%m-%d"),
                        "Company": company_name,
                        "Account": account_num,
                        "Project": project_name,
                        "Payee Name": worker_name,
                        "Amount": amt,
                        "Memo": full_memo,
                    }
                )

                cur_check += 1

            if generated_pdfs:
                save_to_history(records_log)
                st.balloons()
                st.success(
                    f"🎉 成功生成 {len(generated_pdfs)} 张支票！编号: #{start_check} ~ #{cur_check - 1}"
                )

                # 统计汇总区
                st.markdown("### 📊 本期跨项目出账汇总")
                df_batch = pd.DataFrame(records_log)

                col_sum1, col_sum2 = st.columns(2)

                with col_sum1:
                    st.markdown("#### 🏢 1. 按公司 / 账号出账小计")
                    summary_company = (
                        df_batch.groupby(["Company", "Account"])
                        .agg(
                            总金额=("Amount", "sum"),
                            支票张数=("Check Number", "count"),
                        )
                        .reset_index()
                    )
                    st.dataframe(
                        summary_company.style.format({"总金额": "${:,.2f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

                with col_sum2:
                    st.markdown("#### 🏗️ 2. 按工地项目出账小计")
                    summary_project = (
                        df_batch.groupby(["Project", "Company"])
                        .agg(
                            项目总人工费=("Amount", "sum"),
                            工人人数=("Check Number", "count"),
                        )
                        .reset_index()
                    )
                    st.dataframe(
                        summary_project.style.format(
                            {"项目总人工费": "${:,.2f}"}
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

                st.markdown("---")
                st.markdown("### 📥 导出与下载区")

                csv_bytes = df_batch.to_csv(index=False).encode("utf-8-sig")
                merged_pdf_bytes = merge_pdfs([p[3] for p in generated_pdfs])

                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button(
                        label="📄 下载【全项目合并 PDF】(连续打印)",
                        data=merged_pdf_bytes,
                        file_name=f"Payroll_Checks_AllProjects_{pay_date}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                    )

                with d2:
                    st.download_button(
                        label="📊 下载【本期出账总账单 CSV】",
                        data=csv_bytes,
                        file_name=f"Payroll_Summary_{pay_date}.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

                with d3:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for chk, proj, py, pdf_b in generated_pdfs:
                            zf.writestr(
                                f"Check_{chk}_[{proj}]_{py}.pdf", pdf_b
                            )

                    st.download_button(
                        label="📦 下载按项目命名的 ZIP 压缩包",
                        data=zip_buf.getvalue(),
                        file_name=f"Payroll_Checks_{pay_date}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
