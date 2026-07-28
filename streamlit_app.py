import base64
from datetime import date, datetime
import io
import os
import zipfile
import fitz
from num2words import num2words  # 用于将数字转换为英文大写
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Check Generator System", page_icon="🧾", layout="wide"
)

# ----------------- 配置文件与常量 -----------------
DEFAULT_TEMPLATE_PATH = "check_run.pdf"
LOG_FILE = "check_issuance_history.xlsx"

# 预设公司与账号映射
COMPANY_PRESETS = {
    "8652": "ACC-8652",
    "3738": "ACC-3738",
}


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
    """查找 PDF 中的 {{ key }} 占位符并替换为实际数据"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page in doc:
        for key, val in replacements.items():
            str_val = str(val) if val is not None else ""

            # 兼容带有不同空格数量的占位符模式
            patterns = [
                f"{{{{ {key} }}}}",
                f"{{{{{key}}}}}",
                f"{{{{  {key}  }}}}",
                f"{{{{   {key}   }}}}",
            ]

            for pattern in patterns:
                rects = page.search_for(pattern)
                for rect in rects:
                    # 擦除原来的 {{ xxx }} 占位符文字
                    page.add_redact_annot(rect, fill=(1, 1, 1))
                    page.apply_redactions()

                    # 写入新数据
                    point = fitz.Point(rect.x0, rect.y1 - 2)
                    page.insert_text(
                        point, str_val, fontsize=10, color=(0, 0, 0)
                    )

    output_stream = io.BytesIO()
    doc.save(output_stream)
    doc.close()
    return output_stream.getvalue()


def display_pdf_preview(pdf_bytes):
    """在 Streamlit 页面中嵌入 PDF 实时预览"""
    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


def get_next_check_number():
    """读取历史台账，自动推荐下一个支票号"""
    if os.path.exists(LOG_FILE):
        try:
            df_log = pd.read_excel(LOG_FILE)
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
    """追加写入历史发纸台账 Excel"""
    df_new = pd.DataFrame(records)
    if os.path.exists(LOG_FILE):
        try:
            df_old = pd.read_excel(LOG_FILE)
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
            df_combined.to_excel(LOG_FILE, index=False)
            return
        except Exception:
            pass
    df_new.to_excel(LOG_FILE, index=False)


def merge_pdfs(pdf_bytes_list):
    """将多个单页支票 PDF 合并为一个大的 PDF (极度方便直接打印机连续打印)"""
    merged_doc = fitz.open()
    for b in pdf_bytes_list:
        doc = fitz.open(stream=b, filetype="pdf")
        merged_doc.insert_pdf(doc)
        doc.close()

    out_stream = io.BytesIO()
    merged_doc.save(out_stream)
    merged_doc.close()
    return out_stream.getvalue()


# ----------------- 模板文件检测 -----------------
pdf_template_bytes = None
if os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        pdf_template_bytes = f.read()

# ----------------- 侧边栏：场景导航与模板 -----------------
st.sidebar.title("⚙️ 系统导航")
mode = st.sidebar.radio(
    "请选择业务场景：",
    [
        "📝 场景一：单张手动生成",
        "👷 场景二：施工队/工人周薪批量生成",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("📄 PDF 模板状态")
if pdf_template_bytes:
    st.sidebar.success(f"已加载后台模板: `{DEFAULT_TEMPLATE_PATH}`")
else:
    st.sidebar.warning(f"未找到 `{DEFAULT_TEMPLATE_PATH}`，请上传：")
    uploaded_tpl = st.sidebar.file_uploader("上传支票模板 PDF", type=["pdf"])
    if uploaded_tpl:
        pdf_template_bytes = uploaded_tpl.read()


# ==============================================================================
# 场景 1：单张手动生成
# ==============================================================================
if mode == "📝 场景一：单张手动生成":
    st.title("📝 单张支票手动开具")

    if not pdf_template_bytes:
        st.error("请先在左侧栏上传 PDF 模板！")
        st.stop()

    col1, col2 = st.columns([1, 1.2])

    with col1:
        st.subheader("1. 填写支票信息")
        input_date = st.date_input("日期 (date)", value=date.today())
        date_str = input_date.strftime("%m/%d/%Y")

        name = st.text_input("收款人 (name)", value="")

        col_amt1, col_amt2 = st.columns([1, 1])
        with col_amt1:
            amount_num = st.number_input(
                "金额 $ (amount)",
                min_value=0.0,
                value=0.00,
                step=0.01,
                format="%.2f",
            )
            amount_str = f"{amount_num:,.2f}"

        auto_words = number_to_words_usd(amount_num)
        amount_words = st.text_input(
            "金额大写 (amount_words)", value=auto_words
        )

        col_chk1, col_chk2 = st.columns([1, 1])
        with col_chk1:
            check_number = st.text_input(
                "支票编号 (number)", value=str(get_next_check_number())
            )
        with col_chk2:
            selected_option = st.selectbox(
                "选择账号预设",
                options=["8652", "3738", "Other"],
                index=0,
            )
            if selected_option in COMPANY_PRESETS:
                account = COMPANY_PRESETS[selected_option]
            else:
                account = st.text_input("手动输入账号", value="ACC-883921")

        memo = st.text_area("Memo (备注)", value="Payroll")

        replacements = {
            "date": date_str,
            "name": name,
            "amount": amount_str,
            "amount_words": amount_words,
            "memo": memo,
            "number": check_number,
            "account": account,
        }

    with col2:
        st.subheader("2. 实时生成与预览")
        try:
            filled_pdf = fill_pdf_placeholders(pdf_template_bytes, replacements)
            display_pdf_preview(filled_pdf)

            st.download_button(
                label="📥 点击下载当前支票 PDF",
                data=filled_pdf,
                file_name=f"Check_{check_number}_{name}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"处理 PDF 时出错: {str(e)}")


# ==============================================================================
# 场景 2：施工队/工人每周发薪
# ==============================================================================
elif mode == "👷 场景二：施工队/工人周薪批量生成":
    st.title("👷 施工队/工人每周发薪")
    st.caption(
        "支持多公司/多项目快速发薪，支持上传周薪 Excel 或直接在界面表格中新增工人。"
    )

    if not pdf_template_bytes:
        st.error("请先在左侧栏上传 PDF 模板！")
        st.stop()

    # 第一步：基本信息选择
    st.subheader("1. 付款账户与工地项目")
    c1, c2, c3 = st.columns(3)
    with c1:
        company_opt = st.selectbox(
            "选择付款公司/账号预设",
            options=["8652", "3738", "Other"],
            index=0,
        )
        if company_opt in COMPANY_PRESETS:
            selected_account = COMPANY_PRESETS[company_opt]
        else:
            selected_account = st.text_input(
                "手动输入付款公司账号", value="ACC-9901"
            )
    with c2:
        project_site = st.text_input(
            "工地/项目名称 (Project Site)", value="123 Main St"
        )
    with c3:
        start_check = st.number_input(
            "起始支票编号",
            min_value=1,
            value=get_next_check_number(),
        )

    st.markdown("---")
    st.subheader("2. 录入工人发薪名单")

    input_type = st.radio(
        "数据录入模式：",
        ["手动表格录入/修改", "上传周薪 Excel/CSV 文件"],
        horizontal=True,
    )

    df_workers = pd.DataFrame()

    if input_type == "手动表格录入/修改":
        default_workers = pd.DataFrame(
            [
                {
                    "工人姓名 (Payee)": "John Smith",
                    "工资金额 (Amount)": 1200.00,
                    "工作内容 (Memo)": " Framing",
                },
                {
                    "工人姓名 (Payee)": "Carlos Mendez",
                    "工资金额 (Amount)": 950.50,
                    "工作内容 (Memo)": " Drywall",
                },
            ]
        )
        df_workers = st.data_editor(
            default_workers, num_rows="dynamic", use_container_width=True
        )

    else:
        uploaded_payroll = st.file_uploader(
            "上传工人薪资表 (Excel / CSV)", type=["xlsx", "csv"]
        )
        if uploaded_payroll:
            if uploaded_payroll.name.endswith(".csv"):
                df_payroll_raw = pd.read_csv(uploaded_payroll)
            else:
                df_payroll_raw = pd.read_excel(uploaded_payroll)

            st.write("上传结果预览（可在线微调）：")
            df_workers = st.data_editor(
                df_payroll_raw, num_rows="dynamic", use_container_width=True
            )

    st.markdown("---")

    # 第三步：批量生成与导出
    if not df_workers.empty:
        st.subheader("3. 批量生成支票")
        pay_date = st.date_input(
            "发薪日期", value=date.today(), key="payroll_date"
        )

        if st.button(
            "🚀 批量生成施工队支票", type="primary", use_container_width=True
        ):
            generated_pdfs = []
            records_log = []
            cur_check = start_check

            for idx, row in df_workers.iterrows():
                worker_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
                try:
                    amt = float(row.iloc[1])
                except (ValueError, TypeError):
                    amt = 0.0
                detail_memo = str(row.iloc[2]).strip() if len(row) > 2 else ""

                if amt <= 0 or not worker_name:
                    continue

                # 自动拼接 Memo：[项目名称] + 工作内容
                full_memo = (
                    f"{project_site} - {detail_memo}"
                    if detail_memo
                    else project_site
                )

                replacements = {
                    "date": pay_date.strftime("%m/%d/%Y"),
                    "name": worker_name,
                    "amount": f"{amt:,.2f}",
                    "amount_words": number_to_words_usd(amt),
                    "memo": full_memo,
                    "number": str(cur_check),
                    "account": selected_account,
                }

                pdf_res = fill_pdf_placeholders(
                    pdf_template_bytes, replacements
                )
                generated_pdfs.append((cur_check, worker_name, pdf_res))

                records_log.append(
                    {
                        "Check Number": cur_check,
                        "Issue Date": pay_date.strftime("%Y-%m-%d"),
                        "Category": f"Contractor Payroll ({project_site})",
                        "Payee Name": worker_name,
                        "Amount": amt,
                        "Account": selected_account,
                        "Memo": full_memo,
                    }
                )

                cur_check += 1

            if generated_pdfs:
                save_to_history(records_log)
                st.balloons()
                st.success(
                    f"🎉 成功生成 {len(generated_pdfs)} 张工人发薪支票！支票号: #{start_check} ~ #{cur_check - 1}"
                )

                # 提供合并 PDF & ZIP 下载
                merged_pdf_bytes = merge_pdfs([p[2] for p in generated_pdfs])

                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    st.download_button(
                        label="📄 下载【合并打印版 PDF】(推荐给打印机连续打印)",
                        data=merged_pdf_bytes,
                        file_name=f"Payroll_Checks_{project_site}_{pay_date}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                    )

                with d_col2:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for chk, py, pdf_b in generated_pdfs:
                            zf.writestr(
                                f"Check_{chk}_{project_site}_{py}.pdf", pdf_b
                            )

                    st.download_button(
                        label="📦 下载单张 ZIP 压缩包",
                        data=zip_buf.getvalue(),
                        file_name=f"Payroll_Checks_ZIP_{project_site}_{pay_date}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
