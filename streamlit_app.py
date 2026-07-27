import streamlit as st
from io import BytesIO
# import pypandoc
from num2words import num2words  # 用于将数字转换为英文大写
from datetime import datetime
import os
from PyPDF2 import PdfReader, PdfWriter


st.set_page_config(
    page_title="Check Generator", page_icon="🧾", layout="wide"
)


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

                    # 写入新数据（调整 baseline 基线对齐）
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
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="750" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


# ----------------- Streamlit 界面排版 -----------------

st.title("🧾 Checks Generator")

# 1. 加载默认模板文件
DEFAULT_TEMPLATE_PATH = "check_run.pdf"
pdf_bytes = None

col1, col2 = st.columns([1, 1.2])

with col1:
    st.subheader("1. Import Info")

    # 判断本地是否存在 check_run.pdf
    if os.path.exists(DEFAULT_TEMPLATE_PATH):
        st.success(f"✅ 已成功读取后台模板：`{DEFAULT_TEMPLATE_PATH}`")
        with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
            pdf_bytes = f.read()
    else:
        st.warning(
            f"⚠️ 未在同级目录下找到 `{DEFAULT_TEMPLATE_PATH}`，请手动上传："
        )
        uploaded_file = st.file_uploader(
            "上传 PDF 模板", type=["pdf"], key="check_uploader"
        )
        if uploaded_file:
            pdf_bytes = uploaded_file.read()

    st.markdown("---")

    # 变量输入框
    input_date = st.date_input("日期 (date)", value=date.today())
    date_str = input_date.strftime("%m/%d/%Y")

    name = st.text_input("收款人 (name)", value="John Doe")

    col_amt1, col_amt2 = st.columns([1, 1])
    with col_amt1:
        amount_num = st.number_input(
            "金额 $ (amount)",
            min_value=0.0,
            value=1250.50,
            step=0.01,
            format="%.2f",
        )
        amount_str = f"${amount_num:,.2f}"

    # 自动算出大写英文，也可手动微调
    auto_words = number_to_words_usd(amount_num)
    amount_words = st.text_input(
        "金额大写 (amount_words)", value=auto_words
    )

    col_chk1, col_chk2 = st.columns([1, 1])
    with col_chk1:
        check_number = st.text_input("支票编号 (number)", value="1001")
    with col_chk2:
        account = st.text_input("账号 (account)", value="ACC-883921")

    memo = st.text_area("备忘/用途 (memo)", value="Invoice Payment")

    # 建立对应你 PDF 里的变量映射字典
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

    if pdf_bytes is not None:
        try:
            # 替换占位符并生成新的 PDF
            filled_pdf = fill_pdf_placeholders(pdf_bytes, replacements)

            st.download_button(
                label="📥 点击下载生成好的支票 PDF",
                data=filled_pdf,
                file_name=f"Check_{check_number}_{date_str.replace('/', '-')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )

            st.markdown("---")
            display_pdf_preview(filled_pdf)

        except Exception as e:
            st.error(f"处理 PDF 时出错: {str(e)}")
    else:
        st.info("💡 请确保 `check_run.pdf` 放在项目根目录下。")
