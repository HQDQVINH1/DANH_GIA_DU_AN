# app.py
import re
import io
import math
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd
import numpy as np

from docx import Document

# Nếu numpy_financial không có, ta tự hiện IRR; pip install numpy_financial nếu muốn dùng
try:
    import numpy_financial as nf
    has_nf = True
except Exception:
    has_nf = False

# --- Cấu hình Trang Streamlit ---
st.set_page_config(page_title="App Phân Tích Phương Án Kinh Doanh", layout="wide")
st.title("App Phân Tích Phương Án Kinh Doanh từ file Word")

# --- Utility: đọc text từ file .docx ---
def read_docx(file_bytes: io.BytesIO) -> str:
    doc = Document(file_bytes)
    paragraphs = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(paragraphs)

# --- Utility: chuyển chuỗi số có dấu phẩy, dấu chấm sang float ---
def parse_number(text: str) -> Optional[float]:
    if text is None:
        return None
    txt = text.strip()
    if txt == "":
        return None
    # Loại bỏ ký tự không phải số, comma, dot, minus, %
    # Nếu có %, trả về phân số
    is_percent = "%" in txt
    txt = txt.replace("%", "")
    # Thay các dấu không cần thiết
    # Xử lý trường hợp có đơn vị VND, triệu, tỷ
    multiplier = 1.0
    if re.search(r'\b(tỷ|ty|billion)\b', txt, flags=re.I):
        multiplier = 1e9
        txt = re.sub(r'\b(tỷ|ty|billion)\b', '', txt, flags=re.I)
    elif re.search(r'\b(triệu|trieu|million)\b', txt, flags=re.I):
        multiplier = 1e6
        txt = re.sub(r'\b(triệu|trieu|million)\b', '', txt, flags=re.I)
    # Remove non numeric except dot and comma and minus
    txt = re.sub(r'[^\d\.,\-]', '', txt)
    # If both comma and dot present, assume comma thousands, dot decimal if dot after comma or vice versa
    if txt.count(',') > 0 and txt.count('.') > 0:
        # heuristic: if last '.' occurs after last ',' then dot is decimal
        if txt.rfind('.') > txt.rfind(','):
            txt = txt.replace(',', '')
        else:
            txt = txt.replace('.', '').replace(',', '.')
    else:
        # If only comma present and groups of 3, treat comma as thousands
        if txt.count(',') > 0 and re.search(r'\d+,\d{3}(,|\.)', txt) is not None:
            txt = txt.replace(',', '')
        else:
            txt = txt.replace(',', '.')
    try:
        val = float(txt)
        if is_percent:
            return val / 100.0
        return val * multiplier
    except Exception:
        return None

# --- Hàm lọc thông tin quan trọng từ văn bản bằng phương pháp Regex + heuristic ---
def extract_project_info(text: str) -> Dict[str, Any]:
    lower = text.lower()
    info = {
        "Vốn đầu tư": None,
        "Dòng đời dự án (năm)": None,
        "Doanh thu hàng năm": None,   # có thể là list hoặc số trung bình
        "Chi phí hàng năm": None,
        "WACC": None,
        "Thuế suất": None,
        "Ghi chú thô": text[:500]  # lưu một đoạn tóm tắt
    }

    # Patterns tìm kiếm các mục phổ biến
    patterns = {
        "Vốn đầu tư": r'(vốn đầu tư|đầu tư ban đầu|initial investment|initial cost)[\s\:\-–]*([^\n\r,;]+)',
        "Dòng đời dự án (năm)": r'(dòng đời|thời gian khai thác|thời gian dự án|lifetime|project life)[^\d\n\r]*?(\d{1,2})\s*(năm|year)?',
        "WACC": r'(wacc|chi phí vốn|weighted average cost of capital)[\s\:\-–]*([^\n\r,;]+)',
        "Thuế suất": r'(thuế suất|tax rate|tax)[\s\:\-–]*([^\n\r,;]+%)',
        # Doanh thu/chi phí có thể xuất hiện nhiều nơi; lấy các câu có "doanh thu" hoặc "doanh số" và "chi phí"
        "Doanh thu hàng năm": r'(doanh thu|revenue|sales)[\s\:\-–]*([^\n\r,;]+)',
        "Chi phí hàng năm": r'(chi phí|costs|expenditure)[\s\:\-–]*([^\n\r,;]+)'
    }

    for key, pat in patterns.items():
        m = re.search(pat, text, flags=re.I)
        if m:
            # giá trị thường ở nhóm 2
            token = m.groups()[-1].strip()
            num = parse_number(token)
            if key == "Dòng đời dự án (năm)" and num is None:
                # có thể token là chữ số nhưng parse_number không bắt, thử lấy int
                try:
                    num = int(re.search(r'(\d{1,2})', token).group(1))
                except Exception:
                    num = None
            info[key] = num if num is not None else token

    # Nếu không tìm thấy WACC dạng %, thử tìm số có % sau "chi phí vốn"
    if info["WACC"] is None:
        m = re.search(r'chi phí vốn[\s\:\-–]*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            w = parse_number(m.group(1) + "%")
            info["WACC"] = w

    # Thuế suất nếu ở dạng số không có %:
    if info["Thuế suất"] is None:
        m = re.search(r'thuế\s*:\s*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            info["Thuế suất"] = parse_number(m.group(1) + "%")

    # Nếu doanh thu/chi phí xuất hiện nhiều lần, gom các con số xuất hiện gần từ khóa
    revenues = []
    costs = []
    for line in text.splitlines():
        if re.search(r'\b(doanh thu|revenue|sales)\b', line, flags=re.I):
            # tìm các số trong dòng
            for numstr in re.findall(r'[\d\.,]+\s*(tỷ|ty|triệu|trieu)?\b', line, flags=re.I):
                n = parse_number(numstr)
                if n:
                    revenues.append(n)
        if re.search(r'\b(chi phí|cost|costs|expenditure)\b', line, flags=re.I):
            for numstr in re.findall(r'[\d\.,]+\s*(tỷ|ty|triệu|trieu)?\b', line, flags=re.I):
                n = parse_number(numstr)
                if n:
                    costs.append(n)

    if revenues and info["Doanh thu hàng năm"] is None:
        # lấy trung bình nếu có nhiều năm trình bày
        info["Doanh thu hàng năm"] = sum(revenues) / len(revenues)
    if costs and info["Chi phí hàng năm"] is None:
        info["Chi phí hàng năm"] = sum(costs) / len(costs)

    # Try: nếu vốn đầu tư chưa có, thử tìm các cụm "tổng vốn" hoặc "capex"
    if info["Vốn đầu tư"] is None:
        m = re.search(r'(tổng vốn|tổng đầu tư|capex|capital\s*expenditure)[\s\:\-–]*([^\n\r,;]+)', text, flags=re.I)
        if m:
            info["Vốn đầu tư"] = parse_number(m.groups()[-1])

    # Nếu vẫn thiếu Dòng đời, để mặc định 5 năm (giả định hợp lý)
    if info["Dòng đời dự án (năm)"] is None:
        info["Dòng đời dự án (năm)"] = 5

    # Nếu WACC chưa có, dùng mặc định 10%
    if info["WACC"] is None:
        info["WACC"] = 0.10

    # Nếu Thuế suất chưa có, dùng 20%
    if info["Thuế suất"] is None:
        info["Thuế suất"] = 0.20

    return info

# --- Xây bảng dòng tiền dự án từ thông tin đã lọc ---
def build_cashflow_table(info: Dict[str, Any]) -> pd.DataFrame:
    life = int(info.get("Dòng đời dự án (năm)", 5))
    capex = info.get("Vốn đầu tư", 0) or 0
    # Nếu Doanh thu/Chi phí là list hoặc số: dùng as is; nếu None -> 0
    revenue = info.get("Doanh thu hàng năm", 0) or 0
    cost = info.get("Chi phí hàng năm", 0) or 0
    tax = info.get("Thuế suất", 0.2) or 0.2

    # Giả định: Doanh thu và chi phí ổn định qua các năm trừ năm đầu có CapEx
    rows = []
    for year in range(0, life + 1):
        if year == 0:
            # Dòng tiền đầu tư ban đầu âm
            rows.append({
                "Năm": year,
                "Doanh thu": 0.0,
                "Chi phí": 0.0,
                "EBIT": 0.0,
                "Thuế": 0.0,
                "NOPAT": 0.0,
                "CapEx": -abs(capex),
                "Dòng tiền ròng": -abs(capex)
            })
        else:
            ebit = revenue - cost
            tax_amt = ebit * tax if ebit > 0 else 0.0
            nopat = ebit - tax_amt
            cf = nopat  # giả định không có thay đổi vốn lưu động, khấu hao không tách riêng
            rows.append({
                "Năm": year,
                "Doanh thu": revenue,
                "Chi phí": cost,
                "EBIT": ebit,
                "Thuế": tax_amt,
                "NOPAT": nopat,
                "CapEx": 0.0,
                "Dòng tiền ròng": cf
            })
    df = pd.DataFrame(rows)
    return df

# --- Tính các chỉ số tài chính ---
def compute_project_metrics(cashflows: List[float], discount_rate: float) -> Dict[str, Any]:
    # cashflows: list với năm 0..N (năm0 âm)
    # NPV
    npv = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cashflows))
    # IRR
    irr = None
    try:
        if has_nf:
            irr = nf.irr(cashflows)
        else:
            # Newton-Raphson đơn giản cho IRR
            def npv_func(r):
                return sum(cf / ((1 + r) ** i) for i, cf in enumerate(cashflows))
            def npv_derivative(r):
                return sum(-i * cf / ((1 + r) ** (i + 1)) for i, cf in enumerate(cashflows))
            r = 0.1
            for _ in range(100):
                f = npv_func(r)
                df_r = npv_derivative(r)
                if df_r == 0:
                    break
                newr = r - f / df_r
                if abs(newr - r) < 1e-6:
                    r = newr
                    break
                r = newr
            irr = r
            if irr <= -0.9999 or math.isnan(irr) or abs(irr) > 1e6:
                irr = None
    except Exception:
        irr = None

    # PP: thời gian hoàn vốn (không chiết khấu)
    cumulative = 0.0
    pp = None
    for i, cf in enumerate(cashflows):
        cumulative += cf
        if cumulative >= 0 and i > 0:
            # tính phần năm dùng linear interpolation
            prev_cum = cumulative - cf
            frac = (0 - prev_cum) / (cf) if cf != 0 else 0
            pp = (i - 1) + frac
            break

    # DPP: thời gian hoàn vốn có chiết khấu
    cumulative_d = 0.0
    dpp = None
    for i, cf in enumerate(cashflows):
        discounted = cf / ((1 + discount_rate) ** i)
        cumulative_d += discounted
        if cumulative_d >= 0 and i > 0:
            # tìm phần năm
            # recompute previous discounted cumulative
            prev_cum_d = cumulative_d - discounted
            frac = (0 - prev_cum_d) / (discounted) if discounted != 0 else 0
            dpp = (i - 1) + frac
            break

    return {
        "NPV": npv,
        "IRR": irr,
        "PP (years)": pp,
        "DPP (years)": dpp
    }

# --- Hàm gọi AI để phân tích kết quả (thử OpenAI rồi Gemini nếu có) ---
def get_ai_comment(summary_text: str) -> str:
    # Thử OpenAI (openai-python) nếu có key
    openai_key = None
    try:
        openai_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        openai_key = None

    if openai_key:
        try:
            import openai
            openai.api_key = openai_key
            prompt = (
                "Bạn là chuyên gia phân tích tài chính. Dưới đây là tóm tắt các chỉ số dự án và bảng dòng tiền. "
                "Hãy đưa ra nhận xét ngắn gọn, rõ ràng (3-5 đoạn), nêu ưu điểm, rủi ro chính và khuyến nghị (3-4 câu).\n\n"
                + summary_text
            )
            resp = openai.ChatCompletion.create(
                model="gpt-4o" if "gpt-4o" in openai.Model.list().data else "gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=600,
                temperature=0.1
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            # Không crash app khi AI lỗi, trả về lỗi dạng văn bản
            return f"Lỗi gọi OpenAI: {e}"

    # Nếu không có OpenAI, thử Google GenAI (gemini) nếu có key
    try:
        gem_key = st.secrets.get("GEMINI_API_KEY") or None
    except Exception:
        gem_key = None

    if gem_key:
        try:
            from google import genai
            client = genai.Client(api_key=gem_key)
            prompt = (
                "Bạn là chuyên gia phân tích tài chính. Dưới đây là tóm tắt các chỉ số dự án và bảng dòng tiền. "
                "Hãy đưa ra nhận xét ngắn gọn, rõ ràng (3-5 đoạn), nêu ưu điểm, rủi ro chính và khuyến nghị (3-4 câu).\n\n"
                + summary_text
            )
            response = client.models.generate_content(model="gemini-2.5", contents=prompt)
            return response.text
        except Exception as e:
            return f"Lỗi gọi Gemini API: {e}"

    return "Không tìm thấy cấu hình API AI. Vui lòng cấu hình OPENAI_API_KEY hoặc GEMINI_API_KEY trong Streamlit Secrets."

# --- UI chính ---
st.markdown("Hướng dẫn: Tải lên file Word (.docx) chứa phương án kinh doanh. Ứng dụng sẽ cố gắng trích xuất các thông số chính và xây dựng bảng dòng tiền tự động.")

uploaded = st.file_uploader("1. Tải file Word (docx) phương án kinh doanh", type=["docx"])

col1, col2 = st.columns([1, 1])
with col1:
    use_ai_extract = st.checkbox("Sử dụng AI/heuristic để trích xuất thông tin (khuyến nghị)", value=True)
with col2:
    override_life = st.number_input("Ghi đè Dòng đời dự án (năm) (0 để giữ giá trị trích xuất)", min_value=0, step=1, value=0)

if uploaded is not None:
    try:
        # đọc docx
        doc_bytes = io.BytesIO(uploaded.read())
        full_text = read_docx(doc_bytes)

        st.subheader("Tóm tắt nội dung đầu tiên (first 500 chars)")
        st.code(full_text[:1000])

        if st.button("2. Lọc thông tin dự án"):
            # trích xuất thông tin
            info = extract_project_info(full_text) if use_ai_extract else {
                "Vốn đầu tư": None, "Dòng đời dự án (năm)": None, "Doanh thu hàng năm": None,
                "Chi phí hàng năm": None, "WACC": None, "Thuế suất": None, "Ghi chú thô": full_text[:500]
            }

            # Nếu người dùng override dòng đời
            if override_life > 0:
                info["Dòng đời dự án (năm)"] = int(override_life)

            st.subheader("Kết quả trích xuất thông tin dự án")
            display_info = {
                "Vốn đầu tư (VND)": info.get("Vốn đầu tư"),
                "Dòng đời (năm)": info.get("Dòng đời dự án (năm)"),
                "Doanh thu hàng năm (VND)": info.get("Doanh thu hàng năm"),
                "Chi phí hàng năm (VND)": info.get("Chi phí hàng năm"),
                "WACC (decimal)": info.get("WACC"),
                "Thuế suất (decimal)": info.get("Thuế suất")
            }
            st.json(display_info)

            # Xây bảng dòng tiền
            st.subheader("3. Bảng Dòng Tiền Dự Án (giả định đơn giản)")
            df_cf = build_cashflow_table(info)
            st.dataframe(df_cf.style.format({
                "Doanh thu": "{:,.0f}",
                "Chi phí": "{:,.0f}",
                "EBIT": "{:,.0f}",
                "Thuế": "{:,.0f}",
                "NOPAT": "{:,.0f}",
                "CapEx": "{:,.0f}",
                "Dòng tiền ròng": "{:,.0f}"
            }), use_container_width=True)

            # Tính chỉ số
            st.subheader("4. Tính các chỉ số đánh giá hiệu quả")
            discount_rate = st.number_input("Chiết khấu dùng cho NPV/DPP (WACC mặc định nếu có)", min_value=0.0, max_value=1.0, value=float(info.get("WACC") or 0.1), step=0.01)
            cashflows = df_cf["Dòng tiền ròng"].tolist()
            metrics = compute_project_metrics(cashflows, discount_rate)

            # Hiển thị metrics
            mdf = pd.DataFrame([{
                "NPV (VND)": metrics["NPV"],
                "IRR (decimal)": metrics["IRR"],
                "PP (năm)": metrics["PP (years)"],
                "DPP (năm)": metrics["DPP (years)"]
            }])
            st.table(mdf.style.format({
                "NPV (VND)": "{:,.0f}",
                "IRR (decimal)": "{:.2%}",
                "PP (năm)": "{:.2f}",
                "DPP (năm)": "{:.2f}"
            }).applymap(lambda v: "N/A" if (v is None or (isinstance(v, float) and (math.isnan(v) or abs(v) > 1e9))) else v))

            # Nút yêu cầu AI phân tích các chỉ số
            st.subheader("5. Yêu cầu AI phân tích các chỉ số")
            summary = f"Thông tin dự án:\n{display_info}\n\nBảng dòng tiền (tóm tắt):\n{df_cf.to_csv(index=False)}\n\nCác chỉ số:\nNPV={metrics['NPV']}\nIRR={metrics['IRR']}\nPP={metrics['PP (years)']}\nDPP={metrics['DPP (years)']}\nChiết khấu dùng={discount_rate}\n"
            if st.button("Gửi yêu cầu AI phân tích chỉ số"):
                with st.spinner("Đang gửi dữ liệu đến AI để phân tích..."):
                    ai_comment = get_ai_comment(summary)
                    st.markdown("**Kết quả phân tích từ AI:**")
                    st.info(ai_comment)

    except Exception as e:
        st.error(f"Có lỗi khi xử lý file: {e}")
else:
    st.info("Vui lòng tải lên file .docx để bắt đầu phân tích.")
