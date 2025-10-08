# app.py
import re
import io
import math
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd
import numpy as np

from docx import Document

# Nếu numpy_financial không có, ta fallback
try:
    import numpy_financial as nf
    has_nf = True
except Exception:
    has_nf = False

# --- Cấu hình Trang Streamlit ---
st.set_page_config(page_title="App Phân Tích Phương Án Kinh Doanh", layout="wide")
st.title("App Phân Tích Phương Án Kinh Doanh từ file Word — (Phiên bản đã sửa lỗi)")

# --- Utility: đọc text từ file .docx ---
def read_docx(file_bytes: io.BytesIO) -> str:
    doc = Document(file_bytes)
    paragraphs = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(paragraphs)

# --- Utility: chuyển chuỗi số có dấu phẩy, dấu chấm sang float ---
def parse_number(text: Optional[str]) -> Optional[float]:
    if text is None:
        return None
    if isinstance(text, (int, float, np.integer, np.floating)):
        return float(text)
    txt = str(text).strip()
    if txt == "":
        return None
    # Nếu chuỗi chứa chữ (ví dụ tên dự án), trả None
    # Nhưng vẫn chấp nhận chuỗi có số và chữ đơn vị như "1,2 tỷ"
    # Kiểm tra nhanh nếu không có bất kỳ chữ số nào thì trả None
    if not re.search(r'\d', txt):
        return None
    # Phát hiện phần trăm
    is_percent = "%" in txt
    txt = txt.replace("%", "")
    multiplier = 1.0
    if re.search(r'\b(tỷ|ty|billion)\b', txt, flags=re.I):
        multiplier = 1e9
        txt = re.sub(r'\b(tỷ|ty|billion)\b', '', txt, flags=re.I)
    elif re.search(r'\b(triệu|trieu|million)\b', txt, flags=re.I):
        multiplier = 1e6
        txt = re.sub(r'\b(triệu|trieu|million)\b', '', txt, flags=re.I)
    # Loại bỏ ký tự không phải số, dấu chấm, dấu phẩy, dấu âm
    txt = re.sub(r'[^\d\.,\-]', '', txt)
    # Chuyển định dạng
    if txt.count(',') > 0 and txt.count('.') > 0:
        if txt.rfind('.') > txt.rfind(','):
            txt = txt.replace(',', '')
        else:
            txt = txt.replace('.', '').replace(',', '.')
    else:
        txt = txt.replace(',', '.')
    # Xử lý chỉ dấu âm lẫn lộn
    txt = txt.strip('.')
    try:
        val = float(txt)
        if is_percent:
            return val / 100.0
        return val * multiplier
    except Exception:
        return None

# --- Hàm lọc thông tin quan trọng từ văn bản ---
def extract_project_info(text: str) -> Dict[str, Any]:
    info = {
        "Vốn đầu tư": None,
        "Dòng đời dự án (năm)": None,
        "Doanh thu hàng năm": None,
        "Chi phí hàng năm": None,
        "WACC": None,
        "Thuế suất": None,
        "Ghi chú thô": text[:500]
    }

    # Patterns cơ bản
    patterns = {
        "Vốn đầu tư": r'(vốn đầu tư|tổng vốn|tổng đầu tư|initial investment|capex)[\s\:\-–]*([^\n\r,;]+)',
        "Dòng đời dự án (năm)": r'(dòng đời|thời gian khai thác|thời gian dự án|lifetime|project life)[^\d\n\r]*?(\d{1,2})\s*(năm|year)?',
        "WACC": r'(wacc|chi phí vốn|weighted average cost of capital)[\s\:\-–]*([^\n\r,;]+)',
        "Thuế suất": r'(thuế suất|tax rate|tax)[\s\:\-–]*([^\n\r,;]+%)',
        "Doanh thu hàng năm": r'(doanh thu|revenue|sales)[\s\:\-–]*([^\n\r,;]+)',
        "Chi phí hàng năm": r'(chi phí|costs|expenditure)[\s\:\-–]*([^\n\r,;]+)'
    }

    for key, pat in patterns.items():
        m = re.search(pat, text, flags=re.I)
        if m:
            token = m.groups()[-1].strip()
            num = parse_number(token)
            info[key] = num if num is not None else token

    # Thử thêm phát hiện WACC không có ký hiệu %
    if info["WACC"] is None:
        m = re.search(r'chi phí vốn[\s\:\-–]*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            info["WACC"] = parse_number(m.group(1) + "%")
    # Thuế suất dạng "Thuế: 20%" hoặc "thuế suất 20%"
    if info["Thuế suất"] is None:
        m = re.search(r'thuế\s*(suất)?[\s\:\-–]*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            info["Thuế suất"] = parse_number(m.group(2) + "%")

    # Gom các số gần từ khóa doanh thu/chi phí nếu không tìm được giá trị đơn
    revenues = []
    costs = []
    for line in text.splitlines():
        if re.search(r'\b(doanh thu|revenue|sales)\b', line, flags=re.I):
            for numstr in re.findall(r'[\d\.,]+\s*(tỷ|ty|triệu|trieu)?\b', line, flags=re.I):
                n = parse_number(numstr)
                if n is not None:
                    revenues.append(n)
        if re.search(r'\b(chi phí|cost|costs|expenditure)\b', line, flags=re.I):
            for numstr in re.findall(r'[\d\.,]+\s*(tỷ|ty|triệu|trieu)?\b', line, flags=re.I):
                n = parse_number(numstr)
                if n is not None:
                    costs.append(n)
    if revenues and info["Doanh thu hàng năm"] is None:
        info["Doanh thu hàng năm"] = sum(revenues) / len(revenues)
    if costs and info["Chi phí hàng năm"] is None:
        info["Chi phí hàng năm"] = sum(costs) / len(costs)

    # Nếu vốn đầu tư vẫn là chuỗi mô tả (ví dụ tên dự án), cố gắng tìm số trong text chung
    if info["Vốn đầu tư"] is None:
        m = re.search(r'(vốn|tổng vốn|tổng đầu tư|capex)[^\d\n\r]*([\d\.,]+\s*(tỷ|ty|triệu|trieu)?)', text, flags=re.I)
        if m:
            info["Vốn đầu tư"] = parse_number(m.group(2))

    # Mặc định nếu vẫn None
    if info["Dòng đời dự án (năm)"] is None:
        info["Dòng đời dự án (năm)"] = 5
    if info["WACC"] is None:
        info["WACC"] = 0.10
    if info["Thuế suất"] is None:
        info["Thuế suất"] = 0.20

    # Chuyển các giá trị string sang numeric nếu có thể
    for key in ["Vốn đầu tư", "Doanh thu hàng năm", "Chi phí hàng năm", "WACC", "Thuế suất"]:
        val = info.get(key)
        if isinstance(val, str):
            parsed = parse_number(val)
            if parsed is not None:
                info[key] = parsed

    return info

# --- Xây bảng dòng tiền dự án từ thông tin đã lọc ---
def build_cashflow_table(info: Dict[str, Any]) -> pd.DataFrame:
    # Lấy giá trị và ép kiểu an toàn
    try:
        life = int(info.get("Dòng đời dự án (năm)", 5) or 5)
    except Exception:
        life = 5

    capex_raw = info.get("Vốn đầu tư", 0)
    revenue_raw = info.get("Doanh thu hàng năm", 0)
    cost_raw = info.get("Chi phí hàng năm", 0)
    tax = info.get("Thuế suất", 0.2) or 0.2

    capex = parse_number(capex_raw) if capex_raw is not None else None
    revenue = parse_number(revenue_raw) if revenue_raw is not None else None
    cost = parse_number(cost_raw) if cost_raw is not None else None

    # Nếu parse không thành công, đặt mặc định và thông báo
    warnings = []
    if capex is None:
        warnings.append("Vốn đầu tư không xác định hoặc không phải số, dùng capex = 0")
        capex = 0.0
    if revenue is None:
        warnings.append("Doanh thu hàng năm không xác định, dùng doanh thu = 0")
        revenue = 0.0
    if cost is None:
        warnings.append("Chi phí hàng năm không xác định, dùng chi phí = 0")
        cost = 0.0
    if not isinstance(tax, (int, float, np.integer, np.floating)):
        # cố parse nếu là string như "20%"
        tparsed = parse_number(str(tax))
        tax = tparsed if tparsed is not None else 0.2
        if tparsed is None:
            warnings.append("Thuế suất không hợp lệ, dùng thuế suất = 0.2")

    # Hiển thị cảnh báo trong Streamlit nếu có
    if warnings:
        for w in warnings:
            st.warning(w)

    rows = []
    for year in range(0, life + 1):
        if year == 0:
            try:
                capex_val = -abs(float(capex))
            except Exception:
                capex_val = 0.0
            rows.append({
                "Năm": year,
                "Doanh thu": 0.0,
                "Chi phí": 0.0,
                "EBIT": 0.0,
                "Thuế": 0.0,
                "NOPAT": 0.0,
                "CapEx": capex_val,
                "Dòng tiền ròng": capex_val
            })
        else:
            ebit = float(revenue) - float(cost)
            tax_amt = ebit * float(tax) if ebit > 0 else 0.0
            nopat = ebit - tax_amt
            cf = nopat
            rows.append({
                "Năm": year,
                "Doanh thu": float(revenue),
                "Chi phí": float(cost),
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
    npv = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cashflows))
    irr = None
    try:
        if has_nf:
            irr = nf.irr(cashflows)
        else:
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
            if irr is None or irr <= -0.9999 or math.isnan(irr) or abs(irr) > 1e9:
                irr = None
    except Exception:
        irr = None

    cumulative = 0.0
    pp = None
    for i, cf in enumerate(cashflows):
        cumulative += cf
        if cumulative >= 0 and i > 0:
            prev_cum = cumulative - cf
            frac = (0 - prev_cum) / cf if cf != 0 else 0
            pp = (i - 1) + frac
            break

    cumulative_d = 0.0
    dpp = None
    for i, cf in enumerate(cashflows):
        discounted = cf / ((1 + discount_rate) ** i)
        cumulative_d += discounted
        if cumulative_d >= 0 and i > 0:
            prev_cum_d = cumulative_d - discounted
            frac = (0 - prev_cum_d) / discounted if discounted != 0 else 0
            dpp = (i - 1) + frac
            break

    return {"NPV": npv, "IRR": irr, "PP (years)": pp, "DPP (years)": dpp}

# --- Hàm gọi AI (giữ nguyên như trước; tùy cấu hình secrets) ---
def get_ai_comment(summary_text: str) -> str:
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
            return f"Lỗi gọi OpenAI: {e}"

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
    use_ai_extract = st.checkbox("Sử dụng heuristic để trích xuất thông tin (khuyến nghị)", value=True)
with col2:
    override_life = st.number_input("Ghi đè Dòng đời dự án (năm) (0 để giữ giá trị trích xuất)", min_value=0, step=1, value=0)

if uploaded is not None:
    try:
        doc_bytes = io.BytesIO(uploaded.read())
        full_text = read_docx(doc_bytes)

        st.subheader("Tóm tắt nội dung đầu tiên")
        st.code(full_text[:1000])

        if st.button("Trích thông tin dự án"):
            info = extract_project_info(full_text) if use_ai_extract else {
                "Vốn đầu tư": None, "Dòng đời dự án (năm)": None, "Doanh thu hàng năm": None,
                "Chi phí hàng năm": None, "WACC": None, "Thuế suất": None, "Ghi chú thô": full_text[:500]
            }

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

            st.subheader("4. Tính các chỉ số đánh giá hiệu quả")
            discount_rate = st.number_input("Chiết khấu dùng cho NPV/DPP (WACC mặc định nếu có)", min_value=0.0, max_value=1.0, value=float(info.get("WACC") or 0.1), step=0.01)
            cashflows = df_cf["Dòng tiền ròng"].tolist()
            metrics = compute_project_metrics(cashflows, discount_rate)

            mdf = pd.DataFrame([{
                "NPV (VND)": metrics["NPV"],
                "IRR (decimal)": metrics["IRR"],
                "PP (năm)": metrics["PP (years)"],
                "DPP (năm)": metrics["DPP (years)"]
            }])
            # Format hiển thị, xử lý None
            def fmt(x, fmtstr):
                try:
                    if x is None:
                        return "N/A"
                    if isinstance(x, float) and (math.isnan(x) or abs(x) > 1e18):
                        return "N/A"
                    return fmtstr.format(x)
                except Exception:
                    return str(x)
            st.table(mdf.rename(columns={
                "NPV (VND)": "NPV (VND)",
                "IRR (decimal)": "IRR (decimal)",
                "PP (năm)": "PP (năm)",
                "DPP (năm)": "DPP (năm)"
            }).to_dict(orient='records'))

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
