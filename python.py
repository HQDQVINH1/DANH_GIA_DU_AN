# app.py
import re
import io
import math
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd
import numpy as np

from docx import Document

# Try import numpy_financial for IRR; fallback implemented if not available
try:
    import numpy_financial as nf
    has_nf = True
except Exception:
    has_nf = False

# --- Page config ---
st.set_page_config(page_title="App Phân Tích Phương Án Kinh Doanh", layout="wide")
st.title("App Phân Tích Phương Án Kinh Doanh từ file Word (Cập nhật lỗi)")

# --- Utilities ---
def read_docx(file_bytes: io.BytesIO) -> str:
    doc = Document(file_bytes)
    paragraphs = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(paragraphs)

def parse_number(text: Optional[str]) -> Optional[float]:
    """Chuyển nhiều dạng chuỗi số (ví dụ '1,2 tỷ', '20%', '1.234.567') thành float (VND hoặc decimal cho %)."""
    if text is None:
        return None
    if isinstance(text, (int, float, np.integer, np.floating)):
        return float(text)
    txt = str(text).strip()
    if txt == "":
        return None
    # Nếu không có chữ số => không phải số
    if not re.search(r'\d', txt):
        return None
    # detect percent
    is_percent = "%" in txt
    txt = txt.replace("%", "")
    multiplier = 1.0
    # detect units
    if re.search(r'\b(tỷ|ty|billion)\b', txt, flags=re.I):
        multiplier = 1e9
        txt = re.sub(r'\b(tỷ|ty|billion)\b', '', txt, flags=re.I)
    elif re.search(r'\b(triệu|trieu|million)\b', txt, flags=re.I):
        multiplier = 1e6
        txt = re.sub(r'\b(triệu|trieu|million)\b', '', txt, flags=re.I)
    # remove non-numeric except . , -
    txt = re.sub(r'[^\d\.,\-]', '', txt)
    # handle comma/dot
    if txt.count(',') > 0 and txt.count('.') > 0:
        if txt.rfind('.') > txt.rfind(','):
            txt = txt.replace(',', '')
        else:
            txt = txt.replace('.', '').replace(',', '.')
    else:
        txt = txt.replace(',', '.')
    txt = txt.strip('.')
    try:
        val = float(txt)
        if is_percent:
            return val / 100.0
        return val * multiplier
    except Exception:
        return None

# --- Extract project info heuristic with parsing ---
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

    # thêm phát hiện WACC/thuế dạng số
    if info["WACC"] is None:
        m = re.search(r'chi phí vốn[\s\:\-–]*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            info["WACC"] = parse_number(m.group(1) + "%")
    if info["Thuế suất"] is None:
        m = re.search(r'thuế\s*(suất)?[\s\:\-–]*([0-9\.,]+)\s*%', text, flags=re.I)
        if m:
            info["Thuế suất"] = parse_number(m.group(2) + "%")

    # gom số gần doanh thu/chi phí
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

    # cố tìm capex nếu vẫn None
    if info["Vốn đầu tư"] is None:
        m = re.search(r'(vốn|tổng vốn|tổng đầu tư|capex)[^\d\n\r]*([\d\.,]+\s*(tỷ|ty|triệu|trieu)?)', text, flags=re.I)
        if m:
            info["Vốn đầu tư"] = parse_number(m.group(2))

    # defaults
    if info["Dòng đời dự án (năm)"] is None:
        info["Dòng đời dự án (năm)"] = 5
    if info["WACC"] is None:
        info["WACC"] = 0.10
    if info["Thuế suất"] is None:
        info["Thuế suất"] = 0.20

    # convert some string fields to numeric if possible
    for key in ["Vốn đầu tư", "Doanh thu hàng năm", "Chi phí hàng năm", "WACC", "Thuế suất"]:
        val = info.get(key)
        if isinstance(val, str):
            p = parse_number(val)
            if p is not None:
                info[key] = p

    return info

# --- Build cashflow table safely ---
def build_cashflow_table(info: Dict[str, Any]) -> pd.DataFrame:
    try:
        life = int(info.get("Dòng đời dự án (năm)", 5) or 5)
    except Exception:
        life = 5

    capex_raw = info.get("Vốn đầu tư", 0)
    revenue_raw = info.get("Doanh thu hàng năm", 0)
    cost_raw = info.get("Chi phí hàng năm", 0)
    tax_raw = info.get("Thuế suất", 0.2)

    capex = parse_number(capex_raw) if capex_raw is not None else None
    revenue = parse_number(revenue_raw) if revenue_raw is not None else None
    cost = parse_number(cost_raw) if cost_raw is not None else None
    tax = tax_raw
    if not isinstance(tax, (int, float, np.integer, np.floating)):
        tparsed = parse_number(str(tax))
        tax = tparsed if tparsed is not None else 0.2

    warnings = []
    if capex is None:
        warnings.append("Vốn đầu tư không xác định hoặc không phải số. Dùng capex = 0 để tiếp tục.")
        capex = 0.0
    if revenue is None:
        warnings.append("Doanh thu hàng năm không xác định. Dùng doanh thu = 0.")
        revenue = 0.0
    if cost is None:
        warnings.append("Chi phí hàng năm không xác định. Dùng chi phí = 0.")
        cost = 0.0
    if not isinstance(tax, (int, float, np.integer, np.floating)):
        warnings.append("Thuế suất không hợp lệ. Dùng thuế suất = 0.2.")
        tax = 0.2

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

# --- Compute metrics ---
def compute_project_metrics(cashflows: List[float], discount_rate: float) -> Dict[str, Any]:
    npv = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cashflows))
    irr = None
    try:
        if has_nf:
            irr = nf.irr(cashflows)
        else:
            # fallback Newton-Raphson
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

# --- AI calling (returns string or raises) ---
def get_ai_comment(summary_text: str) -> str:
    # Prefer OpenAI if configured, else Gemini if configured
    try:
        openai_key = st.secrets.get("OPENAI_API_KEY")
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
            # chọn model an toàn; tránh gọi openai.Model.list() trong runtime
            model_name = "gpt-4o" if False else "gpt-4o-mini"
            resp = openai.ChatCompletion.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=600,
                temperature=0.1
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            raise RuntimeError(f"Lỗi gọi OpenAI: {e}")

    try:
        gem_key = st.secrets.get("GEMINI_API_KEY")
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
            raise RuntimeError(f"Lỗi gọi Gemini API: {e}")

    raise RuntimeError("Không tìm thấy cấu hình API AI. Vui lòng cấu hình OPENAI_API_KEY hoặc GEMINI_API_KEY trong Streamlit Secrets.")

# --- Main UI ---
st.markdown("Hướng dẫn: Tải file Word (.docx) chứa phương án kinh doanh. Ứng dụng sẽ trích xuất các thông số chính và xây bảng dòng tiền.")

uploaded = st.file_uploader("1. Tải file Word (docx) phương án kinh doanh", type=["docx"])

col1, col2 = st.columns([1, 1])
with col1:
    use_ai_extract = st.checkbox("Sử dụng heuristic để trích xuất thông tin (khuyến nghị)", value=True)
with col2:
    override_life = st.number_input("Ghi đè Dòng đời dự án (năm) (0 để giữ giá trị trích xuất)", min_value=0, step=1, value=0)

# ensure session state keys exist
if 'ai_result' not in st.session_state:
    st.session_state['ai_result'] = None
if 'ai_error' not in st.session_state:
    st.session_state['ai_error'] = None
if 'ai_running' not in st.session_state:
    st.session_state['ai_running'] = False

if uploaded is not None:
    try:
        doc_bytes = io.BytesIO(uploaded.read())
        full_text = read_docx(doc_bytes)

        st.subheader("Tóm tắt nội dung đầu tiên")
        st.code(full_text[:1000])

        if st.button("Trích thông tin dự án", key="extract_info"):
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

            # display metrics with safe formatting
            def safe_fmt(val, numfmt="{:,.0f}"):
                try:
                    if val is None:
                        return "N/A"
                    if isinstance(val, float) and (math.isnan(val) or abs(val) > 1e18):
                        return "N/A"
                    return numfmt.format(val)
                except Exception:
                    return str(val)

            st.table(pd.DataFrame([{
                "NPV (VND)": safe_fmt(metrics["NPV"], "{:,.0f}"),
                "IRR (decimal)": safe_fmt(metrics["IRR"], "{:.2%}") if metrics["IRR"] is not None else "N/A",
                "PP (năm)": safe_fmt(metrics["PP (years)"], "{:.2f}") if metrics["PP (years)"] is not None else "N/A",
                "DPP (năm)": safe_fmt(metrics["DPP (years)"], "{:.2f}") if metrics["DPP (years)"] is not None else "N/A"
            }]))

            # AI analysis section
            st.subheader("5. Yêu cầu AI phân tích các chỉ số")
            summary = (
                f"Thông tin dự án:\n{display_info}\n\n"
                f"Bảng dòng tiền (tóm tắt):\n{df_cf.to_csv(index=False)}\n\n"
                f"Các chỉ số:\nNPV={metrics['NPV']}\nIRR={metrics['IRR']}\nPP={metrics['PP (years)']}\nDPP={metrics['DPP (years)']}\nChiết khấu dùng={discount_rate}\n"
            )

            if st.button("Gửi yêu cầu AI phân tích chỉ số", key="send_ai_analysis"):
                st.session_state['ai_result'] = None
                st.session_state['ai_error'] = None
                st.session_state['ai_running'] = True

                # check API keys
                has_openai = bool(st.secrets.get("OPENAI_API_KEY", "")) if hasattr(st, "secrets") else False
                has_gemini = bool(st.secrets.get("GEMINI_API_KEY", "")) if hasattr(st, "secrets") else False
                if not (has_openai or has_gemini):
                    st.session_state['ai_error'] = "Không tìm thấy OPENAI_API_KEY hoặc GEMINI_API_KEY trong Streamlit Secrets."
                    st.session_state['ai_running'] = False
                else:
                    try:
                        with st.spinner("Đang gửi dữ liệu đến AI để phân tích..."):
                            ai_comment = get_ai_comment(summary)
                            if ai_comment is None:
                                raise RuntimeError("AI trả về None")
                            st.session_state['ai_result'] = ai_comment
                    except Exception as e:
                        st.session_state['ai_error'] = f"Lỗi khi gọi AI: {e}"
                    finally:
                        st.session_state['ai_running'] = False

            # show AI status / result / error from session_state
            if st.session_state.get('ai_running', False):
                st.info("Yêu cầu đang được xử lý, xin chờ...")
            elif st.session_state.get('ai_error'):
                st.error(st.session_state['ai_error'])
            elif st.session_state.get('ai_result'):
                st.markdown("**Kết quả phân tích từ AI:**")
                st.info(st.session_state['ai_result'])

    except Exception as e:
        st.error(f"Có lỗi khi xử lý file: {e}")
else:
    st.info("Vui lòng tải lên file .docx để bắt đầu phân tích.")
