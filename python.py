import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import time
from google import genai
from google.genai.errors import APIError

# --- Cấu hình Trang Streamlit ---
st.set_page_config(
    page_title="App Đánh Giá Phương Án Kinh Doanh (DCF)",
    layout="wide"
)

st.title("Ứng dụng Đánh Giá Dự Án (DCF) bằng AI 🚀")
st.markdown("Sử dụng Gemini AI để trích xuất thông số từ file Word (.txt hoặc .docx) và thẩm định dự án.")

# Khóa API - Lấy từ Streamlit Secrets
API_KEY = st.secrets.get("GEMINI_API_KEY")

# --- 0. Định nghĩa Schema cho Dữ liệu Cần Trích Xuất (Bắt buộc) ---
# Sử dụng JSON Schema để đảm bảo AI trả về dữ liệu có cấu trúc.
FINANCIAL_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "investment_capital": {"type": "NUMBER", "description": "Tổng vốn đầu tư ban đầu của dự án, thường là chi phí ban đầu (năm 0)."},
        "project_lifespan": {"type": "INTEGER", "description": "Số năm hoạt động của dự án (Dòng đời dự án)."},
        "annual_revenue": {"type": "NUMBER", "description": "Doanh thu dự kiến hàng năm (Giả định không đổi trong các năm hoạt động)."},
        "annual_cost": {"type": "NUMBER", "description": "Chi phí hoạt động dự kiến hàng năm (Không bao gồm thuế, khấu hao. Giả định không đổi)."},
        "wacc_rate": {"type": "NUMBER", "description": "Chi phí vốn bình quân (WACC), dưới dạng phần trăm (ví dụ: 10 cho 10%)."},
        "tax_rate": {"type": "NUMBER", "description": "Thuế suất doanh nghiệp, dưới dạng phần trăm (ví dụ: 20 cho 20%)."}
    },
    "required": ["investment_capital", "project_lifespan", "annual_revenue", "annual_cost", "wacc_rate", "tax_rate"]
}

# --- 1. Hàm Trích Xuất Dữ liệu bằng AI (Gemini) ---
def extract_financial_data(doc_content: str, api_key: str):
    """Sử dụng Gemini để trích xuất 6 thông số tài chính chính từ nội dung văn bản."""
    if not api_key:
        st.error("Lỗi: Không tìm thấy Khóa API 'GEMINI_API_KEY'.")
        return None

    try:
        client = genai.Client(api_key=api_key)
        
        system_prompt = (
            "Bạn là một chuyên gia phân tích tài chính. Nhiệm vụ của bạn là đọc nội dung dự án kinh doanh "
            "và trích xuất 6 thông số tài chính chính xác (chỉ số, không phải chuỗi ký tự) vào định dạng JSON đã cho. "
            "Chuyển đổi tất cả các giá trị tiền tệ sang đơn vị triệu đồng (hoặc đơn vị phù hợp) và WACC/Thuế sang số (ví dụ: 10 cho 10%). "
            "Nếu một giá trị không rõ ràng, hãy đưa ra ước tính hợp lý nhất hoặc 0."
        )

        user_prompt = f"""
        Trích xuất 6 thông số tài chính sau từ nội dung dự án kinh doanh:
        1. Vốn đầu tư (Investment Capital)
        2. Dòng đời dự án (Project Lifespan - năm)
        3. Doanh thu hàng năm (Annual Revenue)
        4. Chi phí hoạt động hàng năm (Annual Cost)
        5. Chi phí vốn (WACC - %)
        6. Thuế suất (Tax Rate - %)

        Nội dung dự án:
        ---
        {doc_content[:15000]}
        ---
        """

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=user_prompt,
            config={
                "system_instruction": system_prompt,
                "response_mime_type": "application/json",
                "response_schema": FINANCIAL_SCHEMA
            }
        )
        # Gemini trả về chuỗi JSON, cần parse
        return json.loads(response.text)

    except APIError as e:
        st.error(f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}")
        return None
    except json.JSONDecodeError:
        st.error("Lỗi phân tích JSON: AI không trả về định dạng JSON hợp lệ.")
        return None
    except Exception as e:
        st.error(f"Đã xảy ra lỗi không xác định trong quá trình trích xuất: {e}")
        return None

# --- 2 & 3. Hàm Xây dựng Dòng tiền và Tính toán Chỉ số ---

def calculate_project_metrics(data: dict):
    """
    Xây dựng bảng dòng tiền và tính toán các chỉ số DCF (NPV, IRR, PP, DPP).
    Giả định Dòng tiền hoạt động (Annual CF) = (Doanh thu - Chi phí) * (1 - Thuế)
    """
    
    # Chuẩn hóa dữ liệu
    I = data['investment_capital']
    N = int(data['project_lifespan'])
    R = data['annual_revenue']
    C = data['annual_cost']
    WACC = data['wacc_rate'] / 100.0
    Tax = data['tax_rate'] / 100.0

    # 1. Xây dựng Bảng Dòng tiền
    years = list(range(N + 1))
    
    # Dòng tiền hoạt động hàng năm (Annual Operating Cash Flow - OCF)
    OCF = (R - C) * (1 - Tax)
    
    # Dòng tiền dự án (Project Cash Flow - CF)
    cash_flows = [-I] + [OCF] * N
    
    df_cash_flow = pd.DataFrame({
        'Năm': years,
        'Vốn đầu tư (I)': [-I] + [0] * N,
        'Doanh thu (R)': [0] + [R] * N,
        'Chi phí (C)': [0] + [C] * N,
        'Dòng tiền hoạt động (OCF)': [0] + [OCF] * N,
        'Dòng tiền Dự án (CF)': cash_flows,
    })
    
    # 2. Tính toán các chỉ số
    
    # NPV (Net Present Value)
    npv_value = np.npv(WACC, cash_flows)
    
    # IRR (Internal Rate of Return)
    try:
        irr_value = np.irr(cash_flows)
    except Exception:
        irr_value = np.nan # Không thể tính IRR

    # Payback Period (PP) - Thời gian hoàn vốn
    cumulative_cf = np.cumsum(cash_flows)
    pp_year = next((i for i, cf in enumerate(cumulative_cf) if cf >= 0), N + 1)
    
    if pp_year <= N:
        # Tính chính xác hơn: Năm hoàn vốn + (Vốn còn lại / Dòng tiền năm hoàn vốn)
        pp_value = (pp_year - 1) + (-cumulative_cf[pp_year - 1] / cash_flows[pp_year])
    else:
        pp_value = "Không hoàn vốn"

    # Discounted Payback Period (DPP) - Thời gian hoàn vốn có chiết khấu
    
    # Tính dòng tiền chiết khấu (Discounted Cash Flow - DCF)
    discount_factors = [1 / (1 + WACC)**t for t in years]
    dcf_flows = [cf * df for cf, df in zip(cash_flows, discount_factors)]
    df_cash_flow['Dòng tiền Chiết khấu (DCF)'] = dcf_flows
    
    # Tính tích lũy DCF
    cumulative_dcf = np.cumsum(dcf_flows)
    dpp_year = next((i for i, dcf in enumerate(cumulative_dcf) if dcf >= 0), N + 1)
    
    if dpp_year <= N:
        # Tính chính xác hơn
        dpp_value = (dpp_year - 1) + (-cumulative_dcf[dpp_year - 1] / dcf_flows[dpp_year])
    else:
        dpp_value = "Không hoàn vốn chiết khấu"

    # Thêm cột tích lũy vào DataFrame để dễ theo dõi
    df_cash_flow['Dòng tiền tích lũy'] = cumulative_cf
    df_cash_flow['DCF tích lũy'] = cumulative_dcf
    
    metrics = {
        "NPV": npv_value,
        "IRR": irr_value,
        "PP": pp_value,
        "DPP": dpp_value
    }
    
    return df_cash_flow, metrics

# --- 4. Hàm Phân Tích Chỉ số bằng AI (Gemini) ---

def get_ai_analysis(data: dict, metrics: dict, api_key: str):
    """Gửi các chỉ số đã tính toán đến Gemini API và nhận nhận xét thẩm định."""
    
    # Định dạng các chỉ số để gửi cho AI
    irr_str = f"{metrics['IRR'] * 100:.2f}%" if not pd.isna(metrics['IRR']) else "Không thể tính toán"
    
    metrics_markdown = f"""
| Chỉ số | Giá trị |
| :--- | :--- |
| Vốn đầu tư (I) | {data['investment_capital']:,.0f} |
| Dòng đời dự án (N) | {data['project_lifespan']} năm |
| Doanh thu/năm | {data['annual_revenue']:,.0f} |
| Chi phí/năm | {data['annual_cost']:,.0f} |
| WACC | {data['wacc_rate']:.2f}% |
| Thuế | {data['tax_rate']:.2f}% |
| **NPV** | **{metrics['NPV']:,.0f}** |
| **IRR** | **{irr_str}** |
| **PP (Hoàn vốn)** | **{metrics['PP']}** |
| **DPP (Hoàn vốn chiết khấu)** | **{metrics['DPP']}** |
"""

    try:
        client = genai.Client(api_key=api_key)
        model_name = 'gemini-2.5-flash'

        prompt = f"""
        Bạn là một chuyên gia thẩm định dự án đầu tư. Dựa trên các thông số và chỉ số hiệu quả dự án sau, hãy đưa ra một nhận xét thẩm định chuyên nghiệp, ngắn gọn (khoảng 3-4 đoạn) về tính khả thi và mức độ hấp dẫn của dự án.
        
        Đánh giá tập trung vào:
        1. **Khả năng sinh lời:** Nhận xét về NPV (Chấp nhận/Từ chối) và so sánh IRR với WACC (tỷ suất sinh lời và chi phí vốn).
        2. **Tính thanh khoản:** Đánh giá thời gian hoàn vốn (PP và DPP).
        3. **Rủi ro/Độ nhạy:** Đề xuất một số điểm cần lưu ý hoặc phân tích độ nhạy.
        
        Tất cả các giá trị tiền tệ là đơn vị [Triệu đồng] (giả định).
        
        ---
        {metrics_markdown}
        ---
        """

        with st.spinner('Đang gửi dữ liệu và chờ Gemini phân tích...'):
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            return response.text

    except APIError as e:
        return f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API. Chi tiết lỗi: {e}"
    except Exception as e:
        return f"Đã xảy ra lỗi không xác định: {e}"


# --- Giao diện Streamlit ---

# Chức năng 1: Tải File và Lọc Dữ liệu
st.subheader("1. Tải File & Lọc Dữ liệu Dự án bằng AI")
uploaded_file = st.file_uploader(
    "Tải lên file Word (.txt, .docx). Để đảm bảo trích xuất chính xác, vui lòng tải lên file **.txt** hoặc **.docx** có nội dung rõ ràng (Streamlit sẽ đọc nội dung văn bản).",
    type=['txt', 'docx']
)

if uploaded_file is not None:
    # Đọc nội dung file dưới dạng chuỗi văn bản
    file_content = ""
    try:
        # Giả định file là text/docx và cố gắng decode
        # Trong môi trường Streamlit, việc đọc nội dung văn bản từ các loại file này
        # thường được thực hiện đơn giản bằng cách decode bytes
        bytes_data = uploaded_file.getvalue()
        file_content = bytes_data.decode('utf-8')
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}. Vui lòng thử lại với file .txt hoặc kiểm tra encoding.")
        file_content = ""

    if file_content:
        st.text_area("Nội dung Văn bản Trích xuất (1500 ký tự đầu):", file_content[:1500], height=200)
        
        if 'extracted_data' not in st.session_state:
            st.session_state['extracted_data'] = None

        if st.button("🔴 Kích hoạt AI Lọc Dữ liệu"):
            if API_KEY:
                with st.spinner('Đang phân tích văn bản và trích xuất thông số...'):
                    data = extract_financial_data(file_content, API_KEY)
                    st.session_state['extracted_data'] = data
            else:
                st.error("Vui lòng cấu hình Khóa API 'GEMINI_API_KEY' trong Streamlit Secrets để sử dụng chức năng AI.")

        if st.session_state['extracted_data']:
            st.subheader("Kết quả Trích xuất AI (Bước 1 Hoàn thành)")
            
            # Hiển thị dữ liệu trích xuất dưới dạng DataFrame
            df_extracted = pd.DataFrame([
                {"Chỉ tiêu": "Vốn đầu tư (triệu)", "Giá trị": f"{st.session_state['extracted_data']['investment_capital']:,.0f}"},
                {"Chỉ tiêu": "Dòng đời dự án (năm)", "Giá trị": st.session_state['extracted_data']['project_lifespan']},
                {"Chỉ tiêu": "Doanh thu hàng năm (triệu)", "Giá trị": f"{st.session_state['extracted_data']['annual_revenue']:,.0f}"},
                {"Chỉ tiêu": "Chi phí hàng năm (triệu)", "Giá trị": f"{st.session_state['extracted_data']['annual_cost']:,.0f}"},
                {"Chỉ tiêu": "WACC (%)", "Giá trị": f"{st.session_state['extracted_data']['wacc_rate']:.2f}%"},
                {"Chỉ tiêu": "Thuế suất (%)", "Giá trị": f"{st.session_state['extracted_data']['tax_rate']:.2f}%"},
            ])
            st.dataframe(df_extracted, use_container_width=True, hide_index=True)

            # --- 2 & 3. Xây dựng Dòng tiền và Tính toán Chỉ số ---
            try:
                data = st.session_state['extracted_data']
                df_cash_flow, metrics = calculate_project_metrics(data)
                
                st.markdown("---")
                st.subheader("2. Bảng Dòng tiền Dự án (Cash Flow)")
                # Hiển thị bảng dòng tiền (Cash Flow)
                st.dataframe(df_cash_flow.style.format({
                    col: '{:,.0f}' for col in df_cash_flow.columns if col != 'Năm' and col != 'PP' and col != 'DPP'
                }), use_container_width=True)
                
                st.markdown("---")
                st.subheader("3. Các Chỉ số Đánh giá Hiệu quả Dự án")
                
                col1, col2, col3, col4 = st.columns(4)
                
                # Hiển thị NPV
                col1.metric("NPV (Giá trị hiện tại ròng - Triệu VNĐ)", f"{metrics['NPV']:,.0f}", help="NPV > 0: Dự án khả thi.")
                
                # Hiển thị IRR
                irr_value_str = f"{metrics['IRR'] * 100:.2f}%" if not pd.isna(metrics['IRR']) else "Không thể tính"
                wacc_rate_str = f"{data['wacc_rate']:.2f}%"
                delta_irr = metrics['IRR'] * 100 - data['wacc_rate']
                col2.metric(
                    "IRR (Tỷ suất sinh lời nội bộ)", 
                    irr_value_str, 
                    delta=f"So với WACC ({wacc_rate_str}): {delta_irr:.2f}%", 
                    help="IRR > WACC: Dự án chấp nhận được."
                )

                # Hiển thị PP
                pp_str = f"{metrics['PP']:.2f} năm" if isinstance(metrics['PP'], (int, float)) else metrics['PP']
                col3.metric("PP (Thời gian hoàn vốn)", pp_str, help="Thời gian cần để dòng tiền tích lũy bằng 0.")

                # Hiển thị DPP
                dpp_str = f"{metrics['DPP']:.2f} năm" if isinstance(metrics['DPP'], (int, float)) else metrics['DPP']
                col4.metric("DPP (Hoàn vốn chiết khấu)", dpp_str, help="Thời gian cần để dòng tiền chiết khấu tích lũy bằng 0.")
                
                st.session_state['metrics'] = metrics # Lưu kết quả để phân tích AI
                st.session_state['data'] = data # Lưu dữ liệu thô
                
                st.markdown("---")
                
                # --- 4. Yêu cầu AI Phân tích ---
                st.subheader("4. Phân tích Thẩm định Dự án (AI)")
                if st.button("🧠 Yêu cầu AI Phân tích Hiệu quả Dự án"):
                    if API_KEY:
                        ai_result = get_ai_analysis(st.session_state['data'], st.session_state['metrics'], API_KEY)
                        st.markdown("**Kết quả Phân tích Thẩm định từ Gemini AI:**")
                        st.info(ai_result)
                    else:
                        st.error("Vui lòng cấu hình Khóa API 'GEMINI_API_KEY' trong Streamlit Secrets để sử dụng chức năng AI.")

            except Exception as e:
                st.error(f"Lỗi tính toán: Không thể xây dựng bảng dòng tiền. Vui lòng kiểm tra lại dữ liệu trích xuất: {e}")

else:
    st.info("Vui lòng tải lên file chứa thông tin dự án để bắt đầu quá trình thẩm định.")
