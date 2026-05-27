import streamlit as st
import pandas as pd
from excel_generator import generate_quotation_excel

# 페이지 설정
st.set_page_config(page_title="시험 견적 시스템", layout="wide")
st.title("시험 견적 시스템")

# 데이터 로드
@st.cache_data
def load_data():
    df_p = pd.read_excel("시험수수료표.xlsx", sheet_name="수수료표")
    df_c = pd.read_excel("시험수수료표.xlsx", sheet_name="업체DB")
    df_p["표시명"] = df_p.apply(lambda row: f"{row['시험명']} / {row['시험방법']}" if pd.notna(row['시험방법']) and row['시험방법'] != "" else row['시험명'], axis=1)
    return df_p, df_c

df_price, df_company = load_data()

# UI 구성
col_a, col_b = st.columns(2)
with col_a:
    selected_company = st.selectbox("업체 선택", options=[""] + df_company["업체명"].tolist())
    sample_name = st.text_input("시료명")
with col_b:
    manager = st.text_input("담당자")
    phone = st.text_input("연락처")

selected_tests = st.multiselect("시험 선택", options=df_price["표시명"].unique().tolist())

st.subheader("공통 옵션")
col1, col2, col3 = st.columns(3)
with col1: quantity = st.number_input("시료수", min_value=1, value=1)
with col2: discount_rate = st.number_input("할인율 (%)", min_value=0, max_value=100, value=0)
with col3: service_type = st.radio("진행 유형", ["일반", "지급 (1.5배)", "즉시 (2.5배)"])

col1, col2, col3 = st.columns(3)
with col1: include_basic = st.checkbox("기본료 포함 (15,000원)", value=True)
with col2: include_shipping = st.checkbox("우송료 포함 (3,500원)", value=False)
with col3: include_vat = st.checkbox("부가세 포함", value=True)

manual_test_name = st.text_input("수기 시험명")
manual_price = st.number_input("수기 시험 가격", min_value=0, value=0)

# 견적 생성 로직
if st.button("견적 생성", type="primary"):
    detail_rows = []
    pure_subtotal = 0
    
    for selected in selected_tests:
        row = df_price[df_price["표시명"] == selected].iloc[0]
        price = int(row["기본가격"])
        detail_rows.append({"name": selected, "price": price})
        pure_subtotal += price * quantity
        
    if manual_test_name:
        detail_rows.append({"name": manual_test_name, "price": int(manual_price)})
        pure_subtotal += int(manual_price) * quantity
        
    discount = int(pure_subtotal * (discount_rate / 100))
    extra_rate = 0.5 if "지급" in service_type else (1.5 if "즉시" in service_type else 0)
    extra_fee = int((pure_subtotal - discount) * extra_rate)
    etc_fee = (15000 if include_basic else 0) + (3500 if include_shipping else 0)
    subtotal = (pure_subtotal - discount) + extra_fee + etc_fee
    vat = int(subtotal * 0.1) if include_vat else 0
    total = subtotal + vat

    calc_results = {
        "pure": pure_subtotal, 
        "discount": discount, 
        "subtotal_after": pure_subtotal - discount,
        "extra": extra_fee, 
        "extra_rate": int(extra_rate*100), 
        "etc": etc_fee, 
        "vat": vat, 
        "total": total
    }

    client_info = {"company": selected_company, "manager": manager, "phone": phone, "sample": sample_name}

    try:
        excel_bytes = generate_quotation_excel(
            template_path="견적서템플릿.xlsx",
            detail_rows=detail_rows,
            calc_results=calc_results,
            quantity=quantity,
            discount_rate=discount_rate,
            client_info=client_info
        )
        
        st.success("🎉 견적서가 성공적으로 빌드되었습니다!")
        st.markdown("---")
        st.subheader("💰 견적 즉시 확인 내역")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("항목 합계", f"{calc_results['pure']:,}원")
        c2.metric("총 할인액", f"-{calc_results['discount']:,}원")
        c3.metric("기타 요금 및 부가세", f"{calc_results['etc'] + calc_results['vat']:,}원")
        c4.metric("최종 결제 금액", f"{calc_results['total']:,}원")
        
        # 상세 계산 과정 복구
        with st.expander("📝 상세 계산 산출 내역 보기"):
            st.write(f"• 할인 후 항목수수료: {calc_results['subtotal_after']:,}원")
            st.write(f"• 진행 유형 추가요금 ({calc_results['extra_rate']}%): {calc_results['extra']:,}원")
            st.write(f"• 기본료/우송료 합계: {calc_results['etc']:,}원")
            st.write(f"• 부가세 (10%): {calc_results['vat']:,}원")
        
        st.download_button(
            label="📥 완성된 견적서 다운로드 (Excel)", 
            data=excel_bytes, 
            file_name=f"견적서_{selected_company or '의뢰자'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"엑셀 생성 오류: {e}")