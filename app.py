import streamlit as st
import pandas as pd
from excel_generator import generate_quotation_excel

# 페이지 설정
st.set_page_config(page_title="시험 견적 시스템", layout="wide")
st.image("logo.png", width=300)
st.subheader("###시험견적시스템")
st.write("")

# 데이터 로드 (업체DB는 더 이상 사용하지 않음)
@st.cache_data
def load_data():
    df_p = pd.read_excel("시험수수료표.xlsx", sheet_name="수수료표")
    df_p["표시명"] = df_p.apply(
        lambda row: f"{row['시험명']} / {row['시험방법']}"
        if pd.notna(row['시험방법']) and row['시험방법'] != "" else row['시험명'],
        axis=1,
    )
    return df_p

df_price = load_data()

# UI 구성
col_a, col_b = st.columns(2)
with col_a:
    company_name = st.text_input("업체명")
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

# 수기 시험 (여러 개 추가 가능)
st.subheader("수기 시험")
st.caption("표 아래 ➕ 를 눌러 행을 추가하고, 행을 선택해 🗑️ 로 삭제할 수 있습니다.")
manual_df = st.data_editor(
    pd.DataFrame({"시험명": pd.Series(dtype="str"), "가격": pd.Series(dtype="int")}),
    num_rows="dynamic",
    use_container_width=True,
    key="manual_tests",
    column_config={
        "시험명": st.column_config.TextColumn("시험명", required=True),
        "가격": st.column_config.NumberColumn("가격 (원)", min_value=0, step=1000, format="%d", required=True),
    },
)

# 견적 생성 로직
if st.button("견적 생성", type="primary"):
    detail_rows = []
    pure_subtotal = 0

    for selected in selected_tests:
        row = df_price[df_price["표시명"] == selected].iloc[0]
        price = int(row["기본가격"])
        detail_rows.append({"name": selected, "price": price})
        pure_subtotal += price * quantity

    for _, m in manual_df.iterrows():
        name = str(m["시험명"]).strip() if pd.notna(m["시험명"]) else ""
        if not name:
            continue
        price = int(m["가격"]) if pd.notna(m["가격"]) else 0
        detail_rows.append({"name": name, "price": price})
        pure_subtotal += price * quantity

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
        "extra_rate": int(extra_rate * 100),
        "etc": etc_fee,
        "vat": vat,
        "total": total,
    }

    client_info = {"company": company_name, "manager": manager, "phone": phone, "sample": sample_name}

    try:
        excel_bytes = generate_quotation_excel(
            template_path="견적서템플릿.xlsx",
            detail_rows=detail_rows,
            calc_results=calc_results,
            quantity=quantity,
            discount_rate=discount_rate,
            client_info=client_info,
        )

        st.success("🎉 견적서가 성공적으로 빌드되었습니다!")
        st.markdown("---")
        st.subheader("💰 견적 즉시 확인 내역")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("항목 합계", f"{calc_results['pure']:,}원")
        c2.metric("총 할인액", f"-{calc_results['discount']:,}원")
        c3.metric("기타 요금 및 부가세", f"{calc_results['etc'] + calc_results['vat']:,}원")
        c4.metric("최종 결제 금액", f"{calc_results['total']:,}원")

        with st.expander("📝 상세 계산 산출 내역 보기"):
            st.write(f"• 할인 후 항목수수료: {calc_results['subtotal_after']:,}원")
            st.write(f"• 진행 유형 추가요금 ({calc_results['extra_rate']}%): {calc_results['extra']:,}원")
            st.write(f"• 기본료/우송료 합계: {calc_results['etc']:,}원")
            st.write(f"• 부가세 (10%): {calc_results['vat']:,}원")

        st.download_button(
            label="📥 완성된 견적서 다운로드 (Excel)",
            data=excel_bytes,
            file_name=f"견적서_{company_name.strip() or '의뢰자'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"엑셀 생성 오류: {e}")
