import base64
import io
import json
import math
import os
from datetime import datetime

import openai
import plotly.graph_objects as go
import streamlit as st
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from dotenv import load_dotenv
from PIL import Image

load_dotenv()

st.set_page_config(
    page_title="Eurofins KCTL RF Inspector",
    page_icon="📡",
    layout="wide",
)


# ----------------------------------------------------------------------------
# STEP 3: OpenAI 5.6 Luna Vision 마커 파싱 모듈
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="AI Vision 마커 데이터 추출 중...")
def extract_marker_data(image_bytes: bytes, api_key: str) -> dict:
    """스펙트럼 분석기 화면 우측 상단 Mkr1(녹색 텍스트) 영역을 읽어
    {"freq_val", "freq_unit", "power_val", "power_unit"} 스키마로 반환한다."""
    client = openai.OpenAI(api_key=api_key)

    img_format = (Image.open(io.BytesIO(image_bytes)).format or "PNG").lower()
    b64_str = base64.b64encode(image_bytes).decode("utf-8")
    data_url = f"data:image/{img_format};base64,{b64_str}"

    response = client.chat.completions.create(
        model="gpt-5.6-luna",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "너는 스펙트럼 분석기(Agilent Swept SA) 화면 이미지를 분석하는 RF 계측 "
                    "전문가다. 화면 우측 상단의 녹색 텍스트로 표시된 Mkr1 마커 영역에서 주파수 "
                    "값/단위와 전력 값/단위를 읽어 지정된 JSON 스키마로만 응답하라."
                ),
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Mkr1 마커의 주파수와 전력을 읽어 다음 JSON 형식으로만 응답하라: "
                            '{"freq_val": <number>, "freq_unit": "GHz 또는 MHz", '
                            '"power_val": <number>, "power_unit": "uW"}'
                        ),
                    },
                    {
                        "type": "image_url",
                        "image_url": {"url": data_url},
                    },
                ],
            },
        ],
    )

    return json.loads(response.choices[0].message.content)


# ----------------------------------------------------------------------------
# STEP 3: RF 물리량 계산 엔진
# ----------------------------------------------------------------------------
def compute_rf_metrics(marker: dict, cf_db: float, limit_dbuv: float) -> dict:
    """마커 데이터(freq_val/freq_unit/power_val)와 CF, Limit으로 RF 판정 지표를 계산한다."""
    freq_val = marker["freq_val"]
    freq_unit = str(marker["freq_unit"]).strip().upper()
    power_uw = marker["power_val"]

    freq_mhz = freq_val * 1000 if freq_unit == "GHZ" else freq_val
    dbm = 10 * math.log10(power_uw / 1000)
    dbuv_m = dbm + 107.0 + cf_db
    margin = limit_dbuv - dbuv_m
    verdict = "PASS" if margin >= 0 else "FAIL"

    return {
        "freq_mhz": round(freq_mhz, 2),
        "power_uw": round(power_uw, 2),
        "dbm": round(dbm, 2),
        "dbuv_m": round(dbuv_m, 2),
        "limit": round(limit_dbuv, 2),
        "margin": round(margin, 2),
        "verdict": verdict,
    }


# ----------------------------------------------------------------------------
# STEP 5: docxtpl 워드 성적서 인메모리 렌더링
# ----------------------------------------------------------------------------
def build_report_buffer(context: dict, results: list) -> io.BytesIO:
    """report_template.docx를 컨텍스트로 렌더링해 메모리 버퍼로 반환한다.
    각 측정 결과의 원본 이미지는 InlineImage로 성적서 내 이미지 표에 삽입된다."""
    doc = DocxTemplate("report_template.docx")

    context = dict(context)
    context["r"] = [
        {
            "no": i + 1,
            "freq_mhz": r["freq_mhz"],
            "power_uw": r["power_uw"],
            "dbm": r["dbm"],
            "dbuv_m": r["dbuv_m"],
            "limit": r["limit"],
            "margin": r["margin"],
            "verdict": r["verdict"],
            "image_name": r["image_name"],
            "photo": InlineImage(doc, io.BytesIO(r["image_bytes"]), width=Mm(35)),
        }
        for i, r in enumerate(results)
    ]

    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------------
with st.sidebar:
    st.title("📡 KCTL RF 계측 Inspector")

    api_key = st.text_input(
        "OpenAI API Key",
        type="password",
        value=os.getenv("OPENAI_API_KEY", ""),
    )
    st.caption("AI Engine: OpenAI 5.6 Luna Vision")

    st.divider()

    tester = st.text_input("시험 담당자", value="홍길동 선임연구원")
    reviewer = st.text_input("기술 검토자", value="김선임 기술책임자")
    sample_name = st.text_input("시료명 (EUT)", value="EUT-2026-BLE-MODULE")

    standard = st.selectbox(
        "시험 규격",
        [
            "FCC Part 15 Subpart B Class B (3 m)",
            "CISPR 32 Class B (3 m)",
            "KN 32",
        ],
    )

    cf_db = st.slider(
        "안테나 보정계수 (CF, dB)",
        min_value=15.0,
        max_value=40.0,
        value=28.5,
        step=0.5,
    )

    limit_dbuv = st.slider(
        "규격 기준치 (Limit, dBµV/m)",
        min_value=120.0,
        max_value=150.0,
        value=140.0,
        step=1.0,
    )

# ----------------------------------------------------------------------------
# 메인 상단 - 대제목 및 설명 배너
# ----------------------------------------------------------------------------
st.title("📡 Eurofins KCTL RF 계측 자동화 대시보드")
st.markdown(
    "스펙트럼 분석기(Agilent Swept SA) 캡처 이미지를 업로드하면 AI Vision이 "
    "마커(Mkr1) 데이터를 추출하고, 규격 판정 및 공식 워드 성적서를 자동 발행합니다."
)

# ----------------------------------------------------------------------------
# 메인 상단 - 이미지 업로더 & 샘플 로드 버튼
# ----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "스펙트럼 분석기 계측 이미지 업로드",
    accept_multiple_files=True,
    type=["png", "jpg"],
)

if "use_sample_images" not in st.session_state:
    st.session_state["use_sample_images"] = False

if st.button("📂 기본 샘플 3종(img_1~3) 일괄 불러오기"):
    st.session_state["use_sample_images"] = True

images_to_process = []

if uploaded_files:
    for f in uploaded_files:
        images_to_process.append((f.name, f.getvalue()))

if st.session_state["use_sample_images"]:
    st.info("기본 샘플 이미지 3종(img_1.png, img_2.png, img_3.png)이 로드되었습니다.")
    for fname in ("img_1.png", "img_2.png", "img_3.png"):
        with open(fname, "rb") as fp:
            images_to_process.append((fname, fp.read()))

if images_to_process:
    st.subheader("🖼️ 분석 대상 이미지")
    preview_cols = st.columns(min(len(images_to_process), 4))
    for idx, (name, img_bytes) in enumerate(images_to_process):
        with preview_cols[idx % len(preview_cols)]:
            st.image(img_bytes, caption=name, use_container_width=True)

# ----------------------------------------------------------------------------
# STEP 3: Vision 파싱 + RF 계산 파이프라인 실행
# ----------------------------------------------------------------------------
if images_to_process:
    if not api_key:
        st.warning("사이드바에 OpenAI API Key를 입력하세요.")
        st.stop()

    results = []
    for name, img_bytes in images_to_process:
        try:
            marker = extract_marker_data(img_bytes, api_key)
            metrics = compute_rf_metrics(marker, cf_db, limit_dbuv)
            metrics["image_name"] = name
            metrics["image_bytes"] = img_bytes
            results.append(metrics)
        except Exception as e:
            st.error(f"오류 내용: {e}")

    st.session_state["analysis_results"] = results

# ----------------------------------------------------------------------------
# STEP 4: KPI 요약 카드 & Plotly 규격 비교 차트
# ----------------------------------------------------------------------------
results = st.session_state.get("analysis_results", [])

if results:
    total_cnt = len(results)
    pass_cnt = sum(1 for r in results if r["verdict"] == "PASS")
    fail_cnt = sum(1 for r in results if r["verdict"] == "FAIL")
    overall_verdict = "PASS" if fail_cnt == 0 else "FAIL"

    st.divider()
    st.subheader("📊 KPI 요약")

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("총 측정 건수", total_cnt)
    kpi2.metric("PASS 건수", pass_cnt, delta=f"{pass_cnt}건", delta_color="normal")
    kpi3.metric("FAIL 건수", fail_cnt, delta=f"{fail_cnt}건", delta_color="inverse")
    kpi4.metric(
        "종합 판정",
        overall_verdict,
        delta="규격 만족" if overall_verdict == "PASS" else "규격 위반",
        delta_color="normal" if overall_verdict == "PASS" else "inverse",
    )

    st.subheader("📈 규격 비교 차트")

    pass_rows = [r for r in results if r["verdict"] == "PASS"]
    fail_rows = [r for r in results if r["verdict"] == "FAIL"]

    fig = go.Figure()

    if pass_rows:
        fig.add_trace(
            go.Bar(
                x=[r["freq_mhz"] for r in pass_rows],
                y=[r["dbuv_m"] for r in pass_rows],
                name="PASS",
                marker_color="#003399",
                customdata=[r["margin"] for r in pass_rows],
                hovertemplate=(
                    "주파수: %{x} MHz<br>측정값: %{y} dBµV/m<br>"
                    "마진: %{customdata:.2f} dB<extra>PASS</extra>"
                ),
            )
        )

    if fail_rows:
        fig.add_trace(
            go.Bar(
                x=[r["freq_mhz"] for r in fail_rows],
                y=[r["dbuv_m"] for r in fail_rows],
                name="FAIL",
                marker_color="#E03A3A",
                customdata=[r["margin"] for r in fail_rows],
                hovertemplate=(
                    "주파수: %{x} MHz<br>측정값: %{y} dBµV/m<br>"
                    "마진: %{customdata:.2f} dB<extra>FAIL</extra>"
                ),
            )
        )

    fig.add_hline(
        y=limit_dbuv,
        line_dash="dash",
        line_color="red",
        annotation_text="FCC Limit 기준",
        annotation_position="top left",
    )

    fig.update_layout(
        xaxis_title="주파수 (MHz)",
        yaxis_title="측정 전계강도 (dBµV/m)",
        legend_title="판정",
    )

    st.plotly_chart(fig, use_container_width=True)

    # ------------------------------------------------------------------------
    # STEP 5: docxtpl 워드 성적서 다운로드 연동
    # ------------------------------------------------------------------------
    st.subheader("📄 공식 시험성적서 발행")

    doc_no = st.text_input("문서번호", value=f"KCTL-RE-{datetime.now():%Y%m%d}-01")
    remarks = st.text_area("특이사항", value="")

    report_context = {
        "doc_no": doc_no,
        "test_date": datetime.now().strftime("%Y-%m-%d"),
        "test_name": "방사성 방출 (RE) 측정 결과 보고서",
        "tester": tester,
        "standard": standard,
        "sample_name": sample_name,
        "cf_db": cf_db,
        "reviewer": reviewer,
        "remarks": remarks,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total": total_cnt,
        "pass_cnt": pass_cnt,
        "fail_cnt": fail_cnt,
        "verdict": overall_verdict,
    }

    try:
        report_buffer = build_report_buffer(report_context, results)
        st.download_button(
            label="📥 공식 시험성적서(.docx) 다운로드",
            data=report_buffer,
            file_name=f"KCTL_RE_Report_{datetime.now():%Y%m%d}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"오류 내용: {e}")
