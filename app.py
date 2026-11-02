from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill


st.set_page_config(
    page_title="스탠다드 뱅크 | 인증 프로젝트 현황",
    page_icon="📋",
    layout="wide",
)

REQUIRED_COLUMNS = [
    "의뢰번호",
    "고객사",
    "제품명",
    "모델명",
    "시험분야",
    "적용규격",
    "담당엔지니어",
    "접수일",
    "완료예정일",
    "진행상태",
    "시료상태",
    "메모",
]
FILTER_COLUMNS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
INITIAL_FIELDS = ["EMC", "무선(RF)", "인증서류", "SAR"]
DATA_DIRECTORY = Path(__file__).resolve().parent / "data"
DATA_FILES = [
    DATA_DIRECTORY / "SB_projects_demo.csv",
    DATA_DIRECTORY / "SB_project_demo.csv",
]
MISSING_VALUE_LABEL = "(미입력)"


def read_csv(source) -> pd.DataFrame:
    try:
        return pd.read_csv(source, encoding="utf-8-sig")
    except UnicodeDecodeError:
        if hasattr(source, "seek"):
            source.seek(0)
        return pd.read_csv(source, encoding="cp949")


def add_schedule_columns(projects: pd.DataFrame, reference_date: date) -> pd.DataFrame:
    result = projects.copy()
    result["접수일"] = pd.to_datetime(result["접수일"], errors="coerce")
    result["완료예정일"] = pd.to_datetime(result["완료예정일"], errors="coerce")
    reference_timestamp = pd.Timestamp(reference_date)
    remaining_days = (result["완료예정일"] - reference_timestamp).dt.days.astype("Int64")
    result["잔여일수"] = remaining_days

    schedule_status = pd.Series("진행 예정", index=result.index, dtype="string")
    schedule_status.loc[result["완료예정일"].isna()] = "일정 확인 필요"
    schedule_status.loc[remaining_days.lt(0).fillna(False)] = "지연 후보(일정 확인 필요)"
    schedule_status.loc[remaining_days.eq(0).fillna(False)] = "오늘 마감"
    schedule_status.loc[result["진행상태"].eq("완료")] = "완료"
    result["일정 상태"] = schedule_status
    return result


def highlight_overdue_row(row: pd.Series) -> list[str]:
    if row["일정 상태"] == "지연 후보(일정 확인 필요)":
        return ["background-color: #fde8e8; color: #842029"] * len(row)
    return [""] * len(row)


def build_excel_report(
    projects: pd.DataFrame,
    reference_date: date,
    filter_summary: dict[str, str],
) -> bytes:
    report_projects = add_schedule_columns(projects, reference_date)
    status_summary = (
        report_projects["진행상태"]
        .fillna(MISSING_VALUE_LABEL)
        .value_counts()
        .rename_axis("진행상태")
        .reset_index(name="건수")
    )
    owner_summary = (
        report_projects["담당엔지니어"]
        .fillna(MISSING_VALUE_LABEL)
        .value_counts()
        .rename_axis("담당엔지니어")
        .reset_index(name="건수")
    )
    conditions = pd.DataFrame(
        [
            {"항목": "기준일", "값": reference_date.isoformat()},
            *[{"항목": column, "값": value} for column, value in filter_summary.items()],
        ]
    )
    sheets = {
        "프로젝트목록": report_projects,
        "상태별집계": status_summary,
        "담당자별집계": owner_summary,
        "분석조건": conditions,
    }

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)

    output.seek(0)
    with pd.ExcelWriter(output, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        workbook = writer.book
        for worksheet in workbook.worksheets:
            last_column = max(worksheet.max_column, 1)
            worksheet.merge_cells(
                start_row=1,
                start_column=1,
                end_row=1,
                end_column=last_column,
            )
            title = worksheet.cell(1, 1, "교육용 검토 파일 — 공식 보고서 아님")
            title.font = Font(bold=True, color="FFFFFF", size=12)
            title.fill = PatternFill("solid", fgColor="17365D")
            title.alignment = Alignment(horizontal="left")
            worksheet.cell(
                2,
                1,
                f"생성일: {date.today():%Y-%m-%d} | 기준일: {reference_date:%Y-%m-%d}",
            )
            for cell in worksheet[3]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2F6FCA")
            worksheet.freeze_panes = "A4"
            worksheet.auto_filter.ref = (
                f"A3:{worksheet.cell(worksheet.max_row, worksheet.max_column).coordinate}"
            )
            for column_cells in worksheet.columns:
                letter = worksheet.cell(3, column_cells[0].column).column_letter
                width = min(
                    max(
                        len(str(cell.value)) if cell.value is not None else 0
                        for cell in column_cells
                    )
                    + 2,
                    42,
                )
                worksheet.column_dimensions[letter].width = max(width, 12)

    output.seek(0)
    return output.getvalue()


st.markdown(
    """
    <style>
    .stApp { background: #ffffff; color: #172b4d; }
    [data-testid="stHeader"] { background: rgba(255, 255, 255, 0.94); }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1500px; }
    .notice {
        padding: 0.7rem 1rem; margin-bottom: 1.25rem;
        border: 1px solid #c9d9f2; border-radius: 0.65rem;
        background: #edf4ff; color: #234a7c; font-size: 0.92rem; font-weight: 600;
    }
    .eyebrow {
        color: #52739d; font-size: 0.82rem; font-weight: 700;
        letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 0.25rem;
    }
    h1 { color: #102a50; letter-spacing: -0.035em; margin-bottom: 0.15rem; }
    .subtitle { color: #64748b; margin-bottom: 1.5rem; }
    [data-testid="stMetric"] {
        background: #ffffff; border: 1px solid #e1e9f4; border-top: 3px solid #2f6fca;
        border-radius: 0.75rem; padding: 0.9rem 1rem;
        box-shadow: 0 3px 12px rgba(20, 48, 85, 0.04);
    }
    [data-testid="stMetricLabel"] { color: #52657e; font-weight: 600; }
    [data-testid="stMetricValue"] { color: #15345d; }
    [data-testid="stDataFrame"] {
        border: 1px solid #e1e9f4; border-radius: 0.7rem; overflow: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="notice">교육용 가상 데이터 기준 — 실제 업무 시스템 아님</div>',
    unsafe_allow_html=True,
)
st.caption("카테고리와 데이터 변경은 현재 세션에만 저장되며, 세션 종료 시 사라질 수 있습니다.")
st.markdown('<div class="eyebrow">STANDARD BANK · CERTIFICATION</div>', unsafe_allow_html=True)
st.title("인증 프로젝트 접수 · 진행현황")
st.markdown(
    '<div class="subtitle">접수된 인증 프로젝트와 시험 진행 상태를 한눈에 확인합니다.</div>',
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "CSV 파일 업로드",
    type=["csv"],
    help="기본 파일을 찾을 수 없거나 다른 데이터를 확인할 때 CSV를 업로드하세요.",
)

data_source = uploaded_file or next((path for path in DATA_FILES if path.is_file()), None)
if data_source is None:
    st.info(
        "데이터 파일을 찾을 수 없습니다. `app.py`와 같은 위치에 `data` 폴더를 만들고 "
        "`SB_projects_demo.csv`를 넣거나, 위에서 CSV 파일을 업로드해 주세요."
    )
    st.stop()

try:
    projects = read_csv(data_source)
except (UnicodeDecodeError, pd.errors.ParserError, pd.errors.EmptyDataError) as error:
    st.error(f"CSV 파일을 읽을 수 없습니다. 인코딩과 CSV 형식을 확인해 주세요. ({error})")
    st.stop()

missing_columns = [column for column in REQUIRED_COLUMNS if column not in projects.columns]
if missing_columns:
    st.error(f"CSV에 필수 항목이 없습니다: {', '.join(missing_columns)}")
    st.stop()

if "project_data" not in st.session_state:
    initial_projects = projects[REQUIRED_COLUMNS].copy()
    existing_fields = initial_projects["시험분야"].dropna().astype(str).unique().tolist()
    st.session_state["project_data"] = initial_projects
    st.session_state["exam_fields"] = list(dict.fromkeys(INITIAL_FIELDS + existing_fields))

projects = st.session_state["project_data"].copy()
fields = st.session_state["exam_fields"]

with st.sidebar:
    st.header("카테고리 관리")
    st.caption("시험분야 등록, 이름 변경, 삭제")

    with st.form("register_field_form", clear_on_submit=True):
        new_field_name = st.text_input("새 시험분야")
        register_field = st.form_submit_button("시험분야 등록")
    if register_field:
        normalized_name = new_field_name.strip()
        if not normalized_name:
            st.warning("등록할 시험분야 이름을 입력해 주세요.")
        elif normalized_name in fields:
            st.warning("이미 등록된 시험분야입니다.")
        else:
            selected_fields = st.session_state.get("filter_시험분야", fields)
            all_fields_selected = set(selected_fields) == set(fields)
            fields.append(normalized_name)
            st.session_state["exam_fields"] = fields
            if all_fields_selected:
                st.session_state["filter_시험분야"] = selected_fields + [normalized_name]
            st.success(f"'{normalized_name}' 시험분야를 등록했습니다.")

    with st.form("rename_field_form", clear_on_submit=True):
        rename_from = st.selectbox("변경할 시험분야", options=fields, key="rename_field_from")
        rename_to = st.text_input("새 이름", key="rename_field_to")
        rename_field = st.form_submit_button("이름 변경")
    if rename_field:
        normalized_name = rename_to.strip()
        if not normalized_name:
            st.warning("변경할 이름을 입력해 주세요.")
        elif normalized_name != rename_from and normalized_name in fields:
            st.warning("이미 등록된 시험분야 이름입니다.")
        elif normalized_name != rename_from:
            st.session_state["project_data"].loc[
                st.session_state["project_data"]["시험분야"].eq(rename_from), "시험분야"
            ] = normalized_name
            st.session_state["exam_fields"] = [
                normalized_name if field == rename_from else field for field in fields
            ]
            selected_fields = st.session_state.get("filter_시험분야", fields)
            st.session_state["filter_시험분야"] = [
                normalized_name if field == rename_from else field for field in selected_fields
            ]
            fields = st.session_state["exam_fields"]
            st.success(f"'{rename_from}'을(를) '{normalized_name}'(으)로 변경했습니다.")

    reclassifiable_fields = [
        field
        for field in fields
        if st.session_state["project_data"]["시험분야"].eq(field).any()
    ]
    with st.form("reclassify_field_form", clear_on_submit=True):
        if reclassifiable_fields:
            reclassify_from = st.selectbox(
                "재분류할 시험분야",
                options=reclassifiable_fields,
                key="reclassify_field_from",
            )
            reclassify_targets = [field for field in fields if field != reclassify_from]
            reclassify_to = st.selectbox(
                "이동할 시험분야",
                options=reclassify_targets,
                key="reclassify_field_to",
            )
        else:
            reclassify_from = None
            reclassify_to = None
            st.caption("재분류할 프로젝트가 없습니다.")
        reclassify_field = st.form_submit_button(
            "의뢰 데이터를 다른 분야로 재분류",
            disabled=not reclassifiable_fields or len(fields) < 2,
        )
    if reclassify_field and reclassify_from and reclassify_to:
        st.session_state["project_data"].loc[
            st.session_state["project_data"]["시험분야"].eq(reclassify_from), "시험분야"
        ] = reclassify_to
        selected_fields = st.session_state.get("filter_시험분야", fields)
        st.session_state["filter_시험분야"] = list(
            dict.fromkeys(
                reclassify_to if field == reclassify_from else field
                for field in selected_fields
            )
        )
        st.success(f"'{reclassify_from}' 의뢰를 '{reclassify_to}'로 재분류했습니다.")

    with st.form("delete_field_form", clear_on_submit=True):
        delete_field_name = st.selectbox("삭제할 시험분야", options=fields, key="delete_field_name")
        delete_field = st.form_submit_button("시험분야 삭제")
    if delete_field:
        in_use = st.session_state["project_data"]["시험분야"].eq(delete_field_name).any()
        if in_use:
            st.error(
                f"'{delete_field_name}'은(는) 의뢰 데이터에서 사용 중입니다. "
                "먼저 의뢰 데이터를 다른 시험분야로 재분류한 뒤 삭제해 주세요."
            )
        else:
            st.session_state["exam_fields"] = [
                field for field in fields if field != delete_field_name
            ]
            st.session_state["filter_시험분야"] = [
                field
                for field in st.session_state.get("filter_시험분야", fields)
                if field != delete_field_name
            ]
            fields = st.session_state["exam_fields"]
            st.success(f"'{delete_field_name}' 시험분야를 삭제했습니다.")

    projects = st.session_state["project_data"].copy()
    st.divider()
    st.header("필터")
    selections = {}
    options_by_column = {}
    for column in FILTER_COLUMNS:
        if column == "시험분야":
            options = fields
        else:
            values = projects[column].astype("string").fillna(MISSING_VALUE_LABEL)
            options = sorted(values.unique().tolist())
        options_by_column[column] = options
        filter_key = f"filter_{column}"
        selected_options = st.session_state.get(filter_key, options)
        st.session_state[filter_key] = [
            value for value in selected_options if value in options
        ]
        selections[column] = st.multiselect(
            column,
            options=options,
            key=filter_key,
        )

reference_date = st.date_input(
    "기준일",
    value=date(2026, 10, 23),
    help="이 날짜를 기준으로 잔여일수와 일정 상태를 계산합니다.",
)

filtered_projects = projects.copy()
for column, selected_values in selections.items():
    normalized_values = filtered_projects[column].astype("string").fillna(MISSING_VALUE_LABEL)
    filtered_projects = filtered_projects[normalized_values.isin(selected_values)]

filtered_projects = add_schedule_columns(filtered_projects, reference_date)
status_counts = filtered_projects["진행상태"].value_counts()
summary_columns = st.columns(1 + len(STATUSES))
with summary_columns[0]:
    st.metric("전체 의뢰", f"{len(filtered_projects):,}건")
for column, status in zip(summary_columns[1:], STATUSES):
    with column:
        st.metric(status, f"{int(status_counts.get(status, 0)):,}건")

overdue_count = int(filtered_projects["일정 상태"].eq("지연 후보(일정 확인 필요)").sum())
st.caption(
    f"기준일 {reference_date:%Y.%m.%d} 기준 · 필터 결과 {len(filtered_projects):,}건 · "
    f"지연 후보 {overdue_count:,}건"
)

filter_summary = {}
for column, selected_values in selections.items():
    if len(selected_values) == len(options_by_column[column]):
        filter_summary[column] = "전체"
    elif selected_values:
        filter_summary[column] = ", ".join(selected_values)
    else:
        filter_summary[column] = "(선택 없음)"

today_stamp = datetime.now().strftime("%Y%m%d")
all_report = build_excel_report(
    projects,
    reference_date,
    {column: "전체" for column in FILTER_COLUMNS},
)
filtered_report = build_excel_report(filtered_projects, reference_date, filter_summary)
download_columns = st.columns(2)
with download_columns[0]:
    st.download_button(
        "전체 목록 Excel 다운로드",
        data=all_report,
        file_name=f"standardbank_project_status_all_{today_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_all_excel",
    )
with download_columns[1]:
    st.download_button(
        "현재 필터 결과 Excel 다운로드",
        data=filtered_report,
        file_name=f"standardbank_project_status_filtered_{today_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_filtered_excel",
    )

display_projects = filtered_projects.copy()
display_projects["접수일"] = display_projects["접수일"].dt.strftime("%Y-%m-%d").fillna("")
display_projects["완료예정일"] = (
    display_projects["완료예정일"].dt.strftime("%Y-%m-%d").fillna("")
)

st.subheader("전체 의뢰 목록")
if filtered_projects.empty:
    st.info("조건에 맞는 인증 프로젝트가 없습니다")
else:
    chart_columns = st.columns(2)
    with chart_columns[0]:
        st.subheader("시험분야별 의뢰 건수")
        field_counts = filtered_projects["시험분야"].fillna(MISSING_VALUE_LABEL).value_counts()
        st.bar_chart(field_counts.rename("건수"), width="stretch")
    with chart_columns[1]:
        st.subheader("진행상태 비율")
        progress_counts = (
            filtered_projects["진행상태"]
            .fillna(MISSING_VALUE_LABEL)
            .value_counts()
            .rename_axis("진행상태")
            .reset_index(name="건수")
        )
        st.vega_lite_chart(
            progress_counts,
            {
                "mark": {"type": "arc", "innerRadius": 55},
                "encoding": {
                    "theta": {"field": "건수", "type": "quantitative"},
                    "color": {"field": "진행상태", "type": "nominal"},
                    "tooltip": [
                        {"field": "진행상태", "type": "nominal"},
                        {"field": "건수", "type": "quantitative"},
                    ],
                },
                "width": "container",
                "height": 280,
            },
            width="stretch",
        )

    st.dataframe(
        display_projects.style.apply(highlight_overdue_row, axis=1),
        width="stretch",
        hide_index=True,
        column_config={
            "잔여일수": st.column_config.NumberColumn("잔여일수", format="%d일"),
            "일정 상태": st.column_config.TextColumn("일정 상태"),
            "메모": st.column_config.TextColumn("메모", width="large"),
        },
    )
