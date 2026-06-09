import io
import os
import threading
from datetime import date, datetime, timedelta
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import requests
import streamlit as st
from streamlit_javascript import st_javascript

KST = ZoneInfo("Asia/Seoul")
SPREADSHEET_ID = "13T5pAuqmkF3QEYXSEf0L-cB-shDbp8X2-BBUfsdGl2s"
WORKSHEET_NAME = "근무기록"
MODIFICATION_REQUEST_WORKSHEET_NAME = "수정요청"
WEEKLY_TARGET_HOURS = 8.0
WEEKLY_TARGET_MINUTES = int(WEEKLY_TARGET_HOURS * 60)
WORK_DETAIL_KEY = "work_detail_input"
FALLBACK_TIME_KEY = "_using_local_time_fallback"
DF_SESSION_KEY = "df"
FILTER_ANCHOR_KEY = "_filter_time_anchor"
FILTER_YEAR_KEY = "filter_year"
FILTER_MONTH_KEY = "filter_month"
FILTER_WEEK_KEY = "filter_week"
FILTER_WEEK_ALL = 0  # '전체' 선택 시 주차 필터 bypass용 정수 센티널
SECRETS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "secrets.json")

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_COLUMNS = [
    "연월",
    "월 주차",
    "날짜",
    "출근시간",
    "퇴근시간",
    "당일 근무시간",
    "업무 내용",
]
# UI 표시 전용 — SHEET_COLUMNS(백엔드/시트)와 분리, '연월'은 화면에서만 숨김
DISPLAY_COLUMNS = [
    "월 주차",
    "날짜",
    "출근시간",
    "퇴근시간",
    "당일 근무시간",
    "업무 내용",
]
INTERNAL_COLUMNS = ["연도", "월", "주차", "_week_start", "_sort_date", "_sheet_row"]
SHEET_HEADER_RANGE = "A1:G1"
MODIFICATION_REQUEST_COLUMNS = [
    "요청 일시",
    "대상 날짜",
    "기존 출근시간",
    "기존 퇴근시간",
    "요청 내용",
    "처리 상태",
]
MODIFICATION_REQUEST_HEADER_RANGE = "A1:F1"

ALLOWED_OFFICE_IPS = ["106.251.83.219"]


def get_client_ip():
    # 브라우저 단에서 api.ipify.org를 호출하여 실제 외부 공인 IP를 가져옴
    ip = st_javascript("await fetch('https://api.ipify.org?format=json').then(r => r.json()).then(d => d.ip)")
    return ip


@st.cache_resource
def get_gspread_client():
    """
    [하이브리드 인증 모듈]
    1순위: Streamlit Cloud의 금고(st.secrets)에서 인증 키 탐색
    2순위: 로컬 PC의 물리적 파일(secrets.json) 탐색
    """
    try:
        # Streamlit Cloud 환경 우선 탐색
        creds_dict = dict(st.secrets["gcp_service_account"])
        # 환경에 따른 이스케이프 문자(줄바꿈) 붕괴 방지 처리
        if "\\n" in creds_dict.get("private_key", ""):
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
    except Exception:
        # 로컬 환경으로 Fallback (예외 발생 시 기존 secrets.json 사용)
        return gspread.service_account(filename=SECRETS_PATH, scopes=SCOPES)


@st.cache_resource
def get_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=7)
        worksheet.update(SHEET_HEADER_RANGE, [SHEET_COLUMNS])

    return worksheet


@st.cache_resource
def get_modification_request_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        worksheet = spreadsheet.worksheet(MODIFICATION_REQUEST_WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=MODIFICATION_REQUEST_WORKSHEET_NAME,
            rows=1000,
            cols=len(MODIFICATION_REQUEST_COLUMNS),
        )
        worksheet.update(MODIFICATION_REQUEST_HEADER_RANGE, [MODIFICATION_REQUEST_COLUMNS])

    return worksheet


def ensure_modification_request_headers(worksheet) -> None:
    rows = worksheet.get_all_values()
    if not rows:
        worksheet.update(MODIFICATION_REQUEST_HEADER_RANGE, [MODIFICATION_REQUEST_COLUMNS])
        return

    header = rows[0][: len(MODIFICATION_REQUEST_COLUMNS)]
    if header != MODIFICATION_REQUEST_COLUMNS:
        worksheet.update(MODIFICATION_REQUEST_HEADER_RANGE, [MODIFICATION_REQUEST_COLUMNS])


def ensure_sheet_headers(worksheet, rows: list[list[str]]) -> None:
    if not rows:
        worksheet.update(SHEET_HEADER_RANGE, [SHEET_COLUMNS])
        return

    header = rows[0][: len(SHEET_COLUMNS)]
    if header != SHEET_COLUMNS:
        worksheet.update(SHEET_HEADER_RANGE, [SHEET_COLUMNS])


def fetch_kst_from_server() -> datetime | None:
    """네이버 응답 Date 헤더에서 KST 시각을 가져온다."""
    try:
        response = requests.head("https://www.naver.com", timeout=0.5)
        date_header = response.headers.get("Date")
        if not date_header:
            return None
        utc_dt = parsedate_to_datetime(date_header)
        return utc_dt.astimezone(KST)
    except (requests.RequestException, TypeError, ValueError, OverflowError):
        return None


def now_kst() -> datetime:
    server_time = fetch_kst_from_server()
    if server_time is not None:
        st.session_state[FALLBACK_TIME_KEY] = False
        return server_time

    st.session_state[FALLBACK_TIME_KEY] = True
    return datetime.now(KST)


def today_kst() -> date:
    return now_kst().date()


def format_date_with_weekday(record_date: date) -> str:
    return f"{record_date.isoformat()} ({WEEKDAY_KO[record_date.weekday()]})"


def today_str() -> str:
    return format_date_with_weekday(today_kst())


def time_str(dt: datetime) -> str:
    return dt.strftime("%H:%M:%S")


def format_time_short(value) -> str:
    if pd.isna(value) or not str(value).strip():
        return "-"
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).strftime("%H:%M")
        except ValueError:
            continue
    return text


def request_datetime_str() -> str:
    return now_kst().strftime("%Y-%m-%d %H:%M:%S")


def format_year_month(record_date: date) -> str:
    return f"{record_date.year}{record_date.month:02d}"


def parse_record_date(value) -> date | None:
    if value is None or str(value).strip() == "":
        return None

    text = str(value).strip()
    if "(" in text:
        text = text.split("(")[0].strip()

    try:
        return pd.to_datetime(text, errors="coerce").date()
    except (TypeError, ValueError):
        return None


def get_tuesday_week_start(record_date: date) -> date:
    """화요일~월요일 주의 시작일(화요일)을 반환한다."""
    return record_date - timedelta(days=(record_date.weekday() - 1) % 7)


def first_tuesday_of_month(year: int, month: int) -> date:
    current = date(year, month, 1)
    while current.weekday() != 1:
        current += timedelta(days=1)
    return current


def first_week_start_of_month(year: int, month: int) -> date:
    return get_tuesday_week_start(first_tuesday_of_month(year, month))


def get_month_week_info(record_date: date) -> tuple[int, int, int, str]:
    """
    커스텀 월 주차 규칙:
    - 주는 화~월
    - 매월 1주차는 해당 월 첫 번째 화요일이 포함된 주(화~월)부터 시작
    """
    week_start = get_tuesday_week_start(record_date)
    year = week_start.year
    month = week_start.month

    first_ws = first_week_start_of_month(year, month)
    if week_start < first_ws:
        if month == 1:
            year -= 1
            month = 12
        else:
            month -= 1
            first_ws = first_week_start_of_month(year, month)

    week_number = (week_start - first_ws).days // 7 + 1
    month_week_label = f"{month}월 {week_number}주차"
    return year, month, week_number, month_week_label


def get_today_filter_defaults() -> tuple[int, int, int]:
    """KST(네이버 또는 폴백) 기준 오늘의 연도·월·커스텀 월 주차."""
    today = today_kst()
    year, month, week, _ = get_month_week_info(today)
    return year, month, week


def _inject_filter_option(options: list[int], value: int) -> list[int]:
    if value not in options:
        return sorted(options + [value])
    return sorted(options)


def _sync_time_aware_filter_defaults(today_year: int, today_month: int, today_week: int) -> None:
    anchor = f"{today_year}-{today_month}-{today_week}"
    if st.session_state.get(FILTER_ANCHOR_KEY) == anchor:
        return

    st.session_state[FILTER_ANCHOR_KEY] = anchor
    st.session_state[FILTER_YEAR_KEY] = today_year
    st.session_state[FILTER_MONTH_KEY] = today_month
    st.session_state[FILTER_WEEK_KEY] = today_week


def _ensure_selectbox_value(key: str, options: list[int], preferred: int) -> None:
    if not options:
        return
    if st.session_state.get(key) not in options:
        st.session_state[key] = preferred if preferred in options else options[0]


def get_custom_week_range(reference: date | None = None) -> tuple[date, date]:
    week_start = get_tuesday_week_start(reference or today_kst())
    week_end = week_start + timedelta(days=6)
    return week_start, week_end


def parse_time(value: str, record_date: date | None = None) -> datetime | None:
    if not value or not str(value).strip():
        return None

    base_date = record_date or today_kst()
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.replace(
                year=base_date.year,
                month=base_date.month,
                day=base_date.day,
                tzinfo=KST,
            )
        except ValueError:
            continue
    return None


def parse_work_hours(value) -> float:
    if value is None or str(value).strip() == "":
        return 0.0

    text = str(value).strip()
    if "(" in text and "분" in text:
        try:
            hours_part = text.split("(")[0].strip()
            return float(hours_part)
        except ValueError:
            pass

    if text.endswith("분"):
        try:
            return float(text.replace("분", "").strip()) / 60.0
        except ValueError:
            return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def format_duration(minutes) -> str:
    """전체 분 단위를 'n시간 m분' 형식의 사람 친화적 문자열로 변환한다."""
    total = int(round(minutes)) if minutes is not None else 0
    if total <= 0:
        return "0분"

    hours = total // 60
    mins = total % 60

    if hours == 0:
        return f"{mins}분"
    if mins == 0:
        return f"{hours}시간 0분"
    return f"{hours}시간 {mins}분"


def work_value_to_minutes(value) -> int:
    return int(round(parse_work_hours(value) * 60))


def _format_week_filter_label(value: int | None) -> str:
    """주차 selectbox 표시 — 내부값은 int/None 유지, UI만 라벨 변환."""
    if value in (FILTER_WEEK_ALL, None):
        return "All(전체)"
    return f"{value}주차"


def generate_excel_file(df: pd.DataFrame) -> bytes:
    """필터된 근무 기록을 서식이 적용된 엑셀 바이트로 변환한다."""
    export_df = build_display_dataframe(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = "근무기록"
        export_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
            }
        )
        cell_center_format = workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
            }
        )
        cell_left_format = workbook.add_format(
            {
                "align": "left",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
            }
        )

        # 컬럼별 고정 너비 — 날짜/시간은 좁게, 업무 내용은 넓게
        col_widths = {
            "월 주차": 14,
            "날짜": 22,
            "출근시간": 12,
            "퇴근시간": 12,
            "당일 근무시간": 16,
            "업무 내용": 65,
        }
        for col_idx, col_name in enumerate(export_df.columns):
            worksheet.set_column(col_idx, col_idx, col_widths.get(col_name, 12))

        worksheet.set_row(0, 22)
        for row_idx, col_name in enumerate(export_df.columns):
            worksheet.write(0, row_idx, col_name, header_format)

        for row_idx in range(len(export_df)):
            excel_row = row_idx + 1
            for col_idx, col_name in enumerate(export_df.columns):
                value = export_df.iloc[row_idx, col_idx]
                if pd.isna(value):
                    value = ""
                cell_format = (
                    cell_left_format if col_name == "업무 내용" else cell_center_format
                )
                worksheet.write(excel_row, col_idx, value, cell_format)

    return output.getvalue()


def build_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """화면 출력용 View-Model — st.session_state.df 원본은 건드리지 않는다."""
    display_df = df[DISPLAY_COLUMNS].copy()
    display_df["당일 근무시간"] = df["당일 근무시간"].apply(
        lambda value: format_duration(work_value_to_minutes(value))
        if value is not None and str(value).strip()
        else ""
    )
    return display_df


def sheet_stored_attributes(record_date: date) -> tuple[str, str]:
    return format_year_month(record_date), get_month_week_info(record_date)[3]


def derive_date_attributes(record_date: date) -> dict:
    year, month, week_number, month_week_label = get_month_week_info(record_date)
    week_start = get_tuesday_week_start(record_date)
    return {
        "연도": year,
        "월": month,
        "주차": week_number,
        "월 주차": month_week_label,
        "_week_start": pd.Timestamp(week_start),
    }


def _coalesce_sheet_value(row: pd.Series, column: str, fallback):
    if column not in row.index:
        return fallback
    value = row[column]
    if pd.isna(value) or str(value).strip() == "":
        return fallback
    return value


def enrich_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        empty = pd.DataFrame(columns=SHEET_COLUMNS + INTERNAL_COLUMNS)
        empty["연도"] = empty["연도"].astype("Int64")
        empty["월"] = empty["월"].astype("Int64")
        empty["주차"] = empty["주차"].astype("Int64")
        empty["_sheet_row"] = empty["_sheet_row"].astype("Int64")
        return empty

    result = df.copy()
    parsed_dates = result["날짜"].apply(parse_record_date)
    result["날짜"] = [
        format_date_with_weekday(record_date) if record_date else value
        for record_date, value in zip(parsed_dates, result["날짜"])
    ]

    enriched_rows = []
    for idx, record_date in enumerate(parsed_dates):
        row = result.iloc[idx]
        if record_date is None:
            enriched_rows.append(
                {
                    "연월": "",
                    "월": pd.NA,
                    "월 주차": "",
                    "연도": pd.NA,
                    "주차": pd.NA,
                    "_week_start": pd.NaT,
                    "_sort_date": pd.NaT,
                }
            )
            continue

        computed = derive_date_attributes(record_date)
        enriched_rows.append(
            {
                "연월": _coalesce_sheet_value(
                    row, "연월", format_year_month(record_date)
                ),
                "월": computed["월"],
                "월 주차": _coalesce_sheet_value(row, "월 주차", computed["월 주차"]),
                "연도": computed["연도"],
                "주차": computed["주차"],
                "_week_start": computed["_week_start"],
                "_sort_date": pd.Timestamp(record_date),
            }
        )

    enriched_df = pd.DataFrame(enriched_rows)
    for col in enriched_df.columns:
        result[col] = enriched_df[col].values

    result["월"] = result["월"].astype("Int64")
    result["연도"] = result["연도"].astype("Int64")
    result["주차"] = result["주차"].astype("Int64")
    result = result.sort_values(
        ["_sort_date", "_sheet_row"],
        ascending=[False, False],
        na_position="last",
    )
    return result.reset_index(drop=True)


def sheet_rows_to_dataframe(rows: list[list[str]]) -> pd.DataFrame:
    if len(rows) <= 1:
        return enrich_dataframe(pd.DataFrame(columns=SHEET_COLUMNS + ["_sheet_row"]))

    records = []
    for sheet_row_idx, row in enumerate(rows[1:], start=2):
        padded = row + [""] * len(SHEET_COLUMNS)
        record = dict(zip(SHEET_COLUMNS, padded[: len(SHEET_COLUMNS)]))
        record["_sheet_row"] = sheet_row_idx
        records.append(record)

    return enrich_dataframe(pd.DataFrame(records))


def load_sheet_dataframe() -> pd.DataFrame:
    worksheet = get_worksheet()
    rows = worksheet.get_all_values()
    ensure_sheet_headers(worksheet, rows)
    if rows and rows[0][: len(SHEET_COLUMNS)] != SHEET_COLUMNS:
        rows = worksheet.get_all_values()
    return sheet_rows_to_dataframe(rows)


def apply_record_filters(
    df: pd.DataFrame,
    selected_year: int | None,
    selected_month: int | None,
    selected_week: int | None,
) -> pd.DataFrame:
    filtered = df.copy()

    if selected_year is not None:
        filtered = filtered[filtered["연도"] == selected_year]
    if selected_month is not None:
        filtered = filtered[filtered["월"] == selected_month]
    # 0/None = 'All(전체)' — 주차 조건을 건너뛰고 선택된 연·월 전체를 반환
    if selected_week is not None and selected_week != FILTER_WEEK_ALL:
        filtered = filtered[filtered["주차"] == selected_week]

    if not filtered.empty:
        filtered = filtered.sort_values(
            ["_sort_date", "_sheet_row"],
            ascending=[False, False],
            na_position="last",
        ).reset_index(drop=True)

    return filtered


def sum_work_hours(df: pd.DataFrame) -> float:
    if df.empty:
        return 0.0
    return float(df["당일 근무시간"].apply(parse_work_hours).sum())


def calculate_metrics(filtered_df: pd.DataFrame) -> tuple[float, float, float]:
    current_week_start = pd.Timestamp(get_tuesday_week_start(today_kst()))

    monthly_total = sum_work_hours(filtered_df)

    week_df = filtered_df[filtered_df["_week_start"] == current_week_start]
    weekly_total = sum_work_hours(week_df)
    remaining_hours = max(WEEKLY_TARGET_HOURS - weekly_total, 0.0)

    return monthly_total, weekly_total, remaining_hours


def _append_row_to_sheet(sheet_row: list) -> None:
    try:
        worksheet = get_worksheet()
        worksheet.append_row(sheet_row, value_input_option="USER_ENTERED")
    except Exception as exc:
        print(f"[Google Sheets] append_row failed: {exc}", flush=True)


def _update_row_in_sheet(row_index: int, update_values: list[list[str]]) -> None:
    try:
        worksheet = get_worksheet()
        worksheet.update(
            f"E{row_index}:G{row_index}",
            update_values,
            value_input_option="USER_ENTERED",
        )
    except Exception as exc:
        print(f"[Google Sheets] update failed: {exc}", flush=True)


def _append_modification_request_to_sheet(request_row: list) -> None:
    try:
        worksheet = get_modification_request_worksheet()
        ensure_modification_request_headers(worksheet)
        worksheet.append_row(request_row, value_input_option="USER_ENTERED")
    except Exception as exc:
        print(f"[Google Sheets] modification request append_row failed: {exc}", flush=True)


def _is_checkout_empty(value) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def is_text_area_enabled(df: pd.DataFrame) -> bool:
    """퇴근시간이 비어 있는 활성 출근 기록 존재 여부."""
    return find_open_session_row(df) is not None


def find_open_session_row(df: pd.DataFrame) -> pd.Series | None:
    if df.empty:
        return None

    open_sessions = df[
        df["출근시간"].astype(str).str.strip().ne("")
        & df["퇴근시간"].apply(_is_checkout_empty)
    ]
    if open_sessions.empty:
        return None

    return open_sessions.loc[open_sessions["_sheet_row"].idxmax()]


def apply_clock_in(df: pd.DataFrame) -> tuple[pd.DataFrame, list] | None:
    if is_text_area_enabled(df):
        st.error("아직 퇴근하지 않은 출근 기록이 있습니다. 먼저 퇴근해 주세요.")
        return None

    current = now_kst()
    today = today_kst()
    year_month, month_week_label = sheet_stored_attributes(today)
    date_label = today_str()
    check_in = time_str(current)
    sheet_row = [
        year_month,
        month_week_label,
        date_label,
        check_in,
        "",
        "",
        "",
    ]

    new_record = dict(zip(SHEET_COLUMNS, sheet_row))
    new_sheet_row = int(df["_sheet_row"].max()) + 1 if not df.empty else 2
    new_record["_sheet_row"] = new_sheet_row

    updated_df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
    return enrich_dataframe(updated_df), sheet_row


def apply_clock_out(df: pd.DataFrame, work_content: str) -> tuple[pd.DataFrame, int, list[list[str]]] | None:
    session_row = find_open_session_row(df)
    if session_row is None:
        st.warning("아직 출근 한 기록이 없습니다. 먼저 출근을 등록 해 주세요.")
        return None

    row_index = int(session_row["_sheet_row"])
    record_date = parse_record_date(session_row["날짜"])
    clock_in_time = parse_time(session_row["출근시간"], record_date)
    if clock_in_time is None:
        st.error("출근 시간 형식을 읽을 수 없습니다.")
        return None

    current = now_kst()
    if current <= clock_in_time:
        st.error("퇴근 시간이 출근 시간보다 빠릅니다.")
        return None

    elapsed = current - clock_in_time
    work_hours = elapsed.total_seconds() / 3600.0
    work_minutes = round(elapsed.total_seconds() / 60)
    work_summary = work_content.strip()
    checkout = time_str(current)
    hours_label = f"{work_hours:.2f} ({work_minutes}분)"
    update_values = [[checkout, hours_label, work_summary]]

    updated_df = df.copy()
    row_mask = updated_df["_sheet_row"] == row_index
    updated_df.loc[row_mask, "퇴근시간"] = checkout
    updated_df.loc[row_mask, "당일 근무시간"] = hours_label
    updated_df.loc[row_mask, "업무 내용"] = work_summary

    return enrich_dataframe(updated_df), row_index, update_values


def render_top_filters(df: pd.DataFrame) -> pd.DataFrame:
    today_year, today_month, today_week = get_today_filter_defaults()
    _sync_time_aware_filter_defaults(today_year, today_month, today_week)

    if df.empty:
        year_options = [today_year]
    else:
        year_options = sorted(df["연도"].dropna().unique().astype(int).tolist())
    year_options = _inject_filter_option(year_options, today_year)
    _ensure_selectbox_value(FILTER_YEAR_KEY, year_options, today_year)

    filter_col1, filter_col2, filter_col3 = st.columns(3)

    with filter_col1:
        selected_year = st.selectbox(
            "연도",
            options=year_options,
            index=year_options.index(st.session_state[FILTER_YEAR_KEY]),
            key=FILTER_YEAR_KEY,
            format_func=lambda value: f"{value}년",
        )

    filtered_by_year = df[df["연도"] == selected_year] if not df.empty else df
    if not filtered_by_year.empty:
        month_options = sorted(filtered_by_year["월"].dropna().unique().astype(int).tolist())
    else:
        month_options = []
    if selected_year == today_year:
        month_options = _inject_filter_option(month_options, today_month)
    if not month_options:
        month_options = [today_month if selected_year == today_year else 1]
    preferred_month = today_month if selected_year == today_year else month_options[-1]
    _ensure_selectbox_value(FILTER_MONTH_KEY, month_options, preferred_month)

    with filter_col2:
        selected_month = st.selectbox(
            "월",
            options=month_options,
            index=month_options.index(st.session_state[FILTER_MONTH_KEY]),
            key=FILTER_MONTH_KEY,
            format_func=lambda value: f"{value}월",
        )

    filtered_by_month = (
        filtered_by_year[filtered_by_year["월"] == selected_month]
        if not filtered_by_year.empty
        else filtered_by_year
    )
    if not filtered_by_month.empty:
        week_options = sorted(filtered_by_month["주차"].dropna().unique().astype(int).tolist())
    else:
        week_options = []
    if selected_year == today_year and selected_month == today_month:
        week_options = _inject_filter_option(week_options, today_week)
    if not week_options:
        week_options = [
            today_week
            if selected_year == today_year and selected_month == today_month
            else 1
        ]
    # 정수 0을 '전체' 센티널로 맨 앞에 추가 — int 타입 체계 유지
    week_options = [FILTER_WEEK_ALL] + [w for w in week_options if w != FILTER_WEEK_ALL]
    preferred_week = (
        today_week
        if selected_year == today_year and selected_month == today_month
        else week_options[-1]
    )
    _ensure_selectbox_value(FILTER_WEEK_KEY, week_options, preferred_week)

    with filter_col3:
        selected_week = st.selectbox(
            "주차",
            options=week_options,
            index=week_options.index(st.session_state[FILTER_WEEK_KEY]),
            key=FILTER_WEEK_KEY,
            format_func=_format_week_filter_label,
        )

    return apply_record_filters(df, selected_year, selected_month, selected_week)


def get_eligible_modification_records(filtered_df: pd.DataFrame) -> pd.DataFrame:
    """필터 결과 중 날짜·출근시간이 있는 행만 반환한다."""
    if filtered_df.empty:
        return filtered_df

    mask = (
        filtered_df["날짜"].astype(str).str.strip().ne("")
        & filtered_df["출근시간"].astype(str).str.strip().ne("")
    )
    return filtered_df[mask].copy()


def format_record_option_label(row: pd.Series) -> str:
    date_label = str(row["날짜"]).strip()
    check_in = format_time_short(row["출근시간"])
    check_out = format_time_short(row["퇴근시간"])
    return f"{date_label} | 출근: {check_in} | 퇴근: {check_out}"


def build_modification_record_options(filtered_df: pd.DataFrame) -> list[tuple[str, pd.Series]]:
    eligible = get_eligible_modification_records(filtered_df)
    if eligible.empty:
        return []

    options: list[tuple[str, pd.Series]] = []
    label_counts: dict[str, int] = {}
    for _, row in eligible.iterrows():
        label = format_record_option_label(row)
        label_counts[label] = label_counts.get(label, 0) + 1
        if label_counts[label] > 1:
            label = f"{label} · 기록 {int(row['_sheet_row'])}"
        options.append((label, row))
    return options


def render_modification_request_form(filtered_df: pd.DataFrame) -> None:
    st.subheader("📝 근무 기록 수정 요청 (관리자 승인)")

    with st.container(border=True):
        record_options = build_modification_record_options(filtered_df)
        if not record_options:
            st.caption("현재 필터 조건에 수정 요청 가능한 근무 기록이 없습니다.")
            return

        labels = [label for label, _ in record_options]

        with st.form("modification_request_form", clear_on_submit=True):
            selected_label = st.selectbox(
                "수정할 근무 기록 선택",
                options=labels,
            )
            request_detail = st.text_area(
                "수정 요청 상세 내용 (예: 퇴근 시간 23:00으로 변경 요청)",
                height=100,
            )
            submitted = st.form_submit_button("요청 제출", type="primary", use_container_width=True)

        if submitted:
            if not request_detail or not request_detail.strip():
                st.error("수정 요청 상세 내용을 입력해 주세요.")
                return

            selected_row = next(row for label, row in record_options if label == selected_label)
            request_row = [
                request_datetime_str(),
                str(selected_row["날짜"]).strip(),
                str(selected_row["출근시간"]).strip(),
                "" if _is_checkout_empty(selected_row["퇴근시간"]) else str(selected_row["퇴근시간"]).strip(),
                request_detail.strip(),
                "대기중",
            ]
            threading.Thread(
                target=_append_modification_request_to_sheet,
                args=(request_row,),
                daemon=True,
            ).start()


def main():
    st.set_page_config(page_title="근무 시간 통합 모니터링", layout="wide")
    st.title("근무 시간 통합 모니터링 시스템")

    if DF_SESSION_KEY not in st.session_state:
        try:
            st.session_state[DF_SESSION_KEY] = load_sheet_dataframe()
        except Exception as exc:
            st.error(f"구글 시트 연결 오류: {exc}")
            return

    df = st.session_state[DF_SESSION_KEY]

    # 구역 1: 조회
    filtered_df = render_top_filters(df)
    st.divider()

    # 구역 2: 요약
    monthly_total, weekly_total, _ = calculate_metrics(filtered_df)
    monthly_minutes = work_value_to_minutes(monthly_total)
    weekly_minutes = work_value_to_minutes(weekly_total)

    metric_col1, metric_col2 = st.columns(2)
    with metric_col1:
        st.metric(
            label="이번 달 총 누적 근무시간",
            value=format_duration(monthly_minutes),
        )
    with metric_col2:
        st.metric(
            label="이번 주(화~월) 누적 근무시간",
            value=format_duration(weekly_minutes),
        )

    progress_ratio = min(weekly_minutes / WEEKLY_TARGET_MINUTES, 1.0)
    st.caption(
        f"주간 8시간 달성률 · {format_duration(weekly_minutes)} "
        f"/ {format_duration(WEEKLY_TARGET_MINUTES)} "
        f"({int(progress_ratio * 100)}%)"
    )
    st.progress(progress_ratio)

    week_start, week_end = get_custom_week_range()
    st.caption(
        f"주간 집계 기간: {format_date_with_weekday(week_start)} ~ "
        f"{format_date_with_weekday(week_end)} "
        f"(KST, 화요일 시작 · 월 귀속 기준: 해당 월 첫 화요일 포함 주)"
    )
    st.divider()

    # 구역 3: 액션
    df = st.session_state[DF_SESSION_KEY]
    is_working = is_text_area_enabled(df)

    current_ip = get_client_ip()
    if current_ip == 0:
        st.info("보안 네트워크를 확인하고 있습니다...")
        clock_in_disabled = True
        clock_out_disabled = True
    elif current_ip not in ALLOWED_OFFICE_IPS:
        st.error(
            f"🔒 지정된 사무실 네트워크(Wi-Fi)에서만 출퇴근이 가능합니다. (현재 접속 IP: {current_ip})"
        )
        clock_in_disabled = True
        clock_out_disabled = True
    else:
        clock_in_disabled = is_working
        clock_out_disabled = not is_working

    action_col1, action_col2 = st.columns(2)
    with action_col1:
        if st.button(
            "출근하기",
            type="primary",
            use_container_width=True,
            disabled=clock_in_disabled,
        ):
            result = apply_clock_in(df)
            if result is not None:
                updated_df, sheet_row = result
                _append_row_to_sheet(sheet_row)
                st.session_state[DF_SESSION_KEY] = updated_df
                st.rerun()
    with action_col2:
        if st.button(
            "퇴근하기",
            type="secondary",
            use_container_width=True,
            disabled=clock_out_disabled,
        ):
            work_content = st.session_state.get(WORK_DETAIL_KEY, "")
            if not work_content or not work_content.strip():
                st.warning("오늘 수행한 업무 상세를 반드시 입력해야 퇴근이 완료됩니다.")
            else:
                result = apply_clock_out(df, work_content)
                if result is not None:
                    updated_df, row_index, update_values = result
                    _update_row_in_sheet(row_index, update_values)
                    st.session_state[DF_SESSION_KEY] = updated_df
                    st.session_state[WORK_DETAIL_KEY] = ""
                    st.rerun()

    df = st.session_state[DF_SESSION_KEY]
    is_working = is_text_area_enabled(df)

    st.write("")
    if not is_working:
        st.info(
            "출근 전입니다. 출근 등록 후, 퇴근 시점에 수행 업무를 입력할 수 있습니다."
        )
    else:
        st.write("")

    st.text_area(
        "오늘 수행한 업무 상세",
        placeholder="오늘 수행한 업무 내용을 입력하세요." if is_working else "",
        height=120,
        key=WORK_DETAIL_KEY,
        disabled=not is_working,
    )
    st.divider()

    # 구역 4: 기록
    record_header_col, record_download_col = st.columns([6, 1])
    with record_header_col:
        st.subheader("📋 상세 근무 기록")
    with record_download_col:
        st.write("")
        selected_year = st.session_state.get(FILTER_YEAR_KEY, "")
        selected_month = st.session_state.get(FILTER_MONTH_KEY, "")
        selected_week = st.session_state.get(FILTER_WEEK_KEY)
        week_label = (
            "전체"
            if selected_week in (FILTER_WEEK_ALL, None)
            else f"{selected_week}주차"
        )
        excel_filename = f"근무보고_{selected_year}_{selected_month:02d}_{week_label}.xlsx"
        st.download_button(
            label="EXCEL",
            data=generate_excel_file(filtered_df),
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/table_chart:",
            use_container_width=True,
        )
    st.dataframe(
        build_display_dataframe(filtered_df),
        use_container_width=True,
        hide_index=True,
    )

    render_modification_request_form(filtered_df)

    if st.session_state.get(FALLBACK_TIME_KEY):
        st.caption("경고: 로컬 시간 기준 기록됨")


if __name__ == "__main__":
    main()