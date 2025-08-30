# -*- coding: utf-8 -*-
"""
영업보고서 웹앱 (Streamlit · s18)
- s17 기반 + "Google Sheets 내보내기 (OAuth)" 기능 추가
- 사용자의 OAuth 코드 패턴(credential.json + token.pkl 재사용) 통합: 최초 1회 로그인 후 토큰 재사용
- 현재 입력/전체 목록을 각각 새 스프레드시트로 생성 후 시트(tab) 단위로 DataFrame 저장
- 기존 "Google Drive 업로드(서비스 계정, XLSX)" 기능은 유지

패키지(로컬 PowerShell)
  & "C:\\Users\\customer\\AppData\\Local\\Programs\\Python\\Python313\\python.exe" -m pip install streamlit pandas openpyxl gspread gspread-dataframe google-auth-oauthlib google-auth google-auth-httplib2
실행
  & "C:\\Users\\customer\\AppData\\Local\\Programs\\Python\\Python313\\python.exe" -m streamlit run "C:\\tsct\\s18.py"
"""
from __future__ import annotations

import io
import json
import re
import sys
import uuid
import os
import pickle
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

# ----------------------------
# Streamlit 런타임/모듈 체크
# ----------------------------
try:
    import streamlit as st  # type: ignore
    STREAMLIT_AVAILABLE = True
except ModuleNotFoundError:
    st = None  # type: ignore
    STREAMLIT_AVAILABLE = False

def is_streamlit_runtime() -> bool:
    if not STREAMLIT_AVAILABLE:
        return False
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx  # type: ignore
        return get_script_run_ctx() is not None
    except Exception:
        return False

# ----------------------------
# 상수/경로
# ----------------------------
APP_TITLE = "영업보고서"
DATA_DIR = Path("data")
DB_PATH = DATA_DIR / "sales_records.json"
OAUTH_TOKEN = Path("token.pkl")
OAUTH_CREDENTIALS = Path("credentials.json")

CHARGER_MODELS = [
    "2100A", "1100A", "3050A", "3050B", "3050C",
    "2007CP", "2007A", "2007C", "1007B", "1030A",
]

ANCILLARY_ITEMS = [
    "I형 볼라드", "U형 볼라드", "기초패드(완속)", "기초패드(급속)",
    "캐노피(완속)", "캐노피(급속)", "바닥면도색",
]

STATUS_CHOICES = ["진행중", "완료", "불가"]

MOBILE_PREFIXES = {"010", "011", "016", "017", "018", "019"}
AREA3 = {"031","032","033","041","042","043","044","051","052","053","054","055","061","062","063","064"}
SERVICE_3 = {"070","050","080"}

# ----------------------------
# 유틸
# ----------------------------

def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not DB_PATH.exists():
        DB_PATH.write_text(json.dumps({"records": []}, ensure_ascii=False, indent=2), encoding="utf-8")


def load_db() -> Dict:
    ensure_dirs()
    with DB_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_db(db: Dict) -> None:
    ensure_dirs()
    with DB_PATH.open("w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def fmt_date(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def strip_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def format_korean_phone(raw: str) -> str:
    digits = strip_digits(raw)
    if not digits:
        return ""
    if len(digits) == 8 and digits[:2] in {"15", "16", "18"}:
        return f"{digits[:4]}-{digits[4:8]}"
    if digits.startswith("02"):
        if len(digits) <= 2:
            return digits
        if 3 <= len(digits) <= 5:
            return f"02-{digits[2:]}"
        if 6 <= len(digits) <= 9:
            return f"02-{digits[2:5]}-{digits[5:9]}"
        return f"02-{digits[2:6]}-{digits[6:10]}"
    if digits[:3] in SERVICE_3:
        if len(digits) <= 3:
            return digits
        if 4 <= len(digits) <= 7:
            return f"{digits[:3]}-{digits[3:]}"
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:11]}"
    if digits[:3] in MOBILE_PREFIXES:
        if len(digits) <= 3:
            return digits
        if 4 <= len(digits) <= 7:
            return f"{digits[:3]}-{digits[3:]}"
        if len(digits) == 10:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:10]}"
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:11]}"
    if digits[:3] in AREA3:
        if len(digits) <= 3:
            return digits
        if 4 <= len(digits) <= 6:
            return f"{digits[:3]}-{digits[3:]}"
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:10]}"
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:8]}"
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:10]}"
    if len(digits) >= 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:11]}"
    if len(digits) <= 3:
        return digits
    return f"{digits[:3]}-{digits[3:]}"


def make_record_id() -> str:
    return str(uuid.uuid4())


def summarize_record(rec: Dict) -> str:
    return f"{rec.get('date','')} | {rec.get('site_name','')} | {rec.get('salesperson','')} | {rec.get('status','')}"


def tot_qty(mapping: Dict[str, int]) -> int:
    return int(sum(int(v or 0) for v in mapping.values()))


def blank_extras_df() -> pd.DataFrame:
    return pd.DataFrame({
        "모델명": pd.Series(dtype="string"),
        "수량": pd.Series(dtype="Int64"),
    })

# ----------------------------
# 데이터 모델
# ----------------------------

def build_record(
    *,
    record_id: Optional[str],
    d_date: date,
    salesperson: str,
    site_name: str,
    manager_name: str,
    phone: str,
    remarks: str,
    status: str,
    reason: str,
    charger_counts: Dict[str, int],
    ancillary_counts: Dict[str, int],
    extra_rows: List[Dict[str, str | int]],
) -> Dict:
    if record_id is None:
        record_id = make_record_id()
    rec = {
        "id": record_id,
        "created_at": now_str(),
        "date": fmt_date(d_date),
        "salesperson": salesperson.strip(),
        "site_name": site_name.strip(),
        "manager_name": manager_name.strip(),
        "phone": format_korean_phone(phone),
        "remarks": (remarks or "").strip(),
        "status": status,
        "reason": (reason or "").strip(),
        "chargers": {k: int(charger_counts.get(k, 0) or 0) for k in CHARGER_MODELS},
        "ancillaries": {k: int(ancillary_counts.get(k, 0) or 0) for k in ANCILLARY_ITEMS},
        "extras": [
            {"name": str(r.get("모델명", "")).strip(), "qty": int(r.get("수량", 0) or 0)}
            for r in extra_rows
            if str(r.get("모델명", "")).strip() and int(r.get("수량", 0) or 0) > 0
        ],
    }
    rec["totals"] = {
        "chargers_total": tot_qty(rec["chargers"]),
        "ancillaries_total": tot_qty(rec["ancillaries"]),
        "extras_total": sum(int(x.get("qty", 0) or 0) for x in rec["extras"]),
    }
    return rec

# ----------------------------
# 엑셀 변환 (기존 기능 유지)
# ----------------------------

def excel_bytes_for_record(rec: Dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        base_rows = [
            ("작성일", rec.get("created_at", "")),
            ("날짜", rec.get("date", "")),
            ("영업자", rec.get("salesperson", "")),
            ("현장명", rec.get("site_name", "")),
            ("담당자", rec.get("manager_name", "")),
            ("연락처", rec.get("phone", "")),
            ("진행상태", rec.get("status", "")),
            ("불가 사유", rec.get("reason", "")),
            ("비고", rec.get("remarks", "")),
        ]
        pd.DataFrame(base_rows, columns=["항목", "값"]).to_excel(writer, sheet_name="기본정보", index=False)
        rows_qty = []
        for k, v in rec.get("chargers", {}).items():
            rows_qty.append(("충전기", k, int(v or 0)))
        for k, v in rec.get("ancillaries", {}).items():
            rows_qty.append(("부대공사", k, int(v or 0)))
        pd.DataFrame(rows_qty, columns=["분류", "항목", "수량"]).to_excel(writer, sheet_name="수량", index=False)
        extras = rec.get("extras", [])
        if extras:
            pd.DataFrame(extras, columns=["name", "qty"]).rename(columns={"name": "모델명", "qty": "수량"}).to_excel(
                writer, sheet_name="기타모델", index=False
            )
    return buf.getvalue()


def excel_bytes_for_all(records: List[Dict]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        rows = []
        for r in records:
            rows.append([
                r.get("id", ""), r.get("date", ""), r.get("salesperson", ""), r.get("site_name", ""),
                r.get("manager_name", ""), r.get("phone", ""), r.get("status", ""), r.get("reason", ""), r.get("remarks", ""),
                r.get("totals", {}).get("chargers_total", 0), r.get("totals", {}).get("ancillaries_total", 0), r.get("totals", {}).get("extras_total", 0),
            ])
        pd.DataFrame(rows, columns=[
            "ID", "날짜", "영업자", "현장명", "담당자", "연락처", "진행상태", "불가 사유", "비고",
            "충전기 합계", "부대공사 합계", "기타 합계",
        ]).to_excel(writer, sheet_name="목록", index=False)
        ch_sum = {k: 0 for k in CHARGER_MODELS}
        an_sum = {k: 0 for k in ANCILLARY_ITEMS}
        for r in records:
            for k, v in r.get("chargers", {}).items():
                ch_sum[k] += int(v or 0)
            for k, v in r.get("ancillaries", {}).items():
                an_sum[k] += int(v or 0)
        pd.DataFrame(list(ch_sum.items()), columns=["항목", "수량"]).sort_values("항목").to_excel(writer, sheet_name="충전기합계", index=False)
        pd.DataFrame(list(an_sum.items()), columns=["항목", "수량"]).sort_values("항목").to_excel(writer, sheet_name="부대공사합계", index=False)
    return buf.getvalue()

# ----------------------------
# Google Drive 업로드 (서비스계정) 기존 유지
# ----------------------------

def upload_to_drive_via_service_account(*, file_bytes: bytes, filename: str, folder_id: str, service_account_info: dict) -> Dict:
    try:
        from google.oauth2 import service_account  # type: ignore
        from googleapiclient.discovery import build  # type: ignore
        from googleapiclient.http import MediaIoBaseUpload  # type: ignore
    except Exception as e:
        raise ImportError(
            "구글 드라이브 업로드 모듈이 설치되어 있지 않습니다.\n"
            "pip install google-api-python-client google-auth google-auth-httplib2"
        ) from e
    creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=["https://www.googleapis.com/auth/drive.file"])
    service = build("drive", "v3", credentials=creds)
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=False)
    created = service.files().create(
        body={"name": filename, "parents": [folder_id], "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        media_body=media,
        fields="id, webViewLink, webContentLink",
    ).execute()
    return created

# ----------------------------
# Google Sheets 내보내기 (OAuth, 사용자 계정)
# ----------------------------

def oauth_get_gspread_client() -> "gspread.Client":
    import gspread
    from gspread_dataframe import set_with_dataframe
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = None
    if OAUTH_TOKEN.exists():
        with OAUTH_TOKEN.open("rb") as f:
            creds = pickle.load(f)
    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            creds.refresh(Request())
        else:
            if not OAUTH_CREDENTIALS.exists():
                raise FileNotFoundError("credentials.json 파일이 필요합니다. 상단에서 업로드해 주세요.")
            flow = InstalledAppFlow.from_client_secrets_file(str(OAUTH_CREDENTIALS), SCOPES)
            creds = flow.run_local_server(port=0)
        with OAUTH_TOKEN.open("wb") as f:
            pickle.dump(creds, f)
    return gspread.authorize(creds)


def gsheet_export_record(rec: Dict) -> str:
    import gspread
    from gspread_dataframe import set_with_dataframe
    client = oauth_get_gspread_client()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ss_title = f"영업보고서_현재입력_{ts}"
    ss = client.create(ss_title)
    sh1 = ss.sheet1
    sh1.update_title("기본정보")

    # 기본정보
    base_rows = [
        ("작성일", rec.get("created_at", "")),
        ("날짜", rec.get("date", "")),
        ("영업자", rec.get("salesperson", "")),
        ("현장명", rec.get("site_name", "")),
        ("담당자", rec.get("manager_name", "")),
        ("연락처", rec.get("phone", "")),
        ("진행상태", rec.get("status", "")),
        ("불가 사유", rec.get("reason", "")),
        ("비고", rec.get("remarks", "")),
    ]
    set_with_dataframe(sh1, pd.DataFrame(base_rows, columns=["항목", "값"]))

    # 수량
    sh2 = ss.add_worksheet(title="수량", rows=100, cols=20)
    rows_qty = []
    for k, v in rec.get("chargers", {}).items():
        rows_qty.append(("충전기", k, int(v or 0)))
    for k, v in rec.get("ancillaries", {}).items():
        rows_qty.append(("부대공사", k, int(v or 0)))
    set_with_dataframe(sh2, pd.DataFrame(rows_qty, columns=["분류", "항목", "수량"]))

    # 기타모델
    extras = rec.get("extras", [])
    if extras:
        sh3 = ss.add_worksheet(title="기타모델", rows=200, cols=20)
        df_ex = pd.DataFrame(extras, columns=["name", "qty"]).rename(columns={"name": "모델명", "qty": "수량"})
        set_with_dataframe(sh3, df_ex)

    return ss.url


def gsheet_export_all(records: List[Dict]) -> str:
    import gspread
    from gspread_dataframe import set_with_dataframe
    client = oauth_get_gspread_client()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ss_title = f"영업보고서_목록_{ts}"
    ss = client.create(ss_title)
    sh1 = ss.sheet1
    sh1.update_title("목록")

    rows = []
    for r in records:
        rows.append([
            r.get("id", ""), r.get("date", ""), r.get("salesperson", ""), r.get("site_name", ""),
            r.get("manager_name", ""), r.get("phone", ""), r.get("status", ""), r.get("reason", ""), r.get("remarks", ""),
            r.get("totals", {}).get("chargers_total", 0), r.get("totals", {}).get("ancillaries_total", 0), r.get("totals", {}).get("extras_total", 0),
        ])
    set_with_dataframe(sh1, pd.DataFrame(rows, columns=[
        "ID", "날짜", "영업자", "현장명", "담당자", "연락처", "진행상태", "불가 사유", "비고",
        "충전기 합계", "부대공사 합계", "기타 합계",
    ]))

    # 합계 탭들
    ch_sum = {k: 0 for k in CHARGER_MODELS}
    an_sum = {k: 0 for k in ANCILLARY_ITEMS}
    for r in records:
        for k, v in r.get("chargers", {}).items():
            ch_sum[k] += int(v or 0)
        for k, v in r.get("ancillaries", {}).items():
            an_sum[k] += int(v or 0)
    sh2 = ss.add_worksheet(title="충전기합계", rows=200, cols=20)
    sh3 = ss.add_worksheet(title="부대공사합계", rows=200, cols=20)
    set_with_dataframe(sh2, pd.DataFrame(list(ch_sum.items()), columns=["항목", "수량"]).sort_values("항목"))
    set_with_dataframe(sh3, pd.DataFrame(list(an_sum.items()), columns=["항목", "수량"]).sort_values("항목"))

    return ss.url

# ----------------------------
# UI: 연락처(실시간 하이픈) + 기본 폼 + 내보내기
# ----------------------------

def phone_on_change():
    v = st.session_state.get("ui_form_phone", "")
    st.session_state.ui_form_phone = format_korean_phone(v)


def init_session():
    if not STREAMLIT_AVAILABLE:
        return
    if "db" not in st.session_state:
        st.session_state.db = load_db()
    if "editing_id" not in st.session_state:
        st.session_state.editing_id = None
    if "pending_load_id" not in st.session_state:
        st.session_state.pending_load_id = None
    if "extras_data" not in st.session_state:
        st.session_state.extras_data = blank_extras_df()
    if "ui_form_phone" not in st.session_state:
        st.session_state.ui_form_phone = ""


def set_form_from_record(rec: Dict) -> None:
    st.session_state.form_date = datetime.strptime(rec.get("date", fmt_date(date.today())), "%Y-%m-%d").date()
    st.session_state.form_salesperson = rec.get("salesperson", "")
    st.session_state.form_site = rec.get("site_name", "")
    st.session_state.form_manager = rec.get("manager_name", "")
    st.session_state.ui_form_phone = rec.get("phone", "")
    st.session_state.form_remarks = rec.get("remarks", "")
    st.session_state.form_status = rec.get("status", STATUS_CHOICES[0])
    st.session_state.form_reason = rec.get("reason", "")
    for k in CHARGER_MODELS:
        st.session_state[f"qty_ch_{k}"] = int(rec.get("chargers", {}).get(k, 0) or 0)
    for k in ANCILLARY_ITEMS:
        st.session_state[f"qty_an_{k}"] = int(rec.get("ancillaries", {}).get(k, 0) or 0)
    ex = rec.get("extras", [])
    df = pd.DataFrame([{"모델명": e.get("name", ""), "수량": int(e.get("qty", 0) or 0)} for e in ex])
    if df.empty:
        df = blank_extras_df()
    else:
        df["모델명"] = df["모델명"].astype("string")
        df["수량"] = pd.to_numeric(df["수량"], errors="coerce").astype("Int64")
    st.session_state.extras_data = df


def sidebar_records_ui():
    st.sidebar.header("저장된 영업보고서")
    db = st.session_state.db
    records = db.get("records", [])
    q = st.sidebar.text_input("검색(날짜/현장/영업자/상태)", key="search_q")
    def match(r: Dict) -> bool:
        s = (r.get("date", "") + " " + r.get("site_name", "") + " " + r.get("salesperson", "") + " " + r.get("status", "")).lower()
        return (q or "").lower() in s
    filtered = [r for r in records if match(r)]
    options = {summarize_record(r): r["id"] for r in filtered}
    chosen_id = None
    if options:
        chosen_label = st.sidebar.selectbox("레코드 선택", list(options.keys()), index=0)
        chosen_id = options[chosen_label]
    else:
        st.sidebar.info("저장된 레코드가 없습니다.")
    col1, col2, col3 = st.sidebar.columns(3)
    with col1:
        if st.button("로드", use_container_width=True, disabled=chosen_id is None):
            st.session_state.pending_load_id = chosen_id
            st.session_state.editing_id = chosen_id
            st.rerun()
    with col2:
        if st.button("새 입력", use_container_width=True):
            st.session_state.editing_id = None
            st.session_state.form_date = date.today()
            st.session_state.form_salesperson = ""
            st.session_state.form_site = ""
            st.session_state.form_manager = ""
            st.session_state.ui_form_phone = ""
            st.session_state.form_remarks = ""
            st.session_state.form_status = STATUS_CHOICES[0]
            st.session_state.form_reason = ""
            st.session_state.extras_data = blank_extras_df()
            for k in CHARGER_MODELS:
                st.session_state[f"qty_ch_{k}"] = 0
            for k in ANCILLARY_ITEMS:
                st.session_state[f"qty_an_{k}"] = 0
            st.rerun()
    with col3:
        if st.button("삭제", use_container_width=True, disabled=chosen_id is None):
            if chosen_id is not None:
                db["records"] = [r for r in records if r["id"] != chosen_id]
                save_db(db)
                st.session_state.db = db
                st.toast("삭제 완료", icon="✅")
                st.rerun()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("영업 현장 기록을 표준화하고, 엑셀/구글드라이브/스프레드시트로 바로 공유하세요.")

    init_session()
    sidebar_records_ui()

    if st.session_state.pending_load_id:
        rid = st.session_state.pending_load_id
        st.session_state.pending_load_id = None
        target = next((r for r in st.session_state.db.get("records", []) if r.get("id") == rid), None)
        if target:
            set_form_from_record(target)

    # 1) 연락처 입력(폼 밖, 실시간 하이픈)
    st.subheader("연락처")
    ui_form_phone = st.text_input(
        "연락처 (실시간 자동 하이픈)",
        key="ui_form_phone",
        value=st.session_state.get("ui_form_phone", ""),
        on_change=lambda: st.session_state.update({"ui_form_phone": format_korean_phone(st.session_state.get("ui_form_phone", ""))}),
        placeholder="예) 01012345678, 0212345678, 15881234"
    )
    st.caption(f"현재 입력: {st.session_state.get('ui_form_phone','')}")

    # 2) 나머지 입력(폼)
    with st.form(key="report_form", clear_on_submit=False):
        st.subheader("기본 정보")
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
        form_date = c1.date_input("날짜", key="form_date", value=st.session_state.get("form_date", date.today()))
        form_salesperson = c2.text_input("영업자", key="form_salesperson", value=st.session_state.get("form_salesperson", ""))
        form_site = c3.text_input("현장명", key="form_site", value=st.session_state.get("form_site", ""))
        form_manager = c4.text_input("담당자", key="form_manager", value=st.session_state.get("form_manager", ""))

        form_remarks = st.text_area("비고 (최대 400자)", key="form_remarks", value=st.session_state.get("form_remarks", ""), max_chars=400, height=100)

        form_status = st.radio("진행상태", STATUS_CHOICES, horizontal=True, key="form_status", index=STATUS_CHOICES.index(st.session_state.get("form_status", STATUS_CHOICES[0])))
        form_reason = st.text_area("불가 사유", key="form_reason", value=st.session_state.get("form_reason", ""), height=80)

        # 수량 입력
        st.subheader("충전기/부대공사 수량 입력")
        with st.expander("충전기 수량", expanded=True):
            cols = st.columns(5)
            charger_counts: Dict[str, int] = {}
            for i, name in enumerate(CHARGER_MODELS):
                charger_counts[name] = cols[i % 5].number_input(name, min_value=0, step=1, key=f"qty_ch_{name}", value=int(st.session_state.get(f"qty_ch_{name}", 0)))
            st.caption(f"충전기 합계: **{tot_qty(charger_counts)}** 대")
        with st.expander("부대공사 수량", expanded=True):
            cols = st.columns(5)
            ancillary_counts: Dict[str, int] = {}
            for i, name in enumerate(ANCILLARY_ITEMS):
                ancillary_counts[name] = cols[i % 5].number_input(name, min_value=0, step=1, key=f"qty_an_{name}", value=int(st.session_state.get(f"qty_an_{name}", 0)))
            st.caption(f"부대공사 합계: **{tot_qty(ancillary_counts)}** 건")

        base_df = st.session_state.get("extras_data", blank_extras_df()).copy()
        if not isinstance(base_df, pd.DataFrame) or base_df.empty:
            base_df = blank_extras_df()
        else:
            if "모델명" in base_df.columns:
                base_df["모델명"] = base_df["모델명"].astype("string")
            if "수량" in base_df.columns:
                base_df["수량"] = pd.to_numeric(base_df["수량"], errors="coerce").astype("Int64")
        extras_df = st.data_editor(
            base_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "모델명": st.column_config.TextColumn("모델명", required=False, width="medium"),
                "수량": st.column_config.NumberColumn("수량", required=False, min_value=0, step=1, width="small"),
            },
            key="extras_editor",
        )

        is_editing = st.session_state.editing_id is not None
        submitted = st.form_submit_button("완료" if is_editing else "저장", type="primary")

    # 제출 처리
    if submitted:
        errs = []
        if not form_salesperson.strip():
            errs.append("영업자를 입력하세요.")
        if not form_site.strip():
            errs.append("현장명을 입력하세요.")
        if not form_manager.strip():
            errs.append("담당자를 입력하세요.")
        if not strip_digits(ui_form_phone):
            errs.append("연락처를 입력하세요.")
        if form_status == "불가" and not (form_reason or "").strip():
            errs.append("불가 사유를 입력하세요.")

        if errs:
            st.error("\n".join(["입력 오류:"] + [f"- {e}" for e in errs]))
        else:
            extra_rows: List[Dict[str, str | int]] = []
            df_ex = extras_df.copy() if isinstance(extras_df, pd.DataFrame) else blank_extras_df()
            if "모델명" in df_ex:
                df_ex["모델명"] = df_ex["모델명"].astype("string").fillna("")
            else:
                df_ex["모델명"] = ""
            if "수량" in df_ex:
                df_ex["수량"] = pd.to_numeric(df_ex["수량"], errors="coerce").fillna(0).astype(int)
            else:
                df_ex["수량"] = 0
            for _, row in df_ex.iterrows():
                name = str(row.get("모델명", "")).strip()
                qty = int(row.get("수량", 0) or 0)
                if name and qty > 0:
                    extra_rows.append({"모델명": name, "수량": qty})
            st.session_state.extras_data = df_ex.copy()

            formatted_phone = format_korean_phone(ui_form_phone)
            rec = build_record(
                record_id=st.session_state.editing_id,
                d_date=form_date,
                salesperson=form_salesperson,
                site_name=form_site,
                manager_name=form_manager,
                phone=formatted_phone,
                remarks=form_remarks,
                status=form_status,
                reason=form_reason,
                charger_counts=charger_counts,
                ancillary_counts=ancillary_counts,
                extra_rows=extra_rows,
            )

            db = st.session_state.db
            all_recs = db.get("records", [])
            if st.session_state.editing_id is None:
                all_recs.append(rec)
                st.toast("저장 완료", icon="✅")
                st.session_state.editing_id = rec["id"]
            else:
                for i, r in enumerate(all_recs):
                    if r.get("id") == st.session_state.editing_id:
                        all_recs[i] = rec
                        break
                st.toast("완료", icon="✅")
            db["records"] = all_recs
            save_db(db)
            st.session_state.db = db
            st.rerun()

    # 내보내기/업로드
    st.markdown("---")
    st.subheader("내보내기 & 업로드")

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        st.write("**엑셀 다운로드**")
        dl_choice = st.radio("대상", ["현재 입력", "전체 목록"], horizontal=True, key="dl_choice")
        if dl_choice == "현재 입력":
            cur = None
            if st.session_state.editing_id is not None:
                cur = next((r for r in st.session_state.db.get("records", []) if r.get("id") == st.session_state.editing_id), None)
            if cur is not None:
                xls_bytes = excel_bytes_for_record(cur)
                st.download_button("현재 입력건 다운로드", data=xls_bytes, file_name=f"영업보고서_{cur.get('date','')}_{cur.get('site_name','')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.info("먼저 저장(또는 완료)하여 레코드를 만든 뒤 다운로드할 수 있습니다.")
        else:
            recs = st.session_state.db.get("records", [])
            if recs:
                xls_all = excel_bytes_for_all(recs)
                today_str = datetime.now().strftime("%Y%m%d")
                st.download_button("전체 목록 다운로드", data=xls_all, file_name=f"영업보고서_목록_{today_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.info("저장된 레코드가 없습니다.")

    with c2:
        st.write("**Google Drive 업로드** (서비스 계정, XLSX)")
        up_choice = st.radio("대상", ["현재 입력", "전체 목록"], horizontal=True, key="up_choice")
        service_json_file = st.file_uploader("서비스 계정 JSON 업로드", type=["json"], key="svc_json")
        folder_id = st.text_input("드라이브 폴더 ID", key="gdrive_folder")
        can_upload = service_json_file is not None and folder_id.strip() != ""
        if st.button("구글 드라이브로 업로드", use_container_width=True, disabled=not can_upload):
            try:
                svc_info = json.load(service_json_file)
                if up_choice == "현재 입력":
                    cur = None
                    if st.session_state.editing_id is not None:
                        cur = next((r for r in st.session_state.db.get("records", []) if r.get("id") == st.session_state.editing_id), None)
                    if cur is None:
                        st.error("현재 입력건이 없습니다. 먼저 저장하세요.")
                    else:
                        file_bytes = excel_bytes_for_record(cur)
                        fname = f"영업보고서_{cur.get('date','')}_{cur.get('site_name','')}.xlsx"
                        created = upload_to_drive_via_service_account(file_bytes=file_bytes, filename=fname, folder_id=folder_id, service_account_info=svc_info)
                        st.success(f"업로드 완료! 파일 ID: {created.get('id')}")
                        link = created.get("webViewLink") or created.get("webContentLink")
                        if link:
                            st.markdown(f"[드라이브에서 열기]({link})")
                else:
                    recs = st.session_state.db.get("records", [])
                    if not recs:
                        st.error("업로드할 전체 목록이 없습니다.")
                    else:
                        file_bytes = excel_bytes_for_all(recs)
                        today_str = datetime.now().strftime("%Y%m%d")
                        fname = f"영업보고서_목록_{today_str}.xlsx"
                        created = upload_to_drive_via_service_account(file_bytes=file_bytes, filename=fname, folder_id=folder_id, service_account_info=svc_info)
                        st.success(f"업로드 완료! 파일 ID: {created.get('id')}")
                        link = created.get("webViewLink") or created.get("webContentLink")
                        if link:
                            st.markdown(f"[드라이브에서 열기]({link})")
            except ImportError as e:
                st.error(str(e))
            except Exception as e:
                st.exception(e)

    with c3:
        st.write("**Google Sheets 내보내기 (OAuth)**")
        st.caption("최초 1회 credentials.json을 업로드해 로그인하면 token.pkl이 생성되어 이후 재사용됩니다.")
        cred_file = st.file_uploader("credentials.json 업로드(최초 1회)", type=["json"], key="cred_json")
        colA, colB, colC = st.columns(3)
        with colA:
            if st.button("토큰 초기화", use_container_width=True):
                try:
                    if OAUTH_TOKEN.exists():
                        OAUTH_TOKEN.unlink()
                    st.success("토큰 삭제 완료")
                except Exception as e:
                    st.exception(e)
        with colB:
            if st.button("시트로 내보내기 - 현재 입력", use_container_width=True):
                try:
                    if cred_file is not None:
                        OAUTH_CREDENTIALS.write_bytes(cred_file.getvalue())
                    cur = None
                    if st.session_state.editing_id is not None:
                        cur = next((r for r in st.session_state.db.get("records", []) if r.get("id") == st.session_state.editing_id), None)
                    if cur is None:
                        st.error("현재 입력건이 없습니다. 먼저 저장하세요.")
                    else:
                        url = gsheet_export_record(cur)
                        st.success("Google Sheets로 내보내기 완료")
                        st.markdown(f"[스프레드시트 열기]({url})")
                except Exception as e:
                    st.exception(e)
        with colC:
            if st.button("시트로 내보내기 - 전체 목록", use_container_width=True):
                try:
                    if cred_file is not None:
                        OAUTH_CREDENTIALS.write_bytes(cred_file.getvalue())
                    recs = st.session_state.db.get("records", [])
                    if not recs:
                        st.error("내보낼 데이터가 없습니다.")
                    else:
                        url = gsheet_export_all(recs)
                        st.success("Google Sheets로 내보내기 완료")
                        st.markdown(f"[스프레드시트 열기]({url})")
                except Exception as e:
                    st.exception(e)

    # 요약
    st.markdown("---")
    st.subheader("요약 미리보기")
    recs = st.session_state.db.get("records", [])
    if recs:
        preview = [{
            "날짜": r.get("date", ""), "현장명": r.get("site_name", ""), "영업자": r.get("salesperson", ""),
            "진행": r.get("status", ""), "연락처": r.get("phone", ""),
            "충전기합계": r.get("totals", {}).get("chargers_total", 0), "부대공사합계": r.get("totals", {}).get("ancillaries_total", 0),
        } for r in recs[-15:][::-1]]
        st.dataframe(pd.DataFrame(preview), use_container_width=True, height=320)
    else:
        st.info("아직 저장된 레코드가 없습니다.")

# ----------------------------
# 단위 테스트 (--selftest)
# ----------------------------

def run_tests() -> None:
    print("[TEST] format_korean_phone...")
    assert format_korean_phone("01012345678") == "010-1234-5678"
    assert format_korean_phone("0111234567") == "011-123-4567"
    assert format_korean_phone("07012345678") == "070-1234-5678"
    assert format_korean_phone("0212345678") == "02-1234-5678"
    assert format_korean_phone("021234567") == "02-123-4567"
    assert format_korean_phone("0312345678") == "031-234-5678"
    assert format_korean_phone("15881234") == "1588-1234"
    assert format_korean_phone("010 1234 5678") == "010-1234-5678"
    assert format_korean_phone("") == ""

    print("[TEST] tot_qty...")
    assert tot_qty({"a": 1, "b": 2, "c": 3}) == 6
    assert tot_qty({"x": 0, "y": None, "z": 5}) == 5

    print("[TEST] build_record...")
    rec = build_record(
        record_id=None, d_date=date(2025, 8, 26), salesperson="홍길동", site_name="현장", manager_name="담당",
        phone="01012345678", remarks="비고", status="진행중", reason="",
        charger_counts={k: (1 if k in ("2100A", "3050A") else 0) for k in CHARGER_MODELS},
        ancillary_counts={k: (2 if k in ("I형 볼라드",) else 0) for k in ANCILLARY_ITEMS},
        extra_rows=[{"모델명": "X", "수량": 3}],
    )
    assert rec["phone"] == "010-1234-5678"
    assert rec["totals"]["chargers_total"] == 2
    assert rec["totals"]["ancillaries_total"] == 2
    assert rec["totals"]["extras_total"] == 3

    print("ALL TESTS PASSED")

# ----------------------------
# 엔트리 포인트
# ----------------------------
if __name__ == "__main__":
    if "--selftest" in sys.argv:
        run_tests()
    elif not STREAMLIT_AVAILABLE:
        print(
            "[안내] Streamlit 모듈이 설치되어 있지 않습니다.\n"
            "  & \"C:\\Users\\customer\\AppData\\Local\\Programs\\Python\\Python313\\python.exe\" -m pip install streamlit pandas openpyxl gspread gspread-dataframe google-auth-oauthlib\n"
            "  & 동일 파이썬으로 -m streamlit run \"C:\\tsct\\s18.py\"\n"
        )
    elif not is_streamlit_runtime():
        print(
            "[안내] 이 스크립트는 Streamlit 런타임에서 실행해야 합니다.\n"
            "  & \"C:\\Users\\customer\\AppData\\Local\\Programs\\Python\\Python313\\python.exe\" -m streamlit run \"C:\\tsct\\s18.py\"\n"
        )
    else:
        main()
