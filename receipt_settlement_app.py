"""
공사 정산서 영수증 자동 입력 웹앱 (Streamlit Community Cloud / 로컬 공용)

- 연결된 API 키에서 실제로 사용 가능한 Gemini 모델 목록을 자동 조회합니다.
- 503 혼잡 오류가 발생하면 사용 가능한 다른 모델로 자동 재시도·전환합니다.
- 기존 정산서(.xlsx)의 서식, 병합 셀, S열 소계 및 상단 합계 수식을 그대로 유지하며,
  중복 의심 내역은 '반영' 체크를 자동 해제해 확인 후 건너뛰기/강제입력을 선택할 수 있습니다.
"""

from __future__ import annotations

import datetime as dt
import io
import os
import time
from typing import Literal

from openpyxl import load_workbook
from pydantic import BaseModel, Field

CategoryType = Literal[
    "목자재",
    "전기자재",
    "기타공사",
    "부자재",
    "인건비",
    "간식",
    "식사",
    "차량",
]

CATEGORY_OPTIONS: list[str] = [
    "목자재",
    "전기자재",
    "기타공사",
    "부자재",
    "인건비",
    "간식",
    "식사",
    "차량",
]

# (품목명 열 번호 또는 None, 금액 열 번호) — 1-based index
CATEGORY_COL_MAP: dict[str, tuple[int | None, int]] = {
    "목자재": (4, 5),      # D, E
    "전기자재": (6, 7),    # F, G
    "기타공사": (8, 9),    # H, I
    "부자재": (10, 11),    # J, K
    "인건비": (12, 13),    # L, M
    "간식": (None, 14),    # N
    "식사": (None, 15),    # O
    "차량": (None, 16),    # P (상세 메모는 Q열 17번에 기록)
}

DEFAULT_USERS = ["박경호", "윤선호", "유원오", "김용태", "조중창"]
DATA_START_ROW = 5
DATA_END_ROW = 104

EXCLUDED_MODEL_KEYWORDS = (
    "embedding",
    "imagen",
    "veo",
    "tts",
    "audio",
    "live",
    "robotics",
    "aqa",
    "bisect",
    "banana",
    "lyria",
    "genie",
    "decision",
)


class ReceiptExtract(BaseModel):
    """Gemini Structured Outputs로 영수증에서 추출할 고정 스키마."""

    spend_date: str = Field(
        description="결제일자 (YYYY-MM-DD 형식, 인식 불가 시 빈 문자열)"
    )
    category: CategoryType = Field(
        description=(
            "지출 항목 분류: "
            "목재·합판·각재는 '목자재', 전선·차단기·조명·스위치는 '전기자재', "
            "폐기물 반출·철거·용역 공사는 '기타공사', "
            "공구·컷소날·마스크·노끈·테이프·철물 소모품은 '부자재', "
            "일당·노무비는 '인건비', 커피·음료·빵·편의점 간식은 '간식', "
            "식당·배달·도시락 식대는 '식사', 주유·주차·통행료는 '차량'"
        )
    )
    item_name: str = Field(
        description=(
            "품목명 또는 간단 요약 (예: '공구함', '마스크 외 2건', '폐기물 반출', "
            "'주유비', '주차비'). 간식/식사도 검수 참고용으로 상호나 대표 메뉴를 기입"
        )
    )
    amount: int = Field(description="최종 결제 합계 금액 (정수, 원 단위)")
    merchant: str = Field(description="가맹점 상호명")
    memo: str = Field(
        default="",
        description="차량번호(예: 2023 주유비), 주차비 내역 등 특이사항 메모",
    )


def parse_date_value(date_str: str) -> dt.date | str | None:
    text = (date_str or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return text


def normalize_date_str(val) -> str:
    if isinstance(val, dt.datetime):
        return val.date().isoformat()
    if isinstance(val, dt.date):
        return val.isoformat()
    parsed = parse_date_value(str(val or ""))
    if isinstance(parsed, dt.date):
        return parsed.isoformat()
    return str(val or "").strip()


def extract_existing_entries(
    excel_bytes: bytes,
    target_sheets: list[str],
) -> list[dict]:
    """기존 시트 내역을 읽어 중복 대조용 인덱스를 생성합니다."""
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True, read_only=True)
    entries: list[dict] = []

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        current_date = ""
        for r in range(DATA_START_ROW, DATA_END_ROW + 1):
            c_val = ws.cell(r, 3).value
            if c_val is not None and str(c_val).strip():
                current_date = normalize_date_str(c_val)

            user_val = str(ws.cell(r, 18).value or "").strip()
            for cat, (_, amt_col) in CATEGORY_COL_MAP.items():
                raw_amt = ws.cell(r, amt_col).value
                if isinstance(raw_amt, (int, float)) and raw_amt > 0:
                    entries.append(
                        {
                            "sheet": sheet_name,
                            "row": r,
                            "date": current_date,
                            "category": cat,
                            "amount": int(round(raw_amt)),
                            "user": user_val,
                        }
                    )

    wb.close()
    return entries


def check_duplicates(
    records: list[dict],
    existing_entries: list[dict],
) -> list[dict]:
    """기존 시트 및 이번 업로드 묶음과 대조해 중복 의심 건의 '반영'을 자동 해제합니다."""
    seen_in_batch: set[tuple[str, str, int]] = set()
    checked: list[dict] = []

    for rec in records:
        row = dict(rec)
        d_str = normalize_date_str(row.get("지출일", ""))
        cat = str(row.get("항목", "")).strip()
        amt = int(row.get("금액") or 0)

        exact_hits = [
            e
            for e in existing_entries
            if e["date"] == d_str and e["category"] == cat and e["amount"] == amt
        ]
        same_amt_hits = [
            e
            for e in existing_entries
            if e["date"] == d_str and e["amount"] == amt and e not in exact_hits
        ]

        batch_key = (d_str, cat, amt)
        if exact_hits:
            hit_desc = ", ".join(
                f"{e['sheet']} {e['row']}행({e['user'] or '사용자미기재'})"
                for e in exact_hits[:2]
            )
            row["중복확인"] = f"⚠️ 기존 중복 ({hit_desc})"
            row["반영"] = False
        elif same_amt_hits:
            hit_desc = ", ".join(
                f"{e['sheet']} {e['row']}행({e['category']})"
                for e in same_amt_hits[:2]
            )
            row["중복확인"] = f"⚠️ 동일 날짜·금액 ({hit_desc})"
            row["반영"] = False
        elif amt > 0 and batch_key in seen_in_batch:
            row["중복확인"] = "⚠️ 이번 업로드 내 중복"
            row["반영"] = False
        else:
            row["중복확인"] = "신규"
            row["반영"] = True

        if amt > 0:
            seen_in_batch.add(batch_key)
        checked.append(row)

    return checked


def find_next_append_row(ws) -> int:
    last_used = DATA_START_ROW - 1
    for r in range(DATA_END_ROW, DATA_START_ROW - 1, -1):
        if any(ws.cell(r, c).value is not None for c in range(3, 19)):
            last_used = r
            break
    next_row = last_used + 1
    if next_row > DATA_END_ROW:
        raise ValueError(
            f"시트 '{ws.title}'의 입력 가능 행({DATA_END_ROW}행)이 모두 찼습니다."
        )
    return next_row


def append_receipts_to_excel(
    excel_bytes: bytes,
    target_sheets: list[str],
    records: list[dict],
    omit_duplicate_dates: bool = True,
) -> tuple[bytes, dict[str, list[int]], int]:
    """'반영'이 체크된 행만 기존 서식·수식을 유지한 채 시트에 추가합니다."""
    active_records = [r for r in records if r.get("반영", True)]
    skipped_count = len(records) - len(active_records)

    wb = load_workbook(io.BytesIO(excel_bytes))
    written_info: dict[str, list[int]] = {}

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")

        ws = wb[sheet_name]
        next_row = find_next_append_row(ws)

        date_num_fmt = ws.cell(DATA_START_ROW, 3).number_format or "yyyy-mm-dd"
        amount_num_fmt = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'

        prev_date: dt.date | str | None = None
        for r in range(next_row - 1, DATA_START_ROW - 1, -1):
            val = ws.cell(r, 3).value
            if val is not None:
                prev_date = val.date() if isinstance(val, dt.datetime) else val
                break

        written_rows: list[int] = []
        for rec in active_records:
            if next_row > DATA_END_ROW:
                raise ValueError(
                    f"시트 '{sheet_name}'의 최대 행({DATA_END_ROW}행)을 초과했습니다."
                )

            category = rec.get("항목", "부자재")
            if category not in CATEGORY_COL_MAP:
                raise ValueError(f"지원하지 않는 항목입니다: {category}")

            amount = int(rec.get("금액") or 0)
            if amount <= 0:
                continue

            date_val = parse_date_value(str(rec.get("지출일", "")))
            if date_val is not None:
                if not (omit_duplicate_dates and date_val == prev_date):
                    c_cell = ws.cell(next_row, 3, value=date_val)
                    if isinstance(date_val, dt.date):
                        c_cell.number_format = date_num_fmt
                prev_date = date_val

            item_col, amt_col = CATEGORY_COL_MAP[category]
            item_name = str(rec.get("품목/내용", "")).strip()

            if item_col is not None and item_name:
                ws.cell(next_row, item_col, value=item_name)

            amt_cell = ws.cell(next_row, amt_col, value=amount)
            amt_cell.number_format = amount_num_fmt

            q_note = str(rec.get("차량/기타상세(Q)", "")).strip()
            if q_note:
                ws.cell(next_row, 17, value=q_note)
            elif category == "차량" and item_name:
                ws.cell(next_row, 17, value=item_name)

            user_name = str(rec.get("사용자", "")).strip()
            if user_name:
                ws.cell(next_row, 18, value=user_name)

            extra_note = str(rec.get("비고(T)", "")).strip()
            if extra_note:
                ws.cell(next_row, 20, value=extra_note)

            settle_status = str(rec.get("정산여부(U)", "")).strip()
            if settle_status:
                ws.cell(next_row, 21, value=settle_status)

            written_rows.append(next_row)
            next_row += 1

        written_info[sheet_name] = written_rows

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue(), written_info, skipped_count


def discover_available_models(api_key: str) -> list[str]:
    """
    현재 API 키로 실제 호출 가능한 Gemini 텍스트·이미지 모델 목록을 서버에서 직접 조회합니다.
    """
    from google import genai

    client = genai.Client(api_key=api_key)
    models: list[str] = []
    for m in client.models.list():
        name = (getattr(m, "name", "") or "").replace("models/", "").strip()
        if not name or "gemini" not in name.lower():
            continue
        low = name.lower()
        if any(bad in low for bad in EXCLUDED_MODEL_KEYWORDS):
            continue
        actions = getattr(m, "supported_actions", None) or []
        if actions and "generateContent" not in actions:
            continue
        if name not in models:
            models.append(name)

    # 빠르고 혼잡이 적은 flash-lite -> flash -> 나머지 순으로 우선 정렬
    def sort_key(n: str) -> tuple[int, str]:
        low = n.lower()
        if "flash-lite" in low:
            return (0, n)
        if "flash" in low:
            return (1, n)
        return (2, n)

    models.sort(key=sort_key)
    return models


def extract_receipt_with_gemini(
    image_bytes: bytes,
    api_key: str,
    model_name: str,
    fallback_models: list[str] | None = None,
) -> tuple[ReceiptExtract, str]:
    """
    선택한 모델로 영수증을 인식하고, 503/429/404 발생 시 사용 가능한 예비 모델로 자동 전환합니다.
    Returns:
        (추출 결과, 실제 성공한 모델명)
    """
    from google import genai
    from google.genai import types
    from PIL import Image, ImageOps

    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((1600, 1600))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)

    client = genai.Client(api_key=api_key)
    prompt = (
        "한국 공사 현장 지출 영수증 이미지를 읽고 정산서 항목에 맞게 JSON으로 추출하세요.\n"
        "- spend_date: 결제일(YYYY-MM-DD). 연도가 생략된 경우 2026년 기준으로 보정하세요.\n"
        "- category: 반드시 [목자재, 전기자재, 기타공사, 부자재, 인건비, 간식, 식사, 차량] 중 하나를 선택하세요.\n"
        "- item_name: 정산서 셀에 들어갈 짧은 품목명(예: '공구함', '마스크 외 2건', '폐기물 반출', '주유비', '주차비').\n"
        "- amount: 할인 등이 반영된 최종 결제 금액(원 단위 정수, 쉼표 제외).\n"
        "- memo: 주유·주차 영수증에 차량번호가 있으면 '차량번호 주유비' 형태로 적으세요."
    )

    candidates = [model_name] + [
        m for m in (fallback_models or []) if m and m != model_name
    ]
    last_err: Exception | None = None

    for candidate in candidates[:6]:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=candidate,
                    contents=[
                        types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"),
                        prompt,
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=ReceiptExtract,
                        temperature=0.0,
                    ),
                )
                return ReceiptExtract.model_validate_json(response.text), candidate
            except Exception as exc:
                last_err = exc
                msg = str(exc)
                if "503" in msg or "UNAVAILABLE" in msg or "429" in msg:
                    time.sleep(1.5)
                    continue
                break

    raise RuntimeError(f"사용 가능한 모델 호출 실패 (마지막 오류: {last_err})")


def _get_secret(st_module, key: str) -> str:
    try:
        if key in st_module.secrets:
            return str(st_module.secrets[key]).strip()
    except Exception:
        pass
    return os.environ.get(key, "").strip()


def run_streamlit_app() -> None:
    import pandas as pd
    import streamlit as st

    st.set_page_config(
        page_title="공사 정산서 영수증 자동 입력기",
        layout="wide",
    )

    required_password = _get_secret(st, "APP_PASSWORD")
    if required_password:
        if not st.session_state.get("authenticated", False):
            st.title("공사 정산서 영수증 자동 입력기")
            st.info("팀 공용 비밀번호를 입력하면 정산 화면이 열립니다.")
            pw_input = st.text_input("접속 비밀번호", type="password")
            if st.button("입장하기", type="primary"):
                if pw_input == required_password:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")
            st.stop()

    st.title("공사 정산서 영수증 자동 입력기")
    st.caption(
        "영수증 사진을 올리면 AI가 지출일·항목·품목명·금액을 추출하고, "
        "기존 시트와 중복을 대조한 뒤 원하는 건만 정산서(.xlsx)에 자동 기입합니다."
    )

    secret_api_key = _get_secret(st, "GEMINI_API_KEY")

    @st.cache_data(ttl=600, show_spinner=False)
    def cached_models(key: str) -> tuple[list[str], str]:
        try:
            return discover_available_models(key), ""
        except Exception as exc:
            return [], str(exc)

    with st.sidebar:
        st.header("기본 설정")
        if secret_api_key:
            st.success("API 키가 서버에 안전하게 연결되어 있습니다.")
            api_key = secret_api_key
        else:
            api_key = st.text_input(
                "Gemini API Key",
                value="",
                type="password",
                help="Google AI Studio(aistudio.google.com/app/apikey)에서 발급한 키를 입력하세요.",
            )

        available_models: list[str] = []
        if api_key:
            available_models, list_err = cached_models(api_key)
            if available_models:
                st.caption(f"내 API 키에서 사용 가능한 모델 {len(available_models)}개 확인됨")
            elif list_err:
                st.warning(f"모델 목록 조회 실패: {list_err}")

        model_options = (available_models if available_models else ["gemini-3.8-flash", "gemini-3.6-flash"]) + ["직접 입력"]
        model_choice = st.selectbox("인식 모델 (자동 탐색)", options=model_options, index=0)
        if model_choice == "직접 입력":
            model_name = st.text_input("모델명 입력", value=model_options[0]).strip()
        else:
            model_name = model_choice

        default_user = st.selectbox(
            "기본 사용자(R열)",
            options=DEFAULT_USERS + ["직접 입력"],
            index=0,
        )
        if default_user == "직접 입력":
            default_user = st.text_input("사용자 이름 입력", value="")

        omit_dup_dates = st.checkbox(
            "같은 날짜가 연속될 때 지출일 칸 비우기 (기존 양식 스타일)",
            value=True,
        )

    # 1단계: 기존 정산서 엑셀 업로드
    st.subheader("1단계. 현재 사용 중인 정산서(.xlsx) 업로드")
    excel_file = st.file_uploader("정산서 엑셀 파일 선택", type=["xlsx"])

    excel_bytes: bytes | None = None
    target_sheets: list[str] = []
    if excel_file is not None:
        excel_bytes = excel_file.getvalue()
        wb_preview = load_workbook(io.BytesIO(excel_bytes), read_only=True)
        sheet_names = wb_preview.sheetnames
        wb_preview.close()

        target_sheets = st.multiselect(
            "영수증을 입력할 시트 선택 (전체 시트 + 주차 시트 동시 선택 가능)",
            options=sheet_names,
            default=[sheet_names[0]] if sheet_names else [],
        )

    # 2단계: 영수증 촬영 또는 다중 업로드
    st.subheader("2단계. 영수증 사진 촬영 또는 업로드")
    tab_upload, tab_camera = st.tabs(["파일 여러 장 업로드", "스마트폰·웹캠 촬영"])

    uploaded_images: list[tuple[str, bytes]] = []
    with tab_upload:
        files = st.file_uploader(
            "영수증 이미지 선택 (JPG, PNG)",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
        )
        if files:
            for f in files:
                uploaded_images.append((f.name, f.getvalue()))

    with tab_camera:
        cam_img = st.camera_input("영수증 촬영")
        if cam_img is not None:
            uploaded_images.append(("camera_receipt.jpg", cam_img.getvalue()))

    if "extracted_rows" not in st.session_state:
        st.session_state["extracted_rows"] = []

    if st.button(
        f"영수증 {len(uploaded_images)}건 AI 자동 인식 및 중복 확인",
        type="primary",
        disabled=(not uploaded_images or not api_key),
    ):
        results = []
        used_models = set()
        progress = st.progress(0.0)
        for idx, (fname, img_bytes) in enumerate(uploaded_images):
            try:
                ext, used_model = extract_receipt_with_gemini(
                    img_bytes,
                    api_key=api_key,
                    model_name=model_name,
                    fallback_models=available_models,
                )
                used_models.add(used_model)
                q_detail = ext.memo
                if ext.category == "차량" and not q_detail:
                    q_detail = (
                        f"{default_user} {ext.item_name}".strip()
                        if default_user
                        else ext.item_name
                    )
                results.append(
                    {
                        "반영": True,
                        "중복확인": "신규",
                        "파일명": fname,
                        "지출일": ext.spend_date,
                        "항목": ext.category,
                        "품목/내용": ext.item_name,
                        "금액": ext.amount,
                        "차량/기타상세(Q)": q_detail,
                        "사용자": default_user,
                        "비고(T)": "",
                        "정산여부(U)": "",
                        "가맹점(참고)": ext.merchant,
                    }
                )
            except Exception as exc:
                st.error(f"'{fname}' 인식 실패: {exc}")
            progress.progress((idx + 1) / len(uploaded_images))

        existing = (
            extract_existing_entries(excel_bytes, target_sheets)
            if (excel_bytes and target_sheets)
            else []
        )
        checked_results = check_duplicates(results, existing)
        st.session_state["extracted_rows"] = checked_results

        dup_cnt = sum(1 for r in checked_results if not r["반영"])
        model_info = f" (사용 모델: {', '.join(sorted(used_models))})" if used_models else ""
        if dup_cnt > 0:
            st.warning(
                f"{len(checked_results)}건 중 {dup_cnt}건에서 기존 내역과 중복이 감지되어 "
                f"'반영' 체크를 해제했습니다{model_info}. 그대로 놔두려면 해제 상태를 유지하고, "
                "별개 지출이라 추가해야 하면 '반영'을 다시 체크하세요."
            )
        elif checked_results:
            st.success(f"{len(checked_results)}건 인식을 완료했습니다 (중복 없음){model_info}.")

    # 3단계: 검수 및 중복 선택 테이블
    if st.session_state["extracted_rows"]:
        st.subheader("3단계. 중복 확인 및 검수")
        st.caption(
            "중복으로 감지된 행은 '반영' 체크가 꺼져 있어 엑셀에 추가되지 않고 기존 파일 그대로 유지됩니다. "
            "별개 영수증이라 추가가 필요하면 '반영'을 체크해 주세요."
        )
        df = pd.DataFrame(st.session_state["extracted_rows"])
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "반영": st.column_config.CheckboxColumn(
                    "반영",
                    help="체크하면 정산서에 입력하고, 해제하면 기존 정산서를 그대로 둡니다.",
                    default=True,
                ),
                "중복확인": st.column_config.TextColumn("중복확인", disabled=True),
                "항목": st.column_config.SelectboxColumn(
                    "항목", options=CATEGORY_OPTIONS, required=True
                ),
                "금액": st.column_config.NumberColumn(
                    "금액(원)", min_value=0, step=100, format="%d"
                ),
            },
        )

        # 4단계: 정산서 반영 및 다운로드
        st.subheader("4단계. 정산서 엑셀에 반영 후 다운로드")
        if excel_bytes is None or not target_sheets:
            st.warning("1단계에서 정산서 엑셀 파일과 입력할 시트를 선택해 주세요.")
        else:
            if st.button("선택한 항목만 정산서에 반영", type="primary"):
                try:
                    records = edited_df.fillna("").to_dict(orient="records")
                    updated_bytes, written_info, skipped_cnt = append_receipts_to_excel(
                        excel_bytes=excel_bytes,
                        target_sheets=target_sheets,
                        records=records,
                        omit_duplicate_dates=omit_dup_dates,
                    )
                    written_parts = [
                        f"'{s}' ({rows[0]}~{rows[-1]}행)"
                        for s, rows in written_info.items()
                        if rows
                    ]
                    if written_parts:
                        st.success(
                            f"반영 완료: {', '.join(written_parts)} "
                            f"(중복·미선택 {skipped_cnt}건은 기존 파일 그대로 유지)"
                        )
                    else:
                        st.info(
                            f"반영할 신규 항목이 없어 기존 정산서를 그대로 유지했습니다 "
                            f"(건너뜀 {skipped_cnt}건)."
                        )

                    base_name = excel_file.name.rsplit(".", 1)[0]
                    today_tag = dt.date.today().strftime("%Y%m%d")
                    st.download_button(
                        label="정산서(.xlsx) 다운로드",
                        data=updated_bytes,
                        file_name=f"{base_name}_영수증반영_{today_tag}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as exc:
                    st.error(f"엑셀 반영 중 오류가 발생했습니다: {exc}")


if __name__ == "__main__":
    run_streamlit_app()
