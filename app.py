import io
import re

import google.generativeai as genai
import streamlit as st
from docx import Document
from PIL import Image


def _split_table_row(row: str) -> list[str]:
    row = row.strip()
    if row.startswith("|"):
        row = row[1:]
    if row.endswith("|"):
        row = row[:-1]
    return [c.strip() for c in row.split("|")]


def _is_markdown_table_separator(line: str) -> bool:
    if "|" not in line:
        return False
    cells = _split_table_row(line)
    if not cells:
        return False
    for c in cells:
        t = c.strip().replace(" ", "")
        if not re.match(r"^:?-{2,}:?$", t):
            return False
    return True


def _parse_markdown_segments(markdown: str) -> list[tuple[str, object]]:
    lines = markdown.splitlines()
    segments: list[tuple[str, object]] = []
    i = 0
    while i < len(lines):
        if not lines[i].strip():
            i += 1
            continue

        stripped = lines[i].strip()
        if (
            "|" in stripped
            and i + 1 < len(lines)
            and _is_markdown_table_separator(lines[i + 1])
        ):
            header = _split_table_row(lines[i])
            i += 2
            rows: list[list[str]] = [header]
            while i < len(lines):
                s = lines[i].strip()
                if not s:
                    break
                if "|" not in s:
                    break
                if _is_markdown_table_separator(lines[i]):
                    i += 1
                    continue
                rows.append(_split_table_row(lines[i]))
                i += 1
            segments.append(("table", rows))
            continue

        block_lines: list[str] = []
        while i < len(lines):
            s = lines[i]
            if not s.strip():
                i += 1
                break
            stp = s.strip()
            if "|" in stp and i + 1 < len(lines) and _is_markdown_table_separator(
                lines[i + 1]
            ):
                break
            block_lines.append(s)
            i += 1
        if block_lines:
            segments.append(("text", "\n".join(block_lines)))

    return segments


def markdown_to_docx_buffer(markdown: str) -> io.BytesIO:
    doc = Document()
    for kind, payload in _parse_markdown_segments(markdown):
        if kind == "text":
            block = str(payload)
            current_para: list[str] = []

            def flush_para() -> None:
                if current_para:
                    doc.add_paragraph("\n".join(current_para))
                    current_para.clear()

            for line in block.split("\n"):
                s = line.strip()
                if not s:
                    flush_para()
                    continue
                if s.startswith("#"):
                    flush_para()
                    level = len(s) - len(s.lstrip("#"))
                    level = min(max(level, 1), 3)
                    doc.add_heading(s.lstrip("#").strip(), level=level)
                else:
                    current_para.append(line.rstrip())
            flush_para()

        elif kind == "table":
            rows: list[list[str]] = payload  # type: ignore[assignment]
            if not rows:
                continue
            num_cols = max(len(r) for r in rows)
            norm_rows: list[list[str]] = []
            for r in rows:
                extended = list(r) + [""] * (num_cols - len(r))
                norm_rows.append(extended[:num_cols])

            table = doc.add_table(rows=len(norm_rows), cols=num_cols)
            table.style = "Table Grid"

            for ri, row_cells in enumerate(norm_rows):
                for ci, cell_text in enumerate(row_cells):
                    cell = table.rows[ri].cells[ci]
                    cell.text = cell_text
                    if ri == 0:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.bold = True

            doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="시험지 깔끔 변환기", layout="centered")

GEMINI_PROMPT = (
    "이 이미지에서 볼펜이나 연필로 쓰인 낙서, 빗금, 풀이 과정은 모두 무시해. "
    "인쇄된 문제 텍스트, 보기(①, ②, ③...), 그리고 표 형식만 완벽하게 추출해. "
    "특히 표 구조는 마크다운(Markdown)의 Table 형식으로 정확하게 구현해서 출력해 줘."
)

with st.sidebar:
    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        help="Google AI Studio에서 발급받은 API 키를 입력하세요.",
    )

st.title("시험지 깔끔 변환기")

uploaded_file = st.file_uploader(
    "JPG, PNG, PDF 파일을 선택하세요",
    type=["jpg", "jpeg", "png", "pdf"],
)

key_ok = bool(api_key and api_key.strip())
file_ok = uploaded_file is not None
can_convert = key_ok and file_ok

if st.button("변환 시작", disabled=not can_convert):
    st.session_state.pop("gemini_markdown", None)
    file_bytes = uploaded_file.read()
    name_lower = (uploaded_file.name or "").lower()
    is_pdf = uploaded_file.type == "application/pdf" or name_lower.endswith(".pdf")

    genai.configure(api_key=api_key.strip())
    model = genai.GenerativeModel("gemini-2.5-flash")

    with st.spinner("변환 중..."):
        try:
            if is_pdf:
                content = [
                    GEMINI_PROMPT,
                    {"mime_type": "application/pdf", "data": file_bytes},
                ]
            else:
                image = Image.open(io.BytesIO(file_bytes))
                content = [GEMINI_PROMPT, image]

            response = model.generate_content(content)
            try:
                result_text = response.text
            except ValueError:
                result_text = None
                st.error("응답을 가져올 수 없습니다. 안전 필터 또는 빈 응답일 수 있습니다.")

            if result_text:
                st.session_state["gemini_markdown"] = result_text
        except Exception as e:
            st.error(f"API 호출 오류: {e}")

if st.session_state.get("gemini_markdown"):
    st.divider()
    st.subheader("추출 결과")
    st.markdown(st.session_state["gemini_markdown"])

    doc_buffer = markdown_to_docx_buffer(st.session_state["gemini_markdown"])
    st.download_button(
        label="워드 파일 다운로드",
        data=doc_buffer,
        file_name="시험지_추출결과.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
