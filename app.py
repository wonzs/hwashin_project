import streamlit as st
import pandas as pd
import pyodbc
import fitz
import os
import tempfile
import shutil
from pathlib import Path
from anthropic import Anthropic
import win32com.client
import chromadb
from sentence_transformers import SentenceTransformer
import re
from qbs_generator import (
    save_pm_block,
    pm_block_exists,
    generate_qbs,
    get_pm_block_path,
)


# =========================
# 기본 설정
# =========================

BASE_DIR = Path(__file__).parent
ROOT_DIR = Path(r"\\210.80.98.20\E$\HTML5_EDS\zFile")

st.set_page_config(
    page_title="AI 성과품 지원",
    layout="wide"
)

@st.cache_resource
def get_db_conn():
    return pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=211.237.34.43,1312;"
        "DATABASE=ERP_TW;"
        "UID=sa;"
        "PWD=hsdb7834!;"
    )

conn = get_db_conn()

CHROMA_DIR = BASE_DIR / "chroma_db"

@st.cache_resource
def get_chroma_collection():
    client = chromadb.PersistentClient(path=str(CHROMA_DIR))
    return client.get_or_create_collection("qbs_reports")

qbs_collection = get_chroma_collection()

@st.cache_resource
def get_embed_model():
    return SentenceTransformer("jhgan/ko-sroberta-multitask")

embed_model = get_embed_model()


# =========================
# 공통 함수
# =========================

def convert_server_path(path):
    if not path:
        return ""

    path = str(path).replace("/", "\\").strip()

    if path.upper().startswith(r"E:\HTML5_EDS"):
        path = path.replace(
            r"E:\HTML5_EDS",
            r"\\210.80.98.20\E$\HTML5_EDS"
        )

    if path.startswith(r"\\210.80.98.20\HTML5_EDS"):
        path = path.replace(
            r"\\210.80.98.20\HTML5_EDS",
            r"\\210.80.98.20\E$\HTML5_EDS"
        )

    return path


def extract_hwp_text(hwp_path):
    hwp_path = Path(hwp_path)

    if not hwp_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwp_path}")

    temp_dir = Path(tempfile.mkdtemp())
    local_hwp_path = temp_dir / hwp_path.name
    temp_txt_path = temp_dir / "result.txt"

    hwp = None

    try:
        shutil.copy2(str(hwp_path), str(local_hwp_path))

        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True

        open_ok = hwp.Open(str(local_hwp_path))
        if not open_ok:
            raise Exception(f"HWP Open 실패: {local_hwp_path}")

        save_ok = hwp.SaveAs(str(temp_txt_path), "TEXT", "force:true")
        if not save_ok:
            raise Exception(f"TEXT SaveAs 실패: {temp_txt_path}")

        if not temp_txt_path.exists():
            raise Exception(f"TXT 파일이 생성되지 않았습니다: {temp_txt_path}")

        for enc in ["utf-16", "cp949", "utf-8"]:
            try:
                with open(temp_txt_path, "r", encoding=enc) as f:
                    return f.read()
            except UnicodeError:
                continue

        raise Exception("TXT 인코딩을 읽지 못했습니다.")

    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except:
                pass

        shutil.rmtree(temp_dir, ignore_errors=True)


def extract_pdf_text(pdf_path):
    text = ""

    doc = fitz.open(pdf_path)

    for page in doc:
        text += page.get_text()

    doc.close()

    return text


def extract_file_text(file_path, ext):
    ext = str(ext).lower().replace(".", "")
    file_path = convert_server_path(file_path)

    if ext in ["hwp", "hwpx"]:
        return extract_hwp_text(file_path)

    if ext == "pdf":
        return extract_pdf_text(file_path)

    raise Exception(f"아직 지원하지 않는 확장자입니다: {ext}")

def split_qbs_sections(text):
    pattern = r"(?=\n?\d\.\d(?:\.\d)?\s+[^\n]+)"
    parts = re.split(pattern, text)

    chunks = []

    for part in parts:
        part = part.strip()
        if len(part) < 100:
            continue

        first_line = part.splitlines()[0].strip()
        section_match = re.match(r"^(\d\.\d(?:\.\d)?)\s*(.*)", first_line)

        section_no = section_match.group(1) if section_match else ""
        section_title = section_match.group(2) if section_match else first_line[:50]

        chunks.append({
            "section_no": section_no,
            "section_title": section_title,
            "text": part
        })

    return chunks

def save_qbs_to_chroma(file_name, text, pm_emp_no="", pm_name="", project_type="", project_name=""):
    chunks = split_qbs_sections(text)

    if not chunks:
        return 0

    documents = []
    metadatas = []
    ids = []

    for idx, chunk in enumerate(chunks):
        doc_id = f"{file_name}_{idx}".replace(" ", "_").replace("/", "_").replace("\\", "_")

        documents.append(chunk["text"])
        metadatas.append({
            "file_name": file_name,
            "pm_emp_no": pm_emp_no,
            "pm_name": pm_name,
            "project_type": project_type,
            "project_name": project_name,
            "section_no": chunk["section_no"],
            "section_title": chunk["section_title"]
        })
        ids.append(doc_id)

    embeddings = embed_model.encode(documents).tolist()

    qbs_collection.upsert(
        ids=ids,
        documents=documents,
        metadatas=metadatas,
        embeddings=embeddings
    )

    return len(documents)

# =========================
# 로그인 / 사업 조회
# =========================

def login_user(user_id, password):
    sql = """
        SELECT
            A.USER_ID,
            A.NAME,
            A.DEPT_CODE,
            B.SHORT_NAME
        FROM ERP_TW.dbo.CO_AUTHORITY_USER A LEFT JOIN ERP_TW.dbo.CO_CODE_DEPT B ON A.DEPT_CODE = B.DEPT_CODE
        WHERE A.USER_ID = ?
          AND A.PASSWORD = ?
          AND A.USING_TAG = 'Y'
    """

    df = pd.read_sql(sql, conn, params=[user_id, password])

    if df.empty:
        return None

    return df.iloc[0].to_dict()


def get_my_projects(dept_code):
    base_sql = """
        SELECT
            CONT_NO,
            SVC_SRC_NM,
            CHARGE_DEPT_CD AS DEPT_CODE
        FROM ERP_TW.dbo.SM_CONT_MASTER
        WHERE AGREE_CODE = '11601'
    """

    if str(dept_code) != "20020":
        base_sql += """
          AND CHARGE_DEPT_CD = ?
        """

        return pd.read_sql(
            base_sql + " ORDER BY CONT_NO DESC",
            conn,
            params=[dept_code]
        )

    return pd.read_sql(
        base_sql + " ORDER BY CONT_NO DESC",
        conn
    )

def find_similar_projects(current_cont_no, project_name, dept_code, limit=5):
    sql = f"""
        SELECT TOP {limit}
            A.CONT_NO,
            A.SVC_SRC_NM,
            A.CHARGE_DEPT_CD
        FROM ERP_TW.dbo.SM_CONT_MASTER A
        WHERE A.AGREE_CODE = '11601'
          AND A.CONT_NO <> ?
          AND A.CHARGE_DEPT_CD = ?
          AND EXISTS (
              SELECT 1
              FROM EDS_TW.hint.OUTPUT_FILE_LIST F
              WHERE F.CONT_NO = A.CONT_NO
          )
        ORDER BY
            A.CONT_NO DESC
    """

    return pd.read_sql(sql, conn, params=[current_cont_no, dept_code])


def get_report_files(cont_no):
    sql = """
        SELECT
            CONT_NO,
            FILE_NAME,
            FILE_EXTENSION,
            FILE_PATH,
            FILE_PATH_CLIENT
        FROM EDS_TW.hint.OUTPUT_FILE_LIST
        WHERE CONT_NO = ?
          AND LOWER(FILE_EXTENSION) IN ('hwp', 'hwpx', 'pdf')
        ORDER BY FILE_PATH
    """

    return pd.read_sql(sql, conn, params=[cont_no])


def pick_report_files(file_df, max_files=2):
    if file_df.empty:
        return file_df

    file_df = file_df.copy()

    priority = file_df[
        file_df["FILE_NAME"].str.contains("보고|보고서|최종|요약|성과|본문", na=False)
    ]

    if not priority.empty:
        return priority.head(max_files)

    return file_df.head(max_files)


# =========================
# Claude AI 함수
# =========================

def build_reference_summary(reports):
    client = Anthropic(api_key=os.getenv("HWASHIN_API_KEY"))

    report_texts = []

    for idx, report in enumerate(reports, start=1):
        report_texts.append(f"""
[참고 성과품 {idx}]
사업번호: {report["project_no"]}
사업명: {report["project_name"]}
파일명: {report["file_name"]}

내용:
{report["text"][:10000]}
""")

    prompt = f"""
아래는 유사 사업의 기존 성과품 보고서들입니다.

이 자료들을 바탕으로 신규 보고서 작성에 활용할 수 있는 참고자료 요약본을 만들어주세요.

정리 형식:
1. 유사 사업들의 공통 목적
2. 공통적인 보고서 구성
3. 자주 등장하는 조사/분석 항목
4. 자주 쓰이는 문제점 표현
5. 자주 쓰이는 개선방향 표현
6. 참고 가능한 목차 구조
7. 도면/부록에 자주 수록되는 자료
8. 신규 보고서 작성 시 주의할 점

{chr(10).join(report_texts)}
"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=5000,
        temperature=0.2,
        messages=[
            {"role": "user", "content": prompt}
        ]
    )

    return message.content[0].text


def ask_report_ai(question):
    client = Anthropic(api_key=os.getenv("HWASHIN_API_KEY"))

    current_project = st.session_state.get("current_project", {})
    reference_summary = st.session_state.get("reference_summary", "")
    chat_history = st.session_state.get("chat_history", [])

    history_text = ""
    for msg in chat_history[-6:]:
        role = "사용자" if msg["role"] == "user" else "AI"
        history_text += f"\n[{role}]\n{msg['content']}\n"

    prompt = f"""
너는 엔지니어링 성과품 보고서 작성 보조 AI입니다.

사용자는 현재 아래 사업의 보고서를 작성 중입니다.

[현재 작업 사업]
사업번호: {current_project.get("cont_no", "")}
사업명: {current_project.get("project_name", "")}
부서코드: {current_project.get("dept_code", "")}

[유사 성과품 참고자료 요약본]
{reference_summary}

[이전 대화]
{history_text}

[사용자 질문]
{question}

답변 기준:
- 기존 성과품의 구성과 표현을 참고해서 답변
- 실제 보고서에 바로 붙여넣을 수 있는 문장 중심
- 필요하면 목차, 표 구성, 도면 목록, 작성 방향을 제안
- 모르는 내용은 임의로 단정하지 말고 "확인 필요"라고 표시
"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3500,
        temperature=0.3,
        messages=[
            {"role": "user", "content": prompt}
        ]
    )

    return message.content[0].text


def prepare_reference_for_project(current_project):
    similar_df = find_similar_projects(
        current_project["cont_no"],
        current_project["project_name"],
        current_project["dept_code"],
        limit=5
    )

    collected_reports = []

    for _, project in similar_df.iterrows():
        file_df = get_report_files(project["CONT_NO"])

        if file_df.empty:
            continue

        picked_files = pick_report_files(file_df, max_files=2)

        for _, file in picked_files.iterrows():
            real_path = convert_server_path(file["FILE_PATH"])
            ext = file["FILE_EXTENSION"]

            try:
                text = extract_file_text(real_path, ext)

                if len(text.strip()) == 0:
                    continue

                collected_reports.append({
                    "project_no": project["CONT_NO"],
                    "project_name": project["SVC_SRC_NM"],
                    "file_name": file["FILE_NAME"],
                    "path": real_path,
                    "text": text
                })

            except Exception as e:
                print("파일 읽기 실패:", file["FILE_NAME"], e)

    return collected_reports, similar_df


# =========================
# 로그인 화면
# =========================

if "login_user" not in st.session_state:
    st.title("AI 성과품 지원 시스템")

    st.subheader("로그인")

    user_id = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        user = login_user(user_id, password)

        if user is None:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
        else:
            st.session_state["login_user"] = user
            st.rerun()

    st.stop()


# =========================
# 로그인 후 화면
# =========================

login_user_info = st.session_state["login_user"]
dept_code = str(login_user_info["DEPT_CODE"]).strip()
dept_name = login_user_info["SHORT_NAME"]

st.title("AI 성과품 지원 시스템")

st.sidebar.write(f"👤 {login_user_info['NAME']}님")
st.sidebar.write(f"부서: {dept_name}")

if st.sidebar.button("로그아웃"):
    st.session_state.clear()
    st.rerun()

menu = st.sidebar.radio(
    "메뉴",
    ["AI 보고서 도우미", "QBS 자동생성", "QBS 기초자료 관리"]
)


# =========================
# 실제 사용자용 화면
# =========================

if menu == "AI 보고서 도우미":
    st.header("AI 보고서 도우미")

    if st.button("내 부서 사업 불러오기"):
        project_df = get_my_projects(dept_code)
        st.session_state["my_projects"] = project_df

        for key in ["current_project", "reference_summary", "collected_reports", "chat_history"]:
            if key in st.session_state:
                del st.session_state[key]

    if "my_projects" in st.session_state:
        project_df = st.session_state["my_projects"]

        if project_df.empty:
            st.warning("조회된 사업이 없습니다.")
        else:
            selected_project = st.selectbox(
                "작업할 사업을 선택하세요",
                project_df["CONT_NO"] + " / " + project_df["SVC_SRC_NM"]
            )

            selected_no = selected_project.split(" / ")[0]
            selected_row = project_df[project_df["CONT_NO"] == selected_no].iloc[0]

            st.session_state["current_project"] = {
                "cont_no": selected_row["CONT_NO"],
                "project_name": selected_row["SVC_SRC_NM"],
                "dept_code": selected_row["DEPT_CODE"]
            }

    if "current_project" in st.session_state:
        current_project = st.session_state["current_project"]

        st.subheader("현재 작업 사업")
        st.write(f"**{current_project['cont_no']} / {current_project['project_name']}**")

        use_ai_prepare = st.checkbox("유사 사업 성과품을 자동 분석합니다. 비용이 발생할 수 있습니다.")

        if use_ai_prepare:
            if st.button("AI 참고자료 자동 준비"):
                with st.spinner("유사 사업과 성과품 파일을 찾는 중입니다..."):
                    reports, similar_df = prepare_reference_for_project(current_project)

                st.session_state["collected_reports"] = reports
                st.session_state["similar_projects"] = similar_df

                if not reports:
                    st.warning("참고할 성과품 텍스트를 찾지 못했습니다.")
                else:
                    st.success(f"참고 성과품 {len(reports)}개를 추출했습니다.")

                    with st.spinner("Claude가 참고자료 요약본을 생성하는 중입니다..."):
                        reference_summary = build_reference_summary(reports)

                    st.session_state["reference_summary"] = reference_summary
                    st.session_state["chat_history"] = []

        else:
            st.info("체크박스를 선택해야 AI 참고자료 자동 준비 버튼이 활성화됩니다.")

    if "similar_projects" in st.session_state:
        with st.expander("AI가 참고한 유사 사업 목록"):
            st.dataframe(st.session_state["similar_projects"], use_container_width=True)

    if "collected_reports" in st.session_state:
        with st.expander("AI가 읽은 참고 성과품 파일"):
            for report in st.session_state["collected_reports"]:
                st.write(
                    "📄",
                    report["project_no"],
                    "/",
                    report["project_name"],
                    "/",
                    report["file_name"],
                    f"({len(report['text']):,}자)"
                )

    if "reference_summary" in st.session_state:
        with st.expander("참고자료 요약본 보기"):
            st.markdown(st.session_state["reference_summary"])

        st.subheader("보고서 작성 AI 대화")

        if "chat_history" not in st.session_state:
            st.session_state["chat_history"] = []

        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        question = st.chat_input("예: 보고서 목차를 어떻게 구성하면 좋을까?")

        if question:
            st.session_state["chat_history"].append({
                "role": "user",
                "content": question
            })

            with st.chat_message("user"):
                st.markdown(question)

            with st.chat_message("assistant"):
                with st.spinner("답변 작성 중입니다..."):
                    answer = ask_report_ai(question)

                st.markdown(answer)

            st.session_state["chat_history"].append({
                "role": "assistant",
                "content": answer
            })

# =========================
# QBS 기초자료 관리 (PM 블록 + 표준문구 + 벡터DB 등록 통합)
# =========================

elif menu == "QBS 기초자료 관리":
    st.header("QBS 기초자료 관리")

    tab1, tab2, tab3 = st.tabs(["📁 PM 실적 블록", "📝 표준 문구 관리", "🗄️ QBS 벡터DB 등록"])

    # ── 탭1: PM 실적 블록 ──────────────────────────────────────
    with tab1:
        st.markdown(
            """
            PM(사업책임기술인)의 **"1.2 당해 용역 관련 유사사업 제시"** 절
            (실적 표 + 유사사업 세부설명)을 **별도 hwpx 파일로 추출**하여 업로드해주세요.

            - 동일 PM의 QBS 생성 시 이 블록이 그대로 삽입됩니다.
            - 한글에서 해당 페이지 범위만 블록 선택 → 복사 → 새 문서에 붙여넣기 →
              'HWPX' 형식으로 저장한 파일을 업로드하면 됩니다.
            """
        )

        col1, col2 = st.columns(2)
        with col1:
            pm_emp_no_tab1 = st.text_input("PM 사번", placeholder="예: 2020123", key="pm_block_emp_no")
        with col2:
            pm_name_tab1 = st.text_input("PM 성명", placeholder="예: 김철수", key="pm_block_name")

        if pm_emp_no_tab1:
            existing_path = get_pm_block_path(BASE_DIR, pm_emp_no_tab1)
            if existing_path.exists():
                st.info(
                    f"✅ 현재 등록된 블록: {existing_path.name}  "
                    f"(수정일: {pd.to_datetime(existing_path.stat().st_mtime, unit='s').strftime('%Y-%m-%d %H:%M')})"
                )

        uploaded_block = st.file_uploader(
            "PM 실적 블록 파일 업로드 (.hwpx)",
            type=["hwpx"],
            accept_multiple_files=False,
            key="pm_block_uploader"
        )

        if st.button("PM 실적 블록 저장", key="pm_block_save"):
            if not pm_emp_no_tab1:
                st.warning("PM 사번을 입력해주세요.")
            elif not uploaded_block:
                st.warning("업로드할 hwpx 파일을 선택해주세요.")
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx") as tmp:
                    tmp.write(uploaded_block.getbuffer())
                    tmp_path = tmp.name
                try:
                    saved_path = save_pm_block(BASE_DIR, pm_emp_no_tab1, tmp_path)
                    st.success(f"PM({pm_emp_no_tab1} / {pm_name_tab1}) 실적 블록 저장 완료: {saved_path.name}")
                except Exception as e:
                    st.error(f"저장 실패: {e}")
                finally:
                    try:
                        os.remove(tmp_path)
                    except:
                        pass

        st.divider()
        st.subheader("등록된 PM 블록 목록")
        pm_block_dir = BASE_DIR / "pm_blocks"
        if pm_block_dir.exists():
            blocks = sorted(pm_block_dir.glob("*.hwpx"))
            if blocks:
                for b in blocks:
                    col_a, col_b = st.columns([4, 1])
                    with col_a:
                        st.write(f"📄 사번: {b.stem}  ·  파일: {b.name}")
                    with col_b:
                        if st.button("삭제", key=f"del_pm_{b.stem}"):
                            b.unlink()
                            st.success(f"{b.name} 삭제됨")
                            st.rerun()
            else:
                st.info("등록된 PM 블록이 없습니다.")
        else:
            st.info("등록된 PM 블록이 없습니다.")

    # ── 탭2: 표준 문구 관리 ────────────────────────────────────
    with tab2:
        import json

        STANDARD_TEXTS_PATH = BASE_DIR / "templates" / "standard_texts.json"

        def load_standard_texts():
            if STANDARD_TEXTS_PATH.exists():
                with open(STANDARD_TEXTS_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            # templates 폴더에 없으면 앱 루트에서 찾기 (초기 배포용)
            fallback = BASE_DIR / "standard_texts.json"
            if fallback.exists():
                with open(fallback, "r", encoding="utf-8") as f:
                    return json.load(f)
            return {}

        def save_standard_texts(data):
            STANDARD_TEXTS_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(STANDARD_TEXTS_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        st.markdown("섹션별 표준 문구를 편집합니다. 발주처 양식이 다른 경우 이 화면에서 수정하세요.")

        texts = load_standard_texts()

        if not texts:
            st.warning("standard_texts.json 파일이 없습니다. templates 폴더에 파일을 복사해주세요.")
        else:
            section_labels = {
                "1.3": "1.3 용역 성공을 위한 요건 및 추가제안사항",
                "2.1": "2.1 당해 용역 관련 업무 수행범위",
                "2.3": "2.3 품질보증체계 및 품질관리계획",
                "2.4": "2.4 전문가 활용계획 등 추가 업무관리사항",
            }

            selected_section = st.selectbox(
                "편집할 섹션 선택",
                list(section_labels.keys()),
                format_func=lambda k: section_labels.get(k, k),
                key="std_section_select"
            )

            section_data = texts.get(selected_section, {})
            subsections = section_data.get("subsections", [])

            if not subsections:
                st.info("해당 섹션에 소절이 없습니다.")
            else:
                edited_subsections = []
                for i, sub in enumerate(subsections):
                    with st.expander(f"📌 {sub['title']}", expanded=(i == 0)):
                        new_title = st.text_input(
                            "소절 제목",
                            value=sub["title"],
                            key=f"std_title_{selected_section}_{i}"
                        )
                        new_content = st.text_area(
                            "내용",
                            value=sub["content"],
                            height=250,
                            key=f"std_content_{selected_section}_{i}"
                        )
                        edited_subsections.append({"title": new_title, "content": new_content})

                col_save, col_reset = st.columns([1, 1])
                with col_save:
                    if st.button("💾 저장", key=f"std_save_{selected_section}"):
                        texts[selected_section]["subsections"] = edited_subsections
                        save_standard_texts(texts)
                        st.success("표준 문구가 저장되었습니다.")
                with col_reset:
                    if st.button("↩️ 기본값으로 초기화", key=f"std_reset_{selected_section}"):
                        fallback = BASE_DIR / "standard_texts.json"
                        if fallback.exists():
                            with open(fallback, "r", encoding="utf-8") as f:
                                default_data = json.load(f)
                            texts[selected_section] = default_data.get(selected_section, section_data)
                            save_standard_texts(texts)
                            st.success("기본값으로 초기화되었습니다.")
                            st.rerun()
                        else:
                            st.error("기본값 파일(standard_texts.json)을 찾을 수 없습니다.")

    # ── 탭3: QBS 벡터DB 등록 (기존 "QBS 자료 등록" 기능) ───────
    with tab3:
        st.markdown("기존 QBS 파일을 ChromaDB에 등록하면 AI가 유사 사례로 참고합니다.")

        uploaded_files = st.file_uploader(
            "기존 QBS 파일 업로드",
            type=["hwp", "hwpx", "pdf"],
            accept_multiple_files=True,
            key="qbs_db_uploader"
        )

        col1, col2 = st.columns(2)
        with col1:
            db_pm_emp_no = st.text_input("PM 사번", placeholder="예: 2020123", key="db_pm_emp_no")
            db_project_type = st.text_input("사업유형", placeholder="예: 하수도, 상수도, 도로", key="db_project_type")
        with col2:
            db_pm_name = st.text_input("PM 성명", placeholder="예: 홍길동", key="db_pm_name")
            db_project_name = st.text_input("사업명", placeholder="예: 연일 하수도정비 중점관리지역 정비사업", key="db_project_name")

        if st.button("QBS 벡터DB에 저장", key="qbs_db_save"):
            if not uploaded_files:
                st.warning("업로드할 QBS 파일을 선택해주세요.")
            else:
                total_count = 0
                for uploaded_file in uploaded_files:
                    suffix = Path(uploaded_file.name).suffix.lower().replace(".", "")
                    with tempfile.NamedTemporaryFile(delete=False, suffix="." + suffix) as tmp:
                        tmp.write(uploaded_file.getbuffer())
                        tmp_path = tmp.name
                    try:
                        with st.spinner(f"{uploaded_file.name} 분석 중..."):
                            text = extract_file_text(tmp_path, suffix)
                            count = save_qbs_to_chroma(
                                file_name=uploaded_file.name,
                                text=text,
                                pm_emp_no=db_pm_emp_no,
                                pm_name=db_pm_name,
                                project_type=db_project_type,
                                project_name=db_project_name
                            )
                            total_count += count
                        st.success(f"{uploaded_file.name}: {count}개 문단 저장 완료")
                    except Exception as e:
                        st.error(f"{uploaded_file.name} 처리 실패: {e}")
                    finally:
                        try:
                            os.remove(tmp_path)
                        except:
                            pass
                st.success(f"총 {total_count}개 문단을 ChromaDB에 저장했습니다.")

# =========================
# "QBS 자동생성"
#     - 사업 선택 + PM 정보 입력 -> 전체 파이프라인 실행 -> hwpx 다운로드
# =========================

elif menu == "QBS 자동생성":
    st.header("QBS 자동생성")
 
    if st.button("내 부서 사업 불러오기", key="qbs_gen_load_projects"):
        project_df = get_my_projects(dept_code)
        st.session_state["qbs_gen_projects"] = project_df
 
        for key in ["qbs_gen_current_project", "qbs_gen_result_path"]:
            if key in st.session_state:
                del st.session_state[key]
 
    if "qbs_gen_projects" in st.session_state:
        project_df = st.session_state["qbs_gen_projects"]
 
        if project_df.empty:
            st.warning("조회된 사업이 없습니다.")
        else:
            selected_project = st.selectbox(
                "QBS를 생성할 사업을 선택하세요",
                project_df["CONT_NO"] + " / " + project_df["SVC_SRC_NM"],
                key="qbs_gen_project_select",
            )
 
            selected_no = selected_project.split(" / ")[0]
            selected_row = project_df[project_df["CONT_NO"] == selected_no].iloc[0]
 
            st.session_state["qbs_gen_current_project"] = {
                "cont_no": selected_row["CONT_NO"],
                "project_name": selected_row["SVC_SRC_NM"],
                "dept_code": selected_row["DEPT_CODE"],
            }
 
    if "qbs_gen_current_project" in st.session_state:
        current_project = st.session_state["qbs_gen_current_project"]
 
        st.subheader("선택된 사업")
        st.write(f"**{current_project['cont_no']} / {current_project['project_name']}**")
 
        st.markdown("##### 사업 정보 (표지/머릿글에 들어갈 정보)")
        col1, col2 = st.columns(2)
        with col1:
            announce_date = st.text_input("공고일", placeholder="예: 2024. 02. 23.")
            amount = st.text_input("계약금액", placeholder="예: 1,930,000,000원")
        with col2:
            submit_date = st.text_input("제출일", placeholder="예: 2024. 03. 08.")
 
        st.markdown("##### PM(사업책임기술인) 정보")
        pm_emp_no = st.text_input("PM 사번", placeholder="예: 2020123", key="qbs_gen_pm_emp_no")
        pm_name = st.text_input("PM 성명", placeholder="예: 김철수", key="qbs_gen_pm_name")
        writer_name = st.text_input("작성자 성명", placeholder="예: 김영희", key="qbs_gen_writer_name")
 
        if pm_emp_no and not pm_block_exists(BASE_DIR, pm_emp_no):
            st.warning(
                f"PM({pm_emp_no})의 실적 블록이 등록되어 있지 않습니다. "
                f"'PM 실적 블록 관리' 메뉴에서 먼저 등록해주세요."
            )
 
        if st.button("QBS 생성 시작"):
            if not pm_emp_no or not pm_name:
                st.warning("PM 사번/성명을 입력해주세요.")
            elif not pm_block_exists(BASE_DIR, pm_emp_no):
                st.error("PM 실적 블록이 등록되어 있지 않아 생성할 수 없습니다.")
            else:
                current_project["announce_date"] = announce_date
                current_project["amount"] = amount
                current_project["submit_date"] = submit_date
 
                with st.spinner("유사 사례 검색 및 1.1절 작성 중..."):
                    try:
                        result_path = generate_qbs(
                            base_dir=BASE_DIR,
                            current_project=current_project,
                            pm_emp_no=pm_emp_no,
                            pm_name=pm_name,
                            writer_name=st.session_state.get("qbs_gen_writer_name", ""),
                            qbs_collection=qbs_collection,
                            embed_model=embed_model,
                        )
                        st.session_state["qbs_gen_result_path"] = result_path
                    except Exception as e:
                        st.error(f"QBS 생성 실패: {e}")
 
    if "qbs_gen_result_path" in st.session_state:
        result_path = st.session_state["qbs_gen_result_path"]
 
        st.success(f"QBS 파일이 생성되었습니다: {result_path.name}")
 
        with open(result_path, "rb") as f:
            st.download_button(
                label="QBS 파일 다운로드 (.hwpx)",
                data=f.read(),
                file_name=result_path.name,
                mime="application/octet-stream",
            )
 
        st.info(
            "생성된 1.1절 내용은 AI가 작성한 초안입니다. "
            "한글에서 열어 [확인 필요] 표시된 부분과 전체 내용을 반드시 검토 후 사용하세요."
        )

