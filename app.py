import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn  
import io

def create_lined_notebook(doc, lines_per_page=25, num_pages=5):
   
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        for i in range(lines_per_page):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(20)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            
            
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'CCCCCC')
            pBdr.append(bottom)
            p._element.get_or_add_pPr().append(pBdr)

def create_grid_notebook(doc, rows=20, cols=20, num_pages=5):
    
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        
        
        for row in table.rows:
            row.height = Inches(0.3)
            for cell in row.cells:
                cell.width = Inches(0.3)
                
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)
                cell.paragraphs[0].paragraph_format.space_before = Pt(0)
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                tcMar = OxmlElement('w:tcMar')
                for margin_type in ['top', 'left', 'bottom', 'right']:
                    margin = OxmlElement(f'w:{margin_type}')
                    margin.set(qn('w:w'), '0')
                    margin.set(qn('w:type'), 'dxa')
                    tcMar.append(margin)
                tcPr.append(tcMar)

def create_english_notebook(doc, lines_per_page=15, num_pages=5):
    
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        for i in range(lines_per_page):
            # 각 줄마다 4선 테이블 생성
            table = doc.add_table(rows=1, cols=1)
            table.autofit = False
            
            cell = table.cell(0, 0)
            cell.width = Inches(6.5)
            cell.height = Inches(0.5)
            
          
            p1 = cell.paragraphs[0]
            p1.paragraph_format.space_after = Pt(0)
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.line_spacing = Pt(6)
            
            
            pBdr1 = OxmlElement('w:pBdr')
            top1 = OxmlElement('w:top')
            top1.set(qn('w:val'), 'dotted')
            top1.set(qn('w:sz'), '2')
            top1.set(qn('w:color'), 'CCCCCC')
            pBdr1.append(top1)
            p1._element.get_or_add_pPr().append(pBdr1)
            
           
            p2 = cell.add_paragraph()
            p2.paragraph_format.space_after = Pt(6)
            p2.paragraph_format.space_before = Pt(6)
            
            pBdr2 = OxmlElement('w:pBdr')
            bottom2 = OxmlElement('w:bottom')
            bottom2.set(qn('w:val'), 'single')
            bottom2.set(qn('w:sz'), '4')
            bottom2.set(qn('w:color'), '000000')
            pBdr2.append(bottom2)
            p2._element.get_or_add_pPr().append(pBdr2)
            
           
            p3 = cell.add_paragraph()
            p3.paragraph_format.space_after = Pt(0)
            p3.paragraph_format.space_before = Pt(6)
            
            pBdr3 = OxmlElement('w:pBdr')
            bottom3 = OxmlElement('w:bottom')
            bottom3.set(qn('w:val'), 'single')
            bottom3.set(qn('w:sz'), '6')
            bottom3.set(qn('w:color'), '000000')
            pBdr3.append(bottom3)
            p3._element.get_or_add_pPr().append(pBdr3)
            
            # 줄 간격 추가
            doc.add_paragraph().paragraph_format.space_after = Pt(10)

def create_cornell_notebook(doc, num_pages=5):
    """코넬노트 양식 생성"""
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        # 상단 제목 영역
        title = doc.add_paragraph("제목: ")
        title.paragraph_format.space_after = Pt(12)
        title.runs[0].font.bold = True
        
        # 날짜 영역
        date = doc.add_paragraph("날짜: ")
        date.paragraph_format.space_after = Pt(12)
        date.runs[0].font.bold = True
        
        # 메인 영역 (핵심어 | 노트)
        table = doc.add_table(rows=1, cols=2)
        table.columns[0].width = Inches(2)
        table.columns[1].width = Inches(4.5)
        
        # 왼쪽 열 (핵심어)
        key_cell = table.cell(0, 0)
        key_cell.text = "핵심어/질문"
        key_cell.paragraphs[0].runs[0].font.bold = True
        key_cell.height = Inches(6)
        
        # 오른쪽 열 (노트)
        note_cell = table.cell(0, 1)
        note_cell.text = "노트 영역"
        note_cell.paragraphs[0].runs[0].font.bold = True
        
        # 하단 요약 영역
        doc.add_paragraph()
        summary_title = doc.add_paragraph("요약:")
        summary_title.runs[0].font.bold = True
        summary_title.paragraph_format.space_before = Pt(12)
        
        # 요약 영역 박스
        summary_table = doc.add_table(rows=1, cols=1)
        summary_cell = summary_table.cell(0, 0)
        summary_cell.height = Inches(1.5)

# Streamlit 앱 메인 함수
st.set_page_config(page_title="노트 양식 생성기", page_icon="📝", layout="wide")

st.title("📝 노트 양식 생성기")
st.markdown("다양한 노트 양식을 선택하고 Word 파일로 다운로드하세요!")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("⚙️ 설정")
    
    # 노트 종류 선택
    notebook_type = st.selectbox(
        "노트 종류 선택",
        ["줄공책", "칸공책", "영어노트 (4선)", "코넬노트"]
    )
    
    # 페이지 수
    num_pages = st.number_input("페이지 수", min_value=1, max_value=50, value=5)
    
    # 용지 방향
    orientation = st.radio("용지 방향", ["세로", "가로"])
    
    # 노트별 추가 설정
    if notebook_type == "줄공책":
        lines_per_page = st.slider("페이지당 줄 수", 10, 40, 25)
    elif notebook_type == "칸공책":
        rows = st.slider("행 수", 10, 30, 20)
        cols = st.slider("열 수", 10, 30, 20)
    elif notebook_type == "영어노트 (4선)":
        lines_per_page = st.slider("페이지당 줄 수", 5, 20, 15)
    
    # 생성 버튼
    if st.button("📄 노트 생성", use_container_width=True, type="primary"):
        with st.spinner("노트를 생성하고 있습니다..."):
            # Document 생성
            doc = Document()
            
            # 용지 방향 설정
            section = doc.sections[0]
            if orientation == "가로":
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width, section.page_height = section.page_height, section.page_width
            
            # 여백 설정
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            # 선택된 노트 종류에 따라 생성
            if notebook_type == "줄공책":
                create_lined_notebook(doc, lines_per_page, num_pages)
            elif notebook_type == "칸공책":
                create_grid_notebook(doc, rows, cols, num_pages)
            elif notebook_type == "영어노트 (4선)":
                create_english_notebook(doc, lines_per_page, num_pages)
            elif notebook_type == "코넬노트":
                create_cornell_notebook(doc, num_pages)
            
            # 메모리에 저장
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            # 다운로드 버튼
            st.download_button(
                label="📥 Word 파일 다운로드",
                data=doc_io.getvalue(),
                file_name=f"{notebook_type}_{num_pages}페이지.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            st.success("✅ 노트가 성공적으로 생성되었습니다!")

with col2:
    st.subheader("📖 사용 방법")
    st.markdown("""
    1. **노트 종류 선택**: 원하는 노트 양식을 선택하세요.
    2. **페이지 수 설정**: 생성할 페이지 수를 입력하세요.
    3. **용지 방향 선택**: 세로 또는 가로 방향을 선택하세요.
    4. **추가 설정**: 노트 종류에 따라 줄 수, 칸 수 등을 조정하세요.
    5. **노트 생성**: '노트 생성' 버튼을 클릭하세요.
    6. **다운로드**: 생성된 Word 파일을 다운로드하세요.
    """)
    
    st.subheader("📝 노트 종류 설명")
    with st.expander("줄공책"):
        st.markdown("""
        - 일반적인 줄이 그어진 노트
        - 글쓰기, 일기, 메모 등에 적합
        - 페이지당 줄 수 조정 가능
        """)
    
    with st.expander("칸공책"):
        st.markdown("""
        - 격자 모양의 칸으로 구성된 노트
        - 수학, 도표, 그래프 그리기에 적합
        - 행과 열의 수 조정 가능
        """)
    
    with st.expander("영어노트 (4선)"):
        st.markdown("""
        - 영어 필기체 연습용 4선 노트
        - 알파벳 쓰기 연습에 최적화
        - 점선과 실선으로 구성
        """)
    
    with st.expander("코넬노트"):
        st.markdown("""
        - 효과적인 학습을 위한 노트 양식
        - 핵심어/질문, 노트, 요약 영역으로 구분
        - 체계적인 학습 정리에 적합
        """)
