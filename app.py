import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

def create_lined_notebook(doc, lines_per_page=25, num_pages=5):
    """줄공책 양식 생성 - 테이블 방식"""
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        # 페이지 상단 여백
        top_para = doc.add_paragraph()
        top_para.paragraph_format.space_after = Pt(10)
        
        # 테이블을 사용한 줄 생성
        table = doc.add_table(rows=lines_per_page, cols=1)
        table.autofit = False
        table.style = 'Normal Table'
        
        for i, row in enumerate(table.rows):
            # 행 높이 설정
            row.height = Pt(28)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            
            cell = row.cells[0]
            cell.width = Inches(7.5)
            
            # 셀 내부 단락 설정
            if cell.paragraphs:
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            
            # 셀 테두리 설정
            tc = cell._element
            tcPr = tc.get_or_add_tcPr()
            
            # 기존 테두리 제거
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is not None:
                tcPr.remove(tcBorders)
            
            # 새 테두리 설정
            tcBorders = OxmlElement('w:tcBorders')
            
            # 하단 선만 추가
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), '808080')
            tcBorders.append(bottom)
            
            # 나머지 테두리는 없음
            for border in ['top', 'left', 'right']:
                side = OxmlElement(f'w:{border}')
                side.set(qn('w:val'), 'nil')
                tcBorders.append(side)
            
            tcPr.append(tcBorders)
            
            # 셀 여백 설정
            tcMar = tcPr.find(qn('w:tcMar'))
            if tcMar is not None:
                tcPr.remove(tcMar)
                
            tcMar = OxmlElement('w:tcMar')
            for margin_name in ['top', 'left', 'bottom', 'right']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '50')
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)
            tcPr.append(tcMar)

def create_grid_notebook(doc, rows=15, cols=15, num_pages=5):
    """칸공책 양식 생성"""
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        # 페이지 크기 계산 (A4 기준)
        page_width = 8.27 - 1.0  # 인치 (여백 제외)
        page_height = 11.69 - 1.0
        
        cell_width = page_width / cols
        cell_height = page_height / rows
        
        # 테이블 생성
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.allow_autofit = False
        
        # 각 행 설정
        for row in table.rows:
            # 행 높이 설정
            tr = row._element
            trPr = tr.get_or_add_trPr()
            
            # 기존 높이 설정 제거
            for child in trPr:
                if child.tag.endswith('trHeight'):
                    trPr.remove(child)
            
            # 새 높이 설정
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), str(int(cell_height * 1440)))  # twips
            trHeight.set(qn('w:hRule'), 'exact')
            trPr.append(trHeight)
            
            # 각 셀 설정
            for cell in row.cells:
                # 셀 너비 설정
                cell.width = Inches(cell_width)
                
                # 셀 내용 설정
                if cell.paragraphs:
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = Pt(0)
                
                # 셀 여백 최소화
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                
                # 기존 여백 제거
                tcMar = tcPr.find(qn('w:tcMar'))
                if tcMar is not None:
                    tcPr.remove(tcMar)
                
                # 새 여백 설정
                tcMar = OxmlElement('w:tcMar')
                for margin_name in ['top', 'left', 'bottom', 'right']:
                    margin = OxmlElement(f'w:{margin_name}')
                    margin.set(qn('w:w'), '10')
                    margin.set(qn('w:type'), 'dxa')
                    tcMar.append(margin)
                tcPr.append(tcMar)

def create_english_notebook(doc, lines_per_page=12, num_pages=5):
    """영어노트 양식 생성 (4선 노트)"""
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        # 페이지 상단 여백
        top_margin = doc.add_paragraph()
        top_margin.paragraph_format.space_after = Pt(20)
        
        for i in range(lines_per_page):
            # 4선을 위한 테이블 생성
            table = doc.add_table(rows=4, cols=1)
            table.autofit = False
            table.style = 'Normal Table'
            
            # 첫 번째 선 (상단 점선)
            row1 = table.rows[0]
            row1.height = Pt(10)
            row1.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            cell1 = row1.cells[0]
            cell1.width = Inches(7.5)
            
            # 점선 스타일
            tc1 = cell1._element
            tcPr1 = tc1.get_or_add_tcPr()
            tcBorders1 = OxmlElement('w:tcBorders')
            bottom1 = OxmlElement('w:bottom')
            bottom1.set(qn('w:val'), 'dotted')
            bottom1.set(qn('w:sz'), '4')
            bottom1.set(qn('w:color'), 'CCCCCC')
            tcBorders1.append(bottom1)
            tcPr1.append(tcBorders1)
            
            # 두 번째 선 (상단 실선)
            row2 = table.rows[1]
            row2.height = Pt(10)
            row2.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            cell2 = row2.cells[0]
            
            tc2 = cell2._element
            tcPr2 = tc2.get_or_add_tcPr()
            tcBorders2 = OxmlElement('w:tcBorders')
            bottom2 = OxmlElement('w:bottom')
            bottom2.set(qn('w:val'), 'single')
            bottom2.set(qn('w:sz'), '4')
            bottom2.set(qn('w:color'), '808080')
            tcBorders2.append(bottom2)
            tcPr2.append(tcBorders2)
            
            # 세 번째 선 (기준선 - 굵은 실선)
            row3 = table.rows[2]
            row3.height = Pt(10)
            row3.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            cell3 = row3.cells[0]
            
            tc3 = cell3._element
            tcPr3 = tc3.get_or_add_tcPr()
            tcBorders3 = OxmlElement('w:tcBorders')
            bottom3 = OxmlElement('w:bottom')
            bottom3.set(qn('w:val'), 'single')
            bottom3.set(qn('w:sz'), '6')
            bottom3.set(qn('w:color'), '000000')
            tcBorders3.append(bottom3)
            tcPr3.append(tcBorders3)
            
            # 네 번째 선 (하단 실선)
            row4 = table.rows[3]
            row4.height = Pt(10)
            row4.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            cell4 = row4.cells[0]
            
            tc4 = cell4._element
            tcPr4 = tc4.get_or_add_tcPr()
            tcBorders4 = OxmlElement('w:tcBorders')
            bottom4 = OxmlElement('w:bottom')
            bottom4.set(qn('w:val'), 'single')
            bottom4.set(qn('w:sz'), '4')
            bottom4.set(qn('w:color'), '808080')
            tcBorders4.append(bottom4)
            tcPr4.append(tcBorders4)
            
            # 모든 셀의 다른 테두리 제거
            for row in table.rows:
                cell = row.cells[0]
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is not None:
                    for border in ['top', 'left', 'right']:
                        side = tcBorders.find(qn(f'w:{border}'))
                        if side is None:
                            side = OxmlElement(f'w:{border}')
                            side.set(qn('w:val'), 'nil')
                            tcBorders.append(side)
            
            # 줄 간격
            spacing = doc.add_paragraph()
            spacing.paragraph_format.space_after = Pt(10)

def create_cornell_notebook(doc, num_pages=5):
    """코넬노트 양식 생성"""
    for page in range(num_pages):
        if page > 0:
            doc.add_page_break()
        
        # 상단 영역 (제목, 날짜)
        header_table = doc.add_table(rows=1, cols=2)
        header_table.style = 'Table Grid'
        header_table.columns[0].width = Inches(4)
        header_table.columns[1].width = Inches(2.5)
        
        # 제목 셀
        title_cell = header_table.cell(0, 0)
        title_p = title_cell.paragraphs[0]
        title_p.add_run("제목: ").bold = True
        
        # 날짜 셀
        date_cell = header_table.cell(0, 1)
        date_p = date_cell.paragraphs[0]
        date_p.add_run("날짜: ").bold = True
        
        # 간격
        doc.add_paragraph().paragraph_format.space_after = Pt(12)
        
        # 메인 영역 (핵심어 | 노트)
        main_table = doc.add_table(rows=1, cols=2)
        main_table.style = 'Table Grid'
        main_table.columns[0].width = Inches(2)
        main_table.columns[1].width = Inches(4.5)
        
        # 핵심어 열
        key_cell = main_table.cell(0, 0)
        key_p = key_cell.paragraphs[0]
        key_p.add_run("핵심어/질문").bold = True
        key_p.add_run("\n\n")
        
        # 노트 열
        note_cell = main_table.cell(0, 1)
        note_p = note_cell.paragraphs[0]
        note_p.add_run("노트 영역").bold = True
        note_p.add_run("\n\n")
        
        # 셀 높이 설정
        tr = main_table.rows[0]._element
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '8000')  # 약 5.5인치
        trHeight.set(qn('w:hRule'), 'atLeast')
        trPr.append(trHeight)
        
        # 간격
        doc.add_paragraph().paragraph_format.space_after = Pt(12)
        
        # 하단 요약 영역
        summary_title = doc.add_paragraph("요약:")
        summary_title.runs[0].font.bold = True
        summary_title.paragraph_format.space_after = Pt(6)
        
        # 요약 박스
        summary_table = doc.add_table(rows=1, cols=1)
        summary_table.style = 'Table Grid'
        summary_cell = summary_table.cell(0, 0)
        
        # 요약 영역 높이 설정
        tr = summary_table.rows[0]._element
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '2000')  # 약 1.5인치
        trHeight.set(qn('w:hRule'), 'atLeast')
        trPr.append(trHeight)

# Streamlit 앱 설정
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
        lines_per_page = st.slider("페이지당 줄 수", 10, 35, 25)
    elif notebook_type == "칸공책":
        rows = st.slider("행 수", 5, 25, 15)
        cols = st.slider("열 수", 5, 25, 15)
        st.info("💡 팁: 많은 칸을 만들면 생성 시간이 길어질 수 있습니다.")
    elif notebook_type == "영어노트 (4선)":
        lines_per_page = st.slider("페이지당 줄 수", 5, 15, 10)
    
    # 생성 버튼
    if st.button("📄 노트 생성", use_container_width=True, type="primary"):
        with st.spinner("노트를 생성하고 있습니다..."):
            try:
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
                
            except Exception as e:
                st.error(f"❌ 오류가 발생했습니다: {str(e)}")
                st.info("다른 설정으로 다시 시도해보세요.")

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
        - 회색 선으로 구성
        """)
    
    with st.expander("칸공책"):
        st.markdown("""
        - 격자 모양의 칸으로 구성된 노트
        - 수학, 도표, 그래프 그리기에 적합
        - 행과 열의 수 조정 가능
        - 정사각형에 가까운 칸으로 구성
        """)
    
    with st.expander("영어노트 (4선)"):
        st.markdown("""
        - 영어 필기체 연습용 4선 노트
        - 알파벳 쓰기 연습에 최적화
        - 점선과 실선으로 구성
        - 기준선이 굵게 표시됨
        """)
    
    with st.expander("코넬노트"):
        st.markdown("""
        - 효과적인 학습을 위한 노트 양식
        - 핵심어/질문, 노트, 요약 영역으로 구분
        - 체계적인 학습 정리에 적합
        - 복습과 정리가 용이한 구조
        """)
    
    st.subheader("💡 추가 팁")
    st.markdown("""
    - **인쇄 시**: 프린터 설정에서 '실제 크기'로 인쇄하세요.
    - **양면 인쇄**: 용지 절약을 위해 양면 인쇄를 권장합니다.
    - **PDF 변환**: Word 파일을 PDF로 변환하면 레이아웃이 더 안정적입니다.
    - **문제 해결**: 생성이 안 되면 행/열 수를 줄여보세요.
    """)
    
    st.subheader("🔧 문제 해결")
    with st.expander("자주 묻는 질문"):
        st.markdown("""
        **Q: 줄이 보이지 않아요**
        - A: Word 프로그램의 '보기' 설정에서 '격자선' 옵션을 확인하세요.
        
        **Q: 생성이 너무 오래 걸려요**
        - A: 칸공책의 경우 행/열 수를 10x10 정도로 줄여보세요.
        
        **Q: 다운로드가 안 돼요**
        - A: 브라우저 팝업 차단을 해제하거나 다른 브라우저를 사용해보세요.
        """)
