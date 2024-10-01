from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn


# PlaceholderReplacer 클래스 정의 (앞서 설명한 클래스)
class PlaceholderReplacer:
    def __init__(self, doc, placeholder_values, font_name=None, font_size=None, font_color=None):
        self.doc = doc
        self.placeholder_values = placeholder_values
        self.font_name = font_name
        self.font_size = font_size
        self.font_color = font_color
        self.replaced_placeholders = {}

    def replace_placeholder(self, paragraph, placeholder, value, context):
        if placeholder in paragraph.text:
            full_text = ""
            for run in paragraph.runs:
                full_text += run.text
            new_text = full_text.replace(placeholder, value)

            # clear all runs and add the new text in a single run
            paragraph.clear()
            run = paragraph.add_run(new_text)

            # 설정된 폰트 값이 있을 경우만 적용
            if self.font_name:
                run.font.name = self.font_name
                r = run._element
                r.rPr.rFonts.set(qn('w:eastAsia'), self.font_name)
            if self.font_size:
                run.font.size = Pt(self.font_size)
            if self.font_color:
                run.font.color.rgb = RGBColor(*self.font_color)

            # 주석처리된 로그 (필요 시 사용)
            # print(f"{context}에서 {placeholder} 대체 성공")
            return True
        return False

    def process_paragraphs(self):
        for para in self.doc.paragraphs:
            for placeholder, value in self.placeholder_values.items():
                if placeholder not in self.replaced_placeholders:
                    if self.replace_placeholder(para, placeholder, value, context="문단"):
                        self.replaced_placeholders[placeholder] = "문단"

    def process_tables(self):
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for placeholder, value in self.placeholder_values.items():
                            if placeholder not in self.replaced_placeholders:
                                if self.replace_placeholder(para, placeholder, value, context="테이블"):
                                    self.replaced_placeholders[placeholder] = "테이블"

    def check_replacement_status(self):
        success_count = 0
        fail_count = 0
        for placeholder in self.placeholder_values.keys():
            if placeholder in self.replaced_placeholders:
                location = self.replaced_placeholders[placeholder]
                # 주석처리된 로그 (필요 시 사용)
                # print(f"{placeholder}: {location}에서 성공")
                success_count += 1
            else:
                # 주석처리된 로그 (필요 시 사용)
                # print(f"{placeholder}: 실패")
                fail_count += 1
        # 주석처리된 로그 (필요 시 사용)
        # print(f"\n총 성공한 placeholder: {success_count}, 총 실패한 placeholder: {fail_count}")


# 실제 사용 예시

# 1. 문서 열기
doc = Document('labeled-PCI-DSS-v4-0-SAQ-A-r1.docx')

# 2. placeholder와 교체할 값 설정
placeholder_values = {
    '{company_name}': 'marketian',
    '{dba}': 'hello, world',
    '{appendix_c1_title}': 'Requirements 3.5.2',
    '{appendix_c1_content}': 'Lorem ipsum, duaos einenskd enaoldfj',
    '{appendix_c2_title}': 'Requirements 3.5.2',
    '{appendix_c2_content}': 'Suspendisse ut purus sed quam consectetur cursus. Pellentesque eget metus tristique, tincidunt elit vel, posuere nulla. Aenean maximus purus eget mi consectetur, ac ultrices velit laoreet. Aliquam nulla urna, fermentum tempus interdum non, faucibus.',
}

# 3. PlaceholderReplacer 객체 생성
replacer = PlaceholderReplacer(doc, placeholder_values)

# 4. 문단과 테이블 내에서 placeholder 대체 수행
replacer.process_paragraphs()
replacer.process_tables()

# 5. 대체 성공 여부 출력
replacer.check_replacement_status()

# 6. 변경된 문서 저장
doc.save('output1.docx')

# 주석처리된 로그를 사용하고 싶다면, 클래스 내 주석을 해제하세요.
