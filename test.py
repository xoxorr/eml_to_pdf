import os
import sys
import email
from email import policy
from reportlab.lib.pagesizes import letter
import fitz  # PyMuPDF (PDF 병합)

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                           QVBoxLayout, QWidget, QFileDialog, QProgressBar)

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from pathlib import Path

# PyQt5 플러그인 경로 설정
if hasattr(sys, 'frozen'):
    # PyInstaller로 패키징된 경우
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(sys._MEIPASS, 'PyQt5', 'Qt5', 'plugins')
    os.environ['QT_PLUGIN_PATH'] = os.path.join(sys._MEIPASS, 'PyQt5', 'Qt5', 'plugins')
else:
    # 개발 환경에서 실행되는 경우
    import PyQt5
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')
    os.environ['QT_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')

# 📌 1. EML 파일에서 메일 데이터 추출
def parse_eml(file_path: str):
    """EML 파일에서 제목, 보낸사람, 본문, 첨부파일을 추출"""
    try:
        # 여러 인코딩 시도
        encodings = ['utf-8', 'cp949', 'euc-kr', 'ascii']
        content = None
        
        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            raise ValueError(f"파일을 읽을 수 없습니다: {file_path}")
            
        msg = email.message_from_string(content, policy=policy.default)

        # 제목 디코딩
        subject = msg["subject"] or "제목 없음"
        if isinstance(subject, email.header.Header):
            subject = str(subject)
        
        # 보낸사람 디코딩
        sender = msg["from"] or "알 수 없음"
        if isinstance(sender, email.header.Header):
            sender = str(sender)
            
        # 날짜 처리
        date = msg["date"] or "날짜 없음"
        
        # 본문 추출 및 디코딩
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        body = part.get_payload(decode=True).decode(charset, errors='replace')
                        break
                    except (UnicodeDecodeError, AttributeError) as e:
                        continue
        else:
            try:
                charset = msg.get_content_charset() or 'utf-8'
                body = msg.get_payload(decode=True).decode(charset, errors='replace')
            except (UnicodeDecodeError, AttributeError) as e:
                body = msg.get_payload(decode=False)

        # 첨부파일 처리
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if filename:
                    # 첨부파일 디코딩 및 저장
                    try:
                        if isinstance(filename, email.header.Header):
                            filename = str(filename)
                        
                        # 첨부파일 정보 저장
                        attachments.append({
                            "filename": filename,
                            "content": part.get_payload(decode=True),
                            "content_type": part.get_content_type()
                        })
                    except Exception as e:
                        print(f"첨부파일 처리 중 오류 ({filename}): {str(e)}")

        return {
            "제목": subject,
            "보낸사람": sender,
            "날짜": date,
            "내용": body.strip(),
            "첨부파일": attachments
        }
        
    except Exception as e:
        raise Exception(f"EML 파일 처리 중 오류 발생 ({file_path}): {str(e)}")

# 📌 2. EML 데이터를 PDF로 변환 (ReportLab 사용)
def save_to_pdf(eml_data, filename, attachments_dir=None):
    """추출한 EML 데이터를 PDF로 저장 (검색 가능한 형식)"""
    try:
        doc = SimpleDocTemplate(
            filename,
            pagesize=letter,
            leftMargin=50,
            rightMargin=50,
            topMargin=50,
            bottomMargin=50
        )
        
        # 스타일 설정
        styles = getSampleStyleSheet()
        
        # 한글 폰트 설정
        try:
            fonts = [
                ('맑은고딕', 'malgun.ttf'),
                ('굴림', 'gulim.ttc'),
                ('돋움', 'dotum.ttc')
            ]
            
            for font_name, font_file in fonts:
                try:
                    font_path = f'C:/Windows/Fonts/{font_file}'
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                    font_to_use = font_name
                    break
                except (OSError, IOError) as e:
                    continue
        except ImportError as e:
            font_to_use = "Helvetica"

        # 사용자 정의 스타일
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontName=font_to_use,
            fontSize=14,
            spaceAfter=20
        )
        
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontName=font_to_use,
            fontSize=12,
            spaceAfter=20
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['Normal'],
            fontName=font_to_use,
            fontSize=10,
            leading=15,
            spaceBefore=10,
            spaceAfter=10
        )

        # 문서 구성요소 생성
        story = []
        
        # 제목
        title = Paragraph(f"제목: {eml_data['제목']}", title_style)
        story.append(title)
        
        # 헤더 정보
        header = Paragraph(f"보낸사람: {eml_data['보낸사람']} | 날짜: {eml_data['날짜']}", header_style)
        story.append(header)
        
        # 본문 처리
        paragraphs = eml_data["내용"].split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                p = Paragraph(paragraph.replace('\n', '<br/>'), body_style)
                story.append(p)

        # 첨부파일 목록 추가
        if eml_data["첨부파일"]:
            story.append(Paragraph("<br/><br/>📎 첨부파일:", header_style))
            
            # 첨부파일 저장 및 목록 작성
            if attachments_dir:
                os.makedirs(attachments_dir, exist_ok=True)
                for attachment in eml_data["첨부파일"]:
                    filename = attachment["filename"]
                    save_path = os.path.join(attachments_dir, filename)
                    
                    # 첨부파일 저장
                    with open(save_path, 'wb') as f:
                        f.write(attachment["content"])
                    
                    # PDF에 첨부파일 목록 추가
                    story.append(Paragraph(f"- {filename}", body_style))

        # PDF 생성
        doc.build(story)
        
    except Exception as e:
        raise Exception(f"PDF 저장 중 오류 발생 ({filename}): {str(e)}")

# 📌 3. 폴더 내 모든 EML 파일을 PDF로 변환
def convert_eml_folder_to_pdfs(eml_folder, output_folder):
    """폴더 내 모든 EML 파일을 개별 PDF로 변환"""
    os.makedirs(output_folder, exist_ok=True)

    for file in os.listdir(eml_folder):
        if file.endswith(".eml"):
            try:
                eml_path = os.path.join(eml_folder, file)
                eml_data = parse_eml(eml_path)
                
                # 보낸사람 이메일 주소 추출
                sender_email = eml_data['보낸사람']
                if '<' in sender_email and '>' in sender_email:
                    sender_email = sender_email.split('<')[1].split('>')[0]
                
                # 날짜 처리
                from email.utils import parsedate_to_datetime
                try:
                    date_obj = parsedate_to_datetime(eml_data['날짜'])
                    date_str = date_obj.strftime('%Y%m%d_%H%M')
                except:
                    date_str = "날짜없음"
                
                # 파일명으로 사용할 수 없는 문자 제거
                safe_filename = "".join(c for c in sender_email if c.isalnum() or c in ('._-@'))
                
                # 날짜를 포함한 파일명 생성
                base_filename = f"{date_str}_{safe_filename}"
                
                # PDF 파일 경로
                pdf_path = os.path.join(output_folder, f"{base_filename}.pdf")
                
                # 동일 파일명이 있을 경우 번호 추가
                counter = 1
                while os.path.exists(pdf_path):
                    pdf_path = os.path.join(output_folder, f"{base_filename}_{counter}.pdf")
                    counter += 1
                
                # 첨부파일 폴더 경로
                attachments_dir = os.path.join(output_folder, f"{base_filename}_attachments")
                if counter > 1:
                    attachments_dir = os.path.join(output_folder, f"{base_filename}_{counter-1}_attachments")
                
                save_to_pdf(eml_data, pdf_path, attachments_dir)
                
            except Exception as e:
                print(f"파일 처리 중 오류 발생 ({file}): {str(e)}")
                continue

# 📌 4. 변환된 PDF 파일 하나로 합치기
def merge_pdfs(input_folder, output_filename):
    pdf_writer = fitz.open()
    
    for file in sorted(os.listdir(input_folder)):
        if file.endswith(".pdf"):
            pdf_reader = fitz.open(os.path.join(input_folder, file))
            pdf_writer.insert_pdf(pdf_reader)
    
    pdf_writer.save(output_filename)
    pdf_writer.close()

# 📌 실행 함수
def process_eml_to_pdf(eml_folder, output_folder, merged_pdf_name):
    """전체 실행 파이프라인"""
    print("📂 EML → PDF 변환 중...")
    convert_eml_folder_to_pdfs(eml_folder, output_folder)

    print("📑 PDF 합치는 중...")
    merge_pdfs(output_folder, merged_pdf_name)

    print(f"✅ 변환 완료: {merged_pdf_name}")

# 📌 사용 예시
# eml_folder = "eml_백업폴더"
# output_folder = "pdf_출력폴더"
# merged_pdf_name = "통합_메일.pdf"
# process_eml_to_pdf(eml_folder, output_folder, merged_pdf_name)

class EmlConverterUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
        # 선택된 경로 저장용 변수
        self.eml_folder = ""
        self.output_folder = ""
        self.merged_pdf_name = "통합_메일.pdf"

    def initUI(self):
        """UI 초기화"""
        self.setWindowTitle('EML to PDF 변환기')
        self.setGeometry(100, 100, 600, 400)

        # 메인 위젯 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()

        # EML 폴더 선택
        self.eml_label = QLabel('EML 폴더: 선택되지 않음')
        eml_btn = QPushButton('EML 폴더 선택')
        eml_btn.clicked.connect(self.select_eml_folder)

        # 출력 폴더 선택
        self.output_label = QLabel('출력 폴더: 선택되지 않음')
        output_btn = QPushButton('출력 폴더 선택')
        output_btn.clicked.connect(self.select_output_folder)

        # 진행 상태 표시
        self.progress = QProgressBar()
        
        # 변환 시작 버튼
        convert_btn = QPushButton('변환 시작')
        convert_btn.clicked.connect(self.start_conversion)

        # 상태 메시지
        self.status_label = QLabel('대기 중...')
        self.status_label.setAlignment(Qt.AlignCenter)

        # 레이아웃에 위젯 추가
        layout.addWidget(self.eml_label)
        layout.addWidget(eml_btn)
        layout.addWidget(self.output_label)
        layout.addWidget(output_btn)
        layout.addWidget(self.progress)
        layout.addWidget(convert_btn)
        layout.addWidget(self.status_label)

        main_widget.setLayout(layout)

    def select_eml_folder(self):
        """EML 폴더 선택"""
        folder = QFileDialog.getExistingDirectory(self, 'EML 폴더 선택')
        if folder:
            self.eml_folder = folder
            self.eml_label.setText(f'EML 폴더: {folder}')

    def select_output_folder(self):
        """출력 폴더 선택"""
        folder = QFileDialog.getExistingDirectory(self, '출력 폴더 선택')
        if folder:
            self.output_folder = folder
            self.output_label.setText(f'출력 폴더: {folder}')

    def start_conversion(self):
        """변환 시작"""
        if not self.eml_folder or not self.output_folder:
            self.status_label.setText('❌ 폴더를 모두 선택해주세요!')
            return

        try:
            # 진행 상태 초기화
            self.status_label.setText('📂 EML → PDF 변환 중...')
            self.progress.setValue(0)
            
            # EML 파일 변환
            convert_eml_folder_to_pdfs(self.eml_folder, self.output_folder)
            self.progress.setValue(50)
            
            # PDF 병합
            self.status_label.setText('📑 PDF 파일 병합 중...')
            merged_pdf_path = os.path.join(self.output_folder, self.merged_pdf_name)
            merge_pdfs(self.output_folder, merged_pdf_path)
            
            self.progress.setValue(100)
            self.status_label.setText('✅ 변환 완료!')
            
        except Exception as e:
            self.status_label.setText(f'❌ 오류 발생: {str(e)}')

# 프로그램 실행
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EmlConverterUI()
    ex.show()
    sys.exit(app.exec_())
