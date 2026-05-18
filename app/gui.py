import sys
import zipfile
import tempfile
import os

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QFileDialog,
    QLabel,
    QVBoxLayout,
    QTextEdit,
    QLineEdit,
)

from excel_handler import ExcelHandler
from image_matcher import ImageMatcher
from image_inserter import ImageInserter
from datetime import datetime


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setAcceptDrops(True)

        self.excel_path = ''
        self.zip_path = ''

        self.setWindowTitle('수주서 이미지 자동 삽입기')

        layout = QVBoxLayout()

        self.excel_label = QLabel('Excel 미선택 (드래그 앤 드롭 가능)')
        self.zip_label = QLabel('ZIP 미선택 (드래그 앤 드롭 가능)')

        # 입력받을 데이터 위한 입력창
        self.start_row_input = QLineEdit('9')
        self.jan_col_input = QLineEdit('F')
        self.name_col_input = QLineEdit('H')
        self.image_col_input = QLineEdit('D')

        self.image_width_input = QLineEdit('150')
        self.image_height_input = QLineEdit('150')

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(300)

        excel_btn = QPushButton('Excel 선택')
        excel_btn.clicked.connect(self.select_excel)

        zip_btn = QPushButton('ZIP 선택')
        zip_btn.clicked.connect(self.select_zip)

        run_btn = QPushButton('실행')
        run_btn.clicked.connect(self.run_process)

        layout.addWidget(QLabel('Excel 파일'))
        layout.addWidget(self.excel_label)
        layout.addWidget(excel_btn)

        layout.addWidget(QLabel('ZIP 파일'))
        layout.addWidget(self.zip_label)
        layout.addWidget(zip_btn)

        layout.addWidget(QLabel('시작 행'))
        layout.addWidget(self.start_row_input)

        layout.addWidget(QLabel('JAN 열'))
        layout.addWidget(self.jan_col_input)

        layout.addWidget(QLabel('상품명 열'))
        layout.addWidget(self.name_col_input)

        layout.addWidget(QLabel('이미지 열'))
        layout.addWidget(self.image_col_input)

        layout.addWidget(QLabel('이미지 너비'))
        layout.addWidget(self.image_width_input)

        layout.addWidget(QLabel('이미지 높이'))
        layout.addWidget(self.image_height_input)

        layout.addWidget(run_btn)

        layout.addWidget(QLabel('로그'))
        layout.addWidget(self.log_box)

        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file_path in files:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in ['.xls', '.xlsx']:
                self.excel_path = file_path
                self.excel_label.setText(file_path)
            elif ext == '.zip':
                self.zip_path = file_path
                self.zip_label.setText(file_path)

    def select_excel(self):
        path, _ = QFileDialog.getOpenFileName(self)

        if path:
            self.excel_path = path
            self.excel_label.setText(path)

    def select_zip(self):
        path, _ = QFileDialog.getOpenFileName(self)

        if path:
            self.zip_path = path
            self.zip_label.setText(path)

    def log(self, message):
        self.log_box.append(str(message))
        self.log_box.ensureCursorVisible()

    def run_process(self):
        try:
            self.log_box.clear()
            
            temp_dir = tempfile.mkdtemp()

            # ZIP 압축 해제
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                for zip_info in zip_ref.infolist():

                    try:
                        filename = zip_info.filename.encode('cp437').decode('cp932')
                    except:
                        filename = zip_info.filename

                    filename = os.path.basename(filename)

                    if not filename:
                        continue

                    if filename.startswith('._') or filename.startswith('.'):
                        continue

                    target_path = os.path.join(temp_dir, filename)

                    if zip_info.is_dir():
                        continue

                    with zip_ref.open(zip_info) as source, open(target_path, 'wb') as target:
                        target.write(source.read())

            # Excel 읽기
            excel = ExcelHandler(self.excel_path)

            products = excel.read_products(
                int(self.start_row_input.text()),
                self.jan_col_input.text(),
                self.name_col_input.text(),
            )

            # 이미지 매칭
            matcher = ImageMatcher(temp_dir)
            
            inserter = ImageInserter(excel.ws)

            failed = []

            inserted = 0

            total = len(products)

            for index, product in enumerate(products, start=1):

                if index % 10 == 0 or index == total:
                    self.log(
                        f'진행률: {index}/{total} | 성공:{inserted} 실패:{len(failed)}'
                    )

                QApplication.processEvents()

                image = matcher.find_image(
                    product['jan'],
                    product['name']
                )

                if image:
                    image_col = self.image_col_input.text()

                    inserter.insert_image(
                        temp_dir,
                        image,
                        f"{image_col}{product['row']}",
                        int(self.image_width_input.text()),
                        int(self.image_height_input.text()),
                    )

                    inserted += 1

                else:
                    failed.append(product)

            # 저장
            # 1. 현재 날짜와 시간을 원하는 형식으로 가져오기
            # %m(월), %d(일), %H(시), %M(분)
            # 윈도우 호환을 위해 콜론(:) 대신 하이픈(-)을 사용한 예시입니다.
            now_str = datetime.now().strftime("%m-%d_%H-%M") 

            # 2. 파일명 구성 (예: output_05-15_23-12.xlsx)
            filename = f"output_{now_str}.xlsx"

            # 3. 전체 경로 생성
            output_path = os.path.join(
                os.path.dirname(self.excel_path),
                filename
            )

            excel.save(output_path)

            self.log(
                f'완료! 삽입:{inserted} 실패:{len(failed)}'
            )

            self.log('========== 완료 ==========')
            self.log(f'총 상품 수: {len(products)}')
            self.log(f'삽입 성공: {inserted}')
            self.log(f'매칭 실패: {len(failed)}')
            self.log(f'결과 파일: {output_path}')

            if failed:

                from openpyxl import Workbook
                fail_wb = Workbook()
                fail_ws = fail_wb.active
                fail_ws.title = '매칭실패'

                fail_ws['A1'] = '원본 엑셀 파일명'
                fail_ws['B1'] = '행 번호'
                fail_ws['C1'] = 'JAN'
                fail_ws['D1'] = '상품명'

                excel_basename = os.path.basename(self.excel_path)
                fail_filename = f"output_fail_{now_str}.xlsx"
                fail_output_path = os.path.join(
                    os.path.dirname(self.excel_path),
                    fail_filename
                )

                for idx, item in enumerate(failed, start=2):
                    fail_ws[f'A{idx}'] = excel_basename
                    fail_ws[f'B{idx}'] = item['row']
                    fail_ws[f'C{idx}'] = item['jan']
                    fail_ws[f'D{idx}'] = item['name']

                fail_wb.save(fail_output_path)
                self.log(f'실패 목록 파일: {fail_output_path}')

        except Exception as e:
            self.log('에러 발생:')
            self.log(e)

            self.log(f'에러: {str(e)}')
            
def run_app():
    app = QApplication(sys.argv)

    window = MainWindow()
    window.resize(700, 500)
    window.show()

    sys.exit(app.exec())
    