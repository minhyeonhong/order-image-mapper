import sys
import zipfile
import tempfile
import os
import json

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QFileDialog,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QGroupBox,
    QTextEdit,
    QLineEdit,
    QComboBox,
    QInputDialog,
    QMessageBox,
    QTabWidget,
    QListWidget,
    QListWidgetItem,
)

from excel_handler import ExcelHandler
from image_matcher import ImageMatcher
from image_inserter import ImageInserter
from datetime import datetime


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setAcceptDrops(True)

        self.recipe_file = 'recipes.json'
        self.recipes = self.load_recipes()

        self.excel_path = ''
        self.zip_path = ''

        self.setWindowTitle('수주서 이미지 자동 삽입기')
        self.resize(700, 650)
        self.setStyleSheet("""
            QWidget { font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; font-size: 13px; color: #222; }
            QGroupBox { font-weight: bold; border: 1px solid #ccc; border-radius: 6px; margin-top: 15px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; color: #0056b3; }
            QPushButton { background-color: #f8f9fa; border: 1px solid #ccc; border-radius: 4px; padding: 6px 12px; }
            QPushButton:hover { background-color: #e2e6ea; }
            QPushButton#runBtn { background-color: #007bff; color: white; font-weight: bold; font-size: 14px; padding: 10px; margin-top: 10px; }
            QPushButton#runBtn:hover { background-color: #0069d9; }
            QLineEdit, QComboBox, QListWidget { border: 1px solid #ccc; border-radius: 4px; padding: 5px; background: white; }
            QLineEdit:focus, QComboBox:focus, QListWidget:focus { border: 1px solid #80bdff; }
            QTextEdit { border: 1px solid #ccc; border-radius: 4px; background: #fff; }
            QTabBar::tab { padding: 8px 15px; border: 1px solid #ccc; border-bottom: none; border-top-left-radius: 4px; border-top-right-radius: 4px; background: #f0f0f0; }
            QTabBar::tab:selected { background: #fff; font-weight: bold; }
            QTabWidget::pane { border: 1px solid #ccc; border-radius: 4px; }
            QLabel#fileLabel { border: 2px dashed #bbb; border-radius: 4px; padding: 10px; background: #fdfdfd; color: #555; font-weight: bold; }
        """)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # --- 1. 파일 선택 그룹 ---
        file_group = QGroupBox("📂 파일 선택 (드래그 앤 드롭)")
        file_layout = QGridLayout()
        file_layout.setSpacing(10)

        self.excel_label = QLabel('Excel 미선택')
        self.excel_label.setObjectName("fileLabel")
        self.zip_label = QLabel('ZIP 미선택')
        self.zip_label.setObjectName("fileLabel")

        excel_btn = QPushButton('Excel 찾기')
        excel_btn.clicked.connect(self.select_excel)
        zip_btn = QPushButton('ZIP 찾기')
        zip_btn.clicked.connect(self.select_zip)

        file_layout.addWidget(QLabel('<b>Excel 파일</b>'), 0, 0)
        file_layout.addWidget(self.excel_label, 0, 1)
        file_layout.addWidget(excel_btn, 0, 2)

        file_layout.addWidget(QLabel('<b>ZIP 파일</b>'), 1, 0)
        file_layout.addWidget(self.zip_label, 1, 1)
        file_layout.addWidget(zip_btn, 1, 2)
        
        file_layout.setColumnStretch(1, 1)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # --- 2. 설정 그룹 ---
        settings_group = QGroupBox("⚙️ 상세 설정")
        settings_layout = QVBoxLayout()
        settings_layout.setSpacing(15)

        self.current_recipe_name = "기본 설정"
        self.current_recipe_label = QLabel(f"<b>현재 적용된 레시피:</b> {self.current_recipe_name}")
        settings_layout.addWidget(self.current_recipe_label)

        # 입력 필드 UI (Grid)
        form_layout = QGridLayout()
        form_layout.setSpacing(10)

        self.start_row_input = QLineEdit('9')
        self.jan_col_input = QLineEdit('F')
        self.name_col_input = QLineEdit('H')
        self.image_col_input = QLineEdit('D')
        self.image_width_input = QLineEdit('150')
        self.image_height_input = QLineEdit('150')

        form_layout.addWidget(QLabel('데이터 시작 행:'), 0, 0)
        form_layout.addWidget(self.start_row_input, 0, 1)
        
        form_layout.addWidget(QLabel('JAN 열:'), 0, 2)
        form_layout.addWidget(self.jan_col_input, 0, 3)

        form_layout.addWidget(QLabel('상품명 열:'), 1, 0)
        form_layout.addWidget(self.name_col_input, 1, 1)

        form_layout.addWidget(QLabel('이미지 삽입 열:'), 1, 2)
        form_layout.addWidget(self.image_col_input, 1, 3)

        form_layout.addWidget(QLabel('이미지 너비:'), 2, 0)
        form_layout.addWidget(self.image_width_input, 2, 1)

        form_layout.addWidget(QLabel('이미지 높이:'), 2, 2)
        form_layout.addWidget(self.image_height_input, 2, 3)

        settings_layout.addLayout(form_layout)
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)

        # --- 3. 실행 버튼 ---
        run_btn = QPushButton('▶ 이미지 자동 삽입 실행')
        run_btn.setObjectName("runBtn")
        run_btn.clicked.connect(self.run_process)
        main_layout.addWidget(run_btn)

        # --- 4. 탭 위젯 (로그 & 레시피 관리) ---
        self.tab_widget = QTabWidget()
        
        # 탭 1: 진행 로그
        self.log_tab = QWidget()
        log_layout = QVBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(200)
        log_layout.addWidget(self.log_box)
        self.log_tab.setLayout(log_layout)
        
        # 탭 2: 레시피 관리
        self.recipe_tab = QWidget()
        recipe_tab_layout = QVBoxLayout()
        
        self.recipe_list = QListWidget()
        self.recipe_list.addItem("기본 설정")
        self.recipe_list.addItems(list(self.recipes.keys()))
        self.recipe_list.setCurrentRow(0)
        
        recipe_btn_layout = QHBoxLayout()
        apply_recipe_btn = QPushButton("적용")
        apply_recipe_btn.clicked.connect(self.apply_recipe)
        save_recipe_tab_btn = QPushButton("현재 설정 저장")
        save_recipe_tab_btn.clicked.connect(self.save_recipe)
        delete_recipe_tab_btn = QPushButton("삭제")
        delete_recipe_tab_btn.clicked.connect(self.delete_recipe)
        
        recipe_btn_layout.addWidget(apply_recipe_btn)
        recipe_btn_layout.addWidget(save_recipe_tab_btn)
        recipe_btn_layout.addWidget(delete_recipe_tab_btn)
        
        recipe_tab_layout.addWidget(self.recipe_list)
        recipe_tab_layout.addLayout(recipe_btn_layout)
        self.recipe_tab.setLayout(recipe_tab_layout)

        self.tab_widget.addTab(self.log_tab, "📝 진행 로그")
        self.tab_widget.addTab(self.recipe_tab, "⚙️ 레시피 관리")

        main_layout.addWidget(self.tab_widget)

        self.setLayout(main_layout)

    def load_recipes(self):
        if os.path.exists(self.recipe_file):
            try:
                with open(self.recipe_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_recipes_to_file(self):
        with open(self.recipe_file, 'w', encoding='utf-8') as f:
            json.dump(self.recipes, f, ensure_ascii=False, indent=4)

    def save_recipe(self):
        name, ok = QInputDialog.getText(self, "레시피 저장", "레시피 이름을 입력하세요:")
        if ok and name:
            name = name.strip()
            if not name:
                return
            if name == "기본 설정":
                QMessageBox.warning(self, "경고", "'기본 설정'이라는 이름은 사용할 수 없습니다.")
                return

            if name not in self.recipes:
                self.recipe_list.addItem(name)

            self.recipes[name] = {
                'start_row': self.start_row_input.text(),
                'jan_col': self.jan_col_input.text(),
                'name_col': self.name_col_input.text(),
                'image_col': self.image_col_input.text(),
                'image_width': self.image_width_input.text(),
                'image_height': self.image_height_input.text()
            }
            self.save_recipes_to_file()
            
            self.current_recipe_name = name
            self.current_recipe_label.setText(f"<b>현재 적용된 레시피:</b> {name}")
            self.log(f'레시피 "{name}"이(가) 저장되었습니다.')
            self.tab_widget.setCurrentIndex(0)

    def delete_recipe(self):
        current_item = self.recipe_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "경고", "삭제할 레시피를 선택해주세요.")
            return
            
        name = current_item.text()
        if name == "기본 설정":
            QMessageBox.warning(self, "경고", "기본 설정은 삭제할 수 없습니다.")
            return

        reply = QMessageBox.question(self, "확인", f'레시피 "{name}"을(를) 삭제하시겠습니까?', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            if name in self.recipes:
                del self.recipes[name]
                self.save_recipes_to_file()
                
                row = self.recipe_list.row(current_item)
                self.recipe_list.takeItem(row)
                
                if self.current_recipe_name == name:
                    self.current_recipe_name = "기본 설정"
                    self.current_recipe_label.setText("<b>현재 적용된 레시피:</b> 기본 설정")

                self.log(f'레시피 "{name}"이(가) 삭제되었습니다.')
                self.tab_widget.setCurrentIndex(0)

    def apply_recipe(self):
        current_item = self.recipe_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "경고", "적용할 레시피를 선택해주세요.")
            return

        name = current_item.text()

        if name == "기본 설정":
            self.start_row_input.setText('9')
            self.jan_col_input.setText('F')
            self.name_col_input.setText('H')
            self.image_col_input.setText('D')
            self.image_width_input.setText('150')
            self.image_height_input.setText('150')
        elif name in self.recipes:
            recipe = self.recipes[name]
            self.start_row_input.setText(recipe.get('start_row', '9'))
            self.jan_col_input.setText(recipe.get('jan_col', 'F'))
            self.name_col_input.setText(recipe.get('name_col', 'H'))
            self.image_col_input.setText(recipe.get('image_col', 'D'))
            self.image_width_input.setText(recipe.get('image_width', '150'))
            self.image_height_input.setText(recipe.get('image_height', '150'))
            
        self.current_recipe_name = name
        self.current_recipe_label.setText(f"<b>현재 적용된 레시피:</b> {name}")
        self.log(f'레시피 "{name}"이(가) 적용되었습니다.')
        self.tab_widget.setCurrentIndex(0)

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
    