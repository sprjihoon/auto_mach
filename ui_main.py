"""
PySide6 UI 화면
"""
import sys
from pathlib import Path
from typing import Optional
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QTextEdit, QPushButton,
    QLabel, QLineEdit, QFileDialog, QGroupBox, QSplitter,
    QHeaderView, QMessageBox, QFrame, QCheckBox, QDialog,
    QScrollArea, QGridLayout
)
from PySide6.QtCore import Qt, Slot, QTimer
from PySide6.QtGui import QFont, QColor, QPalette, QIcon
import pandas as pd

from models import ScanResult, ScanEvent
from excel_loader import ExcelLoader


class SummaryDialog(QDialog):
    """구성 요약 다이얼로그 (카드 형태)"""
    
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("📦 구성 요약")
        self.setMinimumSize(800, 600)
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 헤더
        header = QLabel()
        pending = self.df[self.df['used'] == 0]
        total = len(self.df['tracking_no'].unique())
        pending_count = len(pending['tracking_no'].unique())
        header.setText(f"<h2>📦 총 {total}건 중 미처리 {pending_count}건</h2>")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)
        
        # 스크롤 영역
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        # 카드 컨테이너
        container = QWidget()
        grid = QGridLayout(container)
        grid.setSpacing(15)
        
        # 구성별 카드 생성
        combo_data = self._get_combo_data(pending)
        
        row, col = 0, 0
        max_cols = 3
        
        for combo_info in combo_data:
            card = self._create_card(combo_info)
            grid.addWidget(card, row, col)
            col += 1
            if col >= max_cols:
                col = 0
                row += 1
        
        scroll.setWidget(container)
        layout.addWidget(scroll)
        
        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.close)
        close_btn.setMaximumWidth(200)
        layout.addWidget(close_btn, alignment=Qt.AlignCenter)
    
    def _get_combo_data(self, pending):
        """구성별 데이터 추출 (수량 포함)"""
        tracking_groups = pending.groupby('tracking_no')
        combo_counts = {}
        
        for tracking_no, group in tracking_groups:
            barcodes = sorted(group['barcode'].unique())
            combo_key = tuple(barcodes)
            
            if combo_key not in combo_counts:
                combo_counts[combo_key] = {
                    'count': 0,
                    'products': [],
                    'barcodes': barcodes
                }
                for _, row in group.iterrows():
                    product_name = str(row['product_name']) if pd.notna(row['product_name']) else ''
                    option_name = str(row['option_name']) if pd.notna(row['option_name']) else ''
                    qty = int(row['qty']) if pd.notna(row['qty']) else 1
                    
                    product_info = product_name
                    if option_name and option_name != 'nan':
                        product_info += f" ({option_name})"
                    
                    # 수량 뒤에 표시: 1개, 2개, 3개...
                    product_info += f" {qty}개"
                    
                    if product_info and product_info not in combo_counts[combo_key]['products']:
                        combo_counts[combo_key]['products'].append(product_info)
            
            combo_counts[combo_key]['count'] += 1
        
        # 개수 내림차순 정렬
        sorted_combos = sorted(combo_counts.values(), key=lambda x: -x['count'])
        return sorted_combos
    
    def _create_card(self, combo_info):
        """카드 위젯 생성 (전체 품목 가로 나열)"""
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        card.setMinimumWidth(230)
        card.setMaximumWidth(350)
        
        # 개수에 따른 색상
        count = combo_info['count']
        if count >= 10:
            bg_color = "#FFEBEE"  # 빨강 계열
            border_color = "#EF5350"
            count_color = "#D32F2F"
        elif count >= 5:
            bg_color = "#FFF3E0"  # 주황 계열
            border_color = "#FF9800"
            count_color = "#E65100"
        elif count >= 3:
            bg_color = "#E3F2FD"  # 파랑 계열
            border_color = "#2196F3"
            count_color = "#1565C0"
        else:
            bg_color = "#F5F5F5"  # 회색 계열
            border_color = "#9E9E9E"
            count_color = "#616161"
        
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 2px solid {border_color};
                border-radius: 10px;
                padding: 10px;
            }}
        """)
        
        layout = QVBoxLayout(card)
        layout.setSpacing(8)
        
        # 개수 배지 (3자리 지원)
        count_label = QLabel(f"<span style='font-size:24px; font-weight:bold; color:{count_color};'>{count}건</span>")
        count_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(count_label)
        
        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(f"background-color: {border_color};")
        layout.addWidget(line)
        
        # 상품 목록 (◆ 구분자로 명확히 구분)
        products = combo_info['products']
        products_text = "  ◆  ".join(products)
        
        prod_label = QLabel(products_text)
        prod_label.setWordWrap(True)
        prod_label.setStyleSheet("font-size: 11px; color: #333; line-height: 1.4;")
        layout.addWidget(prod_label)
        
        layout.addStretch()
        
        return card
from scanner_listener import ScannerListener
from ezauto_input import EzAutoInput
from pdf_printer import PDFPrinter
from order_processor import OrderProcessor
from utils import get_timestamp


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        
        # 모듈 초기화
        self.excel_loader = ExcelLoader()
        self.scanner = ScannerListener()
        self.ezauto = EzAutoInput()
        self.pdf_printer = PDFPrinter()
        self.processor = OrderProcessor(
            self.excel_loader,
            self.ezauto,
            self.pdf_printer
        )
        
        # UI 초기화
        self._init_ui()
        self._connect_signals()
        
        # 스캐너 시작
        self._scanner_active = False
    
    def _init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("자동출고 프로그램 v1.0")
        self.setMinimumSize(1200, 800)
        
        # 중앙 위젯
        central = QWidget()
        self.setCentralWidget(central)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # === 상단: 파일 로드 및 설정 ===
        top_group = self._create_top_section()
        main_layout.addWidget(top_group)
        
        # === 중간: 스플리터 (테이블들 + 로그) ===
        splitter = QSplitter(Qt.Vertical)
        
        # 테이블 영역
        tables_widget = self._create_tables_section()
        splitter.addWidget(tables_widget)
        
        # 로그 영역
        log_group = self._create_log_section()
        splitter.addWidget(log_group)
        
        splitter.setSizes([500, 200])
        main_layout.addWidget(splitter, 1)
        
        # === 하단: 상태바 ===
        self._create_status_bar()
        
        # 스타일 적용
        self._apply_styles()
    
    def _create_top_section(self) -> QGroupBox:
        """상단 섹션: 파일 로드 및 설정"""
        group = QGroupBox("설정")
        layout = QHBoxLayout(group)
        
        # 엑셀 파일 경로
        layout.addWidget(QLabel("엑셀:"))
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("엑셀 파일 선택")
        self.excel_path_edit.setMaximumWidth(150)
        layout.addWidget(self.excel_path_edit)
        
        # 찾아보기 버튼
        self.browse_btn = QPushButton("찾아보기")
        self.browse_btn.clicked.connect(self._on_browse_excel)
        layout.addWidget(self.browse_btn)
        
        # 로드 버튼
        self.load_btn = QPushButton("불러오기")
        self.load_btn.clicked.connect(self._on_load_excel)
        layout.addWidget(self.load_btn)
        
        # 구성 요약 버튼
        self.summary_btn = QPushButton("📦 구성요약")
        self.summary_btn.clicked.connect(self._on_show_summary)
        layout.addWidget(self.summary_btn)
        
        layout.addSpacing(20)
        
        # PDF 파일 경로
        layout.addWidget(QLabel("PDF:"))
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("PDF 선택")
        self.pdf_path_edit.setMaximumWidth(150)
        layout.addWidget(self.pdf_path_edit)
        
        # PDF 파일 찾아보기 버튼
        self.pdf_browse_btn = QPushButton("파일 선택")
        self.pdf_browse_btn.clicked.connect(self._on_browse_pdf_file)
        layout.addWidget(self.pdf_browse_btn)
        
        layout.addSpacing(20)
        
        # 스캐너 시작/중지
        self.scanner_btn = QPushButton("스캐너 시작")
        self.scanner_btn.setCheckable(True)
        self.scanner_btn.clicked.connect(self._on_toggle_scanner)
        self.scanner_btn.setMinimumWidth(120)
        layout.addWidget(self.scanner_btn)
        
        # EzAuto 창 제목
        layout.addWidget(QLabel("창 제목:"))
        self.ezauto_title_edit = QLineEdit()
        self.ezauto_title_edit.setText("이지오토")
        self.ezauto_title_edit.setMaximumWidth(100)
        self.ezauto_title_edit.textChanged.connect(self._on_ezauto_title_changed)
        layout.addWidget(self.ezauto_title_edit)
        
        # EzAuto 활성화
        self.ezauto_check = QCheckBox("EzAuto 입력")
        self.ezauto_check.setChecked(True)
        self.ezauto_check.toggled.connect(self._on_toggle_ezauto)
        layout.addWidget(self.ezauto_check)
        
        # PDF 출력 활성화
        self.pdf_check = QCheckBox("PDF 출력")
        self.pdf_check.setChecked(True)
        self.pdf_check.toggled.connect(self._on_toggle_pdf)
        layout.addWidget(self.pdf_check)
        
        return group
    
    def _create_tables_section(self) -> QWidget:
        """테이블 섹션"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setSpacing(10)
        
        # === 왼쪽: 현재 송장 상세 ===
        left_group = QGroupBox("현재 작업 중인 송장")
        left_layout = QVBoxLayout(left_group)
        
        # 현재 tracking_no 표시
        tracking_layout = QHBoxLayout()
        tracking_layout.addWidget(QLabel("송장번호:"))
        self.current_tracking_label = QLabel("-")
        self.current_tracking_label.setFont(QFont("Consolas", 14, QFont.Bold))
        self.current_tracking_label.setStyleSheet("color: #2196F3;")
        tracking_layout.addWidget(self.current_tracking_label)
        tracking_layout.addStretch()
        
        # 남은 수량
        tracking_layout.addWidget(QLabel("남은 수량:"))
        self.remaining_label = QLabel("0")
        self.remaining_label.setFont(QFont("Consolas", 14, QFont.Bold))
        self.remaining_label.setStyleSheet("color: #FF5722;")
        tracking_layout.addWidget(self.remaining_label)
        
        left_layout.addLayout(tracking_layout)
        
        # 상세 테이블
        self.detail_table = QTableWidget()
        self.detail_table.setColumnCount(6)
        self.detail_table.setHorizontalHeaderLabels([
            "상품명", "옵션명", "바코드", "필요수량", "스캔수량", "남은수량"
        ])
        self.detail_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.detail_table.setAlternatingRowColors(True)
        left_layout.addWidget(self.detail_table)
        
        layout.addWidget(left_group, 1)  # 5:5 비율
        
        # === 오른쪽: 전체 요약 ===
        right_group = QGroupBox("📦 남은 수량")
        right_layout = QVBoxLayout(right_group)
        
        # 수동 바코드 입력
        manual_layout = QHBoxLayout()
        self.manual_barcode_edit = QLineEdit()
        self.manual_barcode_edit.setPlaceholderText("수동 바코드 입력")
        self.manual_barcode_edit.returnPressed.connect(self._on_manual_scan)
        manual_layout.addWidget(self.manual_barcode_edit)
        
        self.manual_scan_btn = QPushButton("스캔")
        self.manual_scan_btn.clicked.connect(self._on_manual_scan)
        manual_layout.addWidget(self.manual_scan_btn)
        
        right_layout.addLayout(manual_layout)
        
        # 탭으로 구성별/제품별 구분
        from PySide6.QtWidgets import QTabWidget
        self.summary_tabs = QTabWidget()
        
        # 탭1: 구성별 요약
        self.combo_scroll = QScrollArea()
        self.combo_scroll.setWidgetResizable(True)
        self.combo_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_scroll.setStyleSheet("QScrollArea { border: none; background-color: #f0f0f0; }")
        
        self.summary_container = QWidget()
        self.summary_grid = QVBoxLayout(self.summary_container)
        self.summary_grid.setSpacing(8)
        self.summary_grid.setAlignment(Qt.AlignTop)
        self.combo_scroll.setWidget(self.summary_container)
        
        self.summary_tabs.addTab(self.combo_scroll, "구성별")
        
        # 탭2: 제품별 요약
        self.product_scroll = QScrollArea()
        self.product_scroll.setWidgetResizable(True)
        self.product_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.product_scroll.setStyleSheet("QScrollArea { border: none; background-color: #f5f5f5; }")
        
        self.product_container = QWidget()
        self.product_grid = QVBoxLayout(self.product_container)
        self.product_grid.setSpacing(5)
        self.product_grid.setAlignment(Qt.AlignTop)
        self.product_scroll.setWidget(self.product_container)
        
        self.summary_tabs.addTab(self.product_scroll, "제품별")
        
        right_layout.addWidget(self.summary_tabs)
        
        layout.addWidget(right_group, 1)  # 5:5 비율
        
        return widget
    
    def _create_log_section(self) -> QGroupBox:
        """로그 섹션"""
        group = QGroupBox("로그")
        layout = QVBoxLayout(group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setMaximumHeight(200)
        layout.addWidget(self.log_text)
        
        # 로그 제어 버튼
        btn_layout = QHBoxLayout()
        
        clear_log_btn = QPushButton("로그 지우기")
        clear_log_btn.clicked.connect(lambda: self.log_text.clear())
        btn_layout.addWidget(clear_log_btn)
        
        btn_layout.addStretch()
        
        # 저장 버튼
        save_btn = QPushButton("엑셀 저장")
        save_btn.clicked.connect(self._on_save_excel)
        btn_layout.addWidget(save_btn)
        
        # 다른 이름으로 저장 버튼
        save_as_btn = QPushButton("다른 이름으로 저장")
        save_as_btn.clicked.connect(self._on_save_excel_as)
        btn_layout.addWidget(save_as_btn)
        
        layout.addLayout(btn_layout)
        
        return group
    
    def _create_status_bar(self):
        """상태바 생성"""
        status = self.statusBar()
        
        self.status_scanner = QLabel("스캐너: 대기")
        self.status_file = QLabel("파일: 없음")
        self.status_count = QLabel("처리: 0건")
        
        status.addWidget(self.status_scanner)
        status.addWidget(QLabel(" | "))
        status.addWidget(self.status_file)
        status.addWidget(QLabel(" | "))
        status.addWidget(self.status_count)
    
    def _apply_styles(self):
        """스타일 적용"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                gridline-color: #eee;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:checked {
                background-color: #4CAF50;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
            QLineEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 6px;
            }
            QLineEdit:focus {
                border-color: #2196F3;
            }
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #1e1e1e;
                color: #d4d4d4;
            }
            QCheckBox {
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
    
    def _connect_signals(self):
        """시그널 연결"""
        # Excel 시그널
        self.excel_loader.data_loaded.connect(self._on_data_loaded)
        self.excel_loader.data_updated.connect(self._on_data_updated)
        self.excel_loader.error_occurred.connect(self._on_error)
        
        # Scanner 시그널
        self.scanner.barcode_scanned.connect(self._on_barcode_scanned)
        self.scanner.status_changed.connect(self._add_log)
        
        # EzAuto 시그널
        self.ezauto.input_success.connect(self._add_log)
        self.ezauto.input_error.connect(self._on_error)
        
        # PDF 시그널
        self.pdf_printer.print_success.connect(self._add_log)
        self.pdf_printer.print_error.connect(self._on_error)
        self.pdf_printer.index_updated.connect(self._on_pdf_indexed)
    
    @Slot(int)
    def _on_pdf_indexed(self, count: int):
        """PDF 인덱싱 완료"""
        if count > 0:
            self._add_log(f"PDF 인덱스: {count}개 송장번호")
        
        # Processor 시그널
        self.processor.scan_processed.connect(self._on_scan_processed)
        self.processor.tracking_completed.connect(self._on_tracking_completed)
        self.processor.ui_update_required.connect(self._update_tables)
        self.processor.log_message.connect(self._add_log)
    
    # === 이벤트 핸들러 ===
    
    @Slot()
    def _on_browse_excel(self):
        """엑셀 파일 찾아보기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xls *.xlsx);;XLS Files (*.xls);;XLSX Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.excel_path_edit.setText(file_path)
    
    @Slot()
    def _on_show_summary(self):
        """구성 요약 다이얼로그 표시"""
        if self.excel_loader.df is None:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 불러오세요.")
            return
        
        dialog = SummaryDialog(self.excel_loader.df, self)
        dialog.exec()
    
    @Slot()
    def _on_browse_pdf_file(self):
        """PDF 파일 찾아보기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "PDF 라벨 파일 선택",
            "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        if file_path:
            self.pdf_path_edit.setText(file_path)
            self.pdf_printer.set_pdf_file(file_path)
            self._add_log(f"PDF 파일 설정: {file_path}")
            
            # 자동 인덱싱
            self._add_log("PDF 파일 스캔 중...")
            count = self.pdf_printer.build_tracking_index()
            
            if count > 0:
                self._add_log(f"<b style='color:#4CAF50'>✓ PDF 스캔 완료: {count}개 송장번호 발견</b>", html=True)
            else:
                self._add_log("[경고] PDF에서 송장번호를 찾지 못했습니다")
    
    @Slot()
    def _on_load_excel(self):
        """엑셀 파일 로드"""
        file_path = self.excel_path_edit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "경고", "엑셀 파일 경로를 입력하세요.")
            return
        
        if self.excel_loader.load_excel(file_path):
            self._add_log(f"엑셀 파일 로드 성공: {file_path}")
            self.status_file.setText(f"파일: {Path(file_path).name}")
            
            # PDF 폴더 설정
            pdf_path = self.pdf_path_edit.text().strip()
            if pdf_path:
                self.pdf_printer.set_labels_directory(pdf_path)
            
            # 구성 요약 출력
            self._show_load_summary()
    
    @Slot()
    def _on_save_excel(self):
        """엑셀 파일 저장"""
        if self.excel_loader.save_excel():
            self._add_log("엑셀 파일 저장 완료")
            QMessageBox.information(self, "성공", "엑셀 파일이 저장되었습니다.")
        else:
            QMessageBox.warning(self, "오류", "엑셀 파일 저장에 실패했습니다.")
    
    @Slot()
    def _on_save_excel_as(self):
        """엑셀 파일 다른 이름으로 저장"""
        if self.excel_loader.df is None:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 불러오세요.")
            return
        
        # 파일 저장 대화상자
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "엑셀 파일 저장",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            # .xlsx 확장자 보장
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            
            if self.excel_loader.save_excel(file_path):
                self._add_log(f"엑셀 파일 저장 완료: {file_path}")
                QMessageBox.information(self, "성공", f"엑셀 파일이 저장되었습니다.\n{file_path}")
            else:
                QMessageBox.warning(self, "오류", "엑셀 파일 저장에 실패했습니다.")
    
    @Slot()
    def _on_toggle_scanner(self):
        """스캐너 시작/중지"""
        if self.scanner_btn.isChecked():
            if self.scanner.start():
                self._scanner_active = True
                self.scanner_btn.setText("스캐너 중지")
                self.status_scanner.setText("스캐너: 활성")
                self._add_log("스캐너 활성화됨")
            else:
                self.scanner_btn.setChecked(False)
        else:
            self.scanner.stop()
            self._scanner_active = False
            self.scanner_btn.setText("스캐너 시작")
            self.status_scanner.setText("스캐너: 대기")
            self._add_log("스캐너 비활성화됨")
    
    @Slot(bool)
    def _on_toggle_ezauto(self, checked: bool):
        """EzAuto 활성화/비활성화"""
        self.ezauto.enabled = checked
        self._add_log(f"EzAuto 입력: {'활성' if checked else '비활성'}")
    
    @Slot(str)
    def _on_ezauto_title_changed(self, title: str):
        """EzAuto 창 제목 변경"""
        self.ezauto.set_window_title(title)
    
    @Slot(bool)
    def _on_toggle_pdf(self, checked: bool):
        """PDF 출력 활성화/비활성화"""
        self.pdf_printer.enabled = checked
        self._add_log(f"PDF 출력: {'활성' if checked else '비활성'}")
    
    @Slot()
    def _on_manual_scan(self):
        """수동 바코드 스캔"""
        barcode = self.manual_barcode_edit.text().strip()
        if barcode:
            self._on_barcode_scanned(barcode)
            self.manual_barcode_edit.clear()
    
    @Slot(str)
    def _on_barcode_scanned(self, barcode: str):
        """바코드 스캔 이벤트"""
        if self.excel_loader.df is None:
            self._add_log("[경고] 엑셀 파일을 먼저 로드하세요")
            return
        
        self.processor.process_scan(barcode)
    
    @Slot(object)
    def _on_scan_processed(self, event: ScanEvent):
        """스캔 처리 완료"""
        # 결과에 따른 색상
        if event.result == ScanResult.SUCCESS:
            color = "#4CAF50"  # 녹색
        elif event.result == ScanResult.ALREADY_USED:
            color = "#FF9800"  # 주황색
        else:
            color = "#F44336"  # 빨간색
        
        self._add_log(f"<span style='color:{color}'>{event.message}</span>", html=True)
    
    @Slot(str)
    def _on_tracking_completed(self, tracking_no: str):
        """송장 완료"""
        self._add_log(f"<b style='color:#4CAF50'>✓ 송장 {tracking_no} 완료!</b>", html=True)
        self._update_status_count()
    
    @Slot()
    def _on_data_loaded(self):
        """데이터 로드 완료"""
        self._update_tables()
        self._update_status_count()
    
    @Slot()
    def _on_data_updated(self):
        """데이터 업데이트"""
        self._update_tables()
    
    @Slot(str)
    def _on_error(self, message: str):
        """오류 발생"""
        self._add_log(f"<span style='color:#F44336'>[오류] {message}</span>", html=True)
    
    # === UI 업데이트 ===
    
    def _update_tables(self):
        """테이블 업데이트"""
        self._update_detail_table()
        self._update_summary_table()
    
    def _update_detail_table(self):
        """현재 송장 상세 테이블 업데이트"""
        tracking_no = self.processor.current_tracking_no
        
        if not tracking_no:
            self.current_tracking_label.setText("-")
            self.remaining_label.setText("0")
            self.detail_table.setRowCount(0)
            return
        
        self.current_tracking_label.setText(tracking_no)
        
        items = self.processor.get_current_tracking_items()
        if items.empty:
            self.detail_table.setRowCount(0)
            return
        
        # 남은 수량 계산
        remaining = self.excel_loader.get_group_remaining(tracking_no)
        self.remaining_label.setText(str(remaining))
        
        # 테이블 업데이트
        self.detail_table.setRowCount(len(items))
        
        for row, (_, item) in enumerate(items.iterrows()):
            item_remaining = max(0, item['qty'] - item['scanned_qty'])
            
            self.detail_table.setItem(row, 0, QTableWidgetItem(str(item['product_name'])))
            self.detail_table.setItem(row, 1, QTableWidgetItem(str(item['option_name'])))
            self.detail_table.setItem(row, 2, QTableWidgetItem(str(item['barcode'])))
            self.detail_table.setItem(row, 3, QTableWidgetItem(str(item['qty'])))
            self.detail_table.setItem(row, 4, QTableWidgetItem(str(item['scanned_qty'])))
            self.detail_table.setItem(row, 5, QTableWidgetItem(str(item_remaining)))
            
            # 완료된 항목은 녹색으로 표시
            if item_remaining == 0:
                for col in range(6):
                    self.detail_table.item(row, col).setBackground(QColor("#E8F5E9"))
    
    def _update_summary_table(self):
        """요약 카드 업데이트 (구성별 + 제품별)"""
        if self.excel_loader.df is None:
            return
        
        df = self.excel_loader.df
        pending = df[df['used'] == 0]
        
        # === 구성별 카드 업데이트 ===
        while self.summary_grid.count():
            item = self.summary_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if pending.empty:
            empty_label = QLabel("✅ 모든 송장 처리 완료!")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("font-size: 16px; color: #4CAF50; padding: 20px;")
            self.summary_grid.addWidget(empty_label)
        else:
            combo_data = self._get_summary_combo_data(pending)
            for combo_info in combo_data:
                card = self._create_summary_card(combo_info)
                self.summary_grid.addWidget(card)
            self.summary_grid.addStretch()
        
        # === 제품별 요약 업데이트 ===
        while self.product_grid.count():
            item = self.product_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if pending.empty:
            empty_label2 = QLabel("✅ 모든 제품 처리 완료!")
            empty_label2.setAlignment(Qt.AlignCenter)
            empty_label2.setStyleSheet("font-size: 16px; color: #4CAF50; padding: 20px;")
            self.product_grid.addWidget(empty_label2)
        else:
            product_data = self._get_product_summary(pending)
            for prod_info in product_data:
                prod_card = self._create_product_card(prod_info)
                self.product_grid.addWidget(prod_card)
            self.product_grid.addStretch()
    
    def _get_product_summary(self, pending):
        """제품별 남은 수량 계산"""
        product_summary = {}
        
        for _, row in pending.iterrows():
            product_name = str(row['product_name']) if pd.notna(row['product_name']) else ''
            option_name = str(row['option_name']) if pd.notna(row['option_name']) else ''
            qty = int(row['qty']) if pd.notna(row['qty']) else 1
            scanned = int(row['scanned_qty']) if pd.notna(row['scanned_qty']) else 0
            remaining = qty - scanned
            
            key = f"{product_name}|{option_name}"
            if key not in product_summary:
                product_summary[key] = {
                    'product_name': product_name,
                    'option_name': option_name,
                    'total_qty': 0,
                    'remaining': 0
                }
            product_summary[key]['total_qty'] += qty
            product_summary[key]['remaining'] += remaining
        
        # 남은 수량 내림차순 정렬
        return sorted(product_summary.values(), key=lambda x: -x['remaining'])
    
    def _create_product_card(self, prod_info):
        """제품별 카드 생성"""
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        
        remaining = prod_info['remaining']
        if remaining >= 20:
            bg_color = "#FFEBEE"
            text_color = "#D32F2F"
        elif remaining >= 10:
            bg_color = "#FFF3E0"
            text_color = "#E65100"
        elif remaining >= 5:
            bg_color = "#E3F2FD"
            text_color = "#1565C0"
        else:
            bg_color = "#F5F5F5"
            text_color = "#616161"
        
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 5px 8px;
                margin: 1px;
            }}
        """)
        
        layout = QHBoxLayout(card)
        layout.setContentsMargins(5, 3, 5, 3)
        layout.setSpacing(8)
        
        # 남은 수량
        count_label = QLabel(f"<b style='color:{text_color};'>{remaining}</b>")
        count_label.setFixedWidth(35)
        count_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(count_label)
        
        # 제품명 + 옵션
        product_text = prod_info['product_name']
        if prod_info['option_name'] and prod_info['option_name'] != 'nan':
            product_text += f" ({prod_info['option_name']})"
        
        prod_label = QLabel(product_text)
        prod_label.setWordWrap(True)
        prod_label.setStyleSheet("font-size: 11px; color: #333;")
        layout.addWidget(prod_label, 1)
        
        return card
    
    def _get_summary_combo_data(self, pending):
        """구성별 데이터 추출 (수량 포함)"""
        tracking_groups = pending.groupby('tracking_no')
        combo_counts = {}
        
        for tracking_no, group in tracking_groups:
            barcodes = tuple(sorted(group['barcode'].unique()))
            
            if barcodes not in combo_counts:
                combo_counts[barcodes] = {
                    'count': 0,
                    'products': [],
                    'barcodes': list(barcodes)
                }
                for _, row in group.iterrows():
                    product_name = str(row['product_name']) if pd.notna(row['product_name']) else ''
                    option_name = str(row['option_name']) if pd.notna(row['option_name']) else ''
                    qty = int(row['qty']) if pd.notna(row['qty']) else 1
                    
                    product_info = product_name
                    if option_name and option_name != 'nan':
                        product_info += f" ({option_name})"
                    
                    # 수량 뒤에 표시: 1개, 2개, 3개...
                    product_info += f" {qty}개"
                    
                    if product_info and product_info not in combo_counts[barcodes]['products']:
                        combo_counts[barcodes]['products'].append(product_info)
            
            combo_counts[barcodes]['count'] += 1
        
        return sorted(combo_counts.values(), key=lambda x: -x['count'])
    
    def _create_summary_card(self, combo_info):
        """요약 카드 생성 (가로 나열, 전체 품목 표시)"""
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        
        count = combo_info['count']
        if count >= 10:
            bg_color = "#FFEBEE"
            border_color = "#EF5350"
            count_color = "#D32F2F"
        elif count >= 5:
            bg_color = "#FFF3E0"
            border_color = "#FF9800"
            count_color = "#E65100"
        elif count >= 3:
            bg_color = "#E3F2FD"
            border_color = "#2196F3"
            count_color = "#1565C0"
        else:
            bg_color = "#F5F5F5"
            border_color = "#9E9E9E"
            count_color = "#616161"
        
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 2px solid {border_color};
                border-radius: 8px;
                padding: 6px 10px;
                margin: 2px;
            }}
        """)
        
        layout = QHBoxLayout(card)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # 개수 배지 (3자리 지원)
        count_label = QLabel(f"<b style='font-size:16px; color:{count_color};'>{count}</b>")
        count_label.setFixedWidth(50)
        count_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(count_label)
        
        # 상품 목록 (◆ 구분자로 명확히 구분)
        products = combo_info['products']
        products_text = "  ◆  ".join(products)
        
        prod_label = QLabel(products_text)
        prod_label.setWordWrap(True)
        prod_label.setStyleSheet("font-size: 11px; color: #333; line-height: 1.4;")
        layout.addWidget(prod_label, 1)
        
        return card
    
    def _update_status_count(self):
        """상태바 카운트 업데이트"""
        if self.excel_loader.df is None:
            self.status_count.setText("처리: 0건")
            return
        
        total = len(self.excel_loader.df['tracking_no'].unique())
        completed = len(self.excel_loader.df[self.excel_loader.df['used'] == 1]['tracking_no'].unique())
        self.status_count.setText(f"처리: {completed}/{total}건")
    
    def _show_load_summary(self):
        """엑셀 로드 후 구성 요약 로그 표시 (다이얼로그 없음)"""
        if self.excel_loader.df is None:
            return
        
        df = self.excel_loader.df
        pending = df[df['used'] == 0]
        
        # 전체 통계
        total_tracking = len(df['tracking_no'].unique())
        pending_tracking = len(pending['tracking_no'].unique())
        
        self._add_log(f"총 송장: {total_tracking}건, 미처리: {pending_tracking}건")
    
    def _add_log(self, message: str, html: bool = False):
        """로그 추가"""
        timestamp = get_timestamp()
        if html:
            self.log_text.append(f"[{timestamp}] {message}")
        else:
            self.log_text.append(f"[{timestamp}] {message}")
        
        # 스크롤 아래로
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def closeEvent(self, event):
        """프로그램 종료 시"""
        # 스캐너 중지
        self.scanner.stop()
        
        # 데이터 저장 확인
        if self.excel_loader.df is not None:
            reply = QMessageBox.question(
                self, "저장 확인",
                "변경사항을 저장하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            
            if reply == QMessageBox.Yes:
                self.excel_loader.save_excel()
                event.accept()
            elif reply == QMessageBox.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def run_app():
    """애플리케이션 실행"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    return app.exec()

