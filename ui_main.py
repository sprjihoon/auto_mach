"""
PySide6 UI 화면
"""
import sys
import os
from pathlib import Path
from typing import Optional
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QTextEdit, QPushButton,
    QLabel, QLineEdit, QFileDialog, QGroupBox, QSplitter,
    QHeaderView, QMessageBox, QFrame, QCheckBox, QDialog,
    QScrollArea, QGridLayout, QListWidget, QListWidgetItem,
    QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, Slot, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QColor, QPalette, QIcon
import pandas as pd

from models import ScanResult, ScanEvent
from excel_loader import ExcelLoader
from normalize_pdf import normalize_pdf


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
        
        # 우선순위 규칙 초기화 (기본값: 단품 우선)
        from priority_engine import get_default_rules
        self.processor.set_priority_rules(get_default_rules())
        
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
        
        # === 중간: 스플리터 (테이블들 + 우선순위 설정 + 로그) ===
        splitter = QSplitter(Qt.Vertical)
        
        # 테이블 영역
        tables_widget = self._create_tables_section()
        splitter.addWidget(tables_widget)
        
        # 우선순위 설정 영역 (우선순위 설정 + 우선 송장 관리)
        priority_section = self._create_priority_section()
        splitter.addWidget(priority_section)
        
        # 로그 영역
        log_group = self._create_log_section()
        splitter.addWidget(log_group)
        
        splitter.setSizes([400, 200, 200])
        main_layout.addWidget(splitter, 1)
        
        # === 하단: 상태바 ===
        self._create_status_bar()
        
        # 스타일 적용
        self._apply_styles()
    
    def _create_top_section(self) -> QGroupBox:
        """상단 섹션: 파일 로드 및 설정"""
        group = QGroupBox("설정")
        layout = QHBoxLayout(group)
        layout.setSpacing(5)  # 요소간 간격 줄임
        
        # 엑셀 파일 경로
        layout.addWidget(QLabel("엑셀:"))
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("엑셀 파일 선택")
        self.excel_path_edit.setMaximumWidth(180)
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
        
        layout.addSpacing(15)
        
        # PDF 파일 경로
        layout.addWidget(QLabel("PDF:"))
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("PDF 선택")
        self.pdf_path_edit.setMaximumWidth(180)
        layout.addWidget(self.pdf_path_edit)
        
        # PDF 파일 찾아보기 버튼
        self.pdf_browse_btn = QPushButton("파일 선택")
        self.pdf_browse_btn.clicked.connect(self._on_browse_pdf_file)
        layout.addWidget(self.pdf_browse_btn)
        
        layout.addSpacing(15)
        
        # 스캐너 시작/중지
        self.scanner_btn = QPushButton("스캐너 시작")
        self.scanner_btn.setCheckable(True)
        self.scanner_btn.clicked.connect(self._on_toggle_scanner)
        self.scanner_btn.setMinimumWidth(100)
        layout.addWidget(self.scanner_btn)
        
        # EzAuto 창 제목
        layout.addWidget(QLabel("창 제목:"))
        self.ezauto_title_edit = QLineEdit()
        self.ezauto_title_edit.setText("이지오토")
        self.ezauto_title_edit.setMaximumWidth(80)
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
        
        # 오른쪽 여백 (창 최대화 시 벌어짐 방지)
        layout.addStretch()
        
        return group
    
    def _create_priority_section(self) -> QWidget:
        """우선순위 설정 섹션 (우선순위 설정 + 우선 송장 관리)"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setSpacing(10)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 우선순위 설정 패널
        priority_group = self._create_priority_panel()
        layout.addWidget(priority_group, 1)
        
        # 우선 송장 추가 패널
        priority_tracking_group = self._create_priority_tracking_panel()
        layout.addWidget(priority_tracking_group, 1)
        
        return widget
    
    def _create_priority_panel(self) -> QGroupBox:
        """우선순위 설정 패널"""
        group = QGroupBox("우선순위 설정")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(8, 15, 8, 8)
        
        # 상호 배타적 옵션을 라디오 버튼으로 구성
        grid = QGridLayout()
        grid.setSpacing(8)
        
        # 1. 단품/조합 선택 (라디오 버튼 그룹)
        single_combo_group = QButtonGroup(group)
        single_combo_layout = QHBoxLayout()
        single_combo_layout.addWidget(QLabel("품목 유형:"))
        
        self.priority_single_radio = QRadioButton("단품 우선")
        self.priority_single_radio.setChecked(True)  # 기본값: 단품 우선
        self.priority_single_radio.toggled.connect(self._on_priority_changed)
        single_combo_group.addButton(self.priority_single_radio, 0)
        single_combo_layout.addWidget(self.priority_single_radio)
        
        self.priority_combo_radio = QRadioButton("조합 우선")
        self.priority_combo_radio.setChecked(False)
        self.priority_combo_radio.toggled.connect(self._on_priority_changed)
        single_combo_group.addButton(self.priority_combo_radio, 1)
        single_combo_layout.addWidget(self.priority_combo_radio)
        
        single_combo_layout.addStretch()
        grid.addLayout(single_combo_layout, 0, 0, 1, 2)
        
        # 2. 수량 선택 (라디오 버튼 그룹)
        qty_group = QButtonGroup(group)
        qty_layout = QHBoxLayout()
        qty_layout.addWidget(QLabel("수량 기준:"))
        
        self.priority_small_qty_radio = QRadioButton("소량 우선")
        self.priority_small_qty_radio.setChecked(False)
        self.priority_small_qty_radio.toggled.connect(self._on_priority_changed)
        qty_group.addButton(self.priority_small_qty_radio, 0)
        qty_layout.addWidget(self.priority_small_qty_radio)
        
        self.priority_large_qty_radio = QRadioButton("대량 우선")
        self.priority_large_qty_radio.setChecked(False)
        self.priority_large_qty_radio.toggled.connect(self._on_priority_changed)
        qty_group.addButton(self.priority_large_qty_radio, 1)
        qty_layout.addWidget(self.priority_large_qty_radio)
        
        # 선택 안 함 옵션 추가
        self.priority_no_qty_radio = QRadioButton("수량 무관")
        self.priority_no_qty_radio.setChecked(True)  # 기본값: 수량 무관
        self.priority_no_qty_radio.toggled.connect(self._on_priority_changed)
        qty_group.addButton(self.priority_no_qty_radio, 2)
        qty_layout.addWidget(self.priority_no_qty_radio)
        
        qty_layout.addStretch()
        grid.addLayout(qty_layout, 1, 0, 1, 2)
        
        # 3. 주문 시간 선택 (라디오 버튼 그룹)
        order_time_group = QButtonGroup(group)
        order_time_layout = QHBoxLayout()
        order_time_layout.addWidget(QLabel("주문 시간:"))
        
        self.priority_old_order_radio = QRadioButton("오래된 주문 우선")
        self.priority_old_order_radio.setChecked(False)
        self.priority_old_order_radio.toggled.connect(self._on_priority_changed)
        order_time_group.addButton(self.priority_old_order_radio, 0)
        order_time_layout.addWidget(self.priority_old_order_radio)
        
        self.priority_new_order_radio = QRadioButton("최신 주문 우선")
        self.priority_new_order_radio.setChecked(False)
        self.priority_new_order_radio.toggled.connect(self._on_priority_changed)
        order_time_group.addButton(self.priority_new_order_radio, 1)
        order_time_layout.addWidget(self.priority_new_order_radio)
        
        # 선택 안 함 옵션 추가
        self.priority_no_time_radio = QRadioButton("시간 무관")
        self.priority_no_time_radio.setChecked(True)  # 기본값: 시간 무관
        self.priority_no_time_radio.toggled.connect(self._on_priority_changed)
        order_time_group.addButton(self.priority_no_time_radio, 2)
        order_time_layout.addWidget(self.priority_no_time_radio)
        
        order_time_layout.addStretch()
        grid.addLayout(order_time_layout, 2, 0, 1, 2)
        
        layout.addLayout(grid)
        
        # 프리셋 버튼 영역
        preset_layout = QHBoxLayout()
        preset_layout.setSpacing(5)
        
        self.preset_default_btn = QPushButton("📌 기본(단품 우선)")
        self.preset_default_btn.setMaximumHeight(30)
        self.preset_default_btn.clicked.connect(lambda: self._apply_preset("default"))
        preset_layout.addWidget(self.preset_default_btn)
        
        self.preset_backlog_btn = QPushButton("📋 밀린 주문 정리")
        self.preset_backlog_btn.setMaximumHeight(30)
        self.preset_backlog_btn.clicked.connect(lambda: self._apply_preset("backlog"))
        preset_layout.addWidget(self.preset_backlog_btn)
        
        self.preset_bulk_btn = QPushButton("📦 대량 소화")
        self.preset_bulk_btn.setMaximumHeight(30)
        self.preset_bulk_btn.clicked.connect(lambda: self._apply_preset("bulk"))
        preset_layout.addWidget(self.preset_bulk_btn)
        
        layout.addLayout(preset_layout)
        
        # 초기 우선순위 규칙 적용
        self._apply_priority_rules()
        
        return group
    
    def _create_priority_tracking_panel(self) -> QGroupBox:
        """우선 송장 추가 패널 (방식 B: 직접 입력)"""
        group = QGroupBox("⭐ 우선 송장 관리")
        layout = QVBoxLayout(group)
        layout.setSpacing(5)
        layout.setContentsMargins(8, 15, 8, 8)
        
        # 입력 영역
        input_layout = QHBoxLayout()
        
        self.priority_tracking_input = QLineEdit()
        self.priority_tracking_input.setPlaceholderText("송장번호 입력/붙여넣기 (여러 개: 줄바꿈 또는 쉼표 구분)")
        self.priority_tracking_input.returnPressed.connect(self._on_add_priority_tracking)
        input_layout.addWidget(self.priority_tracking_input)
        
        add_btn = QPushButton("추가")
        add_btn.clicked.connect(self._on_add_priority_tracking)
        add_btn.setMaximumWidth(60)
        input_layout.addWidget(add_btn)
        
        remove_btn = QPushButton("해제")
        remove_btn.clicked.connect(self._on_remove_priority_tracking)
        remove_btn.setMaximumWidth(60)
        input_layout.addWidget(remove_btn)
        
        layout.addLayout(input_layout)
        
        # 우선 송장 목록
        list_label = QLabel("우선 송장 목록:")
        layout.addWidget(list_label)
        
        self.priority_tracking_list = QListWidget()
        self.priority_tracking_list.setMaximumHeight(100)
        self.priority_tracking_list.setSelectionMode(QListWidget.SingleSelection)
        layout.addWidget(self.priority_tracking_list)
        
        # 설명 텍스트
        help_label = QLabel("💡 여러 송장번호를 한 번에 입력 가능 (줄바꿈 또는 쉼표로 구분)")
        help_label.setStyleSheet("font-size: 9px; color: #666;")
        layout.addWidget(help_label)
        
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
        
        # 저장 경로 설정
        btn_layout.addWidget(QLabel("저장 위치:"))
        self.save_path_edit = QLineEdit()
        self.save_path_edit.setPlaceholderText("저장 위치 선택")
        self.save_path_edit.setMaximumWidth(200)
        btn_layout.addWidget(self.save_path_edit)
        
        self.save_browse_btn = QPushButton("위치 선택")
        self.save_browse_btn.clicked.connect(self._on_browse_save_path)
        btn_layout.addWidget(self.save_browse_btn)
        
        # 저장 버튼
        save_btn = QPushButton("엑셀 저장")
        save_btn.clicked.connect(self._on_save_excel)
        btn_layout.addWidget(save_btn)
        
        # 제품별 PDF 저장 버튼
        pdf_save_btn = QPushButton("📄 피킹리스트 PDF")
        pdf_save_btn.clicked.connect(self._on_save_product_pdf)
        btn_layout.addWidget(pdf_save_btn)
        
        # 피킹리스트 열기 버튼
        self.open_pdf_btn = QPushButton("📂 피킹리스트 열기")
        self.open_pdf_btn.clicked.connect(self._on_open_picking_pdf)
        self.open_pdf_btn.setEnabled(False)  # 초기에는 비활성화
        btn_layout.addWidget(self.open_pdf_btn)
        
        # 마지막 저장된 PDF 경로
        self._last_pdf_path = None
        
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
        self.excel_loader.priority_cleared.connect(self._on_priority_cleared)
        
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
        self.processor.scanner_pause.connect(self.scanner.pause)
        self.processor.scanner_resume.connect(self.scanner.resume)
    
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
            # PDF 크롭 처리
            try:
                import tempfile
                temp_dir = Path(tempfile.gettempdir()) / "auto_mach_labels"
                temp_dir.mkdir(exist_ok=True)
                
                # 크롭된 PDF 저장 경로
                original_path = Path(file_path)
                cropped_path = temp_dir / f"cropped_{original_path.stem}.pdf"
                
                self._add_log("PDF 크롭 처리 중... (168mm × 107mm)")
                normalize_pdf(file_path, str(cropped_path))
                self._add_log(f"✓ PDF 크롭 완료: {cropped_path}")
                
                # 크롭된 PDF 사용
                self.pdf_path_edit.setText(file_path)  # 원본 경로 표시
                self.pdf_printer.set_pdf_file(str(cropped_path))  # 크롭된 파일 사용
                self._add_log(f"PDF 파일 설정: {file_path} (크롭된 버전 사용)")
            except Exception as e:
                self._add_log(f"[오류] PDF 크롭 실패: {str(e)}. 원본 파일 사용.")
                self.pdf_path_edit.setText(file_path)
                self.pdf_printer.set_pdf_file(file_path)
            
            # 자동 인덱싱
            self._add_log("PDF 파일 스캔 중...")
            
            # 엑셀에서 송장번호 목록 가져오기 (이미지 PDF의 경우 순서대로 매핑)
            excel_tracking_numbers = None
            if self.excel_loader.df is not None and 'tracking_no' in self.excel_loader.df.columns:
                # 순서를 보장하기 위해 drop_duplicates 사용 (첫 번째 출현 순서 유지)
                excel_tracking_numbers = self.excel_loader.df['tracking_no'].drop_duplicates().tolist()
                self._add_log(f"엑셀 송장번호 순서: {', '.join(map(str, excel_tracking_numbers[:5]))}..." if len(excel_tracking_numbers) > 5 else f"엑셀 송장번호: {', '.join(map(str, excel_tracking_numbers))}")
            
            count = self.pdf_printer.build_tracking_index(excel_tracking_numbers)
            
            if count > 0:
                self._add_log(f"<b style='color:#4CAF50'>✓ PDF 스캔 완료: {count}개 송장번호 발견</b>", html=True)
            else:
                if excel_tracking_numbers:
                    self._add_log("[경고] PDF에서 송장번호를 찾지 못했습니다. 이미지 기반 PDF일 수 있습니다.")
                else:
                    self._add_log("[경고] PDF에서 송장번호를 찾지 못했습니다. 엑셀 파일을 먼저 로드하면 자동 매핑됩니다.")
    
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
            
            # PDF 파일이 설정되어 있으면 자동으로 다시 스캔 (이미지 PDF 매핑을 위해)
            pdf_file_path = self.pdf_path_edit.text().strip()
            if pdf_file_path and os.path.exists(pdf_file_path):
                self.pdf_printer.set_pdf_file(pdf_file_path)
                self._add_log("엑셀 로드 후 PDF 재스캔 중...")
                
                # 엑셀에서 송장번호 목록 가져오기 (순서 보장)
                excel_tracking_numbers = None
                if self.excel_loader.df is not None and 'tracking_no' in self.excel_loader.df.columns:
                    # 순서를 보장하기 위해 drop_duplicates 사용 (첫 번째 출현 순서 유지)
                    excel_tracking_numbers = self.excel_loader.df['tracking_no'].drop_duplicates().tolist()
                    self._add_log(f"엑셀 송장번호 순서: {', '.join(map(str, excel_tracking_numbers[:5]))}..." if len(excel_tracking_numbers) > 5 else f"엑셀 송장번호: {', '.join(map(str, excel_tracking_numbers))}")
                
                count = self.pdf_printer.build_tracking_index(excel_tracking_numbers)
                
                if count > 0:
                    self._add_log(f"<b style='color:#4CAF50'>✓ PDF 재스캔 완료: {count}개 송장번호 발견</b>", html=True)
                else:
                    self._add_log("[경고] PDF 재스캔 실패: 송장번호를 찾지 못했습니다.")
            
            # 구성 요약 출력
            self._show_load_summary()
    
    @Slot()
    def _on_save_excel(self):
        """엑셀 파일 저장 (파일명_역매칭.xlsx로 저장)"""
        if self.excel_loader.df is None:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 불러오세요.")
            return
        
        # 저장 경로 확인
        save_path = self.save_path_edit.text().strip()
        
        if save_path:
            # 지정된 경로로 저장
            success, saved_path = self.excel_loader.save_excel(save_path)
            if success:
                self._add_log(f"엑셀 파일 저장 완료: {saved_path}")
                QMessageBox.information(self, "성공", f"엑셀 파일이 저장되었습니다.\n{saved_path}")
            else:
                QMessageBox.warning(self, "오류", "엑셀 파일 저장에 실패했습니다.")
        else:
            # 원본 위치에 _역매칭 붙여서 저장
            success, saved_path = self.excel_loader.save_excel()
            if success:
                self._add_log(f"엑셀 파일 저장 완료: {saved_path}")
                QMessageBox.information(self, "성공", f"엑셀 파일이 저장되었습니다.\n{saved_path}")
            else:
                QMessageBox.warning(self, "오류", "엑셀 파일 저장에 실패했습니다.")
    
    @Slot()
    def _on_save_product_pdf(self):
        """제품별 요약을 PDF로 저장"""
        if self.excel_loader.df is None:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 불러오세요.")
            return
        
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 저장 경로가 지정되어 있으면 해당 폴더에 자동 저장
        save_path = self.save_path_edit.text().strip()
        if save_path:
            # 지정된 경로의 폴더에 피킹리스트 PDF 저장
            save_dir = Path(save_path).parent
            file_path = str(save_dir / f"피킹리스트_{timestamp}.pdf")
        else:
            # 파일 저장 경로 선택 (기본 파일명에 타임스탬프 포함)
            default_name = f"피킹리스트_{timestamp}.pdf"
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "제품별 요약 PDF 저장",
                default_name,
                "PDF Files (*.pdf);;All Files (*)"
            )
            
            if not file_path:
                return
        
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'
        
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            
            # 한글 폰트 등록 (맑은 고딕)
            try:
                pdfmetrics.registerFont(TTFont('MalgunGothic', 'C:/Windows/Fonts/malgun.ttf'))
                font_name = 'MalgunGothic'
            except:
                font_name = 'Helvetica'
            
            # 데이터 준비
            df = self.excel_loader.df
            pending = df[df['used'] == 0]
            
            if pending.empty:
                QMessageBox.information(self, "알림", "처리할 제품이 없습니다.")
                return
            
            # 로케이션 컬럼 확인
            has_location = 'location' in pending.columns
            
            # 제품별 집계 (UI와 동일하게 product_name + option_name으로 그룹화)
            product_data = []
            product_summary = {}
            
            for _, row in pending.iterrows():
                product_name = str(row['product_name']) if pd.notna(row['product_name']) else ''
                option_name = str(row['option_name']) if pd.notna(row['option_name']) else ''
                barcode = str(row['barcode']) if pd.notna(row['barcode']) else ''
                qty = int(row['qty']) if pd.notna(row['qty']) else 1
                scanned = int(row['scanned_qty']) if pd.notna(row['scanned_qty']) else 0
                remaining = qty - scanned
                
                location = ''
                if has_location and 'location' in row and pd.notna(row['location']):
                    location = str(row['location'])
                
                key = f"{product_name}|{option_name}"
                if key not in product_summary:
                    product_summary[key] = {
                        'product_name': product_name,
                        'option_name': option_name,
                        'remaining': 0,
                        'location': location,
                        'barcode': barcode
                    }
                product_summary[key]['remaining'] += remaining
            
            # 남은 수량이 있는 것만 추가
            for item in product_summary.values():
                if item['remaining'] > 0:
                    product_data.append(item)
            
            # 수량 내림차순 정렬
            product_data.sort(key=lambda x: -x['remaining'])
            
            # PDF 생성
            doc = SimpleDocTemplate(file_path, pagesize=A4, 
                                   leftMargin=15*mm, rightMargin=15*mm,
                                   topMargin=15*mm, bottomMargin=15*mm)
            
            elements = []
            
            # 스타일
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'Title',
                parent=styles['Heading1'],
                fontName=font_name,
                fontSize=18,
                alignment=1  # 중앙 정렬
            )
            
            # 제목
            from datetime import datetime
            title = Paragraph(f"제품별 피킹 리스트 ({datetime.now().strftime('%Y-%m-%d %H:%M')})", title_style)
            elements.append(title)
            elements.append(Spacer(1, 10*mm))
            
            # 테이블 헤더
            if has_location:
                headers = ['No', '수량', '로케이션', '제품명', '옵션명', '바코드']
                col_widths = [10*mm, 15*mm, 25*mm, 55*mm, 40*mm, 35*mm]
            else:
                headers = ['No', '수량', '제품명', '옵션명', '바코드']
                col_widths = [10*mm, 15*mm, 70*mm, 50*mm, 35*mm]
            
            # 테이블 데이터
            table_data = [headers]
            for i, item in enumerate(product_data, 1):
                if has_location:
                    row = [
                        str(i),
                        str(item['remaining']),
                        item['location'],
                        item['product_name'][:30],
                        item['option_name'][:20] if item['option_name'] != 'nan' else '',
                        item['barcode']
                    ]
                else:
                    row = [
                        str(i),
                        str(item['remaining']),
                        item['product_name'][:40],
                        item['option_name'][:25] if item['option_name'] != 'nan' else '',
                        item['barcode']
                    ]
                table_data.append(row)
            
            # 테이블 생성
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), font_name),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2196F3')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('ALIGN', (0, 1), (1, -1), 'CENTER'),  # No, 수량 중앙
                ('ALIGN', (2, 1), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            
            elements.append(table)
            
            # 합계
            total_remaining = sum(item['remaining'] for item in product_data)
            summary_style = ParagraphStyle(
                'Summary',
                parent=styles['Normal'],
                fontName=font_name,
                fontSize=12,
                alignment=2  # 오른쪽 정렬
            )
            elements.append(Spacer(1, 5*mm))
            elements.append(Paragraph(f"총 {len(product_data)}개 품목 / {total_remaining}개 수량", summary_style))
            
            # PDF 저장
            doc.build(elements)
            
            self._add_log(f"제품별 PDF 저장 완료: {file_path}")
            
            # 마지막 PDF 경로 저장 및 열기 버튼 활성화
            self._last_pdf_path = file_path
            self.open_pdf_btn.setEnabled(True)
            
            QMessageBox.information(self, "성공", f"PDF가 저장되었습니다.\n{file_path}")
            
        except ImportError:
            QMessageBox.warning(self, "오류", "reportlab 패키지가 필요합니다.\npip install reportlab")
        except Exception as e:
            self._add_log(f"[오류] PDF 저장 실패: {str(e)}")
            QMessageBox.warning(self, "오류", f"PDF 저장 실패: {str(e)}")
    
    @Slot()
    def _on_open_picking_pdf(self):
        """마지막 저장된 피킹리스트 PDF 열기"""
        if self._last_pdf_path and Path(self._last_pdf_path).exists():
            import os
            os.startfile(self._last_pdf_path)
            self._add_log(f"피킹리스트 열기: {self._last_pdf_path}")
        else:
            QMessageBox.warning(self, "경고", "열 수 있는 PDF 파일이 없습니다.\n먼저 피킹리스트 PDF를 저장하세요.")
    
    @Slot()
    def _on_browse_save_path(self):
        """저장 경로 선택"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "엑셀 저장 위치 선택",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            # .xlsx 확장자 보장
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            self.save_path_edit.setText(file_path)
            self._add_log(f"저장 위치 설정: {file_path}")
    
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
    def _on_priority_changed(self):
        """우선순위 설정 변경 (라디오 버튼 자동 상호 배타적)"""
        self._apply_priority_rules()
    
    def _apply_preset(self, preset_name: str):
        """
        프리셋 적용
        
        Args:
            preset_name: 프리셋 이름 ("default", "backlog", "bulk")
        """
        from priority_engine import get_preset_rules
        
        # 프리셋 규칙 가져오기
        rules = get_preset_rules(preset_name)
        
        # 라디오 버튼 UI 상태 업데이트 (시그널 차단하여 무한 루프 방지)
        if hasattr(self, 'priority_single_radio'):
            self.priority_single_radio.blockSignals(True)
            self.priority_combo_radio.blockSignals(True)
            self.priority_small_qty_radio.blockSignals(True)
            self.priority_large_qty_radio.blockSignals(True)
            self.priority_no_qty_radio.blockSignals(True)
            self.priority_old_order_radio.blockSignals(True)
            self.priority_new_order_radio.blockSignals(True)
            self.priority_no_time_radio.blockSignals(True)
            
            self.priority_single_radio.setChecked(rules["single_first"])
            self.priority_combo_radio.setChecked(rules["combo_first"])
            self.priority_small_qty_radio.setChecked(rules["small_qty_first"])
            self.priority_large_qty_radio.setChecked(rules["large_qty_first"])
            # 수량 무관: 둘 다 False일 때
            if not rules["small_qty_first"] and not rules["large_qty_first"]:
                self.priority_no_qty_radio.setChecked(True)
            self.priority_old_order_radio.setChecked(rules["old_order_first"])
            self.priority_new_order_radio.setChecked(rules["new_order_first"])
            # 시간 무관: 둘 다 False일 때
            if not rules["old_order_first"] and not rules["new_order_first"]:
                self.priority_no_time_radio.setChecked(True)
            
            self.priority_single_radio.blockSignals(False)
            self.priority_combo_radio.blockSignals(False)
            self.priority_small_qty_radio.blockSignals(False)
            self.priority_large_qty_radio.blockSignals(False)
            self.priority_no_qty_radio.blockSignals(False)
            self.priority_old_order_radio.blockSignals(False)
            self.priority_new_order_radio.blockSignals(False)
            self.priority_no_time_radio.blockSignals(False)
        
        # 규칙 적용
        self._apply_priority_rules()
        
        # 프리셋 이름 매핑
        preset_names = {
            "default": "기본(단품 우선)",
            "backlog": "밀린 주문 정리",
            "bulk": "대량 소화"
        }
        self._add_log(f"프리셋 적용: {preset_names.get(preset_name, preset_name)}")
    
    def _apply_priority_rules(self):
        """현재 UI 설정을 기반으로 우선순위 규칙 적용"""
        # 라디오 버튼에서 값 읽기
        if hasattr(self, 'priority_single_radio'):
            rules = {
                "single_first": self.priority_single_radio.isChecked(),
                "combo_first": self.priority_combo_radio.isChecked(),
                "small_qty_first": self.priority_small_qty_radio.isChecked(),
                "large_qty_first": self.priority_large_qty_radio.isChecked(),
                "old_order_first": self.priority_old_order_radio.isChecked(),
                "new_order_first": self.priority_new_order_radio.isChecked(),
                "manual_priority": True  # ⭐ 고정 기능 항상 활성화
            }
        else:
            # 초기화 중일 때는 기본값 사용
            rules = {
                "single_first": True,
                "combo_first": False,
                "small_qty_first": False,
                "large_qty_first": False,
                "old_order_first": False,
                "new_order_first": False,
                "manual_priority": True
            }
        
        # processor에 규칙 전달
        self.processor.set_priority_rules(rules)
        
        # 로그 출력 (변경사항만, manual_priority 제외)
        # log_text가 초기화되지 않았을 수 있으므로 안전하게 처리
        if hasattr(self, 'log_text'):
            active_rules = [k for k, v in rules.items() if v and k != "manual_priority"]
            if active_rules:
                self._add_log(f"우선순위 규칙 적용: {', '.join(active_rules)}")
    
    def _on_toggle_tracking_priority(self, tracking_no: str, is_priority: bool):
        """
        송장 ⭐ 고정 상태 토글 (방식 A: 카드의 ⭐ 버튼)
        
        Args:
            tracking_no: 송장번호
            is_priority: True면 ⭐ 고정, False면 해제
        """
        self._set_tracking_priority(tracking_no, is_priority)
        
        # UI 업데이트 (⭐ 버튼 상태 및 목록 반영)
        self._update_summary_table()
        self._update_priority_tracking_list()
        
        # 로그 출력
        status = "⭐ 고정" if is_priority else "⭐ 해제"
        self._add_log(f"송장 {tracking_no} {status}")
    
    def _set_tracking_priority(self, tracking_no: str, is_priority: bool):
        """
        송장 ⭐ 고정 상태 설정 (공통 함수)
        
        Args:
            tracking_no: 송장번호
            is_priority: True면 ⭐ 고정, False면 해제
        """
        self.excel_loader.set_tracking_priority(tracking_no, is_priority)
        
        # 메타데이터 캐시 갱신 (다음 매칭부터 적용)
        if self.excel_loader._metadata_cache:
            # 해당 송장의 메타데이터만 갱신
            if tracking_no in self.excel_loader._metadata_cache:
                meta = self.excel_loader._metadata_cache[tracking_no]
                meta["is_priority"] = is_priority
    
    def _on_add_priority_tracking(self):
        """우선 송장 추가 (방식 B: 직접 입력)"""
        input_text = self.priority_tracking_input.text().strip()
        if not input_text:
            return
        
        # 여러 개 입력 지원: 줄바꿈 또는 쉼표로 구분
        tracking_nos = []
        for line in input_text.replace(',', '\n').split('\n'):
            tn = line.strip()
            if tn:
                tracking_nos.append(tn)
        
        if not tracking_nos:
            return
        
        # 각 송장번호 추가
        added_count = 0
        not_found = []
        
        for tracking_no in tracking_nos:
            # 송장번호 존재 확인
            if self.excel_loader.df is None:
                QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 불러오세요.")
                return
            
            # used=0인 송장만 확인 (처리되지 않은 송장)
            pending = self.excel_loader.df[self.excel_loader.df['used'] == 0]
            if tracking_no not in pending['tracking_no'].values:
                not_found.append(tracking_no)
                continue
            
            # 이미 우선 송장인지 확인
            if not self.excel_loader.get_tracking_priority(tracking_no):
                self._set_tracking_priority(tracking_no, True)
                added_count += 1
        
        # 입력창 초기화
        self.priority_tracking_input.clear()
        
        # 결과 메시지
        if added_count > 0:
            self._add_log(f"⭐ 우선 송장 {added_count}개 추가됨")
            self._update_priority_tracking_list()
            self._update_summary_table()
        
        if not_found:
            not_found_str = ', '.join(not_found[:5])
            if len(not_found) > 5:
                not_found_str += f" 외 {len(not_found) - 5}개"
            QMessageBox.warning(
                self, "경고",
                f"다음 송장번호를 찾을 수 없거나 이미 처리되었습니다:\n{not_found_str}"
            )
    
    def _on_remove_priority_tracking(self):
        """우선 송장 해제 (방식 B: 목록에서 선택 후 해제)"""
        selected_items = self.priority_tracking_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "알림", "해제할 송장을 선택하세요.")
            return
        
        removed_count = 0
        for item in selected_items:
            tracking_no = item.text()
            if self.excel_loader.get_tracking_priority(tracking_no):
                self._set_tracking_priority(tracking_no, False)
                removed_count += 1
        
        if removed_count > 0:
            self._add_log(f"⭐ 우선 송장 {removed_count}개 해제됨")
            self._update_priority_tracking_list()
            self._update_summary_table()
    
    def _update_priority_tracking_list(self):
        """우선 송장 목록 업데이트"""
        if not hasattr(self, 'priority_tracking_list'):
            return
        
        self.priority_tracking_list.clear()
        
        if self.excel_loader.df is None:
            return
        
        # 모든 우선 송장 조회
        all_tracking_nos = self.excel_loader.get_all_tracking_numbers()
        priority_tracking_nos = [
            tn for tn in all_tracking_nos
            if self.excel_loader.get_tracking_priority(tn)
        ]
        
        # 목록에 추가 (정렬)
        for tracking_no in sorted(priority_tracking_nos):
            item = QListWidgetItem(f"⭐ {tracking_no}")
            item.setData(Qt.UserRole, tracking_no)  # tracking_no 저장
            self.priority_tracking_list.addItem(item)
    
    @Slot()
    def _on_manual_scan(self):
        """수동 바코드 스캔"""
        barcode = self.manual_barcode_edit.text().strip()
        if barcode:
            # 스캐너 버퍼 클리어 (이중 처리 방지)
            self.scanner.clear_buffer()
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
            # 스캔 성공 시 카드 반짝임 효과
            QTimer.singleShot(100, lambda: self._highlight_scanned_cards(event.barcode))
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
    
    @Slot(str)
    def _on_priority_cleared(self, tracking_no: str):
        """완료된 우선 송장 자동 해제 (시그널 핸들러)"""
        self._add_log(f"완료된 우선 송장 자동 해제: {tracking_no}")
        # UI 업데이트 (우선 송장 목록 및 카드 ⭐ 상태)
        self._update_priority_tracking_list()
        self._update_summary_table()
    
    @Slot()
    def _on_data_loaded(self):
        """데이터 로드 완료"""
        self._update_tables()
        self._update_status_count()
        # 우선 송장 목록 업데이트
        self._update_priority_tracking_list()
    
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
            # 각 송장별로 별도 카드 생성 (⭐ 기능을 위해)
            tracking_groups = pending.groupby('tracking_no')
            combo_cards = []
            
            for tracking_no, group in tracking_groups:
                # 각 송장에 대한 카드 정보 생성
                combo_info = self._create_combo_info_for_tracking(tracking_no, group)
                combo_cards.append(combo_info)
            
            # ⭐ 고정 송장을 먼저 정렬 (우선순위 반영)
            combo_cards.sort(key=lambda x: (
                not self.excel_loader.get_tracking_priority(x['tracking_nos'][0]),  # ⭐ 고정이 먼저
                -x['count']  # 그 다음 개수 내림차순
            ))
            
            for combo_info in combo_cards:
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
            barcode = str(row['barcode']) if pd.notna(row['barcode']) else ''
            qty = int(row['qty']) if pd.notna(row['qty']) else 1
            scanned = int(row['scanned_qty']) if pd.notna(row['scanned_qty']) else 0
            remaining = qty - scanned
            
            key = f"{product_name}|{option_name}"
            if key not in product_summary:
                product_summary[key] = {
                    'product_name': product_name,
                    'option_name': option_name,
                    'barcode': barcode,
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
        
        # 바코드 정보 저장 (반짝임 효과용)
        card._barcode = prod_info.get('barcode', '')
        
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
        
        # 남은 수량 (4자리까지 표시)
        count_label = QLabel(f"<b style='color:{text_color}; font-size:14px;'>{remaining}</b>")
        count_label.setFixedWidth(50)
        count_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
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
    
    def _create_combo_info_for_tracking(self, tracking_no: str, group: pd.DataFrame) -> dict:
        """
        특정 송장에 대한 카드 정보 생성
        
        Args:
            tracking_no: 송장번호
            group: 해당 송장의 DataFrame 그룹
        
        Returns:
            카드 정보 딕셔너리
        """
        barcodes = sorted(group['barcode'].unique())
        products = []
        
        for _, row in group.iterrows():
            product_name = str(row['product_name']) if pd.notna(row['product_name']) else ''
            option_name = str(row['option_name']) if pd.notna(row['option_name']) else ''
            qty = int(row['qty']) if pd.notna(row['qty']) else 1
            
            product_info = product_name
            if option_name and option_name != 'nan':
                product_info += f" ({option_name})"
            
            # 수량 뒤에 표시: 1개, 2개, 3개...
            product_info += f" {qty}개"
            
            if product_info and product_info not in products:
                products.append(product_info)
        
        return {
            'count': 1,  # 송장당 1개
            'products': products,
            'barcodes': barcodes,
            'tracking_nos': [tracking_no]  # 단일 송장
        }
    
    def _get_summary_combo_data(self, pending):
        """구성별 데이터 추출 (수량 포함) - 기존 함수 유지 (다른 곳에서 사용 가능)"""
        tracking_groups = pending.groupby('tracking_no')
        combo_counts = {}
        
        for tracking_no, group in tracking_groups:
            barcodes = tuple(sorted(group['barcode'].unique()))
            
            if barcodes not in combo_counts:
                combo_counts[barcodes] = {
                    'count': 0,
                    'products': [],
                    'barcodes': list(barcodes),
                    'tracking_nos': []  # 같은 구성의 송장번호 리스트
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
            if tracking_no not in combo_counts[barcodes]['tracking_nos']:
                combo_counts[barcodes]['tracking_nos'].append(tracking_no)
        
        return sorted(combo_counts.values(), key=lambda x: -x['count'])
    
    def _create_summary_card(self, combo_info):
        """요약 카드 생성 (가로 나열, 전체 품목 표시) + ⭐ 토글 버튼"""
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        
        # 바코드 정보 저장 (반짝임 효과용)
        card._barcodes = combo_info.get('barcodes', [])
        # tracking_no 리스트 저장 (⭐ 토글용)
        card._tracking_nos = combo_info.get('tracking_nos', [])
        
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
        
        # ⭐ 토글 버튼 (여러 송장이 있으면 첫 번째 송장 기준)
        # 실제로는 각 송장별로 별도 카드가 생성되므로 첫 번째 송장만 사용
        if card._tracking_nos:
            tracking_no = card._tracking_nos[0]
            is_priority = self.excel_loader.get_tracking_priority(tracking_no)
            
            star_btn = QPushButton("⭐" if is_priority else "☆")
            star_btn.setCheckable(True)
            star_btn.setChecked(is_priority)
            star_btn.setMaximumWidth(30)
            star_btn.setMaximumHeight(30)
            star_btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    background-color: transparent;
                    font-size: 16px;
                }
                QPushButton:checked {
                    color: #FFD700;
                }
            """)
            star_btn.clicked.connect(lambda checked, tn=tracking_no: self._on_toggle_tracking_priority(tn, checked))
            layout.addWidget(star_btn)
        
        return card
    
    def _flash_card(self, card: QFrame, flash_color: str = "#FFEB3B"):
        """카드 반짝임 효과"""
        if not card:
            return
        
        # 원래 스타일 저장
        original_style = card.styleSheet()
        
        # 반짝임 색상으로 변경
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {flash_color};
                border: 3px solid #FFC107;
                border-radius: 8px;
                padding: 6px 10px;
                margin: 2px;
            }}
        """)
        
        # 0.3초 후 원래 스타일로 복원
        QTimer.singleShot(300, lambda: card.setStyleSheet(original_style))
    
    def _highlight_scanned_cards(self, barcode: str):
        """스캔된 바코드에 해당하는 카드들 반짝임"""
        # 구성별 카드에서 찾기
        for i in range(self.summary_grid.count()):
            item = self.summary_grid.itemAt(i)
            if item and item.widget():
                card = item.widget()
                if hasattr(card, '_barcodes') and barcode in card._barcodes:
                    self._flash_card(card)
        
        # 제품별 카드에서 찾기
        for i in range(self.product_grid.count()):
            item = self.product_grid.itemAt(i)
            if item and item.widget():
                card = item.widget()
                if hasattr(card, '_barcode') and card._barcode == barcode:
                    self._flash_card(card, "#4CAF50")  # 녹색 반짝임
    
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
        # log_text가 초기화되지 않았으면 무시
        if not hasattr(self, 'log_text') or self.log_text is None:
            return
        
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
                success, saved_path = self.excel_loader.save_excel()
                if success:
                    self._add_log(f"종료 시 저장 완료: {saved_path}")
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

