"""
주문 처리 로직
qty/scanned_qty 처리, 우선순위 정렬 로직
"""
from typing import Optional, Tuple
from PySide6.QtCore import QObject, Signal
import pandas as pd
import winsound
import threading

from models import ScanResult, ScanEvent
from excel_loader import ExcelLoader
from ezauto_input import EzAutoInput
from pdf_printer import PDFPrinter
from utils import get_timestamp, sanitize_barcode


def play_scan_sound():
    """스캔 성공 신호음 (짧은 비프)"""
    def _play():
        winsound.Beep(1000, 100)  # 1000Hz, 100ms
    threading.Thread(target=_play, daemon=True).start()


def play_complete_sound():
    """송장 완료 신호음 (멜로디)"""
    def _play():
        winsound.Beep(800, 150)   # 낮은 음
        winsound.Beep(1000, 150)  # 중간 음
        winsound.Beep(1200, 200)  # 높은 음
    threading.Thread(target=_play, daemon=True).start()


def play_error_sound():
    """오류 신호음"""
    def _play():
        winsound.Beep(300, 300)  # 낮은 음, 긴 소리
    threading.Thread(target=_play, daemon=True).start()


class OrderProcessor(QObject):
    """주문 처리 핵심 로직"""
    
    # 시그널
    scan_processed = Signal(object)  # ScanEvent
    tracking_completed = Signal(str)  # tracking_no
    ui_update_required = Signal()
    log_message = Signal(str)  # 로그 메시지
    scanner_pause = Signal()  # 스캐너 일시 중지
    scanner_resume = Signal()  # 스캐너 재개
    
    def __init__(
        self,
        excel_loader: ExcelLoader,
        ezauto_input: EzAutoInput,
        pdf_printer: PDFPrinter
    ):
        super().__init__()
        self.excel = excel_loader
        self.ezauto = ezauto_input
        self.pdf = pdf_printer
        
        # 현재 작업 중인 tracking_no
        self._current_tracking_no: Optional[str] = None
        
        # 처리 중 플래그 (재스캔 방지)
        self._is_processing: bool = False
        self._last_barcode: str = ""
        self._last_scan_time: float = 0
    
    @property
    def current_tracking_no(self) -> Optional[str]:
        return self._current_tracking_no
    
    def process_scan(self, barcode: str) -> ScanEvent:
        """
        바코드 스캔 처리 메인 로직
        
        1) 바코드 스캔 감지
        2) (barcode == 입력값) AND (used == 0) 조건으로 후보 행 조회
        3) ORDER BY qty ASC, tracking_no ASC 정렬
        4) candidates.iloc[0] 선택
        5) scanned_qty += 1
        6) remaining == 0 이면 PDF 출력, used = 1
        """
        import time as time_module
        
        barcode = sanitize_barcode(barcode)
        timestamp = get_timestamp()
        current_time = time_module.time()
        
        # 같은 바코드 0.5초 내 재스캔 방지 (스캐너 더블 스캔 방지용)
        if barcode == self._last_barcode and (current_time - self._last_scan_time) < 0.5:
            self.log_message.emit(f"[무시] 더블 스캔 방지: {barcode}")
            return None
        
        self._last_barcode = barcode
        self._last_scan_time = current_time
        
        # 송장번호 형식 감지 (609로 시작하는 13자리 숫자) → 무시
        if barcode.startswith('609') and len(barcode) == 13 and barcode.isdigit():
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=None,
                result=ScanResult.NOT_FOUND,
                message=f"송장번호 스캔 무시: {barcode}"
            )
            self.log_message.emit(f"[정보] 송장번호 스캔 무시: {barcode}")
            return event
        
        self.log_message.emit(f"바코드 스캔: {barcode}")
        
        # 1. 현재 작업 중인 송장이 있으면 그 송장에서만 찾기
        if self._current_tracking_no:
            current_group = self.excel.get_tracking_group(self._current_tracking_no)
            current_match = current_group[
                (current_group['barcode'].astype(str).str.strip() == barcode) & 
                (current_group['scanned_qty'] < current_group['qty'])
            ]
            
            if not current_match.empty:
                # 현재 송장에서 해당 바코드 처리
                candidates = current_match.reset_index(drop=False)
                self.log_message.emit(f"[디버그] 현재 송장 {self._current_tracking_no}에서 처리")
            else:
                # 현재 송장에 해당 바코드 없음 → 경고음 + 무시
                play_error_sound()  # 경고음 🚨
                
                event = ScanEvent(
                    timestamp=timestamp,
                    barcode=barcode,
                    tracking_no=self._current_tracking_no,
                    result=ScanResult.NOT_FOUND,
                    message=f"⚠️ 현재 송장({self._current_tracking_no})에 '{barcode}' 없음!"
                )
                self.scan_processed.emit(event)
                self.log_message.emit(f"[경고] {event.message}")
                return event
        else:
            # 새 송장 검색 (우선순위 정렬됨)
            try:
                candidates = self.excel.find_candidates(barcode)
                self.log_message.emit(f"[디버그] 후보 {len(candidates)}건 찾음")
            except Exception as e:
                self.log_message.emit(f"[오류] 후보 검색 실패: {str(e)}")
                candidates = None
        
        if candidates is None or candidates.empty:
            # 바코드 없음 → 경고음
            play_error_sound()  # 경고음 🚨
            
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=None,
                result=ScanResult.NOT_FOUND,
                message=f"⚠️ 바코드 '{barcode}'를 찾을 수 없습니다"
            )
            self.scan_processed.emit(event)
            self.log_message.emit(f"[경고] {event.message}")
            return event
        
        # 2. 첫 번째 후보 선택 (qty 가장 작고, tracking_no 오름차순)
        selected = candidates.iloc[0]
        tracking_no = str(selected['tracking_no'])
        original_index = selected['index']  # 원본 DataFrame 인덱스
        
        # 3. 이미 사용된 송장인지 확인
        if self.excel.is_tracking_used(tracking_no):
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=tracking_no,
                result=ScanResult.ALREADY_USED,
                message=f"이미 처리된 송장입니다: {tracking_no}"
            )
            self.scan_processed.emit(event)
            self.log_message.emit(f"[경고] {event.message}")
            return event
        
        # 4. scanned_qty 증가
        if not self.excel.increment_scanned(original_index):
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=tracking_no,
                result=ScanResult.ERROR,
                message=f"스캔 수량 업데이트 실패"
            )
            self.scan_processed.emit(event)
            self.log_message.emit(f"[오류] {event.message}")
            return event
        
        # 5. EzAuto 입력 전송 (같은 송장이면 바코드만)
        is_new_tracking = (self._current_tracking_no != tracking_no)
        
        # 처리 시작
        self._is_processing = True
        
        # 스캐너 일시 중지 (EzAuto 입력 중 키 입력 방지)
        self.scanner_pause.emit()
        
        if is_new_tracking:
            # 새 송장: 송장번호 + 바코드 입력
            self._current_tracking_no = tracking_no
            self.ezauto.send_input(tracking_no, barcode)
            self.log_message.emit(f"[EzAuto] 송장번호 + 바코드 입력: {tracking_no} / {barcode}")
            
            # 새 송장 첫 스캔 시 PDF 즉시 출력!
            self.log_message.emit(f"[출력] 송장 {tracking_no} PDF 출력 시작")
            if self.pdf.print_pdf(tracking_no):
                self.log_message.emit(f"[성공] PDF 출력 완료: {tracking_no}")
            else:
                self.log_message.emit(f"[오류] PDF 출력 실패: {tracking_no}")
        else:
            # 같은 송장: 바코드만 입력
            self.ezauto.send_barcode_only(barcode)
            self.log_message.emit(f"[EzAuto] 바코드만 입력: {barcode}")
        
        # 스캐너 재개 전 대기 (입력 완료 후 안정화)
        import time as time_mod
        time_mod.sleep(0.5)
        
        # 스캐너 재개
        self.scanner_resume.emit()
        
        # 7. 남은 수량 계산
        remaining = self.excel.get_group_remaining(tracking_no)
        
        # 8. UI 업데이트 요청
        self.ui_update_required.emit()
        
        # 9. 완료 확인
        if remaining == 0:
            # 송장 완료! (PDF는 이미 첫 스캔 시 출력됨)
            self.log_message.emit(f"[완료] 송장 {tracking_no} 구성 완료!")
            
            # 완료 신호음 🎵
            play_complete_sound()
            
            # used = 1 설정
            self.excel.mark_used(tracking_no)
            self.log_message.emit(f"[완료] 송장 {tracking_no} 처리 완료 (used=1)")
            
            # 스캐너 일시 중지 (다음 송장 자동 시작 방지)
            self.scanner_pause.emit()
            
            # 완료 시그널
            self.tracking_completed.emit(tracking_no)
            self._current_tracking_no = None
            
            # 1초 대기 후 스캐너 재개
            import time as time_mod
            time_mod.sleep(1.0)
            self.scanner_resume.emit()
            self.log_message.emit("[정보] 다음 송장 스캔 준비 완료")
            
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=tracking_no,
                result=ScanResult.SUCCESS,
                message=f"송장 {tracking_no} 구성 완료!"
            )
        else:
            # 스캔 성공 신호음 🔔
            play_scan_sound()
            
            event = ScanEvent(
                timestamp=timestamp,
                barcode=barcode,
                tracking_no=tracking_no,
                result=ScanResult.SUCCESS,
                message=f"스캔 성공 (남은 수량: {remaining})"
            )
        
        # 처리 완료
        self._is_processing = False
        
        self.scan_processed.emit(event)
        self.log_message.emit(f"[정보] {event.message}")
        return event
    
    def get_current_tracking_items(self) -> pd.DataFrame:
        """현재 작업 중인 tracking_no의 항목들 반환"""
        if not self._current_tracking_no:
            return pd.DataFrame()
        return self.excel.get_tracking_group(self._current_tracking_no)
    
    def get_pending_summary(self) -> pd.DataFrame:
        """미처리 항목 요약"""
        return self.excel.get_summary_by_barcode()
    
    def reset_current_tracking(self):
        """현재 tracking_no 초기화"""
        self._current_tracking_no = None
        self.ui_update_required.emit()

