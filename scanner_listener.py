"""
스캐너(HID) 입력 후킹
keyboard 모듈을 사용하여 글로벌 키보드 입력 감지
"""
import threading
from typing import Callable, Optional
from PySide6.QtCore import QObject, Signal
import keyboard


class ScannerListener(QObject):
    """바코드 스캐너 입력 리스너"""
    
    # 바코드 스캔 완료 시그널
    barcode_scanned = Signal(str)
    # 상태 변경 시그널
    status_changed = Signal(str)
    
    # 스캐너 입력 속도 임계값 (ms) - 이보다 느리면 사람 타이핑으로 간주
    SCAN_SPEED_THRESHOLD = 50  # 50ms
    # 최소 바코드 길이
    MIN_BARCODE_LENGTH = 4
    
    def __init__(self):
        super().__init__()
        self._buffer: str = ""
        self._is_running: bool = False
        self._hook = None
        self._lock = threading.Lock()
        self._last_key_time: float = 0
        self._is_fast_input: bool = False  # 빠른 입력 모드 (스캐너)
    
    def start(self) -> bool:
        """스캐너 리스닝 시작"""
        if self._is_running:
            return True
        
        try:
            self._is_running = True
            self._buffer = ""
            
            # 키보드 훅 등록
            keyboard.on_press(self._on_key_press)
            
            self.status_changed.emit("스캐너 리스닝 시작됨")
            return True
            
        except Exception as e:
            self._is_running = False
            self.status_changed.emit(f"스캐너 시작 실패: {str(e)}")
            return False
    
    def stop(self):
        """스캐너 리스닝 중지"""
        if not self._is_running:
            return
        
        try:
            keyboard.unhook_all()
            self._is_running = False
            self._buffer = ""
            self.status_changed.emit("스캐너 리스닝 중지됨")
            
        except Exception as e:
            self.status_changed.emit(f"스캐너 중지 오류: {str(e)}")
    
    def _on_key_press(self, event: keyboard.KeyboardEvent):
        """키 입력 이벤트 핸들러 (스캐너 입력 속도 필터링)"""
        if not self._is_running:
            return
        
        import time
        current_time = time.time() * 1000  # ms
        
        with self._lock:
            key_name = event.name
            
            # 입력 속도 체크
            time_diff = current_time - self._last_key_time
            self._last_key_time = current_time
            
            if key_name == 'enter':
                # Enter 키: 버퍼의 내용을 바코드로 처리
                if self._buffer:
                    barcode = self._buffer.strip()
                    self._buffer = ""
                    self._is_fast_input = False
                    
                    # 최소 길이 확인 및 빠른 입력이었는지 확인
                    if barcode and len(barcode) >= self.MIN_BARCODE_LENGTH:
                        # 시그널 발생 (메인 스레드에서 처리됨)
                        self.barcode_scanned.emit(barcode)
            
            elif key_name == 'backspace':
                # Backspace: 버퍼 초기화 (사람 입력으로 간주)
                self._buffer = ""
                self._is_fast_input = False
            
            elif key_name == 'space':
                # 스페이스: 느린 입력이면 버퍼 초기화
                if time_diff > self.SCAN_SPEED_THRESHOLD * 2:
                    self._buffer = ""
                else:
                    self._buffer += ' '
            
            elif len(key_name) == 1:
                # 일반 문자
                if len(self._buffer) == 0:
                    # 첫 글자: 버퍼 시작
                    self._buffer = key_name
                    self._is_fast_input = True
                elif time_diff <= self.SCAN_SPEED_THRESHOLD:
                    # 빠른 입력: 스캐너로 간주하고 버퍼에 추가
                    self._buffer += key_name
                else:
                    # 느린 입력: 사람 타이핑으로 간주하고 버퍼 초기화 후 새로 시작
                    self._buffer = key_name
                    self._is_fast_input = False
            
            elif key_name.startswith('shift'):
                # Shift 키는 무시
                pass
    
    def clear_buffer(self):
        """버퍼 초기화"""
        with self._lock:
            self._buffer = ""
    
    @property
    def is_running(self) -> bool:
        """실행 상태 반환"""
        return self._is_running
    
    @property
    def current_buffer(self) -> str:
        """현재 버퍼 내용 반환"""
        with self._lock:
            return self._buffer


class ManualScannerInput(QObject):
    """수동 바코드 입력 (UI 입력 필드용)"""
    
    barcode_scanned = Signal(str)
    
    def submit_barcode(self, barcode: str):
        """바코드 수동 제출"""
        barcode = barcode.strip()
        if barcode:
            self.barcode_scanned.emit(barcode)

