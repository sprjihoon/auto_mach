"""
EzAuto 자동입력 (pyautogui)
"""
import time
import pyautogui
import pygetwindow as gw
from typing import Optional
from PySide6.QtCore import QObject, Signal

# pyautogui 안전 설정
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


class EzAutoInput(QObject):
    """EzAuto 프로그램 자동 입력 클래스"""
    
    # 시그널
    input_success = Signal(str)  # 성공 메시지
    input_error = Signal(str)    # 오류 메시지
    
    def __init__(self):
        super().__init__()
        self._enabled = True
        self._typing_interval = 0.02  # 타이핑 간격 (초)
        self._delay_after_tracking = 0.8  # tracking_no 입력 후 대기 시간 (늘림)
        self._delay_after_barcode = 0.3   # barcode 입력 후 대기 시간
        self._window_title = "이지오토"  # EzAuto 창 제목 (부분 매칭)
    
    @property
    def enabled(self) -> bool:
        return self._enabled
    
    @enabled.setter
    def enabled(self, value: bool):
        self._enabled = value
    
    def set_window_title(self, title: str):
        """EzAuto 창 제목 설정"""
        self._window_title = title
    
    def set_delays(self, after_tracking: float = 0.3, after_barcode: float = 0.1):
        """대기 시간 설정"""
        self._delay_after_tracking = after_tracking
        self._delay_after_barcode = after_barcode
    
    def find_and_activate_ezauto(self) -> bool:
        """EzAuto 창을 찾아서 활성화"""
        try:
            # 창 제목에 EzAuto가 포함된 창 찾기
            windows = gw.getWindowsWithTitle(self._window_title)
            
            if not windows:
                # 대소문자 구분 없이 재시도
                all_windows = gw.getAllWindows()
                for win in all_windows:
                    if self._window_title.lower() in win.title.lower():
                        windows = [win]
                        break
            
            if windows:
                win = windows[0]
                # 최소화되어 있으면 복원
                if win.isMinimized:
                    win.restore()
                # 창 활성화
                win.activate()
                time.sleep(0.1)  # 활성화 대기
                return True
            else:
                self.input_error.emit(f"'{self._window_title}' 창을 찾을 수 없습니다")
                return False
                
        except Exception as e:
            self.input_error.emit(f"창 활성화 오류: {str(e)}")
            return False
    
    def send_input(self, tracking_no: str, barcode: str) -> bool:
        """
        EzAuto에 입력 전송
        순서: 창 활성화 → tracking_no → Enter → 대기 → barcode → Enter
        """
        if not self._enabled:
            self.input_error.emit("EzAuto 입력이 비활성화되어 있습니다")
            return False
        
        try:
            # 0. EzAuto 창 찾아서 활성화
            if not self.find_and_activate_ezauto():
                return False
            
            # 1. tracking_no 입력
            pyautogui.typewrite(tracking_no, interval=self._typing_interval)
            pyautogui.press('enter')
            
            # 2. 잠시 대기
            time.sleep(self._delay_after_tracking)
            
            # 3. barcode 입력
            pyautogui.typewrite(barcode, interval=self._typing_interval)
            pyautogui.press('enter')
            
            # 4. 완료 대기
            time.sleep(self._delay_after_barcode)
            
            self.input_success.emit(f"EzAuto 입력 완료: {tracking_no} / {barcode}")
            return True
            
        except pyautogui.FailSafeException:
            self.input_error.emit("안전 모드 발동: 마우스가 화면 모서리로 이동됨")
            return False
            
        except Exception as e:
            self.input_error.emit(f"EzAuto 입력 오류: {str(e)}")
            return False
    
    def send_tracking_only(self, tracking_no: str) -> bool:
        """tracking_no만 입력"""
        if not self._enabled:
            return False
        
        try:
            pyautogui.typewrite(tracking_no, interval=self._typing_interval)
            pyautogui.press('enter')
            time.sleep(self._delay_after_tracking)
            return True
            
        except Exception as e:
            self.input_error.emit(f"tracking_no 입력 오류: {str(e)}")
            return False
    
    def send_barcode_only(self, barcode: str) -> bool:
        """barcode만 입력"""
        if not self._enabled:
            return False
        
        try:
            pyautogui.typewrite(barcode, interval=self._typing_interval)
            pyautogui.press('enter')
            time.sleep(self._delay_after_barcode)
            return True
            
        except Exception as e:
            self.input_error.emit(f"barcode 입력 오류: {str(e)}")
            return False


class EzAutoInputAsync(EzAutoInput):
    """비동기 EzAuto 입력 (별도 스레드에서 실행)"""
    
    def __init__(self):
        super().__init__()
        self._is_busy = False
    
    @property
    def is_busy(self) -> bool:
        return self._is_busy
    
    def send_input_async(self, tracking_no: str, barcode: str):
        """비동기로 입력 전송 (스레드에서 호출)"""
        import threading
        
        if self._is_busy:
            self.input_error.emit("이전 입력이 진행 중입니다")
            return
        
        def _run():
            self._is_busy = True
            try:
                self.send_input(tracking_no, barcode)
            finally:
                self._is_busy = False
        
        thread = threading.Thread(target=_run, daemon=True)
        thread.start()

