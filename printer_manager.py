"""
프린터 관리 모듈
Windows 프린터 목록 조회, 설정 저장/로드, PDF 출력 기능 제공
"""
import os
import json
from pathlib import Path
from typing import Optional, List, Dict
import tempfile

# win32api, win32print는 선택적 (pywin32 설치 시에만 사용)
try:
    import win32api
    import win32print
    HAS_WIN32API = True
except ImportError:
    HAS_WIN32API = False


def get_settings_path() -> Path:
    """설정 파일 경로 반환"""
    base_path = Path(__file__).parent
    return base_path / "settings.json"


def get_printers() -> List[str]:
    """
    Windows에 설치된 프린터 목록 반환
    
    Returns:
        프린터 이름 리스트
    """
    printers = []
    if not HAS_WIN32API:
        return printers
    
    try:
        # 로컬 및 네트워크 프린터 모두 조회
        printer_info = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        printers = [info[2] for info in printer_info]
    except Exception as e:
        print(f"프린터 목록 조회 오류: {str(e)}")
    
    return printers


def save_printer_settings(label_printer: Optional[str] = None, a4_printer: Optional[str] = None) -> bool:
    """
    settings.json에 두 프린터 이름 저장
    
    Args:
        label_printer: 라벨 프린터 이름
        a4_printer: A4 프린터 이름
    
    Returns:
        저장 성공 여부
    """
    settings_path = get_settings_path()
    
    # 기존 설정 로드
    settings = {}
    if settings_path.exists():
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception:
            settings = {}
    
    # 새 설정 업데이트
    if label_printer is not None:
        settings["label_printer"] = label_printer
    if a4_printer is not None:
        settings["a4_printer"] = a4_printer
    
    # 저장
    try:
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"설정 저장 오류: {str(e)}")
        return False


def load_printer_settings() -> Dict[str, Optional[str]]:
    """
    settings.json에서 프린터 이름 로드
    
    Returns:
        {"label_printer": str or None, "a4_printer": str or None}
    """
    settings_path = get_settings_path()
    
    if not settings_path.exists():
        return {"label_printer": None, "a4_printer": None}
    
    try:
        with open(settings_path, 'r', encoding='utf-8') as f:
            settings = json.load(f)
        
        return {
            "label_printer": settings.get("label_printer"),
            "a4_printer": settings.get("a4_printer")
        }
    except Exception as e:
        print(f"설정 로드 오류: {str(e)}")
        return {"label_printer": None, "a4_printer": None}


def print_pdf_with_printer(pdf_path: str, printer_name: Optional[str] = None) -> bool:
    """
    지정된 프린터로 PDF 출력
    
    Args:
        pdf_path: 출력할 PDF 파일 경로
        printer_name: 프린터 이름 (None이면 기본 프린터 사용)
    
    Returns:
        출력 성공 여부
    """
    if not os.path.exists(pdf_path):
        print(f"PDF 파일 없음: {pdf_path}")
        return False
    
    if not HAS_WIN32API:
        # win32api가 없으면 기본 방법 사용
        try:
            import subprocess
            os.startfile(pdf_path, "print")
            return True
        except Exception as e:
            print(f"PDF 출력 오류: {str(e)}")
            return False
    
    try:
        # 기본 프린터 백업
        original_default = None
        try:
            original_default = win32print.GetDefaultPrinter()
        except Exception:
            pass
        
        # 프린터 이름이 지정된 경우 기본 프린터로 임시 설정
        if printer_name:
            try:
                # 프린터 존재 확인
                printers = get_printers()
                if printer_name not in printers:
                    print(f"프린터를 찾을 수 없습니다: {printer_name}")
                    return False
                
                # 기본 프린터로 설정
                win32print.SetDefaultPrinter(printer_name)
            except Exception as e:
                print(f"프린터 설정 오류: {str(e)}")
                return False
        
        # PDF 출력
        try:
            win32api.ShellExecute(
                0,
                "print",
                pdf_path,
                None,
                ".",
                0
            )
            
            # 기본 프린터 복원
            if original_default and printer_name:
                try:
                    win32print.SetDefaultPrinter(original_default)
                except Exception:
                    pass
            
            return True
        except Exception as e:
            print(f"PDF 출력 오류: {str(e)}")
            
            # 기본 프린터 복원
            if original_default and printer_name:
                try:
                    win32print.SetDefaultPrinter(original_default)
                except Exception:
                    pass
            
            return False
            
    except Exception as e:
        print(f"프린터 출력 오류: {str(e)}")
        return False


def check_printer_exists(printer_name: str) -> bool:
    """
    프린터가 시스템에 존재하는지 확인
    
    Args:
        printer_name: 확인할 프린터 이름
    
    Returns:
        존재 여부
    """
    if not printer_name:
        return False
    
    printers = get_printers()
    return printer_name in printers

