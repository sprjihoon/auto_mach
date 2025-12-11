"""
공통 유틸리티 함수
"""
import os
import sys
from datetime import datetime
from pathlib import Path


def get_timestamp() -> str:
    """현재 타임스탬프 반환"""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_base_path() -> Path:
    """실행 파일 기준 경로 반환 (PyInstaller 호환)"""
    if getattr(sys, 'frozen', False):
        # PyInstaller로 빌드된 경우
        return Path(sys.executable).parent
    else:
        # 개발 환경
        return Path(__file__).parent


def get_labels_path() -> Path:
    """라벨 PDF 폴더 경로"""
    labels_dir = get_base_path() / "labels"
    labels_dir.mkdir(exist_ok=True)
    return labels_dir


def get_pdf_path(tracking_no: str) -> Path:
    """송장번호로 PDF 파일 경로 반환"""
    return get_labels_path() / f"{tracking_no}.pdf"


def pdf_exists(tracking_no: str) -> bool:
    """PDF 파일 존재 여부 확인"""
    return get_pdf_path(tracking_no).exists()


def format_log_message(level: str, message: str) -> str:
    """로그 메시지 포맷팅"""
    timestamp = get_timestamp()
    return f"[{timestamp}] [{level}] {message}"


def sanitize_barcode(barcode: str) -> str:
    """바코드 문자열 정리"""
    # 앞뒤 공백 제거, 특수문자 정리
    return barcode.strip().replace('\r', '').replace('\n', '')

