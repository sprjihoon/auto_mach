"""
PDF 파일 검색 모듈 (재출력용)
멀티코어를 활용한 고속 PDF 검색 및 송장번호/주문번호 추출
"""
import os
import re
import sys
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing

# Windows에서 멀티프로세싱 문제 방지 (ThreadPoolExecutor 사용)
if sys.platform == "win32":
    USE_THREADS = True
else:
    USE_THREADS = False

# PDF 처리 라이브러리
try:
    import pdfplumber
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


def _extract_text_from_pdf(pdf_path: Path) -> Tuple[str, Dict[str, str]]:
    """
    PDF에서 텍스트 추출 및 송장번호/주문번호 검색
    
    Args:
        pdf_path: PDF 파일 경로
    
    Returns:
        (추출된_텍스트, {송장번호: 원본형식, 주문번호: 원본형식})
    """
    found_numbers = {}
    
    if not PDF_SUPPORT:
        return "", found_numbers
    
    # 송장번호 패턴 (5-4-4 형식, 11-13자리 숫자)
    # 하이픈, 공백, 다양한 하이픈 문자 지원
    tracking_patterns = [
        r'(\d{5}[-–—\s]+\d{4}[-–—\s]+\d{4})',  # 60914-8682-2638 형식 (하이픈/공백 유연)
        r'(\d{5}[-–—]\d{4}[-–—]\d{4})',  # 하이픈만
        r'(\d{5}\s+\d{4}\s+\d{4})',  # 공백만
        r'등기번호[:\s\-]*(\d{5}[-–—\s]?\d{4}[-–—\s]?\d{4})',  # "등기번호: 60914-8682-2638"
        r'송장번호[:\s\-]*(\d{5}[-–—\s]?\d{4}[-–—\s]?\d{4})',  # "송장번호: 60914-8682-2638"
        r'\b(\d{13})\b',  # 13자리 연속
        r'\b(\d{12})\b',  # 12자리 연속
        r'\b(\d{11})\b',  # 11자리 연속
    ]
    
    # 주문번호 패턴 (날짜-순서 형식 등)
    order_patterns = [
        r'\b(\d{8}[-–—]\d{7})\b',  # 20251212-0000051 형식
        r'\b(\d{8}-\d{7})\b',  # 일반 하이픈
        r'주문번호[:\s]*(\d{8}[-–—]?\d{7})',  # "주문번호: 20251212-0000051"
        r'주문번호[:\s]*(\d{8}-\d{7})',  # 일반 하이픈
    ]
    
    all_text = ""
    found_tracking = set()
    found_order = set()
    
    # 방법 1: pdfplumber로 시도 (가장 정확)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # 표준 텍스트 추출
                text = page.extract_text() or ""
                
                # 고정밀 텍스트 추출 옵션 시도
                if not text or len(text.strip()) < 10:
                    extraction_methods = [
                        {"x_tolerance": 1, "y_tolerance": 1, "layout": True},
                        {"x_tolerance": 3, "y_tolerance": 3, "layout": True},
                        {"x_tolerance": 5, "y_tolerance": 5, "layout": False},
                        {"x_tolerance": 2, "y_tolerance": 2, "layout": True, "x_density": 10, "y_density": 10},
                        {"use_text_flow": True, "layout": True},
                    ]
                    
                    for method in extraction_methods:
                        try:
                            text = page.extract_text(**method) or ""
                            if text and len(text.strip()) >= 10:
                                break
                        except:
                            continue
                
                if text:
                    all_text += text + "\n"
                    
                    # 송장번호 검색
                    for pattern in tracking_patterns:
                        matches = re.findall(pattern, text)
                        for match in matches:
                            clean = re.sub(r'[-–—\s]', '', str(match))
                            if clean.isdigit() and len(clean) >= 11:
                                if clean not in found_tracking:
                                    found_tracking.add(clean)
                                    found_numbers[f"tracking_{clean}"] = match
                    
                    # 주문번호 검색
                    for pattern in order_patterns:
                        matches = re.findall(pattern, text)
                        for match in matches:
                            clean = re.sub(r'[-–—\s]', '', str(match))
                            if clean.isdigit() and len(clean) >= 8:
                                if clean not in found_order:
                                    found_order.add(clean)
                                    found_numbers[f"order_{clean}"] = match
                    
                    # 첫 번째 매칭에서 멈춤 (성능 최적화)
                    if found_tracking or found_order:
                        break
    except Exception:
        pass
    
    # 방법 2: PyMuPDF로 시도 (pdfplumber 실패 시)
    if not all_text or (not found_tracking and not found_order):
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text() or ""
                
                if text:
                    all_text += text + "\n"
                    
                    # 송장번호 검색
                    for pattern in tracking_patterns:
                        matches = re.findall(pattern, text)
                        for match in matches:
                            clean = re.sub(r'[-–—\s]', '', str(match))
                            if clean.isdigit() and len(clean) >= 11:
                                if clean not in found_tracking:
                                    found_tracking.add(clean)
                                    found_numbers[f"tracking_{clean}"] = match
                    
                    # 주문번호 검색
                    for pattern in order_patterns:
                        matches = re.findall(pattern, text)
                        for match in matches:
                            clean = re.sub(r'[-–—\s]', '', str(match))
                            if clean.isdigit() and len(clean) >= 8:
                                if clean not in found_order:
                                    found_order.add(clean)
                                    found_numbers[f"order_{clean}"] = match
                    
                    # 첫 번째 매칭에서 멈춤
                    if found_tracking or found_order:
                        break
            
            doc.close()
        except Exception:
            pass
    
    return all_text, found_numbers


def _search_pdf_file(args: Tuple[Path, str]) -> Optional[Dict]:
    """
    단일 PDF 파일 검색 (멀티프로세싱용)
    
    Args:
        args: (pdf_path, search_value) 튜플
    
    Returns:
        매칭 정보 또는 None
    """
    pdf_path, search_value = args
    
    try:
        text, found_numbers = _extract_text_from_pdf(pdf_path)
        
        if not found_numbers:
            return None
        
        # 검색값 정규화 (하이픈, 공백 제거)
        search_clean = re.sub(r'[-–—\s]', '', str(search_value).strip())
        search_original = str(search_value).strip()
        
        # 1. 정확한 매칭 (하이픈 제거 후 비교) - 가장 빠름
        for key, original in found_numbers.items():
            if key.startswith("tracking_"):
                tracking_clean = key.replace("tracking_", "")
                if tracking_clean == search_clean:
                    return {
                        "pdf_path": str(pdf_path),
                        "tracking_no": tracking_clean,
                        "original": original,
                        "type": "tracking"
                    }
        
        for key, original in found_numbers.items():
            if key.startswith("order_"):
                order_clean = key.replace("order_", "")
                if order_clean == search_clean:
                    return {
                        "pdf_path": str(pdf_path),
                        "order_no": order_clean,
                        "original": original,
                        "type": "order"
                    }
        
        # 2. 원본 형식 직접 비교 (하이픈 포함)
        for key, original in found_numbers.items():
            if key.startswith("tracking_"):
                original_str = str(original).strip()
                # 원본 형식 직접 비교
                if original_str == search_original:
                    return {
                        "pdf_path": str(pdf_path),
                        "tracking_no": key.replace("tracking_", ""),
                        "original": original,
                        "type": "tracking"
                    }
                # 원본 형식 정규화 후 비교
                original_clean = re.sub(r'[-–—\s]', '', original_str)
                if original_clean == search_clean:
                    return {
                        "pdf_path": str(pdf_path),
                        "tracking_no": key.replace("tracking_", ""),
                        "original": original,
                        "type": "tracking"
                    }
        
        for key, original in found_numbers.items():
            if key.startswith("order_"):
                original_str = str(original).strip()
                # 원본 형식 직접 비교
                if original_str == search_original:
                    return {
                        "pdf_path": str(pdf_path),
                        "order_no": key.replace("order_", ""),
                        "original": original,
                        "type": "order"
                    }
                # 원본 형식 정규화 후 비교
                original_clean = re.sub(r'[-–—\s]', '', original_str)
                if original_clean == search_clean:
                    return {
                        "pdf_path": str(pdf_path),
                        "order_no": key.replace("order_", ""),
                        "original": original,
                        "type": "order"
                    }
        
        # 3. 텍스트에서 직접 검색 (하이픈 포함/제거 모두 시도)
        if search_clean in text.replace("-", "").replace("–", "").replace("—", "").replace(" ", ""):
            # 텍스트에서 직접 패턴 매칭 시도
            for pattern in [
                r'(\d{5}[-–—\s]+\d{4}[-–—\s]+\d{4})',
                r'(\d{13})',
                r'(\d{12})',
                r'(\d{11})',
            ]:
                matches = re.findall(pattern, text)
                for match in matches:
                    match_clean = re.sub(r'[-–—\s]', '', str(match))
                    if match_clean == search_clean:
                        # tracking_no로 반환
                        return {
                            "pdf_path": str(pdf_path),
                            "tracking_no": match_clean,
                            "original": match,
                            "type": "tracking"
                        }
        
    except Exception as e:
        # 오류는 무시하고 계속 진행
        pass
    
    return None


def search_pdf_files(
    search_value: str,
    search_dirs: List[Path],
    max_workers: Optional[int] = None,
    use_multicore: bool = True,
    cancel_flag: Optional[object] = None,
    progress_callback: Optional[callable] = None
) -> Optional[Dict]:
    """
    여러 폴더에서 PDF 파일 검색 (멀티코어 활용)
    
    Args:
        search_value: 검색할 송장번호 또는 주문번호
        search_dirs: 검색할 폴더 리스트
        max_workers: 최대 워커 수 (None이면 CPU 코어의 70%)
        use_multicore: 멀티코어 사용 여부
        cancel_flag: 취소 플래그 객체 (hasattr(cancel_flag, 'cancelled') and cancel_flag.cancelled이면 중단)
        progress_callback: 진행 상황 콜백 함수 (message: str) -> None
    
    Returns:
        첫 번째 매칭 결과 또는 None
    """
    if not PDF_SUPPORT:
        if progress_callback:
            progress_callback("PDF 라이브러리가 설치되지 않았습니다.")
        return None
    
    # CPU 코어 수 확인 및 워커 수 결정
    cpu_count = multiprocessing.cpu_count()
    if use_multicore:
        if max_workers is None:
            max_workers = max(1, int(cpu_count * 0.7))  # 70% 사용
        if progress_callback:
            progress_callback(f"멀티스레드 검색 시작: {max_workers}개 워커 (CPU 코어 {cpu_count}개 중 {max_workers}개 사용)")
    else:
        max_workers = 1  # 단일 스레드
        if progress_callback:
            progress_callback(f"단일스레드 검색 시작 (CPU 코어 {cpu_count}개 중 1개 사용)")
    
    # 모든 PDF 파일 수집
    pdf_files = []
    for search_dir in search_dirs:
        if search_dir.exists() and search_dir.is_dir():
            files = list(search_dir.glob("*.pdf"))
            pdf_files.extend(files)
            if progress_callback:
                progress_callback(f"폴더 '{search_dir}'에서 {len(files)}개 PDF 파일 발견")
    
    if not pdf_files:
        if progress_callback:
            progress_callback(f"검색할 PDF 파일이 없습니다. (폴더: {[str(d) for d in search_dirs]})")
        return None
    
    total_files = len(pdf_files)
    if progress_callback:
        progress_callback(f"총 {total_files}개 PDF 파일 검색 중... (검색값: {search_value})")
    
    # 취소 확인
    if cancel_flag and hasattr(cancel_flag, 'cancelled') and cancel_flag.cancelled:
        return None
    
    # 멀티코어 검색 (ThreadPoolExecutor 사용 - Windows 호환성)
    # 첫 번째 매칭에서 즉시 종료하기 위해 as_completed 사용
    executor_class = ThreadPoolExecutor  # Windows에서는 항상 ThreadPoolExecutor 사용
    
    checked_count = 0
    with executor_class(max_workers=max_workers) as executor:
        # 모든 작업 제출
        future_to_pdf = {
            executor.submit(_search_pdf_file, (pdf_path, search_value)): pdf_path
            for pdf_path in pdf_files
        }
        
        # 첫 번째 결과를 찾으면 즉시 종료
        for future in as_completed(future_to_pdf):
            checked_count += 1
            
            # 취소 확인
            if cancel_flag and hasattr(cancel_flag, 'cancelled') and cancel_flag.cancelled:
                # 모든 작업 취소
                for f in future_to_pdf:
                    f.cancel()
                if progress_callback:
                    progress_callback(f"검색 취소됨 ({checked_count}/{total_files} 파일 검사 완료)")
                return None
            
            # 진행 상황 업데이트 (10개마다 또는 마지막)
            if progress_callback and (checked_count % 10 == 0 or checked_count == total_files):
                progress_callback(f"검색 중... ({checked_count}/{total_files} 파일 검사 완료)")
            
            try:
                result = future.result(timeout=30)  # 타임아웃 30초
                if result:
                    # 첫 번째 매칭 발견 시 나머지 작업 취소
                    for f in future_to_pdf:
                        f.cancel()
                    if progress_callback:
                        progress_callback(f"✓ 매칭 발견! ({checked_count}/{total_files} 파일 검사 완료)")
                    return result
            except Exception as e:
                if progress_callback and checked_count % 50 == 0:  # 50개마다 오류 로그
                    progress_callback(f"일부 파일 검색 오류 (계속 진행 중...)")
                continue
    
    if progress_callback:
        progress_callback(f"검색 완료: {total_files}개 파일 모두 검사했으나 매칭 없음")
    return None


def find_pdf_by_tracking_or_order(
    search_value: str,
    base_dirs: Optional[List[str]] = None,
    use_multicore: bool = True,
    cancel_flag: Optional[object] = None,
    progress_callback: Optional[callable] = None
) -> Optional[Dict]:
    """
    송장번호 또는 주문번호로 PDF 파일 찾기
    
    Args:
        search_value: 검색할 송장번호 또는 주문번호
        base_dirs: 검색할 기본 폴더 리스트 (None이면 labels, orders 사용)
        use_multicore: 멀티코어 사용 여부
        cancel_flag: 취소 플래그 객체
        progress_callback: 진행 상황 콜백 함수
    
    Returns:
        매칭 정보 또는 None
    """
    if base_dirs is None:
        base_dirs = ["labels", "orders"]
    
    search_paths = [Path(d) for d in base_dirs]
    
    return search_pdf_files(
        search_value,
        search_paths,
        use_multicore=use_multicore,
        cancel_flag=cancel_flag,
        progress_callback=progress_callback
    )

