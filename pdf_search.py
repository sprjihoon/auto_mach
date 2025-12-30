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

# PDF 처리 라이브러리
try:
    import pdfplumber
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


def _extract_text_from_page(page, use_pdfplumber: bool = True) -> List[str]:
    """
    단일 페이지에서 텍스트 추출 (다양한 방법 시도)
    """
    texts = []
    
    if use_pdfplumber:
        # pdfplumber page
        try:
            text = page.extract_text() or ""
            if text.strip():
                texts.append(text)
            
            # 고정밀 옵션
            for method in [
                {"x_tolerance": 1, "y_tolerance": 1, "layout": True},
                {"x_tolerance": 3, "y_tolerance": 3, "layout": True},
            ]:
                try:
                    alt = page.extract_text(**method) or ""
                    if alt.strip() and alt not in texts:
                        texts.append(alt)
                except:
                    pass
        except:
            pass
    else:
        # PyMuPDF page
        try:
            text1 = page.get_text() or ""
            if text1.strip():
                texts.append(text1)
            
            # 블록
            try:
                blocks = page.get_text("blocks") or []
                block_text = " ".join(str(b[4]) for b in blocks if len(b) >= 5 and isinstance(b[4], str))
                if block_text.strip() and block_text not in texts:
                    texts.append(block_text)
            except:
                pass
            
            # 단어
            try:
                words = page.get_text("words") or []
                word_text = " ".join(str(w[4]) for w in words if len(w) >= 5)
                if word_text.strip() and word_text not in texts:
                    texts.append(word_text)
            except:
                pass
        except:
            pass
    
    return texts


def _find_number_in_text(text: str, search_clean: str) -> bool:
    """
    텍스트에서 검색값(정규화된) 찾기
    줄바꿈, 하이픈, 공백 모두 제거하고 비교
    """
    if not text:
        return False
    
    # 텍스트 정규화 (줄바꿈, 하이픈, 공백 모두 제거)
    text_clean = re.sub(r'[-–—\s\n\r\t]', '', text)
    
    # 직접 포함 여부
    return search_clean in text_clean


def _normalize_text_for_pattern(text: str) -> str:
    """
    패턴 매칭을 위한 텍스트 정규화
    줄바꿈을 공백으로 치환하여 연속된 숫자 패턴을 찾을 수 있게 함
    """
    if not text:
        return ""
    # 줄바꿈을 빈 문자열로 치환 (숫자가 분리되어 있을 때 합쳐짐)
    return re.sub(r'[\n\r]+', '', text)


def _search_single_pdf(
    pdf_path: Path,
    search_clean: str,
    is_order_search: bool = False
) -> Optional[Dict]:
    """
    단일 PDF에서 검색 (찾으면 즉시 반환)
    
    Args:
        pdf_path: PDF 파일 경로
        search_clean: 정규화된 검색값 (하이픈 제거됨)
        is_order_search: 주문번호 검색 여부 (15자리면 True)
    
    Returns:
        매칭 정보 또는 None
    """
    if not PDF_SUPPORT:
        return None
    
    # 송장번호 패턴
    tracking_patterns = [
        r'\d{5}[-–—\s]+\d{4}[-–—\s]+\d{4}',
        r'\d{5}[-–—\s]+\d{4}[-–—\s]+\d{5}',
        r'\d{4}[-–—\s]+\d{4}[-–—\s]+\d{5}',
        r'\d{4}[-–—\s]+\d{4}[-–—\s]+\d{4}',
        r'\d{13}',
        r'\d{12}',
        r'\d{11}',
    ]
    
    # 주문번호 패턴
    order_patterns = [
        r'\d{8}[-–—]\d{7}',     # 20251212-0000051
        r'\d{8}[-–—]\d{6}',     # 20251212-000005
        r'\d{15}',              # 202512120000051
        r'\d{14}',              # 20251212000005
    ]
    
    # 검색 대상 패턴 결정
    if is_order_search:
        patterns = order_patterns + tracking_patterns
    else:
        patterns = tracking_patterns + order_patterns
    
    try:
        # 방법 1: pdfplumber
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    texts = _extract_text_from_page(page, use_pdfplumber=True)
                    
                    for text in texts:
                        # 빠른 체크: 검색값이 텍스트에 있는지
                        if _find_number_in_text(text, search_clean):
                            # 줄바꿈 제거하여 패턴 매칭
                            normalized_text = _normalize_text_for_pattern(text)
                            
                            # 정확한 패턴 매칭 시도
                            for pattern in patterns:
                                matches = re.findall(pattern, normalized_text)
                                for match in matches:
                                    clean = re.sub(r'[-–—\s]', '', str(match))
                                    if clean == search_clean:
                                        result_type = "order" if len(clean) >= 14 else "tracking"
                                        return {
                                            "pdf_path": str(pdf_path),
                                            "tracking_no": clean if result_type == "tracking" else None,
                                            "order_no": clean if result_type == "order" else None,
                                            "original": match,
                                            "type": result_type,
                                            "page": page_num
                                        }
                            
                            # 패턴 매칭 실패해도 텍스트에서 직접 찾았으면 반환 (fallback)
                            result_type = "order" if len(search_clean) >= 14 else "tracking"
                            return {
                                "pdf_path": str(pdf_path),
                                "tracking_no": search_clean if result_type == "tracking" else None,
                                "order_no": search_clean if result_type == "order" else None,
                                "original": search_clean,
                                "type": result_type,
                                "page": page_num
                            }
        except Exception:
            pass
        
        # 방법 2: PyMuPDF
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                texts = _extract_text_from_page(page, use_pdfplumber=False)
                
                for text in texts:
                    # 빠른 체크
                    if _find_number_in_text(text, search_clean):
                        # 줄바꿈 제거하여 패턴 매칭
                        normalized_text = _normalize_text_for_pattern(text)
                        
                        # 정확한 패턴 매칭 시도
                        for pattern in patterns:
                            matches = re.findall(pattern, normalized_text)
                            for match in matches:
                                clean = re.sub(r'[-–—\s]', '', str(match))
                                if clean == search_clean:
                                    doc.close()
                                    result_type = "order" if len(clean) >= 14 else "tracking"
                                    return {
                                        "pdf_path": str(pdf_path),
                                        "tracking_no": clean if result_type == "tracking" else None,
                                        "order_no": clean if result_type == "order" else None,
                                        "original": match,
                                        "type": result_type,
                                        "page": page_num
                                    }
                        
                        # 패턴 매칭 실패해도 텍스트에서 직접 찾았으면 반환 (fallback)
                        doc.close()
                        result_type = "order" if len(search_clean) >= 14 else "tracking"
                        return {
                            "pdf_path": str(pdf_path),
                            "tracking_no": search_clean if result_type == "tracking" else None,
                            "order_no": search_clean if result_type == "order" else None,
                            "original": search_clean,
                            "type": result_type,
                            "page": page_num
                        }
                
                # 특수 패턴 (등기번호, 송장번호, 주문번호 라벨)
                for text in texts:
                    special_patterns = [
                        r'등기번호[:\s\-]*(\d{5}[-–—\s]?\d{4}[-–—\s]?\d{4,5})',
                        r'송장번호[:\s\-]*(\d{5}[-–—\s]?\d{4}[-–—\s]?\d{4,5})',
                        r'주문번호[:\s\-]*(\d{8}[-–—]?\d{6,7})',
                    ]
                    
                    for sp in special_patterns:
                        matches = re.findall(sp, text)
                        for match in matches:
                            clean = re.sub(r'[-–—\s]', '', str(match))
                            if clean == search_clean:
                                doc.close()
                                result_type = "order" if len(clean) >= 14 else "tracking"
                                return {
                                    "pdf_path": str(pdf_path),
                                    "tracking_no": clean if result_type == "tracking" else None,
                                    "order_no": clean if result_type == "order" else None,
                                    "original": match,
                                    "type": result_type,
                                    "page": page_num
                                }
            
            doc.close()
        except Exception:
            pass
        
    except Exception:
        pass
    
    return None


def _search_pdf_file(args) -> Optional[Dict]:
    """
    멀티스레드용 래퍼 함수
    """
    if len(args) >= 3:
        pdf_path, search_value, debug_callback = args[:3]
    else:
        pdf_path, search_value = args[:2]
        debug_callback = None
    
    try:
        # 검색값 정규화
        search_str = str(search_value).strip()
        search_clean = re.sub(r'[-–—\s]', '', search_str)
        
        # 숫자가 아니거나 너무 짧으면 스킵
        if not search_clean.isdigit() or len(search_clean) < 8:
            return None
        
        # 주문번호인지 송장번호인지 판단 (14자리 이상이면 주문번호)
        is_order_search = len(search_clean) >= 14
        
        result = _search_single_pdf(pdf_path, search_clean, is_order_search)
        
        if result and debug_callback:
            debug_callback(f"[FOUND] {pdf_path.name}: {result.get('original')}")
        
        return result
        
    except Exception as e:
        if debug_callback:
            debug_callback(f"[ERROR] {pdf_path.name}: {str(e)}")
        return None


def search_pdf_files(
    search_value: str,
    search_dirs: List[Path],
    max_workers: Optional[int] = None,
    use_multicore: bool = True,
    cancel_flag: Optional[object] = None,
    progress_callback: Optional[callable] = None,
    debug_callback: Optional[callable] = None
) -> Optional[Dict]:
    """
    여러 폴더에서 PDF 파일 검색 (멀티코어 활용)
    """
    if not PDF_SUPPORT:
        if progress_callback:
            progress_callback("PDF 라이브러리가 설치되지 않았습니다.")
        return None
    
    # CPU 코어 수 확인 및 워커 수 결정
    cpu_count = multiprocessing.cpu_count()
    if use_multicore:
        if max_workers is None:
            max_workers = max(1, int(cpu_count * 0.7))
        if progress_callback:
            progress_callback(f"멀티스레드 검색: {max_workers}개 워커")
    else:
        max_workers = 1
        if progress_callback:
            progress_callback("단일스레드 검색")
    
    # 모든 PDF 파일 수집
    pdf_files = []
    for search_dir in search_dirs:
        dir_path = Path(search_dir)
        if dir_path.exists() and dir_path.is_dir():
            files = list(dir_path.glob("*.pdf"))
            pdf_files.extend(files)
            if progress_callback:
                progress_callback(f"폴더 '{dir_path.name}'에서 {len(files)}개 PDF 발견")
    
    if not pdf_files:
        if progress_callback:
            progress_callback("검색할 PDF 파일이 없습니다.")
        return None
    
    total_files = len(pdf_files)
    if progress_callback:
        progress_callback(f"총 {total_files}개 파일 검색 시작")
    
    # 취소 확인
    if cancel_flag and hasattr(cancel_flag, 'cancelled') and cancel_flag.cancelled:
        return None
    
    checked_count = 0
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_pdf = {
            executor.submit(_search_pdf_file, (pdf_path, search_value, debug_callback)): pdf_path
            for pdf_path in pdf_files
        }
        
        for future in as_completed(future_to_pdf):
            checked_count += 1
            
            # 취소 확인
            if cancel_flag and hasattr(cancel_flag, 'cancelled') and cancel_flag.cancelled:
                for f in future_to_pdf:
                    f.cancel()
                if progress_callback:
                    progress_callback(f"검색 취소됨 ({checked_count}/{total_files})")
                return None
            
            # 진행 상황 업데이트
            if progress_callback and (checked_count % 5 == 0 or checked_count == total_files):
                progress_callback(f"검색 중... ({checked_count}/{total_files})")
            
            try:
                result = future.result(timeout=60)
                if result:
                    # 첫 번째 매칭 발견 시 즉시 종료 (break)
                    for f in future_to_pdf:
                        f.cancel()
                    if progress_callback:
                        progress_callback(f"✓ 매칭 발견! ({checked_count}/{total_files})")
                    return result
            except Exception:
                continue
    
    if progress_callback:
        progress_callback(f"검색 완료: 매칭 없음 ({total_files}개)")
    return None


def find_pdf_by_tracking_or_order(
    search_value: str,
    base_dirs: Optional[List[str]] = None,
    use_multicore: bool = True,
    cancel_flag: Optional[object] = None,
    progress_callback: Optional[callable] = None,
    debug_callback: Optional[callable] = None
) -> Optional[Dict]:
    """
    송장번호 또는 주문번호로 PDF 파일 찾기 (메인 엔트리 포인트)
    """
    if base_dirs is None:
        base_dirs = ["labels", "orders"]
    
    search_paths = [Path(d) for d in base_dirs]
    
    return search_pdf_files(
        search_value,
        search_paths,
        use_multicore=use_multicore,
        cancel_flag=cancel_flag,
        progress_callback=progress_callback,
        debug_callback=debug_callback
    )
