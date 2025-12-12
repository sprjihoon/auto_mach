"""
재출력용 PDF 페이지 추출 모듈
2페이지 감지 및 추출 기능 포함
"""
import re
import tempfile
from pathlib import Path
from typing import Optional, Tuple

try:
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


def extract_reprint_page_to_temp(
    pdf_path: Path,
    tracking_no: str,
    is_order_sheet: bool,
    keep_temp_files: bool
) -> Optional[Path]:
    """
    재출력용 PDF 페이지를 임시 파일로 추출합니다.
    2장 송장/주문서 로직을 포함하여 다음 페이지에 수령자 이름이나 제품 정보가 있으면 함께 추출합니다.
    
    Args:
        pdf_path: 원본 PDF 파일 경로
        tracking_no: 검색된 송장번호
        is_order_sheet: 주문서(True)인지 송장(False)인지 여부
        keep_temp_files: 임시 파일을 보관할지 여부
    
    Returns:
        추출된 임시 PDF 파일 경로 또는 None
    """
    prefix = "[주문서 재출력] " if is_order_sheet else "[송장 재출력] "
    
    try:
        clean_tracking_no = re.sub(r'[-–—\s]', '', tracking_no)
        
        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)
        
        # 송장번호가 포함된 페이지 찾기 (정확도를 위해 전체 페이지 스캔)
        found_page_num = -1
        for p_idx in range(total_pages):
            page_text = doc[p_idx].get_text() or ""
            if clean_tracking_no in re.sub(r'[-–—\s]', '', page_text):
                found_page_num = p_idx
                break
        
        if found_page_num == -1:
            print(f"{prefix}오류: PDF에서 송장번호 '{tracking_no}'를 찾을 수 없습니다. 원본 파일 직접 출력 시도.")
            doc.close()
            return None # 송장번호를 찾지 못하면 추출하지 않음
        
        start_page = found_page_num
        end_page = found_page_num
        
        # 2장 송장/주문서 처리 로직
        if found_page_num + 1 < total_pages:
            next_page = doc[found_page_num + 1]
            next_text = next_page.get_text() or ""
            
            # 다음 페이지에 다른 송장번호가 있는지 확인
            next_tracking_patterns = [
                r'\d{5}[-–—\s]+\d{4}[-–—\s]+\d{4}',  # 5-4-4 형식
                r'\b\d{13}\b',  # 13자리
                r'\b\d{12}\b',  # 12자리
                r'\b\d{11}\b',  # 11자리
            ]
            
            next_has_other_tracking = False
            for pattern in next_tracking_patterns:
                matches = re.findall(pattern, next_text)
                for match in matches:
                    clean_match = re.sub(r'[-–—\s]', '', match)
                    if clean_match.isdigit() and len(clean_match) >= 10:
                        if clean_match != clean_tracking_no: # 현재 송장번호와 다르면 다른 송장으로 간주
                            next_has_other_tracking = True
                            break
                if next_has_other_tracking:
                    break
            
            if not next_has_other_tracking:
                # 다음 페이지에 고객 정보나 제품 정보가 있는지 확인
                has_relevant_info = any(keyword in next_text for keyword in [
                    '수령자', '받는분', '수신인', '고객', '주문자',
                    '상품명', '제품명', '품목', '수량', '개', '옵션'
                ])
                
                if has_relevant_info or len(next_text.strip()) > 50: # 내용이 충분히 많으면 포함
                    end_page = found_page_num + 1
                    print(f"{prefix}✓ 2장 감지: 다음 페이지({end_page + 1})도 함께 추출")
        
        extracted_pages_count = end_page - start_page + 1
        print(f"{prefix}PDF 페이지 추출: {tracking_no} (페이지 {start_page + 1}부터 {end_page + 1}까지, 총 {extracted_pages_count}장)")
        
        optimized_doc = fitz.open()
        
        for page_idx in range(start_page, end_page + 1):
            page = doc[page_idx]
            original_rect = page.rect
            original_rotation = page.rotation
            
            # 송장(라벨)은 크롭, 주문서는 원본 전체 사용
            if is_order_sheet:
                # 주문서: 크롭 없이 전체 페이지 사용
                dpi = 300
                zoom = dpi / 72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)  # clip 파라미터 없음 (전체 페이지)
            else:
                # 송장(라벨): 내용 영역만 크롭
                # 내용 영역 감지 (텍스트 블록 기준)
                try:
                    blocks = page.get_text("blocks") or []
                    xs0, ys0, xs1, ys1 = [], [], [], []
                    for block in blocks:
                        if len(block) >= 5:
                            x0, y0, x1, y1, text = block[:5]
                            if isinstance(text, str) and text.strip():
                                xs0.append(x0)
                                ys0.append(y0)
                                xs1.append(x1)
                                ys1.append(y1)
                    if xs0 and ys0 and xs1 and ys1:
                        margin = 10
                        clip_rect = fitz.Rect(
                            max(original_rect.x0, min(xs0) - margin),
                            max(original_rect.y0, min(ys0) - margin),
                            min(original_rect.x1, max(xs1) + margin),
                            min(original_rect.y1, max(ys1) + margin),
                        )
                    else:
                        clip_rect = original_rect
                except Exception:
                    clip_rect = original_rect
                
                dpi = 300
                zoom = dpi / 72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, clip=clip_rect, alpha=False)  # 크롭 적용
            
            # 새 페이지 생성 (원본 크기 및 방향 유지)
            if original_rotation in [90, 270]:
                new_page = optimized_doc.new_page(width=original_rect.height, height=original_rect.width)
            else:
                new_page = optimized_doc.new_page(width=original_rect.width, height=original_rect.height)
            
            target_rect = fitz.Rect(0, 0, new_page.rect.width, new_page.rect.height)
            new_page.insert_image(target_rect, pixmap=pix, rotate=0, keep_proportion=True, overlay=True)
        
        temp_filename = f"{clean_tracking_no}_reprint_{'order' if is_order_sheet else 'label'}.pdf"
        temp_dir = Path(tempfile.gettempdir()) / "auto_mach_reprint_temp"
        temp_dir.mkdir(exist_ok=True)
        temp_path = temp_dir / temp_filename
        if temp_path.exists():
            temp_path.unlink()
        optimized_doc.save(str(temp_path))
        
        optimized_doc.close()
        doc.close()
        
        print(f"{prefix}✅ PDF 생성 완료: {temp_path.name} ({extracted_pages_count}장)")
        return temp_path
        
    except Exception as e:
        print(f"{prefix}페이지 추출 오류: {str(e)}")
        return None


def extract_pages_from_pdf(
    pdf_path: Path,
    tracking_no: str,
    search_start_page: int = 0
) -> Optional[Tuple[Path, int]]:
    """
    PDF에서 송장번호를 찾아 해당 페이지(들)를 추출 (송장용 - 크롭 적용)
    
    Args:
        pdf_path: PDF 파일 경로
        tracking_no: 송장번호
        search_start_page: 검색 시작 페이지 (기본값: 0)
    
    Returns:
        (임시_PDF_경로, 추출된_페이지_수) 또는 None
    """
    # 송장용이므로 크롭 적용
    return _extract_pages_with_crop(pdf_path, tracking_no, search_start_page)


def _extract_pages_with_crop(
    pdf_path: Path,
    tracking_no: str,
    search_start_page: int = 0
) -> Optional[Tuple[Path, int]]:
    """
    PDF에서 송장번호를 찾아 해당 페이지(들)를 추출 (크롭 적용)
    """
    if not PDF_SUPPORT:
        return None
    
    if not pdf_path.exists():
        return None
    
    try:
        clean_tracking_no = re.sub(r'[-–—\s]', '', tracking_no)
        
        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)
        
        # 송장번호 패턴
        tracking_patterns = [
            r'\d{5}[-–—\s]+\d{4}[-–—\s]+\d{4}',  # 5-4-4 형식
            r'\b\d{13}\b',  # 13자리
            r'\b\d{12}\b',  # 12자리
            r'\b\d{11}\b',  # 11자리
        ]
        
        found_page = None
        
        # 송장번호가 있는 페이지 찾기
        for page_num in range(search_start_page, total_pages):
            page = doc[page_num]
            text = page.get_text() or ""
            
            for pattern in tracking_patterns:
                matches = re.findall(pattern, text)
                for match in matches:
                    clean_match = re.sub(r'[-–—\s]', '', match)
                    if clean_match == clean_tracking_no or clean_match.startswith(clean_tracking_no) or clean_tracking_no.startswith(clean_match):
                        found_page = page_num
                        break
                if found_page is not None:
                    break
            if found_page is not None:
                break
        
        if found_page is None:
            doc.close()
            return None
        
        # 2페이지 감지 로직
        start_page = found_page
        end_page = found_page
        
        # 다음 페이지 확인
        if found_page + 1 < total_pages:
            next_page = doc[found_page + 1]
            next_text = next_page.get_text() or ""
            
            # 다음 페이지에 다른 송장번호가 있는지 확인
            next_has_tracking = False
            for pattern in tracking_patterns:
                matches = re.findall(pattern, next_text)
                for match in matches:
                    clean_match = re.sub(r'[-–—\s]', '', match)
                    if clean_match.isdigit() and len(clean_match) >= 10:
                        if clean_match != clean_tracking_no:
                            next_has_tracking = True
                            break
                if next_has_tracking:
                    break
            
            # 다음 페이지에 송장번호가 없고, 고객 정보나 제품 정보가 있으면 포함
            if not next_has_tracking:
                has_customer_info = any(keyword in next_text for keyword in [
                    '수령자', '받는분', '수신인', '고객', '주문자',
                    '상품명', '제품명', '품목', '수량', '개'
                ])
                
                # 현재 페이지에서 수령자 이름 추출 시도
                current_page = doc[found_page]
                current_text = current_page.get_text() or ""
                recipient_name = None
                
                name_patterns = [
                    r'수령자[:\s]*([가-힣]{2,4})',
                    r'받는분[:\s]*([가-힣]{2,4})',
                    r'수신인[:\s]*([가-힣]{2,4})',
                ]
                
                for pattern in name_patterns:
                    match = re.search(pattern, current_text)
                    if match:
                        recipient_name = match.group(1).strip()
                        break
                
                if has_customer_info or (recipient_name and len(next_text.strip()) > 20):
                    end_page = found_page + 1
        
        # 페이지 추출 (송장용 - 크롭 적용)
        temp_dir = Path(tempfile.gettempdir()) / "auto_mach_reprint"
        temp_dir.mkdir(exist_ok=True)
        
        optimized_doc = fitz.open()
        
        for page_idx in range(start_page, end_page + 1):
            page = doc[page_idx]
            original_rect = page.rect
            original_rotation = page.rotation
            
            # 내용 영역 감지 (텍스트 블록 기준)
            try:
                blocks = page.get_text("blocks") or []
                xs0, ys0, xs1, ys1 = [], [], [], []
                for block in blocks:
                    if len(block) >= 5:
                        x0, y0, x1, y1, text = block[:5]
                        if isinstance(text, str) and text.strip():
                            xs0.append(x0)
                            ys0.append(y0)
                            xs1.append(x1)
                            ys1.append(y1)
                if xs0 and ys0 and xs1 and ys1:
                    margin = 10
                    clip_rect = fitz.Rect(
                        max(original_rect.x0, min(xs0) - margin),
                        max(original_rect.y0, min(ys0) - margin),
                        min(original_rect.x1, max(xs1) + margin),
                        min(original_rect.y1, max(ys1) + margin),
                    )
                else:
                    clip_rect = original_rect
            except Exception:
                clip_rect = original_rect
            
            # 고해상도 렌더링 (크롭 적용)
            dpi = 300
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, clip=clip_rect, alpha=False)
            
            # 새 페이지 생성 (원본 크기 및 방향 유지)
            if original_rotation in [90, 270]:
                new_page = optimized_doc.new_page(width=original_rect.height, height=original_rect.width)
            else:
                new_page = optimized_doc.new_page(width=original_rect.width, height=original_rect.height)
            
            # 이미지를 삽입
            target_rect = fitz.Rect(0, 0, new_page.rect.width, new_page.rect.height)
            new_page.insert_image(target_rect, pixmap=pix, rotate=0, keep_proportion=True, overlay=True)
        
        temp_path = temp_dir / f"{clean_tracking_no}.pdf"
        if temp_path.exists():
            temp_path.unlink()
        
        optimized_doc.save(str(temp_path))
        optimized_doc.close()
        doc.close()
        
        extracted_pages = end_page - start_page + 1
        return (temp_path, extracted_pages)
        
    except Exception as e:
        return None
