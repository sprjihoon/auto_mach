"""
PDF 정규화 모듈
입력 PDF를 송장 내용 영역만 감지하여 크롭하고 90도 회전하여 정규화
"""
from pathlib import Path
from typing import Optional, Tuple

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter
        PYPDF_AVAILABLE = True
    except ImportError:
        PYPDF_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False


# mm → pt 변환 상수 (1 inch = 25.4 mm, 1 inch = 72 pt)
MM_TO_PT = 72.0 / 25.4

# 우체국 송장 규격 (mm) - 가로 × 세로
LABEL_WIDTH_MM = 168.0  # 가로
LABEL_HEIGHT_MM = 107.0  # 세로

# 포인트 단위로 변환
LABEL_WIDTH_PT = LABEL_WIDTH_MM * MM_TO_PT  # 약 476.221 pt (가로)
LABEL_HEIGHT_PT = LABEL_HEIGHT_MM * MM_TO_PT  # 약 303.307 pt (세로)


def detect_content_bbox(pdf_path: str, page_num: int = 0) -> Tuple[float, float, float, float]:
    """
    PDF 페이지에서 실제 내용이 있는 영역의 bounding box 감지
    
    Args:
        pdf_path: PDF 파일 경로
        page_num: 페이지 번호 (0부터 시작)
        
    Returns:
        (llx, lly, urx, ury) - 좌측 하단 원점 기준 bounding box
    """
    if not PYMUPDF_AVAILABLE:
        raise ImportError("PyMuPDF가 필요합니다. pip install PyMuPDF")
    
    doc = fitz.open(pdf_path)
    try:
        if page_num >= len(doc):
            raise ValueError(f"페이지 번호 {page_num}가 범위를 벗어났습니다 (총 {len(doc)}페이지)")
        
        page = doc[page_num]
        
        # 텍스트 블록 감지
        text_blocks = page.get_text("blocks")
        
        # 이미지 영역 감지
        image_list = page.get_images()
        image_rects = []
        for img_index, img in enumerate(image_list):
            try:
                xref = img[0]
                base_image = doc.extract_image(xref)
                # 이미지가 페이지에 표시되는 위치 찾기
                image_rects.extend(page.get_image_rects(xref))
            except:
                pass
        
        # 모든 내용 영역의 좌표 수집
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')
        
        # 텍스트 블록 처리
        for block in text_blocks:
            if len(block) >= 5:  # (x0, y0, x1, y1, text, ...)
                x0, y0, x1, y1 = block[0], block[1], block[2], block[3]
                min_x = min(min_x, x0)
                min_y = min(min_y, y0)
                max_x = max(max_x, x1)
                max_y = max(max_y, y1)
        
        # 이미지 영역 처리
        for rect in image_rects:
            min_x = min(min_x, rect.x0)
            min_y = min(min_y, rect.y0)
            max_x = max(max_x, rect.x1)
            max_y = max(max_y, rect.y1)
        
        # 내용이 감지되지 않은 경우 전체 페이지 반환
        if min_x == float('inf'):
            media_box = page.mediabox
            return (0.0, 0.0, float(media_box.width), float(media_box.height))
        
        # 여백 추가 (5mm)
        margin_pt = 5.0 * MM_TO_PT
        min_x = max(0.0, min_x - margin_pt)
        min_y = max(0.0, min_y - margin_pt)
        max_x = min(float(page.mediabox.width), max_x + margin_pt)
        max_y = min(float(page.mediabox.height), max_y + margin_pt)
        
        return (min_x, min_y, max_x, max_y)
        
    finally:
        doc.close()


def normalize_pdf(input_path: str, output_path: str) -> bool:
    """
    PDF를 좌측 상단(0,0) 기준으로 168mm × 107mm 영역만 크롭하여 정규화
    
    처리 과정:
    1. 좌측 상단(0,0) 기준으로 168mm × 107mm 크롭 (가로 × 세로)
    2. 회전 없이 원본 방향 유지
    3. 최종 크기: 168mm × 107mm
    
    Args:
        input_path: 입력 PDF 파일 경로
        output_path: 출력 PDF 파일 경로
        
    Returns:
        성공 여부 (bool)
    """
    if not PYPDF_AVAILABLE:
        raise ImportError("pypdf 또는 PyPDF2 라이브러리가 필요합니다. pip install pypdf")
    
    input_path = Path(input_path)
    output_path = Path(output_path)
    
    if not input_path.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {input_path}")
    
    try:
        reader = PdfReader(str(input_path))
        writer = PdfWriter()
        
        for page_num, page in enumerate(reader.pages):
            # 원본 페이지의 MediaBox 가져오기
            # MediaBox는 [llx, lly, urx, ury] 형식 (좌측 하단 원점 기준)
            media_box = page.mediabox
            
            # 원본 페이지 크기 (포인트)
            original_width = float(media_box.width)
            original_height = float(media_box.height)
            
            # 좌측 상단 기준으로 크롭 영역 계산
            # PDF 좌표계는 좌측 하단이 원점(0,0)이므로:
            # - 좌측 하단 좌표: (0, 0)
            # - 우측 상단 좌표: (original_width, original_height)
            # 
            # 좌측 상단을 기준으로 168mm × 107mm 크롭하려면:
            # - 크롭 영역의 좌측 하단: (0, original_height - LABEL_HEIGHT_PT)
            # - 크롭 영역의 우측 상단: (LABEL_WIDTH_PT, original_height)
            #
            # 단, 원본이 라벨보다 작은 경우를 대비해 최소값 사용
            crop_width = min(LABEL_WIDTH_PT, original_width)
            crop_height = min(LABEL_HEIGHT_PT, original_height)
            
            # 좌측 상단 기준 크롭 박스 계산
            # 좌측 하단 원점 기준이므로:
            # - llx (좌측 하단 x): 0
            # - lly (좌측 하단 y): original_height - crop_height
            # - urx (우측 상단 x): crop_width
            # - ury (우측 상단 y): original_height
            crop_box = [
                0.0,  # llx: 좌측 하단 x (좌측 상단 기준이므로 0)
                float(original_height - crop_height),  # lly: 좌측 하단 y
                float(crop_width),  # urx: 우측 상단 x
                float(original_height)  # ury: 우측 상단 y
            ]
            
            # 페이지 복사
            new_page = page
            
            # CropBox 설정 (실제 보이는 영역)
            new_page.cropbox.lower_left = (crop_box[0], crop_box[1])
            new_page.cropbox.upper_right = (crop_box[2], crop_box[3])
            
            # MediaBox도 동일하게 설정 (페이지 물리적 크기)
            # 크롭 후 크기: crop_width × crop_height (168mm × 107mm)
            new_page.mediabox.lower_left = (0.0, 0.0)
            new_page.mediabox.upper_right = (crop_width, crop_height)
            
            # ArtBox, TrimBox, BleedBox도 동일하게 설정 (일관성 유지)
            if hasattr(new_page, 'artbox'):
                new_page.artbox.lower_left = (0.0, 0.0)
                new_page.artbox.upper_right = (crop_width, crop_height)
            
            if hasattr(new_page, 'trimbox'):
                new_page.trimbox.lower_left = (0.0, 0.0)
                new_page.trimbox.upper_right = (crop_width, crop_height)
            
            if hasattr(new_page, 'bleedbox'):
                new_page.bleedbox.lower_left = (0.0, 0.0)
                new_page.bleedbox.upper_right = (crop_width, crop_height)
            
            # 회전 없이 원본 방향 유지
            writer.add_page(new_page)
        
        # 출력 디렉토리 생성
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # PDF 저장
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return True
        
    except Exception as e:
        raise RuntimeError(f"PDF 정규화 실패: {str(e)}")


def normalize_pdf_batch(input_dir: str, output_dir: str, pattern: str = "*.pdf") -> int:
    """
    디렉토리 내 모든 PDF 파일을 일괄 정규화
    
    Args:
        input_dir: 입력 디렉토리 경로
        output_dir: 출력 디렉토리 경로
        pattern: 파일 패턴 (기본값: "*.pdf")
        
    Returns:
        처리된 파일 수
    """
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    pdf_files = list(input_dir.glob(pattern))
    success_count = 0
    
    for pdf_file in pdf_files:
        try:
            output_file = output_dir / pdf_file.name
            normalize_pdf(str(pdf_file), str(output_file))
            success_count += 1
        except Exception as e:
            print(f"오류 [{pdf_file.name}]: {str(e)}")
            continue
    
    return success_count


if __name__ == "__main__":
    # 테스트 코드
    import sys
    
    if len(sys.argv) < 3:
        print("사용법: python normalize_pdf.py <입력PDF> <출력PDF>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    try:
        normalize_pdf(input_file, output_file)
        print(f"✓ 정규화 완료: {output_file}")
    except Exception as e:
        print(f"✗ 오류: {str(e)}")
        sys.exit(1)

