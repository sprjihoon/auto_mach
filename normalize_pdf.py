"""
PDF 정규화 모듈
입력 PDF에서 송장 내용 영역만 크롭하여 정규화
"""
from pathlib import Path
from typing import Optional, Tuple

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


def normalize_pdf(input_path: str, output_path: str) -> bool:
    """
    PDF에서 실제 내용 영역만 감지하여 크롭
    
    처리 과정:
    1. PyMuPDF로 실제 내용이 있는 영역 감지
    2. 내용 영역만 크롭 (공백 제거)
    3. 168mm × 107mm 크기로 조정
    
    Args:
        input_path: 입력 PDF 파일 경로
        output_path: 출력 PDF 파일 경로
        
    Returns:
        성공 여부 (bool)
    """
    if not PYMUPDF_AVAILABLE:
        raise ImportError("PyMuPDF가 필요합니다. pip install PyMuPDF")
    
    input_path = Path(input_path)
    output_path = Path(output_path)
    
    if not input_path.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {input_path}")
    
    try:
        # 원본 PDF 열기
        doc = fitz.open(str(input_path))
        
        # 새 PDF 생성
        new_doc = fitz.open()
        
        # 1단계: 모든 페이지의 내용 영역을 먼저 감지하여 최대 크기 결정
        all_crop_rects = []
        max_content_width = 0
        max_content_height = 0
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # 원본 페이지 크기
            page_rect = page.rect
            original_width = page_rect.width
            original_height = page_rect.height
            
            # 실제 내용 영역 감지 (텍스트 + 이미지 + 드로잉)
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            
            min_x = float('inf')
            min_y = float('inf')
            max_x = float('-inf')
            max_y = float('-inf')
            
            # 텍스트 블록 처리
            for block in blocks:
                bbox = block.get("bbox", None)
                if bbox:
                    x0, y0, x1, y1 = bbox
                    min_x = min(min_x, x0)
                    min_y = min(min_y, y0)
                    max_x = max(max_x, x1)
                    max_y = max(max_y, y1)
            
            # 이미지 영역 처리
            image_list = page.get_images()
            for img in image_list:
                try:
                    xref = img[0]
                    rects = page.get_image_rects(xref)
                    for rect in rects:
                        min_x = min(min_x, rect.x0)
                        min_y = min(min_y, rect.y0)
                        max_x = max(max_x, rect.x1)
                        max_y = max(max_y, rect.y1)
                except:
                    pass
            
            # 드로잉(선, 사각형 등) 영역 처리
            try:
                drawings = page.get_drawings()
                for drawing in drawings:
                    rect = drawing.get("rect", None)
                    if rect:
                        min_x = min(min_x, rect.x0)
                        min_y = min(min_y, rect.y0)
                        max_x = max(max_x, rect.x1)
                        max_y = max(max_y, rect.y1)
            except:
                pass
            
            # 내용이 감지되지 않은 경우 전체 페이지 사용
            if min_x == float('inf'):
                min_x = 0
                min_y = 0
                max_x = original_width
                max_y = original_height
            
            # 여백 추가 (3mm)
            margin = 3.0 * MM_TO_PT
            min_x = max(0, min_x - margin)
            min_y = max(0, min_y - margin)
            max_x = min(original_width, max_x + margin)
            max_y = min(original_height, max_y + margin)
            
            # 크롭 영역 저장
            crop_rect = fitz.Rect(min_x, min_y, max_x, max_y)
            all_crop_rects.append(crop_rect)
            
            # 최대 크기 업데이트
            max_content_width = max(max_content_width, crop_rect.width)
            max_content_height = max(max_content_height, crop_rect.height)
        
        # 2단계: 모든 페이지를 라벨 용지 크기(168mm × 107mm)로 생성
        # 프린터가 "용지에 맞춤"을 해도 실제 크기가 유지되도록 함
        for page_num in range(len(doc)):
            crop_rect = all_crop_rects[page_num]
            
            # 새 페이지 생성 (라벨 용지 크기로 고정)
            new_page = new_doc.new_page(
                width=LABEL_WIDTH_PT,  # 168mm
                height=LABEL_HEIGHT_PT  # 107mm
            )
            
            # 크롭된 영역의 크기
            content_width = crop_rect.width
            content_height = crop_rect.height
            
            # 스케일 계산 (라벨 크기에 맞추되, 비율 유지)
            scale_x = LABEL_WIDTH_PT / max_content_width
            scale_y = LABEL_HEIGHT_PT / max_content_height
            scale = min(scale_x, scale_y)  # 비율 유지를 위해 작은 값 사용
            
            # 스케일링된 내용 크기
            scaled_width = content_width * scale
            scaled_height = content_height * scale
            
            # 대상 영역 (좌측 상단 정렬, 스케일 적용)
            target_rect = fitz.Rect(0, 0, scaled_width, scaled_height)
            
            # 원본 페이지의 크롭 영역을 새 페이지에 복사 (스케일링됨)
            new_page.show_pdf_page(
                target_rect,  # 대상 영역 (스케일 적용)
                doc,  # 원본 문서
                page_num,  # 원본 페이지 번호
                clip=crop_rect  # 크롭할 영역
            )
        
        # 출력 디렉토리 생성
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # PDF 저장
        new_doc.save(str(output_path))
        new_doc.close()
        doc.close()
        
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
