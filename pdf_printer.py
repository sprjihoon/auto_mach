"""
PDF 자동출력 모듈
Windows os.startfile 방식으로 클릭 없이 기본 프린터로 인쇄
PDF 내용에서 송장번호를 찾아서 해당 페이지만 출력 지원
"""
import os
import tempfile
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from PySide6.QtCore import QObject, Signal

from utils import get_pdf_path, pdf_exists

# PDF 처리 라이브러리
try:
    import pdfplumber
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


class PDFPrinter(QObject):
    """PDF 자동 출력 클래스"""
    
    # 시그널
    print_success = Signal(str)  # 성공 메시지
    print_error = Signal(str)    # 오류 메시지
    index_updated = Signal(int)  # 인덱싱 완료 (페이지 수)
    
    def __init__(self):
        super().__init__()
        self._enabled = True
        self._labels_dir: Optional[Path] = None
        self._pdf_file: Optional[Path] = None  # 단일 PDF 파일
        self._tracking_index: Dict[str, Tuple[Path, int]] = {}  # {tracking_no: (pdf_path, page_num)}
        self._temp_dir = Path(tempfile.gettempdir()) / "auto_mach_labels"
        self._temp_dir.mkdir(exist_ok=True)
    
    @property
    def enabled(self) -> bool:
        return self._enabled
    
    @enabled.setter
    def enabled(self, value: bool):
        self._enabled = value
    
    def set_labels_directory(self, path: str):
        """라벨 PDF 폴더 경로 설정 (하위 호환)"""
        self._labels_dir = Path(path)
    
    def set_pdf_file(self, path: str):
        """단일 PDF 파일 설정"""
        self._pdf_file = Path(path)
        self._labels_dir = self._pdf_file.parent
    
    def build_tracking_index(self) -> int:
        """PDF 파일에서 송장번호 인덱스 생성"""
        if not PDF_SUPPORT:
            self.print_error.emit("PDF 라이브러리가 설치되지 않았습니다 (pdfplumber, PyMuPDF)")
            return 0
        
        self._tracking_index.clear()
        total_pages = 0
        
        # 단일 파일 모드
        if self._pdf_file and self._pdf_file.exists():
            pdf_files = [self._pdf_file]
        elif self._labels_dir and self._labels_dir.exists():
            pdf_files = list(self._labels_dir.glob("*.pdf"))
        else:
            return 0
        
        for pdf_path in pdf_files:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        text = page.extract_text() or ""
                        
                        # 텍스트에서 송장번호 패턴 찾기
                        import re
                        # 다양한 송장번호 패턴 매칭
                        patterns = [
                            r'\b(\d{5}-\d{4}-\d{4})\b',  # 60914-8675-3755 형식
                            r'\b(\d{13})\b',  # 13자리 숫자
                            r'\b(\d{12})\b',  # 12자리 숫자
                            r'\b(\d{10,14})\b',  # 10-14자리 숫자
                        ]
                        
                        for pattern in patterns:
                            matches = re.findall(pattern, text)
                            for match in matches:
                                # 하이픈 제거한 버전도 저장
                                clean_match = match.replace('-', '')
                                if clean_match not in self._tracking_index:
                                    self._tracking_index[clean_match] = (pdf_path, page_num)
                                    total_pages += 1
                                # 원본 형식도 저장
                                if match not in self._tracking_index:
                                    self._tracking_index[match] = (pdf_path, page_num)
                        
            except Exception as e:
                self.print_error.emit(f"PDF 스캔 오류 ({pdf_path.name}): {str(e)}")
                continue
        
        self.index_updated.emit(total_pages)
        return total_pages
    
    def get_indexed_tracking_numbers(self) -> List[str]:
        """인덱싱된 송장번호 목록 반환"""
        return list(self._tracking_index.keys())
    
    def extract_page_to_temp(self, tracking_no: str) -> Optional[Path]:
        """송장번호에 해당하는 페이지를 임시 PDF로 추출"""
        if tracking_no not in self._tracking_index:
            return None
        
        pdf_path, page_num = self._tracking_index[tracking_no]
        
        try:
            # PyMuPDF로 특정 페이지만 추출
            doc = fitz.open(pdf_path)
            
            # 새 PDF 생성
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            
            # 임시 파일로 저장
            temp_path = self._temp_dir / f"{tracking_no}.pdf"
            new_doc.save(str(temp_path))
            new_doc.close()
            doc.close()
            
            return temp_path
            
        except Exception as e:
            self.print_error.emit(f"페이지 추출 오류: {str(e)}")
            return None
    
    def get_pdf_path(self, tracking_no: str) -> Path:
        """tracking_no로 PDF 경로 반환"""
        if self._labels_dir:
            return self._labels_dir / f"{tracking_no}.pdf"
        return get_pdf_path(tracking_no)
    
    def print_pdf(self, tracking_no: str) -> bool:
        """
        PDF 자동 출력
        1. 인덱스에서 송장번호 찾기 → 해당 페이지만 추출하여 출력
        2. 없으면 {tracking_no}.pdf 파일 직접 출력
        """
        if not self._enabled:
            self.print_error.emit("PDF 출력이 비활성화되어 있습니다")
            return False
        
        pdf_path = None
        
        # 1. 인덱스에서 찾기
        if tracking_no in self._tracking_index:
            pdf_path = self.extract_page_to_temp(tracking_no)
            if pdf_path:
                self.print_success.emit(f"인덱스에서 찾음: {tracking_no} (페이지 추출)")
        
        # 2. 직접 파일 찾기
        if pdf_path is None:
            pdf_path = self.get_pdf_path(tracking_no)
            if not pdf_path.exists():
                self.print_error.emit(f"PDF 파일 없음: {pdf_path}")
                return False
        
        try:
            # Windows os.startfile로 기본 프린터에 인쇄
            os.startfile(str(pdf_path), "print")
            
            self.print_success.emit(f"PDF 인쇄 요청: {tracking_no}")
            return True
            
        except FileNotFoundError:
            self.print_error.emit(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
            return False
            
        except OSError as e:
            self.print_error.emit(f"PDF 인쇄 OS 오류: {str(e)}")
            return False
            
        except Exception as e:
            self.print_error.emit(f"PDF 인쇄 오류: {str(e)}")
            return False
    
    def check_pdf_exists(self, tracking_no: str) -> bool:
        """PDF 파일 존재 여부 확인"""
        pdf_path = self.get_pdf_path(tracking_no)
        return pdf_path.exists()


def print_pdf_simple(tracking_no: str, labels_dir: str = "labels") -> bool:
    """
    간단한 PDF 출력 함수 (클래스 없이 사용)
    
    사용예:
        print_pdf_simple("6091486739755")
        print_pdf_simple("6091486739755", "C:/labels")
    """
    pdf_path = Path(labels_dir) / f"{tracking_no}.pdf"
    
    if not pdf_path.exists():
        print(f"[오류] PDF 파일 없음: {pdf_path}")
        return False
    
    try:
        os.startfile(str(pdf_path), "print")
        print(f"[성공] PDF 인쇄 요청: {tracking_no}.pdf")
        return True
    except Exception as e:
        print(f"[오류] PDF 인쇄 실패: {str(e)}")
        return False

