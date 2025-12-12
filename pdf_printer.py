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
    
    def build_tracking_index(self, excel_tracking_numbers: List[str] = None) -> int:
        """
        PDF 파일에서 송장번호 인덱스 생성
        
        Args:
            excel_tracking_numbers: 엑셀에서 가져온 송장번호 목록 (이미지 PDF의 경우 순서대로 매핑)
        """
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
                import re
                # 다양한 송장번호 패턴 매칭
                # 하이픈 변형: 일반 하이픈(-), en-dash(–), em-dash(—), 공백 등
                patterns = [
                    # 5-4-4 형식 (하이픈 변형 포함)
                    r'\b(\d{5}[-–—\s]\d{4}[-–—\s]\d{4})\b',  # 60914-8682-2638 형식 (하이픈 변형 지원)
                    # 5-4-4 형식 (일반 하이픈만)
                    r'\b(\d{5}-\d{4}-\d{4})\b',  # 60914-8675-3755 형식
                    # 연속 숫자 형식
                    r'\b(\d{13})\b',  # 13자리 숫자
                    r'\b(\d{12})\b',  # 12자리 숫자
                    r'\b(\d{10,14})\b',  # 10-14자리 숫자
                ]
                
                # 방법 1: pdfplumber로 텍스트 추출 시도
                text_extracted = False
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        for page_num, page in enumerate(pdf.pages):
                            text = page.extract_text() or ""
                            if text and len(text.strip()) > 0:
                                text_extracted = True
                                found_matches = set()
                                
                                for pattern in patterns:
                                    matches = re.findall(pattern, text)
                                    for match in matches:
                                        # 모든 하이픈 변형과 공백 제거
                                        clean_match = re.sub(r'[-–—\s]', '', match)
                                        
                                        # 숫자만 남았는지 확인 (최소 10자리)
                                        if clean_match.isdigit() and len(clean_match) >= 10:
                                            # 이미 처리한 매치는 건너뛰기
                                            if clean_match in found_matches:
                                                continue
                                            found_matches.add(clean_match)
                                            
                                            # 하이픈 제거한 버전 저장 (주요 인덱스)
                                            if clean_match not in self._tracking_index:
                                                self._tracking_index[clean_match] = (pdf_path, page_num)
                                                total_pages += 1
                                            
                                            # 원본 형식도 저장 (하이픈 포함)
                                            if match != clean_match and match not in self._tracking_index:
                                                self._tracking_index[match] = (pdf_path, page_num)
                except Exception as e:
                    # pdfplumber 실패 시 PyMuPDF로 시도
                    pass
                
                # 방법 2: pdfplumber로 텍스트 추출 실패 시 PyMuPDF로 시도
                if not text_extracted:
                    try:
                        doc = fitz.open(pdf_path)
                        pymupdf_extracted = False
                        for page_num in range(len(doc)):
                            page = doc[page_num]
                            # PyMuPDF로 텍스트 추출 시도
                            text = page.get_text() or ""
                            
                            if text and len(text.strip()) > 0:
                                pymupdf_extracted = True
                                found_matches = set()
                                
                                for pattern in patterns:
                                    matches = re.findall(pattern, text)
                                    for match in matches:
                                        # 모든 하이픈 변형과 공백 제거
                                        clean_match = re.sub(r'[-–—\s]', '', match)
                                        
                                        # 숫자만 남았는지 확인 (최소 10자리)
                                        if clean_match.isdigit() and len(clean_match) >= 10:
                                            # 이미 처리한 매치는 건너뛰기
                                            if clean_match in found_matches:
                                                continue
                                            found_matches.add(clean_match)
                                            
                                            # 하이픈 제거한 버전 저장 (주요 인덱스)
                                            if clean_match not in self._tracking_index:
                                                self._tracking_index[clean_match] = (pdf_path, page_num)
                                                total_pages += 1
                                            
                                            # 원본 형식도 저장 (하이픈 포함)
                                            if match != clean_match and match not in self._tracking_index:
                                                self._tracking_index[match] = (pdf_path, page_num)
                        
                        # PyMuPDF로도 텍스트 추출 실패 시 엑셀 기반 매핑 시도
                        if not pymupdf_extracted:
                            if excel_tracking_numbers and len(excel_tracking_numbers) > 0:
                                page_count = len(doc)
                                
                                # 엑셀의 송장번호를 PDF 페이지 순서대로 매핑
                                for idx, tracking_no in enumerate(excel_tracking_numbers):
                                    if idx < page_count:
                                        # 하이픈 제거한 버전
                                        clean_tracking_no = re.sub(r'[-–—\s]', '', str(tracking_no))
                                        if clean_tracking_no not in self._tracking_index:
                                            self._tracking_index[clean_tracking_no] = (pdf_path, idx)
                                            total_pages += 1
                                        # 원본 형식도 저장
                                        if str(tracking_no) not in self._tracking_index:
                                            self._tracking_index[str(tracking_no)] = (pdf_path, idx)
                                
                                if total_pages > 0:
                                    self.print_success.emit(f"엑셀 송장번호로 PDF 매핑 완료: {total_pages}개 (알PDF 이미지 기반)")
                            else:
                                self.print_error.emit(f"PDF 텍스트 추출 실패 ({pdf_path.name}): 알PDF로 저장된 이미지 PDF입니다. 엑셀 파일을 먼저 로드하면 자동 매핑됩니다.")
                        
                        doc.close()
                    except Exception as e:
                        # 예외 발생 시에도 엑셀 기반 매핑 시도
                        if excel_tracking_numbers and len(excel_tracking_numbers) > 0:
                            try:
                                doc = fitz.open(pdf_path)
                                page_count = len(doc)
                                doc.close()
                                
                                # 엑셀의 송장번호를 PDF 페이지 순서대로 매핑
                                for idx, tracking_no in enumerate(excel_tracking_numbers):
                                    if idx < page_count:
                                        # 하이픈 제거한 버전
                                        clean_tracking_no = re.sub(r'[-–—\s]', '', str(tracking_no))
                                        if clean_tracking_no not in self._tracking_index:
                                            self._tracking_index[clean_tracking_no] = (pdf_path, idx)
                                            total_pages += 1
                                        # 원본 형식도 저장
                                        if str(tracking_no) not in self._tracking_index:
                                            self._tracking_index[str(tracking_no)] = (pdf_path, idx)
                                
                                if total_pages > 0:
                                    self.print_success.emit(f"엑셀 송장번호로 PDF 매핑 완료: {total_pages}개 (알PDF 이미지 기반)")
                            except Exception as e2:
                                self.print_error.emit(f"PDF 매핑 실패 ({pdf_path.name}): {str(e2)}")
                        else:
                            self.print_error.emit(f"PDF 텍스트 추출 실패 ({pdf_path.name}): 알PDF로 저장된 이미지 PDF입니다. 엑셀 파일을 먼저 로드하면 자동 매핑됩니다.")
                        
            except Exception as e:
                self.print_error.emit(f"PDF 스캔 오류 ({pdf_path.name}): {str(e)}")
                continue
        
        self.index_updated.emit(total_pages)
        return total_pages
    
    def get_indexed_tracking_numbers(self) -> List[str]:
        """인덱싱된 송장번호 목록 반환"""
        return list(self._tracking_index.keys())
    
    def extract_page_to_temp(self, tracking_no: str) -> Optional[Path]:
        """
        송장번호에 해당하는 페이지를 임시 PDF로 추출
        다음 페이지에 수령자 이름만 있고 송장번호가 없으면 함께 추출 (2장 송장 처리)
        """
        if tracking_no not in self._tracking_index:
            return None
        
        pdf_path, page_num = self._tracking_index[tracking_no]
        
        try:
            import re
            # 파일명에 사용할 수 있도록 하이픈 제거
            clean_tracking_no = re.sub(r'[-–—\s]', '', tracking_no)
            
            # PyMuPDF로 PDF 열기
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            
            # 페이지 번호 검증 (0-based)
            if page_num < 0 or page_num >= total_pages:
                self.print_error.emit(f"페이지 번호 오류: {page_num} (총 {total_pages}페이지)")
                doc.close()
                return None
            
            # 현재 페이지에서 수령자 이름 추출 시도
            recipient_name = None
            try:
                current_page = doc[page_num]
                current_text = current_page.get_text() or ""
                
                # 수령자 이름 패턴 찾기 (한글 이름, 영문 이름 등)
                # 일반적인 패턴: "수령자", "받는분", "수신인" 등의 키워드 다음에 이름
                name_patterns = [
                    r'수령자[:\s]*([가-힣]{2,4})',
                    r'받는분[:\s]*([가-힣]{2,4})',
                    r'수신인[:\s]*([가-힣]{2,4})',
                    r'받는\s*사람[:\s]*([가-힣]{2,4})',
                    r'수령인[:\s]*([가-힣]{2,4})',
                ]
                
                for pattern in name_patterns:
                    match = re.search(pattern, current_text)
                    if match:
                        recipient_name = match.group(1).strip()
                        break
            except Exception:
                pass
            
            # 인쇄할 페이지 범위 결정
            end_page = page_num
            
            # 다음 페이지 확인 (2장 송장 처리)
            if page_num + 1 < total_pages and recipient_name:
                try:
                    next_page = doc[page_num + 1]
                    next_text = next_page.get_text() or ""
                    
                    # 다음 페이지에 송장번호가 있는지 확인
                    tracking_patterns = [
                        r'\b\d{5}[-–—\s]\d{4}[-–—\s]\d{4}\b',
                        r'\b\d{13}\b',
                        r'\b\d{12}\b',
                    ]
                    
                    has_tracking_no = False
                    for pattern in tracking_patterns:
                        if re.search(pattern, next_text):
                            has_tracking_no = True
                            break
                    
                    # 다음 페이지에 송장번호가 없고, 수령자 이름이 있으면 함께 인쇄
                    if not has_tracking_no and recipient_name in next_text:
                        end_page = page_num + 1
                        self.print_success.emit(f"2장 송장 감지: {tracking_no} (수령자: {recipient_name}, 페이지 {page_num + 1}-{end_page + 1})")
                except Exception:
                    pass
            
            # 새 PDF 생성 (1장 또는 2장)
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_num, to_page=end_page)
            
            # 임시 파일로 저장 (하이픈 제거 버전 사용)
            temp_path = self._temp_dir / f"{clean_tracking_no}.pdf"
            new_doc.save(str(temp_path))
            new_doc.close()
            doc.close()
            
            if end_page > page_num:
                self.print_success.emit(f"페이지 추출 완료: {tracking_no} (페이지 {page_num + 1}-{end_page + 1}/{total_pages}, 2장 송장)")
            else:
                self.print_success.emit(f"페이지 추출 완료: {tracking_no} (페이지 {page_num + 1}/{total_pages})")
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
        
        import re
        
        # 하이픈 제거한 버전으로 정규화
        clean_tracking_no = re.sub(r'[-–—\s]', '', tracking_no)
        
        pdf_path = None
        
        # 1. 인덱스에서 송장번호 찾기 (원본 PDF 파일과 페이지 번호 확인)
        original_pdf_path = None
        page_num = None
        search_keys = [clean_tracking_no, tracking_no]
        for key in search_keys:
            if key in self._tracking_index:
                original_pdf_path, page_num = self._tracking_index[key]
                self.print_success.emit(f"PDF 파일에서 송장번호 찾음: {key} (원본: {original_pdf_path.name}, 페이지: {page_num + 1})")
                break
        
        # 2. 해당 페이지를 임시 파일로 추출하여 실물 프린터로 인쇄
        if original_pdf_path and page_num is not None:
            pdf_path = self.extract_page_to_temp(key if key in self._tracking_index else clean_tracking_no)
            if not pdf_path:
                self.print_error.emit(f"페이지 추출 실패: {tracking_no}")
                return False
        else:
            # 인덱스에 없으면 직접 파일 찾기 (하이픈 제거 버전으로 검색)
            pdf_path = self.get_pdf_path(clean_tracking_no)
            if not pdf_path.exists():
                # 원본 형식으로도 시도
                pdf_path = self.get_pdf_path(tracking_no)
                if not pdf_path.exists():
                    self.print_error.emit(f"PDF 파일 없음: {clean_tracking_no}")
                    return False
        
        try:
            import subprocess
            pdf_path_str = str(pdf_path)
            
            # win32api, win32print는 선택적 (pywin32 설치 시에만 사용)
            try:
                import win32api
                import win32print
                HAS_WIN32API = True
            except ImportError:
                HAS_WIN32API = False
            
            # 실물 프린터로 직접 인쇄 (기본 프린터 사용)
            
            # 방법 1: Adobe Reader로 실물 프린터 인쇄 (가장 확실한 방법)
            # /t 옵션: 기본 프린터로 인쇄 후 자동 종료 (사용자 클릭 불필요)
            adobe_readers = [
                r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            ]
            
            for reader_path in adobe_readers:
                if os.path.exists(reader_path):
                    try:
                        # Adobe Reader/Acrobat로 기본 프린터에 직접 인쇄 (/t: 인쇄 후 종료)
                        subprocess.Popen(
                            [reader_path, "/t", pdf_path_str],
                            shell=False,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                        if HAS_WIN32API:
                            try:
                                default_printer = win32print.GetDefaultPrinter()
                                self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} ({default_printer}로 인쇄 중...)")
                            except:
                                self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} (기본 프린터로 인쇄 중...)")
                        else:
                            self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} (기본 프린터로 인쇄 중...)")
                        return True
                    except Exception as e:
                        self.print_error.emit(f"Adobe Reader 인쇄 오류: {str(e)}")
                        continue
            
            # 방법 2: Windows 기본 PDF 뷰어 찾아서 인쇄
            try:
                import winreg
                # PDF 파일의 기본 프로그램 찾기
                try:
                    key = winreg.OpenKey(
                        winreg.HKEY_CLASSES_ROOT,
                        r".pdf\shell\print\command"
                    )
                    command = winreg.QueryValue(key, None)
                    winreg.CloseKey(key)
                    
                    # 명령어에서 실행 파일 경로 추출
                    if command:
                        # "C:\Program Files\..." "%1" 형식에서 경로 추출
                        import shlex
                        parts = shlex.split(command)
                        if parts:
                            pdf_viewer = parts[0]
                            if os.path.exists(pdf_viewer):
                                # PDF 뷰어로 인쇄 시도
                                subprocess.Popen(
                                    [pdf_viewer, "/t", pdf_path_str] if "/t" in command or "print" in command.lower() else [pdf_viewer, pdf_path_str],
                                    shell=False,
                                    stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL,
                                    creationflags=subprocess.CREATE_NO_WINDOW
                                )
                                self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} (기본 PDF 뷰어로 인쇄)")
                                return True
                except Exception:
                    pass
            except Exception:
                pass
            
            # 방법 3: Windows ShellExecute로 인쇄 시도
            if HAS_WIN32API:
                try:
                    # 기본 프린터 이름 확인
                    default_printer = win32print.GetDefaultPrinter()
                    
                    # 인쇄 동사 사용 (기본 프린터로 직접 인쇄)
                    win32api.ShellExecute(
                        0,
                        "print",
                        pdf_path_str,
                        None,
                        ".",
                        0
                    )
                    self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} ({default_printer}로 인쇄)")
                    return True
                except Exception as e:
                    self.print_error.emit(f"ShellExecute 인쇄 실패: {str(e)}")
            
            # 방법 4: os.startfile로 인쇄 시도
            try:
                os.startfile(pdf_path_str, "print")
                self.print_success.emit(f"실물 프린터 인쇄 요청: {tracking_no} (Windows 기본 인쇄 동사)")
                return True
            except (OSError, FileNotFoundError) as e:
                self.print_error.emit(f"os.startfile 인쇄 실패: {str(e)}")
            
            # 실물 프린터 인쇄 실패
            if HAS_WIN32API:
                try:
                    default_printer = win32print.GetDefaultPrinter()
                    self.print_error.emit(f"실물 프린터 인쇄 실패: {tracking_no} (기본 프린터: {default_printer})")
                except:
                    self.print_error.emit(f"실물 프린터 인쇄 실패: {tracking_no} (기본 프린터 확인 필요)")
            else:
                self.print_error.emit(f"실물 프린터 인쇄 실패: {tracking_no} (기본 프린터 확인 필요)")
            return False
            
        except FileNotFoundError:
            self.print_error.emit(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
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

