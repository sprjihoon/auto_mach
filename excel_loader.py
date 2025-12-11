"""
엑셀 로딩·저장 + DataFrame 관리
"""
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple
from PySide6.QtCore import QObject, Signal

from utils import get_base_path, get_timestamp


class ExcelLoader(QObject):
    """엑셀 데이터 관리 클래스"""
    
    # 시그널 정의
    data_loaded = Signal()
    data_updated = Signal()
    error_occurred = Signal(str)
    
    # 필수 컬럼 (영어)
    REQUIRED_COLUMNS = ['tracking_no', 'barcode', 'product_name', 'option_name', 'qty']
    
    # 한글 → 영어 컬럼 매핑 (상품수량 우선)
    COLUMN_MAPPING = {
        '송장번호': 'tracking_no',
        '바코드': 'barcode',
        '상품명': 'product_name',
        '옵션명': 'option_name',
        '상품수량': 'qty',
        '로케이션': 'location',
        '위치': 'location',
    }
    
    # 대체 컬럼명 (상품수량이 없을 때만 사용)
    FALLBACK_QTY_COLUMNS = ['주문수량']
    
    def __init__(self):
        super().__init__()
        self.df: Optional[pd.DataFrame] = None
        self.file_path: Optional[Path] = None
    
    def load_excel(self, file_path: str) -> bool:
        """엑셀 파일 로드 (xls, xlsx, csv, html 등 지원)"""
        try:
            path = Path(file_path)
            if not path.exists():
                self.error_occurred.emit(f"파일을 찾을 수 없습니다: {file_path}")
                return False
            
            # 파일 내용 확인
            with open(path, 'rb') as f:
                header = f.read(500)
            
            df = None
            last_error = None
            
            # 파일 시그니처로 형식 판단
            header_stripped = header.lstrip()
            
            if header[:4] == b'PK\x03\x04':
                # ZIP (xlsx)
                try:
                    df = pd.read_excel(path, engine='openpyxl')
                except Exception as e:
                    last_error = e
            elif header[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                # OLE2 (xls)
                try:
                    df = pd.read_excel(path, engine='xlrd')
                except Exception as e:
                    last_error = e
            elif header_stripped[:1] == b'<' or b'<html' in header.lower() or b'<table' in header.lower() or b'<meta' in header.lower():
                # HTML 형식 (.xls로 저장된 HTML)
                encodings = ['utf-8', 'cp949', 'euc-kr']
                for enc in encodings:
                    try:
                        dfs = pd.read_html(path, encoding=enc, header=0)
                        if dfs:
                            df = dfs[0]
                            # 첫 번째 행이 데이터인 경우 (헤더가 숫자인 경우)
                            if all(isinstance(c, (int, float)) for c in df.columns):
                                # 첫 번째 행을 헤더로 사용
                                df.columns = df.iloc[0]
                                df = df.iloc[1:].reset_index(drop=True)
                            break
                    except Exception as e:
                        last_error = e
                        continue
            else:
                # CSV 또는 기타 텍스트 형식 시도
                encodings = ['utf-8', 'cp949', 'euc-kr']
                for enc in encodings:
                    try:
                        df = pd.read_csv(path, encoding=enc)
                        break
                    except Exception as e:
                        last_error = e
                        continue
            
            # 위 방법 모두 실패 시 순차 시도
            if df is None:
                # HTML 재시도
                encodings = ['utf-8', 'cp949', 'euc-kr']
                for enc in encodings:
                    try:
                        dfs = pd.read_html(path, encoding=enc, header=0)
                        if dfs:
                            df = dfs[0]
                            # 첫 번째 행이 데이터인 경우 (헤더가 숫자인 경우)
                            if all(isinstance(c, (int, float)) for c in df.columns):
                                df.columns = df.iloc[0]
                                df = df.iloc[1:].reset_index(drop=True)
                            break
                    except Exception as e:
                        last_error = e
                        continue
            
            if df is None:
                engines = ['openpyxl', 'xlrd']
                for engine in engines:
                    try:
                        df = pd.read_excel(path, engine=engine)
                        break
                    except Exception as e:
                        last_error = e
                        continue
            
            if df is None:
                self.error_occurred.emit(f"엑셀 파일을 읽을 수 없습니다: {last_error}")
                return False
            
            self.df = df
            self.file_path = path
            
            # 컬럼명을 모두 문자열로 변환
            self.df.columns = [str(col) for col in self.df.columns]
            
            # 한글 컬럼명 → 영어 컬럼명 매핑
            rename_dict = {}
            for kor, eng in self.COLUMN_MAPPING.items():
                if kor in self.df.columns and eng not in self.df.columns:
                    rename_dict[kor] = eng
            
            if rename_dict:
                self.df = self.df.rename(columns=rename_dict)
            
            # qty 컬럼이 없으면 대체 컬럼 사용
            if 'qty' not in self.df.columns:
                for fallback_col in self.FALLBACK_QTY_COLUMNS:
                    if fallback_col in self.df.columns:
                        self.df = self.df.rename(columns={fallback_col: 'qty'})
                        break
            
            # 필수 컬럼 확인
            missing_cols = [col for col in self.REQUIRED_COLUMNS if col not in self.df.columns]
            if missing_cols:
                # 현재 컬럼 목록 출력
                available = ', '.join([str(c) for c in self.df.columns[:10].tolist()])
                self.error_occurred.emit(f"필수 컬럼 누락: {', '.join(missing_cols)}\n현재 컬럼: {available}...")
                return False
            
            # scanned_qty 컬럼이 없으면 추가
            if 'scanned_qty' not in self.df.columns:
                self.df['scanned_qty'] = 0
            
            # used 컬럼이 없으면 추가
            if 'used' not in self.df.columns:
                self.df['used'] = 0
            
            # 데이터 타입 정리
            self.df['tracking_no'] = self.df['tracking_no'].astype(str)
            self.df['barcode'] = self.df['barcode'].astype(str)
            self.df['qty'] = self.df['qty'].astype(int)
            self.df['scanned_qty'] = self.df['scanned_qty'].fillna(0).astype(int)
            self.df['used'] = self.df['used'].fillna(0).astype(int)
            
            self.data_loaded.emit()
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"엑셀 로드 오류: {str(e)}")
            return False
    
    def save_excel(self, save_path: str = None) -> Tuple[bool, str]:
        """현재 DataFrame을 엑셀로 저장 (xlsx로 저장, 파일명 뒤에 _역매칭 추가)
        
        Args:
            save_path: 저장할 파일 경로. None이면 원본 파일 기반으로 _역매칭 추가
            
        Returns:
            (성공 여부, 저장된 파일 경로)
        """
        if self.df is None:
            return False, ""
        
        try:
            if save_path:
                # 지정된 경로로 저장
                target_path = Path(save_path)
            elif self.file_path:
                # 원본 파일명에 _역매칭 추가
                stem = self.file_path.stem  # 확장자 제외 파일명
                
                # 이미 _역매칭이 있으면 추가하지 않음
                if not stem.endswith('_역매칭'):
                    stem = f"{stem}_역매칭"
                
                target_path = self.file_path.parent / f"{stem}.xlsx"
            else:
                self.error_occurred.emit("저장할 파일 경로가 없습니다")
                return False, ""
            
            # 저장 (항상 openpyxl 사용)
            self.df.to_excel(target_path, index=False, engine='openpyxl')
            return True, str(target_path)
            
        except Exception as e:
            self.error_occurred.emit(f"엑셀 저장 오류: {str(e)}")
            return False, ""
    
    def find_by_barcode(self, barcode: str) -> pd.DataFrame:
        """바코드로 행 검색 (used=0인 것만)"""
        if self.df is None:
            return pd.DataFrame()
        
        # 바코드를 문자열로 변환하고 공백 제거하여 비교
        barcode = str(barcode).strip()
        df_barcodes = self.df['barcode'].astype(str).str.strip()
        
        mask = (df_barcodes == barcode) & (self.df['used'] == 0)
        return self.df[mask].copy()
    
    def find_candidates(self, barcode: str) -> pd.DataFrame:
        """
        바코드로 후보 검색 후 우선순위 정렬
        단품(구성 적은 것) → 소형조합 → 대형조합 순서
        동일 구성 수면 tracking_no 오름차순
        """
        candidates = self.find_by_barcode(barcode)
        if candidates.empty:
            return candidates
        
        # 각 tracking_no의 전체 구성 수 계산
        tracking_counts = self.df[self.df['used'] == 0].groupby('tracking_no').size()
        
        # 후보에 구성 수 추가
        candidates = candidates.copy()
        candidates['_total_items'] = candidates['tracking_no'].map(tracking_counts)
        
        # 우선순위 정렬: 구성 수 오름차순, tracking_no 오름차순
        candidates = candidates.sort_values(
            by=['_total_items', 'tracking_no'],
            ascending=[True, True]
        ).reset_index(drop=False)  # 원본 인덱스 유지
        
        return candidates
    
    def get_tracking_group(self, tracking_no: str) -> pd.DataFrame:
        """tracking_no로 그룹 조회"""
        if self.df is None:
            return pd.DataFrame()
        
        mask = self.df['tracking_no'] == tracking_no
        return self.df[mask].copy()
    
    def get_group_remaining(self, tracking_no: str) -> int:
        """tracking_no 그룹의 남은 수량 계산"""
        group = self.get_tracking_group(tracking_no)
        if group.empty:
            return 0
        
        return int((group['qty'] - group['scanned_qty']).clip(lower=0).sum())
    
    def is_tracking_used(self, tracking_no: str) -> bool:
        """tracking_no가 이미 사용되었는지 확인"""
        group = self.get_tracking_group(tracking_no)
        if group.empty:
            return False
        return group['used'].any() == 1
    
    def increment_scanned(self, original_index: int) -> bool:
        """scanned_qty 증가 (원본 DataFrame 인덱스 사용)"""
        if self.df is None:
            return False
        
        try:
            current_scanned = self.df.at[original_index, 'scanned_qty']
            current_qty = self.df.at[original_index, 'qty']
            
            # qty 초과 방지
            if current_scanned < current_qty:
                self.df.at[original_index, 'scanned_qty'] = current_scanned + 1
                self.data_updated.emit()
                return True
            return False
            
        except Exception as e:
            self.error_occurred.emit(f"스캔 수량 업데이트 오류: {str(e)}")
            return False
    
    def mark_used(self, tracking_no: str) -> bool:
        """tracking_no 그룹 전체를 used=1로 표시"""
        if self.df is None:
            return False
        
        try:
            mask = self.df['tracking_no'] == tracking_no
            self.df.loc[mask, 'used'] = 1
            self.data_updated.emit()
            self.save_excel()  # 즉시 저장
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"used 업데이트 오류: {str(e)}")
            return False
    
    def get_all_pending(self) -> pd.DataFrame:
        """처리되지 않은 모든 항목 조회"""
        if self.df is None:
            return pd.DataFrame()
        
        return self.df[self.df['used'] == 0].copy()
    
    def get_summary_by_barcode(self) -> pd.DataFrame:
        """바코드별 요약 (남은 수량)"""
        if self.df is None:
            return pd.DataFrame()
        
        pending = self.get_all_pending()
        if pending.empty:
            return pd.DataFrame()
        
        pending['remaining'] = pending['qty'] - pending['scanned_qty']
        
        summary = pending.groupby(['barcode', 'product_name', 'option_name']).agg({
            'qty': 'sum',
            'scanned_qty': 'sum',
            'remaining': 'sum'
        }).reset_index()
        
        return summary[summary['remaining'] > 0]

