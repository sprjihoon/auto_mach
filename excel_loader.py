"""
엑셀 로딩·저장 + DataFrame 관리
"""
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List
from datetime import datetime
from PySide6.QtCore import QObject, Signal

from utils import get_base_path, get_timestamp


class ExcelLoader(QObject):
    """엑셀 데이터 관리 클래스"""
    
    # 시그널 정의
    data_loaded = Signal()
    data_updated = Signal()
    error_occurred = Signal(str)
    priority_cleared = Signal(str)  # 완료된 우선 송장 해제 시그널
    
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
        '주문일': 'order_date',
        '주문시간': 'order_time',
        '주문번호': 'order_no',
    }
    
    # 대체 컬럼명 (상품수량이 없을 때만 사용)
    FALLBACK_QTY_COLUMNS = ['주문수량']
    
    def __init__(self):
        super().__init__()
        self.df: Optional[pd.DataFrame] = None
        self.file_path: Optional[Path] = None
        # 송장별 메타데이터 캐시 (성능 최적화)
        self._metadata_cache: Optional[Dict[str, Dict[str, Any]]] = None
        # 우선순위 규칙 (기본값: 단품 우선)
        self._priority_rules: Optional[Dict[str, bool]] = None
        # 송장별 ⭐ 고정 상태 저장 (tracking_no -> is_priority)
        self._priority_tracking: Dict[str, bool] = {}
    
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
            
            # order_datetime 컬럼 생성
            if 'order_datetime' not in self.df.columns:
                # 1. 주문번호(order_no) 컬럼이 있으면 주문번호 순서로 생성 (가장 우선)
                has_order_no = 'order_no' in self.df.columns
                
                if has_order_no:
                    # 주문번호를 기준으로 정렬하여 순서 결정
                    # 주문번호 형식 예: "20251212-0000051" (날짜-순서)
                    def parse_order_datetime_from_order_no(row):
                        """주문번호에서 datetime 추출 또는 주문번호 순서로 datetime 생성"""
                        try:
                            order_no = row['order_no'] if 'order_no' in row.index else None
                            if pd.notna(order_no):
                                order_no_str = str(order_no).strip()
                                
                                # 주문번호 형식 파싱 시도 (예: "20251212-0000051")
                                # 앞부분이 날짜 형식인지 확인 (YYYYMMDD)
                                if '-' in order_no_str:
                                    date_part = order_no_str.split('-')[0]
                                    if len(date_part) == 8 and date_part.isdigit():
                                        # YYYYMMDD 형식 파싱
                                        year = int(date_part[:4])
                                        month = int(date_part[4:6])
                                        day = int(date_part[6:8])
                                        # 기본 시간으로 datetime 생성 (순서는 주문번호 정렬로 결정)
                                        return datetime(year, month, day, 0, 0, 0)
                                
                                # 숫자만 있는 경우 (예: "202512120000051")
                                if order_no_str.isdigit() and len(order_no_str) >= 8:
                                    # 앞 8자리가 날짜일 가능성
                                    date_part = order_no_str[:8]
                                    year = int(date_part[:4])
                                    month = int(date_part[4:6])
                                    day = int(date_part[6:8])
                                    return datetime(year, month, day, 0, 0, 0)
                        except:
                            pass
                        return None
                    
                    # 각 행에 대해 주문번호에서 datetime 추출 시도
                    self.df['order_datetime'] = self.df.apply(parse_order_datetime_from_order_no, axis=1)
                    
                    # 주문번호에서 날짜를 추출하지 못한 경우, 주문번호 순서로 정렬하여 datetime 생성
                    if self.df['order_datetime'].isna().any():
                        # 주문번호로 정렬하여 순서 결정
                        order_no_sorted = self.df['order_no'].drop_duplicates().sort_values().reset_index(drop=True)
                        order_no_to_idx = {no: idx for idx, no in enumerate(order_no_sorted)}
                        
                        # tracking_no별로 첫 번째 주문번호 사용
                        tracking_to_order_no = {}
                        for tracking_no, group in self.df.groupby('tracking_no'):
                            first_order_no = group['order_no'].iloc[0] if 'order_no' in group.columns else None
                            if pd.notna(first_order_no):
                                tracking_to_order_no[tracking_no] = first_order_no
                        
                        # 주문번호 순서를 datetime으로 변환 (2020-01-01부터 시작, 하루 간격)
                        base_date = datetime(2020, 1, 1)
                        def get_datetime_from_order_no(tracking_no):
                            order_no = tracking_to_order_no.get(tracking_no)
                            if order_no is None:
                                return None
                            idx = order_no_to_idx.get(order_no, 0)
                            return base_date.replace(day=1 + (idx % 28))
                        
                        # order_datetime이 없는 행에 대해 주문번호 순서로 생성
                        mask_na = self.df['order_datetime'].isna()
                        if mask_na.any():
                            self.df.loc[mask_na, 'order_datetime'] = self.df.loc[mask_na, 'tracking_no'].map(get_datetime_from_order_no)
                    
                    # 여전히 None이 있으면 로딩 순서로 대체
                    if self.df['order_datetime'].isna().any():
                        tracking_order = self.df['tracking_no'].drop_duplicates().reset_index(drop=True)
                        order_dict = {tn: idx for idx, tn in enumerate(tracking_order)}
                        base_date = datetime(2020, 1, 1)
                        mask_na = self.df['order_datetime'].isna()
                        self.df.loc[mask_na, 'order_datetime'] = self.df.loc[mask_na, 'tracking_no'].map(
                            lambda tn: base_date.replace(day=1 + (order_dict.get(tn, 0) % 28))
                        )
                
                # 2. 주문번호가 없으면 주문일(order_date)과 주문시간(order_time) 컬럼 확인
                elif 'order_date' in self.df.columns or 'order_time' in self.df.columns:
                    has_order_date = 'order_date' in self.df.columns
                    has_order_time = 'order_time' in self.df.columns
                    
                    if has_order_date or has_order_time:
                    def combine_datetime(row):
                        """주문일과 주문시간을 합쳐서 datetime 생성"""
                        order_date = None
                        order_time = None
                        
                        # 주문일 찾기 (pandas Series는 인덱스로 접근)
                        if has_order_date:
                            try:
                                order_date = row['order_date'] if 'order_date' in row.index else None
                            except:
                                order_date = None
                        # 주문시간 찾기
                        if has_order_time:
                            try:
                                order_time = row['order_time'] if 'order_time' in row.index else None
                            except:
                                order_time = None
                        
                        # 주문일 처리
                        if pd.notna(order_date):
                            if isinstance(order_date, datetime):
                                date_part = order_date
                            elif isinstance(order_date, str):
                                try:
                                    date_part = pd.to_datetime(order_date).to_pydatetime()
                                except:
                                    date_part = None
                            else:
                                try:
                                    date_part = pd.to_datetime(order_date).to_pydatetime()
                                except:
                                    date_part = None
                        else:
                            date_part = None
                        
                        # 주문시간 처리
                        if pd.notna(order_time):
                            if isinstance(order_time, datetime):
                                time_part = order_time
                            elif isinstance(order_time, str):
                                try:
                                    # 시간 문자열 파싱 (예: "10:30:00", "10:30")
                                    time_str = str(order_time).strip()
                                    if ':' in time_str:
                                        parts = time_str.split(':')
                                        if len(parts) >= 2:
                                            hour = int(parts[0])
                                            minute = int(parts[1])
                                            second = int(parts[2]) if len(parts) > 2 else 0
                                            time_part = datetime(1900, 1, 1, hour, minute, second).time()
                                        else:
                                            time_part = None
                                    else:
                                        time_part = None
                                except:
                                    time_part = None
                            else:
                                try:
                                    time_part = pd.to_datetime(order_time).to_pydatetime().time()
                                except:
                                    time_part = None
                        else:
                            time_part = None
                        
                        # 날짜와 시간 합치기
                        if date_part:
                            if isinstance(date_part, datetime):
                                if time_part:
                                    if isinstance(time_part, datetime):
                                        return time_part.replace(year=date_part.year, month=date_part.month, day=date_part.day)
                                    elif hasattr(time_part, 'hour'):
                                        return date_part.replace(hour=time_part.hour, minute=time_part.minute, second=time_part.second)
                                return date_part
                            else:
                                try:
                                    date_dt = pd.to_datetime(date_part).to_pydatetime()
                                    if time_part and hasattr(time_part, 'hour'):
                                        return date_dt.replace(hour=time_part.hour, minute=time_part.minute, second=time_part.second)
                                    return date_dt
                                except:
                                    pass
                        
                        return None
                    
                    # 각 행에 대해 datetime 생성
                    self.df['order_datetime'] = self.df.apply(combine_datetime, axis=1)
                    
                    # 생성 실패한 경우를 위해 로딩 순서 기반으로 대체
                    if self.df['order_datetime'].isna().all():
                        # 각 tracking_no의 첫 번째 출현 순서를 기반으로 datetime 생성
                        tracking_order = self.df['tracking_no'].drop_duplicates().reset_index(drop=True)
                        order_dict = {tn: idx for idx, tn in enumerate(tracking_order)}
                        base_date = datetime(2020, 1, 1)
                        self.df['order_datetime'] = self.df['tracking_no'].map(
                            lambda tn: base_date.replace(day=1 + (order_dict.get(tn, 0) % 28))
                        )
                else:
                    # 주문일/주문시간 컬럼도 없으면 로딩 순서 기반으로 생성
                    tracking_order = self.df['tracking_no'].drop_duplicates().reset_index(drop=True)
                    order_dict = {tn: idx for idx, tn in enumerate(tracking_order)}
                    base_date = datetime(2020, 1, 1)
                    self.df['order_datetime'] = self.df['tracking_no'].map(
                        lambda tn: base_date.replace(day=1 + (order_dict.get(tn, 0) % 28))
                    )
            
            # 메타데이터 캐시 초기화
            self._metadata_cache = None
            # ⭐ 고정 상태는 유지 (엑셀 재로드 시에도 보존)
            
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
    
    def set_priority_rules(self, rules: Dict[str, bool]):
        """
        우선순위 규칙 설정
        
        Args:
            rules: 우선순위 규칙 딕셔너리
        """
        self._priority_rules = rules
        # 메타데이터 캐시 무효화 (규칙 변경 시 재계산 필요)
        self._metadata_cache = None
    
    def get_order_metadata(self, tracking_no: str) -> Dict[str, Any]:
        """
        송장 메타데이터 수집
        
        Args:
            tracking_no: 송장번호
        
        Returns:
            메타데이터 딕셔너리
        """
        if self.df is None:
            return {}
        
        # 캐시 확인
        if self._metadata_cache is None:
            self._build_metadata_cache()
        
        return self._metadata_cache.get(tracking_no, {})
    
    def _build_metadata_cache(self):
        """모든 송장의 메타데이터 캐시 구축"""
        if self.df is None:
            self._metadata_cache = {}
            return
        
        self._metadata_cache = {}
        pending = self.df[self.df['used'] == 0]
        
        if pending.empty:
            return
        
        # tracking_no별 그룹화
        for tracking_no, group in pending.groupby('tracking_no'):
            # order_datetime: 첫 번째 행의 order_datetime 사용
            order_datetime = None
            if 'order_datetime' in group.columns:
                first_datetime = group['order_datetime'].iloc[0]
                if pd.notna(first_datetime):
                    if isinstance(first_datetime, datetime):
                        order_datetime = first_datetime
                    elif isinstance(first_datetime, str):
                        try:
                            dt = pd.to_datetime(first_datetime)
                            # pandas Timestamp를 Python datetime으로 변환
                            order_datetime = dt.to_pydatetime() if hasattr(dt, 'to_pydatetime') else dt
                        except:
                            pass
                    elif hasattr(first_datetime, 'to_pydatetime'):
                        # pandas Timestamp인 경우
                        try:
                            order_datetime = first_datetime.to_pydatetime()
                        except:
                            pass
            
            # item_count: qty 합계
            item_count = int(group['qty'].sum())
            
            # sku_count: 고유 바코드 수
            sku_count = group['barcode'].nunique()
            
            # is_single: 단품 여부
            is_single = (sku_count == 1)
            
            # is_priority: 수동 우선순위 (⭐ 고정 상태)
            # _priority_tracking에서 조회, 없으면 False
            is_priority = self._priority_tracking.get(tracking_no, False)
            
            self._metadata_cache[tracking_no] = {
                "tracking_no": tracking_no,
                "order_datetime": order_datetime,
                "item_count": item_count,
                "sku_count": sku_count,
                "is_single": is_single,
                "is_priority": is_priority
            }
    
    def find_candidates(self, barcode: str, priority_rules: Optional[Dict[str, bool]] = None) -> pd.DataFrame:
        """
        바코드로 후보 검색 후 우선순위 엔진을 사용하여 정렬
        
        [기존 단품 우선 로직 제거됨]
        이제 priority_engine을 사용하여 동적으로 우선순위 결정
        
        Args:
            barcode: 검색할 바코드
            priority_rules: 우선순위 규칙 (None이면 기본 규칙 사용)
        
        Returns:
            우선순위 정렬된 후보 DataFrame
        """
        from priority_engine import calc_priority_score, get_default_rules
        
        candidates = self.find_by_barcode(barcode)
        if candidates.empty:
            return candidates
        
        # 우선순위 규칙 결정
        if priority_rules is None:
            if self._priority_rules is None:
                priority_rules = get_default_rules()
            else:
                priority_rules = self._priority_rules
        
        # 메타데이터 캐시 구축 (필요시)
        if self._metadata_cache is None:
            self._build_metadata_cache()
        
        # 각 후보에 대해 우선순위 점수 계산
        candidates = candidates.copy()
        scores = []
        
        for _, row in candidates.iterrows():
            tracking_no = str(row['tracking_no'])
            meta = self.get_order_metadata(tracking_no)
            score = calc_priority_score(meta, priority_rules)
            scores.append(score)
        
        candidates['_priority_score'] = scores
        
        # 우선순위 점수 내림차순 정렬, 동일 점수면 tracking_no 오름차순
        candidates = candidates.sort_values(
            by=['_priority_score', 'tracking_no'],
            ascending=[False, True]
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
                # 메타데이터 캐시 무효화 (데이터 변경 시)
                self._metadata_cache = None
                self.data_updated.emit()
                return True
            return False
            
        except Exception as e:
            self.error_occurred.emit(f"스캔 수량 업데이트 오류: {str(e)}")
            return False
    
    def mark_used(self, tracking_no: str) -> bool:
        """
        tracking_no 그룹 전체를 used=1로 표시
        완료된 우선 송장(⭐)은 자동으로 해제됨
        """
        if self.df is None:
            return False
        
        try:
            mask = self.df['tracking_no'] == tracking_no
            self.df.loc[mask, 'used'] = 1
            
            # 완료된 우선 송장 자동 해제
            self._clear_priority_if_completed(tracking_no)
            
            # 메타데이터 캐시 무효화 (데이터 변경 시)
            self._metadata_cache = None
            self.data_updated.emit()
            self.save_excel()  # 즉시 저장
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"used 업데이트 오류: {str(e)}")
            return False
    
    def _clear_priority_if_completed(self, tracking_no: str):
        """
        완료된 우선 송장 자동 해제
        
        Args:
            tracking_no: 완료된 송장번호
        """
        # is_priority가 True인 경우만 해제
        if self.get_tracking_priority(tracking_no):
            self.set_tracking_priority(tracking_no, False)
            # 메타데이터 캐시에서도 제거 (이미 무효화되지만 명시적으로 처리)
            if self._metadata_cache and tracking_no in self._metadata_cache:
                self._metadata_cache[tracking_no]["is_priority"] = False
            # UI 업데이트를 위한 시그널 발생
            self.priority_cleared.emit(tracking_no)
    
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
    
    def set_tracking_priority(self, tracking_no: str, is_priority: bool):
        """
        송장 ⭐ 고정 상태 설정
        
        Args:
            tracking_no: 송장번호
            is_priority: True면 ⭐ 고정, False면 해제
        """
        self._priority_tracking[tracking_no] = is_priority
        # 메타데이터 캐시 무효화 (다음 조회 시 갱신)
        if self._metadata_cache and tracking_no in self._metadata_cache:
            # 해당 송장만 캐시에서 제거
            del self._metadata_cache[tracking_no]
    
    def get_tracking_priority(self, tracking_no: str) -> bool:
        """
        송장 ⭐ 고정 상태 조회
        
        Args:
            tracking_no: 송장번호
        
        Returns:
            ⭐ 고정 여부
        """
        return self._priority_tracking.get(tracking_no, False)
    
    def get_all_tracking_numbers(self) -> List[str]:
        """
        모든 송장번호 목록 반환 (used=0인 것만)
        
        Returns:
            송장번호 리스트
        """
        if self.df is None:
            return []
        
        pending = self.df[self.df['used'] == 0]
        if pending.empty:
            return []
        
        return pending['tracking_no'].drop_duplicates().tolist()

