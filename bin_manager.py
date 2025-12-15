"""
BIN 관리 모듈
SKU별 BIN 자동 배정 및 송장별 BIN 매핑
"""
from typing import Dict, List, Optional, Tuple
from PySide6.QtCore import QObject, Signal
import pandas as pd


class BinManager(QObject):
    """BIN 주소 관리 클래스"""
    
    # 시그널
    bin_updated = Signal()  # BIN 정보 갱신됨
    bin_reset = Signal()    # BIN 전체 리셋됨
    
    def __init__(self):
        super().__init__()
        # SKU(바코드) → BIN 매핑
        self._sku_bin_map: Dict[str, str] = {}
        # 송장번호 → BIN 매핑
        self._order_bin_map: Dict[str, str] = {}
        # BIN 카운터
        self._bin_counter: int = 0
        # 초기화 여부
        self._initialized: bool = False
    
    @property
    def is_initialized(self) -> bool:
        """BIN 시스템 초기화 여부"""
        return self._initialized
    
    def reset(self):
        """
        BIN 전체 리셋
        - 엑셀 로드 시 반드시 호출
        - 모든 BIN 정보 초기화
        """
        self._sku_bin_map.clear()
        self._order_bin_map.clear()
        self._bin_counter = 0
        self._initialized = False
        self.bin_reset.emit()
    
    def assign_bins_from_dataframe(self, df: pd.DataFrame) -> int:
        """
        DataFrame에서 SKU별 BIN 자동 배정
        
        1) SKU별 총 수량 합계 계산
        2) 총 수량 기준 내림차순 정렬
        3) 정렬 순서대로 BIN-01, BIN-02... 배정
        
        Args:
            df: 엑셀 DataFrame (barcode, qty 컬럼 필수)
        
        Returns:
            배정된 BIN 개수
        """
        # 리셋
        self.reset()
        
        if df is None or df.empty:
            return 0
        
        # used=0인 미처리 항목만 대상
        pending = df[df['used'] == 0] if 'used' in df.columns else df
        
        if pending.empty:
            return 0
        
        # SKU(바코드)별 총 수량 집계
        sku_qty = pending.groupby('barcode')['qty'].sum().reset_index()
        sku_qty.columns = ['barcode', 'total_qty']
        
        # 총 수량 내림차순 정렬
        sku_qty = sku_qty.sort_values('total_qty', ascending=False).reset_index(drop=True)
        
        # BIN 배정
        for idx, row in sku_qty.iterrows():
            barcode = str(row['barcode']).strip()
            if barcode and barcode != 'nan':
                self._bin_counter += 1
                bin_id = f"BIN-{self._bin_counter:02d}"
                self._sku_bin_map[barcode] = bin_id
        
        self._initialized = True
        self.bin_updated.emit()
        
        return len(self._sku_bin_map)
    
    def get_sku_bin(self, barcode: str) -> str:
        """
        SKU(바코드)의 BIN 주소 조회
        
        Args:
            barcode: 바코드
        
        Returns:
            BIN 주소 (예: "BIN-01") 또는 "BIN 미지정"
        """
        if not self._initialized:
            return "BIN 미지정"
        
        barcode = str(barcode).strip()
        return self._sku_bin_map.get(barcode, "BIN 미지정")
    
    def build_order_bin_map(self, df: pd.DataFrame):
        """
        송장별 BIN 매핑 구축
        - 각 송장의 대표 SKU 결정 (첫 번째 바코드)
        - sku_bin_map을 사용하여 order_bin_map 생성
        
        Args:
            df: 엑셀 DataFrame
        """
        self._order_bin_map.clear()
        
        if df is None or df.empty or not self._initialized:
            return
        
        # used=0인 미처리 항목만 대상
        pending = df[df['used'] == 0] if 'used' in df.columns else df
        
        if pending.empty:
            return
        
        # 송장별 첫 번째 바코드를 대표 SKU로 사용
        for tracking_no, group in pending.groupby('tracking_no'):
            tracking_no_str = str(tracking_no).strip()
            
            # 첫 번째 바코드 (대표 SKU)
            first_barcode = str(group.iloc[0]['barcode']).strip()
            
            # BIN 조회
            bin_id = self._sku_bin_map.get(first_barcode, "BIN 미지정")
            self._order_bin_map[tracking_no_str] = bin_id
        
        self.bin_updated.emit()
    
    def get_order_bin(self, tracking_no: str) -> str:
        """
        송장번호의 BIN 주소 조회
        
        Args:
            tracking_no: 송장번호
        
        Returns:
            BIN 주소 (예: "BIN-01") 또는 "BIN 미지정"
        """
        if not self._initialized:
            return "BIN 미지정"
        
        tracking_no = str(tracking_no).strip()
        return self._order_bin_map.get(tracking_no, "BIN 미지정")
    
    def get_all_sku_bins(self) -> List[Tuple[str, str, int]]:
        """
        모든 SKU-BIN 매핑 목록 반환 (정렬용)
        
        Returns:
            [(barcode, bin_id, bin_number), ...] 리스트
        """
        result = []
        for barcode, bin_id in self._sku_bin_map.items():
            # BIN-01 → 1
            try:
                bin_num = int(bin_id.split('-')[1])
            except:
                bin_num = 999
            result.append((barcode, bin_id, bin_num))
        
        # BIN 번호 오름차순 정렬
        result.sort(key=lambda x: x[2])
        return result
    
    def get_sku_bin_map(self) -> Dict[str, str]:
        """SKU-BIN 매핑 딕셔너리 반환"""
        return self._sku_bin_map.copy()
    
    def get_order_bin_map(self) -> Dict[str, str]:
        """송장-BIN 매핑 딕셔너리 반환"""
        return self._order_bin_map.copy()
    
    def get_bin_count(self) -> int:
        """배정된 BIN 개수"""
        return len(self._sku_bin_map)

