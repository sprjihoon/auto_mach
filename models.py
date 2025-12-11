"""
데이터 모델 정의
"""
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum


class ScanResult(Enum):
    """스캔 결과 상태"""
    SUCCESS = "success"
    ALREADY_USED = "already_used"
    NOT_FOUND = "not_found"
    ERROR = "error"


@dataclass
class OrderItem:
    """주문 항목 데이터 모델"""
    tracking_no: str
    barcode: str
    product_name: str
    option_name: str
    qty: int
    scanned_qty: int = 0
    used: int = 0
    
    @property
    def remaining(self) -> int:
        """남은 스캔 수량"""
        return max(0, self.qty - self.scanned_qty)
    
    @property
    def is_complete(self) -> bool:
        """스캔 완료 여부"""
        return self.scanned_qty >= self.qty


@dataclass
class TrackingGroup:
    """송장번호 단위 그룹"""
    tracking_no: str
    items: list[OrderItem] = field(default_factory=list)
    
    @property
    def total_qty(self) -> int:
        """전체 필요 수량"""
        return sum(item.qty for item in self.items)
    
    @property
    def total_scanned(self) -> int:
        """전체 스캔된 수량"""
        return sum(item.scanned_qty for item in self.items)
    
    @property
    def remaining(self) -> int:
        """남은 수량"""
        return sum(item.remaining for item in self.items)
    
    @property
    def is_complete(self) -> bool:
        """그룹 완료 여부"""
        return self.remaining == 0
    
    @property
    def is_used(self) -> bool:
        """사용 완료 여부"""
        return any(item.used == 1 for item in self.items)


@dataclass
class ScanEvent:
    """스캔 이벤트 로그"""
    timestamp: str
    barcode: str
    tracking_no: Optional[str]
    result: ScanResult
    message: str

