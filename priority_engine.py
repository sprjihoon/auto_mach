"""
우선순위 매칭 엔진
송장 매칭 순서를 동적으로 결정하는 엔진
"""
from typing import Dict, Any
from datetime import datetime

# 우선순위 점수 상수
PRIORITY_FIXED_SCORE = 1000000  # ⭐ 고정 송장 점수 (최우선)
SINGLE_COMBO_BONUS = 10000      # 단품/조합 보너스 점수


def calc_priority_score(order_meta: Dict[str, Any], rules: Dict[str, bool]) -> int:
    """
    송장 메타데이터와 우선순위 규칙을 기반으로 우선순위 점수를 계산합니다.
    
    Args:
        order_meta: 송장 메타데이터
            {
                "tracking_no": str,
                "order_datetime": datetime or None,
                "item_count": int,  # qty 합
                "sku_count": int,   # SKU 종류 수
                "is_single": bool,  # sku_count == 1
                "is_priority": bool  # 수동 우선순위 지정
            }
        rules: 우선순위 규칙
            {
                "single_first": bool,      # 단품 우선
                "combo_first": bool,       # 조합 우선
                "small_qty_first": bool,   # 소량 우선
                "large_qty_first": bool,   # 대량 우선
                "old_order_first": bool,    # 오래된 주문 우선
                "new_order_first": bool,    # 최신 주문 우선
                "manual_priority": bool     # 수동 우선순위 활성화
            }
    
    Returns:
        우선순위 점수 (높을수록 우선순위 높음)
    """
    score = 0
    
    # 1. 수동 우선순위 (⭐ 고정 송장 - 최우선)
    # manual_priority가 활성화되고 is_priority가 True인 송장은 항상 최우선
    if rules.get("manual_priority", False) and order_meta.get("is_priority", False):
        score += PRIORITY_FIXED_SCORE  # 매우 높은 점수로 다른 모든 조건보다 우선
    
    # 2. 단품/조합 우선순위
    if rules.get("single_first", False) and order_meta.get("is_single", False):
        score += SINGLE_COMBO_BONUS  # 단품에 가산점
    elif rules.get("combo_first", False) and not order_meta.get("is_single", False):
        score += SINGLE_COMBO_BONUS  # 조합에 가산점
    
    # 3. 수량 기반 우선순위
    item_count = order_meta.get("item_count", 0)
    if rules.get("small_qty_first", False):
        # 소량 우선: item_count가 적을수록 높은 점수
        # 최대 1000개 가정, 역순으로 점수 부여
        score += max(0, 1000 - item_count)
    elif rules.get("large_qty_first", False):
        # 대량 우선: item_count가 많을수록 높은 점수
        score += item_count
    
    # 4. 주문 시간 기반 우선순위
    order_datetime = order_meta.get("order_datetime")
    if order_datetime:
        if isinstance(order_datetime, datetime):
            # 타임스탬프로 변환 (초 단위)
            timestamp = order_datetime.timestamp()
            
            if rules.get("old_order_first", False):
                # 오래된 주문 우선: 타임스탬프가 작을수록 높은 점수
                # 2020-01-01 기준으로 상대 시간 계산 (초 단위)
                base_timestamp = datetime(2020, 1, 1).timestamp()
                score += max(0, int(base_timestamp - timestamp))
            elif rules.get("new_order_first", False):
                # 최신 주문 우선: 타임스탬프가 클수록 높은 점수
                # 2020-01-01 기준으로 상대 시간 계산 (초 단위)
                base_timestamp = datetime(2020, 1, 1).timestamp()
                score += max(0, int(timestamp - base_timestamp))
    
    # 5. 기본 정렬: tracking_no (문자열 정렬로 일관성 유지)
    # 점수에 영향을 주지 않지만, 동일 점수일 때 일관된 순서 보장을 위해
    # tracking_no를 문자열로 변환하여 정렬 키에 포함할 수 있음
    
    return score


def get_default_rules() -> Dict[str, bool]:
    """
    기본 우선순위 규칙 반환 (기존 동작과 동일: 단품 우선)
    
    Returns:
        기본 규칙 딕셔너리
    """
    return {
        "single_first": True,      # 단품 우선 (기본값)
        "combo_first": False,
        "small_qty_first": False,
        "large_qty_first": False,
        "old_order_first": False,
        "new_order_first": False,
        "manual_priority": True     # ⭐ 고정 기능 활성화
    }


def get_preset_rules(preset_name: str) -> Dict[str, bool]:
    """
    프리셋 우선순위 규칙 반환
    
    Args:
        preset_name: 프리셋 이름 ("default", "backlog", "bulk")
    
    Returns:
        프리셋 규칙 딕셔너리
    """
    presets = {
        "default": {
            # [프리셋 1] 기본(단품 우선)
            "single_first": True,
            "combo_first": False,
            "small_qty_first": False,
            "large_qty_first": False,
            "old_order_first": False,
            "new_order_first": False,
            "manual_priority": True  # ⭐ 고정 기능 항상 활성화
        },
        "backlog": {
            # [프리셋 2] 밀린 주문 정리
            "single_first": False,
            "combo_first": False,
            "small_qty_first": True,
            "large_qty_first": False,
            "old_order_first": True,
            "new_order_first": False,
            "manual_priority": True
        },
        "bulk": {
            # [프리셋 3] 대량 소화
            "single_first": False,
            "combo_first": True,
            "small_qty_first": False,
            "large_qty_first": True,
            "old_order_first": False,
            "new_order_first": False,
            "manual_priority": True
        }
    }
    
    return presets.get(preset_name, get_default_rules())

