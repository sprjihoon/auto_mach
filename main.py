"""
자동출고 프로그램 - 메인 진입점
=====================================

제품 바코드를 스캔하면:
1) 엑셀에서 해당 제품의 tracking_no(송장번호)를 역매칭하여 찾고
2) EzAuto 프로그램에 자동으로 키입력을 보내고
3) tracking_no 단위로 qty/scanned_qty를 실시간 추적하고
4) 구성 수량이 모두 충족되면 PDF 송장 라벨을 자동 출력
5) 출력된 tracking_no는 used=1로 저장하여 재스캔 및 재출력 금지

사용법:
    python main.py

빌드:
    pyinstaller -F -w main.py
"""

import sys
import os

# PyInstaller 빌드 시 경로 설정
if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 경우
    os.chdir(os.path.dirname(sys.executable))
else:
    # 개발 환경
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


def main():
    """메인 함수"""
    from ui_main import run_app
    return run_app()


if __name__ == "__main__":
    sys.exit(main())

