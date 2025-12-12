# -*- coding: utf-8 -*-
"""
엑셀 송장번호와 PDF 페이지 매핑 확인 스크립트
실제 매핑이 정확한지 확인하고 수정
"""
import pandas as pd
from pathlib import Path
import re

def check_mapping():
    """엑셀과 PDF 매핑 확인"""
    # 엑셀 파일 경로 (실제 경로로 수정 필요)
    excel_path = Path(r"C:/Users/one/Downloads/확장주문검색_20251212151012_812996788.xls")
    
    if not excel_path.exists():
        print(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    # 엑셀 로드 (여러 방법 시도)
    try:
        # 방법 1: HTML로 읽기 (xls가 HTML 형식인 경우)
        try:
            df = pd.read_html(str(excel_path))[0]  # 첫 번째 테이블
            print(f"엑셀 로드 성공 (HTML): {len(df)}행")
        except:
            # 방법 2: CSV로 읽기
            try:
                df = pd.read_csv(excel_path, encoding='cp949')
                print(f"엑셀 로드 성공 (CSV): {len(df)}행")
            except:
                # 방법 3: Excel 엔진들
                try:
                    df = pd.read_excel(excel_path, engine='xlrd')
                    print(f"엑셀 로드 성공 (xlrd): {len(df)}행")
                except:
                    df = pd.read_excel(excel_path, engine='openpyxl')
                    print(f"엑셀 로드 성공 (openpyxl): {len(df)}행")
        
        if 'tracking_no' not in df.columns:
            print("tracking_no 컬럼을 찾을 수 없습니다.")
            print(f"사용 가능한 컬럼: {df.columns.tolist()}")
            return
        
        # 송장번호 순서 확인 (중복 제거하되 순서 보장)
        tracking_numbers = df['tracking_no'].drop_duplicates().tolist()
        
        print(f"\n엑셀 송장번호 순서 ({len(tracking_numbers)}개):")
        for i, tracking_no in enumerate(tracking_numbers):
            print(f"  {i + 1}. {tracking_no}")
        
        print(f"\n현재 매핑 방식:")
        print(f"  엑셀 순서 1번 → PDF 페이지 1")
        print(f"  엑셀 순서 2번 → PDF 페이지 2")
        print(f"  ...")
        
        print(f"\n⚠️ 문제:")
        print(f"  - 6091486822635 (엑셀 순서 ?)번이 출력되어야 함")
        print(f"  - 6091486822642 (엑셀 순서 ?)번 페이지가 잘못 출력됨")
        
        # 해당 송장번호들의 엑셀 순서 찾기
        target_numbers = ['6091486822635', '6091486822642']
        for target in target_numbers:
            try:
                index = tracking_numbers.index(target)
                print(f"  - {target}: 엑셀 순서 {index + 1}번 → PDF 페이지 {index + 1}")
            except ValueError:
                # 하이픈 제거 버전으로 시도
                clean_target = re.sub(r'[-–—\s]', '', target)
                found = False
                for i, tracking_no in enumerate(tracking_numbers):
                    clean_no = re.sub(r'[-–—\s]', '', str(tracking_no))
                    if clean_no == clean_target:
                        print(f"  - {target} (정규화): 엑셀 순서 {i + 1}번 → PDF 페이지 {i + 1}")
                        found = True
                        break
                if not found:
                    print(f"  - {target}: 엑셀에서 찾을 수 없음")
        
    except Exception as e:
        print(f"엑셀 처리 실패: {str(e)}")

if __name__ == "__main__":
    check_mapping()
