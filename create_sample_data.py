"""
샘플 데이터 생성 스크립트
테스트용 엑셀 파일과 PDF 라벨 생성
"""

import pandas as pd
from pathlib import Path


def create_sample_excel():
    """샘플 엑셀 파일 생성"""
    
    # 샘플 데이터
    data = [
        # 단품 (qty=1) - 우선순위 높음
        {"tracking_no": "1001", "barcode": "BC001", "product_name": "비타민C", "option_name": "30정", "qty": 1},
        {"tracking_no": "1002", "barcode": "BC002", "product_name": "오메가3", "option_name": "60정", "qty": 1},
        {"tracking_no": "1003", "barcode": "BC003", "product_name": "프로바이오틱스", "option_name": "30포", "qty": 1},
        
        # 소형 조합 (qty=2)
        {"tracking_no": "2001", "barcode": "BC001", "product_name": "비타민C", "option_name": "30정", "qty": 2},
        {"tracking_no": "2001", "barcode": "BC002", "product_name": "오메가3", "option_name": "60정", "qty": 2},
        
        {"tracking_no": "2002", "barcode": "BC001", "product_name": "비타민C", "option_name": "30정", "qty": 2},
        {"tracking_no": "2002", "barcode": "BC003", "product_name": "프로바이오틱스", "option_name": "30포", "qty": 2},
        
        # 대형 조합 (qty=3)
        {"tracking_no": "3001", "barcode": "BC001", "product_name": "비타민C", "option_name": "30정", "qty": 3},
        {"tracking_no": "3001", "barcode": "BC002", "product_name": "오메가3", "option_name": "60정", "qty": 3},
        {"tracking_no": "3001", "barcode": "BC003", "product_name": "프로바이오틱스", "option_name": "30포", "qty": 3},
        
        # 추가 단품
        {"tracking_no": "1004", "barcode": "BC001", "product_name": "비타민C", "option_name": "30정", "qty": 1},
        {"tracking_no": "1005", "barcode": "BC002", "product_name": "오메가3", "option_name": "60정", "qty": 1},
    ]
    
    df = pd.DataFrame(data)
    df['scanned_qty'] = 0
    df['used'] = 0
    
    # 엑셀 저장 (xlsx 형식 - xls도 읽기 지원)
    output_path = Path("sample_orders.xlsx")
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"샘플 엑셀 생성: {output_path}")
    print(f"총 {len(df)}개 행, {df['tracking_no'].nunique()}개 송장")
    print("※ xls 파일도 읽기 지원됩니다")
    
    return df


def create_sample_pdfs():
    """샘플 PDF 라벨 생성 (빈 파일)"""
    
    labels_dir = Path("labels")
    labels_dir.mkdir(exist_ok=True)
    
    # 송장번호 목록
    tracking_numbers = ["1001", "1002", "1003", "1004", "1005", "2001", "2002", "3001"]
    
    for tn in tracking_numbers:
        pdf_path = labels_dir / f"{tn}.pdf"
        # 빈 PDF 파일 생성 (실제로는 라벨 PDF가 있어야 함)
        pdf_path.write_text(f"Sample PDF for tracking_no: {tn}")
        print(f"샘플 PDF 생성: {pdf_path}")
    
    print(f"\n총 {len(tracking_numbers)}개 PDF 파일 생성")
    print("※ 실제 사용 시에는 실제 송장 라벨 PDF로 교체하세요!")


def main():
    print("=" * 50)
    print("샘플 데이터 생성")
    print("=" * 50)
    
    # 엑셀 생성
    print("\n[1] 샘플 엑셀 생성 중...")
    create_sample_excel()
    
    # PDF 생성
    print("\n[2] 샘플 PDF 생성 중...")
    create_sample_pdfs()
    
    print("\n" + "=" * 50)
    print("완료!")
    print("=" * 50)
    print("\n사용법:")
    print("1. python main.py 실행")
    print("2. sample_orders.xlsx 파일 로드 (xls도 지원)")
    print("3. PDF 폴더 선택 (labels)")
    print("4. 스캐너 시작 후 바코드 스캔")
    print("\n테스트용 바코드: BC001, BC002, BC003")


if __name__ == "__main__":
    main()

