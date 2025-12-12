# -*- coding: utf-8 -*-
import pdfplumber
import re
from pathlib import Path

pdf_path = Path(r"C:\Users\user\Desktop\송장번호.pdf")

print(f"파일 존재: {pdf_path.exists()}")

if pdf_path.exists():
    with pdfplumber.open(pdf_path) as pdf:
        print(f"페이지 수: {len(pdf.pages)}")
        
        found = {}
        # 하이픈 변형: 일반 하이픈(-), en-dash(–), em-dash(—), 공백 등
        patterns = [
            # 5-4-4 형식 (하이픈 변형 포함)
            r'\b(\d{5}[-–—\s]\d{4}[-–—\s]\d{4})\b',  # 60914-8682-2638 형식 (하이픈 변형 지원)
            # 5-4-4 형식 (일반 하이픈만)
            r'\b(\d{5}-\d{4}-\d{4})\b',  # 60914-8675-3755 형식
            # 연속 숫자 형식
            r'\b(\d{13})\b',  # 13자리 숫자
            r'\b(\d{12})\b',  # 12자리 숫자
        ]
        
        found_matches = set()  # 중복 방지용
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            
            for pattern in patterns:
                matches = re.findall(pattern, text)
                for match in matches:
                    # 모든 하이픈 변형과 공백 제거
                    clean = re.sub(r'[-–—\s]', '', match)
                    
                    # 숫자만 남았는지 확인 (최소 10자리)
                    if clean.isdigit() and len(clean) >= 10:
                        # 이미 처리한 매치는 건너뛰기
                        if clean not in found_matches:
                            found_matches.add(clean)
                            found[clean] = (i+1, match)
                            print(f"페이지 {i+1}: {match} → {clean}")
        
        print(f"\n총 {len(found)}개 송장번호 발견")

