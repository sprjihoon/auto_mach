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
        patterns = [
            r'\b(\d{5}-\d{4}-\d{4})\b',  # 60914-8675-3755 형식
            r'\b(\d{13})\b',  # 13자리 숫자
            r'\b(\d{12})\b',  # 12자리 숫자
        ]
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            
            for pattern in patterns:
                matches = re.findall(pattern, text)
                for match in matches:
                    clean = match.replace('-', '')
                    if clean not in found:
                        found[clean] = (i+1, match)
                        print(f"페이지 {i+1}: {match} → {clean}")
        
        print(f"\n총 {len(found)}개 송장번호 발견")

