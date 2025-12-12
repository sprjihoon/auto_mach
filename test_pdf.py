# -*- coding: utf-8 -*-
"""PDF 송장번호 추출 테스트 스크립트"""
import pdfplumber
import re
from pathlib import Path

pdf_path = Path(r"C:/Users/one/Desktop/베으1212.pdf")

print(f"파일 존재: {pdf_path.exists()}")
print(f"파일 경로: {pdf_path}")

if not pdf_path.exists():
    print("파일을 찾을 수 없습니다!")
    exit(1)

found = {}
found_matches = set()

# 다양한 송장번호 패턴
patterns = [
    # 5-4-4 형식 (하이픈 변형 포함)
    r'\b(\d{5}[-–—\s]\d{4}[-–—\s]\d{4})\b',  # 60914-8682-2638 형식
    # 5-4-4 형식 (일반 하이픈만)
    r'\b(\d{5}-\d{4}-\d{4})\b',  # 60914-8675-3755 형식
    # 연속 숫자 형식
    r'\b(\d{13})\b',  # 13자리 숫자
    r'\b(\d{12})\b',  # 12자리 숫자
    r'\b(\d{10,14})\b',  # 10-14자리 숫자
]

# 방법 1: pdfplumber로 시도
text_extracted = False
try:
    with pdfplumber.open(pdf_path) as pdf:
        print(f"\n페이지 수: {len(pdf.pages)}")
        
        for i, page in enumerate(pdf.pages):
            print(f"\n=== 페이지 {i+1} (pdfplumber) ===")
            text = page.extract_text() or ""
            
            if text and len(text.strip()) > 0:
                text_extracted = True
                print(f"텍스트 추출 성공! 길이: {len(text)}")
                print(f"텍스트 샘플 (처음 200자): {text[:200]}")
                
                # 숫자 패턴 찾기
                all_numbers = re.findall(r'\d+', text)
                print(f"발견된 모든 숫자 패턴: {all_numbers[:20]}")
                
                # 송장번호 패턴 매칭
                for pattern_idx, pattern in enumerate(patterns):
                    matches = re.findall(pattern, text)
                    if matches:
                        print(f"패턴 {pattern_idx+1} 매칭: {matches}")
                    
                    for match in matches:
                        clean = re.sub(r'[-–—\s]', '', match)
                        if clean.isdigit() and len(clean) >= 10:
                            if clean not in found_matches:
                                found_matches.add(clean)
                                found[clean] = (i+1, match)
                                print(f"✓ 송장번호 발견: {match} → {clean} (페이지 {i+1})")
            else:
                print("텍스트 추출 실패 또는 빈 페이지")
except Exception as e:
    print(f"pdfplumber 실패: {str(e)}")

# 방법 2: PyMuPDF로 시도
if not text_extracted:
    try:
        import fitz
        doc = fitz.open(pdf_path)
        print(f"\n페이지 수: {len(doc)} (PyMuPDF)")
        
        for page_num in range(len(doc)):
            print(f"\n=== 페이지 {page_num+1} (PyMuPDF) ===")
            page = doc[page_num]
            text = page.get_text() or ""
            
            if text and len(text.strip()) > 0:
                print(f"텍스트 추출 성공! 길이: {len(text)}")
                print(f"텍스트 샘플 (처음 200자): {text[:200]}")
                
                # 숫자 패턴 찾기
                all_numbers = re.findall(r'\d+', text)
                print(f"발견된 모든 숫자 패턴: {all_numbers[:20]}")
                
                # 송장번호 패턴 매칭
                for pattern_idx, pattern in enumerate(patterns):
                    matches = re.findall(pattern, text)
                    if matches:
                        print(f"패턴 {pattern_idx+1} 매칭: {matches}")
                    
                    for match in matches:
                        clean = re.sub(r'[-–—\s]', '', match)
                        if clean.isdigit() and len(clean) >= 10:
                            if clean not in found_matches:
                                found_matches.add(clean)
                                found[clean] = (page_num+1, match)
                                print(f"✓ 송장번호 발견: {match} → {clean} (페이지 {page_num+1})")
            else:
                print("텍스트 추출 실패 또는 빈 페이지")
        doc.close()
    except Exception as e:
        print(f"PyMuPDF 실패: {str(e)}")
        import traceback
        traceback.print_exc()

print(f"\n총 {len(found)}개 송장번호 발견")
if found:
    print("\n발견된 송장번호 목록:")
    for clean_no, (page_num, original) in found.items():
        print(f"  {clean_no} (원본: {original}, 페이지: {page_num})")
else:
    print("\n송장번호를 찾지 못했습니다.")
    print("PDF가 이미지로만 구성되어 있어 OCR이 필요할 수 있습니다.")

