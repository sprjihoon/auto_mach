# 자동출고 프로그램 v1.0

Windows용 바코드 스캔 기반 자동출고 시스템

## 📦 주요 기능

1. **바코드 스캔 자동 감지**: HID 스캐너로 바코드 스캔 시 자동 인식
2. **송장번호 역매칭**: 바코드로 해당 tracking_no 자동 검색
3. **EzAuto 자동입력**: 송장번호와 바코드를 EzAuto에 자동 입력
4. **실시간 수량 추적**: qty/scanned_qty 실시간 관리
5. **PDF 자동출력**: 구성 완료 시 송장 라벨 자동 인쇄
6. **재처리 방지**: 완료된 송장 재스캔 방지 (used=1)

## 🔧 설치

### 필수 요구사항
- Windows 10 이상
- Python 3.9 이상

### 의존성 설치

```bash
pip install -r requirements.txt
```

## 🚀 실행

### 개발 환경
```bash
python main.py
```

### EXE 빌드
```bash
# 방법 1: spec 파일 사용
pyinstaller auto_mach.spec

# 방법 2: 직접 빌드
pyinstaller -F -w main.py --name 자동출고프로그램
```

## 📂 엑셀 파일 형식

프로그램이 사용하는 필수 컬럼:

| 컬럼명 | 설명 | 예시 |
|--------|------|------|
| tracking_no | 송장번호 | 6091486739755 |
| barcode | 상품 바코드 | 8801234567890 |
| product_name | 상품명 | 건강식품 세트 |
| option_name | 옵션명 | 대용량 |
| qty | 필요 수량 | 3 |
| scanned_qty | 스캔 수량 (자동생성) | 0 |
| used | 처리 완료 여부 (자동생성) | 0 |

## 📁 폴더 구조

```
/auto_mach/
├── main.py              # 메인 진입점
├── ui_main.py           # PySide6 UI
├── excel_loader.py      # 엑셀 관리
├── scanner_listener.py  # 스캐너 입력
├── ezauto_input.py      # EzAuto 입력
├── pdf_printer.py       # PDF 출력
├── order_processor.py   # 주문 처리
├── models.py            # 데이터 모델
├── utils.py             # 유틸리티
├── requirements.txt     # 의존성
├── auto_mach.spec       # PyInstaller spec
├── labels/              # PDF 라벨 폴더
│   ├── 6091486739755.pdf
│   └── ...
└── 주문데이터.xlsx       # 엑셀 데이터
```

## 🔍 스캔 우선순위

바코드 스캔 시 다음 우선순위로 송장 선택:

1. **qty 오름차순**: 단품(qty=1) → 소형조합 → 대형조합
2. **tracking_no 오름차순**: 동일 qty면 송장번호 작은 것 우선

## ⚙️ 사용 방법

1. **엑셀 로드**: "찾아보기" → 엑셀 파일 선택 → "불러오기"
2. **PDF 폴더 설정**: labels 폴더 경로 입력 (기본: labels)
3. **스캐너 시작**: "스캐너 시작" 버튼 클릭
4. **바코드 스캔**: HID 스캐너로 바코드 스캔
5. **자동 처리**: EzAuto 입력 → 수량 추적 → PDF 출력

## ⚠️ 주의사항

- EzAuto 프로그램이 활성화된 상태에서 사용
- 마우스를 화면 모서리로 이동하면 안전모드 발동 (pyautogui)
- PDF 파일명은 `{tracking_no}.pdf` 형식이어야 함
- 기본 프린터가 설정되어 있어야 PDF 자동 인쇄 가능

## 📝 라이선스

MIT License

