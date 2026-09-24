# 중국 SNS 완전정복

중국 SNS의 종류·특징·문화를 소개하는 5페이지 PPT 프레젠테이션입니다.

## 파일 구성

- `중국_SNS_완전정복.pptx` — 최종 발표 자료
- `build/build_deck.py` — `python-pptx`로 덱을 생성하는 스크립트
- `build/assets/icons/` — 슬라이드에 쓰인 각 SNS 앱 아이콘 원본 이미지
- `project/` — 초기 디자인 시안(HTML 목업, PowerShell 생성 스크립트)

## 덱 재생성

```bash
pip install python-pptx Pillow
cd build
python build_deck.py
```

`build/china_sns_deck.pptx`로 결과물이 생성됩니다.

## 슬라이드 구성

1. 표지
2. 개관 — 서구 SNS와 중국 대체 플랫폼 비교
3. 주요 플랫폼 — 위챗·웨이보·더우인·샤오홍슈·빌리빌리
4. 중국 SNS 문화의 특징
5. 감사합니다
