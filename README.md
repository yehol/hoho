# wuzu tokki · 우주토끼 — IP 소개서

온 우주의 물질로 이루어진 토끼들.

## 결과물

- `Wuzu_Tokki_IP_소개서_20260922.pptx` — 16:9 / 10페이지 / 편집 가능
- `build-deck.js` — 편집용 소스 (PptxGenJS)
- `assets/` — 덱이 참조하는 원본 이미지

## 재생성

```bash
npm install
node build-deck.js
```

## 이미지 취급 원칙

`assets/`의 이미지는 원본을 픽셀 단위로 보존한 파일이다. 크롭·확대·변형·리터칭을
하지 않으며, 배치 시에도 원본 종횡비만 사용한다. 텍스트는 이미지 위에 올리지 않는다.
