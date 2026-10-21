# 상품권 자금대사 웹사이트

통장(발행처) 집계 엑셀과 상품권 일별 잔액집계 엑셀을 통장일 기준으로 대사하고, 대조 엑셀을 내려받는 정적 웹사이트입니다.

- 서버 프로그램이 필요 없습니다. 이 폴더를 그대로 웹에 올리면 됩니다.
- 업로드한 파일은 사용자의 브라우저 안에서만 처리되고, 어디에도 전송·저장되지 않습니다.
- 외부 CDN을 쓰지 않습니다(엑셀 라이브러리는 `vendor/`에 포함). 폰트만 Google Fonts를 쓰고, 막혀 있으면 기본 글꼴로 표시됩니다.

## 파일 구성

```
index.html              화면
js/engine.js            대사 로직, 휴일 목록
js/app.js               화면 동작
vendor/exceljs.min.js   엑셀 읽기·쓰기 라이브러리 (ExcelJS 4.4.0, MIT)
test/, package.json     엔진 테스트 (배포에는 필요 없음)
```

## 테스트

샘플 엑셀을 저장소 밖 `../samples/bank.xlsx`, `../samples/voucher.xlsx`에 두고 실행합니다.

```bash
npm install
npm test
```

## 올리는 방법 (택1)

### 1) Netlify Drop — 가장 간단 (계정만 있으면 1분)
1. https://app.netlify.com/drop 접속
2. 이 폴더 전체를 화면에 끌어다 놓기
3. 발급된 주소(예: `https://xxxx.netlify.app`)를 공유

### 2) GitHub Pages
1. GitHub에 새 저장소를 만들고 이 폴더 내용을 올림
2. Settings → Pages → Branch: `main` / root 선택 후 저장
3. `https://<계정>.github.io/<저장소>/` 주소로 접속

### 3) 사내 웹서버
Nginx, Apache, IIS 등 정적 파일을 서비스할 수 있는 곳이면 폴더를 그대로 복사하면 됩니다.

## 접근 제한이 필요하면
누구나 주소만 알면 열 수 있습니다. 데이터는 전송되지 않지만 내부 담당자만 쓰게 하려면
- Netlify: Site settings → Access control → Password protection (유료 플랜)
- Cloudflare Pages + Cloudflare Access (회사 이메일로 로그인 제한, 50명까지 무료)
- 사내망 서버에 올리기

## 휴일 수정
`js/engine.js`에서 `const HOLIDAYS = {` 부분에 `'YYYY-MM-DD':'이름'` 형식으로 추가합니다.
화면에서 추가한 휴일은 그 사용자 브라우저에만 저장됩니다.

## 대사 기준
| 항목 | 통장 | 상품권 파일 | 거래일 |
|---|---|---|---|
| 개인구매금 | 입금·개인 구매금 | 개인구매(구매일) 고객구매금 | D-1 |
| 법인구매금 | 입금·법인 구매금 | 법인구매(입금완료일) 법인구매금 | D-1 |
| 개인환불금 | 출금·개인 환불금 | 개인환불(신청일) 고객금액 | D-1 |
| 법인환불금 | 출금·법인 환불금 | 법인환불(관리자승인일) 고객금액 | D-1 |
| QR추심 | 입금·QR추심/기타 | QR 결제 추심 | D |
| QR정산 | 출금·QR정산/기타 | QR 결제 정산 | D |
| PG정산 | 출금·PG정산 | 온라인 결제 정산 | D |

주말·휴일 금액은 다음 영업일 통장 금액과 합산해 비교합니다.
