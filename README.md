# 연합뉴스 언론사 인사·부고 다이제스트 자동화

평일 오전 8시 40분(KST)에 연합뉴스 [인사](https://www.yna.co.kr/people/personnel) /
[부고](https://www.yna.co.kr/people/obituary-notice) 페이지를 확인해서, 직전 영업일 08:40 ~
당일 08:40 사이에 올라온 항목 중 **언론사(신문·방송·통신사 등) 관련 소식만** 골라 Outlook
메일로 보내는 GitHub Actions 자동화입니다. 월요일 실행분은 주말 동안 건너뛴 만큼(금요일
08:40부터) 모아서 확인합니다.

## 동작 방식

- `.github/workflows/yonhap-media-digest.yml` : 평일(월~금) 08:40 KST에 실행되는 GitHub
  Actions 스케줄(`cron: "40 23 * * 0-4"`, UTC 기준). `workflow_dispatch`로 수동 실행도
  가능합니다.
- `scripts/yonhap_digest.py` : 두 페이지를 요청 → 항목/날짜 파싱 → 기간 윈도우로 필터링 →
  언론사 키워드 매칭 → SMTP로 메일 발송. 기준 시각은 파일 상단의 `CUTOFF_HOUR` /
  `CUTOFF_MINUTE` 상수로 관리합니다 (워크플로우의 cron과 같이 맞춰야 합니다).
- `config/media_keywords.txt` : "언론사"로 판별할 키워드 목록. 코드 수정 없이 이 파일만 편집하면
  필터링 기준을 바꿀 수 있습니다.

## 설정이 필요한 GitHub Secrets

저장소 **Settings → Secrets and variables → Actions** 에서 아래 값을 등록해주세요.

| Secret | 필수 | 설명 |
|---|---|---|
| `SMTP_USERNAME` | ✅ | 메일을 보낼 Outlook 계정 (예: `you@outlook.com`). 발신자이자 기본 수신자로 사용됩니다. |
| `SMTP_PASSWORD` | ✅ | 위 계정의 **앱 비밀번호**(일반 로그인 비밀번호 아님). 아래 "앱 비밀번호 발급" 참고. |
| `MAIL_TO` | ❌ | 받는 사람 주소를 발신 계정과 다르게 하고 싶을 때만. 생략 시 `SMTP_USERNAME`으로 발송. |
| `SMTP_SERVER` | ❌ | 기본값 `smtp-mail.outlook.com` (개인 outlook.com/hotmail.com 계정용). 회사/학교 Microsoft 365 계정이면 `smtp.office365.com`으로 바꿔야 할 수 있습니다. |

### Outlook 앱 비밀번호 발급 방법 (개인 계정)

1. https://account.live.com/proofs/AppPassword 접속 (또는 Microsoft 계정 → 보안 → 고급 보안 옵션)
2. 2단계 인증이 켜져 있어야 앱 비밀번호를 만들 수 있습니다. 꺼져 있다면 먼저 활성화하세요.
3. "새 앱 암호 만들기" → 생성된 문자열을 `SMTP_PASSWORD` 시크릿에 그대로 입력.

> 회사/학교 Microsoft 365(테넌트) 계정은 보안 정책상 SMTP AUTH(기본 인증)가 관리자에 의해
> 막혀 있는 경우가 많습니다. 로그인/발송이 실패하면 관리자에게 "해당 계정의 SMTP AUTH 허용"을
> 요청하거나, Microsoft Graph API 기반 발송 방식으로 전환이 필요합니다.

## 언론사 필터 커스터마이징

`config/media_keywords.txt` 에 키워드를 한 줄씩 추가/삭제하면 됩니다. 제목에 해당 키워드가
포함된 인사/부고 항목만 메일에 포함됩니다. 회사명 접미어(`~일보`, `~신문`, `~방송`, `~뉴스`,
`~통신` 등)는 스크립트에 기본 내장되어 있어 별도 등록 없이도 어느 정도 잡힙니다.

## ⚠ 알려진 한계 (중요)

이 스크립트를 만든 환경은 네트워크 정책상 `yna.co.kr`에 실제로 접속해 페이지 구조를 확인할
수 없었습니다. 그래서 파서는 정확한 CSS 클래스명 대신, "기사 링크는 `/view/...` 형태이고 그
주변에 날짜/시간 텍스트가 있다"는 일반적인 가정으로 관대하게 항목을 추출합니다.

- 처음 몇 번은 **결과를 꼭 확인**해 주세요 (메일 하단에 "[진단] 인사 페이지 링크 N개 발견 /
  부고 페이지 링크 N개 발견" 문구가 함께 옵니다. 이 숫자가 0이면 파싱이 실패한 것입니다).
- 매 실행마다 원본 HTML이 GitHub Actions 실행 결과의 `debug-html-*` 아티팩트로 업로드됩니다.
  파싱이 이상하면 이 아티팩트를 열어서 실제 구조를 보고 `scripts/yonhap_digest.py`의
  `extract_items` / `DATE_PATTERNS`를 조정하면 됩니다.
- 페이지 요청 자체가 실패한 경우(네트워크 오류, 차단 등)에도 조용히 실패하지 않고 오류 내용을
  담은 알림 메일을 보내도록 되어 있습니다.

## 수동 테스트

Actions 탭 → "Yonhap Media Personnel & Obituary Digest" → **Run workflow** 로 즉시 실행해
결과를 확인할 수 있습니다. Secrets만 등록하면 스케줄을 기다리지 않고 바로 테스트 가능합니다.

## 메일 양식 변경

현재는 기본 HTML 템플릿(`scripts/yonhap_digest.py`의 `build_email_html` 함수)으로 발송됩니다.
평소 쓰시는 메일 양식을 알려주시면 그 형식에 맞게 다시 다듬어 드릴 수 있습니다.
