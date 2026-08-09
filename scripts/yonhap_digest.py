#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연합뉴스 인사(人事)·부고 페이지를 확인해서, 직전 영업일 08:40 ~ 당일 08:40(KST) 사이에
올라온 항목 중 "언론사" 관련 소식만 골라 Outlook으로 메일을 보내는 스크립트.
월요일 실행분은 주말을 건너뛴 만큼(금요일 08:40부터) 모아서 확인합니다.

실행 주체: GitHub Actions (평일 08:40 KST 스케줄)

주의(중요):
    이 스크립트를 작성한 환경은 네트워크 정책상 yna.co.kr 에 직접 접속해 실제 HTML 구조를
    확인할 수 없었습니다. 아래 파서는 연합뉴스 사이트의 일반적인 리스트 페이지 구조(기사 링크가
    "/view/AKR..." 형태이고, 근처에 날짜/시간 텍스트가 있다)를 가정한 "관대한(lenient)" 방식으로
    작성되었습니다. 즉, 정확한 CSS 클래스명에 의존하지 않고
        1) 기사 링크로 보이는 <a href="…/view/…"> 를 모두 찾고
        2) 그 앵커 주변 텍스트에서 날짜/시간 패턴을 정규식으로 찾아 매칭합니다.
    이렇게 하면 사이트의 사소한 마크업 변경에는 잘 버티지만, 실제 페이지와 크게 다를 경우
    항목을 하나도 못 찾을 수 있습니다. 그런 경우를 대비해:
        - 각 실행마다 원본 HTML을 debug/ 폴더에 저장해 GitHub Actions 아티팩트로 업로드합니다.
        - 파싱된 항목이 0개면 "정상 결과가 0개인지, 파싱 실패인지" 알 수 있도록 메일 본문에
          진단 정보(찾은 링크 수, 필터링된 개수 등)를 함께 적어 보냅니다.
        - 페이지 요청 자체가 실패하면(네트워크/차단 등) 예외를 삼키지 않고 알림 메일을 보냅니다.
    첫 실행 결과가 이상하면 debug/ 아티팩트를 열어보고 CSS 선택자를 조정해 주세요.
"""

from __future__ import annotations

import os
import re
import smtplib
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup

KST = ZoneInfo("Asia/Seoul")

# 기준 시각: 평일 이 시각까지 올라온 소식을 수집합니다 (workflow의 cron과 함께 맞춰야 함)
CUTOFF_HOUR = 8
CUTOFF_MINUTE = 40

PERSONNEL_URL = "https://www.yna.co.kr/people/personnel"
OBITUARY_URL = "https://www.yna.co.kr/people/obituary-notice"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

DEBUG_DIR = Path(__file__).resolve().parent.parent / "debug"
KEYWORDS_PATH = Path(__file__).resolve().parent.parent / "config" / "media_keywords.txt"

# 접미어만으로도 "언론사스러운" 이름을 넓게 잡기 위한 보조 목록.
# config/media_keywords.txt 의 명시적 키워드와 함께 사용됩니다.
SUFFIX_KEYWORDS = ["일보", "신문", "방송", "뉴스", "통신", "미디어", "타임스", "데일리", "저널", "매체"]

# 기사(항목) 링크로 인식할 URL 패턴 (연합뉴스 기사 URL은 보통 /view/AKR... 형태)
ARTICLE_LINK_RE = re.compile(r"/view/[A-Za-z0-9]+")

# 앵커 주변 텍스트에서 날짜/시간을 뽑아내기 위한 패턴들 (우선순위 순서)
DATE_PATTERNS = [
    # 2026-08-07 09:12 / 2026.08.07 09:12
    re.compile(r"(20\d{2})[.\-](\d{1,2})[.\-](\d{1,2})\.?\s*(\d{1,2}):(\d{2})"),
    # 08-07 09:12 / 08.07 09:12 (연도 없음 -> 현재 연도로 가정)
    re.compile(r"(?<!\d)(\d{1,2})[.\-](\d{1,2})\s+(\d{1,2}):(\d{2})(?!\d)"),
]


@dataclass
class Item:
    title: str
    link: str
    dt: datetime | None
    section: str  # "인사" or "부고"

    @property
    def is_media(self) -> bool:
        return is_media_related(self.title)


def load_media_keywords() -> list[str]:
    keywords = list(SUFFIX_KEYWORDS)
    if KEYWORDS_PATH.exists():
        for line in KEYWORDS_PATH.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            keywords.append(line)
    return keywords


MEDIA_KEYWORDS = load_media_keywords()


def is_media_related(text: str) -> bool:
    return any(keyword in text for keyword in MEDIA_KEYWORDS)


def fetch_html(url: str, label: str) -> str:
    resp = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    html = resp.text
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    (DEBUG_DIR / f"{label}.html").write_text(html, encoding="utf-8")
    return html


def parse_datetime_near(text: str, now: datetime) -> datetime | None:
    for pattern in DATE_PATTERNS:
        m = pattern.search(text)
        if not m:
            continue
        groups = m.groups()
        try:
            if len(groups) == 5:
                year, month, day, hour, minute = (int(g) for g in groups)
            else:
                month, day, hour, minute = (int(g) for g in groups)
                year = now.year
                # 12월 말 -> 1월 항목처럼 연도 경계를 넘는 경우 보정
                candidate = datetime(year, month, day, hour, minute, tzinfo=KST)
                if candidate > now + timedelta(days=1):
                    year -= 1
            return datetime(year, month, day, hour, minute, tzinfo=KST)
        except ValueError:
            continue
    return None


def extract_items(html: str, base_url: str, section: str, now: datetime) -> tuple[list[Item], int]:
    """반환값: (매칭된 Item 리스트, 발견된 기사 링크 총 개수[진단용])"""
    soup = BeautifulSoup(html, "lxml")
    seen_links: set[str] = set()
    items: list[Item] = []
    total_links = 0

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not ARTICLE_LINK_RE.search(href):
            continue
        link = href if href.startswith("http") else requests.compat.urljoin(base_url, href)
        if link in seen_links:
            continue

        # separator=" " 로 텍스트를 뽑아야 "제목08-07 09:12"처럼 단어가 붙어버리는 걸 방지
        raw_title = a.get_text(" ", strip=True)
        if not raw_title or len(raw_title) < 2:
            continue

        total_links += 1
        seen_links.add(link)

        # 날짜는 앵커 자신 -> 부모 -> 조부모 순으로 텍스트를 넓혀가며 탐색
        dt = None
        date_match_str = None
        node = a
        for _ in range(3):
            if node is None:
                break
            text = node.get_text(" ", strip=True)
            for pattern in DATE_PATTERNS:
                m = pattern.search(text)
                if m:
                    date_match_str = m.group(0)
                    break
            dt = parse_datetime_near(text, now)
            if dt is not None:
                break
            node = node.parent

        # 앵커 자체 텍스트에 날짜가 섞여 있으면(제목+시각이 같은 <a> 안에 있는 경우) 제거
        title = raw_title
        if date_match_str and date_match_str in title:
            title = title.replace(date_match_str, " ")
        title = re.sub(r"\s+", " ", title).strip(" -·|")

        if not title:
            continue

        items.append(Item(title=title, link=link, dt=dt, section=section))

    return items, total_links


def in_window(item: Item, start: datetime, end: datetime) -> bool:
    return item.dt is not None and start <= item.dt < end


def build_email_html(
    window_start: datetime,
    window_end: datetime,
    personnel_media: list[Item],
    obituary_media: list[Item],
    diagnostics: dict,
) -> str:
    def section_html(title: str, items: list[Item]) -> str:
        if not items:
            return f"<h3>{title}</h3><p style='color:#888;'>해당 없음</p>"
        rows = "".join(
            f"<li style='margin-bottom:8px;'>"
            f"<a href='{it.link}' style='text-decoration:none;color:#0a5cad;'>{it.title}</a>"
            f"<span style='color:#999;font-size:12px;'> ({it.dt.strftime('%m-%d %H:%M') if it.dt else '시간 미상'})</span>"
            f"</li>"
            for it in items
        )
        return f"<h3>{title} ({len(items)}건)</h3><ul style='padding-left:18px;'>{rows}</ul>"

    diag_html = (
        f"<p style='color:#aaa;font-size:11px;'>"
        f"[진단] 인사 페이지 링크 {diagnostics['personnel_total']}개 발견 / "
        f"부고 페이지 링크 {diagnostics['obituary_total']}개 발견 "
        f"(파싱이 잘못된 것 같으면 GitHub Actions 실행의 debug 아티팩트를 확인하세요)</p>"
    )

    return f"""
    <div style="font-family:'Malgun Gothic',Apple SD Gothic Neo,sans-serif;max-width:640px;">
      <h2>연합뉴스 언론사 인사·부고 다이제스트</h2>
      <p style="color:#555;">
        기간: {window_start.strftime('%Y-%m-%d %H:%M')} ~ {window_end.strftime('%Y-%m-%d %H:%M')} (KST)
      </p>
      {section_html('인사', personnel_media)}
      {section_html('부고', obituary_media)}
      <hr style="margin:24px 0;border:none;border-top:1px solid #eee;">
      <p style="font-size:12px;color:#999;">
        원본: <a href="{PERSONNEL_URL}">인사</a> / <a href="{OBITUARY_URL}">부고</a>
      </p>
      {diag_html}
    </div>
    """


def build_error_email_html(error: Exception) -> str:
    return f"""
    <div style="font-family:'Malgun Gothic',Apple SD Gothic Neo,sans-serif;">
      <h2>⚠ 연합뉴스 다이제스트 실행 실패</h2>
      <p>자동 실행 중 오류가 발생해 페이지를 확인하지 못했습니다. 아래 내용을 확인해 주세요.</p>
      <pre style="background:#f6f6f6;padding:12px;white-space:pre-wrap;">{type(error).__name__}: {error}</pre>
      <p>원본 페이지를 직접 확인해 주세요:</p>
      <ul>
        <li><a href="{PERSONNEL_URL}">인사</a></li>
        <li><a href="{OBITUARY_URL}">부고</a></li>
      </ul>
    </div>
    """


def send_mail(subject: str, html_body: str) -> None:
    smtp_user = os.environ["SMTP_USERNAME"]
    smtp_pass = os.environ["SMTP_PASSWORD"]
    # GitHub Actions는 등록되지 않은 시크릿도 "빈 문자열"로 넘겨준다(변수 자체가 없는 게 아님).
    # os.environ.get(key, default)는 키가 존재하면 빈 문자열이라도 그대로 반환해버리므로
    # 기본값이 무시된다. 그래서 `or`로 빈 문자열도 걸러내야 한다.
    mail_to = os.environ.get("MAIL_TO") or smtp_user
    smtp_server = os.environ.get("SMTP_SERVER") or "smtp-mail.outlook.com"
    smtp_port = int(os.environ.get("SMTP_PORT") or "587")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = mail_to
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    print(f"[INFO] SMTP 연결 시도: {smtp_server}:{smtp_port} (user={smtp_user}, to={mail_to})")
    with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, [mail_to], msg.as_string())


def main() -> int:
    now = datetime.now(KST)
    window_end = now.replace(hour=CUTOFF_HOUR, minute=CUTOFF_MINUTE, second=0, microsecond=0)
    if now < window_end:
        # 기준 시각(08:40) 이전에 수동 실행된 경우 등, 기준을 오늘 08:40으로 맞추기 위한 보정
        window_end -= timedelta(days=1)

    # 평일(월~금)에만 실행되는 걸 전제로 합니다. 월요일 실행분은 주말 동안 건너뛴 만큼
    # (금요일 08:40부터) 모아서 확인하고, 그 외 요일은 전날 08:40부터 확인합니다.
    lookback_days = 3 if window_end.weekday() == 0 else 1  # weekday(): 0=월요일
    window_start = window_end - timedelta(days=lookback_days)

    today_str = now.strftime("%Y-%m-%d")

    try:
        personnel_html = fetch_html(PERSONNEL_URL, "personnel")
        obituary_html = fetch_html(OBITUARY_URL, "obituary")
    except Exception as e:  # noqa: BLE001 - 네트워크 오류 등 무엇이든 알림 메일로 보고
        print(f"[ERROR] 페이지 요청 실패: {e}", file=sys.stderr)
        send_mail(f"[연합뉴스 다이제스트] 실행 실패 - {today_str}", build_error_email_html(e))
        return 1

    personnel_items, personnel_total = extract_items(personnel_html, PERSONNEL_URL, "인사", now)
    obituary_items, obituary_total = extract_items(obituary_html, OBITUARY_URL, "부고", now)

    personnel_in_window = [it for it in personnel_items if in_window(it, window_start, window_end)]
    obituary_in_window = [it for it in obituary_items if in_window(it, window_start, window_end)]

    personnel_media = [it for it in personnel_in_window if it.is_media]
    obituary_media = [it for it in obituary_in_window if it.is_media]

    print(f"[INFO] 인사: 링크 {personnel_total}개 / 기간내 {len(personnel_in_window)}개 / 언론사 {len(personnel_media)}개")
    print(f"[INFO] 부고: 링크 {obituary_total}개 / 기간내 {len(obituary_in_window)}개 / 언론사 {len(obituary_media)}개")

    subject = f"[연합뉴스 다이제스트] 언론사 인사·부고 - {today_str} (인사 {len(personnel_media)} / 부고 {len(obituary_media)})"
    body = build_email_html(
        window_start,
        window_end,
        personnel_media,
        obituary_media,
        diagnostics={"personnel_total": personnel_total, "obituary_total": obituary_total},
    )
    send_mail(subject, body)
    print("[INFO] 메일 발송 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
