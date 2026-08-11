#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연합뉴스 인사(人事)·부고 페이지를 확인해서, 직전 영업일 08:40 ~ 당일 08:40(KST) 사이에
올라온 항목 중 "언론사" 관련 소식만 골라 Outlook으로 메일을 보내는 스크립트.
월요일 실행분은 주말을 건너뛴 만큼(금요일 08:40부터) 모아서 확인합니다.

실행 주체: GitHub Actions (평일 08:40 KST 스케줄)

파서 구조:
    연합뉴스 인사/부고 목록의 각 항목은 <li data-cid="AKR..."> 로 감싸여 있고(사이드바 광고나
    추천기사는 이 속성이 없어 구분됨), 그 안에 제목(.title01), 본문 요약(p.lead), 게시 시각
    (.txt-time)이 이미 다 들어있다. 이 구조를 실제 페이지에서 확인해 반영했다 (extract_items /
    _extract_from_list_items 참고). 혹시 사이트 마크업이 바뀌어 이 구조를 하나도 못 찾으면,
    예전의 "/view/ 링크 전체 훑기 + 제목 패턴 필터" 방식으로 자동 폴백한다 (_extract_fallback).

    이 페이지에는 언론계 인사·부고 외에 스포츠 선수, 연예인 등 일반 유명인의 부고도 같이
    실린다(실사용 데이터로 확인함). 그래서 목록에서 뽑은 제목에 언론사 키워드
    (config/media_keywords.txt)가 있는 항목만 최종적으로 메일에 포함한다 (Item.is_media).
    본문까지 검사하면 기사가 인용한 다른 언론사 이름 때문에 오탐이 생겨서 제목만 본다.
    실제로 보내는 메일에 언론사가 아닌데 잘못 포함되거나, 언론사인데 빠진 항목이 보이면
    config/media_keywords.txt 에 키워드를 추가/조정하면 된다 (코드 수정 불필요).
        - 각 실행마다 원본 HTML을 debug/ 폴더에 저장해 GitHub Actions 아티팩트로 업로드합니다.
        - 페이지 요청 자체가 실패하면(네트워크/차단 등) 예외를 삼키지 않고 알림 메일을 보냅니다.
    폴백이 동작했다는 로그("[WARN] li[data-cid] 구조를 찾지 못해...")가 보이면 마크업이 바뀐
    것이니 debug/ 아티팩트를 열어 구조를 다시 확인해 주세요.
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
SUFFIX_KEYWORDS = ["일보", "신문", "방송", "뉴스", "통신", "미디어", "타임스", "데일리", "저널", "매체", "매일"]

# 기사(항목) 링크로 인식할 URL 패턴 (연합뉴스 기사 URL은 보통 /view/AKR... 형태)
ARTICLE_LINK_RE = re.compile(r"/view/[A-Za-z0-9]+")

# 앵커 주변 텍스트에서 날짜/시간을 뽑아내기 위한 패턴들 (우선순위 순서)
DATE_PATTERNS = [
    # 2026-08-07 09:12 / 2026.08.07 09:12
    re.compile(r"(20\d{2})[.\-](\d{1,2})[.\-](\d{1,2})\.?\s*(\d{1,2}):(\d{2})"),
    # 08-07 09:12 / 08.07 09:12 (연도 없음 -> 현재 연도로 가정)
    re.compile(r"(?<!\d)(\d{1,2})[.\-](\d{1,2})\s+(\d{1,2}):(\d{2})(?!\d)"),
]


WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


def format_korean_date(dt: datetime) -> str:
    return f"{dt.year}년 {dt.month}월 {dt.day}일 ({WEEKDAY_KO[dt.weekday()]})"


# "(전남광주=연합뉴스)" 같은 송고처 표기 + 뒤따르는 "홍길동 기자 = " 바이라인을 제거하기
# 위한 패턴. 부고/인사 정형 공지는 이 표기가 본문 "끝"에 붙지만, 일반 뉴스 기사는 본문
# "맨 앞"에 붙는다 (예: "(서울=연합뉴스) 이충원 기자 = 일본 럭비..."). 앞/뒤 어디에 있든
# 지우지 않으면 "뉴스"라는 접미어 키워드가 "연합뉴스"라는 단어 자체와 매칭돼버려서,
# 언론사와 무관한 일반 기사(예: 스포츠 선수 부고)까지 "언론사 관련"으로 잘못 판정된다.
YONHAP_BYLINE_RE = re.compile(r"\([^()]{0,20}=\s*연합뉴스\)\s*(?:[가-힣]{2,4}\s*기자\s*=\s*)?")


def strip_yonhap_byline(text: str) -> str:
    return re.sub(r"\s+", " ", YONHAP_BYLINE_RE.sub(" ", text)).strip()


@dataclass
class Item:
    title: str
    link: str
    dt: datetime | None
    section: str  # "인사" or "부고"
    company: str = ""
    headline: str = ""  # 기사 페이지에서 다시 뽑은 제목 (목록 제목보다 정확함)
    body: str = ""       # 기사 본문 (▲ 상세 내용)

    @property
    def is_media(self) -> bool:
        # 본문까지 검사하면 "~라고 OO 매체가 보도했다"처럼 기사가 인용한 다른 언론사 이름
        # 때문에 오탐이 생긴다(실사용 데이터로 확인함). 회사명은 항상 제목(또는 제목 안
        # 괄호)에 들어있으므로 제목만 검사한다.
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


def matched_media_keyword(text: str) -> str | None:
    """is_media_related() 판정에 실제로 어떤 키워드가 매치됐는지(디버깅용)."""
    for keyword in MEDIA_KEYWORDS:
        if keyword in text:
            return keyword
    return None


def is_media_related(text: str) -> bool:
    return matched_media_keyword(text) is not None


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
    """반환값: (매칭된 Item 리스트, 발견된 목록 항목 총 개수[진단용])

    실제 인사/부고 목록의 각 항목은 <li data-cid="AKR..."> 로 감싸여 있고(사이드바 광고/추천
    기사는 이 속성이 없음), 그 안에 제목(.title01), 본문 요약(p.lead), 시간(.txt-time)이
    이미 다 들어있다. 이 구조를 우선적으로 쓰고, 혹시 마크업이 바뀌어 하나도 못 찾으면
    예전의 "페이지 전체에서 /view/ 링크 훑기" 방식으로 폴백한다.
    """
    soup = BeautifulSoup(html, "lxml")
    list_items = soup.find_all("li", attrs={"data-cid": True})

    if list_items:
        return _extract_from_list_items(list_items, base_url, section, now)

    print(f"[WARN] li[data-cid] 구조를 찾지 못해 폴백 파서를 사용합니다 ({section})")
    return _extract_fallback(html, base_url, section, now)


def _extract_from_list_items(list_items, base_url: str, section: str, now: datetime) -> tuple[list[Item], int]:
    items: list[Item] = []
    seen_links: set[str] = set()

    for li in list_items:
        a = li.select_one("a.tit-news") or li.find("a", href=True)
        if not a or not a.get("href"):
            continue
        href = a["href"]
        link = href if href.startswith("http") else requests.compat.urljoin(base_url, href)
        if link in seen_links:
            continue
        seen_links.add(link)

        title_el = li.select_one(".title01") or a
        title = re.sub(r"\s+", " ", title_el.get_text(" ", strip=True)).strip()
        if not title:
            continue

        lead_el = li.select_one("p.lead")
        body = re.sub(r"\s+", " ", lead_el.get_text(" ", strip=True)).strip() if lead_el else ""
        body = strip_yonhap_byline(body)

        time_el = li.select_one(".txt-time")
        dt = parse_datetime_near(time_el.get_text(" ", strip=True), now) if time_el else None

        items.append(Item(title=title, link=link, dt=dt, section=section, headline=title, body=body))

    return items, len(list_items)


def _extract_fallback(html: str, base_url: str, section: str, now: datetime) -> tuple[list[Item], int]:
    """li[data-cid] 구조가 안 보일 때 쓰는 예전 방식 (관대하지만 사이드바가 섞일 수 있음)."""
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

        raw_title = a.get_text(" ", strip=True)
        if not raw_title or len(raw_title) < 2:
            continue

        total_links += 1
        seen_links.add(link)

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

        title = raw_title
        if date_match_str and date_match_str in title:
            title = title.replace(date_match_str, " ")
        title = re.sub(r"\s+", " ", title).strip(" -·|")

        if not title:
            continue

        # data-cid 구조가 없을 때만 쓰는 안전장치: 최소한 제목 패턴으로라도 걸러본다.
        if not looks_like_section_item(title, section):
            continue

        items.append(Item(title=title, link=link, dt=dt, section=section))

    return items, total_links


def looks_like_section_item(title: str, section: str) -> bool:
    t = title.strip()
    if section == "인사":
        return t.endswith("인사")
    if section == "부고":
        return t.startswith("[부고]")
    return True


def in_window(item: Item, start: datetime, end: datetime) -> bool:
    return item.dt is not None and start <= item.dt < end


def fetch_article_detail(url: str) -> tuple[str, str]:
    """개별 기사 페이지에서 (제목, 본문)을 뽑는다. 실패/미확인 시 ("", "")."""
    try:
        resp = requests.get(url, headers=REQUEST_HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
    except Exception as e:  # noqa: BLE001 - 기사 하나 실패로 전체를 죽이지 않음
        print(f"[WARN] 기사 본문 요청 실패: {url} ({e})")
        return "", ""

    headline = ""
    h1 = soup.find("h1")
    if h1:
        headline = h1.get_text(" ", strip=True)
    if not headline:
        og_title = soup.find("meta", property="og:title")
        if og_title and og_title.get("content"):
            headline = og_title["content"].strip()

    body = ""
    article = soup.find("article")
    if article:
        paragraphs = [p.get_text(" ", strip=True) for p in article.find_all("p")]
        body = " ".join(p for p in paragraphs if p)
    if not body:
        og_desc = soup.find("meta", property="og:description")
        if og_desc and og_desc.get("content"):
            body = og_desc["content"].strip()
    if not body:
        # 최후 수단: "▲"가 포함된 것 중 적당히 긴(30~2000자) 블록을 본문으로 추정
        best = ""
        for c in soup.find_all(["div", "section"]):
            text = c.get_text(" ", strip=True)
            if 30 < len(text) < 2000 and text.count("▲") >= 1 and len(text) > len(best):
                best = text
        body = best

    headline = re.sub(r"\s+", " ", headline).strip()
    body = re.sub(r"\s+", " ", body).strip()
    body = strip_yonhap_byline(body)
    return headline, body


# 부고 괄호 안 "회사명 직함"에서 직함으로 보이는 꼬리 단어들. 이 단어들을 뒤에서부터
# 떼어내고 남는 부분을 회사명으로 쓴다 (완벽하지 않은 휴리스틱 - 안 맞으면 이 목록에 추가).
JOB_TITLE_WORDS = [
    "기자", "대표", "전문위원", "부국장", "국장대행", "편집국장", "국장", "팀장",
    "부장", "실장", "본부장", "이사", "센터장", "에디터", "위원", "사장", "회장",
    "부사장", "상무", "전무", "주필", "논설위원", "부실장", "정치부장", "산업부장",
    "경제부장", "사회부장", "국제부장", "문화부장", "체육부장", "보도국", "편집제작",
    "광고부장", "디지털뉴스부", "취재부장",
]


# "(향년 71세, 전 MBC 편성국장)"처럼 나이 정보가 같이 들어간 괄호에서 나이 부분을 제거하기 위한 패턴
AGE_PATTERN = re.compile(r"향년\s*(?:만\s*)?\d+세\s*,?\s*")


def strip_trailing_job_title(paren_text: str) -> str:
    paren_text = AGE_PATTERN.sub("", paren_text).strip(" ,")
    tokens = paren_text.split()
    while tokens and any(w in tokens[-1] for w in JOB_TITLE_WORDS):
        tokens.pop()
    return " ".join(tokens).strip(" ,")


def extract_company_label(title: str, body: str) -> str:
    """제목/본문에서 언론사명을 뽑는다 (표의 왼쪽 칸에 쓸 라벨)."""
    combined = f"{title} {body}"

    # 부고 형식: "이영섭(뉴스1 대표)씨 모친상" - 괄호 안에서 나이/직함을 떼어내고 남는 걸 회사명으로 사용
    parens = re.findall(r"\(([^)]*)\)", title)
    if parens:
        stripped = strip_trailing_job_title(parens[0])
        if stripped:
            return stripped

    # 인사 형식: "IT조선 인사" 처럼 " 인사"로 끝나면 그 앞부분을 회사명으로 사용
    m = re.match(r"^(.+?)\s*인사\b", title)
    if m and m.group(1).strip():
        return m.group(1).strip()

    # 최후 수단: 매칭된 키워드 중 가장 긴 것을 라벨로 사용
    for kw in sorted(MEDIA_KEYWORDS, key=len, reverse=True):
        if kw in combined:
            return kw

    return title


# 부고 본문(▲ 이름(소속)씨 관계, ... = 날짜정보)에서 "이름(소속)씨 관계" 부분을 뽑아
# 표준 "[부고] 이름(소속)씨 관계" 제목을 새로 만들기 위한 패턴. 정형 부고 공지(예: 이두헌)는
# 본문이 이 "▲" 불릿 형식이라 이걸로 잡힌다.
OBIT_LEAD_PATTERN = re.compile(r"▲\s*([가-힣]{2,4}(?:\([^)]{0,40}\))?)\s*씨\s*([가-힣]{1,6}상|별세)")

# 유명 인사는 정형 공지 대신 일반 기사 형식으로 올라온다(예: "'사랑이 뭐길래' 수출로 한류
# 개척…박재복 전 MBC 국장 별세"). 이런 기사는 본문에 "▲" 불릿이 없어 OBIT_LEAD_PATTERN으로
# 못 잡으므로, 제목 자체의 맨 끝에서 "이름 (전/현) 소속 직함 별세/OO상" 패턴을 찾는다. 앞에
# 어떤 리드 문구가 붙어도(…로 구분) 문자열 끝에서부터 매칭하므로 영향받지 않는다.
HEADLINE_OBIT_PATTERN = re.compile(
    r"([가-힣]{2,4})\s+((?:전|현)\s?[가-힣A-Za-z0-9]+(?:\s[가-힣A-Za-z0-9·]+){0,4})\s*(별세|[가-힣]{1,6}상)\s*$"
)


def synthesize_obituary_title(headline: str, body: str) -> str | None:
    headline_m = HEADLINE_OBIT_PATTERN.search(headline.strip())
    body_m = OBIT_LEAD_PATTERN.search(body)

    # 헤드라인(예: "박재복 전 MBC 국장 별세")과 본문 불릿(예: "▲ 박재복씨 별세, ...")이
    # 같은 이름을 가리키면 "본인 부고"(=본인 본인상)로 본다. 이 경우 본문에는 소속 정보가
    # 없는 게 보통이라(장례식장/발인 정보만 있음), 소속은 헤드라인에서 가져와 제목에 붙이고
    # "별세"는 다른 항목들의 "이름(소속)씨 OO상" 패턴과 통일되도록 "본인상"으로 표시한다.
    # (본문 텍스트 자체는 건드리지 않는다 - 이미 정리된 형식 그대로 둔다.)
    if headline_m and body_m:
        h_name, affiliation, _h_relation = headline_m.groups()
        b_name, b_relation = body_m.groups()
        if h_name == b_name:
            relation = "본인상" if b_relation == "별세" else b_relation
            return f"[부고] {h_name}({affiliation.strip()})씨 {relation}"

    if body_m:
        name_part, relation = body_m.groups()
        return f"[부고] {name_part}씨 {relation}"

    if headline_m:
        name, affiliation, relation = headline_m.groups()
        return f"[부고] {name}({affiliation.strip()})씨 {relation}"

    return None


def enrich_items(items: list[Item]) -> None:
    """회사명을 채우고, 목록 페이지에 본문(lead)이 없었던 항목만 기사 페이지를 열어 보완한다."""
    for it in items:
        if not it.headline:
            it.headline = it.title
        if not it.body:
            headline, body = fetch_article_detail(it.link)
            it.headline = headline or it.headline
            it.body = body

        # 헤드라인형 기사는 표준 "[부고] 이름(소속)씨 관계" 형식으로 제목을 바꿔서 보여준다.
        if it.section == "부고" and not it.headline.strip().startswith("[부고]"):
            synthesized = synthesize_obituary_title(it.headline, it.body)
            if synthesized:
                it.headline = synthesized

        it.company = extract_company_label(it.headline, it.body)


def build_email_html(window_end: datetime, personnel_media: list[Item], obituary_media: list[Item]) -> str:
    date_str = format_korean_date(window_end)

    def rows_html(items: list[Item]) -> str:
        if not items:
            return (
                "<tr><td colspan='2' style='border:1px solid #ddd;padding:10px;color:#888;'>"
                "해당 없음</td></tr>"
            )
        out = []
        for it in items:
            title_line = (
                f"<div style='font-weight:bold;'>{it.headline}</div>"
                if it.section == "부고" and it.headline
                else ""
            )
            body_text = it.body or "(본문 확인 필요 - 원문 링크 참고)"
            out.append(
                "<tr>"
                f"<td style='border:1px solid #ddd;padding:8px;vertical-align:top;white-space:nowrap;'>{it.company}</td>"
                "<td style='border:1px solid #ddd;padding:8px;vertical-align:top;'>"
                f"{title_line}<div>{body_text}</div>"
                f"<div style='margin-top:4px;'><a href='{it.link}' style='font-size:11px;color:#999;'>원문</a></div>"
                "</td>"
                "</tr>"
            )
        return "".join(out)

    return f"""
    <div style="font-family:'Malgun Gothic',Apple SD Gothic Neo,sans-serif;max-width:700px;">
      <p>안녕하세요,</p>
      <p>{date_str} 인사 / 부고 전달 드립니다.</p>
      <p>감사합니다.</p>
      <hr style="margin:20px 0;border:none;border-top:1px solid #ccc;">
      <table style="border-collapse:collapse;width:100%;font-size:14px;">
        <tr><th colspan="2" style="background:#fff2ac;text-align:left;padding:8px;border:1px solid #ddd;">인사</th></tr>
        {rows_html(personnel_media)}
        <tr><th colspan="2" style="background:#fff2ac;text-align:left;padding:8px;border:1px solid #ddd;">부고</th></tr>
        {rows_html(obituary_media)}
      </table>
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
    target_date_str = os.environ.get("TARGET_DATE")
    if target_date_str:
        # 테스트/재현용: 특정 날짜의 08:40 기준으로 실행한 것처럼 시뮬레이션합니다.
        base_date = datetime.strptime(target_date_str, "%Y-%m-%d").replace(tzinfo=KST)
        now = base_date.replace(hour=CUTOFF_HOUR, minute=CUTOFF_MINUTE) + timedelta(minutes=1)
        print(f"[INFO] TARGET_DATE={target_date_str} 기준으로 테스트 실행 (실제 현재 시각 아님)")
    else:
        now = datetime.now(KST)
    # window_end는 항상 "실행된 날짜의 08:40"으로 고정한다 (실제 실행 시각과 무관).
    # GitHub Actions의 schedule 트리거는 혼잡 시간대(특히 UTC 자정 = 한국시간 오전 9시
    # 부근)에 몰리면 수 분~수십 분씩 지연될 수 있는데(실사용 데이터로 확인함: 예약
    # 08:40 대비 실제 실행 09:02~09:04), window_end를 실행 시각 기준으로 다시 계산하면
    # "08:40 이전 실행 -> 어제 기준으로 취급"하는 예전 로직과 충돌해 지연이 적은 날에는
    # 엉뚱하게 전날 기준으로 발송될 위험이 있었다. 실행 시각이 언제든(빨라도 늦어도)
    # 08:40 컷오프는 항상 "그 날짜"로 고정해야 안전하다. 테스트/재현용(TARGET_DATE)도
    # 이미 그 날짜의 08:40+1분으로 now를 맞춰 만들기 때문에 결과는 동일하다.
    window_end = now.replace(hour=CUTOFF_HOUR, minute=CUTOFF_MINUTE, second=0, microsecond=0)

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

    # 연합뉴스 /people/personnel, /people/obituary-notice 페이지에는 언론계 인사·부고 외에
    # 스포츠 선수, 연예인 등 일반 유명인 부고도 같이 실린다(실사용 데이터로 확인함). 그래서
    # 제목에 언론사 키워드(config/media_keywords.txt)가 있는 항목만 최종 채택한다.
    personnel_media = [it for it in personnel_in_window if it.is_media]
    obituary_media = [it for it in obituary_in_window if it.is_media]

    print(f"[INFO] 인사: 링크 {personnel_total}개 / 기간내 {len(personnel_in_window)}개 / 언론사 {len(personnel_media)}개")
    print(f"[INFO] 부고: 링크 {obituary_total}개 / 기간내 {len(obituary_in_window)}개 / 언론사 {len(obituary_media)}개")

    # 디버그: 기간 내로 잡힌 항목을 전부 로그에 출력 (엉뚱한 기사가 섞이는지 확인용)
    # is_media 판정은 제목만 보므로(본문 인용 오탐 방지), 디버그도 제목만으로 매칭 키워드를 표시한다.
    for it in personnel_in_window:
        kw = matched_media_keyword(it.title)
        print(f"[DEBUG][인사] media_kw={kw!r} | {it.dt} | {it.title} | {it.link}")
    for it in obituary_in_window:
        kw = matched_media_keyword(it.title)
        print(f"[DEBUG][부고] media_kw={kw!r} | {it.dt} | {it.title} | {it.link}")

    # 기간 내 전체 항목의 개별 기사 페이지를 열어 본문/회사명을 채운다
    enrich_items(personnel_media)
    enrich_items(obituary_media)

    # 디버그: 헤드라인형 부고 제목이 표준 형식으로 잘 변환됐는지(또는 왜 안 됐는지) 확인용.
    # 위쪽 [DEBUG][부고] 로그는 원본 제목(it.title)을 찍지만, 이 로그는 enrich 이후
    # 실제 메일에 나가는 최종 제목(it.headline)을 찍는다.
    for it in obituary_media:
        converted = "O" if it.headline.strip().startswith("[부고]") else "X"
        print(f"[DEBUG][부고 enrich] 표준형식변환={converted} | headline={it.headline!r} | company={it.company!r}")
        if converted == "X":
            print(f"[DEBUG][부고 enrich]   body={it.body!r}")

    date_str = format_korean_date(window_end)
    subject = f"{date_str} 인사 / 부고"
    body = build_email_html(window_end, personnel_media, obituary_media)
    send_mail(subject, body)
    print("[INFO] 메일 발송 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
