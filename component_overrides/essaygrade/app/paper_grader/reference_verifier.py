from __future__ import annotations

from dataclasses import asdict, dataclass
from html import unescape
import re
from typing import Iterable

from bs4 import BeautifulSoup
import requests


SEARCH_TIMEOUT_SECONDS = 20
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0 Safari/537.36"
TITLE_STOPWORDS = {
    "研究",
    "分析",
    "应用",
    "设计",
    "实现",
    "系统",
    "问题",
    "对策",
    "思考",
    "探讨",
    "浅析",
    "论文",
}
TITLE_FRAGMENT_SPLIT_RE = re.compile(r"基于|关于|面向|针对|围绕|依托|结合|对于|在|与|和|及其|及|的|之")
GENERIC_FRAGMENT_SUFFIXES = (
    "研究",
    "分析",
    "应用",
    "设计",
    "实现",
    "探析",
    "思考",
    "对策",
    "实践",
    "综述",
    "初探",
    "策略",
    "开发",
    "建模",
)


@dataclass
class ReferenceEntry:
    index: int
    raw_text: str
    title: str
    authors: str
    year: str | None

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ReferenceCheck:
    index: int
    raw_text: str
    title: str
    cited_in_body: bool
    status: str
    confidence: float
    source: str | None
    matched_title: str | None
    matched_url: str | None
    notes: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ReferenceAuditResult:
    entries: list[ReferenceEntry]
    checks: list[ReferenceCheck]
    citation_numbers: list[int]
    dangling_citations: list[int]
    uncited_reference_indexes: list[int]
    verified_count: int
    possible_count: int
    not_found_count: int
    search_error_count: int
    citation_mapping_ok: bool
    notes: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


def audit_references(reference_texts: Iterable[str], body_text: str, online: bool = True) -> ReferenceAuditResult:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    entries = [parse_reference_entry(index + 1, text) for index, text in enumerate(reference_texts)]
    citation_numbers = sorted({int(item) for item in re.findall(r"\[(\d+)\]", body_text)})
    dangling_citations = [number for number in citation_numbers if number > len(entries)]
    uncited_reference_indexes = [entry.index for entry in entries if entry.index not in citation_numbers]

    notes: list[str] = []
    if not citation_numbers:
        notes.append("正文未检测到规范的方括号编号引用，无法建立稳定的一一映射。")
    if dangling_citations:
        notes.append(f"正文出现无对应条目的引用编号: {dangling_citations}。")
    if uncited_reference_indexes:
        notes.append(f"参考文献中存在未在正文出现的条目: {uncited_reference_indexes}。")

    checks: list[ReferenceCheck] = []
    verified_count = 0
    possible_count = 0
    not_found_count = 0
    search_error_count = 0

    for entry in entries:
        cited_in_body = entry.index in citation_numbers
        if not online:
            checks.append(
                ReferenceCheck(
                    index=entry.index,
                    raw_text=entry.raw_text,
                    title=entry.title,
                    cited_in_body=cited_in_body,
                    status="offline",
                    confidence=0.0,
                    source=None,
                    matched_title=None,
                    matched_url=None,
                    notes=["联网核验关闭。"],
                )
            )
            continue

        try:
            check = verify_reference_entry(entry, cited_in_body, session)
        except Exception as exc:
            check = ReferenceCheck(
                index=entry.index,
                raw_text=entry.raw_text,
                title=entry.title,
                cited_in_body=cited_in_body,
                status="search_error",
                confidence=0.0,
                source=None,
                matched_title=None,
                matched_url=None,
                notes=[f"联网核验异常: {type(exc).__name__}: {exc}"],
            )

        checks.append(check)
        if check.status == "verified":
            verified_count += 1
        elif check.status == "possible_match":
            possible_count += 1
        elif check.status == "not_found":
            not_found_count += 1
        elif check.status == "search_error":
            search_error_count += 1

    return ReferenceAuditResult(
        entries=entries,
        checks=checks,
        citation_numbers=citation_numbers,
        dangling_citations=dangling_citations,
        uncited_reference_indexes=uncited_reference_indexes,
        verified_count=verified_count,
        possible_count=possible_count,
        not_found_count=not_found_count,
        search_error_count=search_error_count,
        citation_mapping_ok=bool(entries) and bool(citation_numbers) and not dangling_citations and not uncited_reference_indexes,
        notes=notes,
    )


def parse_reference_entry(index: int, raw_text: str) -> ReferenceEntry:
    cleaned = raw_text.strip()
    title = _extract_reference_title(cleaned)
    authors = _extract_reference_authors(cleaned)
    year_match = re.search(r"(19|20)\d{2}", cleaned)
    year = year_match.group(0) if year_match else None
    return ReferenceEntry(index=index, raw_text=cleaned, title=title, authors=authors, year=year)


def verify_reference_entry(entry: ReferenceEntry, cited_in_body: bool, session: requests.Session) -> ReferenceCheck:
    if not entry.title:
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="not_found",
            confidence=0.0,
            source=None,
            matched_title=None,
            matched_url=None,
            notes=["无法从参考文献条目中解析出标题。"],
        )
    if not _is_plausible_reference_title(entry.title):
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="not_found",
            confidence=0.0,
            source=None,
            matched_title=None,
            matched_url=None,
            notes=["参考文献的题名解析结果异常，当前条目更像作者、页码或期刊信息，不能作为文献属实依据。"],
        )

    notes: list[str] = []
    crossref_match = _attempt_search(lambda: _search_crossref(entry, session), "Crossref", notes)
    if crossref_match and crossref_match["confidence"] >= 0.86:
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="verified",
            confidence=crossref_match["confidence"],
            source="crossref",
            matched_title=crossref_match["title"],
            matched_url=crossref_match["url"],
            notes=notes + ["Crossref 找到高置信度匹配。"],
        )

    cnki_match = _attempt_search(lambda: _search_bing_site(entry.title, ["cnki.net", "cnki.com.cn"], session), "CNKI/Bing", notes)
    if cnki_match and cnki_match["confidence"] >= 0.82:
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="verified",
            confidence=cnki_match["confidence"],
            source="bing_site_cnki",
            matched_title=cnki_match["title"],
            matched_url=cnki_match["url"],
            notes=notes + ["通过 Bing 的 CNKI 域名检索命中高置信度结果。"],
        )

    generic_match = _attempt_search(lambda: _search_bing_site(entry.title, [], session), "Bing", notes)
    if generic_match and generic_match["confidence"] >= 0.88:
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="verified",
            confidence=generic_match["confidence"],
            source="bing_web",
            matched_title=generic_match["title"],
            matched_url=generic_match["url"],
            notes=notes + ["通过通用网页检索命中高置信度结果。"],
        )

    possible = _best_match([crossref_match, cnki_match, generic_match])
    mismatch_note = None
    if possible:
        if _is_possible_title_match(entry.title, possible["title"]):
            return ReferenceCheck(
                index=entry.index,
                raw_text=entry.raw_text,
                title=entry.title,
                cited_in_body=cited_in_body,
                status="possible_match",
                confidence=possible["confidence"],
                source=possible["source"],
                matched_title=possible["title"],
                matched_url=possible["url"],
                notes=notes + ["找到了疑似结果，但匹配度还不足以判定为真。"],
            )
        mismatch_note = "候选结果只命中局部关键词，核心题名片段对不上，不能据此认定文献属实。"

    if notes:
        return ReferenceCheck(
            index=entry.index,
            raw_text=entry.raw_text,
            title=entry.title,
            cited_in_body=cited_in_body,
            status="search_error",
            confidence=0.0,
            source=None,
            matched_title=None,
            matched_url=None,
            notes=notes + ([mismatch_note] if mismatch_note else []),
        )

    return ReferenceCheck(
        index=entry.index,
        raw_text=entry.raw_text,
        title=entry.title,
        cited_in_body=cited_in_body,
        status="not_found",
        confidence=0.0,
        source=None,
        matched_title=None,
        matched_url=None,
        notes=[mismatch_note or "联网检索未找到可信匹配，存在杜撰或格式异常风险。"],
    )


def _attempt_search(search_fn, label: str, notes: list[str]) -> dict | None:
    try:
        return search_fn()
    except Exception as exc:
        notes.append(f"{label} 检索异常: {type(exc).__name__}: {exc}")
        return None


def _search_crossref(entry: ReferenceEntry, session: requests.Session) -> dict | None:
    response = session.get(
        "https://api.crossref.org/works",
        params={"query.bibliographic": entry.raw_text, "rows": 5},
        timeout=SEARCH_TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    items = response.json().get("message", {}).get("items", [])
    best: dict | None = None
    for item in items:
        candidate_title = " ".join(item.get("title", [])).strip()
        if not candidate_title or not _is_plausible_reference_title(candidate_title):
            continue
        confidence = _title_confidence(entry.title, candidate_title, entry.year, item.get("published-print") or item.get("published-online"))
        match = {
            "source": "crossref",
            "title": candidate_title,
            "url": item.get("URL"),
            "confidence": confidence,
        }
        if best is None or match["confidence"] > best["confidence"]:
            best = match
    return best


def _search_bing_site(title: str, domains: list[str], session: requests.Session) -> dict | None:
    quoted_title = f"\"{title}\""
    if domains:
        query = quoted_title + " " + " OR ".join(f"site:{domain}" for domain in domains)
    else:
        query = quoted_title

    response = session.get(
        "https://cn.bing.com/search",
        params={"q": query},
        timeout=SEARCH_TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    return _best_bing_result(title, response.text, domains)


def _best_bing_result(title: str, html: str, domains: list[str]) -> dict | None:
    soup = BeautifulSoup(html, "html.parser")
    best: dict | None = None
    for result in soup.select("li.b_algo")[:5]:
        link = result.select_one("h2 a")
        if not link:
            continue
        candidate_title = unescape(link.get_text(" ", strip=True))
        if not _is_plausible_reference_title(candidate_title):
            continue
        candidate_url = link.get("href")
        snippet = unescape(result.get_text(" ", strip=True))
        if domains and candidate_url and not any(domain in candidate_url for domain in domains):
            continue
        confidence = max(
            _title_similarity(title, candidate_title),
            _title_similarity(title, snippet),
        )
        match = {
            "source": "bing_site_cnki" if domains else "bing_web",
            "title": candidate_title,
            "url": candidate_url,
            "confidence": confidence,
        }
        if best is None or match["confidence"] > best["confidence"]:
            best = match
    return best


def _best_match(matches: list[dict | None]) -> dict | None:
    candidates = [match for match in matches if match]
    if not candidates:
        return None
    return max(candidates, key=lambda item: item["confidence"])


def _extract_reference_title(raw_text: str) -> str:
    text = re.sub(r"^\[\d+\]\s*", "", raw_text)
    text = re.sub(r"（[^）]*）\s*$", "", text).strip()

    year_pattern = re.search(r"[\(（](19|20)\d{2}[\)）]\.?\s*(.+)", text)
    if year_pattern:
        tail = year_pattern.group(2)
        title = _take_until_publication_segment(tail)
        if title:
            return title

    segments = [segment.strip() for segment in re.split(r"[。\.．]", text) if segment.strip()]
    if len(segments) >= 2:
        for segment in segments[1:]:
            if len(segment) >= 4 and not re.fullmatch(r"(19|20)\d{2}", segment):
                return segment
    return text


def _extract_reference_authors(raw_text: str) -> str:
    text = re.sub(r"^\[\d+\]\s*", "", raw_text)
    year_pattern = re.split(r"[\(（](19|20)\d{2}[\)）]", text, maxsplit=1)
    if len(year_pattern) > 1:
        return year_pattern[0].strip(" .。．")
    segments = [segment.strip() for segment in re.split(r"[。\.．]", text) if segment.strip()]
    return segments[0] if segments else ""


def _take_until_publication_segment(tail: str) -> str:
    candidates = [segment.strip() for segment in re.split(r"[。\.．]", tail) if segment.strip()]
    if not candidates:
        return tail.strip()
    return candidates[0]


def _title_confidence(title: str, candidate_title: str, year: str | None, published_part) -> float:
    profile = _title_match_profile(title, candidate_title)
    confidence = profile["confidence"]
    if year and published_part:
        published_year = None
        if isinstance(published_part, dict):
            date_parts = published_part.get("date-parts", [])
            if date_parts and date_parts[0]:
                published_year = str(date_parts[0][0])
        if published_year == year:
            confidence += 0.04
    return min(confidence, 1.0)


def _title_similarity(expected: str, candidate: str) -> float:
    return _title_match_profile(expected, candidate)["confidence"]


def _normalize_title(value: str) -> str:
    text = re.sub(r"[\s\u3000]+", "", value or "")
    text = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", text)
    return text.lower()


def _is_plausible_reference_title(value: str) -> bool:
    raw = (value or "").strip(" .,:;，；、()（）[]{}")
    normalized = _normalize_title(raw)
    if len(normalized) < 4:
        return False

    letter_count = len(re.findall(r"[A-Za-z\u4e00-\u9fff]", normalized))
    digit_count = len(re.findall(r"\d", normalized))
    if letter_count == 0:
        return False
    if digit_count >= max(letter_count * 2, 4) and letter_count < 4:
        return False
    if re.fullmatch(r"[\d\s\-/.:()（）]+", raw):
        return False

    author_like_parts = [part.strip() for part in re.split(r"[;；,，、/&]+", raw) if part.strip()]
    if len(author_like_parts) >= 2:
        normalized_parts = [_normalize_title(part) for part in author_like_parts]
        if all(1 <= len(part) <= 5 for part in normalized_parts):
            return False

    return True


def _split_title_terms(normalized_title: str) -> set[str]:
    terms = set()
    for size in range(2, min(7, len(normalized_title) + 1)):
        for index in range(0, len(normalized_title) - size + 1):
            term = normalized_title[index : index + size]
            if term in TITLE_STOPWORDS:
                continue
            terms.add(term)
    return terms


def _is_strict_title_match(expected: str, candidate: str) -> bool:
    profile = _title_match_profile(expected, candidate)
    if profile["exact"]:
        return True
    if profile["containment_ratio"] >= 0.9 and profile["signature_hit_ratio"] >= 0.5:
        return True
    if profile["edit_similarity"] >= 0.92 and profile["expected_term_coverage"] >= 0.9:
        return True
    return profile["edit_similarity"] >= 0.88 and profile["signature_hit_ratio"] >= 1.0


def _is_possible_title_match(expected: str, candidate: str) -> bool:
    profile = _title_match_profile(expected, candidate)
    if _is_strict_title_match(expected, candidate):
        return True
    return (
        profile["signature_hit_ratio"] >= 0.5
        and profile["edit_similarity"] >= 0.82
        and profile["expected_term_coverage"] >= 0.7
    )


def _title_match_profile(expected: str, candidate: str) -> dict[str, float | bool]:
    expected_norm = _normalize_title(expected)
    candidate_norm = _normalize_title(candidate)
    if not expected_norm or not candidate_norm:
        return {
            "exact": False,
            "containment_ratio": 0.0,
            "edit_similarity": 0.0,
            "expected_term_coverage": 0.0,
            "signature_hit_ratio": 0.0,
            "confidence": 0.0,
        }

    exact = expected_norm == candidate_norm
    containment_ratio = 0.0
    if exact:
        containment_ratio = 1.0
    elif expected_norm in candidate_norm:
        containment_ratio = len(expected_norm) / max(len(candidate_norm), 1)
    elif candidate_norm in expected_norm:
        containment_ratio = len(candidate_norm) / max(len(expected_norm), 1)

    edit_similarity = _edit_similarity(expected_norm, candidate_norm)
    expected_terms = _split_title_terms(expected_norm)
    candidate_terms = _split_title_terms(candidate_norm)
    overlap = len(expected_terms & candidate_terms)
    expected_term_coverage = overlap / max(len(expected_terms), 1)
    signature_hit_ratio = _signature_hit_ratio(expected_norm, candidate_norm)

    if exact:
        confidence = 1.0
    else:
        confidence = 0.45 * edit_similarity + 0.25 * expected_term_coverage + 0.30 * signature_hit_ratio
        if containment_ratio >= 0.95:
            confidence = max(confidence, 0.97)
        elif containment_ratio >= 0.8 and signature_hit_ratio >= 0.5:
            confidence = max(confidence, 0.9)
        if signature_hit_ratio == 0.0 and edit_similarity < 0.9:
            confidence = min(confidence, 0.69)
        elif signature_hit_ratio < 0.5 and edit_similarity < 0.86:
            confidence = min(confidence, 0.78)
        elif signature_hit_ratio < 1.0 and edit_similarity < 0.82:
            confidence = min(confidence, 0.84)

    return {
        "exact": exact,
        "containment_ratio": containment_ratio,
        "edit_similarity": edit_similarity,
        "expected_term_coverage": expected_term_coverage,
        "signature_hit_ratio": signature_hit_ratio,
        "confidence": min(max(confidence, 0.0), 1.0),
    }


def _edit_similarity(expected_norm: str, candidate_norm: str) -> float:
    longer = max(len(expected_norm), len(candidate_norm))
    if longer == 0:
        return 0.0

    previous = list(range(len(candidate_norm) + 1))
    for row, expected_char in enumerate(expected_norm, 1):
        current = [row]
        for col, candidate_char in enumerate(candidate_norm, 1):
            substitution_cost = 0 if expected_char == candidate_char else 1
            current.append(
                min(
                    previous[col] + 1,
                    current[col - 1] + 1,
                    previous[col - 1] + substitution_cost,
                )
            )
        previous = current
    return 1.0 - previous[-1] / longer


def _signature_hit_ratio(expected_norm: str, candidate_norm: str) -> float:
    fragments = _signature_fragments(expected_norm)
    if not fragments:
        return 0.0
    hit_count = sum(1 for fragment in fragments if fragment in candidate_norm or candidate_norm in fragment)
    return hit_count / len(fragments)


def _signature_fragments(normalized_title: str) -> list[str]:
    fragments: list[str] = []
    for raw_fragment in TITLE_FRAGMENT_SPLIT_RE.split(normalized_title):
        fragment = _trim_generic_fragment(raw_fragment)
        if len(fragment) < 2 or fragment in TITLE_STOPWORDS:
            continue
        fragments.append(fragment)

    if not fragments and len(normalized_title) >= 2:
        fragments.append(normalized_title)
    return _dedupe_preserve_order(fragments)


def _trim_generic_fragment(fragment: str) -> str:
    value = fragment.strip()
    while True:
        matched_suffix = next(
            (suffix for suffix in GENERIC_FRAGMENT_SUFFIXES if value.endswith(suffix) and len(value) - len(suffix) >= 2),
            None,
        )
        if not matched_suffix:
            return value
        value = value[: -len(matched_suffix)]


def _dedupe_preserve_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    unique_items: list[str] = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        unique_items.append(item)
    return unique_items
