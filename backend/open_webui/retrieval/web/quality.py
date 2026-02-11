import hashlib
import re
from datetime import datetime, timezone
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

TRACKING_PARAMS = {
    "utm_source",
    "utm_medium",
    "utm_campaign",
    "utm_term",
    "utm_content",
    "gclid",
    "fbclid",
    "mc_cid",
    "mc_eid",
}

STOPWORDS = {
    "the",
    "a",
    "an",
    "and",
    "or",
    "for",
    "to",
    "of",
    "in",
    "on",
    "at",
    "by",
    "with",
    "from",
    "latest",
    "changes",
    "change",
}

NEWS_DOMAINS = {
    "reuters.com",
    "apnews.com",
    "bloomberg.com",
    "wsj.com",
    "ft.com",
    "nytimes.com",
}

OFFICIAL_HINTS = {
    ".gov",
    ".mil",
    "federalregister.gov",
    "congress.gov",
    "whitehouse.gov",
    "uscourts.gov",
    "sec.gov",
    "treasury.gov",
    "ustr.gov",
}


def canonicalize_url(url: str) -> str:
    try:
        parsed = urlparse(url)
        host = (parsed.hostname or "").lower()
        path = parsed.path or "/"

        query = [
            (k, v)
            for k, v in parse_qsl(parsed.query, keep_blank_values=True)
            if k.lower() not in TRACKING_PARAMS
        ]
        query.sort(key=lambda item: item[0])

        netloc = host
        if parsed.port and parsed.port not in (80, 443):
            netloc = f"{host}:{parsed.port}"

        return urlunparse(
            (
                parsed.scheme or "https",
                netloc,
                path.rstrip("/") or "/",
                "",
                urlencode(query),
                "",
            )
        )
    except Exception:
        return url


def clean_content(text: str) -> str:
    text = text or ""
    text = re.sub(r"!\[[^\]]*]\([^)]+\)", " ", text)
    text = re.sub(r"\[([^\]]+)]\([^)]+\)", r"\1", text)
    text = re.sub(r"\[\]\([^)]+\)", " ", text)
    text = re.sub(r"https?://\S+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def sha256_text(text: str) -> str:
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def extract_published_date(text: str) -> str | None:
    if not text:
        return None

    patterns = [
        (r"\b(\d{4})-(\d{2})-(\d{2})\b", "%Y-%m-%d"),
        (r"\b(\d{2})/(\d{2})/(\d{4})\b", "%m/%d/%Y"),
        (r"\b([A-Z][a-z]+ \d{1,2}, \d{4})\b", "%B %d, %Y"),
    ]

    for pattern, fmt in patterns:
        match = re.search(pattern, text)
        if not match:
            continue
        value = match.group(0)
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue

    return None


def infer_source_type(url: str) -> str:
    host = (urlparse(url).hostname or "").lower()
    if any(hint in host for hint in OFFICIAL_HINTS):
        return "official"
    if host.endswith(".edu"):
        return "academic"
    if any(host.endswith(news) or host == news for news in NEWS_DOMAINS):
        return "news"
    if "wikipedia.org" in host:
        return "wiki"
    if any(token in host for token in ("substack", "medium", "blog")):
        return "blog"
    return "commercial"


def build_query_pack(query: str, max_variants: int = 5) -> list[str]:
    base = (query or "").strip()
    if not base:
        return []

    year = datetime.now(timezone.utc).year
    variants = [
        base,
        f"{base} latest updates {year}",
        f"{base} official source government regulator",
        f"\"{base}\"",
        f"{base} timeline effective date",
    ]

    deduped: list[str] = []
    seen: set[str] = set()
    for variant in variants:
        key = variant.lower().strip()
        if key and key not in seen:
            deduped.append(variant)
            seen.add(key)
        if len(deduped) >= max_variants:
            break
    return deduped


def _query_terms(query: str) -> set[str]:
    parts = re.findall(r"[A-Za-z0-9]+", (query or "").lower())
    return {part for part in parts if part not in STOPWORDS and len(part) >= 3}


def compute_source_quality_score(
    *,
    query: str,
    url: str,
    content: str,
    published_date: str | None,
    strict_authority: bool,
) -> float:
    source_type = infer_source_type(url)
    authority_map = {
        "official": 1.0,
        "academic": 0.9,
        "news": 0.78,
        "wiki": 0.55 if strict_authority else 0.68,
        "commercial": 0.42 if strict_authority else 0.58,
        "blog": 0.3 if strict_authority else 0.48,
    }
    authority = authority_map.get(source_type, 0.5)

    cleaned = clean_content(content)
    density = min(len(cleaned) / 2200.0, 1.0)

    terms = _query_terms(query)
    text_terms = set(re.findall(r"[A-Za-z0-9]+", cleaned.lower()))
    overlap = len(terms.intersection(text_terms))
    relevance = min(overlap / max(len(terms), 1), 1.0)

    freshness = 0.5
    if published_date:
        try:
            dt = datetime.fromisoformat(published_date)
            days = (datetime.now() - dt).days
            if days <= 30:
                freshness = 1.0
            elif days <= 180:
                freshness = 0.8
            elif days <= 365:
                freshness = 0.65
            else:
                freshness = 0.45
        except ValueError:
            pass

    spam_penalty = 0.0
    url_lower = url.lower()
    if any(token in url_lower for token in ("affiliate", "sponsored", "utm_")):
        spam_penalty = 0.2

    score = (
        authority * 0.45 + relevance * 0.3 + density * 0.15 + freshness * 0.1
    ) - spam_penalty
    return max(0.0, min(1.0, round(score, 4)))


def extract_evidence_spans(
    *,
    query: str,
    content: str,
    source_url: str,
    published_date: str | None,
    source_quality_score: float,
    max_items: int = 4,
) -> list[dict[str, Any]]:
    cleaned = clean_content(content)
    sentences = re.split(r"(?<=[.!?])\s+", cleaned)
    terms = _query_terms(query)
    now_iso = datetime.now(timezone.utc).isoformat()

    ranked: list[tuple[float, str]] = []
    for sentence in sentences:
        sentence = sentence.strip()
        if len(sentence) < 35:
            continue
        sentence_terms = set(re.findall(r"[A-Za-z0-9]+", sentence.lower()))
        overlap = len(terms.intersection(sentence_terms))
        signal = 0.15 if re.search(r"\b\d+(\.\d+)?%?\b", sentence) else 0.0
        score = overlap + signal
        if score <= 0:
            continue
        ranked.append((score, sentence))

    ranked.sort(key=lambda item: item[0], reverse=True)
    evidence: list[dict[str, Any]] = []
    for score, sentence in ranked[:max_items]:
        confidence = max(0.0, min(1.0, round(source_quality_score * (0.5 + score / 6), 3)))
        evidence.append(
            {
                "claim_text": sentence[:320],
                "supporting_quote": sentence[:320],
                "source_url": source_url,
                "published_date": published_date,
                "retrieved_at": now_iso,
                "confidence": confidence,
            }
        )
    return evidence


def dedupe_candidates(candidates: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    seen_urls: set[str] = set()
    seen_hashes: set[str] = set()
    deduped: list[dict[str, Any]] = []

    for candidate in candidates:
        canonical_url = canonicalize_url(candidate.get("url", ""))
        if not canonical_url:
            continue
        content = clean_content(candidate.get("content") or "")
        content_hash = sha256_text(content[:6000])

        if canonical_url in seen_urls or content_hash in seen_hashes:
            continue

        seen_urls.add(canonical_url)
        seen_hashes.add(content_hash)
        enriched = dict(candidate)
        enriched["canonical_url"] = canonical_url
        enriched["content_hash"] = content_hash
        enriched["content"] = content
        deduped.append(enriched)

    return deduped
