import logging
from datetime import datetime, timezone

import requests
from open_webui.retrieval.web.main import SearchResult
from open_webui.retrieval.web.quality import (
    build_query_pack,
    clean_content,
    compute_source_quality_score,
    dedupe_candidates,
    extract_evidence_spans,
    extract_published_date,
    infer_source_type,
)
from yarl import URL

log = logging.getLogger(__name__)


def search_jina(
    api_key: str,
    query: str,
    count: int,
    base_url: str = "",
    enhanced_mode: bool = False,
    strict_authority: bool = False,
    max_candidates: int = 24,
    max_evidence_items: int = 8,
) -> list[SearchResult]:
    """
    Search using Jina's Search API and return the results as a list of SearchResult objects.
    Args:
        api_key (str): The Jina API key
        query (str): The query to search for
        count (int): The number of results to return
        base_url (str): Optional custom base URL for the Jina API

    Returns:
        list[SearchResult]: A list of search results
    """
    count = max(1, int(count or 3))
    max_candidates = max(1, int(max_candidates or 24))
    max_evidence_items = max(1, int(max_evidence_items or 8))
    jina_search_endpoint = base_url if base_url else "https://s.jina.ai/"

    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": api_key,
        "X-Retain-Images": "none",
    }

    url = str(URL(jina_search_endpoint))
    if not enhanced_mode:
        payload = {"q": query, "count": count if count <= 10 else 10}
        response = requests.post(url, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()

        results = []
        for result in data.get("data", []):
            url_value = result.get("url")
            if not url_value:
                continue
            content = clean_content(result.get("content") or "")
            published_date = extract_published_date(content)
            source_quality_score = compute_source_quality_score(
                query=query,
                url=url_value,
                content=content,
                published_date=published_date,
                strict_authority=strict_authority,
            )
            results.append(
                SearchResult(
                    link=url_value,
                    title=result.get("title"),
                    snippet=content,
                    source_quality_score=source_quality_score,
                    source_type=infer_source_type(url_value),
                    published_date=published_date,
                    retrieved_at=datetime.now(timezone.utc).isoformat(),
                )
            )
        return results

    query_pack = build_query_pack(query=query, max_variants=5)
    max_candidates = max(max_candidates, count)
    per_query_count = min(10, max(1, max_candidates // max(len(query_pack), 1)))

    raw_candidates = []
    for variant in query_pack:
        payload = {"q": variant, "count": per_query_count}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=60)
            response.raise_for_status()
            for result in response.json().get("data", []):
                raw_candidates.append(
                    {
                        "url": result.get("url", ""),
                        "title": result.get("title"),
                        "content": result.get("content") or "",
                        "query_variant": variant,
                    }
                )
        except Exception as e:
            log.warning("Jina enhanced query variant failed (%s): %s", variant, e)

    candidates = dedupe_candidates(raw_candidates)

    scored = []
    for candidate in candidates:
        published_date = extract_published_date(candidate.get("content", ""))
        score = compute_source_quality_score(
            query=query,
            url=candidate["canonical_url"],
            content=candidate.get("content", ""),
            published_date=published_date,
            strict_authority=strict_authority,
        )
        evidence = extract_evidence_spans(
            query=query,
            content=candidate.get("content", ""),
            source_url=candidate["canonical_url"],
            published_date=published_date,
            source_quality_score=score,
            max_items=max_evidence_items,
        )
        scored.append(
            {
                **candidate,
                "published_date": published_date,
                "source_quality_score": score,
                "source_type": infer_source_type(candidate["canonical_url"]),
                "evidence_spans": evidence,
            }
        )

    scored.sort(key=lambda item: item["source_quality_score"], reverse=True)
    top_results = scored[:count]
    retrieved_at = datetime.now(timezone.utc).isoformat()

    results: list[SearchResult] = []
    for item in top_results:
        snippet = item.get("content", "")
        results.append(
            SearchResult(
                link=item["canonical_url"],
                title=item.get("title"),
                snippet=snippet,
                source_quality_score=item["source_quality_score"],
                source_type=item["source_type"],
                published_date=item["published_date"],
                retrieved_at=retrieved_at,
                content_hash=item.get("content_hash"),
                evidence_spans=item.get("evidence_spans"),
                query_variant=item.get("query_variant"),
            )
        )
    return results
