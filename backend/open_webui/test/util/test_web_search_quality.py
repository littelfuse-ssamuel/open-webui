from open_webui.retrieval.web.jina_search import search_jina
from open_webui.retrieval.web.quality import (
    build_query_pack,
    canonicalize_url,
    compute_source_quality_score,
    dedupe_candidates,
    extract_evidence_spans,
)


def test_query_pack_generation_is_deterministic():
    query = "latest tariff changes 2025 United States"
    first = build_query_pack(query, max_variants=5)
    second = build_query_pack(query, max_variants=5)
    assert first == second
    assert len(first) >= 3


def test_url_canonicalization_removes_tracking_query_params():
    url = "https://example.com/path?a=1&utm_source=x&fbclid=y#section"
    assert canonicalize_url(url) == "https://example.com/path?a=1"


def test_dedupe_candidates_uses_canonical_url_and_content_hash():
    candidates = [
        {"url": "https://example.com/a?utm_source=x", "content": "Hello world"},
        {"url": "https://example.com/a", "content": "Hello world"},
        {"url": "https://example.com/b", "content": "Different content"},
    ]
    deduped = dedupe_candidates(candidates)
    assert len(deduped) == 2


def test_quality_score_prefers_official_source_for_same_content():
    query = "federal tariff update"
    content = "Executive Order 14257 effective April 2, 2025."
    official = compute_source_quality_score(
        query=query,
        url="https://www.federalregister.gov/d/2025-06063",
        content=content,
        published_date="2025-04-02",
        strict_authority=True,
    )
    blog = compute_source_quality_score(
        query=query,
        url="https://exampleblog.substack.com/post/tariff-update",
        content=content,
        published_date="2025-04-02",
        strict_authority=True,
    )
    assert official > blog


def test_evidence_extraction_returns_claim_with_citation_fields():
    content = (
        "Executive Order 14257 took effect on April 2, 2025. "
        "A 10% baseline tariff became effective for most imports."
    )
    evidence = extract_evidence_spans(
        query="tariff baseline effective date",
        content=content,
        source_url="https://www.whitehouse.gov/example",
        published_date="2025-04-02",
        source_quality_score=0.9,
        max_items=3,
    )
    assert evidence
    first = evidence[0]
    assert first["source_url"] == "https://www.whitehouse.gov/example"
    assert "claim_text" in first
    assert "confidence" in first


def test_enhanced_jina_search_ranks_official_domain_higher(monkeypatch):
    class _Response:
        def __init__(self, data):
            self._data = data

        def raise_for_status(self):
            return None

        def json(self):
            return {"data": self._data}

    def fake_post(*args, **kwargs):
        return _Response(
            [
                {
                    "url": "https://exampleblog.substack.com/post/tariff-list",
                    "title": "Tariff tracker",
                    "content": "Tariff tracker with estimates and commentary only.",
                },
                {
                    "url": "https://www.federalregister.gov/d/2025-06063",
                    "title": "Executive order",
                    "content": "Executive Order 14257 effective April 2, 2025 sets duties.",
                },
            ]
        )

    monkeypatch.setattr("open_webui.retrieval.web.jina_search.requests.post", fake_post)

    results = search_jina(
        api_key="",
        query="US tariff executive order changes",
        count=2,
        base_url="https://s.jina.ai/",
        enhanced_mode=True,
        strict_authority=True,
        max_candidates=10,
        max_evidence_items=2,
    )
    assert len(results) == 2
    assert results[0].link.startswith("https://www.federalregister.gov/")
    assert results[0].source_quality_score is not None
    assert results[0].evidence_spans is not None
