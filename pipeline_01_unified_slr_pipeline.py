"""
SLR Automation Pipeline for MathTutorAI Literature Review

This script:
1. Loads exported records from CSV/XLSX files.
2. Searches the five databases used in the review - OpenAlex, Crossref,
   ERIC, Semantic Scholar, and arXiv - all free and requiring no paid
   subscription (no Scopus, Web of Science, IEEE Xplore, or ACM access).
   Optional connectors for DBLP, DOAJ, and CORE are included in the code
   but disabled by default, so a default run reproduces exactly the
   sources reported in the manuscript; the supplementary Consensus export
   is incorporated separately as an input file.
3. Retries failed API requests automatically.
4. Normalizes metadata.
5. Deduplicates records by DOI and title.
6. Tags records by RQ1–RQ6.
7. Creates manual screening columns.
8. Generates initial PRISMA-style counts.
9. Exports an Excel workbook and CSV files.

Important:
This script supports the SLR process but does not replace manual review.
Final inclusion/exclusion decisions must be checked manually.
"""

import os
import re
import time
import json
import hashlib
import pandas as pd
import xml.etree.ElementTree as ET

from pathlib import Path
from difflib import SequenceMatcher
from urllib.parse import urlencode, quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

try:
    import requests
except ImportError:
    class _SimpleRequestException(Exception):
        """Fallback request exception used when requests is unavailable."""


    class _SimpleResponse:
        """Minimal response wrapper compatible with the parts we use."""

        def __init__(self, status_code, text):
            self.status_code = status_code
            self.text = text

        def raise_for_status(self):
            if self.status_code >= 400:
                raise _SimpleRequestException(f"HTTP {self.status_code}")

        def json(self):
            return json.loads(self.text)


    class _RequestsFallback:
        class exceptions:
            RequestException = _SimpleRequestException

        @staticmethod
        def get(url, params=None, timeout=60, headers=None):
            if params:
                query_string = urlencode(params, doseq=True)
                separator = "&" if "?" in url else "?"
                url = f"{url}{separator}{query_string}"

            request = Request(url, headers=headers or {})

            try:
                with urlopen(request, timeout=timeout) as response:
                    payload = response.read().decode("utf-8")
                    status_code = getattr(response, "status", response.getcode())
                    return _SimpleResponse(status_code, payload)
            except HTTPError as exc:
                payload = exc.read().decode("utf-8", errors="replace")
                return _SimpleResponse(exc.code, payload)
            except (URLError, OSError) as exc:
                raise _SimpleRequestException(str(exc)) from exc

    requests = _RequestsFallback()

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **_kwargs):
        """Fallback tqdm passthrough when the package is unavailable."""
        return iterable


# ============================================================
# 1. CONFIGURATION
# ============================================================

PROJECT_NAME = "MathTutorAI_SLR"

INPUT_DIR = Path("slr_exports")
OUTPUT_DIR = Path("slr_outputs")

OUTPUT_DIR.mkdir(exist_ok=True)
INPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "MathTutorAI_SLR_screening_table.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "MathTutorAI_SLR_screening_table.csv"
RAW_RECORDS_CSV = OUTPUT_DIR / "MathTutorAI_raw_records.csv"
DUPLICATES_CSV = OUTPUT_DIR / "MathTutorAI_duplicate_log.csv"
PRISMA_COUNTS_CSV = OUTPUT_DIR / "MathTutorAI_PRISMA_initial_counts.csv"
SEARCH_AUDIT_CSV = OUTPUT_DIR / "MathTutorAI_search_audit.csv"

# Replace this with your real email.
# It helps OpenAlex/Crossref identify API requests.
USER_EMAIL = "sobsoire@gmail.com"


def env_flag(name, default=True):
    """Read a boolean flag from the environment."""
    value = os.getenv(name)

    if value is None:
        return default

    return value.strip().lower() not in {"0", "false", "no", "off"}

# API search settings.
USE_OPENALEX = env_flag("SLR_USE_OPENALEX", True)
USE_CROSSREF = env_flag("SLR_USE_CROSSREF", True)
USE_ERIC = env_flag("SLR_USE_ERIC", True)
# Additional FREE sources (no paid subscription required).
USE_SEMANTIC_SCHOLAR = env_flag("SLR_USE_SEMANTIC_SCHOLAR", True)
USE_ARXIV = env_flag("SLR_USE_ARXIV", True)
# DBLP, DOAJ and CORE are optional extra connectors. They are DISABLED by
# default so a default run reproduces exactly the five databases reported in
# the manuscript (OpenAlex, Crossref, ERIC, Semantic Scholar, arXiv). Set the
# matching environment variable to "1" to re-enable one for exploratory use.
USE_DBLP = env_flag("SLR_USE_DBLP", False)
USE_DOAJ = env_flag("SLR_USE_DOAJ", False)
USE_CORE = env_flag("SLR_USE_CORE", False)  # only runs if CORE_API_KEY is set
USE_SEED_REFERENCE_ENRICHMENT = env_flag("SLR_USE_SEED_REFERENCE_ENRICHMENT", True)

# Broader retrieval helps reduce misses for recent arXiv-style papers.
OPENALEX_PAGES_PER_QUERY = 4
OPENALEX_PER_PAGE = 50

CROSSREF_ROWS_PER_QUERY = 75

ERIC_API_URL = "https://api.ies.ed.gov/eric/"
ERIC_ROWS_PER_CALL = 200
ERIC_MAX_RECORDS_PER_QUERY = 2000
ERIC_PEER_REVIEWED_ONLY = env_flag("SLR_ERIC_PEER_REVIEWED_ONLY", False)
ERIC_FIELDS = [
    "id", "title", "author", "source", "publicationdateyear", "description",
    "subject", "peerreviewed", "educationlevel", "publicationtype",
    "publisher", "url", "doi", "isbn", "issn", "language",
    "e_fulltextauth", "iescited", "sourceid", "e_datemodified"
]

# ---- Additional FREE sources (no subscription required) ----
# Semantic Scholar: free, no key needed. An optional free key (set
# S2_API_KEY in the environment) raises the rate limit.
SEMANTIC_SCHOLAR_LIMIT_PER_QUERY = 100  # API hard cap is 100 per request
SEMANTIC_SCHOLAR_API_KEY = os.getenv("S2_API_KEY", "")
SEMANTIC_SCHOLAR_FIELDS = (
    "title,abstract,year,venue,externalIds,authors,url,"
    "publicationTypes,citationCount,openAccessPdf"
)

# arXiv: free, no key. Be polite — the API asks for ~3s between calls.
ARXIV_MAX_RESULTS_PER_QUERY = 80

# DBLP: free, no key. Indexes the bibliographic records of essentially
# every IEEE/ACM computer-science venue, so it recovers the paywalled
# CS conference/journal metadata this review would otherwise miss.
DBLP_HITS_PER_QUERY = 100

# DOAJ: free, no key. Open-access journals only.
DOAJ_PAGE_SIZE = 100
DOAJ_MAX_PAGES = 2

# CORE: free, but needs a FREE API key from
# https://core.ac.uk/services/api . Leave CORE_API_KEY empty to skip it.
CORE_API_KEY = os.getenv("CORE_API_KEY", "")
CORE_LIMIT_PER_QUERY = 100

SEED_REFERENCE_FILENAME = "MathTutorAI_seed_references.csv"
MUST_INCLUDE_FILENAME = "MathTutorAI_must_include_references.csv"

# Year filter.
# Set to None if you do not want filtering.
YEAR_MIN = 2014
YEAR_MAX = 2026

# Retry settings.
MAX_RETRIES = 4
RETRY_SLEEP_SECONDS = 6
REQUEST_TIMEOUT_SECONDS = 60


# ============================================================
# 2. SEARCH STRATEGY
# ============================================================

TERM_GROUPS = {
    "tutoring_core": [
        "intelligent tutoring system",
        "AI-based tutoring",
        "AI tutor",
        "adaptive learning system"
    ],
    "tutoring_extended": [
        "cognitive tutor",
        "automated tutoring",
        "educational recommender"
    ],
    "education_context": [
        "educational system",
        "adaptive learning",
        "education",
        "tutoring"
    ],
    "math_core": [
        "mathematics",
        "math",
        "algebra",
        "geometry",
        "calculus",
        "mathematical reasoning"
    ],
    "math_problem_solving": [
        "math problem solving",
        "problem-solving",
        "step-by-step reasoning"
    ],
    "llm_and_symbolic": [
        "large language model",
        "LLM",
        "generative AI",
        "computer algebra system",
        "CAS",
        "SymPy",
        "multi-agent system",
        "LLM agents"
    ],
    "architecture_terms": [
        "architecture",
        "framework",
        "model",
        "pipeline",
        "system"
    ],
    "multi_agent_terms": [
        "multi-agent system",
        "multi-agent architecture",
        "agent-based architecture",
        "agentic AI",
        "LLM agents",
        "collaborative agents"
    ],
    "multi_agent_llm_terms": [
        "multi-agent",
        "LLM agents",
        "agentic workflow",
        "agent collaboration",
        "agent coordination"
    ],
    "personalization_terms": [
        "personalization",
        "personalisation",
        "personalized",
        "personalised",
        "adaptive",
        "learner model",
        "student model",
        "adaptation"
    ],
    "adaptation_terms": [
        "fine-tuning",
        "parameter-efficient fine-tuning",
        "PEFT",
        "LoRA"
    ],
    "correctness_terms": [
        "correctness",
        "verification",
        "reliability",
        "validation",
        "answer checking"
    ],
    "symbolic_verification_terms": [
        "computer algebra system",
        "CAS",
        "SymPy",
        "symbolic computation",
        "formal verification",
        "verifier"
    ],
    "llm_verification_terms": [
        "verifier",
        "verification",
        "self-correction",
        "answer validation",
        "formal verification",
        "symbolic verification"
    ],
    "evaluation_terms": [
        "evaluation",
        "assessment",
        "experiment",
        "user study",
        "learning gains",
        "usability",
        "effectiveness"
    ],
    "benchmark_terms": [
        "benchmark",
        "accuracy",
        "correctness",
        "reliability"
    ],
    "limitation_terms": [
        "limitations",
        "challenges",
        "research gap",
        "future work"
    ],
    "mathtutorai_gap_terms": [
        "personalization",
        "personalisation",
        "multi-agent",
        "verification",
        "curriculum alignment"
    ],
    "math_education_context": [
        "mathematics education",
        "math tutoring",
        "math exercises"
    ],
    "ai_education_context": [
        "artificial intelligence",
        "intelligent tutoring",
        "adaptive learning",
        "personalization",
        "personalisation"
    ]
}


SEARCH_STRATEGIES = {
    "MAIN_broad_AI_math_tutoring": {
        "description": "Primary broad search across tutoring, mathematics, enabling AI technologies, and adaptation.",
        "variants": [
            {
                "variant_label": "main_primary",
                "groups": [
                    "tutoring_core",
                    "math_core",
                    "llm_and_symbolic",
                    "personalization_terms"
                ]
            }
        ]
    },
    "RQ1_architectures": {
        "description": "Architectures and paradigms used in AI-based mathematics tutoring systems.",
        "variants": [
            {
                "variant_label": "rq1_primary",
                "groups": [
                    "tutoring_core",
                    "math_core",
                    "architecture_terms"
                ]
            }
        ]
    },
    "RQ2_multi_agent": {
        "description": "Multi-agent architectures in tutoring and mathematical reasoning systems.",
        "variants": [
            {
                "variant_label": "rq2_tutoring_agents",
                "groups": [
                    "multi_agent_terms",
                    "education_context",
                    "math_core"
                ]
            },
            {
                "variant_label": "rq2_llm_agents",
                "groups": [
                    ["large language model", "LLM", "generative AI"],
                    "multi_agent_llm_terms",
                    ["mathematics", "mathematical reasoning", "tutoring", "education"]
                ]
            }
        ]
    },
    "RQ3_personalization_adaptation": {
        "description": "Personalization, learner modeling, and model adaptation strategies.",
        "variants": [
            {
                "variant_label": "rq3_tutoring_personalization",
                "groups": [
                    "tutoring_core",
                    ["mathematics", "math", "algebra", "geometry"],
                    "personalization_terms"
                ]
            },
            {
                "variant_label": "rq3_llm_adaptation",
                "groups": [
                    ["large language model", "LLM"],
                    ["mathematics", "mathematical reasoning", "education", "tutoring"],
                    ["fine-tuning", "parameter-efficient fine-tuning", "PEFT", "LoRA", "adaptation", "personalization", "personalisation"]
                ]
            }
        ]
    },
    "RQ4_correctness_reliability": {
        "description": "Correctness, verification, and reliability in mathematical reasoning systems.",
        "variants": [
            {
                "variant_label": "rq4_symbolic_verification",
                "groups": [
                    ["mathematical reasoning", "math problem solving", "step-by-step reasoning"],
                    "correctness_terms",
                    "symbolic_verification_terms"
                ]
            },
            {
                "variant_label": "rq4_llm_verification",
                "groups": [
                    ["large language model", "LLM"],
                    ["mathematics", "mathematical reasoning", "math problem solving"],
                    "llm_verification_terms"
                ]
            }
        ]
    },
    "RQ5_evaluation": {
        "description": "Technical and pedagogical evaluation methods for AI-based mathematics tutoring systems.",
        "variants": [
            {
                "variant_label": "rq5_tutoring_evaluation",
                "groups": [
                    ["intelligent tutoring system", "AI-based tutoring", "AI tutor"],
                    ["mathematics", "math", "algebra", "geometry"],
                    "evaluation_terms"
                ]
            },
            {
                "variant_label": "rq5_reasoning_benchmarks",
                "groups": [
                    ["math problem solving", "mathematical reasoning", "AI tutor"],
                    ["evaluation", "benchmark", "accuracy", "correctness", "reliability"],
                    ["large language model", "LLM", "generative AI", "computer algebra system"]
                ]
            }
        ]
    },
    "RQ6_limitations_gap": {
        "description": "Limitations, research gaps, and local curriculum context motivating MathTutorAI.",
        "variants": [
            {
                "variant_label": "rq6_general_gap",
                "groups": [
                    ["AI-based mathematics tutoring", "intelligent tutoring system", "math tutor"],
                    "limitation_terms",
                    "mathtutorai_gap_terms"
                ]
            }
        ]
    }
}


# ============================================================
# 3. RQ TAGGING KEYWORDS
# ============================================================

RQ_KEYWORDS = {
    "RQ1_architecture": [
        "architecture", "framework", "pipeline", "system design",
        "system architecture", "model architecture", "platform",
        "intelligent tutoring system", "adaptive learning system",
        "cognitive tutor"
    ],

    "RQ2_multi_agent": [
        "multi-agent", "multi agent", "agentic", "llm agent", "llm agents",
        "agent collaboration", "agent coordination", "collaborative agents",
        "planner agent", "critic agent", "verifier agent",
        "multiagent"
    ],

    "RQ3_personalization": [
        "personalization", "personalisation", "personalized", "personalised",
        "adaptive", "adaptation", "learner model", "student model",
        "knowledge tracing", "bayesian knowledge tracing",
        "deep knowledge tracing", "fine-tuning", "fine tuning",
        "parameter-efficient", "peft", "lora"
    ],

    "RQ4_correctness_reliability": [
        "verification", "verifier", "correctness", "reliability",
        "validation", "answer checking", "answer validation",
        "symbolic", "computer algebra", "sympy", "formal verification",
        "self-correction", "self correction", "theorem prover",
        "proof assistant", "lean", "coq", "isabelle"
    ],

    "RQ5_evaluation": [
        "evaluation", "assessment", "experiment", "user study",
        "learning gains", "usability", "effectiveness", "benchmark",
        "accuracy", "performance", "empirical study", "case study"
    ],

    "RQ6_limitations_gap": [
        "limitation", "limitations", "challenge", "challenges",
        "future work", "research gap", "open problem",
        "curriculum alignment", "secondary education"
    ]
}


CORE_RELEVANCE_TERMS = [
    "intelligent tutoring", "ai tutor", "adaptive learning", "cognitive tutor",
    "mathematics", "math", "algebra", "geometry", "calculus",
    "mathematical reasoning", "math problem solving",
    "large language model", "llm", "generative ai",
    "multi-agent", "multi agent", "agentic",
    "personalization", "personalisation",
    "learner model", "student model", "knowledge tracing",
    "verification", "computer algebra", "sympy", "correctness",
    "evaluation", "learning gains"
]


MATH_SIGNAL_TERMS = [
    "mathematics", "math", "algebra", "geometry", "calculus",
    "mathematical reasoning", "math problem solving", "problem solving"
]

EDUCATION_SIGNAL_TERMS = [
    "tutor", "tutoring", "intelligent tutoring", "adaptive learning",
    "student model", "learner model", "education", "educational",
    "pedagogical", "classroom", "curriculum", "secondary education",
    "learning gains", "knowledge tracing"
]

AI_SIGNAL_TERMS = [
    "artificial intelligence", "ai tutor", "large language model", "llm",
    "generative ai", "computer algebra", "sympy", "multi-agent",
    "agentic", "verification", "adaptive", "personalization", "peft", "lora"
]

NON_SCHOLARLY_TITLE_TERMS = [
    "peer review report",
    "editorial",
    "table of contents",
    "preface",
    "correction",
    "erratum",
    "retraction"
]

OFF_SCOPE_DOMAIN_TERMS = [
    "healthcare",
    "medical",
    "clinical",
    "software engineering",
    "developer tools",
    "social science research",
    "high energy physics",
    "circuit analysis",
    "single cell",
    "bioinformatics",
    "fake review detection",
    "security verification",
    "robotics",
    "cell type identification"
]

GENERIC_FINE_TUNING_TERMS = [
    "parameter-efficient fine-tuning",
    "peft",
    "lora",
    "prompt tuning",
    "direct preference optimization"
]

GENERIC_VERIFICATION_TERMS = [
    "formal verification",
    "property checking",
    "correctness by construction",
    "arithmetic datapaths"
]


# ============================================================
# 4. BASIC HELPER FUNCTIONS
# ============================================================

def resolve_terms(group_or_terms):
    """Resolve a named term group or return the explicit term list."""
    if isinstance(group_or_terms, str):
        return TERM_GROUPS[group_or_terms]
    return list(group_or_terms)


def quote_search_term(term):
    """Quote a term when it contains spaces or punctuation."""
    term = clean_text(term)

    if re.fullmatch(r"[A-Za-z0-9]+", term):
        return term

    escaped = term.replace('"', '\\"')
    return f'"{escaped}"'


def build_boolean_group(group_or_terms):
    """Build a Boolean OR group from a named or explicit term set."""
    terms = resolve_terms(group_or_terms)
    return "(" + " OR ".join(quote_search_term(term) for term in terms) + ")"


def build_boolean_query(groups):
    """Build a Boolean AND query from multiple concept groups."""
    return " AND ".join(build_boolean_group(group) for group in groups)


def build_crossref_query(groups, max_terms_per_group=4):
    """
    Build a concise bibliographic query for Crossref.

    Crossref's query parameters are broad free-text searches rather than
    guaranteed Boolean parsing, so we keep a compact representation of the
    same concept groups for better precision.
    """
    pieces = []

    for group in groups:
        terms = resolve_terms(group)
        selected_terms = terms[:max_terms_per_group]
        pieces.extend(quote_search_term(term) for term in selected_terms)

    return " ".join(pieces)


def build_eric_query(groups):
    """Build an ERIC query from the configured concept groups."""
    query = build_boolean_query(groups)

    if ERIC_PEER_REVIEWED_ONLY:
        query = f"({query}) AND peerreviewed:T"

    return query


def build_compact_query(groups, max_terms_per_group=4, max_groups=None):
    """
    Build a compact free-text query.

    Used by Semantic Scholar, DOAJ, CORE, and DBLP, whose search
    endpoints are relevance-ranked free-text searches rather than strict
    Boolean parsers. Keeping the query short avoids over-constraining
    these engines (especially DBLP, which AND-matches every token).
    """
    pieces = []
    used_groups = groups if max_groups is None else groups[:max_groups]

    for group in used_groups:
        terms = resolve_terms(group)[:max_terms_per_group]
        pieces.extend(quote_search_term(term) for term in terms)

    return " ".join(pieces)


def build_arxiv_query(groups, max_terms_per_group=4, max_groups=3):
    """
    Build a Boolean query for the arXiv API.

    arXiv uses field prefixes (all:) with AND/OR and quoted phrases. We
    limit the number of groups and terms to keep the query length and
    selectivity reasonable.
    """
    and_parts = []

    for group in groups[:max_groups]:
        terms = resolve_terms(group)[:max_terms_per_group]
        or_parts = [f"all:{quote_search_term(term)}" for term in terms]

        if or_parts:
            and_parts.append("(" + " OR ".join(or_parts) + ")")

    return " AND ".join(and_parts)


def iter_search_jobs():
    """Yield every configured search job with reproducible query strings."""
    for search_label, strategy in SEARCH_STRATEGIES.items():
        for variant in strategy["variants"]:
            groups = variant["groups"]
            yield {
                "search_label": search_label,
                "variant_label": variant["variant_label"],
                "strategy_description": strategy["description"],
                "groups": groups,
                "boolean_query": build_boolean_query(groups),
                "openalex_query": build_boolean_query(groups),
                "crossref_query": build_crossref_query(groups),
                "eric_query": build_eric_query(groups),
                "semantic_scholar_query": build_compact_query(groups),
                "arxiv_query": build_arxiv_query(groups),
                "dblp_query": build_compact_query(groups, max_terms_per_group=2, max_groups=3),
                "doaj_query": build_compact_query(groups),
                "core_query": build_compact_query(groups)
            }


def search_jobs_to_df():
    """Convert configured search jobs to a reporting table."""
    rows = []

    for job in iter_search_jobs():
        rows.append({
            "search_label": job["search_label"],
            "variant_label": job["variant_label"],
            "strategy_description": job["strategy_description"],
            "boolean_query": job["boolean_query"],
            "openalex_query": job["openalex_query"],
            "crossref_query": job["crossref_query"],
            "eric_query": job["eric_query"],
            "semantic_scholar_query": job["semantic_scholar_query"],
            "arxiv_query": job["arxiv_query"],
            "dblp_query": job["dblp_query"],
            "doaj_query": job["doaj_query"],
            "core_query": job["core_query"]
        })

    return pd.DataFrame(rows)


def is_seed_reference_file(file_name):
    """Return True when the import file is the curated seed reference file."""
    return clean_text(file_name).lower() == SEED_REFERENCE_FILENAME.lower()


def is_control_file(file_name):
    """Return True for local helper files that should not be imported as records."""
    name = clean_text(file_name).lower()
    return name in {MUST_INCLUDE_FILENAME.lower()}


def is_consensus_export_file(file_name):
    """Return True for current imported CSV exports coming from Consensus."""
    name = clean_text(file_name).lower()
    return name.startswith("user_import_")

def clean_text(value):
    """Clean a text value safely."""
    if value is None:
        return ""

    if isinstance(value, (list, tuple, set)):
        cleaned_items = [clean_text(item) for item in value]
        return "; ".join(item for item in cleaned_items if item)

    if isinstance(value, dict):
        return json.dumps(value, sort_keys=True)

    if pd.isna(value):
        return ""

    value = str(value)

    # Remove simple HTML/XML tags sometimes found in Crossref abstracts.
    value = re.sub(r"<[^>]+>", " ", value)

    # Normalize spaces.
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_title(title):
    """Normalize title for deduplication."""
    title = clean_text(title).lower()
    title = title.replace("&amp;", "and")
    title = re.sub(r"[^\w\s]", "", title)
    title = re.sub(r"\s+", " ", title)
    return title.strip()


def normalize_doi(doi):
    """Normalize DOI for deduplication."""
    doi = clean_text(doi).lower()
    doi = doi.replace("https://doi.org/", "")
    doi = doi.replace("https://dx.doi.org/", "")
    doi = doi.replace("http://dx.doi.org/", "")
    doi = doi.replace("http://doi.org/", "")
    doi = doi.replace("doi:", "")
    return doi.strip()


def make_record_id(title, doi, fallback_id=""):
    """Create a stable record ID from DOI, title, or a source-specific ID."""
    doi_norm = normalize_doi(doi)
    title_norm = normalize_title(title)

    base = doi_norm if doi_norm else title_norm

    if not base:
        base = clean_text(fallback_id).lower()

    if not base:
        base = str(time.time())

    return hashlib.md5(base.encode("utf-8")).hexdigest()[:12]


def reconstruct_openalex_abstract(inverted_index):
    """Reconstruct OpenAlex abstract from inverted index."""
    if not inverted_index:
        return ""

    words = []

    for word, positions in inverted_index.items():
        for position in positions:
            words.append((position, word))

    words = sorted(words, key=lambda x: x[0])

    return " ".join(word for _, word in words)


def text_contains_any(text, keywords):
    """Return True if text contains any keyword."""
    text = text.lower()
    return any(keyword.lower() in text for keyword in keywords)


def calculate_relevance_score(title, abstract):
    """Simple keyword-based relevance score."""
    text = f"{title} {abstract}".lower()
    return sum(1 for term in CORE_RELEVANCE_TERMS if term.lower() in text)


def title_similarity(a, b):
    """Similarity score between two titles."""
    return SequenceMatcher(None, normalize_title(a), normalize_title(b)).ratio()


def token_overlap(a, b):
    """Jaccard-style overlap between normalized title tokens."""
    a_tokens = set(normalize_title(a).split())
    b_tokens = set(normalize_title(b).split())

    if not a_tokens or not b_tokens:
        return 0.0

    return len(a_tokens & b_tokens) / len(a_tokens | b_tokens)


def should_merge_titles(row_a, row_b):
    """Conservative duplicate test for title-based merging."""
    title_a = clean_text(row_a.get("title", ""))
    title_b = clean_text(row_b.get("title", ""))

    if not title_a or not title_b:
        return False

    title_norm_a = normalize_title(title_a)
    title_norm_b = normalize_title(title_b)

    if title_norm_a == title_norm_b:
        return True

    similarity = title_similarity(title_a, title_b)
    overlap = token_overlap(title_a, title_b)

    year_a = row_a.get("year_numeric")
    year_b = row_b.get("year_numeric")
    years_close = (
        pd.isna(year_a)
        or pd.isna(year_b)
        or abs(int(year_a) - int(year_b)) <= 1
    )

    return years_close and similarity >= 0.985 and overlap >= 0.9


def strongest_matching_terms(text, keywords):
    """Return matching keywords for reporting."""
    text = clean_text(text).lower()
    return [keyword for keyword in keywords if keyword.lower() in text]


# ============================================================
# 5. RETRY LOGIC FOR API REQUESTS
# ============================================================

def request_with_retries(url, params, max_retries=MAX_RETRIES, extra_headers=None):
    """
    Send a GET request with retries.

    This handles common temporary errors:
    - connection reset
    - timeout
    - server-side interruption
    - rate-limit-like failures

    extra_headers lets callers add per-request headers, e.g. an
    Authorization bearer token (CORE) or an x-api-key (Semantic Scholar).
    """
    headers = {"User-Agent": f"{PROJECT_NAME}/1.0 mailto:{USER_EMAIL}"}

    if extra_headers:
        headers.update(extra_headers)

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(
                url,
                params=params,
                timeout=REQUEST_TIMEOUT_SECONDS,
                headers=headers
            )

            if getattr(response, "status_code", None) == 404:
                print(f"Request returned 404 and will not be retried: {url}")
                return None

            response.raise_for_status()
            return response

        except requests.exceptions.RequestException as e:
            print(f"Request failed, attempt {attempt}/{max_retries}: {e}")

            if attempt < max_retries:
                time.sleep(RETRY_SLEEP_SECONDS)
            else:
                print("Skipping this request after repeated failures.")
                return None


# ============================================================
# 6. LOAD LOCAL DATABASE EXPORTS
# ============================================================

def read_csv_safely(file_path):
    """Read CSV with several encoding fallbacks."""
    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]

    for enc in encodings:
        try:
            return pd.read_csv(file_path, encoding=enc)
        except Exception:
            continue

    raise ValueError(f"Could not read CSV file: {file_path}")


def load_local_exports(input_dir):
    """
    Load CSV/XLSX files exported from databases such as:
    Scopus, Web of Science, IEEE, ACM, ScienceDirect, Springer, ERIC.
    """
    files = list(input_dir.glob("*.csv")) + list(input_dir.glob("*.xlsx"))

    records = []

    for file in files:
        if is_control_file(file.name):
            continue

        try:
            if file.suffix.lower() == ".csv":
                df = read_csv_safely(file)

            elif file.suffix.lower() == ".xlsx":
                df = pd.read_excel(file)

            else:
                continue

            df["import_source_file"] = file.name
            records.append(df)

        except Exception as e:
            print(f"Could not read {file.name}: {e}")

    if not records:
        return pd.DataFrame()

    return pd.concat(records, ignore_index=True)


def standardize_local_records(df):
    """Map different export column names into a common format."""
    if df.empty:
        return pd.DataFrame()

    df = df.copy()

    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.replace(".", "_")
    )

    if df.columns.duplicated().any():
        deduped = pd.DataFrame(index=df.index)

        for column_name in pd.unique(df.columns):
            matching = df.loc[:, df.columns == column_name]

            if matching.shape[1] == 1:
                deduped[column_name] = matching.iloc[:, 0]
            else:
                # Merge same-meaning columns from different import schemas by
                # taking the first non-empty value row by row.
                merged = matching.iloc[:, 0].copy()

                for index in range(1, matching.shape[1]):
                    candidate = matching.iloc[:, index]
                    merged = merged.where(
                        merged.notna() & (merged.astype(str).str.strip() != ""),
                        candidate
                    )

                deduped[column_name] = merged

        df = deduped

    possible_columns = {
        "title": [
            "title", "document_title", "article_title", "publication_title",
            "item_title"
        ],
        "abstract": [
            "abstract", "description", "abstract_note", "summary"
        ],
        "doi": [
            "doi", "document_doi", "digital_object_identifier"
        ],
        "year": [
            "year", "publication_year", "pub_year", "date"
        ],
        "authors": [
            "authors", "author_names", "creators", "author"
        ],
        "source": [
            "source_title", "journal", "publication", "publication_name",
            "booktitle", "conference_name", "source"
        ],
        "url": [
            "url", "link", "document_url", "publication_url"
        ],
        "keywords": [
            "keywords", "author_keywords", "index_keywords"
        ],
        "database": [
            "database", "source_database"
        ],
        "import_source_file": [
            "import_source_file"
        ]
    }

    def find_col(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None

    out = pd.DataFrame()

    for standard_col, candidates in possible_columns.items():
        col = find_col(candidates)
        out[standard_col] = df[col] if col else ""

    out["record_origin"] = "local_export"
    out["search_label"] = ""
    out["variant_label"] = ""
    out["query_text"] = ""
    out["cited_by_count"] = ""
    out["openalex_id"] = ""
    out["work_type"] = ""
    out["source_type"] = ""
    out["language"] = ""

    if "import_source_file" in out.columns:
        seed_mask = out["import_source_file"].apply(is_seed_reference_file)
        out.loc[seed_mask, "record_origin"] = "seed_reference_local"
        out.loc[seed_mask, "search_label"] = "SEED_reference_capture"
        out.loc[seed_mask, "variant_label"] = "seed_manual_input"
        out.loc[seed_mask, "query_text"] = "seed_reference_csv"

        consensus_mask = out["import_source_file"].apply(is_consensus_export_file)
        out.loc[consensus_mask, "record_origin"] = "consensus_automatic_search_result"
        out.loc[consensus_mask, "search_label"] = "consensus_automatic_search_result"
        out.loc[consensus_mask, "variant_label"] = "consensus_export_csv"
        out.loc[consensus_mask, "query_text"] = "consensus_automatic_search_result"
        if "database" in out.columns:
            blank_database_mask = consensus_mask & (out["database"].astype(str).str.strip() == "")
            out.loc[blank_database_mask, "database"] = "Consensus"

    return out


def build_openalex_record(item, query_label, variant_label, query_text):
    """Convert one OpenAlex work item to the normalized record shape."""
    primary_location = item.get("primary_location") or {}
    source = primary_location.get("source") or {}
    authorships = item.get("authorships") or []
    authors = []

    for authorship in authorships:
        author = authorship.get("author") or {}
        display_name = author.get("display_name", "")

        if display_name:
            authors.append(display_name)

    abstract = reconstruct_openalex_abstract(
        item.get("abstract_inverted_index")
    )

    return {
        "title": item.get("title", ""),
        "abstract": abstract,
        "doi": item.get("doi", ""),
        "year": item.get("publication_year", ""),
        "authors": "; ".join(authors),
        "source": source.get("display_name", ""),
        "url": item.get("id", ""),
        "keywords": "",
        "database": "OpenAlex",
        "import_source_file": "",
        "record_origin": "api_openalex",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": item.get("cited_by_count", ""),
        "openalex_id": item.get("id", ""),
        "work_type": item.get("type", ""),
        "source_type": source.get("type", ""),
        "language": item.get("language", "")
    }


def build_crossref_record(item, query_label, variant_label, query_text):
    """Convert one Crossref work item to the normalized record shape."""
    title = ""
    if item.get("title"):
        title = item["title"][0]

    abstract = clean_text(item.get("abstract", ""))

    authors = []

    for author in item.get("author", []):
        given = author.get("given", "")
        family = author.get("family", "")
        full_name = f"{given} {family}".strip()

        if full_name:
            authors.append(full_name)

    year = ""
    published = (
        item.get("published-print")
        or item.get("published-online")
        or item.get("published")
        or item.get("created")
    )

    if published and "date-parts" in published:
        try:
            year = published["date-parts"][0][0]
        except Exception:
            year = ""

    container_title = ""
    if item.get("container-title"):
        container_title = item["container-title"][0]

    return {
        "title": title,
        "abstract": abstract,
        "doi": item.get("DOI", ""),
        "year": year,
        "authors": "; ".join(authors),
        "source": container_title,
        "url": item.get("URL", ""),
        "keywords": "",
        "database": "Crossref",
        "import_source_file": "",
        "record_origin": "api_crossref",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": item.get("is-referenced-by-count", ""),
        "openalex_id": "",
        "work_type": item.get("type", ""),
        "source_type": "",
        "language": item.get("language", "")
    }


def build_eric_record(item, query_label, variant_label, query_text):
    """Convert one ERIC record to the normalized record shape."""
    eric_id = clean_text(item.get("id", ""))
    source = clean_text(item.get("source", "")) or clean_text(item.get("publisher", ""))
    publication_type = clean_text(item.get("publicationtype", ""))

    return {
        "title": clean_text(item.get("title", "")),
        "abstract": clean_text(item.get("description", "")),
        "doi": clean_text(item.get("doi", "")),
        "year": clean_text(item.get("publicationdateyear", "")),
        "authors": clean_text(item.get("author", "")),
        "source": source,
        "url": clean_text(item.get("url", "")) or (f"https://eric.ed.gov/?id={eric_id}" if eric_id else ""),
        "keywords": clean_text(item.get("subject", "")),
        "database": "ERIC",
        "import_source_file": "",
        "record_origin": "api_eric",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": clean_text(item.get("iescited", "")),
        "openalex_id": "",
        "work_type": publication_type,
        "source_type": "",
        "language": clean_text(item.get("language", "")),
        "eric_id": eric_id,
        "peerreviewed": clean_text(item.get("peerreviewed", "")),
        "educationlevel": clean_text(item.get("educationlevel", "")),
        "publicationtype": publication_type,
        "publisher": clean_text(item.get("publisher", "")),
        "fulltext_available_eric": clean_text(item.get("e_fulltextauth", "")),
        "ies_cited": clean_text(item.get("iescited", "")),
        "sourceid": clean_text(item.get("sourceid", "")),
        "date_modified": clean_text(item.get("e_datemodified", ""))
    }


def merge_seed_with_api(seed_row, api_record, api_origin_label):
    """Prefer seed values but fill missing metadata from API records."""
    merged = seed_row.copy()

    for field in [
        "title", "abstract", "doi", "year", "authors", "source", "url",
        "keywords", "database", "cited_by_count", "openalex_id",
        "work_type", "source_type", "language"
    ]:
        seed_value = clean_text(merged.get(field, ""))
        api_value = clean_text(api_record.get(field, ""))

        if (not seed_value) and api_value:
            merged[field] = api_value

    merged["record_origin"] = api_origin_label
    merged["search_label"] = "SEED_reference_capture"
    merged["variant_label"] = "seed_manual_input"

    if clean_text(merged.get("query_text", "")) in {"", "seed_reference_csv"}:
        merged["query_text"] = clean_text(api_record.get("query_text", "")) or "seed_reference_lookup"

    if clean_text(merged.get("database", "")) == "":
        merged["database"] = clean_text(api_record.get("database", "")) or "SeedReference"

    return merged


def search_openalex_seed_reference(seed_row):
    """Look up a specific seed reference in OpenAlex using DOI or title."""
    doi_norm = normalize_doi(seed_row.get("doi", ""))
    title = clean_text(seed_row.get("title", ""))

    url = "https://api.openalex.org/works"
    params = {
        "mailto": USER_EMAIL,
        "per-page": 10
    }

    if doi_norm:
        params["filter"] = f"doi:{doi_norm}"
        query_text = f"seed_doi:{doi_norm}"
    else:
        params["search"] = title
        query_text = f"seed_title:{title}"

    response = request_with_retries(url, params)

    if response is None:
        return None

    try:
        data = response.json()
        results = data.get("results", [])
    except Exception:
        return None

    title_norm = normalize_title(title)

    for item in results:
        item_doi = normalize_doi(item.get("doi", ""))
        item_title_norm = normalize_title(item.get("title", ""))

        if doi_norm and item_doi == doi_norm:
            return build_openalex_record(
                item,
                query_label="SEED_reference_capture",
                variant_label="seed_openalex_lookup",
                query_text=query_text
            )

        if title_norm and item_title_norm == title_norm:
            return build_openalex_record(
                item,
                query_label="SEED_reference_capture",
                variant_label="seed_openalex_lookup",
                query_text=query_text
            )

    return None


def search_crossref_seed_reference(seed_row):
    """Look up a specific seed reference in Crossref using DOI or title."""
    doi_norm = normalize_doi(seed_row.get("doi", ""))
    title = clean_text(seed_row.get("title", ""))

    query_label = "SEED_reference_capture"
    variant_label = "seed_crossref_lookup"

    if doi_norm:
        url = f"https://api.crossref.org/works/{doi_norm}"
        query_text = f"seed_doi:{doi_norm}"
        response = request_with_retries(url, {})

        if response is None:
            return None

        try:
            item = response.json().get("message", {})
        except Exception:
            return None

        if item:
            return build_crossref_record(
                item,
                query_label=query_label,
                variant_label=variant_label,
                query_text=query_text
            )

        return None

    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": 10,
        "mailto": USER_EMAIL
    }
    response = request_with_retries(url, params)

    if response is None:
        return None

    try:
        items = response.json().get("message", {}).get("items", [])
    except Exception:
        return None

    title_norm = normalize_title(title)

    for item in items:
        item_title = ""
        if item.get("title"):
            item_title = item["title"][0]

        if title_norm and normalize_title(item_title) == title_norm:
            return build_crossref_record(
                item,
                query_label=query_label,
                variant_label=variant_label,
                query_text=f"seed_title:{title}"
            )

    return None


def enrich_seed_references(df):
    """Enrich seed reference rows with metadata from OpenAlex/Crossref."""
    if df.empty or not USE_SEED_REFERENCE_ENRICHMENT:
        return df

    df = df.copy()

    if "import_source_file" not in df.columns:
        return df

    seed_mask = df["import_source_file"].apply(is_seed_reference_file)

    if not seed_mask.any():
        return df

    enriched_rows = []

    for _, row in df.iterrows():
        row_dict = row.to_dict()

        if not is_seed_reference_file(row_dict.get("import_source_file", "")):
            enriched_rows.append(row_dict)
            continue

        openalex_match = search_openalex_seed_reference(row_dict)

        if openalex_match:
            row_dict = merge_seed_with_api(
                row_dict,
                openalex_match,
                api_origin_label="seed_reference_enriched_openalex"
            )
        else:
            crossref_match = search_crossref_seed_reference(row_dict)

            if crossref_match:
                row_dict = merge_seed_with_api(
                    row_dict,
                    crossref_match,
                    api_origin_label="seed_reference_enriched_crossref"
                )
            else:
                row_dict["record_origin"] = "seed_reference_local"
                row_dict["search_label"] = "SEED_reference_capture"
                row_dict["variant_label"] = "seed_manual_input"
                row_dict["query_text"] = "seed_reference_csv"

        enriched_rows.append(row_dict)
        time.sleep(0.3)

    return pd.DataFrame(enriched_rows)


# ============================================================
# 7. OPENALEX SEARCH
# ============================================================

def search_openalex(job, pages=2, per_page=50):
    """Search OpenAlex Works API."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["openalex_query"]

    for page in range(1, pages + 1):
        url = "https://api.openalex.org/works"

        params = {
            "search": query,
            "page": page,
            "per-page": per_page,
            "mailto": USER_EMAIL
        }

        response = request_with_retries(url, params)

        if response is None:
            print(f"OpenAlex skipped: {query_label}, page {page}")
            continue

        try:
            data = response.json()

            for item in data.get("results", []):
                results.append(
                    build_openalex_record(
                        item,
                        query_label=query_label,
                        variant_label=variant_label,
                        query_text=query
                    )
                )

            time.sleep(1)

        except Exception as e:
            print(f"OpenAlex parsing error for {query_label}, page {page}: {e}")

    return pd.DataFrame(results)


# ============================================================
# 8. CROSSREF SEARCH
# ============================================================

def search_crossref(job, rows=50):
    """Search Crossref Works API."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["crossref_query"]

    url = "https://api.crossref.org/works"

    params = {
        "query.bibliographic": query,
        "rows": rows,
        "mailto": USER_EMAIL
    }

    filter_parts = []

    if YEAR_MIN is not None:
        filter_parts.append(f"from-pub-date:{YEAR_MIN}")

    if YEAR_MAX is not None:
        filter_parts.append(f"until-pub-date:{YEAR_MAX}")

    if filter_parts:
        params["filter"] = ",".join(filter_parts)

    response = request_with_retries(url, params)

    if response is None:
        print(f"Crossref skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        data = response.json()
        items = data.get("message", {}).get("items", [])

        for item in items:
            results.append(
                build_crossref_record(
                    item,
                    query_label=query_label,
                    variant_label=variant_label,
                    query_text=query
                )
            )

        time.sleep(1)

    except Exception as e:
        print(f"Crossref parsing error for {query_label}: {e}")

    return pd.DataFrame(results)


def extract_eric_num_found(payload):
    """Return ERIC's total hit count from the API payload."""
    if not isinstance(payload, dict):
        return 0

    response = payload.get("response")

    if isinstance(response, dict):
        return int(response.get("numFound", 0) or 0)

    return int(payload.get("numFound", 0) or 0)


def extract_eric_docs(payload):
    """Return the ERIC document list from either supported payload shape."""
    if not isinstance(payload, dict):
        return []

    response = payload.get("response")

    if isinstance(response, dict):
        return response.get("docs", []) or []

    return payload.get("docs", []) or []


def search_eric(job):
    """Search the ERIC API with paging and a per-query safety cap."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["eric_query"]

    params = {
        "search": query,
        "format": "json",
        "start": 0,
        "rows": ERIC_ROWS_PER_CALL,
        "fields": ",".join(ERIC_FIELDS)
    }

    response = request_with_retries(ERIC_API_URL, params)

    if response is None:
        print(f"ERIC skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        first_payload = response.json()
    except Exception as e:
        print(f"ERIC parsing error for {query_label}: {e}")
        return pd.DataFrame(results)

    num_found = extract_eric_num_found(first_payload)
    all_docs = list(extract_eric_docs(first_payload))

    if num_found == 0:
        return pd.DataFrame(results)

    max_to_download = min(num_found, ERIC_MAX_RECORDS_PER_QUERY)
    start = ERIC_ROWS_PER_CALL

    while start < max_to_download:
        page_params = {
            "search": query,
            "format": "json",
            "start": start,
            "rows": ERIC_ROWS_PER_CALL,
            "fields": ",".join(ERIC_FIELDS)
        }
        page_response = request_with_retries(ERIC_API_URL, page_params)

        if page_response is None:
            break

        try:
            page_payload = page_response.json()
        except Exception as e:
            print(f"ERIC parsing error for {query_label}, start {start}: {e}")
            break

        page_docs = extract_eric_docs(page_payload)

        if not page_docs:
            break

        all_docs.extend(page_docs)
        start += ERIC_ROWS_PER_CALL
        time.sleep(0.4)

    for item in all_docs[:max_to_download]:
        results.append(
            build_eric_record(
                item,
                query_label=query_label,
                variant_label=variant_label,
                query_text=query
            )
        )

    return pd.DataFrame(results)


# ============================================================
# 8b. ADDITIONAL FREE SOURCES
#     Semantic Scholar, arXiv, DBLP, DOAJ, CORE
#
#     All free. Semantic Scholar, arXiv, DBLP and DOAJ need no API key.
#     CORE needs a free key (https://core.ac.uk/services/api) and is
#     skipped automatically if CORE_API_KEY is empty.
#
#     Why these were added: this review's corpus is dominated by LLM and
#     multi-agent work, which is concentrated in IEEE Xplore and the ACM
#     Digital Library — both paywalled and unavailable here. DBLP indexes
#     the bibliographic records of essentially every IEEE/ACM CS venue,
#     and Semantic Scholar + arXiv recover the abstracts and preprints of
#     that same literature. Together they recover the large majority of
#     records the paywalled databases would have returned, at no cost.
#
#     Deliberately NOT added (and why): Google Scholar (no official API;
#     scraping violates its terms and is unreliable); BASE (free, but the
#     API requires IP whitelisting/registration); Lens.org and Dimensions
#     (scholarly API access is limited/paid); Scopus / Web of Science /
#     IEEE Xplore / ACM DL (all require paid subscriptions). You can still
#     drop manual CSV exports from any of these into the slr_exports
#     folder and they will be merged through load_local_exports().
# ============================================================

def build_semantic_scholar_record(item, query_label, variant_label, query_text):
    """Convert one Semantic Scholar paper to the normalized record shape."""
    external_ids = item.get("externalIds") or {}
    authors = [
        clean_text(author.get("name", ""))
        for author in (item.get("authors") or [])
        if author.get("name")
    ]
    arxiv_id = external_ids.get("ArXiv", "")
    doi = external_ids.get("DOI", "")
    url = item.get("url", "") or (
        f"https://arxiv.org/abs/{arxiv_id}" if arxiv_id else ""
    )
    publication_types = item.get("publicationTypes") or []

    if isinstance(publication_types, list):
        work_type = "; ".join(publication_types)
    else:
        work_type = clean_text(publication_types)

    return {
        "title": clean_text(item.get("title", "")),
        "abstract": clean_text(item.get("abstract", "")),
        "doi": doi,
        "year": item.get("year", "") or "",
        "authors": "; ".join(authors),
        "source": clean_text(item.get("venue", "")),
        "url": url,
        "keywords": "",
        "database": "SemanticScholar",
        "import_source_file": "",
        "record_origin": "api_semantic_scholar",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": item.get("citationCount", "") or "",
        "openalex_id": "",
        "work_type": work_type,
        "source_type": "",
        "language": ""
    }


def search_semantic_scholar(job, limit=SEMANTIC_SCHOLAR_LIMIT_PER_QUERY):
    """Search the Semantic Scholar Graph API (free, no key required)."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["semantic_scholar_query"]

    if not query:
        return pd.DataFrame(results)

    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    params = {
        "query": query,
        "limit": min(limit, 100),
        "fields": SEMANTIC_SCHOLAR_FIELDS
    }

    if YEAR_MIN is not None or YEAR_MAX is not None:
        low = YEAR_MIN if YEAR_MIN is not None else ""
        high = YEAR_MAX if YEAR_MAX is not None else ""
        params["year"] = f"{low}-{high}"

    extra_headers = (
        {"x-api-key": SEMANTIC_SCHOLAR_API_KEY}
        if SEMANTIC_SCHOLAR_API_KEY else None
    )

    response = request_with_retries(url, params, extra_headers=extra_headers)

    if response is None:
        print(f"Semantic Scholar skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        data = response.json()

        for item in data.get("data", []) or []:
            results.append(
                build_semantic_scholar_record(
                    item,
                    query_label=query_label,
                    variant_label=variant_label,
                    query_text=query
                )
            )

        # Semantic Scholar rate-limits unauthenticated traffic; be gentle.
        time.sleep(1.5)

    except Exception as e:
        print(f"Semantic Scholar parsing error for {query_label}: {e}")

    return pd.DataFrame(results)


ARXIV_ATOM_NS = {
    "atom": "http://www.w3.org/2005/Atom",
    "arxiv": "http://arxiv.org/schemas/atom"
}


def build_arxiv_record(entry, query_label, variant_label, query_text):
    """Convert one arXiv Atom <entry> element to the normalized record."""

    def text_of(tag):
        node = entry.find(f"atom:{tag}", ARXIV_ATOM_NS)
        return clean_text(node.text) if node is not None and node.text else ""

    published = text_of("published")
    year = published[:4] if len(published) >= 4 else ""

    authors = [
        clean_text(name.text)
        for name in entry.findall("atom:author/atom:name", ARXIV_ATOM_NS)
        if name is not None and name.text
    ]

    doi_node = entry.find("arxiv:doi", ARXIV_ATOM_NS)
    doi = clean_text(doi_node.text) if doi_node is not None and doi_node.text else ""

    primary = entry.find("arxiv:primary_category", ARXIV_ATOM_NS)
    category = primary.get("term", "") if primary is not None else ""

    return {
        "title": text_of("title"),
        "abstract": text_of("summary"),
        "doi": doi,
        "year": year,
        "authors": "; ".join(authors),
        "source": "arXiv",
        "url": text_of("id"),
        "keywords": category,
        "database": "arXiv",
        "import_source_file": "",
        "record_origin": "api_arxiv",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": "",
        "openalex_id": "",
        "work_type": "preprint",
        "source_type": "repository",
        "language": ""
    }


def search_arxiv(job, max_results=ARXIV_MAX_RESULTS_PER_QUERY):
    """Search the arXiv API (free, no key required). Returns Atom XML."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["arxiv_query"]

    if not query:
        return pd.DataFrame(results)

    url = "http://export.arxiv.org/api/query"
    params = {
        "search_query": query,
        "start": 0,
        "max_results": max_results,
        "sortBy": "relevance",
        "sortOrder": "descending"
    }

    response = request_with_retries(url, params)

    if response is None:
        print(f"arXiv skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        root = ET.fromstring(response.text)

        for entry in root.findall("atom:entry", ARXIV_ATOM_NS):
            results.append(
                build_arxiv_record(
                    entry,
                    query_label=query_label,
                    variant_label=variant_label,
                    query_text=query
                )
            )

        # arXiv asks callers to wait ~3 seconds between requests.
        time.sleep(3)

    except Exception as e:
        print(f"arXiv parsing error for {query_label}: {e}")

    return pd.DataFrame(results)


def _dblp_author_names(info):
    """Extract author names from a DBLP info block (handles all shapes)."""
    authors_field = info.get("authors") or {}
    author = authors_field.get("author") if isinstance(authors_field, dict) else None
    names = []

    if isinstance(author, list):
        for entry in author:
            if isinstance(entry, dict):
                names.append(clean_text(entry.get("text", "")))
            else:
                names.append(clean_text(entry))
    elif isinstance(author, dict):
        names.append(clean_text(author.get("text", "")))
    elif isinstance(author, str):
        names.append(clean_text(author))

    return [name for name in names if name]


def build_dblp_record(hit, query_label, variant_label, query_text):
    """Convert one DBLP hit to the normalized record shape (no abstract)."""
    info = hit.get("info") or {}

    return {
        "title": clean_text(info.get("title", "")),
        "abstract": "",  # DBLP does not provide abstracts; dedup fills from S2/OpenAlex
        "doi": clean_text(info.get("doi", "")),
        "year": clean_text(info.get("year", "")),
        "authors": "; ".join(_dblp_author_names(info)),
        "source": clean_text(info.get("venue", "")),
        "url": clean_text(info.get("ee", "")) or clean_text(info.get("url", "")),
        "keywords": "",
        "database": "DBLP",
        "import_source_file": "",
        "record_origin": "api_dblp",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": "",
        "openalex_id": "",
        "work_type": clean_text(info.get("type", "")),
        "source_type": "",
        "language": ""
    }


def search_dblp(job, hits=DBLP_HITS_PER_QUERY):
    """Search the DBLP publication API (free, no key required)."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["dblp_query"]

    if not query:
        return pd.DataFrame(results)

    url = "https://dblp.org/search/publ/api"
    params = {"q": query, "format": "json", "h": hits, "f": 0}

    response = request_with_retries(url, params)

    if response is None:
        print(f"DBLP skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        data = response.json()
        hit_list = (
            ((data.get("result") or {}).get("hits") or {}).get("hit")
        ) or []

        if isinstance(hit_list, dict):
            hit_list = [hit_list]

        for hit in hit_list:
            results.append(
                build_dblp_record(
                    hit,
                    query_label=query_label,
                    variant_label=variant_label,
                    query_text=query
                )
            )

        time.sleep(1)

    except Exception as e:
        print(f"DBLP parsing error for {query_label}: {e}")

    return pd.DataFrame(results)


def build_doaj_record(item, query_label, variant_label, query_text):
    """Convert one DOAJ article to the normalized record shape."""
    bibjson = item.get("bibjson") or {}

    authors = [
        clean_text(author.get("name", ""))
        for author in (bibjson.get("author") or [])
        if author.get("name")
    ]

    doi = ""
    for identifier in bibjson.get("identifier", []) or []:
        if str(identifier.get("type", "")).lower() == "doi":
            doi = clean_text(identifier.get("id", ""))
            break

    url = ""
    for link in bibjson.get("link", []) or []:
        if str(link.get("type", "")).lower() in {"fulltext", "homepage"}:
            url = clean_text(link.get("url", ""))
            break

    journal_title = (bibjson.get("journal") or {}).get("title", "")

    return {
        "title": clean_text(bibjson.get("title", "")),
        "abstract": clean_text(bibjson.get("abstract", "")),
        "doi": doi,
        "year": clean_text(bibjson.get("year", "")),
        "authors": "; ".join(authors),
        "source": clean_text(journal_title),
        "url": url,
        "keywords": "; ".join(
            clean_text(keyword) for keyword in (bibjson.get("keywords") or [])
        ),
        "database": "DOAJ",
        "import_source_file": "",
        "record_origin": "api_doaj",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": "",
        "openalex_id": "",
        "work_type": "journal-article",
        "source_type": "open_access_journal",
        "language": "; ".join(
            clean_text(language) for language in (bibjson.get("language") or [])
        )
    }


def search_doaj(job, page_size=DOAJ_PAGE_SIZE, max_pages=DOAJ_MAX_PAGES):
    """Search the DOAJ articles API (free, no key required)."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["doaj_query"]

    if not query:
        return pd.DataFrame(results)

    base = "https://doaj.org/api/v2/search/articles/"

    for page in range(1, max_pages + 1):
        # DOAJ takes the query in the URL path, so it must be encoded there.
        url = base + quote(query, safe="")
        params = {"pageSize": page_size, "page": page}

        response = request_with_retries(url, params)

        if response is None:
            print(f"DOAJ skipped: {query_label}, page {page}")
            break

        try:
            data = response.json()
            page_results = data.get("results", []) or []

            if not page_results:
                break

            for item in page_results:
                results.append(
                    build_doaj_record(
                        item,
                        query_label=query_label,
                        variant_label=variant_label,
                        query_text=query
                    )
                )

            if len(page_results) < page_size:
                break

            time.sleep(1)

        except Exception as e:
            print(f"DOAJ parsing error for {query_label}, page {page}: {e}")
            break

    return pd.DataFrame(results)


def build_core_record(item, query_label, variant_label, query_text):
    """Convert one CORE work to the normalized record shape."""
    authors = []
    for author in item.get("authors", []) or []:
        if isinstance(author, dict):
            authors.append(clean_text(author.get("name", "")))
        else:
            authors.append(clean_text(author))
    authors = [name for name in authors if name]

    url = clean_text(item.get("downloadUrl", "") or "")
    if not url:
        fulltext_urls = item.get("sourceFulltextUrls") or []
        if fulltext_urls:
            url = clean_text(fulltext_urls[0])

    return {
        "title": clean_text(item.get("title", "")),
        "abstract": clean_text(item.get("abstract", "")),
        "doi": clean_text(item.get("doi", "") or ""),
        "year": item.get("yearPublished", "") or "",
        "authors": "; ".join(authors),
        "source": clean_text(item.get("publisher", "")),
        "url": url,
        "keywords": "",
        "database": "CORE",
        "import_source_file": "",
        "record_origin": "api_core",
        "search_label": query_label,
        "variant_label": variant_label,
        "query_text": query_text,
        "cited_by_count": "",
        "openalex_id": "",
        "work_type": clean_text(item.get("documentType", "")),
        "source_type": "open_access_aggregator",
        "language": ""
    }


def search_core(job, limit=CORE_LIMIT_PER_QUERY):
    """Search the CORE v3 API. Requires a free CORE_API_KEY; skipped if unset."""
    results = []
    query_label = job["search_label"]
    variant_label = job["variant_label"]
    query = job["core_query"]

    if not CORE_API_KEY or not query:
        return pd.DataFrame(results)

    url = "https://api.core.ac.uk/v3/search/works"
    params = {"q": query, "limit": min(limit, 100)}
    headers = {"Authorization": f"Bearer {CORE_API_KEY}"}

    response = request_with_retries(url, params, extra_headers=headers)

    if response is None:
        print(f"CORE skipped: {query_label}")
        return pd.DataFrame(results)

    try:
        data = response.json()

        for item in data.get("results", []) or []:
            results.append(
                build_core_record(
                    item,
                    query_label=query_label,
                    variant_label=variant_label,
                    query_text=query
                )
            )

        time.sleep(1)

    except Exception as e:
        print(f"CORE parsing error for {query_label}: {e}")

    return pd.DataFrame(results)


# ============================================================
# 9. NORMALIZATION
# ============================================================

def normalize_records(df):
    """Normalize all records into a stable table."""
    required_cols = [
        "title", "abstract", "doi", "year", "authors", "source", "url",
        "keywords", "database", "import_source_file", "record_origin",
        "search_label", "variant_label", "query_text", "cited_by_count",
        "openalex_id", "work_type", "source_type", "language", "eric_id",
        "peerreviewed", "educationlevel", "publicationtype", "publisher",
        "fulltext_available_eric", "ies_cited", "sourceid", "date_modified"
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    for col in required_cols:
        df[col] = df[col].apply(clean_text)

    df["doi_normalized"] = df["doi"].apply(normalize_doi)
    df["title_normalized"] = df["title"].apply(normalize_title)

    df["record_id"] = df.apply(
        lambda row: make_record_id(row["title"], row["doi"], row["eric_id"]),
        axis=1
    )

    df["year_numeric"] = pd.to_numeric(df["year"], errors="coerce")

    return df


# ============================================================
# 10. DEDUPLICATION
# ============================================================

def deduplicate_records(df):
    """
    Deduplicate records.

    Priority:
    1. DOI-based deduplication.
    2. Exact normalized-title deduplication across all sources.
    3. Conservative near-duplicate title merging.
    """
    if df.empty:
        return df, pd.DataFrame()

    df = df.copy()

    df["abstract_length"] = df["abstract"].apply(lambda x: len(clean_text(x)))
    df["has_abstract"] = df["abstract_length"] > 0
    df["cited_by_numeric"] = pd.to_numeric(df["cited_by_count"], errors="coerce").fillna(0)
    df["peer_review_like"] = df["title"].str.lower().str.contains(
        "peer review report", na=False
    )
    df["repository_like"] = (
        df["source"].str.lower().str.contains("arxiv|zenodo|osf", na=False)
        | df["url"].str.lower().str.contains("arxiv|zenodo|osf", na=False)
    )
    df["has_doi"] = df["doi_normalized"] != ""
    df["eric_id_normalized"] = df["eric_id"].apply(lambda value: clean_text(value).lower())
    df["has_eric_id"] = df["eric_id_normalized"] != ""

    df = df.sort_values(
        by=[
            "has_abstract",
            "abstract_length",
            "has_doi",
            "has_eric_id",
            "peer_review_like",
            "repository_like",
            "cited_by_numeric",
            "year_numeric"
        ],
        ascending=[False, False, False, False, True, True, False, False]
    ).copy()

    duplicate_rows = []

    def remove_group_duplicates(frame, key_column, reason):
        if frame.empty:
            return frame.copy()

        kept_rows = []
        kept_record_ids = {}

        for _, row in frame.iterrows():
            key = clean_text(row[key_column])

            if not key:
                kept_rows.append(row.to_dict())
                continue

            if key not in kept_record_ids:
                kept_record_ids[key] = row["record_id"]
                kept_rows.append(row.to_dict())
                continue

            dropped = row.to_dict()
            dropped["duplicate_reason"] = reason
            dropped["duplicate_of_record_id"] = kept_record_ids[key]
            duplicate_rows.append(dropped)

        return pd.DataFrame(kept_rows, columns=frame.columns)

    deduped = remove_group_duplicates(
        df,
        key_column="eric_id_normalized",
        reason="same_eric_id"
    )

    deduped = remove_group_duplicates(
        deduped,
        key_column="doi_normalized",
        reason="same_doi"
    )

    deduped = remove_group_duplicates(
        deduped,
        key_column="title_normalized",
        reason="same_title"
    )

    def candidate_bucket_keys(row):
        title_key = clean_text(row.get("title_normalized", ""))

        if not title_key:
            return []

        tokens = title_key.split()
        first_token = tokens[0] if tokens else ""
        first_two = " ".join(tokens[:2]) if tokens else ""
        prefix = title_key[:18]
        length_bucket = len(title_key) // 12

        return [
            ("prefix", prefix),
            ("first_two", first_two),
            ("first_token", first_token, length_bucket)
        ]

    kept_rows = []
    bucket_map = {}

    for _, row in deduped.iterrows():
        duplicate_of = None

        candidate_indexes = []
        seen_indexes = set()

        for bucket_key in candidate_bucket_keys(row):
            for kept_index in bucket_map.get(bucket_key, []):
                if kept_index not in seen_indexes:
                    seen_indexes.add(kept_index)
                    candidate_indexes.append(kept_index)

        for kept_index in candidate_indexes:
            kept = kept_rows[kept_index]
            if should_merge_titles(row, kept):
                duplicate_of = kept["record_id"]
                break

        if duplicate_of:
            dropped = row.to_dict()
            dropped["duplicate_reason"] = "near_duplicate_title"
            dropped["duplicate_of_record_id"] = duplicate_of
            duplicate_rows.append(dropped)
        else:
            kept_row = row.to_dict()
            kept_rows.append(kept_row)
            kept_index = len(kept_rows) - 1

            for bucket_key in candidate_bucket_keys(kept_row):
                bucket_map.setdefault(bucket_key, []).append(kept_index)

    deduped = pd.DataFrame(kept_rows, columns=deduped.columns)
    duplicate_log = pd.DataFrame(duplicate_rows)

    return deduped, duplicate_log


# ============================================================
# 11. RQ TAGGING AND SCREENING SUPPORT
# ============================================================

def tag_rqs(df):
    """Tag each record by research question and add screening columns."""
    df = df.copy()

    df["combined_text"] = (
        df["title"].fillna("") + " " +
        df["abstract"].fillna("") + " " +
        df["keywords"].fillna("")
    ).str.lower()

    for rq, keywords in RQ_KEYWORDS.items():
        df[rq] = df["combined_text"].apply(
            lambda text: text_contains_any(text, keywords)
        )

    df["relevance_score"] = df.apply(
        lambda row: calculate_relevance_score(row["title"], row["abstract"]),
        axis=1
    )

    rq_cols = list(RQ_KEYWORDS.keys())
    df["matched_rq_count"] = df[rq_cols].sum(axis=1)

    df["math_signal_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, MATH_SIGNAL_TERMS))
    )
    df["education_signal_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, EDUCATION_SIGNAL_TERMS))
    )
    df["ai_signal_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, AI_SIGNAL_TERMS))
    )
    df["off_scope_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, OFF_SCOPE_DOMAIN_TERMS))
    )
    df["non_scholarly_terms"] = df["title"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, NON_SCHOLARLY_TITLE_TERMS))
    )
    df["generic_adaptation_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, GENERIC_FINE_TUNING_TERMS))
    )
    df["generic_verification_terms"] = df["combined_text"].apply(
        lambda text: "; ".join(strongest_matching_terms(text, GENERIC_VERIFICATION_TERMS))
    )

    df["has_math_signal"] = df["math_signal_terms"] != ""
    df["has_education_signal"] = df["education_signal_terms"] != ""
    df["has_ai_signal"] = df["ai_signal_terms"] != ""
    df["off_scope_domain_flag"] = df["off_scope_terms"] != ""
    df["non_scholarly_flag"] = df["non_scholarly_terms"] != ""
    df["generic_adaptation_flag"] = df["generic_adaptation_terms"] != ""
    df["generic_verification_flag"] = df["generic_verification_terms"] != ""

    def suggested_decision(row):
        title_abs = f"{row['title']} {row['abstract']} {row['keywords']}".lower()
        math_related = row["has_math_signal"]
        education_related = row["has_education_signal"]
        ai_related = row["has_ai_signal"]
        reasoning_related = any(
            term in title_abs for term in [
                "mathematical reasoning",
                "math problem solving",
                "step-by-step reasoning",
                "solution generation"
            ]
        )

        exclusion_reasons = []

        if row["non_scholarly_flag"]:
            exclusion_reasons.append("non-scholarly/update item")

        if row["off_scope_domain_flag"] and not education_related:
            exclusion_reasons.append("off-scope application domain")

        if row["generic_adaptation_flag"] and not (math_related and education_related):
            exclusion_reasons.append("generic model adaptation paper")

        if row["generic_verification_flag"] and not (math_related and (education_related or reasoning_related)):
            exclusion_reasons.append("generic verification paper")

        

        if exclusion_reasons:
            return (
                "Likely exclude - check manually",
                "; ".join(exclusion_reasons)
            )

        if math_related and ai_related and (education_related or reasoning_related) and row["matched_rq_count"] >= 1:
            return (
                "Potentially include - check manually",
                "strong match across math, AI, and tutoring/reasoning"
            )

        if math_related and (education_related or reasoning_related) and (ai_related or row["relevance_score"] >= 3):
            return (
                "Unclear - check manually",
                "partial match that needs manual scope confirmation"
            )

        if math_related and education_related:
            return (
                "Unclear - check manually",
                "education and mathematics present, but AI contribution is unclear"
            )

        return (
            "Likely exclude - check manually",
            "weak concept overlap with SLR scope"
        )

    decisions = df.apply(suggested_decision, axis=1, result_type="expand")
    df["auto_screening_suggestion"] = decisions[0]
    df["auto_exclusion_reason"] = decisions[1]

    # Manual screening columns.
    df["title_abstract_decision"] = ""
    df["full_text_decision"] = ""
    df["exclusion_reason"] = ""
    df["reviewer_notes"] = ""

    # Extraction columns.
    df["math_domain"] = ""
    df["system_type"] = ""
    df["architecture_summary"] = ""
    df["multi_agent_roles"] = ""
    df["personalization_strategy"] = ""
    df["verification_strategy"] = ""
    df["evaluation_method"] = ""
    df["reported_limitations"] = ""
    df["relevance_to_mathtutorai"] = ""

    # Quality appraisal columns.
    df["Q1_architecture_clear_0_2"] = ""
    df["Q2_math_domain_clear_0_2"] = ""
    df["Q3_personalization_clear_0_2"] = ""
    df["Q4_correctness_addressed_0_2"] = ""
    df["Q5_evaluation_credible_0_2"] = ""
    df["Q6_limitations_clear_0_2"] = ""
    df["quality_total_0_12"] = ""

    return df


# ============================================================
# 12. PRISMA COUNTS
# ============================================================

def create_prisma_counts(raw_count, year_filtered_count, deduped_count, screening_df):
    """Create PRISMA-style counts with both current and manual-progress fields."""
    duplicate_count = year_filtered_count - deduped_count
    manual_title = screening_df["title_abstract_decision"].fillna("").str.lower()
    manual_full_text = screening_df["full_text_decision"].fillna("").str.lower()
    auto_suggestion = screening_df["auto_screening_suggestion"].fillna("")

    return pd.DataFrame([{
        "records_identified_total": raw_count,
        "records_after_year_filter": year_filtered_count,
        "duplicate_records_removed": duplicate_count,
        "records_after_deduplication": deduped_count,
        "records_screened_title_abstract": deduped_count,
        "records_excluded_title_abstract": int(manual_title.str.contains("exclude").sum()),
        "reports_sought_for_retrieval": int(manual_title.str.contains("include").sum()),
        "reports_not_retrieved": int(manual_full_text.str.contains("not retrieved").sum()),
        "reports_assessed_full_text": int((manual_full_text != "").sum()),
        "reports_excluded_full_text": int(manual_full_text.str.contains("exclude").sum()),
        "studies_included_final_synthesis": int(manual_full_text.str.contains("include").sum()),
        "auto_potentially_include": int((auto_suggestion == "Potentially include - check manually").sum()),
        "auto_unclear": int((auto_suggestion == "Unclear - check manually").sum()),
        "auto_likely_exclude": int((auto_suggestion == "Likely exclude - check manually").sum())
    }])


# ============================================================
# 13. EXPORT RESULTS
# ============================================================

def export_results(raw_df, screening_df, duplicate_log, prisma_counts, search_audit_df):
    """Export all outputs to CSV and Excel."""
    suggestion_priority = {
        "Potentially include - check manually": 0,
        "Unclear - check manually": 1,
        "Likely exclude - check manually": 2
    }

    screening_df = screening_df.copy()
    screening_df["suggestion_priority"] = screening_df["auto_screening_suggestion"].map(
        suggestion_priority
    ).fillna(9)
    screening_df = screening_df.sort_values(
        by=[
            "suggestion_priority",
            "relevance_score",
            "matched_rq_count",
            "year_numeric"
        ],
        ascending=[True, False, False, False]
    ).drop(columns=["suggestion_priority"])

    raw_df.to_csv(RAW_RECORDS_CSV, index=False, encoding="utf-8-sig")
    screening_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    prisma_counts.to_csv(PRISMA_COUNTS_CSV, index=False, encoding="utf-8-sig")
    search_audit_df.to_csv(SEARCH_AUDIT_CSV, index=False, encoding="utf-8-sig")

    if duplicate_log is not None and not duplicate_log.empty:
        duplicate_log.to_csv(DUPLICATES_CSV, index=False, encoding="utf-8-sig")

    rq_summary = []

    for rq in RQ_KEYWORDS.keys():
        rq_summary.append({
            "research_question_tag": rq,
            "records_tagged": int(screening_df[rq].sum())
        })

    rq_summary = pd.DataFrame(rq_summary)
    screening_summary = (
        screening_df
        .groupby("auto_screening_suggestion", dropna=False)
        .size()
        .reset_index(name="record_count")
    )
    database_summary = (
        screening_df
        .groupby(["database", "work_type"], dropna=False)
        .size()
        .reset_index(name="record_count")
        .sort_values(["database", "record_count"], ascending=[True, False])
    )

    criteria = pd.DataFrame({
        "Inclusion criteria": [
            "Peer-reviewed or academically credible source",
            "Focuses on AI, ITS, adaptive learning, LLMs, CAS, or agent-based systems",
            "Related to mathematics learning, tutoring, reasoning, or problem solving",
            "Provides evidence for architecture, personalization, verification, evaluation, or limitations",
            "Published within the selected date range unless foundational"
        ],
        "Exclusion criteria": [
            "Not related to mathematics education or mathematical reasoning",
            "No AI, tutoring, adaptive learning, LLM, CAS, or agent-based component",
            "Pure opinion article with no technical or pedagogical evidence",
            "Duplicate record",
            "Unavailable full text, if full-text analysis is required"
        ]
    })

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        screening_df.to_excel(writer, sheet_name="screening_table", index=False)
        prisma_counts.to_excel(writer, sheet_name="prisma_counts", index=False)
        search_audit_df.to_excel(writer, sheet_name="search_audit", index=False)
        rq_summary.to_excel(writer, sheet_name="rq_summary", index=False)
        screening_summary.to_excel(writer, sheet_name="screening_summary", index=False)
        database_summary.to_excel(writer, sheet_name="database_summary", index=False)
        criteria.to_excel(writer, sheet_name="criteria", index=False)

        if duplicate_log is not None and not duplicate_log.empty:
            duplicate_log.to_excel(writer, sheet_name="duplicate_log", index=False)

    print("\nExport completed.")
    print(f"Excel file: {OUTPUT_EXCEL}")
    print(f"Screening CSV: {OUTPUT_CSV}")
    print(f"Raw records CSV: {RAW_RECORDS_CSV}")
    print(f"PRISMA counts: {PRISMA_COUNTS_CSV}")
    print(f"Search audit: {SEARCH_AUDIT_CSV}")

    if duplicate_log is not None and not duplicate_log.empty:
        print(f"Duplicate log: {DUPLICATES_CSV}")


# ============================================================
# 14. MAIN PIPELINE
# ============================================================

def main():
    print("\nStarting SLR pipeline...")
    print(f"Input folder: {INPUT_DIR.resolve()}")
    print(f"Output folder: {OUTPUT_DIR.resolve()}")

    all_records = []
    search_jobs = list(iter_search_jobs())
    search_audit_df = search_jobs_to_df()

    # --------------------------------------------------------
    # Load local database exports
    # --------------------------------------------------------
    print("\nLoading local CSV/XLSX exports...")

    local_raw = load_local_exports(INPUT_DIR)

    if not local_raw.empty:
        local_standard = standardize_local_records(local_raw)
        local_standard = enrich_seed_references(local_standard)
        all_records.append(local_standard)
        print(f"Local records loaded: {len(local_standard)}")

    else:
        print("No local export files found.")
        print("You can place Scopus/Web of Science/IEEE/ACM CSV or XLSX exports in the slr_exports folder.")

    # --------------------------------------------------------
    # Search OpenAlex
    # --------------------------------------------------------
    if USE_OPENALEX:
        print("\nSearching OpenAlex...")

        for job in tqdm(search_jobs):
            df_openalex = search_openalex(
                job=job,
                pages=OPENALEX_PAGES_PER_QUERY,
                per_page=OPENALEX_PER_PAGE
            )

            if not df_openalex.empty:
                all_records.append(df_openalex)

    # --------------------------------------------------------
    # Search Crossref
    # --------------------------------------------------------
    if USE_CROSSREF:
        print("\nSearching Crossref...")

        for job in tqdm(search_jobs):
            df_crossref = search_crossref(
                job=job,
                rows=CROSSREF_ROWS_PER_QUERY
            )

            if not df_crossref.empty:
                all_records.append(df_crossref)

    # --------------------------------------------------------
    # Search ERIC
    # --------------------------------------------------------
    if USE_ERIC:
        print("\nSearching ERIC...")

        for job in tqdm(search_jobs):
            df_eric = search_eric(job)

            if not df_eric.empty:
                all_records.append(df_eric)

    # --------------------------------------------------------
    # Search Semantic Scholar (free, no key required)
    # --------------------------------------------------------
    if USE_SEMANTIC_SCHOLAR:
        print("\nSearching Semantic Scholar...")

        for job in tqdm(search_jobs):
            df_s2 = search_semantic_scholar(job)

            if not df_s2.empty:
                all_records.append(df_s2)

    # --------------------------------------------------------
    # Search arXiv (free, no key required)
    # --------------------------------------------------------
    if USE_ARXIV:
        print("\nSearching arXiv...")

        for job in tqdm(search_jobs):
            df_arxiv = search_arxiv(job)

            if not df_arxiv.empty:
                all_records.append(df_arxiv)

    # --------------------------------------------------------
    # Search DBLP (free, no key; recovers IEEE/ACM CS metadata)
    # --------------------------------------------------------
    if USE_DBLP:
        print("\nSearching DBLP...")

        for job in tqdm(search_jobs):
            df_dblp = search_dblp(job)

            if not df_dblp.empty:
                all_records.append(df_dblp)

    # --------------------------------------------------------
    # Search DOAJ (free, no key; open-access journals)
    # --------------------------------------------------------
    if USE_DOAJ:
        print("\nSearching DOAJ...")

        for job in tqdm(search_jobs):
            df_doaj = search_doaj(job)

            if not df_doaj.empty:
                all_records.append(df_doaj)

    # --------------------------------------------------------
    # Search CORE (free WITH a free API key; skipped if unset)
    # --------------------------------------------------------
    if USE_CORE and CORE_API_KEY:
        print("\nSearching CORE...")

        for job in tqdm(search_jobs):
            df_core = search_core(job)

            if not df_core.empty:
                all_records.append(df_core)

    elif USE_CORE and not CORE_API_KEY:
        print("\nCORE enabled but CORE_API_KEY is not set - skipping CORE.")
        print("Get a free key at https://core.ac.uk/services/api and set the")
        print("CORE_API_KEY environment variable to include CORE in the search.")

    # --------------------------------------------------------
    # Stop if no records
    # --------------------------------------------------------
    if not all_records:
        print("\nNo records found.")
        print("Add export files to slr_exports or keep API search enabled.")
        return

    # --------------------------------------------------------
    # Merge raw records
    # --------------------------------------------------------
    raw_df = pd.concat(all_records, ignore_index=True)
    raw_count = len(raw_df)

    print(f"\nRaw records collected: {raw_count}")

    # Save raw records before cleaning.
    raw_df.to_csv(RAW_RECORDS_CSV, index=False, encoding="utf-8-sig")

    # --------------------------------------------------------
    # Normalize records
    # --------------------------------------------------------
    normalized = normalize_records(raw_df)

    # --------------------------------------------------------
    # Year filtering
    # --------------------------------------------------------
    before_year_filter = len(normalized)

    if YEAR_MIN is not None:
        normalized = normalized[
            normalized["year_numeric"].isna()
            | (normalized["year_numeric"] >= YEAR_MIN)
        ]

    if YEAR_MAX is not None:
        normalized = normalized[
            normalized["year_numeric"].isna()
            | (normalized["year_numeric"] <= YEAR_MAX)
        ]

    year_filtered_count = len(normalized)

    print(f"Records after year filter: {year_filtered_count}")
    print(f"Year-filter removed: {before_year_filter - year_filtered_count}")

    # --------------------------------------------------------
    # Deduplication
    # --------------------------------------------------------
    deduped, duplicate_log = deduplicate_records(normalized)

    deduped_count = len(deduped)
    duplicate_count = year_filtered_count - deduped_count

    print(f"Records after deduplication: {deduped_count}")
    print(f"Duplicates removed: {duplicate_count}")

    # --------------------------------------------------------
    # RQ tagging and screening support
    # --------------------------------------------------------
    screening_df = tag_rqs(deduped)

    raw_query_hits = (
        raw_df[
            raw_df["search_label"].fillna("").astype(str).str.strip() != ""
        ]
        .groupby(["search_label", "variant_label"], dropna=False)
        .size()
        .reset_index(name="raw_hits")
    )
    # Per-database hit columns, generated dynamically so the audit adapts
    # to whichever sources actually ran (OpenAlex, Crossref, ERIC,
    # Semantic Scholar, arXiv, DBLP, DOAJ, CORE, local imports, ...).
    def _hit_column_name(db_name):
        slug = re.sub(r"[^a-z0-9]+", "_", str(db_name).lower()).strip("_")
        return f"{slug}_hits" if slug else "unknown_hits"

    databases_present = [
        db for db in raw_df["database"].dropna().astype(str).unique()
        if db.strip()
    ]
    source_query_hits = (
        raw_df[
            raw_df["database"].astype(str).str.strip().isin(databases_present)
            & (raw_df["search_label"].fillna("").astype(str).str.strip() != "")
        ]
        .groupby(["search_label", "variant_label", "database"], dropna=False)
        .size()
        .reset_index(name="source_hits")
        .pivot_table(
            index=["search_label", "variant_label"],
            columns="database",
            values="source_hits",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )
    db_rename_map = {
        db: _hit_column_name(db)
        for db in databases_present
        if db in source_query_hits.columns
    }
    source_query_hits = source_query_hits.rename(columns=db_rename_map)
    db_hit_columns = list(db_rename_map.values())
    source_query_hits = source_query_hits.reindex(
        columns=["search_label", "variant_label"] + db_hit_columns,
        fill_value=0
    )
    screened_query_hits = (
        screening_df
        .groupby(["search_label", "variant_label"], dropna=False)
        .size()
        .reset_index(name="deduped_hits")
    )
    likely_exclude_hits = (
        screening_df[screening_df["auto_screening_suggestion"] == "Likely exclude - check manually"]
        .groupby(["search_label", "variant_label"], dropna=False)
        .size()
        .reset_index(name="auto_likely_exclude_hits")
    )

    search_audit_df = (
        search_audit_df
        .merge(raw_query_hits, on=["search_label", "variant_label"], how="left")
        .merge(source_query_hits, on=["search_label", "variant_label"], how="left")
        .merge(screened_query_hits, on=["search_label", "variant_label"], how="left")
        .merge(likely_exclude_hits, on=["search_label", "variant_label"], how="left")
    )
    audit_count_columns = (
        ["raw_hits"] + db_hit_columns + ["deduped_hits", "auto_likely_exclude_hits"]
    )
    for col in audit_count_columns:
        if col not in search_audit_df.columns:
            search_audit_df[col] = 0
    search_audit_df[audit_count_columns] = (
        search_audit_df[audit_count_columns].fillna(0).astype(int)
    )

    # --------------------------------------------------------
    # PRISMA counts
    # --------------------------------------------------------
    prisma_counts = create_prisma_counts(
        raw_count=raw_count,
        year_filtered_count=year_filtered_count,
        deduped_count=deduped_count,
        screening_df=screening_df
    )

    # --------------------------------------------------------
    # Export
    # --------------------------------------------------------
    export_results(
        raw_df=raw_df,
        screening_df=screening_df,
        duplicate_log=duplicate_log,
        prisma_counts=prisma_counts,
        search_audit_df=search_audit_df
    )

    print("\nNext step:")
    print("Open the Excel file and manually complete these columns:")
    print("1. title_abstract_decision")
    print("2. full_text_decision")
    print("3. exclusion_reason")
    print("4. reviewer_notes")
    print("5. quality appraisal columns")
    print("6. extraction columns for RQ1–RQ6")

    print("\nMain output file:")
    print(OUTPUT_EXCEL.resolve())


if __name__ == "__main__":
    main()
