# MathTutorAI SLR — Audit Package

This repository contains the **audit and reproducibility package** for the systematic literature review (SLR) conducted as part of the MathTutorAI project.

The package provides the complete screening and analysis evidence supporting the SLR, together with the Python pipeline used to automate the literature retrieval, metadata normalization, deduplication, research-question tagging, screening support, and PRISMA-related outputs.

## Files

### `MathTutorAI_SLR_audit_package.xlsx`

This workbook is the **main audit dataset** of the review. It contains the complete post-deduplication screening and analysis information for the SLR.

The workbook includes:

* **`screening_decisions`** — the 2,670 deduplicated records, including bibliographic metadata, search provenance, automated relevance signals, RQ1–RQ6 tags, PRISMA stage, and final screening decisions.
* **`included_studies`** — the final corpus of **73 included studies**, with extracted characteristics, research-question coverage, and quality-appraisal scores.
* **`prisma_counts`** — PRISMA screening and selection counts used to reproduce the study-selection flow.
* **`search_audit`** — the search strategies, Boolean queries, database sources, and retrieval counts.
* **`duplicate_log`** — records identified as duplicates during cross-database deduplication.
* **`database_summary`** — summary of retrieved records by database and publication type.
* **`rq_signal_summary`** — distribution of records associated with RQ1–RQ6.
* **`screening_summary`** — automated screening suggestions.
* **`criteria`** — inclusion and exclusion criteria applied during screening.
* **`quality_threshold_sensitivity`** — sensitivity analysis for different quality-score thresholds.
* **`rq_cocoverage_matrix`** — research-question co-coverage matrix derived from the included studies.
* **`pillar_coverage`** — cross-pillar coverage of the final corpus.
* **`quality_rubric`** — quality-assessment rubric, score distribution, and scoring provenance.

The workbook also documents the provenance of the PRISMA stages and the relationship between the different screening and analysis datasets.

### `slr_pipeline.py`

This Python script provides the **reproducible SLR processing pipeline**. It:

1. Loads bibliographic records from CSV/XLSX exports.
2. Retrieves literature from **OpenAlex, Crossref, ERIC, Semantic Scholar, and arXiv**.
3. Normalizes bibliographic metadata.
4. Deduplicates records using DOI and title similarity.
5. Applies the configured search strategies and RQ1–RQ6 keyword tagging.
6. Generates screening-support fields and PRISMA-style counts.
7. Produces the screening tables, CSV outputs, and audit information.

The pipeline is intended to **support and document the SLR process, not replace manual screening**. Final inclusion/exclusion decisions remain subject to human review.

## Search Sources

The audit package covers the five primary sources reported in the review:

* OpenAlex
* Crossref
* ERIC
* Semantic Scholar
* arXiv

A supplementary Consensus export is also incorporated into the audit package.

The pipeline additionally supports optional connectors for DBLP, DOAJ, and CORE, which are disabled by default so that the default configuration remains aligned with the reported search strategy.

## Reproducibility

The package is intended to allow reviewers and readers to:

* inspect the complete screening evidence;
* trace included studies back to their bibliographic records;
* examine the search strategies and database provenance;
* verify duplicate-removal decisions;
* reproduce RQ tagging and screening-support logic;
* inspect the PRISMA counts;
* verify the quality-assessment procedure and sensitivity analysis;
* reproduce the analytical tables derived from the final corpus.

Because some bibliographic databases are queried through live APIs, re-running the pipeline may produce different retrieval counts over time. The package therefore serves as both a **reproducibility resource and an audit trail for the results reported in the article**.
