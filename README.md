Outlook Signature Intelligence Extraction Engine

A production-grade ETL pipeline that extracts, cleans, enriches, and deduplicates contact signatures from Outlook mailboxes — built to solve a real strategic business problem.

Business Context

During the preparation for the 2026 Marketing & Sales Strategy, our team needed to understand who exactly across multiple client organizations was communicating with us, and how these individuals mapped to departmental and decision-making structures.

Manually identifying roles, departments, and contact profiles from thousands of emails was not feasible.
To support stakeholder mapping, communication planning, and strategic analysis, this system was developed to automate the entire extraction, cleaning, and consolidation process.

Impact

- 11,352 emails scanned
- ≈250–300 high-value structured contact profiles extracted

Automatic detection of:
- Names
- Job titles
- Departments
- Relevant signature context


- Eliminated weeks of manual review
- Helped the marketing team identify the core audience and assisted with preparing the 2026 Marketing Strategy
- Is reusable for ongoing strategic and operational analysis


Project Overview

The Outlook Signature Extraction Engine transforms unstructured, multilingual email signatures into clean, structured datasets for analysis, CRM enrichment, and strategy building.

It is a robust multi-phase system designed for:
- Enterprise Outlook environments
- Volume processing
- Inconsistent HTML signatures
- Multilingual communication
- Deduplication across time periods


Solution Architecture
Outlook Mailbox (11,352 emails)
        ↓
Phase 1 – Extraction & JSONL Caching
        ↓
HTML Cleaning + Signature Boundary Detection
        ↓
Phase 2 – NLP & Rule-Based Enrichment
        ↓
Ranking & Deduplication of Contact Profiles
        ↓
Final Output: 250–300 structured records (Excel/CSV)

Each layer is isolated, testable, and built for reproducibility.

Features
Signature Extraction

Noise-resistant HTML normalization
Removal of reply history & disclaimers
Multilingual courtesy-phrase detection
Boundary detection for signature blocks

NLP & Rule-Based Classification

spaCy name extraction
Keyword-based job title inference
Department classification using multi-dictionary matching
Confidence scoring of extracted fields

Deduplication & Consolidation

Per-sender scoring
“Latest information wins” logic
Sorting by completeness for manual or analytical review

Reliability & Scalability

JSONL cache layer for deterministic processing
Safe handling of Outlook COM iteration
Modular design for easy extension


Example Output
A sample anonymized output file is included.

Engineering Approach & Design Principles
This project was developed with a focus on solving a complex operational problem while applying solid engineering practices suitable for scalable, maintainable data pipelines.

Engineering Highlights

A multi-phase ETL pipeline cleanly separates Outlook extraction, HTML cleaning, NLP enrichment, and output generation.
A deterministic JSONL cache layer ensures reproducibility, fast re-runs, and easier debugging.
A hybrid NLP + rule-based system extracts names, job titles, and departments from inconsistent, multilingual signatures.
HTML signatures are normalized through a dedicated pipeline that removes reply chains, disclaimers, and formatting artifacts.
Fault-tolerant Outlook COM access ensures stable scanning even in large or complex mailbox structures.
A deduplication and scoring system evaluates signature candidates and selects the most complete and recent record for each sender.


Design Principles

Separation of Concerns: Each processing phase is isolated, enabling independent development and testing.
Resilience: Outlook COM interactions are safeguarded against intermittent failures and inconsistent folder structures.
Explainability: Classification rules, keyword sets, and decision criteria are transparent and easy to audit or modify.
Extensibility: New job titles, departments, or ML models can be integrated without major architectural changes.
Data Quality: Prioritizes freshness, completeness, and clarity of extracted contact data to support downstream business workflows.


Why This Architecture?
Extracting structured insights from more than 11,000 heterogeneous emails required:

A robust HTML normalization pipeline
NLP-powered entity extraction
Rule-based classification for job titles & departments
A reproducible multi-stage ETL structure
A consolidation system capable of selecting best-of-multiple signatures
Numeric ranking to handle inconsistent or incomplete signatures

These architectural choices make the pipeline suitable for:

Stakeholder mapping
CRM enrichment
Strategy development
Workflow automation
Organizational network analysis


Tech Stack

Python 3.10+
pandas
spaCy
beautifulsoup4
lxml
pywin32 (Outlook COM API)
openpyxl


Installation
Shellgit clone <repo-url>pip install -r requirements.txtShow more lines

Usage
Shell# Phase 1: scan root folderspython phase1_scan_folder.py# Phase 1: scan only subfolderspython phase1_scan_subfolders_only.py# Phase 2: consolidate and enrich dataShow more lines
Optional utilities are in /tools.
