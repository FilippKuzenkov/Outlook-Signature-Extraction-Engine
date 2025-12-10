# 📧 Outlook Signature Extraction Engine

**A production-grade ETL pipeline that extracts, cleans, enriches, and deduplicates contact signatures from Outlook mailboxes — built to solve a real strategic business problem.**

---

##  Business Context & Problem

During the preparation for the **2026 Marketing & Sales Strategy**, our team needed to understand who exactly across multiple client organizations was communicating with us, and how these individuals mapped to departmental and decision-making structures.

Manually identifying roles, departments, and contact profiles from thousands of emails was **not feasible**.

To support stakeholder mapping, communication planning, and strategic analysis, this system was developed to automate the entire extraction, cleaning, and consolidation process.

---

##  Impact & Value

The system **eliminated weeks of manual review** and helped the marketing team identify the core audience, directly assisting with the 2026 Marketing Strategy preparation.

| Metric | Result |
| :--- | :--- |
| **Emails Scanned** | 11,352 |
| **High-Value Profiles Extracted** | 250–300 |
| **Automatic Detections** | Names, Job Titles, Departments, Relevant Signature Context |

The pipeline is reusable for ongoing strategic and operational analysis.

---

##  Project Overview & Architecture

The **Outlook Signature Extraction Engine** transforms unstructured, multilingual email signatures into clean, structured datasets for analysis, CRM enrichment, and strategy building.

It is a robust multi-phase system designed for:
* **Enterprise Outlook environments**
* **Volume processing** (11,000+ emails)
* Handling **inconsistent HTML signatures**
* **Multilingual communication**
* **Deduplication** across time periods

### Solution Architecture (ETL Pipeline)
The system is built on a clear, isolated, and testable multi-phase structure.

The data flow is structured as follows:

  **Outlook Mailbox (11,352 emails)**

  **Phase 1: Extraction & JSONL Caching**

  **HTML Cleaning + Signature Boundary Detection**

  **Phase 2: NLP & Rule-Based Enrichment**

  **Ranking & Deduplication of Contact Profiles**

##  Features

### Signature Extraction
* Noise-resistant **HTML normalization**.
* Removal of reply history and disclaimers.
* Multilingual courtesy-phrase detection.
* Boundary detection for accurate signature blocks.

### NLP & Rule-Based Classification
* `spaCy` name extraction.
* Keyword-based **job title inference**.
* **Department classification** using multi-dictionary matching.
* Confidence scoring of extracted fields.

### Deduplication & Consolidation
* Per-sender scoring.
* **“Latest information wins”** logic for updated contact details.
* Sorting by completeness for manual or analytical review.

### Reliability & Scalability
* **JSONL cache layer** for deterministic processing and fast re-runs.
* Safe handling of Outlook COM iteration.
* **Modular design** for easy extension.

---

##  Engineering Approach & Design Principles

This project was developed with a focus on solving a complex operational problem while applying solid engineering practices suitable for scalable, maintainable data pipelines.

### Engineering Highlights
* A **multi-phase ETL pipeline** cleanly separates extraction, cleaning, enrichment, and output.
* A **deterministic JSONL cache layer** ensures reproducibility and easier debugging.
* A **hybrid NLP + rule-based system** handles inconsistent, multilingual signatures.
* **Fault-tolerant Outlook COM access** ensures stable scanning.
* A **deduplication and scoring system** selects the most complete and recent record.

### Design Principles
| Principle | Description |
| :--- | :--- |
| **Separation of Concerns** | Each processing phase is isolated, enabling independent development and testing. |
| **Resilience** | Outlook COM interactions are safeguarded against intermittent failures. |
| **Explainability** | Classification rules and decision criteria are transparent and easy to audit or modify. |
| **Extensibility** | New job titles, departments, or ML models can be integrated without major architectural changes. |
| **Data Quality** | Prioritizes freshness, completeness, and clarity of extracted contact data. |

---

##  Tech Stack

| Category | Tools |
| :--- | :--- |
| **Core** | Python 3.10+ |
| **Data Processing** | `pandas` |
| **NLP** | `spaCy` |
| **HTML Parsing** | `beautifulsoup4`, `lxml` |
| **Outlook Access** | `pywin32` (Outlook COM API) |
| **Output** | `openpyxl` |

---

##  Installation & Usage

### Installation
```shell
git clone <repo-url>
pip install -r requirements.txt
