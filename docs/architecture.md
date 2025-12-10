## System Architecture

This document describes the **architectural design** of the Outlook Signature Intelligence Extraction Engine. It outlines the core components, their responsibilities, and how data flows through the system.

---

## High-Level Overview (Data Flow)

The system is structured as a two-phase pipeline using an intermediary cache layer. 

* **Input:** Outlook Desktop

* **Phase 1 – Extraction**
    * Outlook COM access
    * Folder & subfolder traversal
    * HTML cleaning
    * Signature boundary detection
    * JSONL cache writer

* **Cache Layer** (JSONL files)

* **Phase 2 – Consolidation**
    * Cache loader
    * NLP name extraction
    * Rule-based title & department detection
    * Scoring & deduplication

* **Final Output** (Excel / CSV)

---

## Architectural Goals

* **Reproducibility**: All intermediate results are stored as **JSONL** for deterministic reruns.
* **Separation of concerns**: Extraction, cleaning, enrichment, and output stages are independent.
* **Resilience**: Outlook COM access is isolated and safeguarded against common crash patterns.
* **Transparency**: NLP and rule-based decisions are explainable and debuggable.
* **Extensibility**: Adding new rules, dictionaries, or ML models does not require redesigning the pipeline.

---

## Core Components

### 1. Outlook Client (`outlook_client.py`)
Provides **safe access** to Outlook folders using `pywin32`.

**Responsibilities:**
* Resolve root folders
* Validate folder indices
* Provide `MailItem` iterators
* Handle common COM error scenarios

### 2. Folder Iterators (`outlook_iterators.py`)
A robust mechanism for iterating emails in folders, addressing:
* COM crashes
* Unexpected message formats
* Filtering by date
* Excluding ignored senders
* Notification email detection

### 3. HTML Cleaning (`html_cleaner.py`)
Transforms raw HTML into normalized plain text.
* Removes reply histories
* Strips disclaimers
* Eliminates styling and scripting tags
* Normalizes whitespace
* Produces **“core lines”** used in signature extraction

### 4. Signature Extraction (`signature_extractor.py`)
Key logic for identifying signature boundaries and meaningful content. This happens using a combination of heuristics, keyword detection, and structural analysis.

**Extracts:**
* Sender name candidates
* Job title fragments
* Department indicators
* Company references
* Relevant metadata

### 5. NLP Name Extraction (`nlp_extractor.py`)
Uses `spaCy` to:
* Detect person names
* Validate extracted names
* Suppress false positives through blacklist and scoring

### 6. Aggregation Layer (`aggregation.py`)
Responsible for **grouping and ranking** results by sender to select the single best record per email address.

**Ranking criteria include:**
* Presence of name
* Completeness of position
* Presence of department
* Recency

### 7. Export Layer (`exporter.py`)
Outputs final consolidated data in:
* **UTF-8 Excel format**
* **UTF-8-BOM CSV** (for Windows compatibility)

---

## Cache Layer

**JSONL** files store one cleaned record per email message.

**Benefits:**
* Fast reprocessing
* Easier debugging
* Separation of extraction (Outlook-dependent) from analysis (Outlook-independent)

---

## Extensibility

Additions can be made without breaking architecture:
* More rules in `/config` dictionaries
* Additional NLP models
* Machine learning classifiers
* CRM or API integrations
* Web UI layers on top

