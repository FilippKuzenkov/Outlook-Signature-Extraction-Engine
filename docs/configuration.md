
## Configuration Guide

This document explains how to configure the Outlook Signature Extraction Engine for different environments and processing needs.

---

## Folder Selection

The file `/config/folders.txt` contains a list of **Outlook folder paths** to be scanned by the pipeline.

**Example (Anonymized):**
Inbox\2025\03_March Inbox\2025\04_April Inbox\2025\05_May\Client_A


**Scanning Modes:**
Phase 1 allows scanning:
* The root folder only
* Subfolders only
* Both (depending on the mode used during script execution)

---

## Exclusion Lists

The system uses specific configuration files to filter out non-essential or automated communications.

### Ignored Senders
The file `/config/ignored_senders.csv` contains a list of email addresses that should be **explicitly excluded** from processing, as they often belong to automated systems.

**Example:**
no-reply@example.com notification@example.com support@example.com


### Notification Detection
The file `/config/notification_patterns.txt` contains patterns used to exclude generic automated emails (e.g., system notifications, ticketing systems).

---

## Keyword Dictionaries (Rule-Based Classification)

These plain text files contain keywords used during **Phase 2 (NLP & Rule-Based Enrichment)** to classify titles and departments accurately.

| File Path | Purpose | Example Keywords |
| :--- | :--- | :--- |
| `/config/job_title_keywords.txt` | Core terms used to identify job titles. | `manager`, `director`, `specialist`, `lead`, `head of` |
| `/config/department_keywords.txt` | Core terms used to classify departments. | `marketing`, `sales`, `engineering`, `procurement`, `finance` |
| `/config/extra_title_tokens.txt` | Additional variants or synonyms for job title terms. | |
| `/config/extra_department_tokens.txt`| Additional phrasing or synonyms for departments. | |

---

## Meta File List

A helper file, `/config/meta_files.txt`, is used to list the available **JSONL cache files** that are ready for Phase 2 consolidation and enrichment processing.
