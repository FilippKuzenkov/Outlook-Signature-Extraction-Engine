
## Pipeline Overview

This document provides a functional **walk-through** of the full ETL (Extract, Transform, Load) pipeline used to extract structured signatures from Outlook mailboxes.

---

## Phase 1: Extraction Pipeline (E - Extract)

### Step 1 — Select Outlook folders
User selects:
* Monthly folders
* Subfolder-only mode or full root traversal

The system supports complex folder trees.

### Step 2 — Iterate Emails
For each `MailItem`:
* Validate message type
* Extract metadata:
    * sender email
    * sender name
    * received timestamp
    * subject

### Step 3 — Clean HTML
Raw HTML is passed to the HTML cleaner:
* Strip reply chains (EN/DE detection)
* Remove embedded disclaimers
* Remove styling/script blocks
* Normalize whitespace
* Produce **“core lines”** list

**Example output:**
```text
["Best regards","Alice Smith","Senior Marketing Manager","Marketing & Communications"]
```

### Step 4 — Write JSONL Cache
Each processed email is stored as:
```json
{
  "sender_email": "alice.smith@example.com",
  "sender_name": "Alice Smith",
  "received_time": "2025-03-14T10:22:55",
  "core_lines": [
    "Best regards",
    "Alice Smith",
    "Senior Marketing Manager",
    "Marketing & Communications"
  ]
}
```

---

## Phase 2: Consolidation & Enrichment (T - Transform, L - Load)

### Step 5 — Load Cache
Phase 2 is independent from Outlook. JSONL cache is read and aggregated per sender.

### Step 6 — NLP Extraction
The spaCy model identifies:
* Person names  
False positives removed using blacklist.

### Step 7 — Rule-Based Classification
Keyword dictionaries determine:
* Job titles
* Department names
* Additional functional keywords  
Rules handle multilingual input and noisy signature text.

### Step 8 — Scoring & Deduplication
For each sender, a ranking score is computed:

| Field                | Score | Priority       |
|----------------------|-------|---------------|
| Valid extracted name | +2    | High          |
| Job title present    | +2    | High          |
| Department detected  | +1    | Medium        |
| Newer timestamp      | N/A   | Highest Priority |

The best candidate per email is selected.

### Step 9 — Export
Export formats:
* Excel (.xlsx)
* CSV (UTF-8-BOM)

Final output is a neatly structured table with columns such as:
* `sender_email`
* `sender_name`
* `detected_name`
* `position`
* `department`
* `debug_core_lines`
