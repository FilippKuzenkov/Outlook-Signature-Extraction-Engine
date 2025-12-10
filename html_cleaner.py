# html_cleaner.py

import re
from bs4 import BeautifulSoup

REPLY_MARKERS = [
    r"-----Original Message-----",
    r"-----Ursprüngliche Nachricht-----",
    r"From:",
    r"Von:",
    r"Sent:",
    r"Gesendet:",
    r"To:",
    r"An:",
    r"Subject:",
    r"Betreff:",
    r"On .* wrote:",
    r"Am .* schrieb .*:",
]

REPLY_REGEX = re.compile("|".join(REPLY_MARKERS), re.IGNORECASE | re.MULTILINE)

def strip_reply_history_lines(lines: list[str]) -> list[str]:
    """
    Line-level version of reply history removal.
    Stops at first line that matches any reply marker
    (e.g. 'Von:', 'From:', '-----Original Message-----', etc.).
    """
    out = []
    for line in lines:
        if REPLY_REGEX.search(line):
            break
        out.append(line)
    return out

def strip_reply_history_from_html(html: str) -> str:
    """
    Removes any reply chain below the first incoming email that hit the inbox.
    Only keeps the newest email content (top block).
    """

    if not html:
        return html

    soup = BeautifulSoup(html, "html.parser")

    # Convert to plain text for marker detection
    text = soup.get_text("\n")

    match = REPLY_REGEX.search(text)
    if not match:
        return html  # no reply chain found, return original

    # Determine how many characters correspond to the first email block
    cutoff_text = text[:match.start()]

    # Rebuild minimal HTML: keep only elements contributing to that text
    cleaned_html = []
    for line in cutoff_text.splitlines():
        if line.strip():
            cleaned_html.append(f"<p>{line.strip()}</p>")

    return "\n".join(cleaned_html)


def html_to_clean_lines(html: str) -> list[str]:
    """
    Standard HTML → lines cleaning: remove style/script, normalize whitespace.
    """

    # FIRST: strip reply history strictly
    html = strip_reply_history_from_html(html)

    soup = BeautifulSoup(html, "html.parser")

    # Remove irrelevant tags
    for t in soup(["style", "script"]):
        t.decompose()

    # Extract visible text
    text = soup.get_text("\n")

    # Normalize lines
    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    return lines
