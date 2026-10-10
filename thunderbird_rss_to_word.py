#!/usr/bin/env python3
"""Archive Thunderbird RSS subscriptions in one continuously updated Word document."""

import argparse
import hashlib
import html
from html.parser import HTMLParser
import json
from pathlib import Path
import sys
import time
from urllib.error import URLError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

try:
    import feedparser
    from docx import Document
except ImportError as exc:
    raise SystemExit(
        "Missing packages. Install them with: "
        "python -m pip install feedparser python-docx"
    ) from exc


DEFAULT_FEED_LIST = Path(__file__).with_name("rss_feeds.txt")


class TextOnlyHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)


def fetch_feed(url):
    request = Request(
        url,
        headers={"User-Agent": "ThunderbirdRSSWordArchive/1.0"},
    )
    with urlopen(request, timeout=30) as response:
        return feedparser.parse(response.read())


class ArticleJSONLDParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self._in_jsonld = False
        self._current_script = []
        self.scripts = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() != "script":
            return
        attributes = dict(attrs)
        self._in_jsonld = (
            attributes.get("type", "").split(";", 1)[0].strip().lower()
            == "application/ld+json"
        )
        self._current_script = []

    def handle_data(self, data):
        if self._in_jsonld:
            self._current_script.append(data)

    def handle_endtag(self, tag):
        if tag.lower() == "script" and self._in_jsonld:
            self.scripts.append("".join(self._current_script))
            self._current_script = []
            self._in_jsonld = False


def plain_text(value):
    parser = TextOnlyHTML()
    parser.feed(html.unescape(value or ""))
    return " ".join(" ".join(parser.parts).split())


def article_body_from_jsonld(value):
    if isinstance(value, list):
        for item in value:
            body = article_body_from_jsonld(item)
            if body:
                return body
        return ""
    if not isinstance(value, dict):
        return ""

    article_types = value.get("@type", [])
    if isinstance(article_types, str):
        article_types = [article_types]
    if any("article" in item.lower() for item in article_types if isinstance(item, str)):
        body = value.get("articleBody")
        if isinstance(body, str) and body.strip():
            return plain_text(body)
    return article_body_from_jsonld(value.get("@graph", []))


def fetch_article_text(entry):
    # Some feeds provide the full article body as content:encoded.
    for content in entry.get("content", []):
        value = plain_text(content.get("value", ""))
        if value:
            return value

    article_url = entry.get("link", "")
    parsed_url = urlsplit(article_url)
    if parsed_url.scheme not in {"http", "https"} or not parsed_url.hostname:
        raise ValueError(f"invalid article URL: {article_url or '(missing)'}")

    request = Request(
        article_url,
        headers={"User-Agent": "ThunderbirdRSSWordArchive/1.0"},
    )
    with urlopen(request, timeout=30) as response:
        page = response.read().decode(response.headers.get_content_charset() or "utf-8", errors="replace")

    parser = ArticleJSONLDParser()
    parser.feed(page)
    for script in parser.scripts:
        try:
            body = article_body_from_jsonld(json.loads(script))
        except json.JSONDecodeError:
            continue
        if body:
            return body
    raise ValueError("the article page did not include a full article body")


def read_feed_urls(feed_list_path):
    if not feed_list_path.is_file():
        raise FileNotFoundError(
            f"RSS feed list not found: {feed_list_path}. "
            "Create it with one RSS feed URL per line."
        )

    urls = []
    for line_number, raw_line in enumerate(
        feed_list_path.read_text(encoding="utf-8-sig").splitlines(), start=1
    ):
        url = raw_line.strip()
        if not url or url.startswith("#"):
            continue
        parsed_url = urlsplit(url)
        if parsed_url.scheme not in {"http", "https"} or not parsed_url.hostname:
            print(
                f"Warning: ignoring invalid URL on line {line_number} "
                f"of {feed_list_path}: {url}",
                file=sys.stderr,
            )
            continue
        if url not in urls:
            urls.append(url)

    if not urls:
        raise ValueError(f"No valid RSS feed URLs found in {feed_list_path}.")
    return urls


def entry_identity(feed_url, entry):
    identity = entry.get("id") or entry.get("link")
    if not identity:
        identity = "|".join(
            (entry.get("title", ""), entry.get("published", ""), entry.get("summary", ""))
        )
    return hashlib.sha256(f"{feed_url}\n{identity}".encode("utf-8")).hexdigest()


def read_state(state_path):
    if not state_path.exists():
        return set(), set()
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        if isinstance(state, list):
            # Older versions recorded processed items without tracking full text.
            if not all(isinstance(item, str) for item in state):
                raise ValueError("expected string item identifiers")
            return set(state), set()
        if not isinstance(state, dict):
            raise ValueError("expected an item list or state object")
        known = state.get("items", [])
        full_text = state.get("full_text_items", [])
        if not all(
            isinstance(items, list) and all(isinstance(item, str) for item in items)
            for items in (known, full_text)
        ):
            raise ValueError("expected string item identifier lists")
        return set(known), set(full_text)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        raise RuntimeError(f"Cannot read archive state {state_path}: {exc}") from exc


def write_state(state_path, identities, full_text_identities):
    temporary_path = state_path.with_suffix(state_path.suffix + ".tmp")
    temporary_path.write_text(
        json.dumps(
            {
                "items": sorted(identities),
                "full_text_items": sorted(full_text_identities),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    temporary_path.replace(state_path)


def append_to_document(document_path, items):
    document = Document(document_path) if document_path.exists() else Document()
    if not document_path.exists():
        document.add_heading("RSS Feed Archive", level=0)

    paragraphs = document.paragraphs
    for feed_title, entry, full_text in items:
        link = entry.get("link")
        link_label = f"Link: {link}" if link else None
        existing_link_index = next(
            (index for index, paragraph in enumerate(paragraphs)
             if link_label and paragraph.text == link_label),
            None,
        )
        if existing_link_index is not None:
            if existing_link_index + 1 < len(paragraphs):
                paragraphs[existing_link_index + 1].text = full_text
            else:
                paragraphs[existing_link_index].insert_paragraph_before(full_text)
            continue

        title = plain_text(entry.get("title", "")) or "(Untitled item)"
        document.add_heading(title, level=2)
        if feed_title:
            document.add_paragraph(f"Feed: {feed_title}")
        published = entry.get("published") or entry.get("updated")
        if published:
            document.add_paragraph(f"Published: {plain_text(published)}")
        if link:
            document.add_paragraph(f"Link: {link}")
        document.add_paragraph(full_text)
        document.add_paragraph("")
        paragraphs = document.paragraphs

    document.save(document_path)


def archive_new_items(document_path, state_path):
    feed_urls = read_feed_urls(DEFAULT_FEED_LIST)
    known, full_text_known = read_state(state_path)
    new_items = []
    new_identities = set(known)
    new_full_text_identities = set(full_text_known)
    for feed_url in feed_urls:
        try:
            feed = fetch_feed(feed_url)
        except (URLError, TimeoutError, OSError, ValueError) as exc:
            print(f"Warning: could not fetch feed {feed_url}: {exc}", file=sys.stderr)
            continue
        if getattr(feed, "bozo", False):
            print(
                f"Warning: feed parse issue for {feed_url}: {feed.bozo_exception}",
                file=sys.stderr,
            )

        feed_title = plain_text(feed.feed.get("title", "")) or feed_url
        for entry in feed.entries:
            identity = entry_identity(feed_url, entry)
            if identity in new_full_text_identities:
                continue
            try:
                full_text = fetch_article_text(entry)
            except (URLError, TimeoutError, OSError, ValueError) as exc:
                print(
                    f"Warning: could not retrieve full text for "
                    f"{entry.get('link', entry.get('title', '(untitled item)'))}: {exc}",
                    file=sys.stderr,
                )
                continue
            new_items.append((feed_title, entry, full_text))
            new_identities.add(identity)
            new_full_text_identities.add(identity)

    if not new_items:
        print("No new feed items.")
        return

    document_path.parent.mkdir(parents=True, exist_ok=True)
    append_to_document(document_path, new_items)
    write_state(state_path, new_identities, new_full_text_identities)
    print(f"Archived {len(new_items)} item(s) to {document_path}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path.home() / "Documents" / "Thunderbird RSS Archive.docx",
        help="Word document path (default: Documents\\Thunderbird RSS Archive.docx)",
    )
    parser.add_argument(
        "--interval-seconds",
        type=int,
        default=300,
        help="seconds between checks (default: 300)",
    )
    args = parser.parse_args()
    if args.interval_seconds < 30:
        parser.error("--interval-seconds must be at least 30")

    output_path = args.output.expanduser().resolve()
    state_path = output_path.with_suffix(".state.json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not output_path.exists():
        document = Document()
        document.add_heading("RSS Feed Archive", level=0)
        document.save(output_path)
    print(f"Watching RSS feeds listed in {DEFAULT_FEED_LIST}; archive: {output_path}")
    print("The document is created immediately and updated when feed items are found.")
    print("Press Ctrl+C to stop.")

    try:
        while True:
            try:
                archive_new_items(output_path, state_path)
            except (OSError, RuntimeError) as exc:
                print(f"Error: {exc}", file=sys.stderr)
            time.sleep(args.interval_seconds)
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
