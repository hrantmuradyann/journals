#!/usr/bin/env python3
"""Fetch recent articles from several publications' RSS feeds and save them to an Excel file."""

import html
import os
import re
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

from openpyxl import Workbook
from openpyxl.styles import Font

# Each publication gets its own sheet in the Excel file.
# Inside a publication, each entry is  "Section name": "feed url".
# A section can also be  "Section name": ("feed url", "Category")  when a publication has only
# one combined feed whose items carry a <category> tag - then only items with that category
# are kept for the section. The feed is downloaded once and reused for every section.
FEEDS = {
    "WSJ": {
        "World News": "https://feeds.content.dowjones.io/public/rss/RSSWorldNews",
        "US Business": "https://feeds.content.dowjones.io/public/rss/WSJcomUSBusiness",
        "Markets": "https://feeds.content.dowjones.io/public/rss/RSSMarketsMain",
        "Technology": "https://feeds.content.dowjones.io/public/rss/RSSWSJD",
        "Opinion": "https://feeds.content.dowjones.io/public/rss/RSSOpinion",
        "Lifestyle": "https://feeds.content.dowjones.io/public/rss/RSSLifestyle",
        "Arts & Culture": "https://feeds.content.dowjones.io/public/rss/RSSArtsCulture",
        "Personal Finance": "https://feeds.content.dowjones.io/public/rss/RSSPersonalFinance",
    },
    "FT": {
        "Home": "https://www.ft.com/rss/home",
        "World": "https://www.ft.com/world?format=rss",
        "Companies": "https://www.ft.com/companies?format=rss",
        "Markets": "https://www.ft.com/markets?format=rss",
        "Technology": "https://www.ft.com/technology?format=rss",
        "Opinion": "https://www.ft.com/opinion?format=rss",
    },
    "Bloomberg": {
        "Markets": "https://feeds.bloomberg.com/markets/news.rss",
        "Business": "https://feeds.bloomberg.com/business/news.rss",
        "Economics": "https://feeds.bloomberg.com/economics/news.rss",
        "Politics": "https://feeds.bloomberg.com/politics/news.rss",
        "Technology": "https://feeds.bloomberg.com/technology/news.rss",
        "Opinion": "https://feeds.bloomberg.com/bview/news.rss",
    },
    "NYT": {
        "Home Page": "https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml",
        "World": "https://rss.nytimes.com/services/xml/rss/nyt/World.xml",
        "Business": "https://rss.nytimes.com/services/xml/rss/nyt/Business.xml",
        "Economy": "https://rss.nytimes.com/services/xml/rss/nyt/Economy.xml",
        "Technology": "https://rss.nytimes.com/services/xml/rss/nyt/Technology.xml",
        "Opinion": "https://rss.nytimes.com/services/xml/rss/nyt/Opinion.xml",
    },
    "The Atlantic": {
        "All": "https://www.theatlantic.com/feed/all/",
        "Best Of": "https://www.theatlantic.com/feed/best-of/",
    },
    "The New Yorker": {
        "Everything": "https://www.newyorker.com/feed/everything",
        "News": "https://www.newyorker.com/feed/news",
        "Culture": "https://www.newyorker.com/feed/culture",
        "Magazine": "https://www.newyorker.com/feed/magazine/rss",
    },
    "The Economist": {
        "Latest": "https://www.economist.com/latest/rss.xml",
        "The World This Week": "https://www.economist.com/the-world-this-week/rss.xml",
        "Leaders": "https://www.economist.com/leaders/rss.xml",
        "Business": "https://www.economist.com/business/rss.xml",
        "Finance & Economics": "https://www.economist.com/finance-and-economics/rss.xml",
        "International": "https://www.economist.com/international/rss.xml",
        "United States": "https://www.economist.com/united-states/rss.xml",
        "The Americas": "https://www.economist.com/the-americas/rss.xml",
        "Europe": "https://www.economist.com/europe/rss.xml",
        "Britain": "https://www.economist.com/britain/rss.xml",
        "Asia": "https://www.economist.com/asia/rss.xml",
        "China": "https://www.economist.com/china/rss.xml",
        "Middle East & Africa": "https://www.economist.com/middle-east-and-africa/rss.xml",
        "Science & Technology": "https://www.economist.com/science-and-technology/rss.xml",
    },
    "The Guardian": {
        "International": "https://www.theguardian.com/international/rss",
        "World": "https://www.theguardian.com/world/rss",
        "Business": "https://www.theguardian.com/uk/business/rss",
        "Politics": "https://www.theguardian.com/politics/rss",
        "Technology": "https://www.theguardian.com/uk/technology/rss",
        "Opinion": "https://www.theguardian.com/uk/commentisfree/rss",
        "Culture": "https://www.theguardian.com/uk/culture/rss",
    },
    "Le Monde": {
        "Front Page": "https://www.lemonde.fr/en/rss/une.xml",
        "International": "https://www.lemonde.fr/en/international/rss_full.xml",
        "France": "https://www.lemonde.fr/en/france/rss_full.xml",
        "Politics": "https://www.lemonde.fr/en/politics/rss_full.xml",
        "Economy": "https://www.lemonde.fr/en/economy/rss_full.xml",
        "Opinion": "https://www.lemonde.fr/en/opinion/rss_full.xml",
        "Culture": "https://www.lemonde.fr/en/culture/rss_full.xml",
    },
    # Semafor has a single feed for everything; items are tagged with a category.
    "Semafor": {
        "Politics": ("https://www.semafor.com/rss.xml", "Politics"),
        "Business": ("https://www.semafor.com/rss.xml", "Business"),
        "Technology": ("https://www.semafor.com/rss.xml", "Technology"),
        "Gulf": ("https://www.semafor.com/rss.xml", "Gulf"),
        "Africa": ("https://www.semafor.com/rss.xml", "Africa"),
        "China": ("https://www.semafor.com/rss.xml", "China"),
        "Energy": ("https://www.semafor.com/rss.xml", "Energy"),
        "Security": ("https://www.semafor.com/rss.xml", "Security"),
    },
}

# Max articles taken from each individual feed
MAX_ARTICLES = 10

# Bloomberg's feed server either answers within ~5 seconds or hangs forever (measured: a
# request that hasn't answered by 10 s never does). So use a short timeout and simply try
# again several times; a feed is only skipped after every attempt has failed.
FETCH_ATTEMPTS = 5
FETCH_TIMEOUT_SECONDS = 10

# Time zone to show publish dates in (GMT+4)
LOCAL_TZ = timezone(timedelta(hours=4))

# Save the Excel file next to this script
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
)

ATOM_NS = "{http://www.w3.org/2005/Atom}"
# Most RSS feeds put the byline in <dc:creator> (WSJ, NYT, Bloomberg, Guardian, New Yorker).
# FT, The Economist and Le Monde do not include authors in their feeds at all.
DC_NS = "{http://purl.org/dc/elements/1.1/}"


def clean_text(text):
    """Strip HTML tags and unescape entities from feed text."""
    if text is None:
        return ""
    text = html.unescape(text)
    text = re.sub(r"<[^>]+>", "", text)
    return " ".join(text.split())


def to_local_time(published):
    """Convert an RSS/Atom date string to 'YYYY-MM-DD HH:MM' in LOCAL_TZ."""
    if not published:
        return ""
    parsed = None
    # RSS uses "Thu, 17 Sep 2026 08:00:00 GMT"
    try:
        parsed = parsedate_to_datetime(published)
    except (TypeError, ValueError):
        pass
    # Atom uses "2026-09-17T08:00:00-04:00"
    if parsed is None:
        try:
            parsed = datetime.fromisoformat(published.replace("Z", "+00:00"))
        except ValueError:
            return published
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(LOCAL_TZ).strftime("%Y-%m-%d %H:%M")


def parse_rss(root):
    """Parse an RSS 2.0 document into a list of (title, description, link, published, author, category)."""
    articles = []
    for item in root.iter("item"):
        title = clean_text(item.findtext("title"))
        description = clean_text(item.findtext("description"))
        link = (item.findtext("link") or "").strip()
        published = to_local_time((item.findtext("pubDate") or "").strip())
        author = clean_text(item.findtext(DC_NS + "creator"))
        if not author:
            author = clean_text(item.findtext("author"))
        category = clean_text(item.findtext("category"))
        articles.append((title, description, link, published, author, category))
    return articles


def parse_atom(root):
    """Parse an Atom document into a list of (title, description, link, published, author, category)."""
    articles = []
    for entry in root.iter(ATOM_NS + "entry"):
        title = clean_text(entry.findtext(ATOM_NS + "title"))
        description = clean_text(entry.findtext(ATOM_NS + "summary"))
        if not description:
            description = clean_text(entry.findtext(ATOM_NS + "content"))
        link = ""
        for link_tag in entry.findall(ATOM_NS + "link"):
            if link_tag.get("rel") in (None, "alternate"):
                link = link_tag.get("href", "").strip()
                break
        published = entry.findtext(ATOM_NS + "published") or entry.findtext(ATOM_NS + "updated") or ""
        published = to_local_time(published.strip())
        author = clean_text(entry.findtext(ATOM_NS + "author/" + ATOM_NS + "name"))
        category = ""
        category_tag = entry.find(ATOM_NS + "category")
        if category_tag is not None:
            category = clean_text(category_tag.get("term"))
        articles.append((title, description, link, published, author, category))
    return articles


def download(url):
    """Download a URL, retrying a few times because feed servers can be slow."""
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    last_error = None
    for attempt in range(FETCH_ATTEMPTS):
        try:
            with urllib.request.urlopen(request, timeout=FETCH_TIMEOUT_SECONDS) as response:
                return response.read()
        except Exception as error:
            last_error = error
            if attempt < FETCH_ATTEMPTS - 1:
                time.sleep(2)
    raise last_error


# Feeds already downloaded during this run, so a combined feed used by several sections
# (see Semafor) is only fetched once.
_feed_cache = {}


def fetch_feed(url):
    """Download a feed and return a list of (title, description, link, published, author, category) tuples."""
    if url in _feed_cache:
        return _feed_cache[url]
    data = download(url)
    root = ET.fromstring(data)
    if root.tag == ATOM_NS + "feed":
        articles = parse_atom(root)
    else:
        articles = parse_rss(root)
    _feed_cache[url] = articles
    return articles


def main():
    workbook = Workbook()
    workbook.remove(workbook.active)  # drop the default empty sheet

    total = 0
    for publication, sections in FEEDS.items():
        sheet = workbook.create_sheet(title=publication)
        sheet.append(["Section", "Title", "Description", "Link", "Published (GMT+4)", "Author"])
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        for section, feed in sections.items():
            if isinstance(feed, tuple):
                url, category_filter = feed
            else:
                url, category_filter = feed, None
            try:
                articles = fetch_feed(url)
            except Exception as error:
                print(f"Could not fetch {publication} / {section}: {error}", file=sys.stderr)
                continue

            if category_filter is not None:
                articles = [a for a in articles if a[5] == category_filter]

            for title, description, link, published, author, category in articles[:MAX_ARTICLES]:
                sheet.append([section, title, description, link, published, author])
                link_cell = sheet.cell(row=sheet.max_row, column=4)
                link_cell.hyperlink = link
                link_cell.font = Font(color="0563C1", underline="single")
                total += 1

        sheet.column_dimensions["A"].width = 20
        sheet.column_dimensions["B"].width = 60
        sheet.column_dimensions["C"].width = 90
        sheet.column_dimensions["D"].width = 60
        sheet.column_dimensions["E"].width = 20
        sheet.column_dimensions["F"].width = 30
        print(f"{publication}: {sheet.max_row - 1} articles")

    # Always write to the same file so each run replaces the previous one
    output_path = os.path.join(OUTPUT_DIR, "news_articles.xlsx")
    workbook.save(output_path)
    print(f"Saved {total} articles to {output_path}")


if __name__ == "__main__":
    main()
