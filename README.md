# quotes-scraper

Scrapes all 100 quotes from [quotes.toscrape.com](https://quotes.toscrape.com) (a sandbox site made for practicing web scraping), visits each author's page for their details, and saves everything to JSON and Excel with some stats on top.

## Extracted data

For each quote: text, author, tags.

For each author (50 in total): name, birth date, birthplace, country and bio.

The Excel file has 5 sheets:

- **Quotes** - every quote with its author, the author's country and tags
- **Authors** - number of quotes per author, birth date, birthplace and all the tags they've been tagged with
- **Tags** - how many quotes use each tag
- **Tag pairs** - tags that often show up together on the same quote
- **Countries** - where the authors were born

Sample from the JSON:

```json
{
  "text": "The world as we have created it is a process of our thinking. It cannot be changed without changing our thinking.",
  "author": "Albert Einstein",
  "author_url": "https://quotes.toscrape.com/author/Albert-Einstein",
  "tags": ["change", "deep-thoughts", "thinking", "world"],
  "author_born": "1879-03-14"
}
```

## How it works

1. Goes through the quote pages until there's no "Next" button
2. Collects the author links in a set, so 100 quotes only mean 50 author pages to download
3. Downloads the author pages 5 at a time with a thread pool
4. Joins quotes and authors on the author url and counts tags, tag pairs and countries with `Counter`

## Setup

```
pip install -r requirements.txt
```

Needs Python 3.9 or newer.

## Usage

```
python scraper.py
```

On a Mac use `python3` instead of `python`.

It takes about 15 seconds and saves `quotes.json` and `quotes.xlsx` to the `output` folder.

Some options:

```
python scraper.py --delay 1      # wait 1 second between quote pages (default 0.5)
python scraper.py --workers 2    # fetch fewer author pages at the same time (default 5)
python scraper.py --out data     # save somewhere else
```

## Tests

```
pip install pytest
python -m pytest
```

The tests use a small local copy of the site's html, so they don't send requests to the real website.

Built with requests, BeautifulSoup and openpyxl.
