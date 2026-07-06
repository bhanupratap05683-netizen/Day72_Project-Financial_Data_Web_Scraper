# Financial Data Web Scraper

## Overview
An end-to-end financial data collection tool. Scrapes stock screener data
across multiple pages, validates and cleans it, computes summary analytics,
and generates a professionally formatted multi-sheet Excel report with full
logging. Capstone project for Phase 5 of the 84-Day Python & Excel Roadmap.

## Pipeline

- Multi-page fetch (retry + backoff) → Parse (BeautifulSoup) → Validate rows
  → Clean & dedupe (pandas) → Summary analytics → 3-sheet formatted Excel
 
## Features
- Multi-page scraping with polite delays between requests
- Retry logic with exponential backoff (Timeout / ConnectionError handling)
- Row-level validation — skips bad data, never crashes on it
- Duplicate removal and price-sorted output
- Summary sheet: averages, gainers/losers, top movers
- Professional Excel styling (headers, borders, auto column widths)
- Complete audit trail in `scraper_project.log`
- Offline sample mode for testing without a network

## Output
- `financial_data_report.xlsx` — sheets: Market Data, Summary, Top Movers
- `scraper_project.log` — timestamped execution log

## Tech Stack
requests · beautifulsoup4 · pandas · openpyxl · logging

## Installation
```bash
pip install requests beautifulsoup4 pandas openpyxl lxml
