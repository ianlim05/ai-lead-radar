"""
GBS Lead Radar - test version
Searches Google News + Bing News (Malaysia focus) and optional extra RSS feeds for companies planning a GBS / shared services / COE,
scores each news item 1-5, saves results.csv and prints the top leads.
Does NOT touch LinkedIn. The LinkedIn column is just a search link for the salesperson.

Setup:  pip install requests beautifulsoup4
Run:    python main.py
Optional AI scoring (free Gemini key from aistudio.google.com/apikey):
        set GEMINI_API_KEY=your_key      (Windows CMD)
        $env:GEMINI_API_KEY="your_key"   (PowerShell)
"""
import csv
import json
import os
from dotenv import load_dotenv
load_dotenv()  # read GEMINI_API_KEY from .env
import re
import time
import urllib.parse
from collections import Counter
from datetime import datetime, timedelta, timezone
from urllib.parse import parse_qs, urlparse
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
try:
    from playwright_stealth import stealth_sync
except ImportError:
    stealth_sync = None
import requests

# ---------- SETTINGS (edit these) ----------
# Keywords come from this Excel file if it exists (sheet "Keywords": column A = Keyword, column B = Active Yes/No).
KEYWORDS_FILE = "GBS Keywords.xlsx"
OUTPUT_FILE = "results.csv"  # can also be a full path, e.g. a synced SharePoint folder

# Built-in list, only used if the Excel file is missing or cannot be read:
KEYWORDS = [
    '"shared services centre" Malaysia',
    '"global business services" Malaysia',
    '"global capability centre" Malaysia',
    '"global capability center" Malaysia',
    '"set up" "shared services centre"',
    '"establish" "global business services"',
    '"regional hub" "shared services" Kuala Lumpur',
    '"centre of excellence" "shared services"',
    '"GBS" MDEC',
    '"shared services" Penang',
]

# Malaysian outlets, searched through Google News with a site: filter
MALAYSIA_SITES = ["thestar.com.my", "theedgemalaysia.com", "nst.com.my", "malaymail.com", "bernama.com"]
DAYS = 7
MIN_LEAD_SCORE = 4  # 3 for testing so borderline items show up. Use 4 for real leads.
MAX_AI_CALLS = 60
COMPANY_ALIASES = {"tp": "Teleperformance"}  # same company under different names in the news (lowercase key)
OUTREACH_MIN_SCORE = 4   # only leads at this score or higher get target roles and a draft message
MAX_OUTREACH_CALLS = 15  # max companies drafted per run (protects the free AI quota)
YOUR_COMPANY = "Agos Asia Sdn Bhd"  # who the draft message says the sender works for
USE_BING = True  # Bing News RSS as a second search source (adds results Google News does not show)

# Extra RSS feed addresses, one per line inside the brackets.
EXTRA_RSS_FEEDS = [
    "https://www.freemalaysiatoday.com/category/business/feed/",     # Free Malaysia Today, business
    "https://www.malaymail.com/feed/rss/money",                      # Malay Mail, money
    "https://www.techinasia.com/feed",                               # Tech in Asia
    "https://feed.businesswire.com/rss/home/?rss=G1QFDERJXkJeGVpT",  # Business Wire Asia
]

GEMINI_KEY = os.getenv("GEMINI_API_KEY")
# Tried in order. If one is overloaded (503) the next is used. Check current free model names in AI Studio.
GEMINI_MODELS = ["gemini-3.8-flash", "gemini-flash-lite-latest"]

# -------------------------------------------

TOPIC = ["global business services", "shared services", "global capability cent",
         "centre of excellence", "center of excellence"]
ACTION = ["set up", "setting up", "establish", "launch", "open", "new centre", "new center",
          "expand", "expansion", "invest", "hub", "announce"]

def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())[:80]

def clean_company(name):
    name = (name or "").strip()
    return COMPANY_ALIASES.get(name.lower(), name)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}
RELEVANT = re.compile(r"shared services|global business services|global capability|capability cent|"
                      r"centre of excellence|center of excellence|\b(gbs|gcc|coe)\b", re.I)

def parse_rss(content, origin, relevant_only=False):
    """Read an RSS 2.0 feed. Drops items older than DAYS. Optionally keeps only topic-related items."""
    root = ET.fromstring(content)
    cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS)
    out = []
    for it in root.iter("item"):
        title = (it.findtext("title") or "").strip()
        link = (it.findtext("link") or "").strip()
        if not title or not link:
            continue
        if relevant_only:
            text = title + " " + re.sub(r"<[^>]+>", " ", it.findtext("description") or "")
            if not RELEVANT.search(text):
                continue
        date = ""
        pub = it.findtext("pubDate")
        if pub:
            try:
                d = parsedate_to_datetime(pub)
                if d.tzinfo is None:
                    d = d.replace(tzinfo=timezone.utc)
                if d < cutoff:
                    continue
                date = d.strftime("%Y-%m-%d")
            except (TypeError, ValueError):
                pass
        source = (it.findtext("source") or "").strip()
        if not source:
            for child in it:  # Bing puts the publisher in a namespaced <Source> tag
                if child.tag.lower().split("}")[-1] == "source":
                    source = (child.text or "").strip()
        if "bing.com/news/apiclick" in link:  # use the real article address, not Bing's redirect
            real = parse_qs(urlparse(link).query).get("url")
            if real:
                link = real[0]
        out.append({"date": date, "title": title, "source": source, "link": link, "origin": origin})
    return out

def fetch_gnews(query):
    params = {"q": f"{query} when:{DAYS}d", "hl": "en-MY", "gl": "MY", "ceid": "MY:en"}
    r = requests.get("https://news.google.com/rss/search", params=params, headers=HEADERS, timeout=40)
    return parse_rss(r.content, "GoogleNews")

def fetch_bing(query):
    out = []
    for q in (query, query + " loc:my"):  # second pass limits Bing to Malaysian news sources
        r = requests.get("https://www.bing.com/news/search", params={"q": q, "format": "rss"},
                         headers=HEADERS, timeout=40)
        out += parse_rss(r.content, "BingNews")
        time.sleep(1.5)
    return out

def rule_score(title):
    """Free fallback: 4 if topic + action words, 2 if topic only, else 1."""
    t = title.lower()
    topic = any(k in t for k in TOPIC) or re.search(r"\b(gbs|coe)\b", t)
    action = any(k in t for k in ACTION)
    if topic and action:
        return 4
    return 2 if topic else 1

def ai_score(title):
    prompt = (
        'From this news headline, return ONE JSON object (not a list), no other text: {"company": string or "", '
        '"signal": short phrase on whether the company is planning/setting up a GBS, shared services '
        'centre, global capability centre or COE, "location": string, "score": integer 1-5 for how '
        'strong a sales lead this is. Scoring rules: 5 = a named company clearly announces setting up or expanding its OWN GBS, '
        'shared services centre, global capability centre or COE; 3-4 = likely plans, or hiring or a site expansion for such a '
        'centre; 2 = a vendor or outsourcing deal, or a relevant company only mentioned in passing; 1 = executive '
        'appointments, reports, surveys, commentary, government news or anything unrelated}. '
        "Headline: " + title
    )
    return gemini_json(prompt)

def gemini_json(prompt):
    last = ""
    for model in GEMINI_MODELS:
        for attempt in range(2):
            try:
                r = requests.post(
                    f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                    headers={"x-goog-api-key": GEMINI_KEY},
                    json={"contents": [{"parts": [{"text": prompt}]}],
                          "generationConfig": {"responseMimeType": "application/json"}},
                    timeout=45,
                )
            except requests.exceptions.RequestException as e:
                last = f"{model}: {type(e).__name__}"
                time.sleep(8)
                continue
            if r.status_code == 429:  # quota used up for this model, try the next model
                last = f"{model}: 429 {r.text[:200]}"
                break
            if r.status_code in (500, 502, 503, 504):  # Google is busy, temporary
                last = f"{model}: {r.status_code} busy"
                time.sleep(8)
                continue
            if r.status_code != 200:  # key or model problem, retrying will not help
                raise PermissionError(f"Gemini error {r.status_code} on {model}: {r.text[:250]}")
            data = json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
            if isinstance(data, list):  # sometimes the model wraps the answer in a list
                data = data[0] if data else {}
            if not isinstance(data, dict):
                raise ValueError("unexpected AI reply format")
            return data
    raise ConnectionError("Gemini is busy, over its free quota, or unreachable (" + last + ")")

def fetch_linkedin_company_data(company_name):
    """
    Uses local Playwright (Headless Chromium) with stealth evasion.
    No login required; extracts data from public search and LinkedIn SEO metadata.
    Includes full description and the exact verified URL.
    """
    if not company_name:
        return ""

    user_agent = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )

    verified_url = ""
    extracted_info = ""

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage"
            ]
        )
        context = browser.new_context(
            user_agent=user_agent,
            viewport={"width": 1280, "height": 800},
            locale="en-US"
        )
        page = context.new_page()
        if stealth_sync:
            stealth_sync(page)

        # -------------------------------------------------------------
        # STEP 1: Search DuckDuckGo (HTML) to find the verified company URL
        # -------------------------------------------------------------
        try:
            query = f'site:linkedin.com/company "{company_name}"'
            search_url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(query)}"
            page.goto(search_url, timeout=25000, wait_until="domcontentloaded")

            content = page.content()
            soup = BeautifulSoup(content, "html.parser")

            for a in soup.find_all("a", href=True):
                href = a["href"]
                # DuckDuckGo wraps links in /l/?uddg=...
                if "uddg=" in href:
                    parsed = urllib.parse.parse_qs(urllib.parse.urlparse(href).query)
                    if "uddg" in parsed:
                        actual_link = parsed["uddg"][0]
                        if "linkedin.com/company/" in actual_link:
                            verified_url = actual_link
                            break
                elif "linkedin.com/company/" in href:
                    verified_url = href
                    break
        except Exception as e:
            print(f"      Search engine lookup warning: {e}")

        # Fallback to normalized slug if search fails
        if not verified_url:
            clean_name = re.sub(r"[^a-zA-Z0-9-]", "-", company_name.lower().strip())
            verified_url = f"https://www.linkedin.com/company/{clean_name}"

        # -------------------------------------------------------------
        # STEP 2: Visit the LinkedIn Page and Extract Full Public Metadata
        # -------------------------------------------------------------
        try:
            # Set extra headers to mimic typical desktop web navigation
            page.set_extra_http_headers({
                "Accept-Language": "en-US,en;q=0.9",
                "Referer": "https://www.google.com/"
            })
            response = page.goto(verified_url, timeout=30000, wait_until="domcontentloaded")

            if response and response.status in (200, 999):  # 999 is custom LinkedIn anti-bot code
                page_content = page.content()
                soup = BeautifulSoup(page_content, "html.parser")

                # 1. Try hidden structured JSON-LD (often has the full, un-truncated description)
                for block in soup.find_all("script", type="application/ld+json"):
                    try:
                        import json
                        data = json.loads(block.text)
                        if data.get("@type") == "Organization" or "numberOfEmployees" in data:
                            desc = data.get("description", "").strip()
                            employees = data.get("numberOfEmployees", {}).get("value", "Unknown size")
                            if desc:
                                extracted_info = f"Size: {employees} | {desc}"
                                break
                    except Exception:
                        continue

                # 2. Try OpenGraph / Meta tag fallback (take full content without slicing)
                if not extracted_info:
                    meta_tag = soup.find("meta", property="og:description") or soup.find("meta", attrs={"name": "description"})
                    if meta_tag and meta_tag.get("content"):
                        extracted_info = meta_tag["content"].strip()

                if not extracted_info:
                    extracted_info = "Profile loaded, but public bio was restricted."
            else:
                extracted_info = "Company page unavailable or blocked."

        except Exception as e:
            extracted_info = f"Playwright Error: {type(e).__name__}"
        finally:
            browser.close()

    # Clean the text: replace newlines (\n) and carriage returns (\r) with a space
    clean_info = " ".join(extracted_info.split())

    if verified_url and "Error" not in clean_info and "unavailable" not in clean_info:
        return verified_url, clean_info
        
    return "", clean_info

def draft_outreach(i):
    who = f" The sender works at {YOUR_COMPANY}." if YOUR_COMPANY else ""
    prompt = (
        'Return ONE JSON object (not a list), no other text, with two keys. '
        '"target_roles": a comma-separated string of exactly 3 job titles that would most likely lead or influence a GBS, '
        'shared services centre, global capability centre or COE at this company (for example Head of Global Business '
        'Services, Finance Director, HR Director). '
        '"message": a friendly, natural LinkedIn connection message of at most 280 characters from a salesperson. It should '
        'mention the news below, say they would like to connect and learn about the plans, and sound human. Use only facts '
        'from the news. Do not mention prices, do not claim anything about the recipient, do not use emojis, and start '
        'with "Hi [Name],".' + who + " "
        f"Company: {i['company']}. Signal: {i['signal']}. Location: {i['location']}. Headline: {i['title']}"
    )
    j = gemini_json(prompt)
    roles = j.get("target_roles", "")
    if isinstance(roles, list):
        roles = ", ".join(str(r) for r in roles)
    return str(roles).strip(), str(j.get("message", "")).strip()

def add_outreach(items, skip=()):
    """Items must already be sorted best score first. One draft per company."""
    if not GEMINI_KEY:
        print("No GEMINI_API_KEY set: skipping target roles and draft messages.")
        return
    seen, todo = set(), []
    for i in items:
        key = norm(i["company"])
        if i["score"] >= OUTREACH_MIN_SCORE and key and key not in seen and key not in skip:
            seen.add(key)
            todo.append(i)
    todo = todo[:MAX_OUTREACH_CALLS]
    print(f"\nDrafting target roles and LinkedIn messages for {len(todo)} companies...")
    fails = 0
    for i in todo:
        try:
            i["target_roles"], i["linkedin_message"] = draft_outreach(i)
            i["outreach_status"] = "New"
            fails = 0
        except PermissionError as e:
            print("   ", e)
            break
        except Exception as e:
            fails += 1
            print("   Could not draft for", i["company"], ":", e)
            if fails >= 3:
                print("   3 failures in a row. Stopping, please run again in a few minutes.")
                break
        time.sleep(5)

def load_keywords():
    if not os.path.exists(KEYWORDS_FILE):
        print(f"Keyword file not found ({KEYWORDS_FILE}). Using the built-in keyword list.")
        return KEYWORDS
    try:
        from openpyxl import load_workbook
        wb = load_workbook(KEYWORDS_FILE, read_only=True, data_only=True)
        ws = wb["Keywords"] if "Keywords" in wb.sheetnames else wb.active
        found = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            kw = str(row[0]).strip() if row and row[0] else ""
            active = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else "yes"
            if kw and active not in ("no", "n", "false", "0"):
                found.append(kw)
        wb.close()
        if not found:
            print("No active keywords in the Excel file. Using the built-in keyword list.")
            return KEYWORDS
        print(f"Loaded {len(found)} active keywords from {KEYWORDS_FILE}")
        return found
    except Exception as e:
        print("Could not read the keyword Excel file:", e, "\nUsing the built-in keyword list.")
        return KEYWORDS

def collect():
    items = []
    keywords = load_keywords()
    print("Collecting news...")
    for k in keywords:
        print(" ", k)
        try:
            items += fetch_gnews(k)
        except Exception as e:
            print("   Google News error:", e)
        if USE_BING:
            try:
                items += fetch_bing(k)
            except Exception as e:
                print("   Bing News error:", e)
    # Malaysian outlets through Google News site: filter
    for site in MALAYSIA_SITES:
        q = f'("shared services" OR "global business services" OR "global capability centre") site:{site}'
        try:
            items += fetch_gnews(q)
        except Exception as e:
            print("   Google News error:", e)
    for url in EXTRA_RSS_FEEDS:
        try:
            r = requests.get(url, headers=HEADERS, timeout=40)
            try:
                got = parse_rss(r.content, "RSS:" + urlparse(url).netloc, relevant_only=True)
            except ET.ParseError:
                print(f"   Extra feed not readable (HTTP {r.status_code}): {url}")
                print(f"   The reply starts with: {r.text[:100]!r}")
                continue
            print(f"  Extra feed {url}: {len(got)} relevant items")
            items += got
        except Exception as e:
            print("   Extra feed error:", url, e)
    counts = Counter(i["origin"] for i in items)
    print("Collected per source (before removing duplicates):", dict(counts))
    return items

def dedupe(items):
    seen_links, seen_titles, out = set(), set(), []
    for i in items:
        if not i["title"] or not i["link"]:
            continue
        if i["link"] in seen_links or norm(i["title"]) in seen_titles:
            continue
        seen_links.add(i["link"])
        seen_titles.add(norm(i["title"]))
        out.append(i)
    return out

def score_all(items):
    ai_calls = 0
    fails = 0
    for i in items:
        i.update({"company": "", "signal": "", "location": "", "score": rule_score(i["title"])})
    if not GEMINI_KEY:
        print("No GEMINI_API_KEY set: using simple rule-based scoring (company names stay blank).")
        return
    print("Scoring with Gemini...")
    # Highest rule-based scores first; the cap protects the free quota
    for i in sorted(items, key=lambda x: -x["score"]):
        if ai_calls >= MAX_AI_CALLS:
            break
        try:
            j = ai_score(i["title"])
            i.update({"company": clean_company(j.get("company")), "signal": j.get("signal", ""),
                      "location": j.get("location", ""), "score": min(5, max(1, int(float(j.get("score") or i["score"]))))})
            fails = 0
        except RuntimeError:
            print("   Gemini rate limit hit, stopping AI scoring. Rule-based scores kept for the rest.")
            break
        except PermissionError as e:
            print("   ", e)
            print("   Stopping AI scoring. Check the API key and model name.")
            break
        except Exception as e:
            fails += 1
            print("   AI error on one item, kept rule-based score:", e)
            if fails >= 3:
                print("   3 failures in a row. Stopping AI scoring, please run again in a few minutes.")
                break
        ai_calls += 1
        time.sleep(5)

COLS = [
    # 1. Lead Triage
    "score",
    "company",
    "location",
    "signal",
    # 2. Direct Outreach & Action
    "linkedin_url",
    "linkedin_search",
    "target_roles",
    "linkedin_message",
    "outreach_status",
    # 3. Source Evidence & Details
    "title",
    "link",
    "linkedin_description",
    "date",
    "first_seen",
    "source",
    "origin"
]

def load_history():
    """Read the existing results file, so old leads and the salesperson's notes are kept."""
    if not os.path.exists(OUTPUT_FILE):
        return [], True
    try:
        with open(OUTPUT_FILE, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except Exception as e:
        print("Could not read the existing results file:", e)
        return [], False
    # Excel can leave empty lines of commas behind after "select all + delete". Ignore rows with no title or link.
    rows = [r for r in rows if (r.get("title") or r.get("link") or "").strip()]
    for r in rows:
        try:
            r["score"] = int(float(r.get("score") or 1))
        except ValueError:
            r["score"] = 1
    return rows, True

def timestamped():
    base, ext = os.path.splitext(OUTPUT_FILE)
    return f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"

def main():
    history, history_ok = load_history()
    known_links = {r.get("link") for r in history}
    known_titles = {norm(r.get("title")) for r in history}
    found = dedupe(collect())
    items = [i for i in found if i["link"] not in known_links and norm(i["title"]) not in known_titles]
    print(f"\n{len(found)} unique articles found, {len(items)} are new since the last run "
          f"({len(history)} older rows kept).")
    if not items:
        print("Nothing new to add. The results file was not changed.")
        return
    today = f"{datetime.now():%Y-%m-%d}"
    for i in items:
        i["first_seen"] = today
    
    score_all(items)
    
    for i in items:
        i["linkedin_search"] = (
            "https://www.linkedin.com/search/results/companies/?keywords=" + urllib.parse.quote(i["company"])
            if i["company"] else "")
            
        # Call Playwright for high-quality leads
        if i.get("score", 1) >= OUTREACH_MIN_SCORE and i.get("company"):
            print(f"  Scraping LinkedIn info for {i['company']}...")
            url, desc = fetch_linkedin_company_data(i["company"])
            i["linkedin_url"] = url
            i["linkedin_description"] = desc
        else:
            i["linkedin_url"] = ""
            i["linkedin_description"] = ""
            
    items.sort(key=lambda x: (-x["score"], x["date"]), reverse=False)
    add_outreach(items, skip={norm(r.get("company")) for r in history if r.get("linkedin_message")})

    rows = items + history  # new rows first, older rows and their notes kept as they were
    # Best score first (5 down to 1). Python's sort keeps the existing order for equal scores,
    # so new rows stay above older rows with the same score.
    rows.sort(key=lambda x: -int(x.get("score") or 1))
    out_path = OUTPUT_FILE
    if os.path.exists(OUTPUT_FILE) and not history_ok:
        out_path = timestamped()  # never overwrite a file we could not read
        print("The existing results file could not be read, so it is left untouched.")
    try:
        f = open(out_path, "w", newline="", encoding="utf-8-sig")
    except PermissionError:  # the file is open in Excel
        out_path = timestamped()
        print(f"{OUTPUT_FILE} is open in Excel, so saving everything to {out_path} instead. Close Excel next time.")
        f = open(out_path, "w", newline="", encoding="utf-8-sig")
    with f:
        w = csv.DictWriter(f, fieldnames=COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    leads = [i for i in items if i["score"] >= MIN_LEAD_SCORE]
    seen, unique = {}, []
    for i in leads:  # one entry per company, best score first
        key = norm(i["company"]) or norm(i["title"])
        if key in seen:
            seen[key]["extra"] += 1
            continue
        i["extra"] = 0
        seen[key] = i
        unique.append(i)
    print(f"\nNEW LEADS THIS RUN (score >= {MIN_LEAD_SCORE}): {len(unique)} companies from {len(leads)} articles\n" + "-" * 60)
    for i in unique[:20]:
        more = f"  (+{i['extra']} more article(s))" if i["extra"] else ""
        print(f"[{i['score']}] {i['company'] or '(company not extracted)'} | {i['location']} | {i['date']}{more}")
        print(f"    {i['title']}\n    {i['link']}")
        if i.get("linkedin_company_info"):
            print(f"    LinkedIn Data: {i['linkedin_company_info']}")
        elif i["linkedin_search"]:
            print(f"    LinkedIn search: {i['linkedin_search']}")
        if i.get("linkedin_message"):
            print(f"    Roles to contact: {i['target_roles']}\n    Draft message: {i['linkedin_message']}")
        print()
    print(f"Saved to {out_path} ({len(rows)} rows in total)")

if __name__ == "__main__":
    main()