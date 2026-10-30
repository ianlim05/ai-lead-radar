"""
GBS Lead Radar
Searches Google News + Bing News (Malaysia focus) and extra RSS feeds for companies planning a GBS / shared services / COE.
Scores leads 1-5, uses Playwright to extract public LinkedIn company metadata securely, and saves to results.csv.

Setup:  pip install requests openpyxl playwright
        playwright install chromium
Run:    python main.py
Optional AI scoring (free Gemini key from aistudio.google.com/apikey):
        set GEMINI_API_KEY=your_key      (Windows CMD)
        $env:GEMINI_API_KEY="your_key"   (PowerShell)
"""
import csv
print("Script is starting...")
import json
import os
from dotenv import load_dotenv
import urllib
load_dotenv()  # read GEMINI_API_KEY from .env
import re
import time
from collections import Counter
from datetime import datetime, timedelta, timezone
from urllib.parse import parse_qs, urlparse
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
import requests
from agent_reach.channels.web import WebChannel
from agent_reach.runtime.dispatch import run_action
import win32com.client

# ---------- SETTINGS (edit these) ----------
# Keywords come from this Excel file if it exists (sheet "Keywords": column A = Keyword, column B = Active Yes/No).
KEYWORDS_FILE = "GBS Keywords.xlsx"
OUTPUT_FILE = "results.csv"

# Malaysian outlets, searched through Google News with a site: filter
MALAYSIA_SITES = ["thestar.com.my", "theedgemalaysia.com", "nst.com.my", "malaymail.com", "bernama.com"]
DAYS = 7
MIN_LEAD_SCORE = 4  # Only output high-quality leads (4 or 5) to the terminal
MAX_AI_CALLS = 100
COMPANY_ALIASES = {"tp": "Teleperformance"}  # same company under different names in the news (lowercase key)
OUTREACH_MIN_SCORE = 4   # only leads at this score or higher get target roles and a draft message
MAX_DEEP_RESEARCH = 15  # max companies drafted per run (protects the free AI quota)
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

# Variables populated by the Excel file
TOPIC = []
ACTION = []

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}
RELEVANT = re.compile(r"shared services|global business services|global capability|capability cent|"
                      r"centre of excellence|center of excellence|\b(gbs|gcc|coe)\b", re.I)

COLS = ["score", "company", "location", "signal", "trigger_event", "scale", "key_players", "angle",
        "article_status", "outreach_status", "linkedin_company_search", "linkedin_people_search",
        "title", "link", "date", "first_seen", "source", "origin"]

def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())[:80]

def clean_company(name):
    name = (name or "").strip()
    return COMPANY_ALIASES.get(name.lower(), name)

def load_excel_config():
    """
    Reads Keywords directly from GBS Keywords.xlsx.
    Non-technical users can customize everything without touching Python code.
    """
    if not os.path.exists(KEYWORDS_FILE):
        print(f"Keyword file not found ({KEYWORDS_FILE}). Please create it.")
        return []

    try:
        from openpyxl import load_workbook
        wb = load_workbook(KEYWORDS_FILE, read_only=True, data_only=True)
        
        keywords = []
        ws_kw = wb["Keywords"] if "Keywords" in wb.sheetnames else wb.active
        
        for row in ws_kw.iter_rows(min_row=2, values_only=True):
            kw = str(row[0]).strip() if row and row[0] else ""
            active = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else "yes"
            
            if kw and active not in ("no", "n", "false", "0"):
                keywords.append(kw)
                
        wb.close()
        print(f"Loaded {len(keywords)} active keywords from Excel.")
        return keywords

    except ImportError:
        print("Error: Missing openpyxl. Please run 'pip install openpyxl' in your terminal.")
        return []
    except Exception as e:
        print(f"Could not read full Excel config: {e}.")
        return []

def collect():
    items = []
    keywords = load_excel_config()
    print("Collecting news...")
    for k in keywords:
        print(" ", k)
        try:
            items += fetch_gnews(k)
        except Exception as e:
            print("   Google News error:", e)
            
        try:
            items += fetch_tinyfish(k)
        except Exception as e:
            print("   TinyFish error:", e)
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

def fetch_tinyfish(query):
    """Hunts for corporate blogs and PDFs using TinyFish with strict noise filters."""
    api_key = os.getenv("TINYFISH_API_KEY")
    if not api_key:
        return [] 
        
    out = []
    try:
        # LAYER 1: Stricter API Query
        # Forces TinyFish to look for actual announcements and avoid general chatter
        advanced_query = f'{query} AND ("press release" OR announces OR launches OR expansion OR "new office")'
        
        # 1-month date filter to balance freshness and API quota
        thirty_days_ago = (datetime.now(timezone.utc) - timedelta(days=30)).strftime("%Y-%m-%d")
        
        r = requests.get(
            "https://api.search.tinyfish.ai", 
            headers={"X-API-Key": api_key},
            params={
                "query": advanced_query,
                "after_date": thirty_days_ago
            },
            timeout=30
        )
        if r.status_code != 200:
            print(f"   TinyFish API error: {r.status_code} - {r.text}")
            return []
            
        data = r.json()
        
        # LAYER 2: The Python Bouncer
        # If any of these words are in the URL or Title, throw it in the trash immediately
        junk_urls = ["job", "career", "workday", "taleo", "wiki", "glossary", "dictionary", "facebook.com", "instagram.com", "reddit.com", "outsourceaccelerator"]
        junk_titles = ["what is", "meaning", "explained", "vs ", "vs.", "guide", "hiring", "apply now"]
        
        for item in data.get("results", []):
            title = (item.get("title") or "").strip()
            link = (item.get("url") or "").strip()
            
            link_lower = link.lower()
            title_lower = title.lower()
            
            # Instantly block if a junk word is found
            if any(junk in link_lower for junk in junk_urls):
                continue
            if any(junk in title_lower for junk in junk_titles):
                continue
                
            if title and link:
                out.append({
                    "date": "", 
                    "title": title, 
                    "source": item.get("site_name", "TinyFish Corporate"), 
                    "link": link, 
                    "origin": "TinyFish"
                })
    except Exception as e:
        print(f"   TinyFish error: {e}")
        
    return out

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
            
            # LAYER 3: The Empty Company Assassin
            company_name = clean_company(j.get("company"))
            if not company_name or company_name.lower() in ["", "none", "n/a", "not stated", "unknown"]:
                print(f"   Skipping: Gemini found no company for {i['link']}")
                fails = 0
                continue  # Throws the article in the trash, does NOT save it to the CSV!
                
            # If it survives the assassin, it gets saved normally
            i.update({
                "company": company_name, 
                "signal": j.get("signal", ""),
                "location": j.get("location", ""), 
                "score": min(5, max(1, int(float(j.get("score") or i["score"]))))
            })
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

def rule_score(title):
    """Safety net fallback score. 0 indicates Gemini never scored the article."""
    return 0

def ai_score(title):
    prompt = (
        'From this news headline, return ONE JSON object (not a list), no other text: {"company": the organisation that will own the centre (if a vendor builds it for a client, give only the client\'s name), '
        '"signal": short phrase on whether the company is planning/setting up a GBS, shared services '
        'centre, global capability centre or COE, "location": string, "score": integer 1-5 for how '
        'strong a sales lead this is. Scoring rules: 5 = a named company clearly announces setting up or expanding its OWN GBS, '
        'shared services centre, global capability centre or COE; 3-4 = likely plans, or hiring or a site expansion for such a '
        'centre; 2 = a vendor or outsourcing deal, or a relevant company only mentioned in passing; 1 = executive '
        'appointments, reports, surveys, commentary, government news or anything unrelated. '
        'A "Centre of Excellence" counts only if it is an internal shared-services or operations centre; '
        "a technology vendor's own AI or innovation lab for customers scores 2.}. "
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

def add_outreach(items):
    """Items must already be sorted best score first."""
    if not GEMINI_KEY:
        print("No GEMINI_API_KEY set: skipping deep research.")
        return
    
    todo = []
    for i in items:
        if i["score"] >= OUTREACH_MIN_SCORE:
            todo.append(i)
            
    todo = todo[:MAX_DEEP_RESEARCH]
    print(f"\nResearching {len(todo)} companies (reading articles and building cheat sheets)...")
    fails = 0
    for i in todo:
        try:
            print(f"  Deep researching {i['company']}...")
            
            # 1. Use the new Bing lookup to resolve Google News redirects to real publisher URLs
            url = find_publisher_link(i["title"], i.get("company", "")) if "news.google.com" in i["link"] else i["link"]
            
            # 2. Try the primary Agent-Reach stealth reader
            article_text = fetch_article_text(url or i["link"])
            
            # 3. If it failed to get sufficient text, fallback to the standard Playwright browser reader
            if len(article_text) < 500:
                try:
                    final_url, article_text = fetch_article_text_playwright(url or i["link"])
                    print(f"    browser reader: {len(article_text)} chars from {final_url[:80]}")
                except Exception as e:
                    print(f"    browser reader error: {type(e).__name__}")
                    article_text = ""
            
            article_text = article_text[:8000]
            
            i["article_status"] = f"full text ({len(article_text)} chars)" if len(article_text) > 500 else "headline only"
            print(f"    article: {i['article_status']}")
            
            # Get the JSON dictionary from Gemini
            j_raw = generate_executive_brief(i, article_text)
            print("    [DEBUG] AI Output:", j_raw)  
            
            # Safely extract the strings
            i["trigger_event"] = str(j_raw.get("trigger_event", ""))
            i["scale"] = str(j_raw.get("scale", ""))
            
            # MUST call this to verify the names and set i["first_contact_name"] for the LinkedIn link
            apply_players(i, j_raw, article_text)
            
            if i["article_status"] == "headline only":
                i["angle"] = "Not enough detail (article could not be read)"
            else:
                i["angle"] = str(j_raw.get("angle", ""))
                
            # Generate the two new separated LinkedIn links
            comp_search, people_link = build_linkedin_link(i)
            i["linkedin_company_search"] = comp_search
            i["linkedin_people_search"] = people_link
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

def find_publisher_link(title, company=""):
    """Google News links cannot be read. Find the same story on Bing to get the real article address."""
    clean = re.sub(r"\s+-\s+[^-]+$", "", title)
    kw = [w for w in re.findall(r"[a-z0-9]+", clean.lower()) if len(w) > 3]
    words = set(kw)
    for query in (f'"{clean}"', clean, f"{company} {' '.join(kw[:6])}".strip()):
        try:
            r = requests.get("https://www.bing.com/news/search", params={"q": query, "format": "rss"},
                             headers=HEADERS, timeout=30)
            best, best_score = "", 0
            for it in parse_rss(r.content, "BingNews"):
                if "news.google.com" in it["link"]:
                    continue
                t_words = {w for w in re.findall(r"[a-z0-9]+", it["title"].lower()) if len(w) > 3}
                overlap = len(words & t_words) / max(1, len(words))
                if overlap > best_score:
                    best, best_score = it["link"], overlap
            if best_score >= 0.5:
                print(f"    publisher link found ({best_score:.0%} match): {best}")
                return best
        except Exception as e:
            print(f"    Bing lookup error: {e}")
    print("    no publisher link found on Bing")
    return ""

def fetch_article_text(url):
    text = ""
    print(f"    reading: {url}")
    try:
        out = run_action(WebChannel(), "read", {"target": url}, timeout=25.0)
        if out.get("ok") and out.get("items"):
            text = str(out["items"][0].get("text", ""))
        else:
            print("    Agent-Reach could not read it, handing off to browser...")
    except Exception as e:
        print(f"    Agent-Reach error: {e}")
        
    return text[:8000]

def fetch_article_text_playwright(url):
    """Open a page in a normal headless browser. Follows Google News redirects. Returns (final_url, text)."""
    from playwright.sync_api import sync_playwright
    final_url, text = url, ""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.goto(url, timeout=30000, wait_until="domcontentloaded")
            try:  # a Google News link redirects to the publisher; wait until we leave news.google.com
                page.wait_for_url(lambda u: "news.google.com" not in u, timeout=10000)
            except Exception:
                pass
            page.wait_for_timeout(1500)
            final_url = page.url
            el = page.query_selector("article") or page.query_selector("main") or page.query_selector("body")
            text = " ".join((el.inner_text() if el else "").split())
        finally:
            browser.close()
    return final_url, text

def generate_executive_brief(i, article_text):
    note = ""
    if len(article_text) < 500:
        note = "\nArticle text is NOT available. Use only the headline, write Not stated for scale, and use an empty list for key_players."
    prompt = (
        'Analyze the provided news article and return ONE JSON object with EXACTLY these four keys: '
        '"trigger_event" (1-sentence summary of the launch/expansion), '
        '"scale" (exact numbers for headcount/investment. If none, write "Not stated"), '
        '"key_players" (a list of objects like {"name": "...", "title": "..."} for each executive named or quoted in the article, or an empty list), '
        '"angle" (1 sentence that must cite a specific fact from the article, such as headcount, function or location; if there is none, write "Not enough detail").\n'
        'CRITICAL RULE: Do not invent names or numbers. If a number is not in the text, write "Not stated". '
        'If no executive is named, use an empty list.\n\n'
        f"Company: {i.get('company', '')}\nHeadline: {i.get('title', '')}\nArticle: {article_text}"
        f"{note}\n"
    )
    return gemini_json(prompt)

def apply_players(i, j, article_text):
    """Keep only executives that really appear in the article text."""
    players = j.get("key_players", [])
    if not isinstance(players, list):
        players = []
    players = [p for p in players if isinstance(p, dict) and str(p.get("name", "")).strip()]
    players = [p for p in players if p["name"].strip().lower() in article_text.lower()]
    i["key_players"] = "; ".join(f'{p["name"].strip()} ({str(p.get("title", "")).strip()})' for p in players) or "Not stated"
    i["first_contact_name"] = players[0]["name"].strip() if players else ""

def build_linkedin_link(i):
    company = i.get("company", "")
    name = i.get("first_contact_name", "")
    
    company_search = ""
    people_search = ""
    
    if company:
        # 1. LinkedIn Company Keyword Search (100% reliable)
        company_search = f"https://www.linkedin.com/search/results/companies/?keywords={urllib.parse.quote(company)}"
        
        # 2. LinkedIn People Search (Finds the executives)
        if name:
            people_search = f"https://www.linkedin.com/search/results/people/?keywords={urllib.parse.quote(f'{name} {company}')}"
        else:
            people_search = f"https://www.linkedin.com/search/results/people/?keywords={urllib.parse.quote(company)}"
            
    return company_search, people_search

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
            r["score"] = int(float(r.get("score") or 0))
        except ValueError:
            r["score"] = 0
    return rows, True

def timestamped():
    base, ext = os.path.splitext(OUTPUT_FILE)
    return f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"

def send_sales_email(leads):
    """Packages high-scoring leads into an HTML email and sends via local Outlook app."""
    if not leads:
        return

    print(f"\nPreparing email report for {len(leads)} leads...")
    
    # 1. Build the HTML email body
    html = f"<h2 style='font-family: Arial;'>New GBS Leads Found: {datetime.now():%Y-%m-%d}</h2>"
    for i in leads[:20]:
        company = i.get('company') or 'Unknown Company'
        html += f"""
        <div style="font-family: Arial; margin-bottom: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
            <h3 style="margin-top: 0; color: #2c3e50;">[{i['score']}] {company} | {i['location']}</h3>
            <p><strong>Trigger:</strong> {i.get('trigger_event', '')}</p>
            <p><strong>Scale:</strong> {i.get('scale', '')}</p>
            <p><strong>Key Players:</strong> {i.get('key_players', '')}</p>
            <p><strong>Angle:</strong> {i.get('angle', '')}</p>
            <p style="margin-top: 15px;">
                <a href="{i['link']}" style="background-color: #0073b1; color: white; padding: 5px 10px; text-decoration: none; border-radius: 3px; margin-right: 10px;">Read Article</a>
                <a href="{i.get('linkedin_company_search', '')}" style="background-color: #0073b1; color: white; padding: 5px 10px; text-decoration: none; border-radius: 3px; margin-right: 10px;">Company LinkedIn</a>
                <a href="{i.get('linkedin_people_search', '')}" style="background-color: #0073b1; color: white; padding: 5px 10px; text-decoration: none; border-radius: 3px;">People Search</a>
            </p>
        </div>
        """

    # 2. Connect to local Outlook and send
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # 0 represents a standard Mail Item
        
        mail.To = "ianlim@agosasia.com"  # Separate multiple emails with a semicolon (;) for Outlook
        mail.Subject = f"🚀 GBS Lead Radar: {len(leads)} New Qualified Leads"
        mail.HTMLBody = html
        
        # mail.Display()  # Optional: Un-comment this line if you want to preview the email before it sends
        mail.Send()
        print("   Email successfully sent to the sales team via local Outlook!")
        
    except Exception as e:
        print(f"   Failed to send email via Outlook: {e}")

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
    
    # 2. Call the updated outreach function
    add_outreach(items)

    rows = items + history  # new rows first, older rows and their notes kept as they were
    rows.sort(key=lambda x: -int(float(x.get("score") or 0)))
    
    # Format '0' scores as 'Not Scored' for the CSV output
    for r in rows:
        if str(r.get("score")) == "0":
            r["score"] = "Not Scored"
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
    
    # 3. Print the new 4-Column Executive Brief to the terminal
    for i in unique[:20]:
        display_score = "Not Scored" if str(i.get('score')) == "0" else i['score']
        more = f"  ({i['extra']} more article(s))" if i['extra'] else ""
        print(f"[{display_score}] {i['company'] or '(company not extracted)'} | {i['location']} | {i['date']}{more}")
        print(f"    {i['title']}\n    {i['link']}")
        if i.get("trigger_event"):
            print(f"    Article: {i.get('article_status', '')}")
            print(f"    Trigger: {i['trigger_event']}")
            print(f"    Scale: {i['scale']}")
            print(f"    Key Players: {i['key_players']}")
            print(f"    Angle: {i['angle']}")
        if i.get("linkedin_company_search"):
            print(f"    Company Search: {i['linkedin_company_search']}")
        if i.get("linkedin_people_search"):
            print(f"    People Search: {i['linkedin_people_search']}")
        print()
        
    print(f"Saved to {out_path} ({len(rows)} rows in total)")

    send_sales_email(unique)
    
if __name__ == "__main__":
    main()