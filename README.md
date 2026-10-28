# AI Lead Radar

An automated Python pipeline designed to accelerate B2B sales research for Global Business Services (GBS), Shared Services, and Center of Excellence (COE) setups. 

**Key Features:**
* **News Monitoring & Redirect Bypass:** Aggregates and deduplicates articles from Google News, Bing News, and custom RSS feeds, utilizing a smart Bing headline-lookup to successfully bypass Google News redirect walls.
* **Hybrid Stealth Scraping:** Employs `agent-reach` as a lightning-fast primary stealth reader to extract full article text, backed by a headless Playwright browser as a robust safety net for heavy anti-bot walls.
* **AI Scoring & Executive Briefs:** Uses the Gemini API to filter noise by scoring headlines (1-5). For high-value leads, it deeply analyzes the full article text to extract a structured intelligence brief (Trigger Event, Scale, Key Players, and Angle).
* **100% Safe LinkedIn Prospecting:** Completely eliminates anti-bot walls and account ban risks by generating targeted, reliable LinkedIn Keyword Search URLs for both the Company and the identified Executives.
* **Historical Tracking:** Exports data to a continuous CSV file, preserving previously discovered leads and historical salesperson notes across pipeline runs.