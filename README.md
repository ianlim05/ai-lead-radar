# AI Lead Radar

An automated Python pipeline designed to accelerate B2B sales research for Global Business Services (GBS), Shared Services, and Center of Excellence (COE) setups. 

Built specifically for non-technical sales teams, this tool acts as a "magic button" that hunts, grades, researches, and emails high-value leads entirely in the background.

## 🚀 Key Features

*   **Dynamic News Radar & Universal Bouncer:** Aggregates and deduplicates articles from Google News, Bing News, and custom RSS feeds. Features a hard-coded "Universal Bouncer" that instantly intercepts and deletes job boards, executive appointments, financial reports, and data centre/cloud region infrastructure news before they reach the AI.
*   **Excel-Driven Configuration:** Allows sales teams to control the exact search strategy, keywords, and active/inactive search terms directly from a simple `GBS Keywords.xlsx` file—zero Python coding required.
*   **Hybrid Stealth Scraping & Redirect Bypass:** Uses a smart Bing headline-lookup to successfully bypass Google News redirect walls. Employs `agent-reach` as a lightning-fast primary stealth reader to extract full article text, backed by a headless Playwright browser and PyPDF2 as robust safety nets for heavy anti-bot walls and PDF links.
*   **AI Scoring & Self-Cleaning Memory:** Uses the Gemini API to rapidly grade headlines (1-5) based on strict GBS/COE rules. Features a brilliant "Self-Cleaning Memory" that intentionally forgets rate-limited ("Not Scored") articles, automatically retrying them on subsequent runs to maximize free API quotas.
*   **Deep-Dive Executive Briefs:** For high-value leads (Score 4 & 5), the AI deeply analyzes up to 8,000 characters of scraped article text to extract a structured intelligence brief (Trigger Event, Scale, Key Players, and Angle).
*   **100% Safe LinkedIn Prospecting:** Completely eliminates anti-bot walls and account ban risks by dynamically generating targeted, reliable LinkedIn Keyword Search URLs for both the Company and the identified Executives.
*   **Automated Outlook Dispatch:** Wraps the final intelligence briefs in a clean HTML report and seamlessly dispatches it to the sales team by taking background control of the local Windows Classic Outlook app (`win32com.client`). It forces the correct sender address without requiring the Outlook window to be open.
*   **Zero-Friction Automation:** Ships with a custom `run.bat` file for instant manual execution and foolproof integration with Windows Task Scheduler. The tool runs invisibly in the background on a daily schedule, requiring zero clicks from the end user.
*   **Historical Tracking:** Exports data to a continuous CSV file, preserving previously discovered leads and historical salesperson notes across pipeline runs without ever sending duplicate leads.

---

## 🛠️ One-Time Computer Setup

### 1. Install Required Software
Open **Command Prompt** or **PowerShell** on your Windows machine and install the required dependencies:
```bash
pip install requests openpyxl playwright pywin32 python-dotenv PyPDF2
playwright install chromium
```

### 2. Configure Your Emails
Open `main.py` and update the email settings at the very top of the file:
```python
RECEIVER_EMAIL = "yourname@yourcompany.com" 
SENDER_EMAIL = "yourname@yourcompany.com" # Must match an account logged into Classic Outlook
```

### 3. Add Your API Keys
Create a file named exactly `.env` in the root folder and add your free API keys:
```text
GEMINI_API_KEY=your_gemini_api_key_here
TINYFISH_API_KEY=your_tinyfish_api_key_here
```

---

## 🎯 Daily Operation & Usage

### 1. Managing Your Search Words (`GBS Keywords.xlsx`)
You control what the radar searches for using the included Excel file. You do not need to edit code to change your search strategy.
*   **Exact Matches:** Wrap specific phrases in quotes (e.g., `"shared services centre"`).
*   **Active/Inactive:** Toggle keywords on or off using the active column in the spreadsheet.

### 2. Running the Radar Manually
Double-click the **`run.bat`** file in the project folder. The script will automatically open the terminal, run the pipeline, dispatch the email, and exit cleanly.

### 3. Full Automation (Windows Task Scheduler)
You can set the tool to run invisibly in the background every morning:
1. Open Windows **Task Scheduler** and click **Create Basic Task**.
2. Name it `GBS Lead Radar` and set the Trigger to **Daily** at your preferred time (e.g., 8:00 AM).
3. Set the Action to **Start a program**.
4. Browse and select the **`run.bat`** file in your project folder.
5. *(Optional)* To make it run invisibly, edit the task properties and check the **Hidden** box.

*Note: Your computer must be awake at the scheduled time. Classic Outlook does not need to be visibly open; the script will wake it up in the background.*

---

## ⚙️ Advanced Configuration (Master Control Panel)

At the very top of `main.py`, you can safely adjust the following global settings:

*   `DAYS = 7`: How far back the news radar looks. Change to 14 or 30 for broader historical searches.
*   `JUNK_TITLES` & `JUNK_URLS`: Add new spam words or blocked websites here to instantly reject them.
*   `COMPANY_ALIASES`: Auto-corrects common news abbreviations (e.g., `{"tp": "Teleperformance"}`).
*   `MALAYSIA_SITES`: A VIP list of local news websites the radar automatically searches to ensure you never miss a localized lead.

> **⚠️ API Safety Warning:** Do not alter `MAX_AI_CALLS = 200` or `MAX_DEEP_RESEARCH = 30`. These are strict safety limits designed to protect your free Google Gemini API quota from rate-limiting bans.