#!/usr/bin/env python3
"""
NY State Jobs Automator  v3
============================
1. Ask for resume PDF(s) + optional transcript PDF
2. Scrape NY State Jobs search results with Selenium (real browser)
3. For each posting open the page, click every tab, extract all text
4. LLM decides: email-apply / portal-apply / fax / mail-only
5. LLM picks best resume and writes a tailored cover letter
6. Cover letter rendered as a professional PDF (Times New Roman, letterhead)
7. Per-job folder created with resume, cover letter PDF, transcript (if needed)
8. Email-apply  -> Outlook Web draft saved via Selenium with attachments
9. Non-email    -> HOW_TO_APPLY.txt + portal/fax/mail instructions saved
10. Master summary files written at the end

Requirements:
    pip install selenium webdriver-manager pymupdf reportlab requests beautifulsoup4 ollama

Ollama must be running:
    ollama pull llama3
    ollama serve
"""

import os, re, sys, json, time, shutil, argparse
from pathlib import Path
from datetime import datetime

# -- PyMuPDF ------------------------------------------------------------------
try:
    import fitz
except ImportError:
    sys.exit("Missing PyMuPDF.  Run:  pip install pymupdf")

# -- requests + bs4 -----------------------------------------------------------
try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("Missing requests/bs4.  Run:  pip install requests beautifulsoup4")

# -- ReportLab ----------------------------------------------------------------
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False
    print("[WARN] reportlab not found - cover letters will be .txt")
    print("       Fix: pip install reportlab")

# -- Ollama -------------------------------------------------------------------
try:
    import ollama as _ollama_lib
    OLLAMA_SDK = True
except ImportError:
    OLLAMA_SDK = False
    print("[WARN] ollama SDK not found - will use REST API.")

# -- Selenium -----------------------------------------------------------------
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False
    print("[WARN] selenium not found. Run: pip install selenium webdriver-manager")

# -- win32com (Outlook desktop app - Windows only) ----------------------------
try:
    import win32com.client as win32
    WIN32_OK = True
except ImportError:
    WIN32_OK = False
    # Not critical - will print reminder at runtime if email apply is found

# =============================================================================
#  CONFIG
# =============================================================================
SEARCH_URL = (
    "https://statejobsny.com/public/vacancyTable.cfm"
    "?searchResults=Yes&Keywords=computer+science"
    "&title=&JurisClassID=&AgID=&isnyhelp=&minDate=&maxDate="
    "&employmentType=&gradeCompareType=GT&grade=&SalMin="
)
VACANCY_BASE  = "https://statejobsny.com/public/vacancyDetailsView.cfm?id="
OLLAMA_MODEL  = "llama3"
OLLAMA_URL    = "http://localhost:11434/api/generate"
OUTPUT_ROOT   = Path("./NYJobs_Applications")
MAX_JOBS      = 100   # set to None for unlimited
VACANCY_TABS  = ["Basics", "Schedule", "Location", "Job Specifics", "How to Apply"]

# =============================================================================
#  SELENIUM DRIVERS
# =============================================================================

def make_headless_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1400,900")
    opts.add_argument("--log-level=3")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts
    )


# =============================================================================
#  PDF READING
# =============================================================================

def read_pdf_text(path: str) -> str:
    doc  = fitz.open(path)
    text = "\n".join(p.get_text() for p in doc)
    doc.close()
    return text.strip()

# =============================================================================
#  LLM
# =============================================================================

def llm(prompt: str, max_tokens: int = 1500) -> str:
    if OLLAMA_SDK:
        resp = _ollama_lib.generate(model=OLLAMA_MODEL, prompt=prompt)
        return resp["response"].strip()
    payload = {
        "model": OLLAMA_MODEL, "prompt": prompt,
        "stream": False, "options": {"num_predict": max_tokens},
    }
    r = requests.post(OLLAMA_URL, json=payload, timeout=180)
    r.raise_for_status()
    return r.json()["response"].strip()


def pick_best_resume(resumes: dict, job_title: str, job_desc: str) -> str:
    if len(resumes) == 1:
        return list(resumes.keys())[0]
    listing = "\n".join(f"RESUME [{n}]:\n{t[:1500]}\n---" for n, t in resumes.items())
    answer = llm(
        f"You are a hiring expert. Job: {job_title}\n\nDescription:\n{job_desc[:1500]}\n\n"
        f"Which resume filename is the BEST match? Reply with ONLY the filename.\n\n{listing}\n\nBEST:",
        max_tokens=60,
    ).strip()
    for name in resumes:
        if name in answer or answer in name:
            return name
    return list(resumes.keys())[0]


def detect_apply_method(how_to_apply_text: str) -> dict:
    """Ask LLM to parse How-to-Apply and return structured dict."""
    prompt = (
        "You are a job application assistant. Read the NY State job 'How to Apply' section below "
        "and extract the application method.\n\n"
        f"HOW TO APPLY:\n{how_to_apply_text[:2000]}\n\n"
        "Respond ONLY with a JSON object (no markdown, no backticks) with keys:\n"
        '  "method"  : "email" | "portal" | "fax" | "mail" | "unknown"\n'
        '  "email"   : email address string or null\n'
        '  "portal"  : full URL string or null\n'
        '  "fax"     : fax number string or null\n'
        '  "contact" : contact person name or null\n'
        '  "notes"   : 1-2 sentence plain-English summary\n\nJSON:'
    )
    raw = llm(prompt, max_tokens=350)
    raw = re.sub(r"```json|```", "", raw).strip()
    # isolate the first { ... } block
    m = re.search(r"\{.*\}", raw, re.S)
    if m:
        raw = m.group(0)
    try:
        result = json.loads(raw)
        for k in ("method", "email", "portal", "fax", "contact", "notes"):
            result.setdefault(k, None)
        if not result.get("notes"):
            result["notes"] = how_to_apply_text[:300]
        return result
    except Exception:
        # regex fallback
        email  = re.search(r"[\w._%+\-]+@[\w.\-]+\.[a-zA-Z]{2,}", how_to_apply_text)
        portal = re.search(r"https?://\S+", how_to_apply_text)
        fax    = re.search(r"[Ff]ax[:\s]+([0-9()\-\s]{7,20})", how_to_apply_text)
        method = ("email" if email else "portal" if portal else
                  "fax"   if fax   else "mail"   if re.search(r"\bmail|send|submit\b",
                                                               how_to_apply_text, re.I)
                  else "unknown")
        return {
            "method":  method,
            "email":   email.group(0)   if email  else None,
            "portal":  portal.group(0)  if portal else None,
            "fax":     fax.group(1).strip() if fax else None,
            "contact": None,
            "notes":   how_to_apply_text[:300],
        }


def write_cover_letter_text(resume_text, transcript_text, job_title,
                             agency, job_desc, apply_info, applicant_name="") -> str:
    transcript_bit = (
        f"\nApplicant Transcript (excerpt):\n{transcript_text[:600]}\n"
        if transcript_text else ""
    )
    notes = apply_info.get("notes", "")
    ref_match = re.search(r"[Rr]ef(?:erence)?\s*[Ii]tem\s*#?([\w\-]+)", notes)
    ref_item = (f"\n- IMPORTANT: Reference Item #{ref_match.group(1)} on the cover letter\n"
                if ref_match else "")

    prompt = (
        "You are a professional career coach writing a cover letter for a NY State government job.\n\n"
        f"JOB TITLE: {job_title}\nAGENCY: {agency}\n\n"
        f"JOB DESCRIPTION:\n{job_desc[:2500]}\n\n"
        f"APPLICANT RESUME:\n{resume_text[:2500]}\n"
        f"{transcript_bit}\n"
        f"SPECIAL INSTRUCTIONS:\n{notes[:400]}\n{ref_item}\n"
        "Rules:\n"
        "- Start with 'Dear Hiring Manager,'\n"
        "- 3-4 paragraphs, formal professional tone, NO bullet points\n"
        "- Highlight 3-4 specific skills from the resume that match the job\n"
        "- Mention the agency by name in the opening paragraph\n"
        "- Final paragraph expresses enthusiasm and requests an interview\n"
        f"- End with 'Sincerely,\\n\\n{applicant_name or '[YOUR NAME]'}'\n"
        "- Output ONLY the letter, no preamble or commentary\n\n"
        "COVER LETTER:"
    )
    return llm(prompt, max_tokens=950)

# =============================================================================
#  COVER LETTER -> PROFESSIONAL PDF
# =============================================================================

def render_cover_letter_pdf(cover_text, out_path, applicant_name,
                             job_title, agency, contact_name="", contact_email=""):
    """Times New Roman letterhead PDF."""
    if not REPORTLAB_OK:
        fallback = out_path.with_suffix(".txt")
        fallback.write_text(cover_text, encoding="utf-8")
        print(f"     [pdf] saved as .txt (install reportlab for PDF)")
        return

    doc = SimpleDocTemplate(
        str(out_path), pagesize=LETTER,
        leftMargin=1.1*inch, rightMargin=1.1*inch,
        topMargin=1.0*inch,  bottomMargin=1.0*inch,
    )

    TIMES      = "Times-Roman"
    TIMES_BOLD = "Times-Bold"
    TIMES_ITAL = "Times-Italic"
    BLACK      = colors.black
    DARK       = colors.HexColor("#1a1a1a")

    header_s = ParagraphStyle("Hdr",  fontName=TIMES_BOLD, fontSize=16,
                               leading=20, textColor=DARK, alignment=TA_CENTER, spaceAfter=2)
    date_s   = ParagraphStyle("Dt",   fontName=TIMES,      fontSize=11,
                               leading=14, textColor=BLACK, spaceBefore=14, spaceAfter=10)
    recip_s  = ParagraphStyle("Rec",  fontName=TIMES,      fontSize=11,
                               leading=15, textColor=BLACK, spaceAfter=14)
    re_s     = ParagraphStyle("Re",   fontName=TIMES_BOLD, fontSize=11,
                               leading=14, textColor=BLACK, spaceAfter=8)
    body_s   = ParagraphStyle("Body", fontName=TIMES,      fontSize=11,
                               leading=17, textColor=BLACK,
                               alignment=TA_JUSTIFY, spaceAfter=10)
    sig_s    = ParagraphStyle("Sig",  fontName=TIMES,      fontSize=11,
                               leading=15, textColor=BLACK, spaceBefore=18)

    story = []

    # Letterhead
    story.append(Paragraph(applicant_name or "Applicant", header_s))
    story.append(HRFlowable(width="100%", thickness=1,
                            color=colors.HexColor("#333333")))
    story.append(Spacer(1, 6))

    # Date
    story.append(Paragraph(datetime.now().strftime("%B %d, %Y"), date_s))

    # Recipient block
    rec_lines = []
    if contact_name:
        rec_lines.append(f"<b>{contact_name}</b>")
    rec_lines.append(agency)
    if contact_email:
        rec_lines.append(contact_email)
    story.append(Paragraph("<br/>".join(rec_lines), recip_s))

    # RE line
    story.append(Paragraph(f"<b>Re:</b> Application for {job_title}", re_s))
    story.append(Spacer(1, 2))

    # Body — find "Dear" start, split on blank lines
    body_text = cover_text.strip()
    m = re.search(r"Dear\b", body_text)
    if m:
        body_text = body_text[m.start():]

    in_sig = False
    for para in re.split(r"\n{2,}", body_text):
        para = para.strip()
        if not para:
            continue
        if re.match(r"^(Sincerely|Regards|Best regards|Yours truly|Respectfully)", para, re.I):
            in_sig = True
        style = sig_s if in_sig else body_s
        story.append(Paragraph(para.replace("\n", "<br/>"), style))

    doc.build(story)
    print(f"     [pdf] cover_letter.pdf saved")

# =============================================================================
#  WEB SCRAPING (Selenium)
# =============================================================================

def scrape_job_listings(max_jobs=MAX_JOBS) -> list:
    print("  [browser] Loading job listings...")
    driver   = make_headless_driver()
    jobs     = []
    page_url = SEARCH_URL

    try:
        while page_url:
            driver.get(page_url)
            time.sleep(2)

            # If the page uses DataTables, switch it to show all entries
            try:
                driver.execute_script("""
                    var sel = document.querySelector('select[name$="_length"]');
                    if (sel) {
                        var opt = document.createElement('option');
                        opt.value = -1; opt.text = 'All';
                        sel.appendChild(opt);
                        sel.value = -1;
                        sel.dispatchEvent(new Event('change'));
                    }
                """)
                time.sleep(2)
            except Exception:
                pass

            # Scroll to bottom to trigger any lazy-load
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)
            soup  = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.find("table")
            if not table:
                break

            for row in table.find_all("tr")[1:]:
                cols = row.find_all("td")
                if len(cols) < 6:
                    continue
                link   = cols[1].find("a")
                job_id = cols[0].get_text(strip=True)
                href   = (link["href"] if link and link.get("href") else "")

                if href.startswith("http"):
                    vurl = href
                elif href.startswith("/"):
                    vurl = "https://statejobsny.com" + href
                elif href:
                    vurl = "https://statejobsny.com/public/" + href
                else:
                    vurl = f"{VACANCY_BASE}{job_id}"

                jobs.append({
                    "id": job_id,
                    "title":    cols[1].get_text(strip=True),
                    "grade":    cols[2].get_text(strip=True),
                    "posted":   cols[3].get_text(strip=True),
                    "deadline": cols[4].get_text(strip=True),
                    "agency":   cols[5].get_text(strip=True),
                    "county":   cols[6].get_text(strip=True) if len(cols) > 6 else "",
                    "vacancy_url": vurl,
                })
                if max_jobs and len(jobs) >= max_jobs:
                    return jobs

            nxt = soup.find("a", string=re.compile(r"next", re.I))
            if nxt and nxt.get("href"):
                h = nxt["href"]
                page_url = (h if h.startswith("http")
                            else "https://statejobsny.com/public/" + h)
            else:
                page_url = None
    finally:
        driver.quit()

    return jobs


def scrape_vacancy(vacancy_url: str) -> dict:
    """
    Open vacancy page in headless Chrome.
    Click each tab, capture the rendered text per tab.
    """
    print(f"     URL: {vacancy_url}")
    driver = make_headless_driver()
    wait   = WebDriverWait(driver, 12)
    tab_data = {t: "" for t in VACANCY_TABS}

    try:
        driver.get(vacancy_url)
        time.sleep(2.5)

        for tab_label in VACANCY_TABS:
            try:
                # Click the tab
                tab_el = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    f"//*[self::button or self::a or self::li or self::span]"
                    f"[normalize-space(text())='{tab_label}']"
                )))
                driver.execute_script("arguments[0].scrollIntoView(true);", tab_el)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", tab_el)
                time.sleep(1.5)

                # Now grab ONLY the visible tab content area
                # Try multiple content container selectors in priority order
                grabbed = False
                for sel in [
                    # Most specific first
                    "//div[contains(@class,'tab-pane') and contains(@class,'active') and contains(@class,'show')]",
                    "//div[contains(@class,'tab-pane') and contains(@class,'active')]",
                    "//div[@id='tabContent' or contains(@class,'tabContent')]//div[contains(@class,'active')]",
                    "//div[@role='tabpanel' and not(contains(@class,'hidden'))]",
                    # Broader fallback
                    "//main//div[not(contains(@class,'hidden'))]",
                ]:
                    try:
                        els = driver.find_elements(By.XPATH, sel)
                        for el in els:
                            text = el.text.strip()
                            # Filter out navigation junk
                            if (text and 
                                len(text) > 100 and 
                                "Skip to Content" not in text[:50] and
                                "How to Get a State Job" not in text[:100]):
                                tab_data[tab_label] = text
                                grabbed = True
                                print(f"     [{tab_label}] {len(text)} chars captured")
                                break
                        if grabbed:
                            break
                    except Exception:
                        continue

                # Last resort: grab visible body but strip known nav elements
                if not grabbed:
                    try:
                        body_text = driver.find_element(By.TAG_NAME, "body").text.strip()
                        # Strip navigation cruft from the start
                        for cut_phrase in [
                            "Skip to Content",
                            "How to Get a State Job",
                            "Search Vacancies",
                            "Other State Listings",
                        ]:
                            if cut_phrase in body_text[:300]:
                                idx = body_text.find(cut_phrase)
                                # Find first real paragraph after nav (starts with capital letter + space)
                                match = re.search(r'[A-Z][a-z]', body_text[idx+100:])
                                if match:
                                    body_text = body_text[idx+100+match.start():]
                                    break
                        if len(body_text) > 200:
                            tab_data[tab_label] = body_text
                            print(f"     [{tab_label}] {len(body_text)} chars (fallback)")
                    except Exception:
                        pass

            except Exception as e:
                print(f"     [warn] Could not read tab '{tab_label}': {e}")
                pass

        # Grab full page text as ultimate fallback
        try:
            full_text = driver.find_element(By.TAG_NAME, "body").text.strip()
        except Exception:
            full_text = ""

    finally:
        try:
            driver.quit()
        except Exception:
            pass  # ignore quit timeout errors

    job_spec   = tab_data.get("Job Specifics", "") or full_text
    how_to_app = tab_data.get("How to Apply",  "") or ""

    return {
        "basics":        tab_data.get("Basics", ""),
        "job_specifics": job_spec,
        "how_to_apply":  how_to_app,
        "full_text":     full_text,
        "url":           vacancy_url,
    }

# =============================================================================
#  FILE HELPERS
# =============================================================================

def sanitize(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "", name)[:80].strip()


def job_folder(job: dict) -> Path:
    f = OUTPUT_ROOT / sanitize(f"{job['id']}_{job['title'][:50]}")
    f.mkdir(parents=True, exist_ok=True)
    return f


def copy_file(src: str, dest_dir: Path):
    dest = dest_dir / Path(src).name
    if not dest.exists():
        shutil.copy2(src, dest)


def save_how_to_apply_txt(folder, job, vacancy, apply_info):
    lines = [
        f"JOB:      {job['title']}",
        f"AGENCY:   {job['agency']}",
        f"ID:       {job['id']}",
        f"DEADLINE: {job['deadline']}",
        f"URL:      {vacancy['url']}",
        "", "=" * 56, "HOW TO APPLY", "=" * 56, "",
        f"METHOD:  {(apply_info.get('method') or 'unknown').upper()}",
    ]
    for k, label in [("email","EMAIL"), ("portal","PORTAL"),
                     ("fax","FAX"), ("contact","CONTACT")]:
        if apply_info.get(k):
            lines.append(f"{label}:   {apply_info[k]}")
    lines += [
        "", "NOTES:", apply_info.get("notes",""), "",
        "-" * 56, "FULL HOW-TO-APPLY TEXT:", "-" * 56, "",
        vacancy.get("how_to_apply",""),
    ]
    (folder / "HOW_TO_APPLY.txt").write_text("\n".join(lines), encoding="utf-8")


def save_job_summary(folder, job, vacancy, apply_info, chosen_resume):
    (folder / "job_summary.json").write_text(json.dumps({
        "generated_at":  datetime.now().isoformat(),
        "vacancy_id":    job["id"],
        "title":         job["title"],
        "agency":        job["agency"],
        "grade":         job["grade"],
        "deadline":      job["deadline"],
        "url":           vacancy["url"],
        "chosen_resume": chosen_resume,
        "apply_method":  apply_info.get("method"),
        "email":         apply_info.get("email"),
        "portal":        apply_info.get("portal"),
        "fax":           apply_info.get("fax"),
    }, indent=2), encoding="utf-8")


def transcript_required(job_specifics, how_to_apply) -> bool:
    combined = (job_specifics + how_to_apply).lower()
    return any(kw in combined for kw in
               ["transcript", "academic record", "degree verification", "college record"])

# =============================================================================
#  OUTLOOK WEB DRAFT
# =============================================================================

# =============================================================================
#  OUTLOOK DRAFT via win32com (uses your installed Outlook desktop app)
# =============================================================================

def save_outlook_draft(to_email, subject, body, attachments, sender_email="") -> bool:
    """
    Creates a draft in Outlook desktop app. If sender_email provided,
    tries to set it as the From account.
    """
    if not WIN32_OK:
        print("  [skip] pywin32 not installed. Run: pip install pywin32")
        return False

    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)  # 0 = olMailItem

        mail.To      = to_email
        mail.Subject = subject
        mail.Body    = body

        # Attach files
        for att_path in attachments:
            resolved = str(Path(att_path).resolve())
            if Path(resolved).exists():
                mail.Attachments.Add(resolved)
                print(f"     attached: {Path(resolved).name}")

        # Try to set sender account if specified
        if sender_email:
            try:
                # Access the NameSpace and find matching account
                namespace = outlook.GetNamespace("MAPI")
                accounts  = namespace.Accounts
                target_account = None
                
                for i in range(1, accounts.Count + 1):
                    acc = accounts.Item(i)
                    if acc.SmtpAddress.lower() == sender_email.lower():
                        target_account = acc
                        break
                
                if target_account:
                    # Set the SendUsingAccount before saving
                    mail.SendUsingAccount = target_account
                    print(f"     from: {sender_email}")
                else:
                    print(f"  [warn] Account '{sender_email}' not found in Outlook")
                    available = [accounts.Item(i).SmtpAddress 
                                 for i in range(1, accounts.Count + 1)]
                    print(f"         Available: {available}")
            except Exception as e:
                print(f"  [warn] Could not set sender account: {e}")
                print(f"         Draft will use default account")

        mail.Save()
        print(f"  [outlook] Draft saved -> {to_email}")
        return True

    except Exception as e:
        print(f"  [outlook ERROR] {e}")
        return False

# =============================================================================
#  INPUT COLLECTION
# =============================================================================

def collect_inputs():
    print("\n" + "=" * 60)
    print("  NY STATE JOBS AUTOMATOR  v3")
    print("=" * 60)

    applicant_name = input("\nYour full name (for cover letter header): ").strip()

    sender_email = input("Your email address (for Outlook drafts, e.g. jpatel@albany.edu): ").strip()

    print("\nStep 1 - Resume(s)")
    print("  Enter full path to each PDF. Press ENTER when done.\n")
    resumes_text, resume_paths = {}, {}
    while True:
        raw = input("  Resume path (ENTER to finish): ").strip().strip('"')
        if not raw:
            if not resumes_text:
                print("  [error] Need at least one resume.")
                continue
            break
        p = Path(raw)
        if not p.exists() or p.suffix.lower() != ".pdf":
            print(f"  [error] Not found / not PDF: {raw}")
            continue
        text = read_pdf_text(str(p))
        resumes_text[p.name] = text
        resume_paths[p.name] = str(p.resolve())
        print(f"  OK {p.name}  ({len(text)} chars)")

    print("\nStep 2 - Transcript (optional)")
    raw = input("  Transcript PDF path (ENTER to skip): ").strip().strip('"')
    transcript_path = None
    if raw:
        tp = Path(raw)
        if tp.exists() and tp.suffix.lower() == ".pdf":
            transcript_path = str(tp.resolve())
            print(f"  OK {tp.name}")
        else:
            print("  [warn] Not found - skipping.")

    return resumes_text, resume_paths, transcript_path, applicant_name, sender_email

# =============================================================================
#  MAIN
# =============================================================================

def main():
    global OLLAMA_MODEL

    parser = argparse.ArgumentParser(description="NY State Jobs Automator v3")
    parser.add_argument("--max-jobs",   type=int, default=MAX_JOBS)
    parser.add_argument("--no-outlook", action="store_true")
    parser.add_argument("--model",      default=OLLAMA_MODEL)
    args = parser.parse_args()
    OLLAMA_MODEL = args.model

    # Inputs
    resumes_text, resume_paths, transcript_path, applicant_name, sender_email = collect_inputs()
    transcript_text = read_pdf_text(transcript_path) if transcript_path else ""

    # Scrape listings
    print(f"\n[1/4] Scraping listings (max {args.max_jobs})...")
    jobs = scrape_job_listings(max_jobs=args.max_jobs)
    print(f"  -> {len(jobs)} postings.")

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    email_jobs, other_jobs = [], []

    # Process each job
    for i, job in enumerate(jobs, 1):
        print(f"\n[{i}/{len(jobs)}] {job['title']}")
        print(f"  Agency: {job['agency']}  |  Deadline: {job['deadline']}")

        try:
            # Scrape vacancy (all tabs)
            print("  -> scraping vacancy tabs...")
            vacancy = scrape_vacancy(job["vacancy_url"])

            # Show a preview of what was captured
            spec_preview = vacancy["job_specifics"][:200].replace("\n", " ").strip()
            apply_preview = vacancy["how_to_apply"][:150].replace("\n", " ").strip()
            print(f"     Job Specifics ({len(vacancy['job_specifics'])} chars): {spec_preview}...")
            print(f"     How to Apply  ({len(vacancy['how_to_apply'])} chars): {apply_preview}...")

            # Detect how to apply
            print("  -> detecting apply method...")
            how_text   = vacancy["how_to_apply"] or vacancy["full_text"]
            apply_info = detect_apply_method(how_text)
            print(f"     method={apply_info['method']}  "
                  f"email={apply_info.get('email')}  "
                  f"portal={apply_info.get('portal')}")

            # Pick best resume
            print("  -> picking best resume...")
            best_name = pick_best_resume(
                resumes_text, job["title"], vacancy["job_specifics"][:2000]
            )
            print(f"     -> {best_name}")

            # Generate cover letter
            print("  -> generating cover letter...")
            cl_text = write_cover_letter_text(
                resume_text     = resumes_text[best_name],
                transcript_text = transcript_text,
                job_title       = job["title"],
                agency          = job["agency"],
                job_desc        = vacancy["job_specifics"],
                apply_info      = apply_info,
                applicant_name  = applicant_name,
            )

            # Create folder + save everything
            folder = job_folder(job)
            print(f"  -> folder: {folder.name}")

            # Cover letter PDF
            cl_pdf = folder / "cover_letter.pdf"
            render_cover_letter_pdf(
                cover_text     = cl_text,
                out_path       = cl_pdf,
                applicant_name = applicant_name,
                job_title      = job["title"],
                agency         = job["agency"],
                contact_name   = apply_info.get("contact") or "",
                contact_email  = apply_info.get("email")   or "",
            )

            # Resume
            copy_file(resume_paths[best_name], folder)

            # Transcript if required
            needs_tr = transcript_required(
                vacancy["job_specifics"], vacancy["how_to_apply"]
            )
            if transcript_path and needs_tr:
                copy_file(transcript_path, folder)
                print("     transcript included (required by posting)")

            # Instructions text file
            save_how_to_apply_txt(folder, job, vacancy, apply_info)
            save_job_summary(folder, job, vacancy, apply_info, best_name)

            # Outlook draft
            method   = apply_info.get("method", "unknown")
            to_email = apply_info.get("email")

            if method == "email" and to_email and not args.no_outlook:
                print(f"  -> Outlook draft -> {to_email}")
                attachments = [resume_paths[best_name], str(cl_pdf)]
                if transcript_path and needs_tr:
                    attachments.append(transcript_path)
                ok = save_outlook_draft(
                    to_email     = to_email,
                    subject      = f"Application - {job['title']} (Vacancy {job['id']})",
                    body         = cl_text,
                    attachments  = attachments,
                    sender_email = sender_email,
                )
                email_jobs.append({**job, "email": to_email,
                                   "folder": str(folder), "draft_saved": ok})
            else:
                other_jobs.append({**job, "method": method,
                                   "apply_info": apply_info, "folder": str(folder)})

            print("  DONE")
            time.sleep(1)

        except Exception as e:
            import traceback
            print(f"  [ERROR] {e}")
            traceback.print_exc()
            continue

    # Master summary files
    print("\n[4/4] Writing summaries...")

    if email_jobs:
        lines = ["EMAIL-APPLY JOBS\n" + "=" * 56]
        for j in email_jobs:
            ok = "draft saved" if j.get("draft_saved") else "DRAFT FAILED"
            lines.append(
                f"\n{j['title']}\n  Agency:  {j['agency']}\n"
                f"  Email:   {j['email']}\n  Status:  {ok}\n  Folder:  {j['folder']}"
            )
        (OUTPUT_ROOT / "EMAIL_APPLY_JOBS.txt").write_text("\n".join(lines), encoding="utf-8")
        print(f"  EMAIL_APPLY_JOBS.txt  ({len(email_jobs)} jobs)")

    if other_jobs:
        lines = ["OTHER JOBS (portal / fax / mail)\n" + "=" * 56]
        lines.append("See HOW_TO_APPLY.txt in each folder.\n")
        for j in other_jobs:
            ai = j.get("apply_info", {})
            lines.append(
                f"\n{j['title']}\n  Agency:   {j['agency']}\n"
                f"  Method:   {j['method'].upper()}\n"
                + (f"  Portal:   {ai.get('portal')}\n" if ai.get("portal") else "")
                + (f"  Fax:      {ai.get('fax')}\n"    if ai.get("fax")    else "")
                + f"  Deadline: {j['deadline']}\n  Folder:   {j['folder']}"
            )
        (OUTPUT_ROOT / "OTHER_APPLY_JOBS.txt").write_text("\n".join(lines), encoding="utf-8")
        print(f"  OTHER_APPLY_JOBS.txt  ({len(other_jobs)} jobs)")

    print(f"\n{'=' * 60}")
    print(f"  COMPLETE  ->  {OUTPUT_ROOT.resolve()}")
    print(f"  Email: {len(email_jobs)}  |  Other: {len(other_jobs)}")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()