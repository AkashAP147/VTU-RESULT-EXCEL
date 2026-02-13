from flask import Flask, render_template, request, Response, session
import requests as req_lib
from bs4 import BeautifulSoup
import uuid
import time
import urllib3

# Suppress SSL warnings since VTU cert chain is incomplete
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = "vtu-result-portal-secret-key"

# ==========================================
# SEMESTER URLS
# ==========================================
SEM_INDEX_URLS = {
    "sem1": "https://results.vtu.ac.in/DJcbcs24/index.php",
    "sem2": "https://results.vtu.ac.in/JJEcbcs24/index.php",
    "sem3": "https://results.vtu.ac.in/DJcbcs25/index.php",
    "sem4": "https://results.vtu.ac.in/JJEcbcs25/index.php",
    "sem5": "https://results.vtu.ac.in/D25J26Ecbcs/index.php",
}

SEM_RESULT_URLS = {
    "sem1": "https://results.vtu.ac.in/DJcbcs24/resultpage.php",
    "sem2": "https://results.vtu.ac.in/JJEcbcs24/resultpage.php",
    "sem3": "https://results.vtu.ac.in/DJcbcs25/resultpage.php",
    "sem4": "https://results.vtu.ac.in/JJEcbcs25/resultpage.php",
    "sem5": "https://results.vtu.ac.in/D25J26Ecbcs/resultpage.php",
}

CAPTCHA_URL = "https://results.vtu.ac.in/captcha/vtu_captcha.php"

# ==========================================
# SUBJECT MAPS PER SEMESTER
# ==========================================
SEM_SUBJECT_MAPS = {
    "1": {
        "MATHEMATICS FOR CSE STREAM-I": "MATHS",
        "PHYSICS FOR CSE STREAM": "PHY",
        "PRINCIPLES OF PROGRAMMING USING C": "C",
        "COMMUNICATIVE ENGLISH": "ENG",
        "INDIAN CONSTITUTION": "IC",
        "INNOVATION AND DESIGN THINKING": "IDT",
        "INTRODUCTION TO CIVIL ENGINEERING": "CIVIL",
        "RENEWABLE ENERGY SOURCES": "RES",
    },
    "2": {
        "MATHEMATICS-II FOR CSE STREAM": "MATHS2",
        "APPLIED CHEMISTRY FOR CSE STREAM": "CHEM",
        "COMPUTER-AIDED ENGINEERING DRAWING": "CAED",
        "PROFESSIONAL WRITING SKILLS IN ENGLISH": "PWSE",
        "SAMSKRUTIKA KANNADA": "SK",
        "SCIENTIFIC FOUNDATIONS OF HEALTH": "SFH",
        "INTRODUCTION TO PYTHON PROGRAMMING": "PY",
        "INTRODUCTION TO ELECTRONICS COMMUNICATION": "ELC",
    },
    "3": {
        "MATHEMATICS FOR COMPUTER SCIENCE": "M3",
        "DIGITAL DESIGN & COMPUTER ORGANIZATION": "DDCO",
        "OPERATING SYSTEMS": "OS",
        "DATA STRUCTURES AND APPLICATIONS": "DSA",
        "DATA STRUCTURES LAB": "DSL",
        "SOCIAL CONNECT AND RESPONSIBILITY": "SCR",
        "NATIONAL SERVICE SCHEME": "NSS",
        "DATA ANALYTICS WITH EXCEL": "DAE",
        "OBJECT ORIENTED PROGRAMMING WITH JAVA": "OOPJ",
    },
    "4": {
        "ANALYSIS & DESIGN OF ALGORITHMS": "ADA",
        "ARTIFICIAL INTELLIGENCE": "AI",
        "DATABASE MANAGEMENT SYSTEMS": "DBMS",
        "ANALYSIS & DESIGN OF ALGORITHMS LAB": "ADAL",
        "BIOLOGY FOR COMPUTER ENGINEERS": "BIO",
        "UNIVERSAL HUMAN VALUES COURSE": "UHV",
        "NATIONAL SERVICE SCHEME": "NSS",
        "DISCRETE MATHEMATICAL STRUCTURES": "DMS",
        "TECHNICAL WRITING USING LATEX LAB": "TWL",
    },
    "5": {
        "SOFTWARE ENGINEERING AND PROJECT MANAGEMENT": "SEPM",
        "COMPUTER NETWORKS": "CN",
        "THEORY OF COMPUTATION": "TOC",
        "DATA VISUALIZATION LAB": "DVL",
        "MINI PROJECT": "MINI",
        "RESEARCH METHODOLOGY AND IPR": "RMIPR",
        "ENVIRONMENTAL STUDIES AND E-WASTE MANAGEMENT": "EVS",
        "NATIONAL SERVICE SCHEME": "NSS",
        "UNIX SYSTEM PROGRAMMING": "UNIX",
    },
}

# Browser-like headers so VTU doesn't block our requests / CAPTCHAs
BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

# Store sessions per user (keyed by a session token)
user_sessions = {}
# Store hidden Token values per user per semester
user_tokens = {}


# ==========================================
# EXTRACT RESULT (with subject mapping)
# ==========================================
def extract_result(html, sem):
    """Parse VTU result HTML and return structured data."""
    soup = BeautifulSoup(html, "html.parser")

    if "University Seat Number" not in soup.text:
        return None

    usn = ""
    name = ""

    tables = soup.find_all("table")
    for table in tables:
        for row in table.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) >= 2:
                label = cols[0].text.strip()
                value = cols[1].text.strip()
                if "University Seat Number" in label:
                    usn = value.replace(":", "").strip()
                if "Student Name" in label:
                    name = value.replace(":", "").strip()

    subject_map = SEM_SUBJECT_MAPS.get(sem, {})
    subjects = []

    div_rows = soup.find_all("div", class_="divTableRow")
    for row in div_rows:
        cols = row.find_all("div", class_="divTableCell")
        if len(cols) == 7 and cols[0].text.strip() != "Subject Code":
            subject_code = cols[0].text.strip()
            subject_name = cols[1].text.strip().upper()
            internal = cols[2].text.strip()
            external = cols[3].text.strip()
            total = cols[4].text.strip()
            result = cols[5].text.strip()

            # Map subject to short name
            short = subject_name
            for full, s in subject_map.items():
                if full in subject_name:
                    short = s
                    break

            subjects.append({
                "code": subject_code,
                "name": subject_name,
                "short": short,
                "internal": internal,
                "external": external,
                "total": total,
                "result": result,
            })

    # Calculate totals
    total_marks = 0
    count = 0
    for s in subjects:
        try:
            total_marks += int(s["total"])
            count += 1
        except ValueError:
            pass

    percentage = round((total_marks / (count * 100)) * 100, 2) if count > 0 else 0

    return {
        "usn": usn,
        "name": name,
        "subjects": subjects,
        "total": total_marks,
        "percentage": percentage,
        "count": count,
    }


# ==========================================
# SESSION HELPERS
# ==========================================
def _create_session(index_url):
    """Create a requests.Session with browser headers, initialize VTU cookies,
    and extract the hidden Token from the index page."""
    s = req_lib.Session()
    s.verify = False
    s.headers.update(BROWSER_HEADERS)
    vtu_token = ""
    try:
        resp = s.get(index_url, headers={"Referer": "https://results.vtu.ac.in/"})
        soup = BeautifulSoup(resp.text, "html.parser")
        token_input = soup.find("input", {"name": "Token"})
        if token_input:
            vtu_token = token_input.get("value", "")
    except Exception:
        pass
    return s, vtu_token


def get_or_create_sessions(token):
    """Create a requests.Session for each semester and fetch the index page to initialize cookies."""
    if token not in user_sessions:
        user_sessions[token] = {}
        user_tokens[token] = {}
        for sem_key, index_url in SEM_INDEX_URLS.items():
            s, vtu_token = _create_session(index_url)
            user_sessions[token][sem_key] = s
            user_tokens[token][sem_key] = vtu_token
    return user_sessions[token]


# ==========================================
# ROUTES
# ==========================================
@app.route("/")
def index():
    token = str(uuid.uuid4())
    session["token"] = token
    get_or_create_sessions(token)
    return render_template("index.html")


@app.route("/captcha/<sem_key>")
def captcha_image(sem_key):
    """Proxy the VTU CAPTCHA image for a given semester."""
    token = session.get("token")
    if not token or token not in user_sessions or sem_key not in user_sessions.get(token, {}):
        return "Session expired", 400

    s = user_sessions[token][sem_key]
    index_url = SEM_INDEX_URLS.get(sem_key, "")
    captcha_with_ts = f"{CAPTCHA_URL}?_CAPTCHA&t={time.time()}"
    resp = s.get(captcha_with_ts, headers={"Referer": index_url})

    content_type = resp.headers.get("Content-Type", "image/png")
    return Response(resp.content, content_type=content_type,
                    headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


@app.route("/refresh_captcha/<sem_key>")
def refresh_captcha(sem_key):
    """Refresh a single semester's CAPTCHA by re-creating the session."""
    token = session.get("token")
    if not token or token not in user_sessions:
        return "Session expired", 400

    # Re-create session for this semester with fresh cookies
    index_url = SEM_INDEX_URLS[sem_key]
    s, vtu_token = _create_session(index_url)
    user_sessions[token][sem_key] = s
    user_tokens[token][sem_key] = vtu_token

    captcha_with_ts = f"{CAPTCHA_URL}?_CAPTCHA&t={time.time()}"
    resp = s.get(captcha_with_ts, headers={"Referer": index_url})
    content_type = resp.headers.get("Content-Type", "image/png")
    return Response(resp.content, content_type=content_type,
                    headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


@app.route("/submit", methods=["POST"])
def submit():
    token = session.get("token")
    if not token or token not in user_sessions:
        return "Session expired. Please go back and refresh.", 400

    usn = request.form["usn"].strip().upper()
    results = {}
    sessions = user_sessions[token]

    for i in range(1, 6):
        sem_key = f"sem{i}"
        captcha = request.form.get(f"captcha{i}", "").strip()
        if not captcha:
            results[sem_key] = None
            continue

        s = sessions[sem_key]
        vtu_token = user_tokens.get(token, {}).get(sem_key, "")
        payload = {"Token": vtu_token, "lns": usn, "captchacode": captcha}
        result_url = SEM_RESULT_URLS[sem_key]
        response = s.post(result_url, data=payload,
                          headers={"Referer": SEM_INDEX_URLS[sem_key]})

        parsed = extract_result(response.text, str(i))
        results[sem_key] = parsed

    # Clean up stored sessions
    user_sessions.pop(token, None)
    user_tokens.pop(token, None)

    return render_template("result.html", results=results, usn=usn)


if __name__ == "__main__":
    app.run(debug=True)
