from flask import Flask, render_template, request, Response, session
import requests as req_lib
from bs4 import BeautifulSoup
import base64
import uuid
import urllib3

# Suppress SSL warnings since VTU cert chain is incomplete
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = "vtu-result-portal-secret-key"

# Base URLs for each semester (index pages)
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

# Store sessions per user (keyed by a session token)
user_sessions = {}


def get_or_create_sessions(token):
    """Create a requests.Session for each semester and fetch the index page to initialize cookies."""
    if token not in user_sessions:
        user_sessions[token] = {}
        for sem_key, index_url in SEM_INDEX_URLS.items():
            s = req_lib.Session()
            s.verify = False  # VTU SSL cert chain is incomplete
            try:
                s.get(index_url)  # initialize session cookies
            except Exception:
                pass
            user_sessions[token][sem_key] = s
    return user_sessions[token]


@app.route("/")
def index():
    # Generate a unique token for this user visit
    token = str(uuid.uuid4())
    session["token"] = token
    # Pre-create sessions to VTU for each semester
    get_or_create_sessions(token)
    return render_template("index.html")


@app.route("/captcha/<sem_key>")
def captcha_image(sem_key):
    """Proxy the VTU CAPTCHA image for a given semester, using the stored session."""
    token = session.get("token")
    if not token or token not in user_sessions or sem_key not in user_sessions.get(token, {}):
        return "Session expired", 400
    s = user_sessions[token][sem_key]
    resp = s.get(CAPTCHA_URL)
    return Response(resp.content, content_type=resp.headers.get("Content-Type", "image/png"))


@app.route("/refresh_captcha/<sem_key>")
def refresh_captcha(sem_key):
    """Refresh a single semester's CAPTCHA by re-fetching."""
    token = session.get("token")
    if not token or token not in user_sessions:
        return "Session expired", 400
    # Re-create session for this semester
    s = req_lib.Session()
    s.verify = False
    s.get(SEM_INDEX_URLS[sem_key])
    user_sessions[token][sem_key] = s
    resp = s.get(CAPTCHA_URL)
    return Response(resp.content, content_type=resp.headers.get("Content-Type", "image/png"))


@app.route("/submit", methods=["POST"])
def submit():
    token = session.get("token")
    if not token or token not in user_sessions:
        return "Session expired. Please go back and refresh.", 400

    usn = request.form["usn"]
    results = {}
    sessions = user_sessions[token]

    for i in range(1, 6):
        sem_key = f"sem{i}"
        captcha = request.form.get(f"captcha{i}", "")
        if not captcha:
            results[sem_key] = None
            continue

        s = sessions[sem_key]
        payload = {"lns": usn, "captchacode": captcha}
        response = s.post(SEM_RESULT_URLS[sem_key], data=payload)
        soup = BeautifulSoup(response.text, "html.parser")

        if "University Seat Number" not in soup.text:
            results[sem_key] = None
        else:
            results[sem_key] = response.text

    # Clean up stored sessions
    user_sessions.pop(token, None)

    return render_template("result.html", results=results)


if __name__ == "__main__":
    app.run(debug=True)
