"""
app.py  -  simple web form for the WhatsApp sender.

Run:   python app.py        (or double-click start.bat)
Open:  http://127.0.0.1:5000   (opens by itself)

The page lets you: add contacts (Excel/CSV upload or pasted numbers), write the message, pick attachments,
press "Send messages". The automation (wa_engine.py) then drives WhatsApp Web in Chrome on THIS computer.
"""
import csv
import hmac
import html
import io
import os
import re
import secrets
import threading
import time
import webbrowser
from datetime import datetime, timedelta

from flask import Flask, Response, jsonify, redirect, request, session

import wa_engine

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
wa_engine.set_data_dir(DATA_DIR)
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ============================================================================
#  LOGIN PIN  -  CHANGE THIS before using the tool.
#  Everyone who knows this PIN can use the tool and see its results.
#  (You can also set it without editing the file:  set WA_SENDER_PIN=your-pin  before running.)
# ============================================================================
DEFAULT_PIN = "246810"
APP_PIN = os.environ.get("WA_SENDER_PIN") or DEFAULT_PIN
MAX_TRIES = 5                  # wrong PINs allowed ...
LOCK_SECONDS = 300             # ... then that address is locked out for 5 minutes
SESSION_HOURS = 8              # login expires after this many hours
HOST = os.environ.get("WA_SENDER_HOST", "127.0.0.1")      # keep 127.0.0.1 unless hosting behind HTTPS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024


def _secret_key():
    """Random key that signs the login cookie. Kept in data/secret.key so logins survive a restart."""
    path = os.path.join(DATA_DIR, "secret.key")
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
    with open(path) as f:
        return f.read().strip()


app.secret_key = _secret_key()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_SECURE=os.environ.get("WA_SENDER_HTTPS") == "1",
                  PERMANENT_SESSION_LIFETIME=timedelta(hours=SESSION_HOURS))
FAILS = {}                     # address -> [wrong attempts, locked until (epoch seconds)]

LOCK = threading.Lock()
SILENT = {"invalid_format", "landline", "no_number", "duplicate", "skipped_already_messaged", "ready"}


def _new_job():
    return {"state": "idle", "id": "", "error": "", "stop": False, "started": "", "finished": "",
            "counts": {}, "rows": [], "log": [], "ready": 0, "will_send": 0, "limit": 0,
            "attachments": [], "message_chars": 0}


JOB = _new_job()


def _log(text):
    with LOCK:
        JOB["log"].append(f"{datetime.now().strftime('%H:%M:%S')}  {text}")
        del JOB["log"][:-400]


wa_engine.set_logger(_log)


def _emit(ev):
    with LOCK:
        if ev["event"] == "plan":
            JOB["ready"], JOB["will_send"] = ev["ready"], ev["will_send"]
            return
        status = ev["status"]
        JOB["counts"][status] = JOB["counts"].get(status, 0) + 1
        if status not in SILENT:
            JOB["rows"].append({"number": ev["number"], "name": ev["name"], "status": status, "detail": ev["detail"]})


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def safe_name(name):
    name = os.path.basename((name or "").replace("\\", "/"))
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or "file"


def _int(value, default, lo, hi):
    try:
        return max(lo, min(hi, int(float(value))))
    except (TypeError, ValueError):
        return default


def _read_contacts(form, files, folder):
    f = files.get("contacts_file")
    if f and f.filename:
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, safe_name(f.filename))
        f.save(path)
        return wa_engine.parse_contacts_file(path)
    text = (form.get("numbers_text") or "").strip()
    if text:
        return wa_engine.parse_pasted(text)
    raise ValueError("Add your contacts: upload an Excel/CSV file or paste phone numbers.")


def _options(form):
    min_d = _int(form.get("min_delay"), 60, 20, 3600)
    max_d = _int(form.get("max_delay"), 150, min_d, 7200)
    return {"limit": _int(form.get("limit"), 30, 1, 500), "min_delay": min_d, "max_delay": max_d,
            "skip_previous": form.get("skip_previous", "1") not in ("0", "false", "")}


# ----------------------------------------------------------------------------
# LOGIN (PIN)
# ----------------------------------------------------------------------------
LOGIN_PAGE = r"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>WhatsApp Sender - Login</title><style>
:root{--g:#128c7e;--g2:#25d366;--ink:#1c2b27;--mut:#5f716c;--line:#dfe7e4;--bg:#f2f6f4;--card:#fff}
@media (prefers-color-scheme:dark){:root{--ink:#e6efec;--mut:#9fb3ad;--line:#2b3a36;--bg:#111a17;--card:#18231f}}
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;background:var(--bg);color:var(--ink);font:15px/1.5 "Segoe UI",system-ui,sans-serif}
.box{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:28px 26px;width:min(360px,92vw);text-align:center}
.logo{width:52px;height:52px;border-radius:50%;background:var(--g);color:#fff;display:grid;place-items:center;font-size:26px;margin:0 auto 10px}
h1{font-size:19px;margin:0 0 4px}p{margin:0 0 16px;color:var(--mut);font-size:13.5px}
input{width:100%;padding:12px;border:1px solid var(--line);border-radius:10px;background:transparent;color:var(--ink);font:22px/1 Consolas,monospace;text-align:center;letter-spacing:6px}
button{margin-top:14px;width:100%;background:var(--g2);color:#06301f;border:0;border-radius:10px;padding:12px;font-size:16px;font-weight:700;cursor:pointer}
.err{margin-top:12px;color:#c0392b;font-size:13.5px;min-height:18px}
</style></head><body><form class="box" method="post" action="/login" autocomplete="off">
<div class="logo">&#128274;</div><h1>WhatsApp Sender</h1><p>Enter your PIN to continue</p>
<input type="password" name="pin" inputmode="numeric" autofocus required maxlength="64" aria-label="PIN">
<button type="submit">Unlock</button><div class="err">{{ERR}}</div></form></body></html>"""


@app.before_request
def guard():
    """Nothing except the login page works without a valid login."""
    if request.endpoint == "login":
        return None
    if not session.get("auth"):
        if request.path.startswith(("/api/", "/download/")):
            return jsonify(ok=False, error="Please log in."), 401
        return redirect("/login")
    # forged cross-site requests cannot add this header, our own page always does
    if request.method == "POST" and request.path.startswith("/api/") \
            and request.headers.get("X-Requested-With") != "wa-sender":
        return jsonify(ok=False, error="Blocked."), 403
    return None


@app.after_request
def no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"        # the browser's Back button must not show old data after logout
    return resp


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("auth"):
        return redirect("/")
    err, ip = "", request.remote_addr or "?"
    if request.method == "POST":
        with LOCK:
            wait = int(FAILS.get(ip, [0, 0.0])[1] - time.time())
        if wait > 0:
            err = f"Too many wrong attempts. Try again in {wait // 60 + 1} minute(s)."
        else:
            pin = (request.form.get("pin") or "").strip()
            if hmac.compare_digest(pin.encode("utf-8"), APP_PIN.encode("utf-8")):
                with LOCK:
                    FAILS.pop(ip, None)
                session.clear()
                session["auth"] = True
                session.permanent = True
                return redirect("/")
            time.sleep(0.7)                           # slows down guessing
            with LOCK:
                rec = FAILS.get(ip, [0, 0.0])
                rec[0] += 1
                left = MAX_TRIES - rec[0]
                if left <= 0:
                    rec = [0, time.time() + LOCK_SECONDS]
                FAILS[ip] = rec
            err = (f"Wrong PIN. {left} attempt(s) left." if left > 0
                   else f"Too many wrong attempts. Locked for {LOCK_SECONDS // 60} minutes.")
    return Response(LOGIN_PAGE.replace("{{ERR}}", html.escape(err)), mimetype="text/html")


@app.post("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ----------------------------------------------------------------------------
# API
# ----------------------------------------------------------------------------
@app.post("/api/preview")
def api_preview():
    """Check the list and show what the first message will look like. Sends nothing."""
    try:
        contacts = _read_contacts(request.form, request.files, os.path.join(UPLOAD_DIR, "_preview"))
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400
    if not contacts:
        return jsonify(ok=False, error="No phone numbers were found in that list."), 400
    done = wa_engine.previously_done()
    counts = {}
    for c in contacts:
        key = "already_messaged" if c["check"] == "ok" and c["number"] in done else c["check"]
        counts[key] = counts.get(key, 0) + 1
    first = next((c for c in contacts if c["check"] == "ok"), None)
    msg = (request.form.get("message") or "").strip()
    return jsonify(ok=True, total=len(contacts), counts=counts,
                   message_preview=wa_engine.render_message(msg, first) if (first and msg) else "",
                   sample=[{"number": c["number"], "name": c["name"], "check": c["check"]} for c in contacts[:8]])


@app.post("/api/start")
def api_start():
    with LOCK:
        if JOB["state"] in ("running", "stopping"):
            return jsonify(ok=False, error="A campaign is already running."), 409
    cid = datetime.now().strftime("%Y%m%d-%H%M%S")
    folder = os.path.join(UPLOAD_DIR, cid)
    try:
        contacts = _read_contacts(request.form, request.files, folder)
        message = (request.form.get("message") or "").strip()
        attachments, used = [], set()
        for f in request.files.getlist("attachments"):
            if not f or not f.filename:
                continue
            os.makedirs(folder, exist_ok=True)
            name = safe_name(f.filename)
            base, ext = os.path.splitext(name)
            n = 1
            while name.lower() in used:
                n += 1
                name = f"{base} ({n}){ext}"
            used.add(name.lower())
            path = os.path.join(folder, name)
            f.save(path)
            attachments.append(path)
        if not contacts:
            raise ValueError("No phone numbers were found in that list.")
        if not message and not attachments:
            raise ValueError("Write a message or add at least one attachment.")
        opts = _options(request.form)
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400

    with LOCK:
        JOB.clear()
        JOB.update(_new_job())
        JOB.update(state="running", id=cid, started=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                   limit=opts["limit"], attachments=[os.path.basename(a) for a in attachments],
                   message_chars=len(message))
    threading.Thread(target=_worker, args=(cid, contacts, message, attachments, opts), daemon=True).start()
    return jsonify(ok=True, id=cid, contacts=len(contacts))


def _worker(cid, contacts, message, attachments, opts):
    try:
        wa_engine.run_campaign(contacts, message, attachments, emit=_emit, campaign=cid,
                               should_stop=lambda: JOB["stop"], **opts)
        state, err = ("stopped" if JOB["stop"] else "done"), ""
    except Exception as e:
        state, err = "error", f"{type(e).__name__}: {str(e)[:300]}"
        _log("ERROR: " + err)
    with LOCK:
        JOB.update(state=state, error=err, finished=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


@app.post("/api/stop")
def api_stop():
    with LOCK:
        if JOB["state"] == "running":
            JOB["stop"] = True
            JOB["state"] = "stopping"
            JOB["log"].append(f"{datetime.now().strftime('%H:%M:%S')}  Stop requested - finishing the current contact...")
    return jsonify(ok=True)


@app.get("/api/status")
def api_status():
    with LOCK:
        return jsonify(state=JOB["state"], id=JOB["id"], error=JOB["error"], started=JOB["started"],
                       finished=JOB["finished"], counts=JOB["counts"], rows=JOB["rows"][-300:],
                       log=JOB["log"][-200:], ready=JOB["ready"], will_send=JOB["will_send"],
                       attachments=JOB["attachments"])


def _csv_response(rows, filename):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(wa_engine.LOG_COLS)
    for r in rows:
        w.writerow(r)
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={filename}"})


def _log_rows(campaign=None):
    if not os.path.exists(wa_engine.LOG_PATH):
        return []
    with open(wa_engine.LOG_PATH, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))[1:]
    return [r for r in rows if campaign is None or (len(r) > 6 and r[6] == campaign)]


@app.get("/download/results.csv")
def dl_results():
    return _csv_response(_log_rows(JOB["id"] or "-"), f"results_{JOB['id'] or 'none'}.csv")


@app.get("/download/log.csv")
def dl_log():
    return _csv_response(_log_rows(), "send_log_all.csv")


# ----------------------------------------------------------------------------
# PAGE
# ----------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>WhatsApp Sender</title>
<style>
:root{--g:#128c7e;--g2:#25d366;--ink:#1c2b27;--mut:#5f716c;--line:#dfe7e4;--bg:#f2f6f4;--card:#fff;--red:#c0392b;--amb:#b7791f}
@media (prefers-color-scheme:dark){:root{--ink:#e6efec;--mut:#9fb3ad;--line:#2b3a36;--bg:#111a17;--card:#18231f;--g:#2bb5a3}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Segoe UI",system-ui,sans-serif}
header{background:var(--g);color:#fff;padding:18px 20px;display:flex;justify-content:space-between;align-items:center;gap:12px}header .out{background:transparent;color:#fff;border:1px solid rgba(255,255,255,.6);border-radius:8px;padding:7px 14px;cursor:pointer;font-size:13.5px}header h1{margin:0;font-size:20px}header p{margin:2px 0 0;opacity:.9;font-size:13px}
main{max-width:860px;margin:18px auto;padding:0 14px 60px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:14px}
.card h2{margin:0 0 10px;font-size:16px}.num{display:inline-block;width:24px;height:24px;line-height:24px;text-align:center;border-radius:50%;background:var(--g);color:#fff;font-size:13px;margin-right:6px}
label{display:block;font-weight:600;margin:10px 0 4px}small,.hint{color:var(--mut);font-size:12.5px}
input[type=text],input[type=number],textarea{width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--ink);font:inherit}
textarea{min-height:170px;resize:vertical}input[type=file]{width:100%}
.tabs{display:flex;gap:6px;margin-bottom:8px}.tab{padding:6px 12px;border:1px solid var(--line);border-radius:20px;cursor:pointer;font-size:13px}
.tab.on{background:var(--g);color:#fff;border-color:var(--g)}
.row{display:flex;gap:10px;flex-wrap:wrap}.row>div{flex:1;min-width:140px}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin:6px 0}.chip{padding:3px 10px;border-radius:14px;border:1px solid var(--line);font-size:12.5px;cursor:pointer;user-select:none}
.files{margin-top:8px}.f{display:flex;align-items:center;justify-content:space-between;border:1px solid var(--line);border-radius:8px;padding:6px 10px;margin-top:6px;font-size:13.5px}
.f button{border:0;background:none;color:var(--red);cursor:pointer;font-size:16px}
.btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}
button.p{background:var(--g2);color:#06301f;border:0;border-radius:10px;padding:12px 22px;font-size:16px;font-weight:700;cursor:pointer}
button.s{background:transparent;color:var(--ink);border:1px solid var(--line);border-radius:10px;padding:12px 18px;font-size:15px;cursor:pointer}
button.d{background:var(--red);color:#fff;border:0;border-radius:10px;padding:10px 18px;font-weight:600;cursor:pointer}
button:disabled{opacity:.5;cursor:not-allowed}
.note{background:#fff8e6;border:1px solid #f0d58a;color:#6b4e00;padding:9px 12px;border-radius:8px;font-size:13px;margin-top:10px}
@media (prefers-color-scheme:dark){.note{background:#2c2610;border-color:#6b5a1a;color:#f1d98a}}
.err{background:#fdecea;border:1px solid #f1a9a0;color:#8a1f12;padding:9px 12px;border-radius:8px;margin-top:10px;display:none}
.prev{margin-top:10px;display:none;border:1px dashed var(--line);border-radius:8px;padding:10px 12px}
.bub{white-space:pre-wrap;background:#dcf8c6;color:#1c2b27;border-radius:8px;padding:8px 10px;margin-top:6px;font-size:14px}
.stats{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.stat{border:1px solid var(--line);border-radius:10px;padding:6px 12px;text-align:center;min-width:92px}
.stat b{display:block;font-size:20px}.stat span{font-size:12px;color:var(--mut)}
.bar{height:10px;background:var(--line);border-radius:6px;overflow:hidden}.bar i{display:block;height:100%;width:0;background:var(--g2);transition:width .4s}
.badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:700;background:var(--line)}
.badge.running,.badge.stopping{background:#d6f5e4;color:#0b5a37}.badge.done{background:#d9e8ff;color:#123a7a}.badge.error{background:#fdecea;color:#8a1f12}.badge.stopped{background:#fff0cc;color:#6b4e00}
pre#log{background:#0f1714;color:#c9f1dc;border-radius:8px;padding:10px;height:170px;overflow:auto;font:12.5px/1.45 Consolas,monospace;white-space:pre-wrap;margin:10px 0}
table{width:100%;border-collapse:collapse;font-size:13.5px}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
.tw{max-height:300px;overflow:auto;border:1px solid var(--line);border-radius:8px}
.st-sent{color:#0b8a4b;font-weight:600}.st-warn{color:var(--amb);font-weight:600}.st-bad{color:var(--red);font-weight:600}
a.lnk{color:var(--g);font-weight:600;text-decoration:none;margin-right:14px}
</style></head><body>
<header><div><h1>WhatsApp Sender</h1><p>Send a message and files to your contact list through WhatsApp Web</p></div><form method="post" action="/logout"><button class="out">Log out</button></form></header>
<main>

<div class="card" id="formCard">
 <h2><span class="num">1</span>Contacts</h2>
 <div class="tabs"><span class="tab on" id="tabFile">Upload Excel / CSV</span><span class="tab" id="tabPaste">Paste numbers</span></div>
 <div id="boxFile"><input type="file" id="contacts" accept=".xlsx,.xls,.csv,.txt">
  <div class="hint">Needs a column called <b>WhatsApp Number</b> (or Mobile / Phone). Optional columns: <b>College Name</b>, <b>Contact Name</b>.</div></div>
 <div id="boxPaste" style="display:none"><textarea id="numbers" style="min-height:110px" placeholder="919876543210, Name of college&#10;9876543211&#10;+91 98765 43212, Another college"></textarea>
  <div class="hint">One per line: number, then an optional name. 10-digit Indian numbers get +91 added automatically.</div></div>

 <h2 style="margin-top:20px"><span class="num">2</span>Message</h2>
 <div class="chips"><span class="hint" style="align-self:center">Insert:</span>
  <span class="chip" data-t="{name}">{name} college / organisation</span><span class="chip" data-t="{contact}">{contact} contact person</span></div>
 <textarea id="message" placeholder="Hello {name},&#10;&#10;Write your message here..."></textarea>
 <small><span id="cnt">0</span> characters &middot; leave empty to send only attachments</small>

 <h2 style="margin-top:20px"><span class="num">3</span>Attachments <small>(optional)</small></h2>
 <input type="file" id="atts" multiple>
 <div class="files" id="fileList"></div>
 <div class="hint">PDF, images, documents, videos. Each file is sent as its own message after the text.</div>

 <h2 style="margin-top:20px"><span class="num">4</span>Options</h2>
 <div class="row">
  <div><label for="limit">Max contacts this run</label><input type="number" id="limit" value="30" min="1" max="500"></div>
  <div><label for="mind">Wait between contacts (sec)</label><div class="row"><input type="number" id="mind" value="60" min="20"><input type="number" id="maxd" value="150" min="20"></div></div>
 </div>
 <label style="font-weight:400;margin-top:10px"><input type="checkbox" id="skip" checked> Skip numbers I have already messaged before</label>
 <div class="note">&#9888; WhatsApp can ban numbers that send many messages to people who haven't contacted them. Keep batches small (about 30 a day) and the waits long.</div>

 <div class="err" id="err"></div>
 <div class="btns"><button class="s" id="btnCheck">Check my list</button><button class="p" id="btnSend">Send messages</button></div>
 <div class="prev" id="prev"></div>
</div>

<div class="card" id="runCard" style="display:none">
 <h2>Progress <span class="badge" id="badge">idle</span></h2>
 <div class="note" id="runNote" style="display:none">Chrome opens WhatsApp Web. If a QR code appears, scan it with your phone (WhatsApp &gt; Linked devices). Don't close that Chrome window and keep this page open.</div>
 <div class="bar" style="margin-top:10px"><i id="barfill"></i></div><small id="barText"></small>
 <div class="stats" id="stats"></div>
 <pre id="log"></pre>
 <div class="tw"><table><thead><tr><th>Number</th><th>Name</th><th>Result</th></tr></thead><tbody id="rows"></tbody></table></div>
 <div class="btns"><button class="d" id="btnStop" style="display:none">Stop after current contact</button>
  <a class="lnk" href="/download/results.csv">Download this run (CSV)</a><a class="lnk" href="/download/log.csv">Download full history</a></div>
</div>
</main>

<script>
const $=s=>document.querySelector(s);let files=[],mode='file',busy=false;
async function api(u,o){o=o||{};o.headers=Object.assign({'X-Requested-With':'wa-sender'},o.headers||{});o.credentials='same-origin';const r=await fetch(u,o);if(r.status===401){location='/login';throw new Error('login required')}return r}
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const LABEL={sent:'Sent',sent_unverified:'Sent (check phone)',text_sent_attachment_failed:'Text sent, attachment failed',not_on_whatsapp:'Not on WhatsApp',failed:'Failed',queued:'Queued'};
const CLS={sent:'st-sent',sent_unverified:'st-warn',text_sent_attachment_failed:'st-warn',not_on_whatsapp:'st-warn',failed:'st-bad'};
$('#tabFile').onclick=()=>setMode('file');$('#tabPaste').onclick=()=>setMode('paste');
function setMode(m){mode=m;$('#tabFile').classList.toggle('on',m=='file');$('#tabPaste').classList.toggle('on',m=='paste');
 $('#boxFile').style.display=m=='file'?'':'none';$('#boxPaste').style.display=m=='paste'?'':'none'}
$('#message').oninput=()=>$('#cnt').textContent=$('#message').value.length;
document.querySelectorAll('.chip').forEach(c=>c.onclick=()=>{const t=$('#message'),s=t.selectionStart;
 t.value=t.value.slice(0,s)+c.dataset.t+t.value.slice(t.selectionEnd);t.focus();t.selectionStart=t.selectionEnd=s+c.dataset.t.length;$('#cnt').textContent=t.value.length});
$('#atts').onchange=e=>{for(const f of e.target.files)if(!files.some(x=>x.name==f.name&&x.size==f.size))files.push(f);e.target.value='';drawFiles()};
function size(n){return n>1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round(n/1024))+' KB'}
function drawFiles(){$('#fileList').innerHTML=files.map((f,i)=>`<div class="f"><span>${esc(f.name)} <small>(${size(f.size)})</small></span><button title="Remove" onclick="files.splice(${i},1);drawFiles()">&times;</button></div>`).join('')}
function showErr(t){const e=$('#err');e.textContent=t||'';e.style.display=t?'block':'none'}
function build(withFiles){const fd=new FormData();fd.append('message',$('#message').value);
 if(mode=='file'&&$('#contacts').files[0])fd.append('contacts_file',$('#contacts').files[0]);
 if(mode=='paste')fd.append('numbers_text',$('#numbers').value);
 fd.append('limit',$('#limit').value);fd.append('min_delay',$('#mind').value);fd.append('max_delay',$('#maxd').value);fd.append('skip_previous',$('#skip').checked?'1':'0');
 if(withFiles)files.forEach(f=>fd.append('attachments',f,f.name));return fd}
$('#btnCheck').onclick=async()=>{showErr('');$('#prev').style.display='none';
 const r=await api('/api/preview',{method:'POST',body:build(false)});const j=await r.json();
 if(!j.ok){showErr(j.error);return}
 const c=j.counts,ok=c.ok||0;let h=`<b>${j.total}</b> rows found &middot; <b style="color:#0b8a4b">${ok} ready to message</b>`;
 const bad=[['already_messaged','already messaged before'],['duplicate','duplicates'],['invalid_format','invalid numbers'],['landline','landlines (no WhatsApp)'],['no_number','empty']].filter(x=>c[x[0]]).map(x=>`${c[x[0]]} ${x[1]}`);
 if(bad.length)h+=` &middot; skipped: ${bad.join(', ')}`;
 if(j.message_preview)h+=`<div class="bub">${esc(j.message_preview)}</div><small>Preview for the first contact</small>`;
 $('#prev').innerHTML=h;$('#prev').style.display='block'};
$('#btnSend').onclick=async()=>{showErr('');
 const chk=await api('/api/preview',{method:'POST',body:build(false)});const cj=await chk.json();
 if(!cj.ok){showErr(cj.error);return}
 if(!$('#message').value.trim()&&!files.length){showErr('Write a message or add at least one attachment.');return}
 const n=Math.min(cj.counts.ok||0,+$('#limit').value||30);
 if(!(cj.counts.ok)){showErr('No sendable numbers in this list (check the numbers / "already messaged" skips).');return}
 if(!confirm(`Send to up to ${n} contact(s)?\n\nChrome will open WhatsApp Web and send the messages one by one.`))return;
 const r=await api('/api/start',{method:'POST',body:build(true)});const j=await r.json();
 if(!j.ok){showErr(j.error);return}poll()};
$('#btnStop').onclick=async()=>{$('#btnStop').disabled=true;await api('/api/stop',{method:'POST'});poll()};
async function poll(){let s;try{s=await (await api('/api/status')).json()}catch(e){return}
 const run=s.state=='running'||s.state=='stopping';busy=run;
 $('#runCard').style.display=s.state=='idle'?'none':'';
 $('#btnSend').disabled=run;$('#btnCheck').disabled=run;
 const b=$('#badge');b.textContent=({running:'running',stopping:'stopping...',done:'finished',stopped:'stopped',error:'error'})[s.state]||s.state;b.className='badge '+s.state;
 $('#runNote').style.display=run?'':'none';$('#btnStop').style.display=run?'':'none';$('#btnStop').disabled=s.state=='stopping';
 const c=s.counts||{},sum=(...k)=>k.reduce((a,x)=>a+(c[x]||0),0);
 const attempted=sum('sent','sent_unverified','text_sent_attachment_failed','not_on_whatsapp','failed');
 const pct=s.will_send?Math.min(100,Math.round(attempted*100/s.will_send)):(run?0:100);
 $('#barfill').style.width=pct+'%';$('#barText').textContent=s.will_send?`${attempted} of ${s.will_send} contacts processed`:(run?'Preparing...':'');
 const chips=[['Sent',sum('sent','sent_unverified')],['Partly sent',sum('text_sent_attachment_failed')],['Not on WhatsApp',sum('not_on_whatsapp')],['Failed',sum('failed')],
  ['Skipped',sum('skipped_already_messaged','duplicate')],['Invalid',sum('invalid_format','landline','no_number')]];
 $('#stats').innerHTML=chips.map(x=>`<div class="stat"><b>${x[1]}</b><span>${x[0]}</span></div>`).join('');
 const lg=$('#log'),stick=lg.scrollTop+lg.clientHeight>=lg.scrollHeight-30;lg.textContent=(s.log||[]).join('\n')+(s.error?'\n'+s.error:'');if(stick)lg.scrollTop=lg.scrollHeight;
 $('#rows').innerHTML=(s.rows||[]).slice().reverse().map(r=>`<tr><td>${esc(r.number)}</td><td>${esc(r.name)}</td><td class="${CLS[r.status]||''}">${esc(LABEL[r.status]||r.status)}${r.detail?`<br><small>${esc(r.detail)}</small>`:''}</td></tr>`).join('')}
poll();setInterval(poll,1500);
</script></body></html>"""


@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html")


def main():
    host, port = HOST, 5000
    if len(APP_PIN) < 4:
        raise SystemExit("The PIN must be at least 4 characters. Edit APP_PIN in app.py.")
    if APP_PIN == DEFAULT_PIN:
        print("\n*** WARNING: you are still using the default PIN. Change APP_PIN in app.py (or set WA_SENDER_PIN). ***")
    print(f"\nWhatsApp Sender is running at http://{host}:{port}\nKeep this window open. Press Ctrl+C to quit.\n")
    threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host=host, port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
