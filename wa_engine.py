"""
wa_engine.py  -  WhatsApp Web sender (Selenium). WhatsApp only - no email code.

Used by app.py (the web form), but can also be imported on its own:
    import wa_engine
    contacts = wa_engine.parse_contacts_file("list.xlsx")
    wa_engine.run_campaign(contacts, "Hello {name}", ["one_pager.pdf"])
"""
import os
import random
import re
import subprocess
import time
import urllib.parse
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (NoAlertPresentException, TimeoutException,
                                        UnexpectedAlertPresentException)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

# ----------------------------------------------------------------------------
# SETTINGS / STATE
# ----------------------------------------------------------------------------
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
BASE_DIR = DATA_DIR                                        # debug screenshots go here
CHROME_PROFILE = os.path.join(DATA_DIR, "chrome_profile")  # keeps the WhatsApp login (scan the QR only once)
LOG_PATH = os.path.join(DATA_DIR, "send_log.csv")
PAGE_TIMEOUT = 40
CURRENT_FILE = ""                                          # the attachment being sent right now
_LOG_CB = print


def set_data_dir(path):
    global DATA_DIR, BASE_DIR, CHROME_PROFILE, LOG_PATH
    DATA_DIR = BASE_DIR = path
    CHROME_PROFILE = os.path.join(path, "chrome_profile")
    LOG_PATH = os.path.join(path, "send_log.csv")
    os.makedirs(path, exist_ok=True)


def set_logger(fn):
    global _LOG_CB
    _LOG_CB = fn


def say(text):
    _LOG_CB(text)


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _upload_extra():
    """Bigger files need longer to upload before we may leave the chat."""
    try:
        return min(90, os.path.getsize(CURRENT_FILE) / (1024 * 1024) * 3)
    except Exception:
        return 0


# ----------------------------------------------------------------------------
# CONTACTS: numbers, files, pasted text, message templates
# ----------------------------------------------------------------------------
MOBILE_RE = re.compile(r"^91[6-9]\d{9}$")


def normalize_number(raw, number_type=""):
    """-> (digits for WhatsApp, check) where check is ok | no_number | landline | invalid_format."""
    d = re.sub(r"\D", "", str(raw or ""))
    if not d:
        return "", "no_number"
    if str(number_type).strip().lower() == "landline":
        return d, "landline"
    if len(d) == 10 and d[0] in "6789":                    # Indian mobile typed without country code
        return "91" + d, "ok"
    if len(d) == 11 and d[0] == "0" and d[1] in "6789":
        return "91" + d[1:], "ok"
    if d.startswith("91") and len(d) == 12:
        return (d, "ok") if MOBILE_RE.match(d) else (d, "invalid_format")
    if len(d) == 10:
        return d, "landline"                               # 10 digits not starting 6-9: landline / not a mobile
    if 11 <= len(d) <= 15 and not d.startswith("0"):
        return d, "ok"                                     # other countries, already with country code
    return d, "invalid_format"


def _pick(cols, keys, exclude=()):
    for k in keys:
        for c in cols:
            if c not in exclude and k in c.lower():
                return c
    return None


def contacts_from_frame(df):
    df = df.fillna("").astype(str)
    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)
    num_col = _pick(cols, ["whatsapp", "mobile", "phone", "number", "contact no", "cell"])
    if num_col is None and len(cols) == 1:
        num_col = cols[0]
        if len(re.sub(r"\D", "", cols[0])) >= 10:     # header row was actually a number
            df = pd.concat([pd.DataFrame({cols[0]: [cols[0]]}), df], ignore_index=True)
    if num_col is None:
        raise ValueError("Could not find a phone-number column. Name it 'WhatsApp Number', 'Mobile' or 'Phone'. "
                         "Columns found: " + ", ".join(cols))
    name_col = _pick(cols, ["college name", "college", "institut", "organi", "company", "name"], exclude=(num_col,))
    contact_col = _pick(cols, ["tpo", "contact", "person", "owner"], exclude=(num_col, name_col))
    type_col = _pick(cols, ["number type"], exclude=(num_col,))
    out, seen = [], set()
    for _, r in df.iterrows():
        number, check = normalize_number(r[num_col], r[type_col] if type_col else "")
        if check == "ok" and number in seen:
            check = "duplicate"
        if check == "ok":
            seen.add(number)
        if not number and not any(str(v).strip() for v in r.values):
            continue                                       # fully blank row
        out.append({"number": number, "name": r[name_col].strip() if name_col else "",
                    "contact": r[contact_col].strip() if contact_col else "", "check": check})
    return out


def parse_contacts_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt", ".tsv"):
        df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding_errors="replace")
    else:
        df = pd.read_excel(path, dtype=str)
    return contacts_from_frame(df)


def parse_pasted(text):
    """One contact per line:  919876543210   or   919876543210, Name of college"""
    rows = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or not re.search(r"\d{8,}", re.sub(r"[\s\-+()]", "", line.split(",")[0])):
            continue
        parts = [p.strip() for p in re.split(r"[,;\t]", line, maxsplit=1)]
        rows.append({"WhatsApp Number": parts[0], "College Name": parts[1] if len(parts) > 1 else ""})
    if not rows:
        return []
    return contacts_from_frame(pd.DataFrame(rows))


def render_message(template, c):
    """Replace {name}/{college} (college or organisation) and {contact} (contact person). Nothing else is touched."""
    name = c.get("name") or "your college"
    contact = c.get("contact") or "Sir/Madam"
    out = template
    for key, val in (("{name}", name), ("{college}", name), ("{contact}", contact)):
        out = out.replace(key, val)
    return out


# ----------------------------------------------------------------------------
# WHATSAPP WEB CORE (the proven sending code)
# ----------------------------------------------------------------------------
INPUT_BOX_XPATH = ("//div[@id='main']//div[@contenteditable='true'][@role='textbox'] | "
                   "//footer//div[@contenteditable='true'] | "
                   "//div[@contenteditable='true'][@data-tab='10']")
INVALID_XPATH = ("//div[@role='dialog']//*[contains(text(),'invalid') or contains(text(),'not on WhatsApp') "
                 "or contains(text(),'not registered')] | "
                 "//div[@data-animate-modal-popup='true']//*[contains(text(),'invalid')]")
SEND_BTN_XPATH = ("//*[@aria-label='Send' or @aria-label='send' or @aria-label='SEND'] | "
                  "//span[contains(@data-icon,'send')] | //*[@data-testid='send']")

# Every outgoing chat bubble. The number of these going UP is our proof that something was really sent.
OUTGOING_XPATH = ("//div[@id='main']//div[contains(@class,'message-out')] | "
                  "//div[@id='main']//div[starts-with(@data-id,'true_')]")


def _pdf_name_xpath():
    """File name shown in the attachment preview header (NOT inside a chat bubble)."""
    name = os.path.basename(CURRENT_FILE)
    if "'" in name:
        return None
    return f"//*[contains(text(),'{name}')][not(ancestor::div[@data-id])]"


# Signals that the attachment-preview screen is open. Several are used on purpose, because WhatsApp changes
# this screen often ("No preview available" / "kB - PDF" in older versions, "1 page" + Adobe in the new one).
PREVIEW_TEXT_XPATH = ("//*[normalize-space(text())='No preview available'] | "
                      "//*[contains(text(),'kB - PDF') or contains(text(),'MB - PDF')] | "
                      "//*[contains(text(),'Powered by Adobe')]")

# JS: find the green Send button of the preview. 1) any visible element in the lower-right corner whose
# aria-label / data-icon / data-testid says "send";  2) whatever sits at the usual button position.
JS_SEND_BUTTON = """
const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
  return r.width > 8 && r.height > 8 && s.visibility !== 'hidden' && s.display !== 'none'; };
const attrs = e => (e.getAttribute('aria-label')||'') + ' ' + (e.getAttribute('data-icon')||'') + ' ' + (e.getAttribute('data-testid')||'');
const W = window.innerWidth, H = window.innerHeight;
const inCorner = e => { const r = e.getBoundingClientRect(); return r.left + r.width/2 > W*0.6 && r.top + r.height/2 > H*0.6; };
const found = [...document.querySelectorAll('[aria-label],[data-icon],[data-testid]')]
   .filter(e => /send/i.test(attrs(e)) && vis(e) && inCorner(e));
if (found.length) { const el = found[found.length-1]; return el.closest('button,[role=button]') || el; }
let el = document.elementFromPoint(W-58, H-62);
for (let i = 0; el && i < 6; i++, el = el.parentElement) {
  if (/send/i.test(attrs(el))) return el.closest('button,[role=button]') || el;
}
return null;"""

JS_SEND_BUTTON_EXISTS = "return !!(function(){" + JS_SEND_BUTTON + "})();"


def ensure_window(driver, min_w=1100, min_h=700):
    """WhatsApp Web changes its layout in small windows / high zoom -> force a big, normal-sized window."""
    try:
        driver.maximize_window()
    except Exception:
        pass
    time.sleep(1)
    w, h = driver.execute_script("return [window.innerWidth, window.innerHeight];")
    if w < min_w or h < min_h:
        try:
            driver.set_window_position(0, 0)
            driver.set_window_size(1400, 900)
            time.sleep(1)
            w, h = driver.execute_script("return [window.innerWidth, window.innerHeight];")
        except Exception:
            pass
    if w < min_w or h < min_h:
        say(f"WARNING: browser viewport is only {w}x{h}. Use a bigger screen or press Ctrl+0 in the "
              f"Chrome window to reset zoom to 100%, otherwise the Send button may not be found.")
    return w, h


def start_browser():
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={CHROME_PROFILE}")
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    driver = webdriver.Chrome(options=opts)
    ensure_window(driver)
    driver.get("https://web.whatsapp.com")
    say("If you see a QR code, scan it from your phone (WhatsApp > Linked devices). Waiting...")
    WebDriverWait(driver, 300).until(
        lambda d: d.find_elements(By.XPATH, "//div[@id='pane-side']") or
        d.find_elements(By.XPATH, "//div[@contenteditable='true']"))
    time.sleep(5)
    return driver


def dismiss_alert(driver):
    try:
        driver.switch_to.alert.accept()
    except NoAlertPresentException:
        pass


def open_chat(driver, number, text=""):
    """Returns 'ready' | 'not_on_whatsapp' | 'timeout'.  text='' opens the chat with an empty message box."""
    url = f"https://web.whatsapp.com/send?phone={number}"
    if text:
        url += f"&text={urllib.parse.quote(text)}"
    try:
        driver.get(url)
    except UnexpectedAlertPresentException:
        dismiss_alert(driver)
        driver.get(url)

    def state(d):
        if d.find_elements(By.XPATH, INPUT_BOX_XPATH):
            return "ready"
        if d.find_elements(By.XPATH, INVALID_XPATH):
            time.sleep(3)                               # confirm: popup still there AND no message box
            if d.find_elements(By.XPATH, INPUT_BOX_XPATH):
                return "ready"
            if d.find_elements(By.XPATH, INVALID_XPATH):
                return "not_on_whatsapp"
        return False

    try:
        result = WebDriverWait(driver, PAGE_TIMEOUT).until(state)
    except (TimeoutException, UnexpectedAlertPresentException):
        dismiss_alert(driver)
        result = "timeout"
    if result in ("not_on_whatsapp", "timeout"):
        try:
            driver.save_screenshot(os.path.join(BASE_DIR, f"debug_{number}_{result}.png"))
        except Exception:
            pass
    return result


# ----------------------------------------------------------------------------
# "Did it REALLY get sent?" helpers - every send is verified against the chat itself
# ----------------------------------------------------------------------------

JS_CHAT_COUNT = """
const sels = ['[data-id^="true_"]', '.message-out', '[data-icon^="msg-"]',
              '[data-testid="msg-container"]', '[role="row"]'];
const side = document.getElementById('pane-side');
const seen = new Set();
for (const s of sels) {
  for (const e of document.querySelectorAll(s)) { if (side && side.contains(e)) continue; seen.add(e); }
}
return seen.size;"""

# The chat list (left side) shows "<file name>" as the last message of a chat once the PDF is sent.
JS_SIDEBAR_FILE = """
const side = document.getElementById('pane-side'); if (!side) return 0;
const name = arguments[0];
return [...side.querySelectorAll('span,div')].filter(e => e.children.length === 0 && (e.textContent||'').includes(name)).length;"""


def snapshot(driver):
    """(messages visible in the open chat, chat-list rows showing the PDF name). Either one going UP = something was sent."""
    try:
        main_n = driver.execute_script(JS_CHAT_COUNT)
        side_n = driver.execute_script(JS_SIDEBAR_FILE, os.path.basename(CURRENT_FILE))
        return (int(main_n or 0), int(side_n or 0))
    except Exception:
        return (0, 0)


def grew(now_, base):
    return now_[0] > base[0] or now_[1] > base[1]


def composer_text(driver):
    """Text currently in the message box ('' = empty, None = box not found)."""
    try:
        els = [e for e in driver.find_elements(By.XPATH, INPUT_BOX_XPATH) if e.is_displayed()]
        return (els[-1].text or "").strip() if els else None
    except Exception:
        return None



def wait_sent(driver, base, secs):
    """True as soon as the chat (or the chat list) shows something new."""
    end = time.time() + secs
    while time.time() < end:
        if grew(snapshot(driver), base):
            return True
        time.sleep(0.5)
    return grew(snapshot(driver), base)


def preview_open(d):
    """True while the attachment-preview screen (with its Send button) is open."""
    try:
        if d.find_elements(By.XPATH, PREVIEW_TEXT_XPATH):
            return True
        xp = _pdf_name_xpath()
        if xp and any(e.is_displayed() for e in d.find_elements(By.XPATH, xp)):
            return True
        return bool(d.execute_script(JS_SEND_BUTTON_EXISTS))
    except Exception:
        return False


def _click_real(driver, el):
    ActionChains(driver).move_to_element(el).pause(0.3).click().perform()


def _cdp_click(driver, x, y):
    """A genuine mouse click at screen position (x, y), sent through Chrome DevTools. Cannot be 'intercepted'."""
    driver.execute_cdp_cmd("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y,
                                                        "button": "none", "clickCount": 0})
    time.sleep(0.15)
    for typ in ("mousePressed", "mouseReleased"):
        driver.execute_cdp_cmd("Input.dispatchMouseEvent", {"type": typ, "x": x, "y": y,
                                                            "button": "left", "clickCount": 1})
        time.sleep(0.1)


def _find_send(driver):
    el = driver.execute_script(JS_SEND_BUTTON)
    if el is None:
        raise RuntimeError("no send-like element found")
    return el


# ---- send strategies for the attachment preview (tried in this order until the chat shows the file) ----
def _s_enter(driver):                         # Enter: focus is normally in the caption box right after paste
    ActionChains(driver).send_keys(Keys.ENTER).perform()


def _s_real_click(driver):                    # real mouse click on the element that looks like Send
    _click_real(driver, _find_send(driver))


def _s_cdp_on_element(driver):                # DevTools click on that element's centre
    x, y = driver.execute_script(
        "const r=arguments[0].getBoundingClientRect(); return [r.left+r.width/2, r.top+r.height/2];",
        _find_send(driver))
    _cdp_click(driver, x, y)


def _s_cdp_corner(driver):                    # DevTools click where the green button normally is (bottom-right)
    w, h = driver.execute_script("return [window.innerWidth, window.innerHeight];")
    _cdp_click(driver, w - 58, h - 62)


def _s_caption_enter(driver):                 # click into the caption box, then Enter
    boxes = [e for e in driver.find_elements(By.XPATH, "//div[@contenteditable='true']") if e.is_displayed()]
    if not boxes:
        raise RuntimeError("no caption box")
    boxes[-1].click()
    boxes[-1].send_keys(Keys.ENTER)


def _s_selenium_click(driver):
    _find_send(driver).click()


def _s_js_click(driver):                      # last resort: JS click on the element and its parents
    driver.execute_script(
        "let el = arguments[0]; for (let i = 0; el && i < 4; i++, el = el.parentElement) { el.click(); }",
        _find_send(driver))


PREVIEW_STRATEGIES = (_s_enter, _s_real_click, _s_cdp_on_element, _s_cdp_corner,
                      _s_caption_enter, _s_selenium_click, _s_js_click)


def dump_debug(driver, number):
    """Save a screenshot + what WhatsApp's page looks like around the Send button, so it can be fixed precisely."""
    try:
        driver.save_screenshot(os.path.join(BASE_DIR, f"debug_{number}_pdf_failed.png"))
    except Exception:
        pass
    try:
        info = driver.execute_script("""
            const W = window.innerWidth, H = window.innerHeight, out = [];
            let el = document.elementFromPoint(W-58, H-62);
            for (let i = 0; el && i < 6; i++, el = el.parentElement) {
              out.push([el.tagName, el.getAttribute('role'), el.getAttribute('aria-label'),
                        el.getAttribute('data-icon'), el.getAttribute('data-testid'),
                        (el.className||'').toString().slice(0,80)].join(' | '));
            }
            const corner = [...document.querySelectorAll('*')].filter(e => {
                const r = e.getBoundingClientRect();
                return r.width > 20 && r.height > 20 && r.left > W*0.7 && r.top > H*0.7 &&
                       (e.getAttribute('aria-label') || e.getAttribute('data-icon') || e.getAttribute('role'));
            }).slice(0, 40).map(e => e.tagName + ' | ' + e.getAttribute('role') + ' | ' +
                       e.getAttribute('aria-label') + ' | ' + e.getAttribute('data-icon'));
            const lab = [...document.querySelectorAll('[aria-label]')].map(e => e.tagName+':'+e.getAttribute('aria-label')).slice(0,80);
            const ico = [...document.querySelectorAll('[data-icon]')].map(e => e.getAttribute('data-icon')).slice(-50);
            return 'AT SEND POSITION (child -> parents):\\n' + out.join('\\n') +
                   '\\n\\nBOTTOM-RIGHT ELEMENTS:\\n' + corner.join('\\n') +
                   '\\n\\naria-labels: ' + lab.join(', ') + '\\n\\ndata-icons: ' + ico.join(', ');""")
        with open(os.path.join(BASE_DIR, f"debug_{number}_send_button.txt"), "w", encoding="utf-8") as f:
            f.write(str(info))
    except Exception:
        pass


def close_preview(driver):
    """Close a stuck preview WITHOUT sending. Returns True when it is gone."""
    for _ in range(2):
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
        time.sleep(1.2)
        if not preview_open(driver):
            return True
    for el in driver.find_elements(By.XPATH, "//*[@aria-label='Close' or @data-icon='x' or "
                                              "@data-icon='x-viewer' or @data-icon='close']"):
        try:
            _click_real(driver, el)
            time.sleep(1.2)
            if not preview_open(driver):
                return True
        except Exception:
            continue
    return not preview_open(driver)



def _closed_outcome(driver, base):
    """Preview is gone. 'sent' = proof seen; 'assumed' = preview closed and stayed closed; None = it came back."""
    if wait_sent(driver, base, 15):
        return "sent"
    for _ in range(3):
        time.sleep(0.7)
        if preview_open(driver):
            return None
    return "assumed"


def finish_attachment(driver, base, number=""):
    """The file is staged in the preview screen: click Send. Returns 'sent' or 'assumed'."""
    WebDriverWait(driver, 25).until(lambda d: preview_open(d) or grew(snapshot(d), base))
    if grew(snapshot(driver), base):
        return "sent"
    time.sleep(random.uniform(2.5, 4))                    # let the preview finish rendering

    def done(result):
        time.sleep(random.uniform(6, 9) + _upload_extra())   # let the upload finish before leaving the chat
        return result

    for _round in range(2):
        for strategy in PREVIEW_STRATEGIES:
            if not preview_open(driver):
                out = _closed_outcome(driver, base)
                if out:
                    return done(out)
            try:
                strategy(driver)
            except Exception:
                continue
            if wait_sent(driver, base, 6):
                return done("sent")
        time.sleep(2)
    if not preview_open(driver):
        out = _closed_outcome(driver, base)
        if out:
            return done(out)
    if wait_sent(driver, base, 10):
        return done("sent")
    raise RuntimeError("PDF preview still open - Send could not be clicked")


def copy_file_to_clipboard(path):
    safe = path.replace("'", "''")
    subprocess.run(["powershell", "-NoProfile", "-Command", f"Set-Clipboard -LiteralPath '{safe}'"],
                   check=True, capture_output=True)


def attach_by_paste(driver):
    copy_file_to_clipboard(CURRENT_FILE)
    box = driver.find_element(By.XPATH, INPUT_BOX_XPATH)
    box.click()
    time.sleep(1)
    box.send_keys(Keys.CONTROL, "v")


def attach_by_menu(driver):
    attach = WebDriverWait(driver, 15).until(lambda d: d.find_elements(
        By.XPATH, "//span[@data-icon='plus-rounded' or @data-icon='plus' or @data-icon='clip' "
                  "or @data-icon='attach-menu-plus'] | //*[@aria-label='Attach']"))[0]
    attach.click()
    time.sleep(1.5)
    file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
    if not file_inputs:
        raise RuntimeError("no file input in attach menu")
    doc_input = next((i for i in file_inputs if (i.get_attribute("accept") or "") in ("*", "*/*")), file_inputs[0])
    doc_input.send_keys(CURRENT_FILE)



def send_text(driver):
    """Press Enter on the pre-filled message and confirm it left the message box (click Send if Enter did nothing)."""
    box = driver.find_element(By.XPATH, INPUT_BOX_XPATH)
    if not composer_text(driver):
        raise RuntimeError("message box is empty - the pre-filled text did not load")
    base = snapshot(driver)
    time.sleep(random.uniform(1.5, 3))
    box.send_keys(Keys.ENTER)

    def confirmed(secs):
        end = time.time() + secs
        while time.time() < end:
            if grew(snapshot(driver), base) or composer_text(driver) == "":
                return True
            time.sleep(0.5)
        return False

    if not confirmed(10):
        for el in driver.find_elements(By.XPATH, SEND_BTN_XPATH)[::-1]:
            try:
                _click_real(driver, el)
                break
            except Exception:
                continue
        if not confirmed(10):
            raise RuntimeError("text message was not sent (message box still full)")
    time.sleep(random.uniform(2, 4))



def send_pdf(driver, number=""):
    """Attach the PDF and send it. Returns 'sent' (proof seen) or 'assumed' (preview closed). Raises if not sent."""
    base = snapshot(driver)
    errors = []
    for attach in (attach_by_paste, attach_by_menu):
        try:
            attach(driver)
            return finish_attachment(driver, base, number)
        except Exception as e:
            errors.append(f"{attach.__name__}: {type(e).__name__} {e}".strip())
            if grew(snapshot(driver), base):              # it actually went out
                return "sent"
            dump_debug(driver, number)
            if preview_open(driver) and not close_preview(driver):
                break                                     # cannot clear the screen -> never stack a 2nd copy
            if grew(snapshot(driver), base):
                return "sent"
    raise RuntimeError("PDF not sent (" + "; ".join(errors) + ")")


# ----------------------------------------------------------------------------
# WhatsApp run loop
# ----------------------------------------------------------------------------


# ----------------------------------------------------------------------------
# CAMPAIGN: log, attachments, main loop
# ----------------------------------------------------------------------------
LOG_COLS = ["number", "name", "contact", "status", "time", "detail", "campaign"]
DONE_STATUSES = ("sent", "sent_unverified", "text_sent_attachment_failed", "not_on_whatsapp")


def log_row(number, name, contact, status, detail="", campaign=""):
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
    row = {"number": number, "name": name, "contact": contact, "status": status,
           "time": now(), "detail": detail, "campaign": campaign}
    pd.DataFrame([row], columns=LOG_COLS).to_csv(LOG_PATH, mode="a", header=not os.path.exists(LOG_PATH),
                                                 index=False, encoding="utf-8-sig")


def previously_done():
    """Numbers that already got a message (or are not on WhatsApp) in an earlier campaign."""
    if not os.path.exists(LOG_PATH):
        return set()
    log = pd.read_csv(LOG_PATH, dtype=str).fillna("")
    if log.empty:
        return set()
    last = log.drop_duplicates("number", keep="last")
    return set(last[last["status"].isin(DONE_STATUSES)]["number"])


def send_attachment(driver, number, path):
    """Send one file. Returns 'sent' (proof seen) or 'assumed' (preview closed). Raises if it could not be sent."""
    global CURRENT_FILE
    CURRENT_FILE = path
    return send_pdf(driver, number)


def _sleep(seconds, should_stop):
    end = time.time() + seconds
    while time.time() < end:
        if should_stop():
            return
        time.sleep(min(1, max(0, end - time.time())))


def run_campaign(contacts, message, attachments, *, limit=30, min_delay=60, max_delay=150,
                 long_break_every=10, long_break_range=(300, 600), skip_previous=True,
                 check_only=False, emit=None, should_stop=None, campaign=""):
    """
    contacts     list of dicts from parse_contacts_file / parse_pasted
    message      text with optional {name} {college} {contact}; '' = send attachments only
    attachments  list of file paths, sent one after another after the text
    emit(event)  called with {'event': 'row', 'index': i, 'status': ..., ...} for every contact
    Returns a dict of counts.
    """
    emit = emit or (lambda ev: None)
    should_stop = should_stop or (lambda: False)
    message = (message or "").strip()
    attachments = [a for a in (attachments or []) if a]
    if not message and not attachments:
        raise ValueError("Nothing to send: write a message or add an attachment.")
    for a in attachments:
        if not os.path.exists(a):
            raise FileNotFoundError(f"Attachment not found: {a}")

    counts = {}

    def record(i, c, status, detail="", save=True):
        counts[status] = counts.get(status, 0) + 1
        if save:
            log_row(c["number"], c["name"], c["contact"], status, detail, campaign)
        emit({"event": "row", "index": i, "number": c["number"], "name": c["name"], "status": status, "detail": detail})

    done = previously_done() if skip_previous else set()
    todo = []
    for i, c in enumerate(contacts):
        if c["check"] != "ok":
            record(i, c, c["check"], save=False)             # invalid / landline / duplicate: reported, not logged
        elif c["number"] in done:
            record(i, c, "skipped_already_messaged", save=False)
        else:
            todo.append((i, c))
    say(f"{len(todo)} contact(s) ready to message. This run will send to at most {limit}.")
    emit({"event": "plan", "ready": len(todo), "will_send": min(len(todo), limit)})
    if check_only or not todo:
        for i, c in todo:
            record(i, c, "ready", save=False)
        return counts

    try:
        driver = start_browser()
    except Exception as e:
        say("Could not open Chrome: " + str(e)[:200])
        say("Close any Chrome window opened by this tool earlier, then try again.")
        raise
    processed = sent = 0
    try:
        for i, c in todo:
            if processed >= limit or should_stop():
                break
            number = c["number"]
            text = render_message(message, c).strip() if message else ""
            status, detail, pause = "failed", "", None
            say(f"[{processed + 1}] {number} ({c['name'] or 'no name'}) ...")
            try:
                state = open_chat(driver, number, text)
                if state == "not_on_whatsapp":
                    status, pause = "not_on_whatsapp", random.uniform(8, 15)
                elif state == "timeout":
                    status, detail = "failed", "chat did not load (timeout)"
                else:
                    if text:
                        send_text(driver)                  # raises if the text did not really go out
                    unverified, failed = 0, []
                    for k, path in enumerate(attachments):
                        try:
                            if send_attachment(driver, number, path) != "sent":
                                unverified += 1
                        except Exception as e:
                            failed.append(f"{os.path.basename(path)}: {str(e)[:90]}")
                        if k < len(attachments) - 1:
                            time.sleep(random.uniform(2, 4))
                    if failed:
                        status = "text_sent_attachment_failed" if text else "failed"
                        detail = "; ".join(failed)
                    else:
                        status = "sent_unverified" if unverified else "sent"
                        if unverified:
                            detail = "attachment preview closed but delivery not auto-confirmed - check your phone"
                    if status in ("sent", "sent_unverified", "text_sent_attachment_failed"):
                        sent += 1
                    pause = (random.randint(*long_break_range) if sent and sent % long_break_every == 0
                             and status != "failed" else random.uniform(min_delay, max_delay))
            except Exception as e:
                status, detail = "failed", f"{type(e).__name__}: {str(e)[:140]}"
                dismiss_alert(driver)
            record(i, c, status, detail)
            say(f"    -> {status}" + (f" ({detail})" if detail else ""))
            processed += 1
            if processed < limit and not should_stop():
                pause = pause if pause is not None else random.uniform(min_delay, max_delay)
                say(f"    waiting {int(pause)}s before the next contact (keeps your number safe)")
                _sleep(pause, should_stop)
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        say(f"Finished. Processed {processed} contact(s), sent to {sent}.")
    return counts
