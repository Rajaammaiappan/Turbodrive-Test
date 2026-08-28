"""

flask_app.py

=========================================================================

JETPULL - Flask web UI  (RR / ALTEN internal tool)



Universal document downloader with portal presets:

  [x] BDC (BlueData Connect)  - URL logic HARDCODED below (live)

  [ ] Maximo / B2B / One Stop Portal - placeholders, logic to follow

  [ ] nothing ticked          - universal mode: user types any URL



BDC routing (hardcoded, per TV number):

  5xxxxx -> .../technicalVariances/Raising/LegacyTV        (always)

  6xxxxx -> .../technicalVariances/Raising/TV, and if the number isEvent

            not found there -> .../technicalVariances/Raising/RepeaterTV

  other  -> tries LegacyTV, TV, RepeaterTV in that order



Output layout: <download folder>/<TV number>/<files>

  Zip downloads are unpacked into the TV folder and the .zip removed.



Usage logging (NOT shown in the tool): every processed number is

appended to a SQLite DB (schema mirrors processed_emails.db) so yearly

savings can be calculated straight from the file:

    columns: portal, tv_number, url_used, downloaded_by, machine,

             target_folder, file_count, file_names, status, message,

             duration_seconds, processed_at



Runs on Flask 1.1.2 / Python 3.9 (corporate Anaconda) - port 5001.

=========================================================================

"""

import getpass

import re

import socket

import sqlite3

import tempfile

import threading

import time

import webbrowser

import pandas as pd

from datetime import datetime

from pathlib import Path


from flask import Flask, jsonify, render_template_string, request


from cdp_automator import (
    AutomationEngine,
    extract_zips,
    load_config,
    read_search_numbers_from_excel,
    safe_resolve,
    write_report_to_excel,
)

app = Flask(__name__)

PORT = 5001


# ------------------------------------------------------------------ #

#  HARDCODED SETTINGS

# ------------------------------------------------------------------ #

# BlueData Connect base - the section (LegacyTV / TV / RepeaterTV) is

# chosen per TV number by bdc_paths_for() below.

BDC_BASE = "https://bluedata.connect.rolls-royce.com/technicalVariances/Raising"


# Usage database. Lives next to the tool by default, so if you deploy

# this folder on a shared drive, EVERY user's runs land in the same DB.

# Change this constant to a fixed network path if you prefer, e.g.

# r"\\portfolioeng_nlr\EFS\BDC_Downloader\bdc_usage_log.db"

USAGE_DB_PATH = str(safe_resolve(Path(__file__).parent) / "bdc_usage_log.db")


# Logos - two ways (local file wins if both are set):

#   a) Save images into the ./static folder as alten_logo.png / rr_logo.png

#   b) Or paste DIRECT image URLs below (must point at the image itself,

#      e.g. https://.../alten.png - NOT a Bing/Google image-search page).

# Email typed automatically into the SSO login page when the form's

# Username field is left empty. Change it here, not in the UI.

# Login is manual - the tool never stores or types passwords. This

# constant is retained (empty) only for backward compatibility.

DEFAULT_LOGIN_EMAIL = ""


ALTEN_LOGO_URL = (
    "https://companieslogo.com/img/orig/ATE.PA_BIG-569b3a98.png?t=1750661371"
)

RR_LOGO_URL = "https://www.rolls-royce.com/~/media/Images/R/Rolls-Royce/logo/rr-logo-svg.svg?h=96&iar=0&w=59"


def bdc_paths_for(number: str):
    """Which BlueData sections to try, in order, for this TV number."""

    n = number.strip()

    if n.startswith("5"):

        return ["LegacyTV"]  # 5xxxxx: always LegacyTV

    if n.startswith("6"):

        return ["TV", "RepeaterTV"]  # 6xxxxx: TV, else RepeaterTV

    return ["LegacyTV", "TV", "RepeaterTV"]  # anything else: try all


# ------------------------------------------------------------------ #

#  Usage DB (silent - never displayed in the tool)

# ------------------------------------------------------------------ #


def _db() -> sqlite3.Connection:

    con = sqlite3.connect(USAGE_DB_PATH, timeout=15)

    con.execute("""

        CREATE TABLE IF NOT EXISTS usage_log (

            id               INTEGER PRIMARY KEY AUTOINCREMENT,

            portal           TEXT,

            tv_number        TEXT,

            url_used         TEXT,

            downloaded_by    TEXT,

            machine          TEXT,

            target_folder    TEXT,

            file_count       INTEGER DEFAULT 0,

            file_names       TEXT,

            status           TEXT,

            message          TEXT,

            duration_seconds REAL,

            processed_at     TEXT

        )""")

    return con


def log_usage(
    portal,
    tv_number,
    url_used,
    target_folder,
    file_count,
    file_names,
    status,
    message,
    duration,
):
    """Append one row per processed number. Failures never stop the run."""

    try:

        con = _db()

        con.execute(
            "INSERT INTO usage_log (portal, tv_number, url_used, "
            "downloaded_by, machine, target_folder, file_count, file_names, "
            "status, message, duration_seconds, processed_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                portal,
                tv_number,
                url_used,
                getpass.getuser(),
                socket.gethostname(),
                target_folder,
                file_count,
                "; ".join(file_names),
                status,
                message,
                round(duration, 1),
                f"{datetime.now():%Y-%m-%d %H:%M:%S}",
            ),
        )

        con.commit()

        con.close()

    except Exception as exc:  # noqa: BLE001

        ui_log(f"(usage log skipped: {exc})")


# ------------------------------------------------------------------ #

#  Single-job state

# ------------------------------------------------------------------ #

JOB = {
    "running": False,
    "logs": [],
    "stop": threading.Event(),
    "report": None,
    "summary": "",
    "results": [],
    "done": None,
    "progress": None,
}

LOCK = threading.Lock()


def ui_log(msg: str) -> None:

    with LOCK:

        JOB["logs"].append(f"{datetime.now():%H:%M:%S}  {msg}")


# ------------------------------------------------------------------ #

#  Sign-in: manual gate, or the old automatic SSO attempt

# ------------------------------------------------------------------ #


def wait_for_manual_signin(engine, cfg) -> None:
    """
    Open a warm-up tab, pause while the user signs in, then reload the portal.

    Two problems this solves:

    * Going straight to the BDC link fails Microsoft's MFA with "Sorry, we're
      having trouble verifying your account". Signing in to an ordinary
      corporate site first (Engine Room) completes the SSO cleanly, and BDC
      then inherits that session. So a second tab is opened on
      `sso_warmup_url` and left in front for the user to sign in on.

    * Detecting when sign-in has finished is unreliable, because Microsoft
      varies the wording of its MFA pages. So the tool simply waits
      `manual_login_wait_seconds` and then carries on.

    Afterwards the tool's own tab is brought back to the front and the portal
    URL is loaded again - the page fetched before sign-in is normally the
    Microsoft error page, so it has to be re-fetched on the valid session.
    """
    wait_s = float(cfg.get("manual_login_wait_seconds", 60))
    settle = float(cfg.get("post_login_settle_seconds", 5))
    warmup_url = cfg.get("sso_warmup_url", "")

    # ---- 1. Warm-up tab -------------------------------------------------
    warm_tid = ""
    if warmup_url:
        warm_tid = engine.open_warmup_tab(warmup_url)
        if warm_tid:
            ui_log("A second tab has opened for sign-in. Use THAT tab - "
                   "signing in there also authorises the portal.")

    ui_log(f"Sign in now - you have {int(wait_s)} seconds. Complete the "
           "Authenticator approval; the download starts automatically when "
           "the time is up.")

    # ---- 2. Countdown ---------------------------------------------------
    deadline = time.time() + wait_s
    next_notice = wait_s - 10
    while True:
        remaining = deadline - time.time()
        if remaining <= 0:
            break
        if JOB["stop"].is_set():
            engine.close_target(warm_tid)
            raise RuntimeError("Stopped while waiting for sign-in.")
        if remaining <= next_notice:
            ui_log(f"Starting in {int(remaining)} seconds...")
            next_notice = (remaining - 10) if remaining > 15 else (remaining - 5)
        time.sleep(0.5)

    # ---- 3. Back to our own tab, on a valid session ---------------------
    if warm_tid and cfg.get("close_warmup_tab", True):
        engine.close_target(warm_tid)
        ui_log("Sign-in tab closed.")
    engine.focus_main_tab()
    time.sleep(settle)                       # let any last redirect finish

    if cfg.get("reload_after_signin", True):
        engine.reload_last_site()

    ui_log("Continuing with the download...")
    try:
        engine.wait_for_page_ready()
    except Exception:                        # noqa: BLE001
        pass



def do_signin(engine, credentials, cfg) -> None:
    """

    Sign in, either manually (default) or by the older automatic SSO flow.



    Set "manual_login": false in config.json to go back to the automatic

    attempt - the code for it is untouched in cdp_automator.login().

    """

    if cfg.get("manual_login", True):

        wait_for_manual_signin(engine, cfg)

    else:

        engine.login(*credentials)


# ------------------------------------------------------------------ #

#  One item = one TV number

# ------------------------------------------------------------------ #


def process_number(
    engine,
    portal,
    base_url,
    number,
    base_folder,
    parent_history_df,
    logged_in_flag,
    credentials,
):
    """

    Process one number: navigate (BDC: try candidate sections), search,

    open, download into <base_folder>/<number>, unzip archives.

    Returns the report row (also written to the usage DB).

    """

    started = time.time()

    tv_folder = str(Path(base_folder) / re.sub(r'[\\/:*?"<>|]', "_", number))

    row = {
        "search_number": number,
        "files": [],
        "timestamp": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
    }

    url_used = ""

    # ---- B2B / Maximo branch -------------------------------------

    if portal == "B2B":

        from maximo_b2b import MaximoWorkflow, MAXIMO_URL

        url_used = MAXIMO_URL

        wf = MaximoWorkflow(engine, load_config(), base_folder, parent_history_df)

        # login once (SSO handled by engine.login when the page appears)

        if not logged_in_flag["done"]:

            try:

                engine.open_site(MAXIMO_URL)

                do_signin(engine, credentials, load_config())

                logged_in_flag["done"] = True

            except Exception as exc:  # noqa: BLE001

                row.update(status="FAILED", reason="ERROR", message=str(exc))

                log_usage(
                    "B2B",
                    number,
                    url_used,
                    base_folder,
                    0,
                    [],
                    "FAILED",
                    str(exc),
                    time.time() - started,
                )

                return row

        res = wf.run(number, base_folder)

        status = "SUCCESS" if res.get("status") == "SUCCESS" else "FAILED"

        reason = (
            ""
            if status == "SUCCESS"
            else ("FILE_NOT_AVAILABLE" if res.get("status") == "NO_FILES" else "ERROR")
        )

        row.update(
            status=status,
            reason=reason,
            files=res.get("files", []),
            message=res.get("message", ""),
        )

        ui_log(f"{number}: {row['message']}")

        log_usage(
            "B2B",
            number,
            url_used,
            "; ".join(res.get("folders", [])) or base_folder,
            len(row["files"]),
            row["files"],
            status,
            row["message"],
            time.time() - started,
        )

        return row

    candidates = (
        [f"{BDC_BASE}/{p}" for p in bdc_paths_for(number)]
        if portal == "BDC"
        else [base_url]
    )

    engine.set_download_dir(tv_folder)

    last_err = None

    for cand in candidates:

        try:

            url_used = cand

            engine.open_site(cand)

            if not logged_in_flag["done"]:

                do_signin(engine, credentials, load_config())

                logged_in_flag["done"] = True

            engine.search_record(number)

            engine.open_result(number)

            files = engine.download_files()

            files = extract_zips(files, tv_folder, engine.log)

            names = [Path(f).name for f in files]

            row.update(
                status="SUCCESS",
                files=names,
                message=f"{len(names)} file(s) in {tv_folder}",
            )

            ui_log(f"{number}: {len(names)} file(s) saved to {tv_folder}")

            break

        except Exception as exc:  # noqa: BLE001

            last_err = exc

            if cand != candidates[-1]:

                ui_log(
                    f"{number}: not found via {cand.rsplit('/',1)[-1]} - "
                    f"trying next section..."
                )

            continue

    else:

        msg = str(last_err)

        low = msg.lower()

        if "no result found" in low or "no search box" in low:

            reason = "NUMBER_NOT_FOUND"

            human = "Number not found in portal"

        elif "no download" in low or "no files" in low or "no file links" in low:

            reason = "FILE_NOT_AVAILABLE"

            human = "No downloadable file available for this number"

        else:

            reason = "ERROR"

            human = msg

        row.update(status="FAILED", reason=reason, message=human)

        ui_log(f"FAILED {number}: {human}")

    log_usage(
        portal,
        number,
        url_used,
        tv_folder,
        len(row["files"]),
        row["files"],
        row.get("status", "FAILED"),
        row.get("message", ""),
        time.time() - started,
    )

    return row


def _already_downloaded(base_folder, portal, number):
    """

    Resume helper: return existing files for this number, or [].

    BDC/Universal put files in <base>/<number>/ ; B2B nests under

    <base>/<number>/... so we look recursively there.

    """

    folder = Path(base_folder) / re.sub(r'[\\/:*?"<>|]', "_", number)

    if not folder.exists():

        return []

    files = [str(p) for p in folder.rglob("*") if p.is_file()]

    return files


def _save_report(report_rows, folder, partial=False):
    """Write the Excel report. Called periodically (partial) and at the end."""

    if not report_rows:

        return None

    try:

        name = (
            "download_report_PARTIAL.xlsx"
            if partial
            else f"download_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )

        path = str(Path(folder) / name)

        write_report_to_excel(report_rows, path)

        if partial:

            ui_log(f"(checkpoint saved: {len(report_rows)} rows -> {name})")

        with LOCK:

            JOB["report"] = path

        return path

    except Exception as exc:  # noqa: BLE001

        ui_log(f"Could not write report: {exc}")

        return None


def read_parent_history_excel(path):

    raw = pd.read_excel(path, header=None)

    header_row = None

    for idx, row in raw.iterrows():

        values = {str(v).strip().lower() for v in row.tolist()}

        if "event item serial number" in values and "event date time" in values:

            header_row = idx

            break

    if header_row is None:

        raise ValueError(
            "Could not find required columns "
            "'Event Item Serial Number' and "
            "'Event Date Time'."
        )

    df = pd.read_excel(
        path,
        header=header_row,
    )

    cols = {str(col).strip().lower(): col for col in df.columns}

    serial_col = cols.get("event item serial number")

    date_col = cols.get("event date time")

    action_col = cols.get("action code")

    if not serial_col or not date_col or not action_col:

        raise ValueError(
            "Required columns not found: "
            "'Event Item Serial Number', "
            "'Event Date Time' and/or "
            "'Action Code'"
        )

    df = df.rename(
        columns={
            serial_col: "event_serial",
            date_col: "event_date",
            action_col: "action_code",
        }
    )

    df = df[
        df["action_code"].fillna("").astype(str).str.strip().str.upper().eq("REMOVAL")
    ]

    df["event_serial"] = df["event_serial"].fillna("").astype(str).str.strip()

    df = df[df["event_serial"] != ""]

    return df


def run_job(
    portal,
    url,
    username,
    password,
    numbers,
    folder,
    headless,
    browser,
    skip_existing=False,
    parent_history_df=None,
    checkpoint_every=25,
):

    report_rows = []

    engine = None

    logged_in = {"done": False}

    original_order = list(numbers)  # keep the user's typed order for output

    try:

        engine = AutomationEngine(
            config=load_config(),
            download_dir=folder,
            status_callback=ui_log,
            headless=headless,
            browser=browser,
        )

        engine.start_browser()

        # Group by PRIMARY link so all numbers sharing a link run back-to-back

        # instead of ping-ponging (e.g. 5,5,5,6,5,6 -> all 5s, then all 6s).

        # Order within each group is preserved. The per-number fallback

        # (6-series /TV -> /RepeaterTV) still happens inside process_number.

        if portal == "BDC":

            def primary(n):

                secs = bdc_paths_for(n)

                return secs[0] if secs else "LegacyTV"

            # stable sort keeps typed order inside each link group

            group_order, seen = [], set()

            for n in numbers:

                key = primary(n)

                if key not in seen:

                    seen.add(key)
                    group_order.append(key)

            ordered = sorted(numbers, key=lambda n: group_order.index(primary(n)))

            if ordered != numbers:

                ui_log(
                    "Grouped by portal link to avoid switching back and "
                    "forth: "
                    + ", ".join(
                        f"{k}({sum(1 for n in numbers if primary(n)==k)})"
                        for k in group_order
                    )
                )

            numbers = ordered

        for idx, number in enumerate(numbers, start=1):

            if JOB["stop"].is_set():

                ui_log(f"Stopped by user before item {idx}/{len(numbers)}.")

                break

            ui_log(f"===== Item {idx}/{len(numbers)}: {number} " f"[{portal}] =====")

            # Live progress (drives the progress bar for large batches).

            with LOCK:

                JOB["progress"] = {
                    "done": idx - 1,
                    "total": len(numbers),
                    "current": number,
                }

            # RESUME support: skip numbers already downloaded in a previous

            # run (their folder exists and contains at least one file).

            if skip_existing:

                existing = _already_downloaded(folder, portal, number)

                if existing:

                    ui_log(
                        f"{number}: already downloaded "
                        f"({len(existing)} file(s)) - skipping."
                    )

                    report_rows.append(
                        {
                            "search_number": number,
                            "status": "SKIPPED",
                            "reason": "ALREADY_DOWNLOADED",
                            "files": [Path(f).name for f in existing],
                            "message": "Skipped - files already present",
                            "timestamp": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
                        }
                    )

                    continue

            # Universal mode keeps the login-once + go-back flow;

            # BDC navigates between sections inside the same session.

            if portal != "BDC" and idx > 1:

                try:

                    engine.return_to_search()

                except Exception:

                    engine.open_site(url)

            report_rows.append(
                process_number(
                    engine,
                    portal,
                    url,
                    number,
                    folder,
                    parent_history_df,
                    logged_in,
                    (username, password),
                )
            )

            # CHECKPOINT: on long batches, save the report periodically so a

            # crash or closed window never loses hours of work.

            if checkpoint_every and idx % checkpoint_every == 0:

                _save_report(report_rows, folder, partial=True)

    except Exception as exc:  # noqa: BLE001

        ui_log(f"FATAL: {exc}")

    finally:

        if engine:

            engine.quit()

    # Restore the user's typed order for the report and on-screen table

    # (execution was grouped by link for speed, but output should match

    # what they entered). Stable for duplicates.

    pos = {n: i for i, n in enumerate(original_order)}

    report_rows.sort(key=lambda r: pos.get(r.get("search_number", ""), 1e9))

    if report_rows:

        path = _save_report(report_rows, folder, partial=False)

        if path:

            ui_log(f"Report written: {path}")

    # Safety net: the same file arriving for several different numbers means

    # the tool opened the wrong record or grabbed a page-template link.

    # Flag it loudly rather than reporting a clean-looking but wrong run.

    seen_files = {}

    for r in report_rows:

        for fn in r.get("files", []):

            seen_files.setdefault(fn, []).append(r.get("search_number", ""))

    repeated = {f: nums for f, nums in seen_files.items() if len(set(nums)) > 1}

    if repeated:

        ui_log(
            "*** WARNING: the same file was downloaded for several "
            "different numbers - this usually means a wrong record was "
            "opened or a page-template link was picked up: "
            + "; ".join(
                f"{f} -> {', '.join(sorted(set(n)))[:60]}"
                for f, n in list(repeated.items())[:5]
            )
        )

        for r in report_rows:

            if any(fn in repeated for fn in r.get("files", [])):

                r["message"] = "CHECK: possible duplicate/wrong file. " + r.get(
                    "message", ""
                )

    ok = sum(1 for r in report_rows if r.get("status") == "SUCCESS")

    skipped = sum(1 for r in report_rows if r.get("status") == "SKIPPED")

    total = len(report_rows)

    results = [
        {
            "number": r.get("search_number", ""),
            "status": r.get("status", "FAILED"),
            "reason": r.get("reason", ""),
            "files": r.get("files", []),
            "message": r.get("message", ""),
        }
        for r in report_rows
    ]

    with LOCK:

        JOB["summary"] = f"DONE - {ok}/{total} item(s) succeeded" + (
            f", {skipped} skipped." if skipped else "."
        )

        JOB["results"] = results

        JOB["done"] = {"ok": ok, "total": total, "skipped": skipped}

        JOB["running"] = False

    ui_log(JOB["summary"])


# ------------------------------------------------------------------ #

#  Routes

# ------------------------------------------------------------------ #


@app.route("/")
def index():

    static_dir = Path(app.root_path) / "static"

    return render_template_string(
        PAGE,
        alten_logo=(static_dir / "alten_logo.png").exists(),
        rr_logo=(static_dir / "rr_logo.png").exists(),
        alten_logo_url=ALTEN_LOGO_URL.strip(),
        rr_logo_url=RR_LOGO_URL.strip(),
    )


@app.route("/start", methods=["POST"])
def start():

    with LOCK:

        if JOB["running"]:

            return jsonify(ok=False, error="A job is already running.")

    if request.form.get("portal_bdc") == "on":

        portal = "BDC"

    elif request.form.get("portal_b2b") == "on":

        portal = "B2B"

    else:

        portal = "Universal"

    url = (request.form.get("url") or "").strip()

    username = ""  # login is manual; tool never types credentials

    password = ""

    typed = (request.form.get("numbers") or "").strip()

    folder = (request.form.get("folder") or "").strip()

    headless = request.form.get("headless") == "on"

    browser = request.form.get("browser") or "edge"

    skip_existing = request.form.get("skip_existing") == "on"

    excel = request.files.get("excel")

    parent_history_excel = request.files.get("parent_history_excel")

    if portal not in ("BDC", "B2B") and not url:

        return jsonify(
            ok=False,
            error=(
                "Enter a website URL, or tick BDC/B2B " "to use a built-in portal link."
            ),
        )

    if not folder:

        return jsonify(
            ok=False,
            error="Please enter a download folder path.",
        )

    if portal == "B2B" and (
        not parent_history_excel or not parent_history_excel.filename
    ):

        return jsonify(
            ok=False,
            error="Please upload a Parent History Excel file.",
        )

    try:

        Path(folder).mkdir(parents=True, exist_ok=True)

    except OSError as exc:

        return jsonify(
            ok=False,
            error=f"Cannot use that folder: {exc}",
        )

    numbers = []

    if excel and excel.filename:

        try:

            tmp = Path(tempfile.gettempdir()) / f"rr_batch_{datetime.now():%H%M%S}.xlsx"

            excel.save(str(tmp))

            numbers = read_search_numbers_from_excel(str(tmp))

        except Exception as exc:  # noqa: BLE001

            return jsonify(
                ok=False,
                error=f"Could not read the Excel file: {exc}",
            )

    else:

        numbers = [n.strip() for n in re.split(r"[,;\n]", typed) if n.strip()]

    if not numbers:

        return jsonify(
            ok=False,
            error=("Enter at least one TV/search number " "or upload an Excel file."),
        )

    parent_history_df = None

    if portal == "B2B" and parent_history_excel and parent_history_excel.filename:

        try:

            parent_tmp = Path(tempfile.gettempdir()) / (
                f"parent_history_" f"{datetime.now():%H%M%S}.xlsx"
            )

            parent_history_excel.save(str(parent_tmp))

            parent_history_df = read_parent_history_excel(str(parent_tmp))

        except Exception as exc:

            return jsonify(
                ok=False,
                error=("Could not read the Parent History " f"Excel file: {exc}"),
            )

    with LOCK:

        JOB.update(
            running=True,
            logs=[],
            report=None,
            summary="",
            results=[],
            done=None,
            progress=None,
            stop=threading.Event(),
        )

    ui_log(f"Starting [{portal}]: " f"{len(numbers)} number(s).")

    threading.Thread(
        target=run_job,
        args=(
            portal,
            url,
            username,
            password,
            numbers,
            folder,
            headless,
            browser,
            skip_existing,
            parent_history_df,
        ),
        daemon=True,
    ).start()

    return jsonify(ok=True)


@app.route("/pick_folder", methods=["POST"])
def pick_folder():
    """

    Open the NATIVE Windows 'Select Folder' dialog on the server (this PC)

    and return the chosen absolute path. This works because the Flask server

    runs locally and has real disk access - unlike the browser, which is

    sandboxed and can never see a true folder path.

    """

    result = {"path": ""}

    def ask():

        try:

            import tkinter as tk

            from tkinter import filedialog

            root = tk.Tk()

            root.withdraw()

            root.attributes("-topmost", True)

            chosen = filedialog.askdirectory(title="Select the output folder")

            root.destroy()

            result["path"] = chosen or ""

        except Exception as exc:  # noqa: BLE001

            result["error"] = str(exc)

    # Tkinter must run on the main thread of its own; run and join briefly.

    t = threading.Thread(target=ask)

    t.start()

    t.join(timeout=120)

    if result.get("error"):

        return jsonify(ok=False, error=result["error"])

    return jsonify(ok=True, path=result["path"])


@app.route("/stop", methods=["POST"])
def stop():

    JOB["stop"].set()

    ui_log("Stop requested - finishing the current item...")

    return jsonify(ok=True)




@app.route("/status")
def status():

    after = request.args.get("after", type=int) or 0

    with LOCK:

        return jsonify(
            running=JOB["running"],
            logs=JOB["logs"][after:],
            next=len(JOB["logs"]),
            report=JOB["report"],
            summary=JOB["summary"],
            results=JOB["results"],
            done=JOB["done"],
            progress=JOB["progress"],
        )


# ------------------------------------------------------------------ #

#  Page - ALTEN mark on the LEFT, Rolls-Royce badge on the RIGHT

# ------------------------------------------------------------------ #

PAGE = r"""<!DOCTYPE html>

<html lang="en">

<head>

<meta charset="utf-8">

<title>JETPULL — RR / ALTEN</title>

<style>

  :root { --bg:#eef2f9; --surface:#ffffff; --border:#c8d4e8;

          --accent:#1a4fad; --accent2:#1d6fdb; --text:#1e293b;

          --muted:#475569; --green:#059669; --red:#dc2626; }

  * { box-sizing:border-box; }

  body { margin:0; background:var(--bg); color:var(--text);

         font:14px/1.5 "Segoe UI", Arial, sans-serif; }

  header { background:#fff; border-bottom:3px solid var(--accent);

           padding:12px 22px; display:flex; align-items:center; gap:16px; }

  .alten-mark { width:34px; height:46px; border:2px solid #111;

                background:linear-gradient(135deg,#e42313 0 48%,#ffd500 48% 100%);

                display:flex; align-items:flex-end; justify-content:center;

                font:700 10px/1.6 "Segoe UI"; color:#111; }

  .alten-word { font:700 17px "Segoe UI"; color:#111; letter-spacing:1px; }

  .titles { flex:1; margin-left:6px; }

  .titles h1 { font-size:20px; margin:0; color:#111; }

  .titles p  { margin:2px 0 0; font-size:12px; color:var(--muted);

               font-family:Consolas, monospace; }

  .logo-left  { height:46px; }

  .logo-right { height:54px; }

  .rr-badge { width:44px; height:56px; background:#10069f; color:#fff;

              border-radius:4px; display:flex; flex-direction:column;

              align-items:center; justify-content:center;

              font:700 8px "Segoe UI"; }

  .rr-badge span { font:700 20px Georgia, serif; letter-spacing:-2px; }

  main { max-width:900px; margin:18px auto; padding:0 16px; }

  fieldset { background:var(--surface); border:1px solid var(--border);

             border-radius:6px; margin:0 0 14px; padding:14px 16px; }

  legend { color:var(--accent); font-weight:600; padding:0 6px; }

  label { display:block; margin:8px 0 3px; color:var(--muted); font-size:13px; }

  input[type=text], input[type=password], select, textarea {

    width:100%; padding:7px 9px; border:1px solid var(--border);

    border-radius:4px; font:13px Consolas, monospace; }

  input:disabled { background:#e9eef7; color:#8aa0c4; }

  .row { display:flex; gap:14px; flex-wrap:wrap; }

  .row > div { flex:1; min-width:240px; }

  .portals { display:flex; gap:22px; flex-wrap:wrap; }

  .portals label { display:flex; align-items:center; gap:6px; margin:0;

                   font-size:14px; color:var(--text); }

  .portals .soon { color:#9db4d8; }

  .portals .soon small { font-size:10px; }

  .opts { display:flex; align-items:center; gap:18px; margin-top:10px; }

  .opts label { display:inline; margin:0; }

  button { border:0; border-radius:4px; padding:10px 20px; font-weight:600;

           cursor:pointer; font-size:14px; }

  #startBtn { background:var(--accent); color:#fff; }

  #startBtn:hover { background:var(--accent2); }

  #startBtn:disabled { background:#9db4d8; cursor:not-allowed; }

  #stopBtn { background:var(--red); color:#fff; margin-left:10px; }

  #stopBtn:disabled { background:#e3a5a5; cursor:not-allowed; }

  #log { background:#0f172a; color:#e2e8f0; font:12px Consolas, monospace;

         border-radius:6px; padding:12px; height:260px; overflow-y:auto;

         white-space:pre-wrap; }

  #summary { margin:10px 0; font-weight:600; }

  #summary.ok { color:var(--green); } #summary.err { color:var(--red); }

  footer { text-align:center; color:var(--muted); font-size:11px; padding:14px; }

  .hint { font-size:12px; color:var(--muted); margin-top:4px; }

</style>

</head>

<body>

<header>

  {% if alten_logo %}

    <img src="/static/alten_logo.png" class="logo-left" alt="ALTEN">

  {% elif alten_logo_url %}

    <img src="{{ alten_logo_url }}" class="logo-left" alt="ALTEN">

  {% else %}

    <div class="alten-mark">&#9650;</div><span class="alten-word">ALTEN</span>

  {% endif %}

  <div class="titles">

    <h1>JETPULL</h1>

  </div>

  {% if rr_logo %}

    <img src="/static/rr_logo.png" class="logo-right" alt="Rolls-Royce">

  {% elif rr_logo_url %}

    <img src="{{ rr_logo_url }}" class="logo-right" alt="Rolls-Royce">

  {% else %}

    <div class="rr-badge"><span>RR</span>ROLLS-ROYCE</div>

  {% endif %}

</header>


<main>

  <form id="form">

    <fieldset>

      <legend>Portal</legend>

      <div class="portals">

        <label><input type="checkbox" name="portal_bdc" id="bdc"> BDC (BlueData Connect)</label>

        <label><input type="checkbox" name="portal_b2b" id="b2b_tick"> B2B (Maximo)</label>

        <label class="soon"><input type="checkbox" disabled> Maximo <small>(coming soon)</small></label>

        <label class="soon"><input type="checkbox" disabled> One Stop Portal <small>(coming soon)</small></label>

      </div>

      <p class="hint">Tick BDC to use the built-in BlueData link with automatic

         LegacyTV / TV / RepeaterTV routing by TV number. Leave everything

         unticked to use this as a universal downloader with your own URL.</p>

    </fieldset>

 

    <fieldset>

      <legend>Step 1 — Website &amp; login</legend>

      <label>Website URL <span id="urlNote">*</span></label>

      <input type="text" name="url" id="url"

             placeholder="https://portal.example.rolls-royce.com">

      <p class="hint">Sign-in is manual. The tool opens the portal and then
         pauses for 60 seconds &mdash; sign in and approve the Authenticator
         prompt in that time, and the download starts automatically. The tool
         never stores or types your password. Change
         <code>manual_login_wait_seconds</code> in config.json if you need
         longer.</p>

    </fieldset>

 

    <fieldset>

      <legend>Step 2 — TV / search numbers</legend>

      <label>Type numbers (separate with commas)</label>

      <textarea name="numbers" rows="2"

                placeholder="512345, 612345, 698765"></textarea>

      <label>…or upload Excel (numbers in column A)</label>

      <input type="file" name="excel" accept=".xlsx,.xlsm">

    </fieldset>

 

    <fieldset id="parent_history_section" style="display:none;">

      <legend>Step 2.1 — Parent History</legend>

 

      <label>

        Upload Parent History Excel

        (Event Item Serial Number + Event Date TIme)

      </label>

 

      <input

          type="file"

          name="parent_history_excel"

          accept=".xlsx,.xlsm,.xls"

      >

 

      <div class="hint">

        For Parent History Dates.

      </div>

    </fieldset>

 

 

    <fieldset>

      <legend>Step 3 — Output &amp; options</legend>

      <label>Download folder on THIS PC * (a sub-folder is created per TV number; zips are auto-unpacked)</label>

      <div style="display:flex; gap:8px;">

        <input type="text" name="folder" id="folder" style="flex:1"

               placeholder="e.g. C:\Users\you\Downloads\TV_Downloads">

        <button type="button" id="browseBtn" style="background:#64748b;color:#fff;white-space:nowrap;">Browse…</button>

      </div>

      <p class="hint" id="browseHint">Browse opens the standard Windows

        &ldquo;Select Folder&rdquo; dialog on this PC and fills in the real

        path. You can also paste a full path directly.</p>

      <div class="opts">

        <label><input type="checkbox" name="headless"> Headless (no browser window)</label>

        <label><input type="checkbox" name="skip_existing" checked>

               Skip numbers already downloaded <small>(resume large batches)</small></label>

        <span class="hint">Headless sign-in: the Authenticator number appears in the log below.

        Your session is remembered, so after the first sign-in later runs need no approval.</span>

        <label>Browser:

          <select name="browser" style="width:auto">

            <option value="edge" selected>edge</option>

            <option value="chrome">chrome</option>

          </select>

        </label>

      </div>

    </fieldset>

 

    <button type="submit" id="startBtn">&#9654;&nbsp; Start download</button>

    <button type="button" id="stopBtn" disabled>&#9632;&nbsp; Stop after current item</button>

  </form>

 

  <div id="progressWrap" style="display:none; margin:10px 0;">

    <div style="display:flex; justify-content:space-between; font-size:12px;

                color:var(--muted); margin-bottom:3px;">

      <span id="progText">Starting…</span><span id="progPct">0%</span>

    </div>

    <div style="background:#dbe3f3; border-radius:6px; height:12px; overflow:hidden;">

      <div id="progBar" style="background:var(--accent); height:100%; width:0%;

           transition:width .3s;"></div>

    </div>

  </div>

 

  <div id="summary"></div>

  <div id="results"></div>

  <div id="log">Ready.</div>

</main>

<footer>&copy; 2026 Alten-Rolls-Royce. All rights reserved.

        Confidential &ndash; Internal Use Only.</footer>

<script>

var form = document.getElementById('form');

var startBtn = document.getElementById('startBtn');

var stopBtn = document.getElementById('stopBtn');

var logBox = document.getElementById('log');

var summary = document.getElementById('summary');

var resultsBox = document.getElementById('results');

var bdc = document.getElementById('bdc');

var urlInput = document.getElementById('url');

var urlNote = document.getElementById('urlNote');

var folderInput = document.getElementById('folder');

var browseBtn = document.getElementById('browseBtn');

var cursor = 0, timer = null;

 

var b2bTick = document.getElementById('b2b_tick');

 

function updatePortalUI() {

  var anyPreset = bdc.checked || b2bTick.checked;

  urlInput.disabled = anyPreset;

 

  var parentSection =

      document.getElementById(

          'parent_history_section'

      );

 

  parentSection.style.display =

      b2bTick.checked ? '' : 'none';

 

  if (anyPreset) {

    urlInput.value = '';

    urlInput.placeholder = bdc.checked

      ? 'Auto: BlueData Connect (LegacyTV / TV / RepeaterTV by number)'

      : 'Auto: Maximo portal (built-in link)';

    urlNote.textContent = '(handled by preset)';

  } else {

    urlInput.placeholder = 'https://portal.example.rolls-royce.com';

    urlNote.textContent = '*';

  }

}

bdc.addEventListener('change', function () {

  if (bdc.checked) b2bTick.checked = false;

  updatePortalUI();

});

b2bTick.addEventListener('change', function () {

  if (b2bTick.checked) bdc.checked = false;

  updatePortalUI();

});

// Browse: ask the LOCAL Flask server to open the native Windows folder

// dialog. The server has real disk access; the browser does not. This

// returns the true absolute path (e.g. C:\Users\...\TV_Downloads).

browseBtn.addEventListener('click', function () {

  var original = browseBtn.textContent;

  browseBtn.textContent = 'Choose in dialog…';

  browseBtn.disabled = true;

  fetch('/pick_folder', { method: 'POST' })

    .then(function (r) { return r.json(); })

    .then(function (j) {

      if (j.ok && j.path) { folderInput.value = j.path; folderInput.focus(); }

      else if (j.error) { alert('Could not open folder dialog: ' + j.error); }

    })

    .catch(function (err) { alert('Folder dialog failed: ' + err); })

    .finally(function () {

      browseBtn.textContent = original; browseBtn.disabled = false;

    });

});

 

function portalName() {

  return bdc.checked ? 'BDC (BlueData Connect)' : (b2bTick.checked ? 'B2B (Maximo)' : 'the target website');

}

 

form.addEventListener('submit', function (e) {

  e.preventDefault();

  // Pre-start reminder popup, portal-aware.

  var proceed = confirm(

    'Please ensure you are already signed in to ' + portalName() +

    ' in Edge before continuing.\n\n' +

    'This tool does not store or enter passwords — complete any login/MFA ' +

    'yourself. Once signed in, click OK to start.');

  if (!proceed) return;

 

  fetch('/start', { method: 'POST', body: new FormData(form) })

    .then(function (r) { return r.json(); })

    .then(function (j) {

      if (!j.ok) { alert(j.error); return; }

      logBox.textContent = ''; summary.textContent = '';

      resultsBox.innerHTML = ''; cursor = 0;

      startBtn.disabled = true; stopBtn.disabled = false;

      timer = setInterval(poll, 1000);

    })

    .catch(function (err) { alert('Request failed: ' + err); });

});

 

stopBtn.addEventListener('click', function () {

  fetch('/stop', { method: 'POST' });

});

 

function renderResults(results) {

  if (!results || !results.length) { resultsBox.innerHTML = ''; return; }

  var rows = results.map(function (r) {

    var color, label;

    if (r.status === 'SUCCESS') {

      color = '#059669';

      label = r.files.length + ' file(s)';

    } else if (r.status === 'SKIPPED') {

      color = '#64748b';

      label = 'Skipped (already downloaded)';

    } else if (r.reason === 'NUMBER_NOT_FOUND') {

      color = '#dc2626'; label = 'Number not found';

    } else if (r.reason === 'FILE_NOT_AVAILABLE') {

      color = '#dc2626'; label = 'File not available';

    } else {

      color = '#d97706'; label = r.message || 'Error';

    }

    return '<tr>' +

      '<td style="padding:4px 8px;font-family:Consolas,monospace;">' + r.number + '</td>' +

      '<td style="padding:4px 8px;color:' + color + ';font-weight:600;">' + label + '</td>' +

      '<td style="padding:4px 8px;color:#475569;font-size:12px;">' +

        (r.files.join(', ') || '') + '</td></tr>';

  }).join('');

  resultsBox.innerHTML =

    '<table style="width:100%;border-collapse:collapse;background:#fff;' +

    'border:1px solid #c8d4e8;border-radius:6px;margin:8px 0;">' +

    '<thead><tr style="background:#1a4fad;color:#fff;text-align:left;">' +

    '<th style="padding:6px 8px;">Number</th>' +

    '<th style="padding:6px 8px;">Status</th>' +

    '<th style="padding:6px 8px;">Files</th></tr></thead><tbody>' +

    rows + '</tbody></table>';

}

 

function poll() {

  fetch('/status?after=' + cursor)

    .then(function (r) { return r.json(); })

    .then(function (j) {

      if (j.logs.length) {

        j.logs.forEach(function (line) {

          logBox.textContent += line + '\n';

        });

        // Keep the log panel fast on very large batches: cap displayed lines.

        var lines = logBox.textContent.split('\n');

        if (lines.length > 1200) {

          logBox.textContent = '… (earlier lines trimmed; full history is in '

            + 'the logs folder) …\n' + lines.slice(-1000).join('\n');

        }

        logBox.scrollTop = logBox.scrollHeight;

        cursor = j.next;

      }

      // Progress bar for large batches

      if (j.progress && j.progress.total) {

        var p = j.progress;

        var pct = Math.round((p.done / p.total) * 100);

        document.getElementById('progressWrap').style.display = '';

        document.getElementById('progBar').style.width = pct + '%';

        document.getElementById('progPct').textContent = pct + '%';

        document.getElementById('progText').textContent =

          'Item ' + (p.done + 1) + ' of ' + p.total +

          (p.current ? '  —  ' + p.current : '');

      }

      if (!j.running) {

        clearInterval(timer); timer = null;

        startBtn.disabled = false; stopBtn.disabled = true;

        document.getElementById('progBar').style.width = '100%';

        document.getElementById('progPct').textContent = '100%';

        renderResults(j.results);

        if (j.summary) {

          summary.textContent = j.summary +

            (j.report ? '  Report: ' + j.report : '');

          summary.className =

            j.summary.indexOf('0/') === -1 ? 'ok' : 'err';

        }

        if (j.done) {

          var ok = j.done.ok, total = j.done.total;

          var skipped = j.done.skipped || 0;

          var failed = total - ok - skipped;

          var msg;

          if (failed === 0 && skipped === 0) {

            msg = 'Success! ' + ok + '/' + total + ' downloaded.';

          } else {

            msg = 'Completed: ' + ok + '/' + total + ' downloaded'

                + (skipped ? ', ' + skipped + ' skipped (already present)' : '')

                + (failed ? ', ' + failed + ' failed' : '')

                + '. See the table for details.';

          }

          setTimeout(function () { alert(msg); }, 200);

        }

      }

    });

}

 


 

</script>

</body>

</html>"""


if __name__ == "__main__":

    (Path(__file__).parent / "static").mkdir(exist_ok=True)

    threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.5:{PORT}")).start()

    print(f"JETPULL UI: http://127.0.0.5:{PORT}")

    print(f"Usage DB: {USAGE_DB_PATH}")

    app.run(host="127.0.0.5", port=PORT, debug=False, threaded=True)
