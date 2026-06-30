import os, json, uuid, re, csv, io
from datetime import datetime, date, timedelta
from urllib.parse import quote
import streamlit as st
from streamlit_echarts import st_echarts
from werkzeug.security import generate_password_hash, check_password_hash
from supabase import create_client, Client

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG / CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
SUPABASE_URL = "https://mvoxhdbcxmmozulenlvh.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im12b3hoZGJjeG1tb3p1bGVubHZoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5MTczNDIsImV4cCI6MjA5NDQ5MzM0Mn0.6Fhrqo6sfMnO3KklN5dwLup0BVbp0_ga8k5hi3LgXQU"

@st.cache_resource
def get_supabase() -> Client:
    try:
        sb = create_client(SUPABASE_URL, SUPABASE_KEY)
        sb.table("users").select("email").limit(1).execute()
        return sb
    except Exception as e:
        st.error(f"❌ Supabase connection failed: {e}")
        st.stop()

STATUSES   = ["New Idea","Assigned","WIP","UAT","Completed","Hold/Park","Rejected"]
PROJECTS   = ["EFS CA-MRO","EFS BA-MRO","EFS BA-LCE","EFS CA-LCE","EFS Controls","EFS Technical Response","Others"]
CATEGORIES = ["Customer Requirement","Internal"]
AUTO_CATS  = ["Automation-Personal Productivity","Automation-Process Improvement","Automation-Defined Product and Sales","Automation-Quality Enhancement","AI-Personal Productivity","AI-Process Improvement","AI-Defined Product and Sales"]
FREQ_MULT  = {"Daily":260,"Weekly":52,"Monthly":12,"Yearly":1}
REJ_REASONS= ["Technical Rejection","Business Rejection"]
ROLES_LIST = ["super user","normal user","automation engineer","automation pl","pl/spl"]
DEFAULT_PW = "admin123"

SUPPORT_NAME  = "Manoj JAGADEESH, Raja AMMAIAPPAN, Naveen KONNUR"
SUPPORT_EMAIL = "manoj.jagadeesh@alten-india.com"

ALTEN_LOGO_URL = "https://www.alten.com/wp-content/uploads/2019/01/favicon-alten.png"

CUSTOMERS = ["Rolls-Royce"]
REGIONS   = ["INDIA", "UK", "USA", "Germany"]

BLOCKED_DOMAINS = {
    "gmail.com","yahoo.com","hotmail.com","rediff.com","outlook.com",
    "live.com","icloud.com","aol.com","protonmail.com","yandex.com",
    "mail.com","zoho.com","gmx.com","inbox.com",
}

ROLE_PAGES = {
    "super user":         ["Dashboard","Submit Idea","PL Assignment","Feasibility","Approval","Admin","Email"],
    "normal user":        ["Submit Idea"],
    "automation engineer":["Dashboard","Submit Idea","Feasibility"],
    "automation pl":      ["Dashboard","Submit Idea","PL Assignment","Feasibility","Approval"],
    "pl/spl":             ["Dashboard","Submit Idea","Approval"],
}
PW_ROLES = {"super user","automation engineer","automation pl","pl/spl"}

DEFAULT_USERS = [{"email":"ravi.manoharan@alten-india.com","role":"super user"}]

AUTO_CAT_COLORS = {
    "Automation-Personal Productivity":"#1a4fad",
    "Automation-Process Improvement":"#7c3aed",
    "Automation-Defined Product and Sales":"#059669",
    "Automation-Quality Enhancement":"#0d9488",
    "AI-Personal Productivity":"#0369a1",
    "AI-Process Improvement":"#9333ea",
    "AI-Defined Product and Sales":"#0891b2",
}
CAT_COLORS = {"Customer Requirement":"#0623E3","Internal":"#0ee95e"}

STATUS_COLORS = {
    "New Idea":"#1a4fad","Assigned":"#7c3aed","WIP":"#0d9488",
    "UAT":"#0ea5e9","Completed":"#059669","Hold/Park":"#b45309","Rejected":"#dc2626"
}

STATUS_ICONS = {
    "New Idea":"💡","Assigned":"📋","WIP":"⚙️",
    "UAT":"🧪","Completed":"✅","Hold/Park":"⏸","Rejected":"❌"
}

THEMES = {
    "ALTEN Red & Blue":   {"primary":"#E30613","secondary":"#00AEEF","bg":"#f5f6fa","sidebar":"#0a0a0a"},
    "Ocean Blue":         {"primary":"#1a4fad","secondary":"#0ea5e9","bg":"#f0f4ff","sidebar":"#0d1b3e"},
    "Forest Green":       {"primary":"#059669","secondary":"#0d9488","bg":"#f0fdf4","sidebar":"#052e16"},
    "Purple Haze":        {"primary":"#7c3aed","secondary":"#a855f7","bg":"#faf5ff","sidebar":"#2e1065"},
    "Midnight Dark":      {"primary":"#e2e8f0","secondary":"#94a3b8","bg":"#0f172a","sidebar":"#020617"},
}

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE  (Supabase backend)
# ══════════════════════════════════════════════════════════════════════════════
def _run_sql(sb, sql):
    try:
        sb.rpc("pg_query", {"query": sql}).execute()
    except Exception:
        pass

def init_db():
    sb = get_supabase()
    _run_sql(sb, """
        CREATE TABLE IF NOT EXISTS users (
            email         text PRIMARY KEY,
            role          text,
            password_hash text
        );
    """)
    _run_sql(sb, """
        CREATE TABLE IF NOT EXISTS ideas (
            id                   text PRIMARY KEY,
            name                 text,
            submitter_email      text,
            idea_name            text,
            idea                 text,
            project              text,
            category             text,
            automation_category  text,
            pl_name              text,
            status               text,
            roi                  float8,
            assigned_engineer    text,
            feasibility_data     text,
            feasibility_comments text,
            decision             text,
            rejection_reason     text,
            approval_comment     text,
            priority_label       text,
            sprint_start         text,
            sprint_end           text,
            delivery_date        text,
            vsm_meeting_date     text,
            sprint_meeting_date  text,
            hold_reason          text,
            customer             text,
            region               text,
            created_date         text,
            assigned_date        text,
            wip_date             text,
            uat_date             text,
            completion_date      text
        );
    """)
    idea_cols = [
        ("submitter_email","text"),("automation_category","text"),
        ("priority_label","text"),("sprint_start","text"),("sprint_end","text"),
        ("delivery_date","text"),("vsm_meeting_date","text"),
        ("sprint_meeting_date","text"),("hold_reason","text"),
        ("customer","text"),("region","text"),("parent_id","text"),
    ]
    for col, dtype in idea_cols:
        _run_sql(sb, f"ALTER TABLE ideas ADD COLUMN IF NOT EXISTS {col} {dtype};")
    _run_sql(sb, "ALTER TABLE users ADD COLUMN IF NOT EXISTS password_hash text;")

    dh = generate_password_hash(DEFAULT_PW)
    for u in DEFAULT_USERS:
        try:
            existing = sb.table("users").select("email,password_hash").eq("email", u["email"].lower()).execute()
            if not existing.data:
                sb.table("users").insert({"email":u["email"].lower(),"role":u["role"],"password_hash":dh}).execute()
            elif not existing.data[0].get("password_hash"):
                sb.table("users").update({"password_hash":dh}).eq("email",u["email"].lower()).execute()
        except Exception:
            pass

def get_all():
    sb   = get_supabase()
    resp = sb.table("ideas").select("*").order("created_date", desc=True).execute()
    rows = []
    for r in (resp.data or []):
        try: r["feasibility_data"] = json.loads(r.get("feasibility_data") or "{}")
        except: r["feasibility_data"] = {}
        rows.append(r)
    return rows

def add_idea(idea):
    row = {
        "id":idea["id"],"name":idea.get("name",""),"submitter_email":idea.get("submitter_email",""),
        "idea_name":idea.get("idea_name",""),"idea":idea.get("idea",""),"project":idea.get("project",""),
        "category":idea.get("category",""),"automation_category":idea.get("automation_category",""),
        "pl_name":idea.get("pl_name",""),"status":idea.get("status","New Idea"),"roi":idea.get("roi",0),
        "assigned_engineer":idea.get("assigned_engineer",""),
        "feasibility_data":json.dumps(idea.get("feasibility_data",{})),
        "feasibility_comments":idea.get("feasibility_comments",""),"decision":idea.get("decision",""),
        "rejection_reason":idea.get("rejection_reason",""),"approval_comment":idea.get("approval_comment",""),
        "priority_label":idea.get("priority_label",""),"sprint_start":idea.get("sprint_start",""),
        "sprint_end":idea.get("sprint_end",""),"delivery_date":idea.get("delivery_date",""),
        "vsm_meeting_date":idea.get("vsm_meeting_date",""),"sprint_meeting_date":idea.get("sprint_meeting_date",""),
        "hold_reason":idea.get("hold_reason",""),"created_date":datetime.now().strftime("%Y-%m-%d %H:%M"),
        "assigned_date":"","wip_date":"","uat_date":"","completion_date":"",
        "customer":idea.get("customer",""),"region":idea.get("region",""),
    }
    get_supabase().table("ideas").insert(row).execute()

def update_idea(iid, fields):
    payload = {}
    for k,v in fields.items():
        payload[k] = json.dumps(v) if k=="feasibility_data" else v
    get_supabase().table("ideas").update(payload).eq("id",iid).execute()

def get_users():
    resp = get_supabase().table("users").select("*").order("email").execute()
    return resp.data or []

def add_user(email, role):
    sb  = get_supabase()
    dh  = generate_password_hash(DEFAULT_PW)
    existing = sb.table("users").select("password_hash").eq("email",email.lower()).execute()
    if existing.data:
        sb.table("users").update({"role":role}).eq("email",email.lower()).execute()
    else:
        sb.table("users").insert({"email":email.lower(),"role":role,"password_hash":dh}).execute()

def delete_user(email):
    get_supabase().table("users").delete().eq("email",email.lower()).execute()

def update_role(email, role):
    get_supabase().table("users").update({"role":role}).eq("email",email.lower()).execute()

def set_password(email, new_pw):
    get_supabase().table("users").update({"password_hash":generate_password_hash(new_pw)}).eq("email",email.lower()).execute()

def reset_password(email):
    set_password(email, DEFAULT_PW)

# ══════════════════════════════════════════════════════════════════════════════
#  DATE / SPRINT ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def next_workday(d):
    while d.weekday() >= 5: d += timedelta(days=1)
    return d

def add_workdays(d, n):
    d = next_workday(d); count=0
    while count < n:
        d += timedelta(days=1)
        if d.weekday() < 5: count+=1
    return d

def sprint_dates(start):
    s = next_workday(start)
    return s, add_workdays(s, 9)

def fmt_d(d): return d.strftime("%Y-%m-%d") if d else ""
def parse_d(s):
    for f in ("%Y-%m-%d %H:%M","%Y-%m-%d"):
        try: return datetime.strptime(str(s).strip(), f).date()
        except: pass
    return None

# ══════════════════════════════════════════════════════════════════════════════
#  PRIORITY ENGINE
# ══════════════════════════════════════════════════════════════════════════════
CAT_RANK = {"Customer Requirement":0,"Internal":1}

def priority_key(i):
    return (CAT_RANK.get(i.get("category",""),2), -float(i.get("roi",0) or 0), i.get("created_date",""))

def rank_ideas(ideas):
    ordered = sorted(ideas, key=priority_key)
    for idx,i in enumerate(ordered,1):
        cat = i.get("category","")
        roi = float(i.get("roi",0) or 0)
        label = ("🔴 P1-Customer" if cat=="Customer Requirement" else "🟡 P2-Internal")
        if roi: label += f" (ROI {roi:.1f})"
        label += f" #{idx}"
        i["priority_rank"] = idx; i["priority_label"] = label
    return ordered

def compute_delivery(all_ideas, eng_email, new_idea):
    active = [i for i in all_ideas
              if i.get("assigned_engineer")==eng_email
              and i.get("status") in {"Assigned","WIP","UAT","Hold/Park"}
              and i.get("id")!=new_idea.get("id")]
    combined = rank_ideas(active + [new_idea])
    today = next_workday(date.today())
    cursor = today; result = None
    for i in combined:
        s,e = sprint_dates(cursor)
        if i.get("id")==new_idea.get("id"):
            result = {"sprint_start":s,"sprint_end":e,
                      "priority_rank":i["priority_rank"],"priority_label":i["priority_label"]}
        cursor = e
    return result

def engineer_queue(all_ideas, eng_email):
    active = rank_ideas([i for i in all_ideas
                         if i.get("assigned_engineer")==eng_email
                         and i.get("status") in {"Assigned","WIP","UAT","Hold/Park"}])
    cursor = next_workday(date.today()); out=[]
    for i in active:
        s,e = sprint_dates(cursor)
        out.append({**i,"sprint_start":s,"sprint_end":e}); cursor=e
    return out

# ══════════════════════════════════════════════════════════════════════════════
#  KANBAN PARENT/CHILD RELATIONSHIP ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def set_parent(child_id, parent_id):
    """Link child_id under parent_id. Pass parent_id=None/'' to make standalone."""
    if child_id == parent_id:
        return False  # can't parent itself
    update_idea(child_id, {"parent_id": parent_id or ""})
    return True

def get_children(ideas, parent_id):
    return [i for i in ideas if i.get("parent_id") == parent_id]

def get_standalone_and_parents(ideas):
    """Ideas with no parent (either true standalone, or a parent of others)."""
    return [i for i in ideas if not i.get("parent_id")]

def parent_summary(ideas, parent_id):
    children = get_children(ideas, parent_id)
    if not children:
        return None
    total = len(children)
    completed = len([c for c in children if c.get("status") == "Completed"])
    pct = round(completed / total * 100, 1) if total else 0
    roi_sum = round(sum(float(c.get("roi", 0) or 0) for c in children), 2)
    hrs_sum = round(sum(idea_hours(c) for c in children), 1)
    status_counts = {}
    for c in children:
        s = c.get("status", "")
        status_counts[s] = status_counts.get(s, 0) + 1
    return {
        "count": total, "completion_pct": pct,
        "roi": roi_sum, "hours": hrs_sum, "status_counts": status_counts,
    }

# ══════════════════════════════════════════════════════════════════════════════
#  OUTLOOK COMPOSE LINK
# ══════════════════════════════════════════════════════════════════════════════
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
def is_email(e): return bool(e) and bool(EMAIL_RE.match(str(e).strip()))

def outlook_link(to_list, subject, body):
    to = ",".join(e for e in to_list if is_email(e))
    return (f"https://outlook.office.com/mail/deeplink/compose?"
            f"to={quote(to)}&subject={quote(subject)}&body={quote(body)}")

def build_assign_outlook(idea, eng, queue_info):
    body = f"""An idea has been assigned to you for Feasibility Study.

Idea Name   : {idea.get('idea_name','')}
Project     : {idea.get('project','')}
Category    : {idea.get('category','')}
Description : {idea.get('idea','')}
PL / SPL    : {idea.get('pl_name','')}
Priority    : {queue_info.get('priority_label','') if queue_info else ''}"""
    if queue_info:
        body += f"\nSprint Start: {fmt_d(queue_info['sprint_start'])}\nDelivery    : {fmt_d(queue_info['sprint_end'])}"
    body += "\n\nPlease begin Feasibility Study and update status in Turbo Drive."
    return outlook_link([eng], f"[Turbo Drive] New Idea Assigned: {idea.get('idea_name','')}", body)

def build_feasibility_outlook(idea, roi, vsm_date, queue_info):
    body = f"""Feasibility Study is complete — please review and provide GO / NO-GO decision.

Idea Name   : {idea.get('idea_name','')}
Project     : {idea.get('project','')}
Category    : {idea.get('category','')}
Engineer    : {idea.get('assigned_engineer','')}
ROI         : {round(roi,2)}
Priority    : {idea.get('priority_label','')}
VSM Date    : {fmt_d(vsm_date)} at 11:00 AM"""
    if queue_info:
        body += f"\nDelivery    : {fmt_d(queue_info['sprint_end'])}"
    body += "\n\nPlease log in to Turbo Drive to approve or reject."
    return outlook_link([idea.get("pl_name","")], f"[Turbo Drive] Feasibility Complete — Action Needed: {idea.get('idea_name','')}", body)

def build_meeting_outlook(mtype, idea, mdate, recipients):
    times  = {"vsm":"11:00 AM","sprint":"10:00 AM","delivery":"03:00 PM"}
    titles = {"vsm":"VSM Session","sprint":"Sprint Planning","delivery":"Delivery/Demo Review"}
    descs  = {
        "vsm":      "A Value Stream Mapping session to review current vs future-state process before GO/NO-GO.",
        "sprint":   "Sprint Planning for the 2-week delivery cycle.",
        "delivery": "Delivery / Demo review — please review the output and confirm acceptance.",
    }
    body = f"""{descs.get(mtype,'')}

Idea Name   : {idea.get('idea_name','')}
Project     : {idea.get('project','')}
Category    : {idea.get('category','')}
Date        : {fmt_d(mdate)} at {times.get(mtype,'')}
Engineer    : {idea.get('assigned_engineer','-')}
PL / SPL    : {idea.get('pl_name','-')}"""
    return outlook_link(recipients, f"[Turbo Drive] {titles.get(mtype,'Meeting')}: {idea.get('idea_name','')} — {fmt_d(mdate)}", body)

# ══════════════════════════════════════════════════════════════════════════════
#  SHARED UI COMPONENTS
# ══════════════════════════════════════════════════════════════════════════════
def render_support_bar():
    st.markdown(
        f'<p style="font-size:11px;color:#94a3b8;margin-top:-10px;margin-bottom:10px;">'
        f'For any queries, reach out to <b>{SUPPORT_NAME}</b> — '
        f'<a href="mailto:{SUPPORT_EMAIL}" style="color:#00AEEF;">{SUPPORT_EMAIL}</a></p>',
        unsafe_allow_html=True
    )

def render_copyright():
    st.markdown("---")
    st.markdown(
        '<p style="text-align:center;font-size:11px;color:#94a3b8;padding:4px 0 8px;">'
        '© 2025 ALTEN Engineering Services. All rights reserved. '
        'Turbo Drive is an internal tool developed and maintained by ALTEN India.'
        '</p>',
        unsafe_allow_html=True
    )

def page_header(title: str):
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:4px;">
      <img src="{ALTEN_LOGO_URL}" style="height:28px;object-fit:contain;" alt="ALTEN"/>
      <span style="font-size:24px;font-weight:800;color:#E30613;letter-spacing:0.5px;">
         {title}
      </span>
    </div>""", unsafe_allow_html=True)
    render_support_bar()

# ══════════════════════════════════════════════════════════════════════════════
#  THEME / CSS INJECTION  (no forced page-bg override on any page)
# ══════════════════════════════════════════════════════════════════════════════
def apply_theme(theme_name):
    t = THEMES.get(theme_name, THEMES["ALTEN Red & Blue"])
    dark_bg = theme_name == "Midnight Dark"
    text_color = "#e2e8f0" if dark_bg else "#1e293b"
    surface = "#1e293b" if dark_bg else "#ffffff"
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html,body,[data-testid="stApp"]{{
        font-family:'Inter',sans-serif;
        background:{t['bg']} !important;
        color:{text_color};
        font-size:clamp(12px,1.1vw,15px);
    }}
    [data-testid="stSidebar"]{{background:{t['sidebar']} !important;}}
    [data-testid="stSidebar"] *{{color:#e2e8f0 !important;}}
    [data-testid="stSidebar"] .stRadio label{{
        font-size:clamp(11px,1vw,14px);padding:6px 10px;border-radius:8px;
        transition:background .15s;cursor:pointer;
    }}
    [data-testid="stSidebar"] .stRadio label:hover{{background:rgba(255,255,255,.1);}}
    h1{{
        background:linear-gradient(135deg,{t['primary']},{t['secondary']});
        -webkit-background-clip:text;background-clip:text;color:transparent;
        font-size:clamp(20px,2vw,28px);font-weight:800;margin-bottom:4px;
    }}
    h2{{color:{t['primary']};font-size:clamp(15px,1.5vw,20px);font-weight:700;}}
    h3{{color:{t['secondary']};font-size:clamp(13px,1.2vw,16px);font-weight:600;}}
    .kpi-card{{
        background:{surface};border-radius:14px;padding:14px 16px;
        border-left:5px solid {t['primary']};
        box-shadow:0 2px 12px rgba(0,0,0,.08);margin-bottom:6px;
    }}
    .kpi-val{{font-size:clamp(16px,1.6vw,22px);font-weight:800;}}
    .kpi-lbl{{font-size:clamp(9px,0.85vw,11px);color:#64748b;font-weight:500;margin-top:2px;}}
    .kpi-sub{{font-size:clamp(8px,0.75vw,10px);color:#94a3b8;margin-top:3px;}}
    .idea-card{{
        background:{surface};border-radius:12px;padding:14px 16px;
        border:1px solid #e2e8f0;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:10px;
    }}
    .outlook-btn{{
        display:inline-block;background:{t['primary']};color:#fff;
        padding:8px 18px;border-radius:8px;text-decoration:none;
        font-weight:600;font-size:clamp(11px,1vw,13px);margin-top:8px;
    }}
    .stButton>button{{
        background:linear-gradient(135deg,{t['primary']},{t['secondary']});
        color:#fff;border:none;border-radius:8px;font-weight:600;padding:8px 20px;
    }}
    .stButton>button:hover{{opacity:.88;}}
    div[data-testid="stForm"]{{background:{surface};border-radius:12px;padding:12px;}}
    .login-box{{
        max-width:440px;margin:40px auto;background:{surface};border-radius:16px;
        padding:32px 36px;box-shadow:0 8px 32px rgba(0,0,0,.12);
        border:1px solid rgba(255,255,255,.08);
    }}
    /* filter bar */
    .filter-bar{{
        background:{surface};border-radius:10px;padding:10px 14px;
        border:1px solid #e2e8f0;margin-bottom:10px;
        box-shadow:0 1px 4px rgba(0,0,0,.05);
    }}
    </style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  SESSION / AUTH HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def ss(key, default=None):
    return st.session_state.get(key, default)

def logged_in(): return bool(ss("email"))
def user_role(): return ss("role","")
def user_pages(): return ROLE_PAGES.get(user_role(),[])
def can(page): return page in user_pages()

def is_org_email(email: str) -> bool:
    if "@" not in email: return False
    domain = email.strip().lower().split("@")[-1]
    return domain not in BLOCKED_DOMAINS

def check_login(email, password):
    resp = get_supabase().table("users").select("*").eq("email",email.lower()).execute()
    if not resp.data: return None, "Email not found in system."
    row  = resp.data[0]
    role = row.get("role","")
    if role in PW_ROLES:
        if not password: return None, "Password required for your role."
        if not row.get("password_hash") or not check_password_hash(row["password_hash"], password):
            return None, "Incorrect password."
    return row, None

# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def idea_hours(i):
    fd = i.get("feasibility_data",{}) or {}
    try:
        return float(fd.get("manual",0) or 0)*float(fd.get("fte",0) or 0)*FREQ_MULT.get(fd.get("freq","Daily"),1)
    except: return 0

def kpi_card(value, label, color, sub="", icon=""):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{color}">
      <div style="font-size:clamp(14px,1.4vw,18px);margin-bottom:2px;">{icon}</div>
      <div class="kpi-val" style="color:{color}">{value}</div>
      <div class="kpi-lbl">{label}</div>
      {"<div class='kpi-sub'>"+sub+"</div>" if sub else ""}
    </div>""", unsafe_allow_html=True)

def cnt_cat_status(cat, st_, ideas):
    return len([i for i in ideas if i.get("category")==cat and i.get("status")==st_])

def cnt_cat_wip(ideas, cat):
    return len([i for i in ideas if i.get("category")==cat and i.get("status") in ("WIP","UAT")])

# ══════════════════════════════════════════════════════════════════════════════
#  PREMIUM KPI CARD — split illustration (left) + metrics (right)
# ══════════════════════════════════════════════════════════════════════════════
KPI_ILLUSTRATIONS = {
    "bulb": '''<svg viewBox="0 0 100 100" width="100%" height="100%">
        <defs><linearGradient id="g_bulb" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="#FFE066"/><stop offset="100%" stop-color="#F5A623"/>
        </linearGradient></defs>
        <circle cx="50" cy="40" r="26" fill="url(#g_bulb)" opacity="0.95"/>
        <rect x="40" y="62" width="20" height="10" rx="2" fill="#94a3b8"/>
        <rect x="42" y="74" width="16" height="6" rx="2" fill="#64748b"/>
        <path d="M50 22 L50 14 M34 30 L28 24 M66 30 L72 24" stroke="#F5A623" stroke-width="3" stroke-linecap="round"/>
        <path d="M42 40 L48 46 L58 32" stroke="#fff" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>''',
    "trophy": '''<svg viewBox="0 0 100 100" width="100%" height="100%">
        <defs><linearGradient id="g_troph" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#FFD700"/><stop offset="100%" stop-color="#D4A017"/>
        </linearGradient></defs>
        <path d="M35 20 H65 V40 C65 52 57 60 50 60 C43 60 35 52 35 40 Z" fill="url(#g_troph)"/>
        <path d="M35 24 C25 24 22 36 32 40" stroke="#D4A017" stroke-width="4" fill="none"/>
        <path d="M65 24 C75 24 78 36 68 40" stroke="#D4A017" stroke-width="4" fill="none"/>
        <rect x="46" y="60" width="8" height="12" fill="#D4A017"/>
        <rect x="38" y="72" width="24" height="8" rx="2" fill="#B8860B"/>
        <circle cx="50" cy="38" r="7" fill="#fff" opacity="0.5"/>
      </svg>''',
    "clock": '''<svg viewBox="0 0 100 100" width="100%" height="100%">
        <defs><linearGradient id="g_clock" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="#5EEAD4"/><stop offset="100%" stop-color="#0D9488"/>
        </linearGradient></defs>
        <circle cx="50" cy="50" r="32" fill="url(#g_clock)"/>
        <circle cx="50" cy="50" r="26" fill="#fff" opacity="0.15"/>
        <line x1="50" y1="50" x2="50" y2="30" stroke="#fff" stroke-width="3" stroke-linecap="round"/>
        <line x1="50" y1="50" x2="64" y2="56" stroke="#fff" stroke-width="3" stroke-linecap="round"/>
        <circle cx="50" cy="50" r="3" fill="#fff"/>
      </svg>''',
    "growth": '''<svg viewBox="0 0 100 100" width="100%" height="100%">
        <defs><linearGradient id="g_growth" x1="0" y1="1" x2="0" y2="0">
          <stop offset="0%" stop-color="#34D399"/><stop offset="100%" stop-color="#059669"/>
        </linearGradient></defs>
        <rect x="18" y="58" width="12" height="24" rx="2" fill="url(#g_growth)" opacity="0.6"/>
        <rect x="38" y="44" width="12" height="38" rx="2" fill="url(#g_growth)" opacity="0.8"/>
        <rect x="58" y="28" width="12" height="54" rx="2" fill="url(#g_growth)"/>
        <path d="M18 50 L40 36 L58 22 L80 12" stroke="#10B981" stroke-width="3" fill="none" stroke-linecap="round"/>
        <path d="M68 12 L80 12 L80 24" stroke="#10B981" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>''',
    "folder": '''<svg viewBox="0 0 100 100" width="100%" height="100%">
        <defs><linearGradient id="g_folder" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="#60A5FA"/><stop offset="100%" stop-color="#1A4FAD"/>
        </linearGradient></defs>
        <path d="M16 32 H42 L48 40 H84 V76 H16 Z" fill="url(#g_folder)"/>
        <rect x="16" y="40" width="68" height="36" fill="#1A4FAD" opacity="0.15"/>
        <rect x="28" y="52" width="44" height="4" rx="2" fill="#fff" opacity="0.7"/>
        <rect x="28" y="60" width="32" height="4" rx="2" fill="#fff" opacity="0.5"/>
      </svg>''',
}

def kpi_card_premium(illustration_key, value, label, sub, color):
    """Premium split-layout KPI card: left = illustration, right = metrics."""
    svg = KPI_ILLUSTRATIONS.get(illustration_key, KPI_ILLUSTRATIONS["bulb"])
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:10px;
         background:rgba(255,255,255,0.65);backdrop-filter:blur(10px);
         border-radius:16px;padding:12px 14px;margin-bottom:8px;
         border:1px solid rgba(255,255,255,0.4);
         box-shadow:0 4px 20px rgba(0,0,0,.08);
         border-left:4px solid {color};
         transition:transform .2s,box-shadow .2s;"
         onmouseover="this.style.transform='translateY(-3px)';this.style.boxShadow='0 8px 28px rgba(0,0,0,.14)'"
         onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 4px 20px rgba(0,0,0,.08)'">
      <div style="flex:0 0 38%;max-width:80px;">{svg}</div>
      <div style="flex:1;min-width:0;">
        <div style="font-size:clamp(16px,1.8vw,24px);font-weight:800;color:{color};line-height:1.1;">{value}</div>
        <div style="font-size:clamp(9px,0.9vw,11px);color:#475569;font-weight:600;margin-top:2px;">{label}</div>
        <div style="font-size:clamp(8px,0.78vw,10px);color:#94a3b8;margin-top:3px;line-height:1.3;">{sub}</div>
      </div>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ROBOT ILLUSTRATIONS — animated, pointing arm/gesture toward selected category
#  rotation_deg: angle applied to the pointing limb to "aim" at the chosen item
# ══════════════════════════════════════════════════════════════════════════════
def render_automation_robot(rotation_deg=0, glow_color="#1a4fad"):
    """Industrial robotic arm — base fixed, forearm rotates toward selection."""
    st.markdown(f"""
    <div style="display:flex;justify-content:center;padding:6px 0;">
      <svg viewBox="0 0 220 220" width="170" height="170">
        <defs>
          <radialGradient id="armGlow" cx="50%" cy="50%" r="50%">
            <stop offset="0%" stop-color="{glow_color}" stop-opacity="0.35"/>
            <stop offset="100%" stop-color="{glow_color}" stop-opacity="0"/>
          </radialGradient>
          <linearGradient id="armMetal" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stop-color="#CBD5E1"/><stop offset="100%" stop-color="#64748B"/>
          </linearGradient>
        </defs>
        <circle cx="110" cy="110" r="100" fill="url(#armGlow)"/>
        <!-- base -->
        <rect x="75" y="170" width="70" height="22" rx="6" fill="url(#armMetal)"/>
        <circle cx="110" cy="170" r="20" fill="#475569"/>
        <circle cx="110" cy="170" r="9" fill="{glow_color}">
          <animate attributeName="opacity" values="1;0.4;1" dur="1.6s" repeatCount="indefinite"/>
        </circle>
        <!-- rotating arm assembly -->
        <g transform="rotate({rotation_deg} 110 170)" style="transition:transform 0.8s cubic-bezier(.4,1.4,.6,1);">
          <rect x="104" y="90" width="12" height="80" rx="6" fill="url(#armMetal)"/>
          <circle cx="110" cy="90" r="10" fill="#334155"/>
          <rect x="106" y="40" width="8" height="52" rx="4" fill="{glow_color}" opacity="0.85"/>
          <circle cx="110" cy="38" r="9" fill="{glow_color}"/>
          <!-- gripper -->
          <path d="M100 30 L92 18 M120 30 L128 18" stroke="{glow_color}" stroke-width="5" stroke-linecap="round"/>
        </g>
      </svg>
    </div>""", unsafe_allow_html=True)

def render_ai_robot(rotation_deg=0, glow_color="#9333ea"):
    """Futuristic humanoid AI robot head — head/eye turns toward selection."""
    st.markdown(f"""
    <div style="display:flex;justify-content:center;padding:6px 0;">
      <svg viewBox="0 0 220 220" width="170" height="170">
        <defs>
          <radialGradient id="aiGlow" cx="50%" cy="50%" r="50%">
            <stop offset="0%" stop-color="{glow_color}" stop-opacity="0.35"/>
            <stop offset="100%" stop-color="{glow_color}" stop-opacity="0"/>
          </radialGradient>
          <linearGradient id="aiBody" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="#E2E8F0"/><stop offset="100%" stop-color="#94A3B8"/>
          </linearGradient>
        </defs>
        <circle cx="110" cy="110" r="100" fill="url(#aiGlow)"/>
        <!-- torso -->
        <rect x="75" y="130" width="70" height="60" rx="14" fill="url(#aiBody)"/>
        <circle cx="110" cy="160" r="12" fill="{glow_color}" opacity="0.8">
          <animate attributeName="opacity" values="0.5;1;0.5" dur="2s" repeatCount="indefinite"/>
        </circle>
        <!-- neck -->
        <rect x="102" y="112" width="16" height="20" fill="#64748B"/>
        <!-- rotating head -->
        <g transform="rotate({rotation_deg} 110 90)" style="transition:transform 0.8s cubic-bezier(.4,1.4,.6,1);">
          <rect x="78" y="58" width="64" height="54" rx="20" fill="url(#aiBody)"/>
          <rect x="92" y="76" width="14" height="14" rx="4" fill="{glow_color}">
            <animate attributeName="opacity" values="1;0.3;1" dur="1.4s" repeatCount="indefinite"/>
          </rect>
          <rect x="114" y="76" width="14" height="14" rx="4" fill="{glow_color}">
            <animate attributeName="opacity" values="1;0.3;1" dur="1.4s" repeatCount="indefinite"/>
          </rect>
          <rect x="96" y="96" width="28" height="4" rx="2" fill="#475569"/>
          <rect x="104" y="50" width="12" height="12" rx="6" fill="{glow_color}"/>
        </g>
      </svg>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  SESSION TIMEOUT (interaction-based — resets on every widget rerun)
#  True silent mouse/keyboard idle detection is not possible in pure Streamlit;
#  this tracks time since the last server-side interaction (button, dropdown,
#  form submit, page nav, etc.) which covers everything that actually causes
#  a rerun in this app.
# ══════════════════════════════════════════════════════════════════════════════
SESSION_TIMEOUT_SECONDS = 300   # 5 minutes
SESSION_WARNING_AT      = 240   # show warning after 4 minutes

def touch_activity():
    st.session_state["_last_activity"] = datetime.now()

def check_session_timeout():
    if not logged_in():
        return
    last = st.session_state.get("_last_activity")
    if last is None:
        touch_activity()
        return
    elapsed = (datetime.now() - last).total_seconds()

    if elapsed >= SESSION_TIMEOUT_SECONDS:
        for k in ["email","role","name","_last_activity"]:
            st.session_state.pop(k, None)
        st.session_state["_session_expired"] = True
        st.rerun()

    elif elapsed >= SESSION_WARNING_AT:
        remaining = int(SESSION_TIMEOUT_SECONDS - elapsed)
        st.warning(f"⚠️ Your session will expire in **{remaining}s** due to inactivity.")
        wc1, wc2 = st.columns(2)
        with wc1:
            if st.button("✅ Continue Session", use_container_width=True, key="_continue_session"):
                touch_activity(); st.rerun()
        with wc2:
            if st.button("🚪 Logout Now", use_container_width=True, key="_logout_now"):
                for k in ["email","role","name","_last_activity"]:
                    st.session_state.pop(k, None)
                st.rerun()

def render_session_countdown():
    """Live countdown shown in the sidebar."""
    last = st.session_state.get("_last_activity")
    if last is None:
        return
    elapsed   = (datetime.now() - last).total_seconds()
    remaining = max(0, int(SESSION_TIMEOUT_SECONDS - elapsed))
    mins, secs = divmod(remaining, 60)
    color = "#dc2626" if remaining < 60 else ("#b45309" if remaining < 90 else "#94a3b8")
    st.markdown(f"""
    <div style="text-align:center;font-size:10px;color:#94a3b8;margin-top:6px;">
      Session Timeout<br>
      <span style="font-size:15px;font-weight:700;color:{color};font-family:monospace;">
        {mins:02d}:{secs:02d}
      </span>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  KANBAN: NATIVE STREAMLIT COLUMN + EXPANDER BOARD
# ══════════════════════════════════════════════════════════════════════════════
def _render_idea_card_body(idea, ideas_all, status, color, indent=0):
    """Renders one card's detail body — used for both top-level and child cards."""
    eng      = idea.get("assigned_engineer") or ""
    proj     = idea.get("project") or "-"
    delivery = idea.get("delivery_date") or ""
    hold     = idea.get("hold_reason") or ""
    name     = idea.get("name") or "-"
    eng_name = eng.split("@")[0] if "@" in eng else (eng or "—")

    st.markdown(
        f'<div style="border-left:3px solid {color};padding-left:8px;margin-bottom:6px;">'
        f'<span style="font-size:clamp(9px,0.85vw,11px);color:#64748b;">📌 {proj}</span><br>'
        f'<span style="font-size:clamp(9px,0.85vw,11px);color:#64748b;">👤 {name}</span><br>'
        f'<span style="font-size:clamp(9px,0.85vw,11px);color:#64748b;">👷 {eng_name}</span>'
        +(f'<br><span style="font-size:clamp(8px,0.75vw,10px);color:#0369a1;">📅 {delivery}</span>' if delivery else "")
        +(f'<br><span style="font-size:clamp(8px,0.75vw,10px);color:#b45309;">⏸ {hold[:30]}</span>' if hold else "")
        +f'</div>', unsafe_allow_html=True,
    )
    new_status = st.selectbox("Move to", STATUSES,
                              index=STATUSES.index(status) if status in STATUSES else 0,
                              key=f"kanban_sel_{idea['id']}",
                              label_visibility="collapsed")
    hold_input = ""
    if new_status == "Hold/Park":
        hold_input = st.text_input("Reason *", key=f"kanban_hold_{idea['id']}", placeholder="Hold reason…")
    if st.button("Update", key=f"kanban_btn_{idea['id']}", use_container_width=True):
        if new_status=="Hold/Park" and not hold_input:
            st.error("Enter a hold reason.")
        else:
            upd = {"status":new_status}
            if new_status=="Completed":
                upd["completion_date"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            upd["hold_reason"] = hold_input if new_status=="Hold/Park" else ""
            update_idea(idea["id"], upd)
            st.rerun()

    # ── Parent/Child relationship controls ────────────────────────────────
    st.markdown("<div style='font-size:10px;color:#94a3b8;margin-top:6px;'>🔗 Relationship</div>", unsafe_allow_html=True)
    current_parent = idea.get("parent_id") or ""
    other_ideas = [i for i in ideas_all if i["id"] != idea["id"] and not i.get("parent_id")]
    parent_opts = ["— Standalone —"] + [f"{i.get('idea_name','(no name)')}" for i in other_ideas]
    parent_id_map = {i.get("idea_name","(no name)"): i["id"] for i in other_ideas}
    cur_label = "— Standalone —"
    for i in other_ideas:
        if i["id"] == current_parent:
            cur_label = i.get("idea_name","(no name)"); break
    sel_parent = st.selectbox("Set Parent", parent_opts,
                              index=parent_opts.index(cur_label) if cur_label in parent_opts else 0,
                              key=f"kanban_parent_{idea['id']}", label_visibility="collapsed")
    if st.button("🔗 Apply Relationship", key=f"kanban_parent_btn_{idea['id']}", use_container_width=True):
        new_parent_id = "" if sel_parent == "— Standalone —" else parent_id_map.get(sel_parent, "")
        if set_parent(idea["id"], new_parent_id):
            st.success("Relationship updated." if new_parent_id else "Converted to standalone card.")
            st.rerun()
        else:
            st.error("A card cannot be its own parent.")
    st.divider()

def render_kanban_board(ideas):
    cols = st.columns(len(STATUSES))
    for col, status in zip(cols, STATUSES):
        color = STATUS_COLORS.get(status,"#888")
        icon  = STATUS_ICONS.get(status,"")
        count = len([i for i in ideas if i.get("status")==status])
        col.markdown(
            f'<div style="background:{color};color:#fff;border-radius:8px;'
            f'padding:6px 10px;text-align:center;font-size:clamp(10px,0.9vw,12px);font-weight:700;'
            f'margin-bottom:6px;">{icon} {status} &nbsp;'
            f'<span style="background:rgba(255,255,255,0.25);border-radius:10px;'
            f'padding:1px 7px;">{count}</span></div>',
            unsafe_allow_html=True,
        )

    cols = st.columns(len(STATUSES))
    for col, status in zip(cols, STATUSES):
        color  = STATUS_COLORS.get(status,"#888")
        # Only top-level cards (no parent) appear directly in their status column.
        # A card that has children is rendered as a 📁 Parent card regardless of
        # where its children individually sit — children render nested inside it.
        bucket = [i for i in ideas if i.get("status")==status and not i.get("parent_id")]
        with col:
            if not bucket:
                st.caption("_Empty_")
            for idea in bucket:
                children = get_children(ideas, idea["id"])
                label = (idea.get("idea_name") or "No Name")[:26]

                if children:
                    summary = parent_summary(ideas, idea["id"])
                    parent_label = f"📁 {label}  ({summary['count']})"
                    with st.expander(parent_label, expanded=False):
                        st.markdown(f"""
                        <div style="background:rgba(0,0,0,0.03);border-radius:8px;padding:8px 10px;margin-bottom:8px;font-size:10px;">
                          <b>Children:</b> {summary['count']} &nbsp;|&nbsp;
                          <b>Completion:</b> {summary['completion_pct']}%<br>
                          <b>Hours Saved:</b> {summary['hours']} &nbsp;|&nbsp;
                          <b>ROI:</b> {summary['roi']}
                        </div>""", unsafe_allow_html=True)
                        _render_idea_card_body(idea, ideas, status, color)

                        st.markdown("<div style='font-size:10px;font-weight:700;color:#64748b;margin:4px 0;'>▼ Child Cards</div>", unsafe_allow_html=True)
                        for child in children:
                            child_color = STATUS_COLORS.get(child.get("status",""), "#888")
                            child_label = f"📄 {(child.get('idea_name') or 'No Name')[:22]} · {child.get('status','')}"
                            with st.expander(child_label, expanded=False):
                                _render_idea_card_body(child, ideas, child.get("status",""), child_color, indent=1)
                else:
                    with st.expander(f"📄 {label}", expanded=False):
                        _render_idea_card_body(idea, ideas, status, color)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def page_login():
    t = THEMES.get(ss("theme","ALTEN Red & Blue"), THEMES["ALTEN Red & Blue"])
    dark_bg = ss("theme","") == "Midnight Dark"
    surface = "#2a61b8" if dark_bg else "#CACDE3"

    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
        st.markdown(f"""
        <div style="background:{surface};border-radius:16px;padding:32px 36px;
             box-shadow:0 8px 32px rgba(0,0,0,.13);border:1px solid rgba(200,200,200,.2);
             margin-top:20px;">
          <div style="display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:6px;">
            <img src="{ALTEN_LOGO_URL}" style="height:50px;object-fit:contain;" alt="ALTEN"/>
            <span style="font-size:clamp(20px,2vw,26px);font-weight:900;
                 background:linear-gradient(135deg,{t['primary']},{t['secondary']});
                 -webkit-background-clip:text;background-clip:text;color:transparent;
                 letter-spacing:1px;"> TURBO DRIVE</span>
          </div>
          <div style="text-align:center;color:#64748b;font-size:clamp(11px,1vw,13px);margin-bottom:2px;">
            Ideation &amp; Automation Workflow Manager
          </div>
        </div>""", unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=False):
            email    = st.text_input("📧 Email (Username)", placeholder="you@company.com")
            password = st.text_input("🔒 Password", type="password",
                                     placeholder="Required for Admin / PL / Engineer roles")
            submitted = st.form_submit_button("🚀 Login", use_container_width=True)
            if submitted:
                if not email:
                    st.error("Please enter your email.")
                else:
                    row, err = check_login(email.strip(), password.strip())
                    if err:
                        st.error(f"🚫 {err}")
                    elif not ROLE_PAGES.get(row["role"]):
                        st.error("No pages configured for your role.")
                    else:
                        st.session_state["email"] = row["email"]
                        st.session_state["role"]  = row["role"]
                        st.session_state["name"]  = row["email"].split("@")[0].replace("."," ").title()
                        st.success("Welcome!")
                        st.rerun()

        col1, col2, col3 = st.columns(3)
        with col2:
            if st.button("🔑 Change Password"):
                st.session_state["_page_override"] = "change_password"; st.rerun()
        with col3:
            if st.button("📝 Register"):
                st.session_state["_page_override"] = "register"; st.rerun()

        st.markdown(
            f'<p style="font-size:clamp(9px,0.8vw,10px);text-align:center;color:#94a3b8;margin-top:6px;">'
            f'For queries: <a href="mailto:{SUPPORT_EMAIL}" style="color:#00AEEF;">'
            f'{SUPPORT_NAME} — {SUPPORT_EMAIL}</a></p>',
            unsafe_allow_html=True
        )
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: REGISTER
# ══════════════════════════════════════════════════════════════════════════════
def page_register():
    page_header("Create Account")
    st.caption("Only **organisational email addresses** are accepted (free providers like Gmail, Yahoo, Hotmail, Rediff are not allowed). Self-registered users get **Submit Idea** access; an Admin can upgrade your role.")
    with st.form("reg_form"):
        name  = st.text_input("Full Name")
        email = st.text_input("Work Email", placeholder="you@company.com")
        pw    = st.text_input("Password", type="password")
        cpw   = st.text_input("Confirm Password", type="password")
        if st.form_submit_button("🚀 Create Account", use_container_width=True):
            if not all([name,email,pw,cpw]):
                st.error("All fields are required.")
            elif not is_org_email(email.strip()):
                st.error("🚫 Free-mail providers (Gmail, Yahoo, Hotmail, Rediff, etc.) are not allowed. Please use your organisational email.")
            elif pw != cpw:
                st.error("Passwords do not match.")
            elif len(pw) < 4:
                st.error("Password must be at least 4 characters.")
            else:
                resp = get_supabase().table("users").select("email").eq("email",email.lower()).execute()
                if resp.data:
                    st.error("This email is already registered — please log in.")
                else:
                    get_supabase().table("users").insert({
                        "email":email.lower(),"role":"normal user","password_hash":generate_password_hash(pw)
                    }).execute()
                    st.success("✅ Registered! You can now log in.")
    if st.button("← Back to Login"):
        st.session_state.pop("_page_override",None); st.rerun()
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: CHANGE PASSWORD
# ══════════════════════════════════════════════════════════════════════════════
def page_change_password():
    page_header("Change Password")
    prefill = ss("email","")
    with st.form("cpw_form"):
        email  = st.text_input("Email", value=prefill)
        cur_pw = st.text_input("Current Password", type="password")
        new_pw = st.text_input("New Password", type="password")
        cnf_pw = st.text_input("Confirm New Password", type="password")
        if st.form_submit_button("Update Password", use_container_width=True):
            if not all([email,cur_pw,new_pw,cnf_pw]):
                st.error("All fields required.")
            elif new_pw != cnf_pw:
                st.error("New passwords do not match.")
            elif len(new_pw) < 4:
                st.error("Minimum 4 characters.")
            else:
                resp = get_supabase().table("users").select("*").eq("email",email.lower()).execute()
                if not resp.data:
                    st.error("Email not found.")
                elif not check_password_hash(resp.data[0].get("password_hash") or "", cur_pw):
                    st.error("🚫 Current password is incorrect.")
                else:
                    set_password(email.lower(), new_pw)
                    st.success("✅ Password changed successfully.")
    if st.button("← Back"):
        st.session_state.pop("_page_override",None); st.rerun()
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: SUBMIT IDEA
# ══════════════════════════════════════════════════════════════════════════════
def page_submit():
    page_header("Submit New Idea 💡")
    users     = get_users()
    pl_emails = [u["email"] for u in users if u["role"] in ("pl/spl","automation pl","super user")]
    with st.form("submit_form", clear_on_submit=True):
        name      = st.text_input("Your Full Name", value=ss("name",""))
        sub_email = st.text_input("Your Email", value=ss("email",""))
        idea_name = st.text_input("Idea Name *", placeholder="Short title for the idea")
        idea_desc = st.text_area("Idea Description *", placeholder="Describe the automation idea in detail")
        col1, col2 = st.columns(2)
        with col1:
            project  = st.selectbox("Project", PROJECTS)
            category = st.selectbox("Idea Category", CATEGORIES)
            customer = st.selectbox("Customer *", CUSTOMERS)
        with col2:
            region  = st.selectbox("Region *", REGIONS)
            pl_name = st.selectbox("Assign PL / SPL", pl_emails) if pl_emails else st.text_input("PL/SPL Email")
        if st.form_submit_button("🚀 Submit Idea", use_container_width=True):
            if not idea_name or not idea_desc:
                st.error("Idea Name and Description are required.")
            else:
                new_id = str(uuid.uuid4())
                add_idea({"id":new_id,"name":name,"submitter_email":sub_email,
                          "idea_name":idea_name,"idea":idea_desc,"project":project,
                          "category":category,"pl_name":pl_name,"status":"New Idea",
                          "customer":customer,"region":region})
                st.success("✅ Idea Submitted Successfully")
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: PL ASSIGNMENT  (+ engineer load bar chart top-right)
# ══════════════════════════════════════════════════════════════════════════════
def page_pl_assignment():
    page_header("PL Assignment 🧑‍💼")

    all_ideas = get_all()
    users     = get_users()
    engineers = [u["email"] for u in users if u["role"]=="automation engineer"]

    # ── Engineer load bar chart — top right ──────────────────────────────
    left_col, right_col = st.columns([2, 1])

    with right_col:
        st.markdown("<span style='font-size:13px;font-weight:600;'>📊 Engineer Task Load</span>", unsafe_allow_html=True)
        if engineers:
            MAX_TASKS = 10   # max expected tasks — gauge full at 10
            for eng in engineers:
                active = len([
                    i for i in all_ideas
                    if i.get("assigned_engineer")==eng
                    and i.get("status") in {"Assigned","WIP","UAT","Hold/Park"}
                ])
                eng_label = eng.split("@")[0].replace("."," ").title()
                load_pct  = min(active / MAX_TASKS, 1.0) * 100
                # colour: green→amber→red based on load
                needle_color = ("#059669" if load_pct < 40
                                else "#b45309" if load_pct < 75
                                else "#dc2626")
                gauge_opt = {
                    "series":[{
                        "type":"gauge",
                        "radius":"85%",
                        "startAngle":200,"endAngle":-20,
                        "min":0,"max":MAX_TASKS,
                        "splitNumber":5,
                        "axisLine":{
                            "lineStyle":{
                                "width":10,
                                "color":[[0.4,"#059669"],[0.75,"#b45309"],[1,"#dc2626"]]
                            }
                        },
                        "pointer":{"itemStyle":{"color":"auto"},"length":"60%","width":4},
                        "axisTick":{"distance":-15,"length":6,"lineStyle":{"color":"#fff","width":1}},
                        "splitLine":{"distance":-20,"length":12,"lineStyle":{"color":"#fff","width":2}},
                        "axisLabel":{"color":"inherit","distance":18,"fontSize":8},
                        "detail":{
                            "valueAnimation":True,
                            "formatter":f"{active} tasks",
                            "color":"inherit","fontSize":11,"offsetCenter":[0,"60%"]
                        },
                        "title":{"offsetCenter":[0,"85%"],"fontSize":9,"color":"#64748b"},
                        "data":[{"value":active,"name":eng_label}],
                    }]
                }
                st_echarts(gauge_opt, height="160px", key=f"gauge_{eng}")
        else:
            st.info("No automation engineers configured.")

    with left_col:
        st.caption("⭐ Auto-priority: **Customer Requirement → ROI (high first) → FIFO**")
        new_ideas = rank_ideas([i for i in all_ideas if i.get("status")=="New Idea"])

        if not new_ideas:
            st.info("No new ideas pending assignment.")
        else:
            if ss("_assign_outlook_url"):
                url = ss("_assign_outlook_url"); lbl = ss("_assign_outlook_label","")
                st.markdown(f"""
                <div class="idea-card" style="border-left:4px solid #0ea5e9;">
                  📤 <b>Send assignment notification via Outlook:</b><br>
                  <span style="color:#64748b;font-size:12px;">{lbl}</span><br>
                  <a href="{url}" target="_blank" class="outlook-btn">📤 Open in Outlook</a>
                </div>""", unsafe_allow_html=True)
                st.session_state.pop("_assign_outlook_url",None)
                st.session_state.pop("_assign_outlook_label",None)

            for idea in new_ideas:
                with st.expander(f"💡 {idea.get('idea_name','(no name)')}  —  {idea.get('priority_label','')}"):
                    c1,c2 = st.columns(2)
                    with c1:
                        st.markdown(f"**Submitted by:** {idea.get('name','-')}")
                        st.markdown(f"**Project:** {idea.get('project','-')} | **Category:** {idea.get('category','-')}")
                        st.markdown(f"**Customer:** {idea.get('customer','-')} | **Region:** {idea.get('region','-')}")
                        st.markdown(f"**Description:** {idea.get('idea','-')}")
                    with c2:
                        st.markdown(f"**PL/SPL:** {idea.get('pl_name','-')}")
                        st.markdown(f"**Submitted:** {idea.get('created_date','-')[:10]}")

                    if not engineers:
                        st.warning("No automation engineers in the system — add them in Admin first.")
                        continue

                    with st.form(f"assign_{idea['id']}"):
                        eng = st.selectbox("Assign Engineer", engineers, key=f"eng_{idea['id']}")
                        if st.form_submit_button("✅ Assign"):
                            qi = compute_delivery(all_ideas, eng, idea)
                            update_idea(idea["id"],{
                                "assigned_engineer":eng,"status":"Assigned",
                                "assigned_date":datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "priority_label":qi["priority_label"] if qi else "",
                                "sprint_start":fmt_d(qi["sprint_start"]) if qi else "",
                                "sprint_end":fmt_d(qi["sprint_end"]) if qi else "",
                                "delivery_date":fmt_d(qi["sprint_end"]) if qi else "",
                            })
                            fresh = next((i for i in get_all() if i["id"]==idea["id"]),idea)
                            url = build_assign_outlook(fresh, eng, qi)
                            st.session_state["_assign_outlook_url"]   = url
                            st.session_state["_assign_outlook_label"] = f"Notify {eng} — {idea.get('idea_name','')}"
                            st.success(f"Assigned to {eng}. Click 📤 Open in Outlook above to notify them.")
                            st.rerun()

    # ── Engineer Workload & Sprint Schedule ───────────────────────────────
    st.divider()
    st.markdown("#### 📅 Engineer Workload & Sprint Schedule")
    st.caption("Active tasks per engineer ordered by auto-priority (Customer → ROI → FIFO), with rolling 2-week sprint dates.")
    all_eng = [u["email"] for u in users if u["role"]=="automation engineer"]
    if not all_eng:
        st.info("No automation engineers configured — add them in Admin.")
    else:
        sel_engs = st.multiselect("Select Engineer(s) to view", all_eng,
                                  default=None, placeholder="Choose one or more engineers…")
        import pandas as pd
        for eng in (sel_engs or []):
            queue = engineer_queue(all_ideas, eng)
            with st.expander(f"👷 {eng}  —  {len(queue)} active task(s)", expanded=True):
                if not queue:
                    st.caption("No active tasks — fully available.")
                else:
                    df = pd.DataFrame([{
                        "Queue #":i["priority_rank"],"Idea":i.get("idea_name",""),
                        "Category":i.get("category",""),"Priority":i.get("priority_label",""),
                        "Sprint Start":fmt_d(i["sprint_start"]),"Delivery (Sprint End)":fmt_d(i["sprint_end"]),
                    } for i in queue])
                    st.dataframe(df, use_container_width=True, hide_index=True)

    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: FEASIBILITY STUDY
# ══════════════════════════════════════════════════════════════════════════════
def page_feasibility():
    page_header("Feasibility Study 🔍")
    all_ideas = get_all()
    assigned  = rank_ideas([i for i in all_ideas if i.get("status")=="Assigned"])

    if not assigned:
        st.info("No ideas pending feasibility study.")
        render_copyright(); return

    if ss("_feas_outlook_url"):
        url = ss("_feas_outlook_url"); lbl = ss("_feas_outlook_label","")
        st.markdown(f"""
        <div class="idea-card" style="border-left:4px solid #059669;">
          📤 <b>Notify PL/SPL via Outlook:</b><br>
          <span style="color:#64748b;font-size:12px;">{lbl}</span><br>
          <a href="{url}" target="_blank" class="outlook-btn" style="background:#059669;">📤 Open in Outlook</a>
        </div>""", unsafe_allow_html=True)
        st.session_state.pop("_feas_outlook_url",None)
        st.session_state.pop("_feas_outlook_label",None)

    for idea in assigned:
        with st.expander(f"💡 {idea.get('idea_name','(no name)')}  —  {idea.get('priority_label','')}"):
            st.markdown(f"**Engineer:** {idea.get('assigned_engineer','-')}  |  "
                        f"**PL/SPL:** {idea.get('pl_name','-')}  |  **Category:** {idea.get('category','-')}")
            if idea.get("delivery_date"):
                st.caption(f"📅 Provisional delivery: {idea['delivery_date']}")

            with st.form(f"feas_{idea['id']}"):
                st.markdown("##### ROI Calculator")
                col1,col2,col3 = st.columns(3)
                with col1: manual = st.number_input("Manual Effort (hrs)", min_value=0.0, step=0.5, key=f"m_{idea['id']}")
                with col2: fte    = st.number_input("FTE Count", min_value=0.0, step=0.1, key=f"f_{idea['id']}")
                with col3: eng_ef = st.number_input("Automation Effort (hrs)", min_value=0.01, step=0.5, value=1.0, key=f"e_{idea['id']}")
                col4,col5 = st.columns(2)
                with col4: freq     = st.selectbox("Frequency", list(FREQ_MULT.keys()), key=f"fr_{idea['id']}")
                with col5: auto_cat = st.selectbox("Automation Category *", AUTO_CATS, key=f"ac_{idea['id']}")
                comments = st.text_area("Comments / Observations", key=f"co_{idea['id']}")
                roi = round((manual*fte*FREQ_MULT[freq])/eng_ef, 2)
                st.info(f"📈 Computed ROI: **{roi}**")
                if st.form_submit_button("✅ Submit Feasibility & Notify PL via Outlook"):
                    vsm_date = next_workday(date.today()+timedelta(days=1))
                    qi = compute_delivery(all_ideas, idea.get("assigned_engineer",""), {**idea,"roi":roi})
                    update_idea(idea["id"],{
                        "status":"WIP","roi":roi,"automation_category":auto_cat,
                        "feasibility_comments":comments,
                        "feasibility_data":{"manual":manual,"fte":fte,"eng":eng_ef,"freq":freq},
                        "wip_date":datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "vsm_meeting_date":fmt_d(vsm_date),
                        **({"priority_label":qi["priority_label"],
                            "sprint_start":fmt_d(qi["sprint_start"]),
                            "sprint_end":fmt_d(qi["sprint_end"]),
                            "delivery_date":fmt_d(qi["sprint_end"])} if qi else {}),
                    })
                    fresh = next((i for i in get_all() if i["id"]==idea["id"]),idea)
                    url = build_feasibility_outlook(fresh, roi, vsm_date, qi)
                    st.session_state["_feas_outlook_url"]   = url
                    st.session_state["_feas_outlook_label"] = (
                        f"Notify PL/SPL ({idea.get('pl_name','')}) — Feasibility complete for {idea.get('idea_name','')}")
                    st.success(f"✅ Submitted. ROI: {roi} | VSM: {fmt_d(vsm_date)}. Click Outlook above to notify PL/SPL.")
                    st.rerun()
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: APPROVAL
# ══════════════════════════════════════════════════════════════════════════════
def page_approval():
    page_header("PL/SPL Approval ✅")
    all_ideas = get_all()
    my_email  = ss("email","")
    wip_ideas = rank_ideas([i for i in all_ideas
                            if i.get("status")=="WIP" and i.get("pl_name","").lower()==my_email.lower()])
    if not wip_ideas:
        other = rank_ideas([i for i in all_ideas if i.get("status")=="WIP"])
        if other and user_role()=="super user":
            st.info("No ideas assigned to your PL/SPL email — showing all WIP ideas (super user view).")
            wip_ideas = other
        else:
            st.info("No ideas pending your approval.")
            render_copyright(); return

    for idea in wip_ideas:
        fd = idea.get("feasibility_data",{}) or {}
        with st.expander(f"💡 {idea.get('idea_name','(no name)')}  —  ROI: {round(idea.get('roi',0),2)}  |  {idea.get('priority_label','')}"):
            col1,col2 = st.columns(2)
            with col1:
                st.markdown(f"**Engineer:** {idea.get('assigned_engineer','-')}")
                st.markdown(f"**Category:** {idea.get('category','-')} / {idea.get('automation_category','-')}")
                st.markdown(f"**Project:** {idea.get('project','-')}")
            with col2:
                st.markdown(f"**Manual Effort:** {fd.get('manual','-')} hrs | **FTE:** {fd.get('fte','-')} | **Freq:** {fd.get('freq','-')}")
                st.markdown(f"**Automation Effort:** {fd.get('eng','-')} hrs")
                st.markdown(f"**ROI:** {round(idea.get('roi',0),2)}")
            if idea.get("feasibility_comments"):
                st.caption(f"💬 {idea['feasibility_comments']}")
            if idea.get("delivery_date"):
                st.caption(f"📅 Sprint: {idea.get('sprint_start','-')} → {idea.get('sprint_end','-')}")

            with st.form(f"appr_{idea['id']}"):
                decision = st.radio("Decision", ["GO ✅","NO-GO ❌"], horizontal=True, key=f"dec_{idea['id']}")
                comment  = st.text_area("Comments", key=f"acom_{idea['id']}")
                reason   = ""
                if "NO-GO" in decision:
                    reason = st.selectbox("Rejection Reason", REJ_REASONS, key=f"rej_{idea['id']}")
                if st.form_submit_button("Submit Decision"):
                    is_go = "GO" in decision and "NO-GO" not in decision
                    now   = datetime.now().strftime("%Y-%m-%d %H:%M")
                    if is_go:
                        eng = idea.get("assigned_engineer","")
                        qi  = compute_delivery(get_all(), eng, idea)
                        sprint_start = qi["sprint_start"] if qi else None
                        sprint_end   = qi["sprint_end"]   if qi else None
                        update_idea(idea["id"],{
                            "status":"WIP","decision":"GO","approval_comment":comment,"wip_date":now,
                            **({"sprint_start":fmt_d(sprint_start),"sprint_end":fmt_d(sprint_end),
                                "delivery_date":fmt_d(sprint_end),
                                "sprint_meeting_date":fmt_d(sprint_start)} if sprint_start else {}),
                        })
                        st.success("✅ GO — Idea is now In Progress (WIP).")
                    else:
                        update_idea(idea["id"],{"status":"Rejected","decision":"NO-GO",
                                               "rejection_reason":reason,"approval_comment":comment})
                        st.warning(f"❌ NO-GO — Idea rejected. Reason: {reason}")
                    st.rerun()
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: DASHBOARD  — with interlinked filters
# ══════════════════════════════════════════════════════════════════════════════
def page_dashboard():
    page_header("Dashboard ")
    all_ideas_raw = get_all()
    if not all_ideas_raw:
        st.info("No ideas yet.")
        render_copyright(); return

    # ── FILTER BAR ────────────────────────────────────────────────────────
    users    = get_users()
    all_pls  = sorted({i.get("pl_name","") for i in all_ideas_raw if i.get("pl_name","")})
    all_regs = sorted({i.get("region","")   for i in all_ideas_raw if i.get("region","")})
    all_cats = CATEGORIES

    st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
    fc1, fc2, fc3, fc4 = st.columns([1.2, 1.2, 1.2, 0.5])
    with fc1:
        f_cat = st.multiselect("Category", all_cats, key="f_cat",
                               placeholder="All categories", label_visibility="collapsed")
        st.caption("🗂 Category")
    with fc2:
        f_pl  = st.multiselect("PL/SPL", all_pls, key="f_pl",
                               placeholder="All PLs", label_visibility="collapsed")
        st.caption("🧑‍💼 PL / SPL")
    with fc3:
        f_reg = st.multiselect("Region", all_regs, key="f_reg",
                               placeholder="All regions", label_visibility="collapsed")
        st.caption("🌍 Region")
    with fc4:
        st.write("")
        if st.button("🔄 Reset", use_container_width=True):
            for k in ["f_cat","f_pl","f_reg"]: st.session_state.pop(k,None)
            st.rerun()
        st.caption("Reset filters")
    st.markdown('</div>', unsafe_allow_html=True)

    # Apply filters — interlinked (all three narrow the same set)
    ideas = all_ideas_raw
    if f_cat: ideas = [i for i in ideas if i.get("category","") in f_cat]
    if f_pl:  ideas = [i for i in ideas if i.get("pl_name","") in f_pl]
    if f_reg: ideas = [i for i in ideas if i.get("region","") in f_reg]

    active_filters = bool(f_cat or f_pl or f_reg)
    if active_filters:
        st.caption(f"📌 Showing **{len(ideas)}** of **{len(all_ideas_raw)}** ideas after filters.")

    if not ideas:
        st.warning("No ideas match the selected filters.")
        render_copyright(); return

    # ── Local helper lambdas (scoped to filtered set) ─────────────────────
    def cnt(s):        return len([i for i in ideas if i.get("status")==s])
    def cnt_cat(cat):  return len([i for i in ideas if i.get("category")==cat])
    def cnt_cat_ac(cat,ac): return len([i for i in ideas if i.get("category")==cat and i.get("automation_category")==ac])
    def cnt_ac(ac):    return len([i for i in ideas if i.get("automation_category")==ac])
    def hrs_cat(cat):  return round(sum(idea_hours(i) for i in ideas if i.get("category")==cat),1)
    def roi_cat(cat):  return round(sum(float(i.get("roi",0) or 0) for i in ideas if i.get("category")==cat),2)
    def roi_ac(ac):    return round(sum(float(i.get("roi",0) or 0) for i in ideas if i.get("automation_category")==ac),2)

    total     = len(ideas)
    cust_hrs  = hrs_cat("Customer Requirement")
    int_hrs   = hrs_cat("Internal")
    cust_roi  = roi_cat("Customer Requirement")
    int_roi   = roi_cat("Internal")
    cust_cnt  = cnt_cat("Customer Requirement")
    int_cnt   = cnt_cat("Internal")
    completed = cnt("Completed")

    # ── ROW 1: KPI Cards ─────────────────────────────────────────────────
    st.markdown("##### 📦 Key Metrics")
    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card(total,"Total Ideas","#1a4fad",f"{completed} completed | {cnt('Rejected')} rejected","💡")
    with c2: kpi_card(completed,"Completed","#059669",f"{round(completed/total*100,1) if total else 0}% completion rate","🏁")
    with c3: kpi_card(f"{cust_hrs:,.0f} hrs","Customer Hrs Saved / yr","#00498F",f"ROI: {cust_roi} | {cust_cnt} ideas | WIP: {cnt('WIP')}","⏱️")
    with c4: kpi_card(f"{int_hrs:,.0f} hrs","Internal Hrs Saved / yr","#0ea5e9",f"ROI: {int_roi} | {int_cnt} ideas | UAT: {cnt('UAT')}","⏳")

    # ── ROW 2: Auto Category Breakdown ───────────────────────────────────
    st.markdown("##### 🔧 Automation Category Breakdown")
    auto_cols = st.columns(len(AUTO_CATS))
    for idx, ac in enumerate(AUTO_CATS):
        ac_total = cnt_ac(ac)
        ac_cust  = cnt_cat_ac("Customer Requirement",ac)
        ac_int   = cnt_cat_ac("Internal",ac)
        ac_roi   = roi_ac(ac)
        color    = AUTO_CAT_COLORS.get(ac,"#888")
        with auto_cols[idx]:
            kpi_card(ac_total, ac, color,
                     f"🔴 Cust: {ac_cust}  🔵 Int: {ac_int}  |  ROI: {ac_roi}", "🔧")

    # ── ROW 3: Charts row (Status + Project + Customer+Region combo pies) ─
    st.markdown("##### 📈 Charts")
    ch1, ch2, ch3 = st.columns(3)

    with ch1:
        st.markdown("<span style='font-size:clamp(10px,1vw,13px);font-weight:600;'>Ideas by Status</span>", unsafe_allow_html=True)
        status_data = [{"value":cnt(s),"name":s,"itemStyle":{"color":STATUS_COLORS.get(s,"#888")}}
                       for s in STATUSES if cnt(s)>0]
        st_echarts({"tooltip":{"trigger":"item","formatter":"{b}: {c} ({d}%)"},
                    "series":[{"type":"pie","radius":["38%","68%"],
                               "data":status_data,"label":{"fontSize":"clamp(8px,0.8vw,10px)"}}]},
                   height="220px")

    with ch2:
        st.markdown("<span style='font-size:clamp(10px,1vw,13px);font-weight:600;'>Customer — Count &amp; ROI</span>", unsafe_allow_html=True)
        cust_data = {}
        for i in ideas:
            c = i.get("customer","") or "Unknown"
            if c not in cust_data: cust_data[c] = {"count":0,"roi":0.0}
            cust_data[c]["count"] += 1
            cust_data[c]["roi"]   += float(i.get("roi",0) or 0)
        cust_palette = ["#E30613","#00AEEF","#7c3aed","#059669","#0d9488","#b45309","#0369a1"]
        c_pie_cnt = [{"value":v["count"],
                      "name":f'{k}\n({round(v["roi"],1)} ROI)',
                      "itemStyle":{"color":cust_palette[idx%len(cust_palette)]}}
                     for idx,(k,v) in enumerate(cust_data.items())]
        st_echarts({
            "tooltip":{"trigger":"item","formatter":"{b}: {c} ideas ({d}%)"},
            "series":[{"type":"pie","radius":["35%","65%"],"data":c_pie_cnt,
                       "label":{"fontSize":9,"formatter":"{b}"},
                       "labelLine":{"length":8,"length2":5}}]
        }, height="220px")

    with ch3:
        st.markdown("<span style='font-size:clamp(10px,1vw,13px);font-weight:600;'>Hours Saved by Project</span>", unsafe_allow_html=True)
        proj_hrs = {}
        for i in ideas:
            proj_hrs[i.get("project","Other")] = proj_hrs.get(i.get("project","Other"),0)+idea_hours(i)
        if proj_hrs:
            st_echarts({
                "tooltip":{"trigger":"axis"},
                "grid":{"left":"3%","right":"4%","bottom":"28%","containLabel":True},
                "xAxis":{"type":"category","data":list(proj_hrs.keys()),
                         "axisLabel":{"rotate":30,"fontSize":8,"interval":0}},
                "yAxis":{"type":"value","name":"hrs/yr","nameTextStyle":{"fontSize":8}},
                "series":[{"type":"bar","data":[round(v,1) for v in proj_hrs.values()],
                           "itemStyle":{"color":"#7c3aed"},"barMaxWidth":32}]},
                height="220px")

    # ── ROW 4: Ideation Tree + Region chart ──────────────────────────────
    tr_col, wl_col = st.columns([1.4, 1])

    with tr_col:
        st.markdown("##### 🌳 Ideation Workflow Tree")
        def cs(cat,st_): return len([i for i in ideas if i.get("category")==cat and i.get("status")==st_])
        def rs(r): return len([i for i in ideas if i.get("status")=="Rejected" and i.get("rejection_reason")==r])
        def add_label_boxes(node):
            color = node.get("itemStyle",{}).get("color","#FFFFFF")
            node["label"] = {"show":True,"backgroundColor":color,"color":"#FFFFFF",
                             "borderRadius":5,"padding":[4,8],"position":"inside",
                             "align":"center","fontSize":10,"fontWeight":"bold"}
            for child in node.get("children",[]): add_label_boxes(child)
        # Tree flow: Ideation → Triage/Feasibility(Queued) → Accepted → Customer → WIP / Deployed
        #                                                              → Internal
        #                                              → Rejected
        tree_data = {
            "name":f"Ideation ({total})","itemStyle":{"color":"#1a4fad"},
            "children":[
                {"name":f"Triage / Feasibility Study\n(Queued: {cnt('Assigned')})","itemStyle":{"color":"#1a4fad"},
                 "children":[
                     {"name":f"Accepted ({cnt('WIP')+cnt('UAT')+cnt('Completed')})","itemStyle":{"color":"#059669"},
                      "children":[
                          {"name":f"Customer ({cust_cnt})","itemStyle":{"color":"#00498F"},
                           "children":[
                               {"name":f"WIP ({cs('Customer Requirement','WIP')+cs('Customer Requirement','UAT')})","itemStyle":{"color":"#0d9488"}},
                               {"name":f"Deployed ({cs('Customer Requirement','Completed')})","itemStyle":{"color":"#059669"}},
                           ]},
                          {"name":f"Internal ({int_cnt})","itemStyle":{"color":"#0ea5e9"},
                           "children":[
                               {"name":f"WIP ({cs('Internal','WIP')+cs('Internal','UAT')})","itemStyle":{"color":"#0d9488"}},
                               {"name":f"Deployed ({cs('Internal','Completed')})","itemStyle":{"color":"#059669"}},
                           ]},
                      ]},
                     {"name":f"Rejected ({cnt('Rejected')})","itemStyle":{"color":"#dc2626"},
                      "children":[
                          {"name":f"Technical ({rs('Technical Rejection')})","itemStyle":{"color":"#ef4444"}},
                          {"name":f"Business ({rs('Business Rejection')})","itemStyle":{"color":"#f97316"}},
                      ]},
                 ]},
            ]
        }
        add_label_boxes(tree_data)
        st_echarts({
            "backgroundColor":"#0B0B0D",
            "tooltip":{"trigger":"item","triggerOn":"mousemove"},
            "series":[{"type":"tree","data":[tree_data],
                       "top":"5%","left":"7%","bottom":"5%","right":"15%",
                       "symbol":"rect","symbolSize":1,
                       "lineStyle":{"color":"#f97316","width":2},
                       "label":{"position":"left","verticalAlign":"middle","align":"right","fontSize":10},
                       "leaves":{"label":{"position":"right","verticalAlign":"middle","align":"left","fontSize":9}},
                       "emphasis":{"focus":"descendant"},
                       "expandAndCollapse":True,"animationDuration":550,"initialTreeDepth":2}]
        }, height="400px")

    with wl_col:
        st.markdown("##### 🌍 Region — Word Map")
        region_data = {}
        for i in ideas:
            r = i.get("region","") or ""
            if not r:
                continue   # skip ideas with no region set
            if r not in region_data: region_data[r] = {"count":0,"roi":0.0}
            region_data[r]["count"] += 1
            region_data[r]["roi"]   += float(i.get("roi",0) or 0)

        no_region_count = len([i for i in ideas if not (i.get("region","") or "").strip()])

        if not region_data:
            st.info(f"No region data yet — {no_region_count} idea(s) have no region assigned.")
        else:
            # ── Word-map: sized coloured bubbles (true word-cloud feel) ──
            max_count = max(v["count"] for v in region_data.values()) or 1
            sorted_regions = sorted(region_data.items(), key=lambda x: -x[1]["count"])
            colors = [
                "#E30613","#1a4fad","#059669","#7c3aed",
                "#0891b2","#b45309","#9333ea","#0d9488",
                "#0369a1","#16a34a","#dc2626","#d97706",
            ]
            wc_html = (
                '<div style="display:flex;flex-wrap:wrap;gap:10px;align-items:center;' +
                'justify-content:center;padding:18px 8px;'  +
                'background:#f8fafc;border-radius:12px;border:1px solid #e2e8f0;'
                'min-height:180px;">'
            )
            for idx,(region,val) in enumerate(sorted_regions):
                ratio   = val["count"] / max_count          # 0.0 – 1.0
                fs      = int(13 + ratio * 26)              # 13 – 39 px
                pad_h   = int(8  + ratio * 14)
                pad_v   = int(4  + ratio * 8)
                alpha   = int(18 + ratio * 20)              # hex opacity 18–38
                col_hex = colors[idx % len(colors)]
                bg_hex  = col_hex + hex(alpha)[2:].zfill(2) # colour + alpha
                tooltip = f"{region}: {val['count']} idea(s) · ROI {round(val['roi'],1)}"
                wc_html += (
                    f'<span title="{tooltip}" style="' +
                    f'font-size:{fs}px;font-weight:{600 if ratio>0.5 else 500};' +
                    f'color:{col_hex};background:{bg_hex};' +
                    f'border-radius:8px;padding:{pad_v}px {pad_h}px;' +
                    f'cursor:default;line-height:1.5;white-space:nowrap;'
                    f'box-shadow:0 1px 3px rgba(0,0,0,.06);">' +
                    f'{region}'
                    f'<span style="font-size:{max(9,fs-10)}px;vertical-align:super;'
                    f'margin-left:3px;opacity:0.7;">{val['count']}</span>'
                    f'</span>'
                )
            wc_html += '</div>'
            st.markdown(wc_html, unsafe_allow_html=True)
            if no_region_count:
                st.caption(f"ℹ️ {no_region_count} idea(s) have no region assigned and are excluded.")
            st.caption("Font size = idea count · Hover for count & ROI")

    # ── All Ideas table + CSV (above Kanban) ────────────────────────────
    st.markdown("##### 📄 All Ideas")
    search = st.text_input("🔎 Search ideas", placeholder="Filter by name, project, status…")
    import pandas as pd
    cols_show = ["idea_name","name","project","category","automation_category","status",
                 "priority_label","assigned_engineer","roi","sprint_start","delivery_date",
                 "customer","region","created_date"]
    df = pd.DataFrame([{c:i.get(c,"") for c in cols_show} for i in ideas])
    if search:
        mask = df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)
        df   = df[mask]
    st.dataframe(df, use_container_width=True, hide_index=True)
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False)
    st.download_button("⬇️ Download CSV", csv_buf.getvalue(), "turbodrive_ideas.csv", "text/csv")

    # ── Kanban Board ──────────────────────────────────────────────────────
    st.markdown("##### 📋 Kanban Board")
    render_kanban_board(ideas)

    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: EMAIL
# ══════════════════════════════════════════════════════════════════════════════
def page_email():
    page_header("Send Meeting Invites 📤")
    st.caption("All dates are auto-computed from sprint scheduling. Click 'Open in Outlook' to review & send manually.")
    all_ideas = get_all()
    idea_opts = {f"{i.get('idea_name','(no name)')} — {i.get('status','')}":i for i in all_ideas}
    if not idea_opts:
        st.info("No ideas yet."); render_copyright(); return

    col1,col2 = st.columns(2)
    with col1: sel_name = st.selectbox("Select Idea", list(idea_opts.keys()))
    with col2:
        mtype = st.selectbox("Meeting Type",
                             ["vsm — VSM Session (11:00 AM)",
                              "sprint — Sprint Planning (10:00 AM)",
                              "delivery — Delivery/Demo Review (03:00 PM)"])

    idea  = idea_opts[sel_name]
    mkey  = mtype.split(" — ")[0]
    date_map = {"vsm":idea.get("vsm_meeting_date"),"sprint":idea.get("sprint_meeting_date"),"delivery":idea.get("delivery_date")}
    mdate = parse_d(date_map.get(mkey,""))

    st.markdown(f"""
    | Field | Value |
    |---|---|
    | Idea | {idea.get('idea_name','-')} |
    | Project | {idea.get('project','-')} |
    | Category | {idea.get('category','-')} / {idea.get('automation_category','-')} |
    | Customer | {idea.get('customer','-')} — Region | {idea.get('region','-')} |
    | Engineer | {idea.get('assigned_engineer','-')} |
    | PL/SPL | {idea.get('pl_name','-')} |
    | Meeting Date | **{fmt_d(mdate) if mdate else '⚠ Not set yet'}** |
    """)

    recips_map = {
        "vsm":     [idea.get("submitter_email",""),idea.get("assigned_engineer",""),idea.get("pl_name","")],
        "sprint":  [idea.get("assigned_engineer",""),idea.get("pl_name","")],
        "delivery":[idea.get("submitter_email",""),idea.get("assigned_engineer",""),idea.get("pl_name","")],
    }
    if mdate:
        url = build_meeting_outlook(mkey, idea, mdate, recips_map[mkey])
        st.markdown(f'<a href="{url}" target="_blank" class="outlook-btn">📤 Open in Outlook</a>', unsafe_allow_html=True)
    else:
        st.warning("⚠️ No date set for this meeting type — complete the workflow first (Feasibility / Approval).")

    st.divider()
    st.markdown("**Auto-Email Event Reference**")
    st.markdown("""
    | Event | Trigger | Recipients |
    |---|---|---|
    | ✅ Feasibility Complete | Engineer submits feasibility | PL/SPL (via Outlook button) |
    | 🎉 Idea Approved (GO) | PL/SPL approves | Submitter + Engineer |
    | ❌ Idea Rejected (NO-GO) | PL/SPL rejects | Submitter |
    | 🧩 VSM Session | After feasibility (manual) | Submitter + Engineer + PL/SPL |
    | 🚀 Sprint Planning | After approval GO (manual) | Engineer + PL/SPL |
    | 🏁 Delivery/Demo | Sprint end date (manual) | Submitter + Engineer + PL/SPL |
    """)
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: ADMIN
# ══════════════════════════════════════════════════════════════════════════════
def page_admin():
    page_header("Admin Panel ⚙️")
    users = get_users()
    tab1,tab2,tab3 = st.tabs(["👥 Users","➕ Add User","🔑 Password Reset"])

    with tab1:
        st.markdown(f"**{len(users)} registered users**")
        for u in users:
            with st.expander(f"📧 {u['email']}  —  {u['role']}"):
                col1,col2,col3 = st.columns([2,2,1])
                with col1:
                    new_role = st.selectbox("Role", ROLES_LIST,
                                            index=ROLES_LIST.index(u["role"]) if u["role"] in ROLES_LIST else 0,
                                            key=f"role_{u['email']}")
                with col2:
                    if st.button("💾 Update Role", key=f"upd_{u['email']}"):
                        update_role(u["email"], new_role); st.success("Role updated.")
                with col3:
                    if st.button("🗑 Delete", key=f"del_{u['email']}"):
                        delete_user(u["email"]); st.warning("User deleted.")

    with tab2:
        with st.form("add_user_form", clear_on_submit=True):
            new_email = st.text_input("Email")
            new_role  = st.selectbox("Role", ROLES_LIST)
            if st.form_submit_button("➕ Add User"):
                if new_email:
                    add_user(new_email.strip().lower(), new_role)
                    st.success(f"Added {new_email} as {new_role} (default pw: {DEFAULT_PW})")
                    st.rerun()

    with tab3:
        st.caption(f"Reset any user's password back to default: **{DEFAULT_PW}**")
        emails = [u["email"] for u in users]
        with st.form("reset_pw_form"):
            target = st.selectbox("User", emails)
            col1,col2 = st.columns(2)
            with col1: new_pw = st.text_input("Set New Password (optional)", type="password")
            with col2:
                do_reset = st.form_submit_button(f"Reset to '{DEFAULT_PW}'")
                do_set   = st.form_submit_button("Set Specific Password")
            if do_reset:
                reset_password(target); st.success(f"Password reset to '{DEFAULT_PW}' for {target}")
            elif do_set:
                if not new_pw or len(new_pw)<4: st.error("Minimum 4 characters.")
                else: set_password(target, new_pw); st.success(f"Password updated for {target}")
    render_copyright()

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN ROUTER
# ══════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Turbo Drive", page_icon="", layout="wide",
                       initial_sidebar_state="expanded")
    init_db()
    if "theme" not in st.session_state:
        st.session_state["theme"] = "ALTEN Red & Blue"
    apply_theme(ss("theme"))

    override = ss("_page_override")
    if override == "register":    page_register(); return
    if override == "change_password": page_change_password(); return

    if not logged_in():
        if ss("_session_expired"):
            st.warning("⚠️ Session expired due to 5 minutes of inactivity. Please login again.")
            st.session_state.pop("_session_expired", None)
        page_login(); return

    # ── Auto-logout check (runs on every page load / interaction) ─────────
    check_session_timeout()
    if not logged_in():   # may have just been logged out by the check above
        st.rerun()
    touch_activity()       # any successful page render counts as activity

    t = THEMES.get(ss("theme"), THEMES["ALTEN Red & Blue"])
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align:center;padding:10px 0 6px;">
          <img src="{ALTEN_LOGO_URL}" style="height:22px;object-fit:contain;margin-bottom:4px;" alt="ALTEN"/><br>
          <span style="font-size:clamp(14px,1.5vw,20px);font-weight:900;
               background:linear-gradient(135deg,{t['primary']},{t['secondary']});
               -webkit-background-clip:text;background-clip:text;color:transparent;">
             TURBO DRIVE
          </span>
          <div style="color:#94a3b8;font-size:10px;margin-top:2px;">Automation Workflow</div>
        </div>""", unsafe_allow_html=True)
        st.divider()

        pages = user_pages()
        icons = {"Dashboard":"📊","Submit Idea":"💡","PL Assignment":"🧑‍💼",
                 "Feasibility":"🔍","Approval":"✅","Admin":"⚙️","Email":"📤"}
        nav = st.radio("Navigation",
                       [f"{icons.get(p,'')} {p}" for p in pages],
                       label_visibility="collapsed")
        current_page = nav.split(" ",1)[1].strip() if nav else pages[0]

        st.divider()
        st.markdown(f"""
        <div style="color:#94a3b8;font-size:12px;">
          👤 <b style="color:#e2e8f0;">{ss('name','')}</b><br>
          <span style="font-size:11px;">{ss('role','')}</span>
        </div>""", unsafe_allow_html=True)

        col1,col2 = st.columns(2)
        with col1:
            if st.button("🔑 PW", help="Change Password"):
                st.session_state["_page_override"] = "change_password"; st.rerun()
        with col2:
            if st.button("🚪 Logout"):
                for k in ["email","role","name","theme","_last_activity"]: st.session_state.pop(k,None)
                st.rerun()

        render_session_countdown()

        st.divider()
        st.markdown("**🎨 Theme**")
        chosen = st.selectbox("", list(THEMES.keys()),
                              index=list(THEMES.keys()).index(ss("theme","ALTEN Red & Blue")),
                              label_visibility="collapsed", key="theme_sel")
        if chosen != ss("theme"):
            st.session_state["theme"] = chosen; st.rerun()

        st.markdown("---")
        st.markdown(
            f'<p style="font-size:10px;color:#64748b;">Queries?<br>'
            f'<a href="mailto:{SUPPORT_EMAIL}" style="color:#00AEEF;">{SUPPORT_NAME}</a></p>',
            unsafe_allow_html=True
        )

    if   current_page == "Dashboard":     page_dashboard()
    elif current_page == "Submit Idea":   page_submit()
    elif current_page == "PL Assignment": page_pl_assignment()
    elif current_page == "Feasibility":   page_feasibility()
    elif current_page == "Approval":      page_approval()
    elif current_page == "Email":         page_email()
    elif current_page == "Admin":         page_admin()

if __name__ == "__main__":
    main()
