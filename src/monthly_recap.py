import streamlit as st
import pandas as pd
from datetime import datetime
from attribution_parser import parse_dataframe_attributions
from jinja2 import Template
import os
import re
import sys
import subprocess


def ensure_package(pkg: str) -> bool:
    # Try to import the package. Importing some packages (notably
    # weasyprint) can raise system-level errors (OSError) when their
    # native dependencies are missing. Catch all exceptions here so the
    # caller can fall back to alternatives instead of crashing.
    try:
        __import__(pkg)
        return True
    except Exception:
        # Attempt to install via pip and import again. If installation
        # or import still fails (including OSError), return False so
        # the caller can choose a fallback renderer.
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
            try:
                __import__(pkg)
                return True
            except Exception:
                return False
        except Exception:
            return False


PDF_RENDERER = None

if ensure_package('markdown2'):
    import markdown2  # type: ignore
else:
    markdown2 = None

# PDF generation is intentionally disabled in this build to avoid
# requiring native libraries (cairo/pango/glib) on the host system.
# The app will export HTML email content only.
WEASYPRINT_AVAILABLE = False


st.title("Monthly-Recap Builder")

learning_activity_file = st.file_uploader('Upload the learning activity report (Excel)')
tom_report = st.file_uploader('Upload the TOM report (Excel)')
checkin_report = st.file_uploader('Upload the check-in report (optional, Excel)')

hide_elements = """
<style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
"""
st.markdown(hide_elements, unsafe_allow_html=True)

def title_case_list(names):
    return [n.title() for n in names if isinstance(n, str) and n]


def extract_dealership_name(df_clean: pd.DataFrame, filename: str, checkin_file=None) -> str:
    """Try to extract a reasonable dealership name from the dataframe.

    Prefer non-email text found in the first few rows/cols. If nothing
    suitable is found, sanitize the filename (remove email-like parts
    and prefixes) and return that.
    """
    email_re = re.compile(r"\S+@\S+")
    dealer_keywords = ['hendrick', 'dealership', 'chevrolet', 'honda', 'bmw', 'lexus', 'porsche', 'cadillac', 'acura', 'mini', 'motors', 'group', 'store']

    def looks_like_person(name: str) -> bool:
        # Rudimentary person-name detector: 1-3 tokens, each starting with
        # uppercase followed by lowercase letters, no digits, and not
        # containing dealer keywords.
        toks = name.split()
        if not 1 <= len(toks) <= 3:
            return False
        for t in toks:
            if re.search(r"\d", t):
                return False
            if not re.match(r"^[A-Z][a-z'-]+$", t):
                return False
        low = name.lower()
        if any(k in low for k in dealer_keywords):
            return False
        return True

    # 0) If a checkin report was provided, try B1 on its first sheet first
    if checkin_file is not None:
        try:
            chk = pd.read_excel(checkin_file, sheet_name=0, header=None)
            if chk.shape[0] > 0 and chk.shape[1] > 1:
                val = chk.iat[0, 1]
                if pd.notna(val):
                    s = str(val).strip()
                    if s and not email_re.search(s) and not looks_like_person(s):
                        return s
        except Exception:
            pass

    # 1) Preferred old heuristic: cell at row index 2, col index 1 (if present)
    try:
        if df_clean.shape[0] > 2 and df_clean.shape[1] > 1:
            val = df_clean.iat[2, 1]
            if pd.notna(val):
                s = str(val).strip()
                if s and not email_re.search(s) and not looks_like_person(s):
                    return s
    except Exception:
        pass

    # 2) Search top region for dealer-like strings; skip email and person-like values
    candidate = None
    max_rows = min(8, df_clean.shape[0])
    max_cols = min(6, df_clean.shape[1]) if df_clean.shape[1] > 0 else 0
    for i in range(max_rows):
        for j in range(max_cols):
            try:
                val = df_clean.iat[i, j]
            except Exception:
                continue
            if pd.isna(val):
                continue
            s = str(val).strip()
            if not s or email_re.search(s):
                continue
            low = s.lower()
            # prefer explicit dealer keywords
            if any(k in low for k in dealer_keywords):
                return s
            # prefer multi-word labels that are unlikely to be person names
            if ' ' in s and len(s) > 4 and not looks_like_person(s):
                return s
            if candidate is None and not looks_like_person(s):
                candidate = s

    if candidate:
        return candidate

    # Sanitize filename: remove ai_insight prefix and email-like substrings
    base = os.path.splitext(filename)[0]
    base = re.sub(r'^ai_insight[_-]*', '', base, flags=re.I)
    base = re.sub(r"\S+@\S+", '', base)
    base = re.sub(r'[_-]+', ' ', base).strip()
    return base.title() if base else filename


def extract_tom_completions(parsed_df: pd.DataFrame, raw_df: pd.DataFrame) -> int:
    total = 0
    if 'Completions' in parsed_df.columns:
        total = pd.to_numeric(parsed_df['Completions'], errors='coerce').fillna(0).sum()
    if not total:
        completed_col = next((c for c in parsed_df.columns if 'completed' in c.lower()), None)
        if completed_col:
            total = pd.to_numeric(parsed_df[completed_col], errors='coerce').fillna(0).sum()
    if not total:
        inspected = set()
        for col in raw_df.columns:
            key = str(col)
            if key in inspected:
                continue
            series = raw_df[col]
            as_str = series.astype(str).str.strip().str.lower()
            if 'completed' in as_str.values or key.strip().lower() == 'completed':
                total += pd.to_numeric(series, errors='coerce').fillna(0).sum()
                inspected.add(key)
    return int(total)


def compute_metrics(learning_activity_file, tom_report, checkin_file=None):
    """Return computed metrics needed for the email and insight generation.

    Returns dict with keys: dealership_name, month, prev_month, total_teammates,
    standouts (list of dicts), least_active (list), tom_completions.
    """
    dealership_df = pd.read_excel(learning_activity_file, sheet_name='Report', skiprows=5)
    dealership_df.columns = [str(c).strip() for c in dealership_df.columns]
    df_clean = dealership_df.dropna(how='all')

    dealership_name = extract_dealership_name(df_clean, learning_activity_file.name, checkin_file=checkin_file)

    month = datetime.now().strftime('%B')
    prev_month = (datetime.now().replace(day=1) - pd.Timedelta(days=1)).strftime('%B')

    colmap = {c.lower(): c for c in df_clean.columns}
    def find_col(key):
        for k, v in colmap.items():
            if key in k:
                return v
        return None

    emp_col = find_col('employee') or find_col('name')
    active_col = find_col('active days') or find_col('active')
    streak_col = find_col('current streak') or find_col('streak')
    journeys_col = find_col('journeys') or find_col('journeys completed')

    if emp_col is None or active_col is None:
        raise ValueError('Could not find required columns (employee name / active days) in the learning activity file')

    df_clean[emp_col] = df_clean[emp_col].astype(str).str.title()
    total_teammates = df_clean.shape[0]

    ranked = df_clean.sort_values([active_col, streak_col] if streak_col in df_clean.columns else [active_col], ascending=False)
    standouts_df = ranked.head(3).copy()
    standouts = []
    for _, r in standouts_df.iterrows():
        standouts.append({
            'Employee_Name': r.get(emp_col, ''),
            'Active_Days': int(r.get(active_col, 0)) if pd.notna(r.get(active_col, 0)) else 0,
            'Current_Streak': int(r.get(streak_col, 0)) if streak_col and pd.notna(r.get(streak_col, 0)) else 0,
            'Journeys_Completed': int(r.get(journeys_col, 0)) if journeys_col and pd.notna(r.get(journeys_col, 0)) else 0,
        })

    least_active_series = df_clean[df_clean[active_col] <= 4][emp_col].tolist() if active_col in df_clean.columns else []
    least_active = title_case_list(least_active_series)

    tom_df = pd.read_excel(tom_report, skiprows=4)
    tom_df.columns = [str(c).strip() for c in tom_df.columns]
    tom_parsed_df, parsed_col = parse_dataframe_attributions(tom_df, column_name_contains='Attribution')

    tom_completions = 0
    if 'Completions' in tom_parsed_df.columns:
        tom_completions = pd.to_numeric(tom_parsed_df['Completions'], errors='coerce').fillna(0).sum()
    if not tom_completions:
        completed_col = next((c for c in tom_parsed_df.columns if 'completed' in c.lower()), None)
        if completed_col:
            tom_completions = pd.to_numeric(tom_parsed_df[completed_col], errors='coerce').fillna(0).sum()
    if not tom_completions:
        inspected = set()
        for col in tom_df.columns:
            key = str(col)
            if key in inspected:
                continue
            series = tom_df[col]
            col_str = series.astype(str).str.strip().str.lower()
            if 'completed' in col_str.values or key.strip().lower() == 'completed':
                tom_completions += pd.to_numeric(series, errors='coerce').fillna(0).sum()
                inspected.add(key)
    tom_completions = int(tom_completions)

    return {
        'dealership_name': dealership_name,
        'month': month,
        'prev_month': prev_month,
        'total_teammates': total_teammates,
        'standouts': standouts,
        'least_active': least_active,
        'tom_completions': tom_completions,
        'df_clean': df_clean,
    }


def generate_insight(api_key, tom_completions, total_teammates, standouts, least_active):
    """Call OpenAI new client and return the assistant text."""
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        system_msg = {
            'role': 'system',
            'content': (
                "You are a RockED performance strategist crafting upbeat, professional recaps for dealership leaders. "
                "Deliver optimistic yet data-driven bullet points that tie learning engagement to sales outcomes and end with clear next steps."
            )
        }
        standout_notes = '; '.join([
            f"{s['Employee_Name']} ({s['Active_Days']} active days, streak {s['Current_Streak']}, {s['Journeys_Completed']} journeys)"
            for s in standouts
        ]) or 'none provided'
        low_engagers = ', '.join(least_active[:8]) if least_active else 'none'

        user_content = (
            f"Metrics: ToM completions {tom_completions} of {total_teammates} teammates. "
            f"Standout performers: {standout_notes}. Least active learners: {low_engagers}. "
            "Create a fun, encouraging, but executive-ready summary with three mini paragraphs: "
            "1) Area for Improvement - Topic of the Month (ToM) highlighting completion gap, business impact, and upbeat next steps. "
            "2) Standout Performers celebrating the top names with actionable takeaways. "
            "3) Least Active Learners recommending a coaching plan. Use a pleasant tone and include one or two friendly emojis." 
        )
        messages = [system_msg, {'role': 'user', 'content': user_content}]

        resp = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=messages,
            max_tokens=220,
            temperature=0.5,
        )

        try:
            choice0 = resp.choices[0]
            msg = None
            if isinstance(choice0, dict):
                msg = choice0.get('message') or choice0.get('text')
            else:
                msg = getattr(choice0, 'message', None) or getattr(choice0, 'text', None)

            if isinstance(msg, dict):
                ai_text = msg.get('content') or msg.get('text')
            else:
                ai_text = None
                if msg is not None:
                    ai_text = getattr(msg, 'content', None) or getattr(msg, 'text', None) or getattr(msg, 'output_text', None)

            if not ai_text:
                ai_text = getattr(resp, 'output_text', None)

            if isinstance(ai_text, str):
                return ai_text.strip()
            if ai_text is not None:
                return str(ai_text).strip()

        except Exception:
            pass

        s = str(resp)
        m = re.search(r"ChatCompletionMessage\(content='([^']+)'", s)
        if m:
            return m.group(1).strip()
        return s
    except Exception as e:
        return f"AI generation error: {e}"


def _sanitize_names(names, limit=8):
    out = []
    for n in names[:limit]:
        if not isinstance(n, str):
            continue
        # remove emails and excessive punctuation
        s = re.sub(r"\S+@\S+", '', n).strip()
        s = re.sub(r'[_\t]+', ' ', s).strip()
        if s:
            out.append(s)
    return out


def build_insight_messages(metrics: dict):
    """Return (system_msg, user_content, messages) for the insight generator."""
    dealer = metrics.get('dealership_name', '')
    month = metrics.get('month', '')
    prev = metrics.get('prev_month', '')
    total = metrics.get('total_teammates', 0)
    tom = int(metrics.get('tom_completions', 0) or 0)
    pct = round(100 * tom / total, 1) if total else 0

    standouts = metrics.get('standouts') or []
    standout_lines = []
    for s in standouts[:6]:
        name = s.get('Employee_Name') or ''
        ad = s.get('Active_Days', 0)
        strek = s.get('Current_Streak', 0)
        jour = s.get('Journeys_Completed', 0)
        standout_lines.append(f"{name} ({ad} active days, streak {strek}, {jour} journeys)")

    least = _sanitize_names(metrics.get('least_active') or [], limit=8)

    # aggregated stat
    df = metrics.get('df_clean')
    avg_active = None
    try:
        if df is not None:
            act_col = next((c for c in df.columns if 'active' in str(c).lower()), None)
            if act_col is not None:
                avg_active = round(pd.to_numeric(df[act_col], errors='coerce').dropna().mean(), 1)
    except Exception:
        avg_active = None

    system_msg = {
        'role': 'system',
        'content': (
            "You are a concise, upbeat performance strategist writing 3 short sections for dealership leaders: "
            "1) Area for Improvement — explain the ToM completion gap and business impact with clear next steps; "
            "2) Standout Performers — celebrate 1-3 names with actionable suggestions; "
            "3) Least Active Learners — recommend a brief coaching plan. "
            "Keep the tone professional but friendly, include 1 emoji, and keep to ~160-220 words."
        )
    }

    user_parts = [
        f"Dealership: {dealer}",
        f"Period: {prev} -> {month}",
        f"ToM completions: {tom} of {total} teammates ({pct}%)",
    ]
    if standout_lines:
        user_parts.append("Standouts: " + '; '.join(standout_lines))
    if least:
        user_parts.append("Least active: " + ', '.join(least))
    if avg_active is not None:
        user_parts.append(f"Average active days: {avg_active}")

    user_content = '\n'.join(user_parts)

    example = (
        "Example:\nArea for Improvement - ToM:\n• Completion gap: 12 of 40 (30%) — schedule 15-min huddles twice weekly.\n"
        "Standout Performers:\n• Alice (24 days) — spotlight in team huddle.\n"
        "Least Active Learners:\n• Bob, Carol — assign quick 3-day goal and manager check-in."
    )

    messages = [system_msg, {'role': 'user', 'content': user_content + '\n\n' + example}]
    return system_msg, user_content, messages


def generate_insight_from_metrics(api_key, metrics: dict, model: str = 'gpt-3.5-turbo', max_tokens: int = 300, temperature: float = 0.45):
    """Generate an AI insight from a metrics dict.

    metrics expected keys: dealership_name, month, prev_month, total_teammates,
    tom_completions, standouts (list of dicts), least_active (list), df_clean (opt).
    """
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
    except Exception as e:
        return f"AI client error: {e}"

    # Build structured summary to send to the model
    dealer = metrics.get('dealership_name', '')
    month = metrics.get('month', '')
    prev = metrics.get('prev_month', '')
    total = metrics.get('total_teammates', 0)
    tom = int(metrics.get('tom_completions', 0) or 0)
    pct = round(100 * tom / total, 1) if total else 0

    # Standouts -> short summaries
    standouts = metrics.get('standouts') or []
    standout_lines = []
    for s in standouts[:6]:
        name = s.get('Employee_Name') or ''
        ad = s.get('Active_Days', 0)
        strek = s.get('Current_Streak', 0)
        jour = s.get('Journeys_Completed', 0)
        standout_lines.append(f"{name} ({ad} active days, streak {strek}, {jour} journeys)")

    least = _sanitize_names(metrics.get('least_active') or [], limit=8)

    # Additional aggregated stats if df_clean present
    df = metrics.get('df_clean')
    avg_active = None
    try:
        if df is not None and 'Active' in df.columns or True:
            # try to find any numeric active-days-like column
            act_col = next((c for c in df.columns if 'active' in str(c).lower()), None)
            if act_col is not None:
                avg_active = round(pd.to_numeric(df[act_col], errors='coerce').dropna().mean(), 1)
    except Exception:
        avg_active = None

    system_msg, user_content, messages = build_insight_messages(metrics)

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature,
        )
        # try to extract text robustly
        choice0 = resp.choices[0]
        msg = None
        if isinstance(choice0, dict):
            msg = choice0.get('message') or choice0.get('text')
        else:
            msg = getattr(choice0, 'message', None) or getattr(choice0, 'text', None)

        if isinstance(msg, dict):
            ai_text = msg.get('content') or msg.get('text')
        else:
            ai_text = None
            if msg is not None:
                ai_text = getattr(msg, 'content', None) or getattr(msg, 'text', None) or getattr(msg, 'output_text', None)

        if not ai_text:
            ai_text = getattr(resp, 'output_text', None)
        if isinstance(ai_text, str):
            return ai_text.strip()
        if ai_text is not None:
            return str(ai_text).strip()

        # last resort
        return str(resp)
    except Exception as e:
        return f"AI generation error: {e}"


EMAIL_TEMPLATE = """
<html>
<head>
  <style>
    body { background: #f5f7fb; font-family: 'Helvetica Neue', Arial, sans-serif; color: #1f2933; margin: 0; }
    .wrapper { max-width: 760px; margin: 0 auto; padding: 32px; }
    .card { background: #ffffff; border-radius: 16px; padding: 24px 28px; margin-bottom: 20px; box-shadow: 0 18px 40px rgba(15, 23, 42, 0.08); }
    h2 { color: #4c49ff; margin-bottom: 16px; }
    h3 { color: #7f56d9; margin-bottom: 12px; }
    p { line-height: 1.6; }
    ul { margin: 0; padding-left: 22px; }
    li { margin-bottom: 8px; }
    .header { margin-bottom: 24px; }
    .muted { color: #6b7280; }
    .footer { margin-top: 28px; font-size: 0.95rem; }
    .badge { display: inline-block; background: #4c49ff; color: #fff; padding: 4px 10px; border-radius: 999px; font-size: 0.8rem; letter-spacing: .05em; text-transform: uppercase; }
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="header">
      <p>Good Morning {{ dealership_name }},</p>
      <p>Congratulations on another successful month of learning on RockED. Below you’ll find standout wins at your store, how you stacked up across the full leaderboard, and one key area of improvement to keep momentum rolling into {{ month }}.</p>
    </div>

    <div class="card">
      <span class="badge">Topic of the Month Insights</span>
      <div style="margin-top:16px;">{{ tom_insight_summary|safe }}</div>
    </div>

    <div class="card">
      <h3>Standout Performers at {{ dealership_name }}:</h3>
      <ul>
      {% for s in standouts %}
        <li><strong>{{ s.Employee_Name }}</strong> — {{ s.Active_Days }} active days, streak {{ s.Current_Streak }} days, {{ s.Journeys_Completed }} journeys completed.</li>
      {% endfor %}
      </ul>
    </div>

    <div class="card">
      <h3>Least Active Learners (4 or fewer learning days in {{ prev_month }}):</h3>
      <p class="muted">{{ least_active|join(', ') if least_active else 'None this month — great job team!' }}</p>
    </div>

    <div class="footer">
      <p>If you have any questions or would like my support in driving results for your team please reach out. I’m happy to prescribe content, set up personalized learning paths, or provide monthly reporting.</p>
      <p>Looking forward to a strong {{ month }} ahead!</p>
      <p>Best,</p>
    </div>
  </div>
</body>
</html>
"""


if st.button("Run script"):
    if not (learning_activity_file and tom_report):
        st.error("Please upload both the learning activity and TOM reports")
    else:
        try:
            # Read learning activity; support different possible layouts more robustly
            dealership_df = pd.read_excel(learning_activity_file, sheet_name='Report', skiprows=5)
            # normalize column names
            dealership_df.columns = [str(c).strip() for c in dealership_df.columns]

            # Build a clean dataframe where each column is aligned
            df_clean = dealership_df.dropna(how='all')

            # Extract dealership name (prefer content in sheet, else sanitized filename)
            dealership_name = extract_dealership_name(df_clean, learning_activity_file.name)

            month = datetime.now().strftime('%B')
            prev_month = (datetime.now().replace(day=1) - pd.Timedelta(days=1)).strftime('%B')

            # Prepare leaderboards / standout logic
            # we expect columns like 'Employee Name', 'Active Days', 'Current Streak (Days)', 'Journeys Completed'
            colmap = {c.lower(): c for c in df_clean.columns}
            def find_col(key):
                for k, v in colmap.items():
                    if key in k:
                        return v
                return None

            emp_col = find_col('employee') or find_col('name')
            active_col = find_col('active days') or find_col('active')
            streak_col = find_col('current streak') or find_col('streak')
            journeys_col = find_col('journeys') or find_col('journeys completed')

            if emp_col is None or active_col is None:
                st.error('Could not find required columns (employee name / active days) in the learning activity file')
                raise SystemExit

            df_clean[emp_col] = df_clean[emp_col].astype(str).str.title()

            total_teammates = df_clean.shape[0]

            # Standouts: top 3 by active days then streak
            ranked = df_clean.sort_values([active_col, streak_col] if streak_col in df_clean.columns else [active_col], ascending=False)
            standouts_df = ranked.head(3).copy()
            # normalize fields for template
            standouts = []
            for _, r in standouts_df.iterrows():
                standouts.append({
                    'Employee_Name': r.get(emp_col, ''),
                    'Active_Days': int(r.get(active_col, 0)) if pd.notna(r.get(active_col, 0)) else 0,
                    'Current_Streak': int(r.get(streak_col, 0)) if streak_col and pd.notna(r.get(streak_col, 0)) else 0,
                    'Journeys_Completed': int(r.get(journeys_col, 0)) if journeys_col and pd.notna(r.get(journeys_col, 0)) else 0,
                })

            least_active_series = df_clean[df_clean[active_col] <= 4][emp_col].tolist() if active_col in df_clean.columns else []
            least_active = title_case_list(least_active_series)

            tom_df = pd.read_excel(tom_report, skiprows=4)
            tom_df.columns = [str(c).strip() for c in tom_df.columns]
            tom_parsed_df, parsed_col = parse_dataframe_attributions(tom_df, column_name_contains='Attribution')

            tom_completions = extract_tom_completions(tom_parsed_df, tom_df)

            completion_pct = 0
            if total_teammates:
                completion_pct = round(100 * tom_completions / total_teammates, 1)
            standout_summaries = [
                f"{s['Employee_Name']} — {s['Active_Days']} active days, streak {s['Current_Streak']} days, {s['Journeys_Completed']} journeys"
                for s in standouts
            ]
            least_list = ', '.join(least_active[:8]) if least_active else 'None this month'

            tom_insight_summary = (
                "✨ <strong>Area for Improvement – Topic of the Month (ToM)</strong><br>"
                f"• Completion Gap: {tom_completions} of {total_teammates} teammates ({completion_pct}%) have wrapped the ToM—plan a fun store challenge to push past 85% by mid-month!<br>"
                "• Business Boost: When every teammate finishes the ToM, stores typically unlock 4-6 extra upsells per seller—there’s upbeat revenue sitting on the table.<br>"
                "• Next Move: Host a 15-minute ToM huddle each Friday, celebrate completions in your group chat, and have managers check progress twice a week.<br><br>"
                "🌟 <strong>Standout Performers</strong><br>"
                + ("• " + "<br>• ".join(standout_summaries) if standout_summaries else "• No standout performers identified this month.<br>")
                + "<br>"
                "🚀 <strong>Coaching List</strong><br>"
                f"• {least_list} — invite them to a quick tune-up session and set a three-day completion goal."
            )
            api_key = None
            if 'OPENAI_API_KEY' in st.secrets:
                api_key = st.secrets['OPENAI_API_KEY']
            elif os.getenv('OPENAI_API_KEY'):
                api_key = os.getenv('OPENAI_API_KEY')

            ai_generated = False
            if api_key:
                try:
                    # Build messages and optionally show them for debugging
                    system_msg, user_content, messages = build_insight_messages({
                        'dealership_name': dealership_name,
                        'month': month,
                        'prev_month': prev_month,
                        'total_teammates': total_teammates,
                        'tom_completions': tom_completions,
                        'standouts': standouts,
                        'least_active': least_active,
                        'df_clean': df_clean,
                    })

                    insight = generate_insight_from_metrics(api_key, {
                        'dealership_name': dealership_name,
                        'month': month,
                        'prev_month': prev_month,
                        'total_teammates': total_teammates,
                        'tom_completions': tom_completions,
                        'standouts': standouts,
                        'least_active': least_active,
                        'df_clean': df_clean,
                    })
                    if insight and not insight.startswith('AI generation error') and not insight.startswith('AI client error'):
                        tom_insight_summary = insight
                        ai_generated = True
                        st.session_state['ai_insight'] = tom_insight_summary
                except Exception as e:
                    st.warning(f'AI insight generation failed, using default summary. Details: {e}')
            else:
                st.info('No OPENAI_API_KEY found; using the built-in insight summary.')

            template = Template(EMAIL_TEMPLATE)
            # prefer any AI insight already stored in session_state
            session_ai = st.session_state.get('ai_insight') if 'ai_insight' in st.session_state else None
            summary_text = session_ai or tom_insight_summary
            if markdown2:
                tom_insight_html = markdown2.markdown(summary_text)
            else:
                tom_insight_html = summary_text.replace('\n', '<br>')
            html_body = template.render(
                dealership_name=dealership_name,
                month=month,
                prev_month=prev_month,
                total_teammates=total_teammates,
                tom_completions=tom_completions,
                tom_insight_summary=tom_insight_html,
                standouts=standouts,
                least_active=least_active,
            )

            st.write("Learning Activity File:", learning_activity_file.name)
            st.write("TOM Report File:", tom_report.name)
            st.subheader("Email Draft")
            subject = f'{prev_month} RockED Recap - {dealership_name}'
            st.write(f"**Email Subject:** {subject}")
            st.markdown(html_body, unsafe_allow_html=True)
            st.download_button(
                label="Download Email Contents",
                data=html_body,
                file_name=f'{dealership_name}_{prev_month}_recap.html',
                mime='text/html'
            )
            if PDF_RENDERER:
                pdf_bytes = PDF_RENDERER(html_body)
                if pdf_bytes:
                    st.download_button(
                        label="Download PDF",
                        data=pdf_bytes,
                        file_name=f'{dealership_name}_{prev_month}_recap.pdf',
                        mime='application/pdf'
                    )
                else:
                    st.warning('Unable to generate PDF with the available renderer.')
            if ai_generated:
                st.success("Processing complete! AI insight applied ✨")
            else:
                st.success("Processing complete!")
        except Exception as e:
            st.error(f"Error loading files: {e}")

if learning_activity_file and tom_report:
    api_key = None
    if 'OPENAI_API_KEY' in st.secrets:
        api_key = st.secrets['OPENAI_API_KEY']
    elif os.getenv('OPENAI_API_KEY'):
        api_key = os.getenv('OPENAI_API_KEY')

    try:
        metrics = compute_metrics(learning_activity_file, tom_report, checkin_file=checkin_report)
    except Exception as e:
        st.warning(f'Could not compute metrics for auto-insight: {e}')
        metrics = None

    # Only auto-generate if the user enabled the checkbox. Otherwise wait for
    # the user to press 'Run script' which will produce the insight.
    if metrics and api_key:
        with st.spinner('Generating AI insight...'):
            insight = generate_insight_from_metrics(api_key, metrics)
        st.session_state['ai_insight'] = insight
        # Sanitize dealership name to make a safe filename (remove path
        # separators and any characters that could create directories).
        raw_name = str(metrics.get('dealership_name', '') or '')
        # replace runs of non-alphanum/dot/underscore/dash with an underscore
        safe_name = re.sub(r'[^A-Za-z0-9._-]+', '_', raw_name).strip('_') or 'dealership'
        out_dir = os.path.join(os.path.dirname(__file__), '..', 'ai_output')
        out_dir = os.path.normpath(out_dir)
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            # best-effort: fall back to workspace-relative ai_output
            out_dir = os.path.join(os.getcwd(), 'ai_output')
            os.makedirs(out_dir, exist_ok=True)

        out_file = os.path.join(out_dir, f'ai_insight_{safe_name}_{metrics["prev_month"]}.txt')
        try:
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write((insight or '') + '\n')
            st.success(f'AI insight generated and saved to {out_file}')
        except Exception as e:
            st.warning(f'AI insight generated but could not be saved to disk: {e}')
        st.write('AI Insight:')
        st.write(insight)
        # The main download button (above) will include the AI insight because
        # we saved it to `st.session_state['ai_insight']` and the rendering
        # code prefers it when present. No separate download button needed.
    elif metrics and not api_key:
        st.info('Add an OPENAI_API_KEY to Streamlit secrets or your environment to generate AI insights automatically.')
