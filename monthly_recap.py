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


st.title("Hendrick Monthly Report Builder")

user_engagement_report = st.file_uploader('Upload the User Engagement Report')
tom_report = st.file_uploader('Upload the ToM report (optional)')

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


def extract_dealership_name(df_clean: pd.DataFrame, filename: str, checkin_file=None) -> dict:
    """Group employees by dealership using the Dealership Name column.

    Returns a dictionary where keys are dealership names and values are lists of employee metrics.
    Each employee metric includes first name, last name, learning days, and total learning units.
    """
    dealership_mapping = {}

    # Iterate through rows to map dealerships and employee metrics
    for _, row in df_clean.iterrows():
        dealership_name = row.get('Dealership Name', '').strip()
        first_name = row.get('First Name', '').strip()
        last_name = row.get('Last Name', '').strip()
        learning_days = row.get('Learning Days', 0)
        total_units = row.get('Total Learning Units', 0)

        if dealership_name and first_name and last_name:
            if dealership_name not in dealership_mapping:
                dealership_mapping[dealership_name] = []
            dealership_mapping[dealership_name].append({
                'First Name': first_name,
                'Last Name': last_name,
                'Learning Days': learning_days,
                'Total Learning Units': total_units
            })

    return dealership_mapping


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


def compute_metrics(user_engagement_report, tom_report=None):
    """Return computed metrics needed for the email and insight generation.

    Returns dict with keys: dealership_name, month, prev_month, total_teammates,
    standouts (list of dicts), least_active (list), tom_completions, df_clean, dealership_mapping.
    """
    # Read the Excel file and try to locate the correct sheet.
    try:
        # Streamlit UploadedFile is file-like; ensure pointer is at start before reading
        try:
            user_engagement_report.seek(0)
        except Exception:
            pass
        xls = pd.ExcelFile(user_engagement_report)
        sheets = list(xls.sheet_names)
    except Exception as e:
        raise ValueError(f'Could not open user engagement Excel: {e}')

    # Prefer exact match 'Report', then any sheet with 'report' in the name, else fall back to first sheet
    report_sheet = None
    for s in sheets:
        if s.strip().lower() == 'report':
            report_sheet = s
            break
    if report_sheet is None:
        for s in sheets:
            if 'report' in s.strip().lower():
                report_sheet = s
                break
    if report_sheet is None:
        report_sheet = sheets[0]

    # Try reading with the expected skiprows; if that fails, try without skiprows
    try:
        try:
            user_engagement_report.seek(0)
        except Exception:
            pass
        dealership_df = pd.read_excel(user_engagement_report, sheet_name=report_sheet, skiprows=5)
    except Exception:
        try:
            user_engagement_report.seek(0)
        except Exception:
            pass
        dealership_df = pd.read_excel(user_engagement_report, sheet_name=report_sheet)

    # If a dict was returned (some code-paths or older logic may call read_excel with sheet_name=None),
    # pick the 'Report' sheet if present, otherwise the first sheet.
    if isinstance(dealership_df, dict):
        # normalize keys
        keys = list(dealership_df.keys())
        picked = None
        for k in keys:
            if str(k).strip().lower() == 'report':
                picked = k
                break
        if picked is None:
            for k in keys:
                if 'report' in str(k).strip().lower():
                    picked = k
                    break
        if picked is None:
            picked = keys[0]
        dealership_df = dealership_df[picked]

    dealership_df.columns = [str(c).strip() for c in dealership_df.columns]

    # Heuristic: sometimes pandas picks a data row as the header (common when
    # skiprows is wrong). If none of the expected header-like substrings are
    # present in the column names, try to find the real header row within the
    # first N rows and re-read the sheet using that row as the header.
    def _looks_like_header(cols):
        header_text = ' '.join([str(c).lower() for c in cols if isinstance(c, str)])
        clues = ['dealership', 'first name', 'last name', 'learning', 'total learning', 'employee', 'name', 'active', 'days']
        return any(clue in header_text for clue in clues)

    df_clean = dealership_df.dropna(how='all')

    # If the current columns don't look like a header, inspect the first few
    # rows to find a likely header row and re-read the sheet with header=index.
    if not _looks_like_header(df_clean.columns):
        try:
            # read a small preview of the sheet as raw values (no header)
            try:
                user_engagement_report.seek(0)
            except Exception:
                pass
            preview = pd.read_excel(user_engagement_report, sheet_name=report_sheet, header=None, nrows=12)
            header_row = None
            aliases = ['dealership', 'first name', 'last name', 'learning', 'total learning', 'employee', 'name', 'active', 'days']
            for i in range(min(len(preview), 12)):
                row_vals = preview.iloc[i].astype(str).str.lower().tolist()
                matches = 0
                for v in row_vals:
                    if any(a in v for a in aliases):
                        matches += 1
                # If two or more alias-like terms appear in the row, treat it as header
                if matches >= 2:
                    header_row = i
                    break

            if header_row is not None:
                try:
                    user_engagement_report.seek(0)
                except Exception:
                    pass
                # re-read using the discovered header row
                dealership_df = pd.read_excel(user_engagement_report, sheet_name=report_sheet, header=header_row)
                dealership_df.columns = [str(c).strip() for c in dealership_df.columns]
                df_clean = dealership_df.dropna(how='all')
        except Exception:
            # If anything goes wrong in the recovery attempt, continue with the
            # original df_clean and let downstream detection handle it.
            df_clean = dealership_df.dropna(how='all')

    # extract mapping (dealership -> list of employee metrics)
    dealership_mapping = extract_dealership_name(df_clean, user_engagement_report.name)

    # pick a default dealership name for single-dealer workflows (first key)
    dealership_name = ''
    if isinstance(dealership_mapping, dict) and dealership_mapping:
        dealership_name = next(iter(dealership_mapping.keys()))

    today = datetime.now()
    month = today.strftime('%B')
    # last calendar month name
    prev_month = (today.replace(day=1) - pd.Timedelta(days=1)).strftime('%B')

    # Decide which month the report is for. If we're in the very first few
    # days of the month (e.g., 1st-3rd) the report typically targets the
    # previous month; otherwise it targets the current month.
    if today.day <= 3:
        report_month = prev_month
        # report_date is last day of previous month
        report_date = today.replace(day=1) - pd.Timedelta(days=1)
    else:
        report_month = month
        report_date = today
    report_year = report_date.year

    # Build lookup of lowercase column names -> original
    colmap = {str(c).strip().lower(): c for c in df_clean.columns}

    def find_col_by_alias(aliases):
        # aliases: list of possible substrings/variants to match
        for a in aliases:
            a = a.strip().lower()
            # exact match first
            if a in colmap:
                return colmap[a]
        # contains match
        for k, orig in colmap.items():
            for a in aliases:
                if a in k:
                    return orig
        return None

    # Common alias lists
    employee_aliases = ['employee', 'name', 'employee name', 'full name', 'user name', 'first name']
    first_name_aliases = ['first name', 'firstname', 'given name']
    last_name_aliases = ['last name', 'lastname', 'surname']
    active_aliases = ['learning days', 'active days', 'active', 'days active', 'learning_days', 'learningdays']
    journeys_aliases = ['journeys', 'journeys completed', 'journey', 'journeys_completed']

    emp_col = find_col_by_alias(employee_aliases)
    active_col = find_col_by_alias(active_aliases)
    streak_col = find_col_by_alias(['current streak', 'streak'])
    journeys_col = find_col_by_alias(journeys_aliases)

    # If emp_col not found, attempt to combine First Name + Last Name
    if emp_col is None:
        first_col = find_col_by_alias(first_name_aliases)
        last_col = find_col_by_alias(last_name_aliases)
        if first_col and last_col:
            # create a synthetic column in df_clean
            df_clean['__employee_combined__'] = (df_clean[first_col].fillna('').astype(str).str.strip() + ' ' + df_clean[last_col].fillna('').astype(str).str.strip()).str.strip()
            emp_col = '__employee_combined__'

    # If still missing required columns, provide an interactive Streamlit fallback so user can choose
    missing = []
    if emp_col is None:
        missing.append('employee name')
    if active_col is None:
        missing.append('active days')

    if missing:
        # If running inside Streamlit, offer select boxes to choose columns; otherwise raise informative error
        try:
            # show available columns
            cols = list(df_clean.columns)
            st.warning(f"Could not auto-detect required columns: {', '.join(missing)}. Please select them below.")
            emp_choice = None
            act_choice = None
            if 'employee name' in missing:
                emp_choice = st.selectbox('Select employee name column', options=[''] + cols, index=0)
                if emp_choice:
                    emp_col = emp_choice
            if 'active days' in missing:
                act_choice = st.selectbox('Select active days column', options=[''] + cols, index=0)
                if act_choice:
                    active_col = act_choice
            # if user provided selections, continue; else raise
            if emp_col is None or active_col is None:
                raise ValueError(f"Could not find required columns ({', '.join(missing)}) in the learning activity file. Available columns: {cols}")
        except Exception:
            cols = list(df_clean.columns)
            raise ValueError(f"Could not find required columns ({', '.join(missing)}) in the learning activity file. Available columns: {cols}")

    # Normalize types
    if emp_col not in df_clean.columns:
        raise ValueError(f'Employee column {emp_col} not in dataframe columns: {list(df_clean.columns)}')
    if active_col not in df_clean.columns:
        raise ValueError(f'Active days column {active_col} not in dataframe columns: {list(df_clean.columns)}')

    # Make sure active_col is numeric
    try:
        df_clean[active_col] = pd.to_numeric(df_clean[active_col], errors='coerce').fillna(0).astype(int)
    except Exception:
        df_clean[active_col] = pd.to_numeric(df_clean[active_col].astype(str).str.extract(r'(\d+)').fillna(0)[0], errors='coerce').fillna(0).astype(int)

    # Ensure employee names are title-cased
    df_clean[emp_col] = df_clean[emp_col].astype(str).str.strip().replace('nan', '')
    df_clean[emp_col] = df_clean[emp_col].apply(lambda x: ' '.join([p.capitalize() for p in str(x).split()]) if x else '')

    total_teammates = df_clean.shape[0]

    # Rankings and standouts
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

    tom_completions = 0
    if tom_report is not None:
        try:
            try:
                tom_report.seek(0)
            except Exception:
                pass
            tom_df = pd.read_excel(tom_report, skiprows=4)
            tom_df.columns = [str(c).strip() for c in tom_df.columns]
            tom_parsed_df, parsed_col = parse_dataframe_attributions(tom_df, column_name_contains='Attribution')

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
        except Exception:
            tom_completions = 0

    return {
        'dealership_name': dealership_name,
        'month': month,
        # keep legacy keys: 'prev_month' is the month being recapped (report_month)
        'prev_month': report_month,
        'report_year': report_year,
        'total_teammates': total_teammates,
        'standouts': standouts,
        'least_active': least_active,
        'tom_completions': tom_completions,
        'df_clean': df_clean,
        'dealership_mapping': dealership_mapping,
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
    if not user_engagement_report:
        st.error("Please upload the user engagement report")
    else:
        try:
            # Use compute_metrics to robustly read the user engagement sheet and compute derived metrics
            metrics = compute_metrics(user_engagement_report, tom_report)
            df_clean = metrics.get('df_clean')
            dealership_mapping = metrics.get('dealership_mapping') or {}
            # choose a default dealership name if present
            dealership_name = metrics.get('dealership_name') or (next(iter(dealership_mapping.keys())) if dealership_mapping else '')
            month = metrics.get('month')
            prev_month = metrics.get('prev_month')
            total_teammates = metrics.get('total_teammates')
            standouts = metrics.get('standouts') or []
            least_active = metrics.get('least_active') or []
            tom_completions = metrics.get('tom_completions') or 0

            # df_clean should already be normalized by compute_metrics; proceed to rendering

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

            st.write("Learning Activity File:", user_engagement_report.name)
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
                    report_month_name = metrics.get('prev_month', prev_month) if 'metrics' in locals() else prev_month
                    report_year_val = metrics.get('report_year', datetime.now().year) if 'metrics' in locals() else datetime.now().year
                    st.download_button(
                        label="Download PDF",
                        data=pdf_bytes,
                        file_name=f'{dealership_name}_{report_month_name}_{report_year_val}_recap.pdf',
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

if user_engagement_report:
    api_key = None
    if 'OPENAI_API_KEY' in st.secrets:
        api_key = st.secrets['OPENAI_API_KEY']
    elif os.getenv('OPENAI_API_KEY'):
        api_key = os.getenv('OPENAI_API_KEY')

    try:
        metrics = compute_metrics(user_engagement_report, tom_report)
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

def generate_per_dealer_reports(user_engagement_report, tom_report, output_dir='ai_output', generate_pdf=False, api_key=None):
    """Generate one HTML (and optionally PDF) report per dealership.

    - Reads user_engagement_report to build per-dealer metrics
    - Calls the insight generator to produce an ai summary for each dealer
    - Renders the EMAIL_TEMPLATE and writes HTML files into output_dir
    - If generate_pdf=True and weasyprint is importable, render PDFs as well
    Returns a list of generated file paths.
    """
    os.makedirs(output_dir, exist_ok=True)

    # If PDF generation was requested, ensure reportlab is installed. We will
    # attempt to install it automatically using ensure_package; if that fails
    # we raise so the UI can notify the user rather than silently falling back.
    reportlab_available = False
    if generate_pdf:
        reportlab_available = ensure_package('reportlab')
        if not reportlab_available:
            raise RuntimeError("reportlab is not available and could not be installed automatically. Run `pip install reportlab` and try again.")
        # import the ReportLab symbols now that the package is available
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet

    metrics = compute_metrics(user_engagement_report, tom_report)
    df_clean = metrics.get('df_clean')
    month = metrics.get('month')
    prev_month = metrics.get('prev_month')

    dealership_mapping = metrics.get('dealership_mapping') or {}
    generated = []

    for dealer, employees in dealership_mapping.items():
        # Build a small DataFrame for this dealer to compute standouts/least_active
        dealer_df = pd.DataFrame(employees)
        # normalize column names
        if 'Learning Days' in dealer_df.columns:
            active_col = 'Learning Days'
        else:
            # fallback
            active_col = dealer_df.columns[2] if len(dealer_df.columns) > 2 else None

        # Compute per-dealer aggregates
        total_teammates = len(employees)
        total_units = int(dealer_df['Total Learning Units'].astype(float).sum()) if 'Total Learning Units' in dealer_df.columns else 0
        avg_active = round(dealer_df['Learning Days'].astype(float).mean(), 1) if 'Learning Days' in dealer_df.columns and total_teammates else None

        # standouts: top 3 by Learning Days
        standouts = []
        try:
            sorted_df = dealer_df.sort_values(by='Learning Days', ascending=False)
            top_df = sorted_df.head(3)
            for _, r in top_df.iterrows():
                standouts.append({
                    'Employee_Name': f"{r.get('First Name','')} {r.get('Last Name','')}",
                    'Active_Days': int(r.get('Learning Days') or 0),
                    'Current_Streak': 0,
                    'Journeys_Completed': 0,
                })
        except Exception:
            standouts = []

        least_active = dealer_df[dealer_df['Learning Days'].astype(float) <= 4][['First Name','Last Name']].apply(lambda r: f"{r[0]} {r[1]}", axis=1).tolist() if 'Learning Days' in dealer_df.columns else []

        # Build metrics for the AI generator
        dealer_metrics = {
            'dealership_name': dealer,
            'month': month,
            'prev_month': prev_month,
            'total_teammates': total_teammates,
            'standouts': standouts,
            'least_active': least_active,
            'tom_completions': metrics.get('tom_completions', 0),
            'df_clean': dealer_df,
        }

        # Generate AI insight (best-effort; may return error string)
        ai_text = ''
        if api_key:
            try:
                ai_text = generate_insight_from_metrics(api_key, dealer_metrics)
            except Exception as e:
                ai_text = f"AI generation error: {e}"
        else:
            # fallback short summary
            ai_text = f"{dealer}: {total_teammates} teammates — {total_units} total learning units, avg days: {avg_active}"

        # Render template to HTML string (for WeasyPrint) but do NOT save HTML to disk; we will create PDF only
        tpl = Template(EMAIL_TEMPLATE)
        if markdown2:
            tom_insight_html = markdown2.markdown(ai_text)
        else:
            tom_insight_html = ai_text.replace('\n', '<br>')
        html_body = tpl.render(
            dealership_name=dealer,
            month=month,
            prev_month=prev_month,
            total_teammates=total_teammates,
            tom_completions=metrics.get('tom_completions', 0),
            tom_insight_summary=tom_insight_html,
            standouts=standouts,
            least_active=least_active,
        )

        # Create a safe name for files (preserve spaces so filenames read like the dealership name)
        safe_name = re.sub(r"[^0-9A-Za-z ._-]", '', dealer).strip()
        # Use the report month/year determined during compute_metrics
        date_suffix = f"{metrics.get('prev_month', prev_month)} {metrics.get('report_year', datetime.now().year)}"

        # We require reportlab (the ensure_package call above should have
        # installed/imported it). Generate the PDF for this dealer.
        file_name = f"{safe_name} - {date_suffix}.pdf"
        pdf_path = os.path.join(output_dir, file_name)

        try:
            doc = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elems = []
            elems.append(Paragraph(f"{dealer} - {metrics.get('prev_month', month)} {metrics.get('report_year', datetime.now().year)} Recap", styles['Title']))
            elems.append(Spacer(1, 12))

            # Build table data
            table_data = [['First Name', 'Last Name', 'Learning Days', 'Total Learning Units']]
            for emp in employees:
                table_data.append([
                    emp.get('First Name', ''),
                    emp.get('Last Name', ''),
                    str(emp.get('Learning Days', '')),
                    str(emp.get('Total Learning Units', '')),
                ])

            tbl = Table(table_data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4c49ff')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ]))
            elems.append(tbl)
            doc.build(elems)
            generated.append(pdf_path)
        except Exception as e:
            generated.append(f"PDF_not_generated:{dealer}:{e}")

    return generated


# Streamlit helper: generate per-dealer reports (PDF only) on demand
if st.button('Generate per-dealer PDFs'):
    if not user_engagement_report:
        st.error('Upload a User Engagement Report first')
    else:
        api_key = None
        if 'OPENAI_API_KEY' in st.secrets:
            api_key = st.secrets['OPENAI_API_KEY']
        elif os.getenv('OPENAI_API_KEY'):
            api_key = os.getenv('OPENAI_API_KEY')
        out_dir = os.path.join(os.getcwd(), 'ai_output')
        os.makedirs(out_dir, exist_ok=True)
        with st.spinner('Generating per-dealer PDFs...'):
            results = generate_per_dealer_reports(user_engagement_report, tom_report, output_dir=out_dir, generate_pdf=True, api_key=api_key)

        # Persist successful PDF paths in session_state so download buttons remain
        if 'generated_pdfs' not in st.session_state:
            st.session_state['generated_pdfs'] = []

        # Show each generated PDF immediately under its success message
        just_generated = []
        for p in results:
            if isinstance(p, str) and p.startswith('PDF_not_generated:'):
                st.warning(p)
                continue

            if p.lower().endswith('.pdf'):
                name = os.path.basename(p)
                st.success(f'PDF generated: {name}')
                # add to persistent list if not already present
                if p not in st.session_state['generated_pdfs']:
                    st.session_state['generated_pdfs'].append(p)
                just_generated.append(p)

                # render the download button directly under the success message
                try:
                    if os.path.exists(p):
                        with open(p, 'rb') as fh:
                            pdfb = fh.read()
                        st.download_button(f'Download {name}', data=pdfb, file_name=name, mime='application/pdf', key=f'dl-{name}-{len(st.session_state["generated_pdfs"]) }')
                    else:
                        st.warning(f'Generated file missing after creation: {p}')
                except Exception as e:
                    st.warning(f'Could not render download for {p}: {e}')

        # Show previously generated PDFs (from earlier runs) that were not in this just-generated batch
        previous = [p for p in st.session_state.get('generated_pdfs', []) if p not in just_generated]
        if previous:
            st.markdown('---')
            st.markdown('**Previously generated PDFs**')
            for p in previous:
                try:
                    if os.path.exists(p):
                        name = os.path.basename(p)
                        with open(p, 'rb') as fh:
                            pdfb = fh.read()
                        st.download_button(f'Download {name}', data=pdfb, file_name=name, mime='application/pdf', key=f'prev-{name}-{st.session_state["generated_pdfs"].index(p)}')
                    else:
                        st.warning(f'Previously generated file missing: {p}')
                except Exception as e:
                    st.warning(f'Could not render download for {p}: {e}')


# Ensure persisted generated PDFs are always visible across reruns (downloads trigger reruns)
if 'generated_pdfs' in st.session_state and st.session_state['generated_pdfs']:
    st.markdown('---')
    st.markdown('**Available generated PDFs**')
    for idx, p in enumerate(st.session_state['generated_pdfs']):
        try:
            if os.path.exists(p):
                name = os.path.basename(p)
                with open(p, 'rb') as fh:
                    pdfb = fh.read()
                # use a stable key per file so Streamlit doesn't remap widgets on rerun
                st.download_button(f'Download {name}', data=pdfb, file_name=name, mime='application/pdf', key=f'available-{idx}-{name}')
            else:
                st.warning(f'Previously generated file missing: {p}')
        except Exception as e:
            st.warning(f'Could not render download for {p}: {e}')
