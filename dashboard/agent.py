"""
agent.py  –  Gemini-powered financial data extraction agent.

Replaces the brittle regex/alias-based parser with a Gemini 2.5 Flash
model that reads the raw CSV content and returns a fully structured,
period-keyed financial dictionary — the same format that excel_writer.py expects.

The agent:
  1. Reads all uploaded CSV files as plain text.
  2. Sends them to Gemini with a detailed extraction prompt.
  3. Receives a clean JSON dict { period_key: { metric: value } }.
  4. Runs compute_derived pass for formulas not directly available.
  5. Returns the final data dict ready for excel_writer.py.
"""

import json
import re
import os
from pathlib import Path
from google import genai
from google.genai import types


# ---------------------------------------------------------------------------
# System prompt — tells Gemini exactly what to extract and how to format it
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """
You are a senior financial analyst and data extraction expert specialising in
Indian listed companies. You will receive raw CSV exports from a financial data
platform (Screener / Ace Equity / ARQ).

There are typically 6 CSV files per company:
  File A – Quarterly P&L (columns are quarters like Jun-15, Sep-15, Dec-15, Mar-16…)
  File B – Half-yearly Balance Sheet (columns are Sep-XX, Mar-XX…)
  File C – Half-yearly Cash Flow
  File D – Annual P&L (detailed, note-level; columns are FY years like Mar 2016, Mar 2017…)
  File E – Annual Balance Sheet (detailed annual)
  File F – Annual Cash Flow (detailed annual)

YOUR TASK:
Extract ALL financial data from ALL files and return a single JSON object.

JSON FORMAT:
{
  "1QFY16": { ...metrics... },
  "2QFY16": { ...metrics... },
  "3QFY16": { ...metrics... },
  "4QFY16": { ...metrics... },
  "FY16":   { ...metrics... },
  "1QFY17": { ...metrics... },
  ...and so on for all periods found
}

PERIOD KEY RULES (Indian Fiscal Year — April to March):
- Quarters map by month-end:
    Jun → 1Q  (Apr–Jun)
    Sep → 2Q  (Jul–Sep)
    Dec → 3Q  (Oct–Dec)
    Mar → 4Q  (Jan–Mar)
- FY number = calendar year of March end
    Jun-15 → 1QFY16   (FY16 = Apr 2015 – Mar 2016)
    Sep-15 → 2QFY16
    Dec-15 → 3QFY16
    Mar-16 → 4QFY16 AND FY16
    Jun-16 → 1QFY17
    Mar 2017 annual → FY17
- Use 2-digit FY: FY16, FY17, FY18 … FY26.

METRICS TO EXTRACT (use EXACTLY these key names):
P&L (quarterly and annual):
  gross_sales        - Gross Sales / Revenue from Operations (Gross)
  excise_duty        - Less: Excise Duty (if present, else 0)
  net_sales          - Net Sales / Revenue from Operations (Net)
  other_op_income    - Other Operating Income / Other Operating Revenue
  total_income       - Total Income from Operations
  raw_material       - Raw Material Consumed / Cost of Materials Consumed
  stock_adj          - Changes in Inventories (can be negative)
  purchase_fg        - Purchase of Finished Goods / Purchases of Stock-in-Trade
  employee_exp       - Employee Benefits Expense / Employee Cost
  power_fuel         - Power & Fuel / Electricity (extract separately for quarterly)
  other_exp          - Other Expenses (for quarterly: exclude Power/Fuel if listed separately)
  total_expenditure  - Total Expenses / Total Expenditure
  other_income       - Other Income (non-operating)
  interest           - Finance Costs / Interest
  depreciation       - Depreciation & Amortisation
  exceptional_items  - Exceptional Items (positive = income, negative = charge)
  pbt                - Profit Before Tax
  tax_current        - Current Tax
  deferred_tax       - Deferred Tax
  total_tax          - Total Tax (current + deferred)
  net_profit_after_tax - Net Profit After Tax / PAT
  minority_interest  - Minority Interest / Non-Controlling Interest
  net_profit_mi      - Net Profit after Minority Interest
  eps_basic          - Basic EPS

Balance Sheet (quarterly from half-yearly file, annual from annual file):
  equity             - Share Capital / Equity Share Capital
  reserves           - Reserves & Surplus / Other Equity
  lt_borrowings      - Long-Term Borrowings
  st_borrowings      - Short-Term Borrowings / Current Borrowings
  deferred_tax_liab  - Deferred Tax Liabilities
  net_block          - Net Block / Fixed Assets / Property Plant & Equipment
  cash               - Cash & Cash Equivalents
  inventories        - Inventories
  trade_receivables  - Trade Receivables
  total_current_assets     - Total Current Assets
  total_current_liabilities - Total Current Liabilities
  trade_payables     - Trade Payables
  total_assets       - Total Assets

Cash Flow:
  ocf               - Net Cash from Operating Activities
  capex             - Purchase of Fixed Assets / Capital Expenditure (report as negative)
  dividend          - Dividend Paid
  investing_cf      - Net Cash from Investing Activities
  financing_cf      - Net Cash from Financing Activities

CRITICAL RULES:
1. For QUARTERLY P&L (File A):
   - Other Expenses row in this file DOES NOT include Power & Fuel.
   - Extract Power & Fuel separately as power_fuel.
   - Set other_exp = (Other Expenses row value) + power_fuel.
   - If Power & Fuel is not listed separately, treat other_exp as-is.

2. For ANNUAL P&L (File D):
   - The "Other Expenses" in annual files is already an aggregate that
     includes manufacturing expenses, so use it directly as other_exp.
   - Do NOT add power_fuel again for annual periods.

3. Balance Sheet data is a SNAPSHOT (not a flow) — do not sum quarters.
   Use the Mar-XX BS as the year-end snapshot for FYxx.
   Use the Sep-XX BS as the 2Q snapshot for the FY ending next March.

4. If a value is not found in any file for a period, omit that key entirely.
   Do NOT put 0 for unknown values — only omit them.

5. All monetary values should be in the SAME UNIT as the source file
   (typically Rs. Crores). Do not convert units.

6. Numbers can be negative (e.g. stock_adj, exceptional_items).

7. Be thorough — extract ALL periods available across ALL files.
   Cross-reference: quarterly files may cover more recent periods than annual files.

Return ONLY the JSON object — no markdown, no explanation, no code fences.
Just the raw JSON starting with { and ending with }.
"""


# ---------------------------------------------------------------------------
# Main agent function
# ---------------------------------------------------------------------------

def run_agent(csv_dir: str, api_key: str, progress_callback=None) -> dict:
    """
    Use Gemini to extract financial data from all CSV files in csv_dir.
    
    Parameters
    ----------
    csv_dir        : directory containing the uploaded CSV files
    api_key        : Gemini API key
    progress_callback : optional callable(message: str) for UI progress updates
    
    Returns
    -------
    dict  { period_key: { metric: value } }  ready for excel_writer.py
    """
    csv_dir = Path(csv_dir)
    
    def log(msg):
        if progress_callback:
            progress_callback(msg)
    
    # ---- 1. Read all CSVs ----
    log("Reading CSV files...")
    csv_files = sorted(csv_dir.glob("*.csv"))
    if not csv_files:
        raise ValueError(f"No CSV files found in {csv_dir}")
    
    file_contents = []
    for f in csv_files:
        try:
            text = f.read_text(encoding='utf-8', errors='replace')
            file_contents.append(f"=== FILE: {f.name} ===\n{text}\n")
        except Exception as e:
            log(f"Warning: Could not read {f.name}: {e}")
    
    combined_content = "\n".join(file_contents)
    log(f"Loaded {len(csv_files)} files ({len(combined_content):,} characters). Sending to AI agent...")
    
    # ---- 2. Call Gemini ----
    client = genai.Client(api_key=api_key)
    
    user_message = (
        f"Here are {len(csv_files)} CSV financial export files for a company.\n"
        f"Please extract all financial data and return the JSON as instructed.\n\n"
        f"{combined_content}"
    )
    
    log("AI Agent analyzing financial statements (this may take 30-60 seconds)...")
    
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=user_message,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            temperature=0.1,   # low temperature for factual extraction
            max_output_tokens=65536,
            thinking_config=types.ThinkingConfig(
                thinking_budget=8000  # give the model time to think carefully
            )
        )
    )
    
    raw_text = response.text.strip()
    log("AI response received. Parsing JSON...")
    
    # ---- 3. Parse JSON response ----
    # Strip markdown code fences if present
    raw_text = re.sub(r'^```(?:json)?\s*', '', raw_text, flags=re.MULTILINE)
    raw_text = re.sub(r'\s*```\s*$', '', raw_text, flags=re.MULTILINE)
    raw_text = raw_text.strip()
    
    # Find the JSON object
    json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
    if not json_match:
        raise ValueError(f"No JSON found in agent response. Got: {raw_text[:500]}")
    
    try:
        raw_data = json.loads(json_match.group())
    except json.JSONDecodeError as e:
        # Try to fix common JSON issues
        fixed = _repair_json(json_match.group())
        try:
            raw_data = json.loads(fixed)
        except Exception:
            raise ValueError(f"Could not parse agent JSON: {e}\nRaw: {raw_text[:1000]}")
    
    log(f"Extracted {len(raw_data)} periods from AI response.")
    
    # ---- 4. Normalise data ----
    data = _normalize_agent_data(raw_data)
    log(f"Normalised to {len(data)} periods.")
    
    # ---- 5. Compute derived metrics ----
    log("Computing derived metrics (EBITDA, PBT, PAT, ratios)...")
    _compute_derived_agent(data)
    _compute_annual_from_quarters_agent(data)
    _propagate_bs_agent(data)
    
    log(f"Done. {len(data)} periods ready for Excel generation.")
    return data


# ---------------------------------------------------------------------------
# JSON repair helper
# ---------------------------------------------------------------------------

def _repair_json(text: str) -> str:
    """Attempt simple repairs on malformed JSON."""
    # Remove trailing commas before } or ]
    text = re.sub(r',\s*([}\]])', r'\1', text)
    # Replace single quotes with double quotes (basic)
    # Only do this for keys
    text = re.sub(r"'([^']+)':", r'"\1":', text)
    return text


# ---------------------------------------------------------------------------
# Data normalisation
# ---------------------------------------------------------------------------

def _safe_float(val):
    if val is None:
        return None
    try:
        if isinstance(val, str):
            val = val.replace(',', '').strip()
            if val in ('', '-', 'N/A', 'NA', 'n/a', 'nan', 'None'):
                return None
        return float(val)
    except (TypeError, ValueError):
        return None


def _normalize_agent_data(raw: dict) -> dict:
    """
    Convert raw agent output to clean { period_key: { metric: float|None } }.
    Validates period key format and coerces all values to float.
    """
    VALID_PERIOD = re.compile(r'^(\dQFY\d{2}|FY\d{2})$')
    data = {}
    
    for period, metrics in raw.items():
        period = str(period).strip()
        if not VALID_PERIOD.match(period):
            # Try to fix common mistakes like "FY2016" → "FY16", "Q1FY16" → "1QFY16"
            period = _fix_period_key(period)
            if not period:
                continue
        
        if not isinstance(metrics, dict):
            continue
        
        clean = {}
        for k, v in metrics.items():
            k = str(k).strip().lower()
            fv = _safe_float(v)
            if fv is not None:
                clean[k] = fv
        
        if clean:
            data[period] = clean
    
    return data


def _fix_period_key(raw: str) -> str:
    """Try to fix non-standard period keys."""
    # "FY2016" → "FY16"
    m = re.match(r'^FY(20)?(\d{2})$', raw, re.IGNORECASE)
    if m:
        return f'FY{m.group(2)}'
    # "Q1FY16" → "1QFY16"
    m = re.match(r'^Q(\d)FY(\d{2})$', raw, re.IGNORECASE)
    if m:
        return f'{m.group(1)}QFY{m.group(2)}'
    # "1QFY2016" → "1QFY16"
    m = re.match(r'^(\d)QFY(20)?(\d{2})$', raw, re.IGNORECASE)
    if m:
        return f'{m.group(1)}QFY{m.group(3)}'
    return None


# ---------------------------------------------------------------------------
# Derived computations (mirrors parser.py logic)
# ---------------------------------------------------------------------------

def _compute_derived_agent(data: dict):
    """Compute EBITDA, PBT, PAT and all ratio metrics for every period."""
    import re as _re
    
    for period, d in data.items():
        g = d.get

        # Net Sales
        if not g('net_sales'):
            gs = g('gross_sales') or 0
            ex = g('excise_duty') or 0
            if gs:
                d['net_sales'] = gs - ex

        # Total Income
        if not g('total_income'):
            d['total_income'] = (g('net_sales') or 0) + (g('other_op_income') or 0)

        # Other Exp for quarterly: add power_fuel if listed separately
        is_annual = bool(_re.match(r'^FY\d{2}$', period))
        if not is_annual:
            pf = g('power_fuel') or 0
            oe = g('other_exp') or 0
            if pf and oe:
                d['other_exp'] = oe + pf  # agent may have already included it
            elif pf and not oe:
                d['other_exp'] = pf

        # Total Expenditure
        rm  = g('raw_material') or 0
        stk = g('stock_adj') or 0
        pfg = g('purchase_fg') or 0
        emp = g('employee_exp') or 0
        oe  = g('other_exp') or 0
        dep = g('depreciation') or 0

        if not g('total_expenditure'):
            d['total_expenditure'] = rm + stk + pfg + emp + oe

        # EBITDA
        ti = g('total_income') or 0
        tot_exp = g('total_expenditure') or 0
        ebitda = ti - tot_exp
        if not g('ebitda') and ebitda:
            d['ebitda'] = ebitda
        ebitda = g('ebitda') or ebitda

        # PBDT = EBITDA - Interest + Other Income
        interest  = g('interest') or 0
        other_inc = g('other_income') or 0
        d['pbdt'] = ebitda - interest + other_inc

        # PBT
        pbt = g('pbt') or (d['pbdt'] - dep)
        d['pbt'] = pbt

        # Tax
        tax_c = g('tax_current') or g('tax') or 0
        tax_d = g('deferred_tax') or 0
        total_tax = g('total_tax') or (tax_c + tax_d)
        d['total_tax']   = total_tax
        d['tax_current'] = tax_c
        d['deferred_tax'] = tax_d

        # PAT
        pat = g('net_profit_after_tax') or g('net_profit_mi') or (pbt - total_tax)
        d['pat'] = pat

        # PBIT
        d['pbit'] = ebitda - dep

        # COGS
        d['cogs'] = rm + stk + pfg

        # EBITDA Margin
        if ti:
            d['ebitda_margin'] = ebitda / ti
            d['ebit_margin']   = d['pbit'] / ti

        # BS ratios
        eq   = g('equity') or 0
        res  = g('reserves') or 0
        lt   = g('lt_borrowings') or 0
        st   = g('st_borrowings') or 0
        debt = lt + st
        if debt == 0:
            debt = g('debt') or 0
        d['debt'] = debt
        networth = eq + res
        if networth:
            d['de_ratio']     = debt / networth
            d['roe']          = pat / networth if pat else None
        cap_emp = networth + debt
        d['cap_employed']  = cap_emp
        cash = g('cash') or 0
        d['invested_cap'] = cap_emp - cash
        if cap_emp:
            d['roce'] = d['pbit'] / cap_emp if d['pbit'] else None

        # Working Capital
        tca = g('total_current_assets') or 0
        tcl = g('total_current_liabilities') or 0
        if tca:
            d['working_capital'] = tca - tcl

        # EPS / Book Value
        fv = g('face_value') or 2.0
        d['face_value'] = fv
        if eq and fv:
            shares = eq / fv
            d['num_shares']   = shares
            d['eps']          = pat / shares if pat and shares else None
            d['book_value_ps'] = (eq + res) / shares if shares else None


def _compute_annual_from_quarters_agent(data: dict):
    """Fill missing annual P&L metrics by summing 4 quarters."""
    FLOW_METRICS = [
        'gross_sales','excise_duty','net_sales','other_op_income','total_income',
        'raw_material','stock_adj','purchase_fg','employee_exp','other_exp',
        'total_expenditure','other_income','interest','depreciation',
        'total_tax','tax_current','deferred_tax','net_profit_after_tax',
        'pbt','exceptional_items','minority_interest','ocf','capex','dividend',
    ]
    fy_re = re.compile(r'^FY(\d{2})$')
    for period in list(data.keys()):
        m = fy_re.match(period)
        if not m:
            continue
        fy2 = int(m.group(1))
        annual_data = data[period]
        for metric in FLOW_METRICS:
            if metric in annual_data and annual_data[metric] is not None:
                continue
            q_vals = [
                data[f'{q}QFY{fy2:02d}'].get(metric)
                for q in range(1, 5)
                if f'{q}QFY{fy2:02d}' in data
            ]
            valid = [v for v in q_vals if v is not None]
            if valid:
                annual_data[metric] = sum(valid)


def _propagate_bs_agent(data: dict):
    """Copy BS snapshot from 4QFYxx to FYxx and vice versa."""
    BS_SNAP = [
        'equity','reserves','lt_borrowings','st_borrowings','debt',
        'cash','net_block','face_value','total_current_assets',
        'total_current_liabilities','working_capital','trade_receivables',
        'inventories','trade_payables',
    ]
    fy_re = re.compile(r'^FY(\d{2})$')
    for period in list(data.keys()):
        m = fy_re.match(period)
        if not m:
            continue
        fy2 = int(m.group(1))
        ann = data[period]
        q4  = data.get(f'4QFY{fy2:02d}', {})
        for metric in BS_SNAP:
            if ann.get(metric) is None and q4.get(metric) is not None:
                ann[metric] = q4[metric]
            if q4.get(metric) is None and ann.get(metric) is not None:
                q4[metric] = ann[metric]
        # Recompute debt if missing
        lt = ann.get('lt_borrowings') or 0
        st = ann.get('st_borrowings') or 0
        if lt + st > 0 and not ann.get('debt'):
            ann['debt'] = lt + st


def get_sorted_periods_agent(data: dict) -> list:
    """Return period keys sorted chronologically."""
    fy_re  = re.compile(r'^(\d)QFY(\d{2})$')
    ann_re = re.compile(r'^FY(\d{2})$')

    def sort_key(p):
        m = fy_re.match(p)
        if m:
            return (int(m.group(2)), int(m.group(1)), 0)
        m = ann_re.match(p)
        if m:
            return (int(m.group(1)), 5, 1)
        return (99, 9, 9)

    return sorted(data.keys(), key=sort_key)
