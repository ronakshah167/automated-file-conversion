"""
parser.py  –  Reads the 6 CSV exports from Screener / external source
and returns a unified, period-keyed dict of financial data.

Period key format
-----------------
Quarters : '1QFY16', '2QFY17', …
Annual   : 'FY16', 'FY17', …

CSV ↔ content map
-----------------
export (4).csv  – Quarterly P&L  (Jun-15 → Mar-26)
export (5).csv  – Balance Sheet   (Jun-15, Dec-15, Mar-16 … Mar-26)
export (6).csv  – Cash-flow       (Sep-18 → Mar-26, half-yearly)
export (1).csv  – Detailed BS     (annual, very granular)
export (2).csv  – Annual CF       (detailed annual)
export (3).csv  – Annual CF alt   (annual summary)
"""

import re
import pandas as pd
from pathlib import Path


# ---------------------------------------------------------------------------
# Date-column  →  period-key conversion
# ---------------------------------------------------------------------------

def _parse_col_to_period(col: str):
    """
    Convert a raw date column header to (period_key, period_type).
    e.g.
      'Jun-15'  → ('1QFY16', 'Q')
      'Mar 2016' → ('4QFY16', 'Q')  AND  ('FY16', 'A')
    """
    col_str = str(col).strip()
    # Normalize spaces to hyphens
    col_str = col_str.replace(' ', '-')
    m = re.match(r'([a-zA-Z]{3})-(\d{2,4})$', col_str)
    if not m:
        return None
    month = m.group(1).capitalize()
    yr_str = m.group(2)
    
    if len(yr_str) == 4:
        yr2 = int(yr_str[2:])
    else:
        yr2 = int(yr_str)

    # Fiscal year: Indian FY ends March, so Mar-XX → FY XX; Jun-XX → FY XX+1 etc.
    fy_map = {
        'Apr': yr2 + 1, 'May': yr2 + 1, 'Jun': yr2 + 1,
        'Jul': yr2 + 1, 'Aug': yr2 + 1, 'Sep': yr2 + 1,
        'Oct': yr2 + 1, 'Nov': yr2 + 1, 'Dec': yr2 + 1,
        'Jan': yr2 + 1, 'Feb': yr2 + 1, 'Mar': yr2,
    }
    q_map = {
        'Apr': 1, 'May': 1, 'Jun': 1,
        'Jul': 2, 'Aug': 2, 'Sep': 2,
        'Oct': 3, 'Nov': 3, 'Dec': 3,
        'Jan': 4, 'Feb': 4, 'Mar': 4,
    }
    if month not in fy_map:
        return None
    fy = fy_map[month]
    q  = q_map[month]
    period_key = f'{q}QFY{fy:02d}'
    is_annual  = (month == 'Mar')        # Mar column doubles as annual
    annual_key = f'FY{fy:02d}' if is_annual else None
    return period_key, annual_key


def _safe_float(val):
    try:
        if pd.isna(val):
            return None
        return float(val)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Generic row-reader helpers
# ---------------------------------------------------------------------------

def _read_csv_to_rowdict(path: str) -> dict:
    """Return {row_label: {col_header: value}} for a Screener-style CSV."""
    df = pd.read_csv(path)
    result = {}

    # Find label column
    label_col_idx = 0
    if 'Title' in df.columns:
        label_col_idx = df.columns.get_loc('Title')
    else:
        for i in range(min(3, len(df.columns))):
            non_null_str = [x for x in df.iloc[:, i].dropna() if isinstance(x, str) and str(x).strip() not in ('', 'nan', 'CON', 'SA', 'SA/CON')]
            if len(non_null_str) > 5:
                label_col_idx = i
                break

    for _, row in df.iterrows():
        label = str(row.iloc[label_col_idx]).strip()
        if label in ('', 'nan', 'SA/CON', 'None'):
            continue
        if label not in result:
            result[label] = {
                str(col).strip(): _safe_float(row[col])
                for col in df.columns if col != df.columns[label_col_idx]
            }
    return result


ALIASED_METRICS = {
    # P&L
    'gross_sales': [
        'grosssalesincomefromoperations', 'grosssales', 'revenuefromoperationsgross',
        'revenuegross', 'salesgross', 'grossrevenue'
    ],
    'excise_duty': [
        'lessexciseduty', 'exciseduty', 'excise', 'lessexcise'
    ],
    'net_sales': [
        'netsalesincomefromoperations', 'netsales', 'revenuefromoperationsnet',
        'revenuefromoperations', 'sales', 'revenue', 'netrevenue', 'reportedsales'
    ],
    'other_op_income': [
        'otheroperatingincome', 'otheroperatingrevenues', 'otheroperatingrevenue'
    ],
    'total_income': [
        'totalincomefromoperationsnet', 'totalincome', 'totalrevenue', 'netincome'
    ],
    'raw_material': [
        'rawmaterialconsumed', 'costofmaterialsconsumed', 'rawmaterials', 'materialsconsumed'
    ],
    'stock_adj': [
        'changesininventoriesoffinishedgoodsworkinprogressandstockintrade',
        'changesininventoriesoffinishedgoodswipandstockintrade',
        'changesininventories', 'stockadjustment', 'stockadj', 'increaseinwip'
    ],
    'purchase_fg': [
        'purchaseoffinishedgoods', 'purchasesofstockintrade', 'purchasefg', 'purchaseofstock'
    ],
    'employee_exp': [
        'employeecost', 'employeebenefitsexpense', 'employeebenefitsexpenses',
        'employeeexpenses', 'staffcost', 'salaries', 'personnelexpenses'
    ],
    'other_exp': [
        'otherexpenses', 'administrativeandsellingexpenses', 'adminandsellingexpenses',
        'sellinganddistributionexpenses', 'operatingexpenses', 'otheroperatingexpenses'
    ],
    'total_expenses': [
        'totalexpenses', 'totaloperatingexpenses', 'totalexpenditure'
    ],
    'other_income': [
        'otherincome', 'nonoperatingincome', 'otherincomes', 'othermiscellaneousincome'
    ],
    'interest': [
        'financecosts', 'interest', 'interestexpenses', 'financecost', 'interestcost'
    ],
    'interest2': [
        'otherinterestexpenses'
    ],
    'depreciation': [
        'depreciationamortizationanddepletionexpense',
        'depreciationandamortisationexpense', 'depreciation',
        'depreciationandamortisation', 'depdepreciation'
    ],
    'total_tax': [
        'totaltax', 'taxexpense'
    ],
    'tax': [
        'tax', 'provisionfortax', 'currentyeartax', 'currenttax'
    ],
    'current_tax': [
        'currentyeartax', 'currenttax'
    ],
    'deferred_tax': [
        'deferredtax', 'deferredtaxliability', 'deferredtaxcurrent'
    ],
    'exceptional_items': [
        'exceptionalitems', 'extraordinaryitems', 'exceptionalandextraordinaryitems'
    ],
    'net_profit_mi': [
        'netprofitaftertaxesminorityinterestandshareofprofitlossofassociates',
        'netprofitaftertaxafterminorityinterest'
    ],
    'net_profit_after_tax': [
        'netprofitfromordinaryactivitiesaftertax',
        'netprofitaftertaxfortheperiod', 'profitfortheperiod', 'netprofit', 'pat',
        'netprofitaftertax'
    ],
    'pat_owners': [
        'ownersofthecompany', 'profitattributabletoowners'
    ],
    'minority_interest': [
        'noncontrollinginterest', 'minorityinterest'
    ],
    'eps_basic': [
        'epsbeforeexceptionalextraordinaryitemsbasic',
        'epsbeforeexceptionalextraordinaryitemsdiluted',
        'basiceps'
    ],
    'eps_diluted': [
        'epsafterexceptionalextraordinaryitemsbasic',
        'epsafterexceptionalextraordinaryitemsdiluted',
        'dilutedeps'
    ],
    'face_value': [
        'facevalue', 'facevaluers'
    ],
    'equity_shares': [
        'equityshares', 'equitysharecapital', 'sharecapital'
    ],
    'dps': [
        'dividendpersharers', 'dividendpershagers', 'dividendpershare', 'dps'
    ],
    'dividend_pct': [
        'dividend'
    ],
    'pbt': [
        'profitfromordinaryactivitiesbeforetax', 'profitbeforetax', 'pbt', 'profitbeforeexceptionalitemsandtax'
    ],
    'exceptional_income': [
        'totalexceptionalincome'
    ],
    'exceptional_expense': [
        'totalexceptionalexpenses'
    ],
    'power_fuel': [
        'poweroilfuel', 'powerfuel', 'powerandfuel'
    ],
    'ebit_pre_other': [
        'profitfromoperationsbeforeotherincomefinancecostsandexceptionalitems'
    ],
    
    # Balance Sheet
    'equity': [
        'sharecapital', 'equitysharecapital', 'equity'
    ],
    'reserves': [
        'reservessurplus',       # 'Reserves & Surplus'  → & stripped → reservessurplus
        'reservesandsurplus',    # some exports use 'Reserves and Surplus'
        'reserves', 'otherequity', 'othershareholdersequity',
        'surplusintheprofitandlossaccount'
    ],
    'shareholders_fund': [
        'shareholdersfund', 'totalshareholdersfunds'
    ],
    'lt_borrowings': [
        'longtermborrowings', 'noncurrentborrowings', 'longtermdebt'
    ],
    'st_borrowings': [
        'shorttermborrowings', 'currentborrowings', 'shorttermdebt',
        'workingcapitalborrowing'
    ],
    'deferred_tax_liab': [
        'deferredtaxliabilitiesnet', 'deferredtaxliabilities'
    ],
    'net_block': [
        'fixedassetsinclcapitalworkinprogress', 'tangibleassets', 'netblock',
        'propertyplantandequipment', 'fixedassets', 'noncurrentassets',
        'propertyplantequipment'
    ],
    'cash': [
        'cashandcashequivalents', 'cashcashequivalents', 'cash', 'cashandbankbalances'
    ],
    'inventories': [
        'inventories'
    ],
    'trade_receivables': [
        'tradereceivables', 'sundrydebtors'
    ],
    'total_current_assets': [
        'totalcurrentassets'
    ],
    'total_current_liabilities': [
        'totalcurrentliabilities'
    ],
    'trade_payables': [
        'tradepayables', 'sundrycreditors'
    ],
    'total_assets': [
        'totalassets', 'totalequityliabilities'
    ],
    'lt_provisions': [
        'longtermprovisions'
    ],
    
    # Cash Flow
    'ocf': [
        'netcashfromoperatingactivities',
        'netcashgeneratedfromusedinoperations',
        'cashflowfromoperatingactivities'
    ],
    'capex': [
        'purchasedoffixedassets', 'purchaseoffixedassets', 'capitalexpenditure', 'purchasesoffixedassets'
    ],
    'dividend': [
        'dividendpaid'
    ],
    'interest_paid': [
        'interestpaid'
    ],
    'investing_cf': [
        'netcashusedininvestingactivities', 'cashflowfrominvestingactivities'
    ],
    'financing_cf': [
        'netcashusedinfinancingactivities', 'cashflowfromfinancingactivities'
    ],
    'net_change_cash': [
        'netincdecincashandcashequivalent', 'netincreasedecreaseincashandcashequivalent'
    ]
}


def _map_raw_to_internal_keys(raw: dict) -> dict:
    """
    Given the raw row dict from a file, maps the raw labels to internal keys.
    Returns a dict of { internal_key: { col: val } }
    """
    mapped = {}
    normalized_raw = {}
    for raw_label in raw.keys():
        norm = re.sub(r'[^a-zA-Z0-9]', '', raw_label).lower()
        normalized_raw[norm] = raw_label

    for int_key, aliases in ALIASED_METRICS.items():
        # Match exact normalized alias only (no substring matching to prevent wrong row bindings)
        for alias in aliases:
            if alias in normalized_raw:
                mapped[int_key] = raw[normalized_raw[alias]]
                break
    return mapped


def detect_csv_type(file_path: str) -> str:
    """
    Identify ARQ statement type from CSV content fingerprints.
    Matches all 6 ARQ export types precisely.

    Fingerprints (from ARQ_to_VV_Mapping doc):
      export_1 – Annual P&L (full detail)  : has 'Title' col; row1 label = 'Revenue From Operations(Gross)'
      export_2 – Annual BS (full detail)   : has 'Title' col; row1 = 'Share Capital'; 500+ rows
      export_3 – Annual Cash Flow          : has 'Title' col; row1 = 'Net Profit'; row2 = 'Total Adjustments'
      export_4 – Quarterly P&L             : row1 = 'Gross Sales/Income from operations'; 40+ data cols
      export_5 – Semi-annual BS (summary)  : row1 = 'Shareholders Fund'; ~30-33 rows
      export_6 – Semi-annual Cash Flow     : row1 = 'Net Profit before Tax & Extraordinary Items'
    """
    try:
        df = pd.read_csv(file_path)
        nrows = len(df)
        ncols = len(df.columns)

        # Detect label column
        title_col = None
        if 'Title' in df.columns:
            title_col = 'Title'
        else:
            for i in range(min(3, ncols)):
                non_null = [x for x in df.iloc[:, i].dropna()
                            if isinstance(x, str) and str(x).strip() not in ('', 'nan', 'CON', 'SA', 'SA/CON')]
                if len(non_null) > 5:
                    title_col = df.columns[i]
                    break

        if title_col is None:
            return 'CSV File'

        # Get first few labels
        labels = [str(df[title_col].iloc[i]).strip() for i in range(min(5, nrows))
                  if str(df[title_col].iloc[i]).strip() not in ('nan', '', 'SA/CON', 'CON', 'SA')]

        if not labels:
            return 'CSV File'

        first_label = labels[0]
        first_label_norm = re.sub(r'[^a-zA-Z0-9]', '', first_label).lower()

        # export_4 – Quarterly P&L: many columns (40+), starts with Gross Sales
        if 'grosssalesincomefromoperations' in first_label_norm and ncols >= 35:
            return 'Quarterly P&L'

        # export_1 – Annual P&L: has 'Revenue From Operations(Gross)' as first label
        if 'revenuefromoperations' in first_label_norm:
            return 'Annual P&L'

        # export_3 – Annual Cash Flow: starts with 'Net Profit' then 'Total Adjustments'
        if first_label_norm == 'netprofit' and len(labels) > 1:
            if 'totaladjustments' in re.sub(r'[^a-zA-Z0-9]', '', labels[1]).lower():
                return 'Annual Cash Flow'

        # export_6 – Semi-annual Cash Flow: starts with 'Net Profit before Tax & Extraordinary Items'
        if 'netprofitbeforetax' in first_label_norm and 'extraordinary' in first_label_norm.lower():
            return 'Cash Flow (Half-Yearly)'

        # Distinguish export_2 (Annual BS, 500+ rows) vs export_5 (Semi-annual BS, ~30 rows)
        if 'sharecapital' in first_label_norm:
            if nrows > 300:
                return 'Annual Balance Sheet'
            else:
                return 'Balance Sheet (Half-Yearly)'

        if 'shareholdersfund' in first_label_norm:
            return 'Balance Sheet (Half-Yearly)'

        # Fallback: use label set detection
        all_labels = set()
        for label in df[title_col].dropna():
            norm = re.sub(r'[^a-zA-Z0-9]', '', str(label)).lower()
            all_labels.add(norm)

        if 'grosssalesincomefromoperations' in all_labels:
            if ncols >= 35:
                return 'Quarterly P&L'
        if 'revenuefromoperationsgross' in all_labels:
            return 'Annual P&L'
        if 'shareholdersfund' in all_labels or 'longtermborrowings' in all_labels:
            if nrows > 300:
                return 'Annual Balance Sheet'
            return 'Balance Sheet (Half-Yearly)'
        if 'netcashgeneratedfromusedinoperations' in all_labels:
            return 'Annual Cash Flow'
        if 'netcashfromoperatingactivities' in all_labels or 'purchasedoffixedassets' in all_labels:
            return 'Cash Flow (Half-Yearly)'

        return 'Financial Statement'
    except Exception:
        return 'CSV File'



# ---------------------------------------------------------------------------
# Per-CSV parsers
# ---------------------------------------------------------------------------

def _parse_quarterly_pl(path: str) -> dict:
    """
    export (4).csv  –  Quarterly P&L
    Returns { period_key: { metric: value } }

    ARQ export_4 Hierarchy Notes:
      idx 12 : Power, Oil & Fuel  (separate from Other Expenses — must be ADDED to other_exp)
      idx 18 : Other Expenses (FIRST, real occurrence)
      idx 19 : Other Expenses (DUPLICATE ECHO — ignore)
      idx 33 : Other Expenses (under Exceptional section — always 0, ignore)
    VV Row 21 "Other Expenses" = idx18 + idx12 (Power/Fuel)
    """
    df = pd.read_csv(path)
    raw = _read_csv_to_rowdict(path)
    mapped = _map_raw_to_internal_keys(raw)

    # Build row-position index for safe duplicate-label access
    # Use the same label column detection as _read_csv_to_rowdict
    label_col_idx = 0
    if 'Title' in df.columns:
        label_col_idx = df.columns.get_loc('Title')
    else:
        for i in range(min(3, len(df.columns))):
            non_null_str = [x for x in df.iloc[:, i].dropna()
                            if isinstance(x, str) and str(x).strip() not in ('', 'nan', 'CON', 'SA', 'SA/CON')]
            if len(non_null_str) > 5:
                label_col_idx = i
                break

    # Extract Power/Fuel and Other Expenses by fixed row position (0-indexed)
    # These are always at the same position in ARQ export_4 regardless of company
    EXPORT4_OTHER_EXP_IDX  = 18   # "Other Expenses" (first real occurrence)
    EXPORT4_POWER_FUEL_IDX = 12   # "Power, Oil & Fuel"

    def _get_row_by_idx(row_idx, col):
        try:
            return _safe_float(df.iloc[row_idx][col])
        except Exception:
            return None

    BS_METRICS = {
        'face_value', 'equity_shares', 'eps_basic', 'eps_diluted', 'dps', 'dividend_pct'
    }

    out = {}
    annual_snap = {}

    all_cols = list(next(iter(raw.values())).keys()) if raw else []
    for col in all_cols:
        parsed = _parse_col_to_period(col)
        if parsed is None:
            continue
        q_key, a_key = parsed
        if q_key not in out:
            out[q_key] = {}

        for int_key, col_vals in mapped.items():
            val = col_vals.get(col)
            if val is not None:
                out[q_key][int_key] = val

        # Override other_exp and power_fuel with row-position-based reads
        # to avoid the duplicate label problem
        other_exp  = _get_row_by_idx(EXPORT4_OTHER_EXP_IDX, col)
        power_fuel = _get_row_by_idx(EXPORT4_POWER_FUEL_IDX, col)
        if other_exp is not None:
            out[q_key]['other_exp']  = other_exp
        if power_fuel is not None:
            out[q_key]['power_fuel'] = power_fuel

        if a_key:
            if a_key not in annual_snap:
                annual_snap[a_key] = {}
            for int_key in BS_METRICS:
                if int_key in mapped:
                    val = mapped[int_key].get(col)
                    if val is not None:
                        annual_snap[a_key][int_key] = val

    for a_key, vals in annual_snap.items():
        if a_key not in out:
            out[a_key] = {}
        out[a_key].update(vals)
    return out



def _parse_balance_sheet(path: str) -> dict:
    """
    export (5).csv  –  Balance Sheet (bi-annual/half-yearly)
    Returns { period_key: { bs_metric: value } }
    """
    raw = _read_csv_to_rowdict(path)
    mapped = _map_raw_to_internal_keys(raw)
    
    out = {}
    all_cols = list(next(iter(raw.values())).keys()) if raw else []
    for col in all_cols:
        parsed = _parse_col_to_period(col)
        if parsed is None:
            continue
        q_key, a_key = parsed
        if q_key not in out:
            out[q_key] = {}
        for int_key, col_vals in mapped.items():
            val = col_vals.get(col)
            if val is not None:
                out[q_key][int_key] = val
        if a_key:
            if a_key not in out:
                out[a_key] = {}
            for int_key, col_vals in mapped.items():
                val = col_vals.get(col)
                if val is not None:
                    out[a_key][int_key] = val
    return out


def _parse_cashflow_biannual(path: str) -> dict:
    """
    export (6).csv  –  Cash Flow (bi-annual)
    Returns { period_key: { cf_metric: value } }
    """
    raw = _read_csv_to_rowdict(path)
    mapped = _map_raw_to_internal_keys(raw)
    
    CF_METRICS = {
        'ocf', 'capex', 'dividend', 'interest_paid', 'investing_cf',
        'financing_cf', 'net_change_cash'
    }

    out = {}
    all_cols = list(next(iter(raw.values())).keys()) if raw else []
    for col in all_cols:
        parsed = _parse_col_to_period(col)
        if parsed is None:
            continue
        q_key, a_key = parsed
        if q_key not in out:
            out[q_key] = {}
        for int_key in CF_METRICS:
            if int_key in mapped:
                val = mapped[int_key].get(col)
                if val is not None:
                    out[q_key][int_key] = val
        if a_key:
            if a_key not in out:
                out[a_key] = {}
            for int_key in CF_METRICS:
                if int_key in mapped:
                    val = mapped[int_key].get(col)
                    if val is not None:
                        out[a_key][int_key] = val
    return out


def _parse_annual_bs(path: str) -> dict:
    """
    export (1).csv  –  Annual detailed Balance Sheet
    Returns { 'FYxx': { bs_metric: value } }
    """
    raw = _read_csv_to_rowdict(path)
    mapped = _map_raw_to_internal_keys(raw)
    
    out = {}
    all_cols = list(next(iter(raw.values())).keys()) if raw else []
    for col in all_cols:
        parsed = _parse_col_to_period(col)
        if parsed is None:
            continue
        _, a_key = parsed
        if a_key:
            if a_key not in out:
                out[a_key] = {}
            for int_key, col_vals in mapped.items():
                val = col_vals.get(col)
                if val is not None:
                    out[a_key][int_key] = val
    return out


def _parse_annual_cf(path: str) -> dict:
    """
    export (2) or (3).csv  –  Annual Cash Flow
    Returns { 'FYxx': { cf_metric: value } }
    """
    raw = _read_csv_to_rowdict(path)
    mapped = _map_raw_to_internal_keys(raw)
    
    out = {}
    all_cols = list(next(iter(raw.values())).keys()) if raw else []
    for col in all_cols:
        parsed = _parse_col_to_period(col)
        if parsed is None:
            continue
        _, a_key = parsed
        if a_key:
            if a_key not in out:
                out[a_key] = {}
            for int_key, col_vals in mapped.items():
                val = col_vals.get(col)
                if val is not None:
                    out[a_key][int_key] = val
    return out


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def parse_all(csv_dir: str) -> dict:
    """
    Parse all CSV files inside csv_dir dynamically.
    Implements DOUBLE-PASS processing:
      Pass 1: Load all 6 ARQ files in correct priority order.
      Pass 2: Re-parse with relaxed matching to fill any gaps left by Pass 1.
    Returns a merged dict { period_key: { metric_key: float_value } }
    """
    csv_dir = Path(csv_dir)

    # ---- categorise files by type ----
    categorised = {
        'quarterly_pl':    [],
        'annual_pl':       [],
        'annual_bs':       [],
        'annual_cf':       [],
        'half_yearly_bs':  [],
        'half_yearly_cf':  [],
        'unknown':         [],
    }
    for csv_file in sorted(csv_dir.glob('*.csv')):
        f_path = str(csv_file)
        st_type = detect_csv_type(f_path)
        if 'Quarterly P&L' in st_type:
            categorised['quarterly_pl'].append(f_path)
        elif 'Annual P&L' in st_type:
            categorised['annual_pl'].append(f_path)
        elif st_type == 'Annual Balance Sheet':
            categorised['annual_bs'].append(f_path)
        elif st_type == 'Annual Cash Flow':
            categorised['annual_cf'].append(f_path)
        elif st_type == 'Balance Sheet (Half-Yearly)':
            categorised['half_yearly_bs'].append(f_path)
        elif 'Cash Flow (Half-Yearly)' in st_type:
            categorised['half_yearly_cf'].append(f_path)
        else:
            categorised['unknown'].append(f_path)

    data = {}  # { period: { metric: val } }

    def merge(source: dict, overwrite: bool = False):
        """Merge source into data. If overwrite=True, newer values replace existing."""
        for period, metrics in source.items():
            if period not in data:
                data[period] = {}
            for k, v in metrics.items():
                if v is not None and (overwrite or k not in data[period] or data[period][k] is None):
                    data[period][k] = v

    # ========================================================================
    # PASS 1: Load all files in data-quality priority order
    # Priority (highest first): annual_pl > annual_cf > half_yearly_bs >
    #                            half_yearly_cf > quarterly_pl
    # ========================================================================
    for f in categorised['half_yearly_bs']:
        merge(_parse_balance_sheet(f))
    for f in categorised['half_yearly_cf']:
        merge(_parse_cashflow_biannual(f))
    for f in categorised['quarterly_pl']:
        merge(_parse_quarterly_pl(f))
    for f in categorised['annual_cf']:
        merge(_parse_annual_cf(f))
    for f in categorised['annual_pl']:
        merge(_parse_annual_bs(f))  # _parse_annual_bs handles annual P&L rows too

    # ---- After Pass 1: Compute derived metrics ----
    _compute_derived(data)
    _compute_annual_from_quarters(data)
    _propagate_bs_to_annual(data)

    # ========================================================================
    # PASS 2: Re-merge all sources again — second iteration fills any gaps
    # that the first pass missed (e.g. metrics skipped due to dict ordering).
    # This pass uses overwrite=False so authoritative data is preserved.
    # ========================================================================
    for f in categorised['half_yearly_bs']:
        merge(_parse_balance_sheet(f), overwrite=False)
    for f in categorised['half_yearly_cf']:
        merge(_parse_cashflow_biannual(f), overwrite=False)
    for f in categorised['quarterly_pl']:
        merge(_parse_quarterly_pl(f), overwrite=False)
    for f in categorised['annual_cf']:
        merge(_parse_annual_cf(f), overwrite=False)
    for f in categorised['annual_pl']:
        merge(_parse_annual_bs(f), overwrite=False)
    for f in categorised['unknown']:
        # Unknown files: try all parsers, keep any data found
        for parser_fn in [_parse_quarterly_pl, _parse_balance_sheet, _parse_cashflow_biannual]:
            try:
                merge(parser_fn(f), overwrite=False)
            except Exception:
                pass

    # ---- After Pass 2: Recompute everything for complete accuracy ----
    _compute_derived(data)
    _compute_annual_from_quarters(data)
    _propagate_bs_to_annual(data)

    return data


def _propagate_bs_to_annual(data: dict):
    """
    For each annual period FYxx, ensure BS snapshot metrics are populated
    from the 4QFYxx (Mar year-end) quarter — since the half-yearly BS only
    has Mar and Sep snapshots. The Mar-end of a FY is the same as 4QFYxx.

    Also: carry forward equity/face_value from adjacent periods when missing.
    """
    BS_SNAPSHOT = [
        'equity', 'reserves', 'lt_borrowings', 'st_borrowings', 'debt',
        'cash', 'net_block', 'face_value', 'num_shares',
        'total_current_assets', 'total_current_liabilities', 'working_capital',
        'shareholders_fund',
    ]
    fy_re = re.compile(r'^FY(\d{2})$')

    for period in list(data.keys()):
        m = fy_re.match(period)
        if not m:
            continue
        fy2 = int(m.group(1))
        annual_data = data[period]
        q4_key = f'4QFY{fy2:02d}'

        if q4_key in data:
            q4_data = data[q4_key]
            for metric in BS_SNAPSHOT:
                # Fill annual from 4Q if missing in annual
                if (metric not in annual_data or annual_data[metric] is None) and q4_data.get(metric) is not None:
                    annual_data[metric] = q4_data[metric]
                # Also fill 4Q from annual if 4Q is missing
                if (metric not in q4_data or q4_data[metric] is None) and annual_data.get(metric) is not None:
                    q4_data[metric] = annual_data[metric]

        # Propagate debt = lt_borrowings + st_borrowings if debt missing
        lt = annual_data.get('lt_borrowings') or 0
        st = annual_data.get('st_borrowings') or 0
        if lt + st > 0 and (annual_data.get('debt') is None or annual_data.get('debt') == 0):
            annual_data['debt'] = lt + st

        # Working Capital = TCA - TCL
        tca = annual_data.get('total_current_assets') or 0
        tcl = annual_data.get('total_current_liabilities') or 0
        if tca > 0:
            annual_data['working_capital'] = tca - tcl




def _compute_annual_from_quarters(data: dict):
    """
    For each annual period FYxx, fill in MISSING P&L flow metrics by summing quarterly values
    from 1QFYxx + 2QFYxx + 3QFYxx + 4QFYxx.

    Priority: export_1 (audited annual) data takes precedence. Quarterly sums only fill gaps.
    This matches the VV template design where annual columns use the audited annual filing.
    """
    FLOW_METRICS = [
        'gross_sales', 'excise_duty', 'net_sales', 'other_op_income', 'total_income',
        'raw_material', 'stock_adj', 'purchase_fg', 'employee_exp', 'other_exp', 'total_expenditure', 'cogs',
        'other_exp2', 'other_income', 'interest', 'interest2',
        'depreciation', 'total_tax', 'tax', 'current_tax', 'deferred_tax',
        'net_profit_mi', 'net_profit_after_tax', 'pat_owners', 'exceptional_items',
        'pbt', 'ebit_pre_other', 'minority_interest', 'ocf', 'capex', 'dividend',
    ]
    # Identify all annual FY keys
    fy_re = re.compile(r'^FY(\d{2})$')
    for period in list(data.keys()):
        m = fy_re.match(period)
        if not m:
            continue
        fy2 = int(m.group(1))
        annual_data = data[period]
        for metric in FLOW_METRICS:
            # Only fill if annual value is genuinely missing (export_1/3 data takes priority)
            if metric in annual_data and annual_data[metric] is not None:
                continue  # Already have authoritative annual data — don't overwrite
            # Fill from quarterly sum if all 4 quarters present
            q_vals = [data[f'{q}QFY{fy2:02d}'].get(metric) for q in range(1, 5) if f'{q}QFY{fy2:02d}' in data]
            valid_q_vals = [v for v in q_vals if v is not None]
            if len(valid_q_vals) > 0:
                annual_data[metric] = sum(valid_q_vals)



def _get_prev_period(current: str, all_data: dict):
    """Find the same-type period from 1 year earlier."""
    qm = re.match(r'^(\d)QFY(\d{2})$', current)
    am = re.match(r'^FY(\d{2})$', current)
    if qm:
        q, fy = int(qm.group(1)), int(qm.group(2))
        prev = f'{q}QFY{fy-1:02d}'
        return prev if prev in all_data else None
    if am:
        fy = int(am.group(1))
        prev = f'FY{fy-1:02d}'
        return prev if prev in all_data else None
    return None


def _compute_derived(data: dict):
    """
    Compute derived / formula-based metrics for each period.
    Mirrors the Excel formulas.
    """
    # Pass 1: Compute static metrics
    for period, d in data.items():
        # Net Sales = Gross Sales - Excise Duty
        if 'net_sales' not in d or d['net_sales'] is None:
            if 'gross_sales' in d and 'excise_duty' in d:
                gs = d['gross_sales'] or 0
                ex = d['excise_duty'] or 0
                d['net_sales'] = gs - ex

        # Total Income = Net Sales + Other Operating Income
        if 'total_income' not in d or d['total_income'] is None:
            ns = d.get('net_sales') or 0
            oi = d.get('other_op_income') or 0
            d['total_income'] = ns + oi

        # Total Expenditure / Total Expenses
        tot_exp = d.get('total_expenditure') or d.get('total_expenses')
        dep     = d.get('depreciation') or 0
        rm      = d.get('raw_material') or 0
        stk     = d.get('stock_adj') or 0
        pfg     = d.get('purchase_fg') or 0
        emp     = d.get('employee_exp') or 0

        # Other Expenses computation — CRITICAL difference between quarterly and annual sources:
        # - Quarterly (export_4): other_exp (r18) does NOT include power_fuel (r12) — must ADD them
        # - Annual (export_1): other_exp (r78) ALREADY includes power_fuel inside Manufacturing expenses — do NOT add
        parsed_other = d.get('other_exp') or 0.0
        power_fuel   = d.get('power_fuel') or 0.0
        is_annual = bool(re.match(r'^FY\d{2}$', period))
        if not is_annual and (parsed_other > 0.0 or power_fuel > 0.0):
            # Quarterly: add power_fuel to other_exp
            d['other_exp'] = parsed_other + power_fuel
        elif is_annual and parsed_other > 0.0:
            # Annual: other_exp from export_1 r78 is already the full aggregate — use as-is
            d['other_exp'] = parsed_other
        elif tot_exp:
            # Fallback: derive from Total Expenses if both are missing
            d['other_exp'] = max(0.0, tot_exp - dep - (rm + stk + pfg + emp))


        # COGS = raw_material + stock_adj + purchase_fg
        cogs = rm + stk + pfg
        d['cogs'] = cogs

        # Total Expenditure = COGS + employee + other_exp
        if not tot_exp:
            tot_exp = cogs + emp + (d.get('other_exp') or 0)
        d['total_expenditure'] = tot_exp

        # EBITDA = Total Income - Total Expenditure
        ti = d.get('total_income') or 0
        ebitda = ti - tot_exp
        if ebitda != 0:
            d.setdefault('ebitda', ebitda)

        # PBDT = EBITDA - Interest + Other Income
        ebitda_val = d.get('ebitda') or 0
        interest   = d.get('interest') or d.get('interest2') or 0
        other_inc  = d.get('other_income') or 0
        d['pbdt'] = ebitda_val - interest + other_inc

        # PBT = PBDT - Depreciation
        dep = d.get('depreciation') or 0
        pbt = d.get('pbt') or (d['pbdt'] - dep)
        d['pbt'] = pbt

        # Tax normalization — ensure tax_current and deferred_tax are set
        tax_current  = d.get('tax_current') or d.get('current_tax') or d.get('tax') or 0
        deferred_tax = d.get('deferred_tax') or 0
        if tax_current:
            d['tax_current'] = tax_current
        if deferred_tax:
            d['deferred_tax'] = deferred_tax

        # total_tax = current + deferred (for formula reference)
        total_tax = d.get('total_tax') or (tax_current + deferred_tax)
        d['total_tax'] = total_tax

        # PAT = PBT - total_tax
        tax = total_tax
        pat = d.get('net_profit_after_tax') or d.get('pat_owners') or (pbt - tax)
        d['pat'] = pat


        # PBIT = EBITDA - Depreciation
        d['pbit'] = ebitda_val - dep

        # EBITDA Margin
        if ti and ti != 0:
            d['ebitda_margin'] = ebitda_val / ti
        # EBIT Margin
        if ti and ti != 0:
            d['ebit_margin'] = d['pbit'] / ti

        # D:E ratio
        eq  = d.get('equity') or 0
        res = d.get('reserves') or 0
        lt  = d.get('lt_borrowings') or 0
        st  = d.get('st_borrowings') or 0
        debt = lt + st
        d['debt'] = debt
        networth = eq + res
        if networth and networth != 0:
            d['de_ratio'] = debt / networth
            d['roe']      = pat / networth if pat else None
            d['book_value'] = networth  # before dividing by shares

        # ROCE = PBIT / (Equity + Reserves + Debt)
        cap_emp = networth + debt
        if cap_emp and cap_emp != 0:
            d['roce'] = d['pbit'] / cap_emp if d['pbit'] else None

        # Working capital = Current Assets - Current Liabilities
        tca = d.get('total_current_assets') or 0
        tcl = d.get('total_current_liabilities') or 0
        d['working_capital'] = tca - tcl

        # EPS (Rs.)
        fv = d.get('face_value') or 2.0
        if pat is not None and eq != 0:
            shares_cr = eq / fv
            d['eps'] = pat / shares_cr if shares_cr != 0 else None
        else:
            d['eps'] = None

        # Book Value per share
        if eq != 0:
            shares_cr = eq / fv
            d['book_value_ps'] = (eq + res) / shares_cr if shares_cr != 0 else None
        else:
            d['book_value_ps'] = None

        # Number of shares
        d['num_shares'] = eq / fv if eq else None

        # NOPAT
        tax_rate = tax / pbt if pbt and pbt != 0 else 0.0
        d['nopat'] = d['pbit'] * (1 - tax_rate)

        # Invested Capital
        d['cap_employed'] = cap_emp
        d['invested_cap'] = cap_emp - (d.get('cash') or 0.0)

        # ROIC
        ic = d['invested_cap']
        d['roic'] = d['nopat'] / ic if ic and ic != 0 else None

    # Pass 2: Compute YOY growth metrics
    GROWTH_MAP = {
        'gs_growth':    'gross_sales',
        'ns_growth':    'net_sales',
        'ti_growth':    'total_income',
        'ebitda_growth':'ebitda',
        'pbit_growth':  'pbit',
        'pbt_growth':   'pbt',
        'pat_growth':   'pat',
    }
    for period, d in data.items():
        prev_period = _get_prev_period(period, data)
        for growth_key, base_metric in GROWTH_MAP.items():
            curr_val = d.get(base_metric)
            if prev_period and curr_val is not None:
                prev_val = data[prev_period].get(base_metric)
                if prev_val and prev_val != 0:
                    d[growth_key] = (curr_val - prev_val) / abs(prev_val)
                else:
                    d[growth_key] = None
            else:
                d[growth_key] = None



def get_sorted_periods(data: dict):
    """Return all period keys sorted chronologically (Q before A)."""
    fy_re  = re.compile(r'^(\d)QFY(\d{2})$')
    ann_re = re.compile(r'^FY(\d{2})$')

    def sort_key(p):
        m = fy_re.match(p)
        if m:
            return (int(m.group(2)), int(m.group(1)), 0)
        m = ann_re.match(p)
        if m:
            return (int(m.group(1)), 5, 1)  # annual after Q4
        return (99, 9, 9)

    return sorted(data.keys(), key=sort_key)


# ---------------------------------------------------------------------------
# CLI test
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    import json, sys
    csv_dir = sys.argv[1] if len(sys.argv) > 1 else '.'
    data = parse_all(csv_dir)
    periods = get_sorted_periods(data)
    print(f'Parsed {len(data)} periods: {periods[:10]} …')
    # Print sample
    for p in periods[:5]:
        print(f'\n{p}:', json.dumps(data[p], indent=2, default=str))
