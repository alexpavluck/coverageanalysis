"""
CSAT v10 analysis pipeline — ESPEN data format.

Rewritten from v2.7 to consume ESPEN Coverage Evaluation Survey response data
(demo_ces_9999_2_couverture_RESPONSE_DATA.xlsx / sheet 'data').

Column schema: all ESPEN p_* columns — see COL_* constants below.

⚠ Unmapped elements (no ESPEN equivalent):
  • admin3 (sub-district level) — ESPEN hierarchy is region → district → site only
  • cddvisit_ov           — ESPEN does not record whether CDD visited household
  • absent_ov             — ESPEN does not record respondent absence during visit
  Cascade simplified from 6 steps to 3 (Total → Received → Swallowed).
  Sub-district section replaced by MDA Delivery Location (p_mda_location).

Public entry point:
    generate_report(xlsx_bytes: bytes, user_config: dict) -> str
"""

import io, os, re, base64, warnings
from datetime import datetime
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D

plt.rcParams.update({
    'figure.dpi': 120,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.grid': True,
    'grid.alpha': 0.3,
    'font.family': 'DejaVu Sans',
})

NAVY   = '#0b2a5e'
BLUE   = '#1a4dbd'
ORANGE = '#FF9F33'
GREEN  = '#0DCF00'
YELLOW = '#F7B500'
RED    = '#FF4444'
GREY   = '#D4D6D8'

def fig_to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)
    return b64

def scorecard_metric(label, value, status='neutral'):
    colors = {'good':'#0DCF00','low':'#FF4444','warn':'#F7B500','neutral':'#e8ecf0'}
    bg = colors.get(status, colors['neutral'])
    txt_color = 'white' if status == 'low' else '#1a1a1a'
    return f'''<div class="scorecard-item" style="background:{bg};color:{txt_color}">
  <div class="sc-label">{label}</div>
  <div class="sc-value">{value}</div>
</div>'''

def cluster_prop_logit_ci(df_sub, outcome_col, cluster_col='cluster',
                          level=0.95, one_sided_lower=False, drop_na=True):
    sub = df_sub.copy()
    if drop_na:
        sub = sub[sub[outcome_col].notna()]
    y = sub[outcome_col].astype(float).values
    c = sub[cluster_col].astype(str).values
    n = len(y)
    if n == 0:
        return dict(phat=np.nan, lower=np.nan, upper=np.nan, se=np.nan,
                    deff=np.nan, n=0, n_clusters=0)
    phat = y.mean()
    clusters = np.unique(c); C = len(clusters)
    if C <= 1:
        se = np.sqrt(phat*(1-phat)/n) if 0 < phat < 1 else 0.0; deff = 1.0
    else:
        e_c = np.array([y[c==ci].sum() - y[c==ci].size*phat for ci in clusters])
        var_phat = (C/(C-1)) * np.sum(e_c**2) / n**2
        se = np.sqrt(var_phat)
        srs_var = phat*(1-phat)/n if n>0 else np.nan
        deff = round(var_phat/srs_var, 1) if srs_var and srs_var > 0 else np.nan
    z = stats.norm.ppf(level) if one_sided_lower else stats.norm.ppf((1+level)/2)
    if phat <= 0 or phat >= 1 or se == 0:
        lower = max(0.0, phat - z*se); upper = min(1.0, phat + z*se)
    else:
        logit_p = np.log(phat/(1-phat)); se_logit = se/(phat*(1-phat))
        lower = np.exp(logit_p - z*se_logit) / (1+np.exp(logit_p - z*se_logit))
        upper = np.exp(logit_p + z*se_logit) / (1+np.exp(logit_p + z*se_logit))
    if one_sided_lower:
        upper = 1.0
    return dict(phat=round(phat,4), lower=round(lower,4), upper=round(upper,4),
                se=round(se,4), deff=deff, n=n, n_clusters=C)

def wilson_ci(x, n, conf=0.95):
    if n == 0: return np.nan, np.nan
    z = stats.norm.ppf((1+conf)/2); phat = x/n; denom = 1 + z**2/n
    center = (phat + z**2/(2*n)) / denom
    margin = z * np.sqrt(phat*(1-phat)/n + z**2/(4*n**2)) / denom
    return max(0, center-margin)*100, min(1, center+margin)*100

def svychisq_sex(df_sub, outcome_col):
    from scipy.stats import chi2_contingency
    try:
        sub = df_sub[df_sub[outcome_col].notna()]
        ct = pd.crosstab(sub['sex'], sub[outcome_col])
        if ct.shape[0] < 2 or ct.shape[1] < 2: return np.nan
        _, p, _, _ = chi2_contingency(ct)
        return p
    except Exception:
        return np.nan

def parse_multiselect(series, separator=r'\s+'):
    vals = []
    for v in series.dropna().astype(str):
        s = v.strip()
        if s in ('', 'nan', '---', 'None'): continue
        for tok in re.split(separator, s):
            tok = tok.strip()
            if tok and tok not in ('---','nan'):
                vals.append(tok)
    return pd.Series(vals).value_counts()

def relabel(counts, codebook):
    out = {}
    for code, n in counts.items():
        out[codebook.get(str(code), str(code))] = n
    return pd.Series(out).sort_values(ascending=False)

print("✅  Libraries imported and helpers defined.")


def generate_report(xlsx_bytes, user_config=None):
    """Run the ESPEN-native v10 pipeline against xlsx_bytes and return HTML."""
    user_config = dict(user_config or {})

    # ── Config ──────────────────────────────────────────────────────────────────
    SHEET_NAME   = user_config.get("sheet_name", "data")   # ESPEN sheet name
    COUNTRY      = user_config.get("country", "Country")
    DISEASE      = user_config.get("disease", "Disease")
    DRUG         = user_config.get("drug", "Ivermectin (IVM)")
    DRUG_CODE    = user_config.get("drug_code", "ivm")     # ivm | alb | pzq
    MDA_ROUND    = user_config.get("mda_round", 1)
    REPORT_DATE  = user_config.get("report_date", datetime.today().strftime("%B %Y"))
    REPORTED_COVERAGE = user_config.get("reported_coverage", None)
    THRESH_EPI   = int(user_config.get("thresh_epi",   65))
    THRESH_THERA = int(user_config.get("thresh_thera", 80))

    # ── ESPEN column mappings ────────────────────────────────────────────────────
    _d = DRUG_CODE.lower()   # drug suffix for column name construction

    COL_REGION       = "p_admin1"
    COL_DISTRICT     = "p_district"
    COL_SITE         = "p_site"           # cluster / village equivalent
    COL_HH           = "p_hh_number"
    COL_SEX          = "p_sex"
    COL_AGE          = "p_age"
    COL_RESPONDENT   = "p_respondent"
    COL_OFFERED      = f"p_received_{_d}"          # e.g. p_received_ivm
    COL_SWALLOW      = f"p_swalllowed_{_d}"        # note: ESPEN typo — 3 l's
    COL_N_SWALLOWED  = f"p_number_{_d}_swallowed"
    COL_REASONS_NSW  = f"p_reason_not_swallowed_{_d}"
    COL_REASONS_SW   = f"p_reason_swallowed_{_d}"
    COL_NT_TREAT     = f"p_nt_swallowed_{_d}"
    COL_NT_WHEN      = f"p_nt_when_swallowed_{_d}"
    COL_SIDE_EFFECTS = "p_side_effect"
    COL_SIDE_TYPE    = "p_side_effect_type"
    COL_HEARD        = "p_heard_ivm_camp"    # IVM-specific; use for IVM analysis
    COL_INFO_CHANNEL = "p_info_chanel"       # ESPEN typo — missing 'n'
    COL_SATISFACTION = "p_satisfied"
    COL_MDA_LOCATION = "p_mda_location"
    COL_TREATED_TEAM = "p_treated_team"

    # ⚠ Unmapped (no ESPEN equivalent) — see module docstring for details:
    #   admin3 (sub-district), cddvisit_ov, absent_ov

    # Label dictionaries — ESPEN provides real category strings, no codebook needed
    SATISFACTION_LABELS = {
        "yes":          "Satisfied",
        "no":           "Not satisfied",
        "do not know":  "Do not know",
    }
    INFO_CHANNEL_LABELS = {
        "Community.leader":  "Community leader",
        "Family":            "Family member",
        "Health.worker":     "Health worker",
        "Town.crier":        "Town crier",
        "Radio":             "Radio",
        "School":            "School",
        "Poster":            "Poster / notice",
        "CDD":               "Drug distributor",
    }
    REASONS_NSW_LABELS = {
        "Fear.side.effects":    "Fear of side effects",
        "Minor.elderly":        "Minor / elderly exclusion",
        "Pregnancy.breastfeed": "Pregnant / breastfeeding",
        "Lack.of.trust":        "Lack of trust",
        "Absent":               "Was absent",
        "Refused":              "Refused",
    }
    MDA_LOCATION_LABELS = {
        "House":       "Household visit",
        "Fixed.post":  "Fixed post",
        "School":      "School",
    }

    DHS_U5_PCT = 15.6
    USE_AI = False
    LLM_PROVIDER = "template"
    LLM_MODEL = None

    def ai_write(section_name, user_prompt, fallback_text):
        return fallback_text

    # ── Data loading ─────────────────────────────────────────────────────────────
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=SHEET_NAME)
    print(f"Loaded {len(raw):,} rows × {len(raw.columns)} columns from '{SHEET_NAME}'")

    required = [COL_DISTRICT, COL_SITE, COL_SEX, COL_AGE, COL_SWALLOW]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Missing required ESPEN columns: {missing}. "
            "Ensure the file uses ESPEN p_* column format and the correct drug code."
        )

    df = raw.copy()
    df['district']  = df[COL_DISTRICT].astype(str).str.strip()
    df['region']    = df[COL_REGION].astype(str).str.strip()   if COL_REGION in df.columns else ''
    df['site']      = df[COL_SITE].astype(str).str.strip()
    df['village']   = df['site']     # alias kept for statistical function compatibility
    df['cluster']   = df['site']     # cluster unit for variance estimation
    df['site_name'] = df['site']

    df['sex']       = df[COL_SEX].astype(str).str.strip().str.capitalize()
    df['age']       = pd.to_numeric(df[COL_AGE], errors='coerce')

    # Derive age groups from p_age (not pre-computed in ESPEN)
    def _age_group(a):
        if pd.isna(a): return 'unknown'
        if a < 5:  return 'under 5'
        if a <= 14: return '5-14'
        if a <= 29: return '15-29'
        if a <= 44: return '30-44'
        if a <= 59: return '45-59'
        if a <= 74: return '60-74'
        return '75 plus'
    df['age_grp'] = df['age'].apply(_age_group)

    df['offered']      = df[COL_OFFERED].astype(str).str.strip().str.lower()      if COL_OFFERED      in df.columns else ''
    df['swallow']      = df[COL_SWALLOW].astype(str).str.strip().str.lower()      if COL_SWALLOW      in df.columns else ''
    df['reasons_nsw']  = df[COL_REASONS_NSW].astype(str).str.strip()              if COL_REASONS_NSW  in df.columns else ''
    df['reasons_sw']   = df[COL_REASONS_SW].astype(str).str.strip()               if COL_REASONS_SW   in df.columns else ''
    df['nt_treat']     = df[COL_NT_TREAT].astype(str).str.strip().str.lower()     if COL_NT_TREAT     in df.columns else ''
    df['nt_when']      = df[COL_NT_WHEN].astype(str).str.strip().str.lower()      if COL_NT_WHEN      in df.columns else ''
    df['side_eff']     = df[COL_SIDE_EFFECTS].astype(str).str.strip().str.lower() if COL_SIDE_EFFECTS in df.columns else ''
    df['side_type']    = df[COL_SIDE_TYPE].astype(str).str.strip()                if COL_SIDE_TYPE    in df.columns else ''
    df['heard']        = df[COL_HEARD].astype(str).str.strip().str.lower()        if COL_HEARD        in df.columns else ''
    df['info_channel'] = df[COL_INFO_CHANNEL].astype(str).str.strip()             if COL_INFO_CHANNEL in df.columns else ''
    df['satisfaction'] = df[COL_SATISFACTION].astype(str).str.strip().str.lower() if COL_SATISFACTION in df.columns else ''
    df['mda_location'] = df[COL_MDA_LOCATION].astype(str).str.strip()             if COL_MDA_LOCATION in df.columns else ''
    df['respondent']   = df[COL_RESPONDENT].astype(str).str.strip()               if COL_RESPONDENT   in df.columns else ''

    # Coverage indicators — derived from ESPEN (not pre-computed)
    # Epi: swallowed / total sampled population
    df['cov_epi'] = np.where(df['swallow'] == 'yes', 1,
                    np.where(df['offered'] == 'no',  0,
                    np.where(df['swallow'] == 'no',  0, np.nan)))
    # Therapeutic: swallowed / offered (those reached)
    df['cov_thera'] = np.where(
        (df['offered'] == 'yes') & (df['swallow'] == 'yes'), 1,
        np.where((df['offered'] == 'yes') & (df['swallow'] == 'no'), 0, np.nan))

    df['is_female']    = (df['sex'].str.lower() == 'female').astype(int)
    df['is_male']      = (df['sex'].str.lower() == 'male').astype(int)
    df['is_under5']    = (df['age_grp'] == 'under 5').astype(int)
    df['offered_yes']  = (df['offered'] == 'yes').astype(int)
    df['offered_no']   = (df['offered'] == 'no').astype(int)
    df['swallowed_yes']= (df['swallow'] == 'yes').astype(int)
    df['heard_yes']    = (df['heard']   == 'yes').astype(int)
    df['side_eff_yes'] = (df['side_eff'] == 'yes').astype(int)
    # ESPEN never-treated: 'never' = truly never treated
    df['never_truly']  = (df['nt_treat'] == 'never').astype(int)
    # Person responding for themselves: ESPEN uses 'Himself'
    df['person_responding'] = np.where(df['respondent'] == 'Himself', 'themselves', 'proxy')

    before = len(df)
    df = df.dropna(subset=['district', 'site'])
    df = df[df['district'].str.lower() != 'nan']
    df = df[df['site'].str.lower() != 'nan']
    print(f"After cleaning: {len(df):,} rows (dropped {before - len(df):,})")

    districts   = sorted(df['district'].unique())
    n_districts = len(districts)
    print(f"Districts: {districts}")
    print(f"Sites (clusters): {df['site'].nunique()} | Interviews: {len(df):,}")

    # ── Analysis engine ──────────────────────────────────────────────────────────
    FIGS, TABLES, STATS = {}, {}, {}
    print("Running analyses…")

    # 1. SAMPLE SIZES & AGE BREAKDOWN
    sample_by_dist = {d: int((df['district']==d).sum()) for d in districts}
    STATS['sample'] = {
        'n_total':     len(df),
        'by_district': sample_by_dist,
        'n_sites':     df['site'].nunique(),
        'n_villages':  df['site'].nunique(),   # alias
        'n_female':    int(df['is_female'].sum()),
        'n_male':      int(df['is_male'].sum()),
        'pct_female':  round(100*df['is_female'].mean(), 1),
        'age_median':  round(float(df['age'].median()), 0) if df['age'].notna().any() else None,
    }

    STATS['admin_breakdown'] = {
        'overall': {
            'districts': len(districts),
            'sites':     int(df['site'].nunique()),
        },
        'by_district': {
            d: {
                'n_sites':      int(df.loc[df['district']==d, 'site'].nunique()),
                'n_interviews': int((df['district']==d).sum()),
                'site_list':    sorted(df.loc[df['district']==d,'site'].dropna().unique().tolist()),
            } for d in districts
        },
    }

    AGE_ORDER = ['under 5','5-14','15-29','30-44','45-59','60-74','75 plus']
    age_rows = []
    for grp in AGE_ORDER:
        row = {'Age group': grp}
        for d in districts:
            sub = df[df['district']==d]
            row[d] = f"{100*(sub['age_grp']==grp).sum()/len(sub):.1f}" if len(sub) else '—'
        age_rows.append(row)
    age_df = pd.DataFrame(age_rows)
    TABLES['age_breakdown'] = age_df.to_html(index=False, border=0, classes='ces-table')

    u5_by_dist = {d: 100*(df[df['district']==d]['age_grp']=='under 5').sum()/max(1,(df['district']==d).sum())
                  for d in districts}
    STATS['age'] = {
        'u5_by_district': {k: round(v,1) for k,v in u5_by_dist.items()},
        'dhs_u5_pct': DHS_U5_PCT,
        'u5_under_dhs_districts': [d for d,v in u5_by_dist.items() if v < DHS_U5_PCT - 2],
    }

    fig_a, ax_a = plt.subplots(figsize=(max(8, n_districts*2.0), 5))
    x = np.arange(len(AGE_ORDER)); w = 0.8/max(1, n_districts)
    colors_d = plt.cm.tab10(np.linspace(0, 1, n_districts))
    for i, d in enumerate(districts):
        pct_vals = [100*(df[df['district']==d]['age_grp']==g).sum()/max(1,(df['district']==d).sum())
                    for g in AGE_ORDER]
        ax_a.bar(x + i*w - (n_districts-1)*w/2, pct_vals, w,
                 color=colors_d[i], edgecolor='black', lw=0.4, label=d)
    ax_a.axhline(DHS_U5_PCT, color=RED, lw=1.2, linestyle='--',
                 label=f'National u5 % (DHS {DHS_U5_PCT}%)')
    ax_a.set_xticks(x); ax_a.set_xticklabels(AGE_ORDER, fontsize=9)
    ax_a.set_ylabel('% of sampled population', fontweight='bold')
    ax_a.set_title('Age Breakdown of Sampled Population', fontweight='bold')
    ax_a.legend(fontsize=9, loc='upper right')
    plt.tight_layout()
    FIGS['fig_age_breakdown'] = fig_to_b64(fig_a)
    print("  ✓ Age breakdown")

    # 2. EPIDEMIOLOGICAL COVERAGE
    epi_rows = []
    for d in districts:
        sub = df[df['district']==d]
        sc    = cluster_prop_logit_ci(sub, 'cov_epi')
        sc_f  = cluster_prop_logit_ci(sub[sub['sex'].str.lower()=='female'], 'cov_epi')
        sc_m  = cluster_prop_logit_ci(sub[sub['sex'].str.lower()=='male'],   'cov_epi')
        sc_1s = cluster_prop_logit_ci(sub, 'cov_epi', level=0.90, one_sided_lower=True)
        p_sex = svychisq_sex(sub, 'cov_epi')
        epi_rows.append(dict(district=d, n=sc['n'], n_clusters=sc['n_clusters'],
            pct=sc['phat']*100, lo=sc['lower']*100, hi=sc['upper']*100,
            lo_1s=sc_1s['lower']*100, deff=sc['deff'],
            pct_f=sc_f['phat']*100, pct_m=sc_m['phat']*100, p_sex=p_sex,
            meets=sc_1s['lower']*100 >= THRESH_EPI))
    epi = pd.DataFrame(epi_rows)
    sc_all_epi    = cluster_prop_logit_ci(df, 'cov_epi')
    sc_all_epi_1s = cluster_prop_logit_ci(df, 'cov_epi', level=0.90, one_sided_lower=True)
    sc_all_epi_f  = cluster_prop_logit_ci(df[df['sex'].str.lower()=='female'], 'cov_epi')
    sc_all_epi_m  = cluster_prop_logit_ci(df[df['sex'].str.lower()=='male'],   'cov_epi')

    STATS['coverage_epi'] = {
        'overall_pct':     round(sc_all_epi['phat']*100, 1),
        'lower_95':        round(sc_all_epi['lower']*100, 1),
        'upper_95':        round(sc_all_epi['upper']*100, 1),
        'lower_1sided':    round(sc_all_epi_1s['lower']*100, 1),
        'deff':            sc_all_epi['deff'],
        'meets_threshold': bool(sc_all_epi_1s['lower']*100 >= THRESH_EPI),
        'threshold':       THRESH_EPI,
        'female_pct':      round(sc_all_epi_f['phat']*100, 1),
        'male_pct':        round(sc_all_epi_m['phat']*100, 1),
        'n_good_districts':int((epi['meets']).sum()),
        'n_low_districts': int((~epi['meets']).sum()),
        'by_district': {r['district']: {'pct': round(r['pct'],1),
                                         'lo': round(r['lo'],1), 'hi': round(r['hi'],1),
                                         'deff': r['deff'], 'meets': bool(r['meets'])}
                        for _, r in epi.iterrows()},
    }

    # Executive summary headlines
    _by_dist = STATS['coverage_epi']['by_district']
    _n_total = len(_by_dist)
    _eff_pass = [d for d, v in _by_dist.items() if v['meets']]
    _n_eff    = len(_eff_pass)
    _eff_all  = _n_total > 0 and _n_eff == _n_total
    _eff_none = _n_total > 0 and _n_eff == 0

    _align_pass, _align_skip = [], []
    if REPORTED_COVERAGE:
        for d, v in _by_dist.items():
            if d in REPORTED_COVERAGE:
                rep_pct = REPORTED_COVERAGE[d] * 100
                if v['lo'] <= rep_pct <= v['hi']:
                    _align_pass.append(d)
            else:
                _align_skip.append(d)
        _n_align_total = _n_total - len(_align_skip)
    else:
        _n_align_total = 0
    _n_align   = len(_align_pass)
    _align_all = _n_align_total > 0 and _n_align == _n_align_total
    _align_none= _n_align_total > 0 and _n_align == 0

    def _headline_effective():
        use_binary = (_n_total <= 2) and (_eff_all or _eff_none)
        if use_binary:
            return ('Effective coverage achieved' if _eff_all else 'Effective coverage NOT achieved'), _eff_all
        return (f'Effective coverage achieved in {_n_eff}/{_n_total} districts', _eff_all)

    def _headline_alignment():
        if _n_align_total == 0:
            return ('Reported coverage not supplied — cannot validate against CES results', None)
        use_binary = (_n_align_total <= 2) and (_align_all or _align_none)
        if use_binary:
            return (('Reported coverage aligns with CES results' if _align_all
                     else 'Reported coverage does NOT align with CES results'), _align_all)
        return (f'Reported coverage aligns with CES results in {_n_align}/{_n_align_total} districts', _align_all)

    _eff_text, _eff_ok = _headline_effective()
    _aln_text, _aln_ok = _headline_alignment()
    STATS['headlines'] = {
        'effective': {'text': _eff_text, 'all_ok': _eff_ok,
                      'pass': _eff_pass, 'n_pass': _n_eff, 'n_total': _n_total,
                      'threshold': THRESH_EPI},
        'alignment': {'text': _aln_text, 'all_ok': _aln_ok,
                      'pass': _align_pass, 'n_pass': _n_align,
                      'n_total': _n_align_total, 'skipped': _align_skip},
    }
    print(f"  ✓ Executive headlines: effective {_n_eff}/{_n_total}; alignment {_n_align}/{_n_align_total}")

    # 3. THERAPEUTIC COVERAGE
    thera_rows = []
    for d in districts:
        sub = df[df['district']==d]
        sc    = cluster_prop_logit_ci(sub, 'cov_thera')
        sc_1s = cluster_prop_logit_ci(sub, 'cov_thera', level=0.90, one_sided_lower=True)
        thera_rows.append(dict(district=d, n=sc['n'], pct=sc['phat']*100,
            lo=sc['lower']*100, hi=sc['upper']*100, lo_1s=sc_1s['lower']*100,
            deff=sc['deff'], meets=sc_1s['lower']*100 >= THRESH_THERA))
    thera = pd.DataFrame(thera_rows)
    sc_all_thera    = cluster_prop_logit_ci(df, 'cov_thera')
    sc_all_thera_1s = cluster_prop_logit_ci(df, 'cov_thera', level=0.90, one_sided_lower=True)
    STATS['coverage_thera'] = {
        'overall_pct':     round(sc_all_thera['phat']*100, 1),
        'lower_95':        round(sc_all_thera['lower']*100, 1),
        'upper_95':        round(sc_all_thera['upper']*100, 1),
        'lower_1sided':    round(sc_all_thera_1s['lower']*100, 1),
        'meets_threshold': bool(sc_all_thera_1s['lower']*100 >= THRESH_THERA),
        'threshold':       THRESH_THERA,
        'n_good_districts':int((thera['meets']).sum()),
        'by_district': {r['district']: {'pct': round(r['pct'],1),
                                         'lo': round(r['lo'],1), 'hi': round(r['hi'],1),
                                         'deff': r['deff'], 'meets': bool(r['meets'])}
                        for _, r in thera.iterrows()},
    }

    REPORTED_COLOR = '#8338EC'
    fig_c, axes = plt.subplots(1, 2, figsize=(max(10, n_districts*2.4), 5.2), sharey=True)
    for ax, lbl, dfc, thr, sc_all in [
            (axes[0], 'Epidemiological Coverage', epi, THRESH_EPI, sc_all_epi),
            (axes[1], 'Therapeutic Coverage',     thera, THRESH_THERA, sc_all_thera)]:
        x = np.arange(len(dfc))
        bar_colors = [GREEN if r['meets'] else RED for _, r in dfc.iterrows()]
        ax.bar(x, dfc['pct'], color=bar_colors, alpha=0.30, edgecolor='black', lw=0.5,
               width=0.55, zorder=2)
        ax.errorbar(x - 0.12, dfc['pct'],
            yerr=[dfc['pct']-dfc['lo'], dfc['hi']-dfc['pct']],
            fmt='o', color='black', ms=8, capsize=5, zorder=6,
            label='Surveyed coverage (95% CI)')
        rep_x, rep_y = [], []
        for xi, (_, r) in zip(x, dfc.iterrows()):
            d = r['district']
            if REPORTED_COVERAGE and d in REPORTED_COVERAGE:
                rep_x.append(xi + 0.12)
                rep_y.append(REPORTED_COVERAGE[d] * 100)
        if rep_x:
            ax.scatter(rep_x, rep_y, marker='D', s=95, color=REPORTED_COLOR,
                       edgecolors='black', lw=0.6, zorder=7, label='Reported coverage')
            for rx, ry in zip(rep_x, rep_y):
                ax.text(rx + 0.04, ry, f"{ry:.1f}%", va='center', ha='left',
                        fontsize=8.5, color=REPORTED_COLOR, fontweight='bold')
        ax.axhline(thr, color=RED, lw=1.5, linestyle='--', label=f'WHO {thr}%')
        ax.set_xticks(x); ax.set_xticklabels(dfc['district'].tolist(), rotation=20, ha='right', fontsize=10)
        ax.set_title(lbl, fontweight='bold')
        y_ceil = max(112, dfc['hi'].max() + 18)
        ax.set_ylim(0, y_ceil)
        ax.legend(fontsize=8.5, loc='lower right', framealpha=0.95)
        for xi, (_, r) in zip(x, dfc.iterrows()):
            ax.text(xi - 0.12, r['hi'] + 6, f"{r['pct']:.1f}%",
                    ha='center', va='bottom', fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.25', facecolor='white', edgecolor='none', alpha=0.85))
    axes[0].set_ylabel(f'Coverage (%) — {DRUG}', fontweight='bold')
    plt.suptitle(f'Survey Coverage by District — {COUNTRY}, {REPORT_DATE}', fontweight='bold', y=1.02)
    plt.tight_layout()
    FIGS['fig_coverage_district'] = fig_to_b64(fig_c)
    print("  ✓ Coverage by district (epi + therapeutic)")

    # Village-level dot plots
    for d in districts:
        sub = df[df['district'] == d]
        vill_rows = []
        for v in sub['site'].unique():
            vsub = sub[(sub['site'] == v) & (sub['cov_epi'].notna())]
            n = len(vsub)
            if n == 0: continue
            n_sw = int(vsub['cov_epi'].sum())
            pct  = 100 * n_sw / n
            lo, hi = wilson_ci(n_sw, n)
            vill_rows.append({'village': str(v), 'n': n,
                              'pct': round(pct, 1), 'lo': round(lo, 1), 'hi': round(hi, 1)})
        if not vill_rows: continue
        vdf = pd.DataFrame(vill_rows).sort_values('pct').reset_index(drop=True)
        n_v = len(vdf)
        fig_v, ax_v = plt.subplots(figsize=(max(9, n_v * 0.55), 5))
        x = np.arange(n_v)
        dot_colors = [GREEN if r['pct'] >= THRESH_EPI else RED for _, r in vdf.iterrows()]
        ax_v.errorbar(x, vdf['pct'],
                      yerr=[(vdf['pct'] - vdf['lo']).clip(lower=0),
                            (vdf['hi'] - vdf['pct']).clip(lower=0)],
                      fmt='none', ecolor='#555', capsize=4, lw=1.2, zorder=3)
        ax_v.scatter(x, vdf['pct'], c=dot_colors, s=72, zorder=5, edgecolors='black', lw=0.6)
        if REPORTED_COVERAGE and d in REPORTED_COVERAGE:
            rep_pct = REPORTED_COVERAGE[d] * 100
            ax_v.axhline(rep_pct, color='#E05080', lw=1.5, linestyle=':',
                         label=f'Reported coverage ({rep_pct:.1f}%)')
        ax_v.axhline(THRESH_EPI, color='#228B22', lw=2, linestyle='-',
                     label=f'WHO target {THRESH_EPI}%')
        ax_v.set_xticks(x)
        ax_v.set_xticklabels(vdf['village'].tolist(), rotation=45, ha='right',
                             fontsize=max(6, 9 - n_v//10))
        ax_v.set_ylabel('Epidemiological coverage (%)', fontweight='bold')
        ax_v.set_xlabel('Site (cluster) — sorted by coverage', fontweight='bold')
        ax_v.set_ylim(0, 108)
        ax_v.set_title(f'Survey Coverage by Site — {d}\n{COUNTRY}, {REPORT_DATE}', fontweight='bold')
        ax_v.legend(fontsize=9, loc='upper left')
        plt.tight_layout()
        FIGS[f'fig_village_{d}'] = fig_to_b64(fig_v)
    print("  ✓ Site-level coverage dot plots")

    # 4. GEOGRAPHIC COVERAGE — % sites with ≥1 person treated
    geo_rows = []
    for d in districts:
        sub = df[df['district']==d]
        by_site = sub.groupby('site').apply(lambda x: (x['swallow']=='yes').any())
        n_s  = len(by_site)
        n_ok = int(by_site.sum())
        geo_rows.append(dict(district=d, n_villages=n_s, n_treated=n_ok,
                             pct=100*n_ok/n_s if n_s else 0))
    geo = pd.DataFrame(geo_rows)
    STATS['coverage_geo'] = {r['district']: {'n_villages':int(r['n_villages']),
                                              'n_treated':int(r['n_treated']),
                                              'pct':round(r['pct'],1)}
                              for _, r in geo.iterrows()}

    missed = []
    for d in districts:
        sub = df[df['district']==d]
        grp = sub.groupby('site').apply(lambda x: (x['swallow']=='yes').sum())
        for v, n in grp.items():
            if n == 0:
                missed.append({'district': d, 'site': v,
                               'n_interviews': int((sub['site']==v).sum())})
    STATS['missed_villages'] = missed

    # 5. SIMPLIFIED CASCADE — Total → Received → Swallowed
    # (CDD Visited and Present steps removed — not collected in ESPEN)
    CASCADE_STEPS = ['Total', 'Received', 'Swallowed']
    casc_rows = []
    for d in districts:
        sub     = df[df['district'] == d]
        n_total = len(sub)
        if n_total == 0: continue
        n_received  = int((sub['offered'] == 'yes').sum())
        n_swallowed = int((sub['swallow'] == 'yes').sum())
        counts = [n_total, n_received, n_swallowed]
        casc_rows.append({
            'district':  d,
            'Total':     100.0,
            'Received':  round(100 * n_received  / n_total, 1),
            'Swallowed': round(100 * n_swallowed / n_total, 1),
            '_counts':   counts,
        })
    casc_df = pd.DataFrame(casc_rows)
    STATS['cascade'] = {r['district']: {s: r[s] for s in CASCADE_STEPS}
                        for _, r in casc_df.iterrows()}

    n_steps = len(CASCADE_STEPS)
    n_rows  = len(casc_df)
    dist_palette = [BLUE, ORANGE, GREEN, NAVY, '#9B59B6', '#E67E22']
    box_w, box_h, gap = 1.75, 0.85, 0.6
    left_pad   = 2.0
    right_pad  = 0.4
    fig_w      = left_pad + n_steps * box_w + (n_steps - 1) * gap + right_pad
    row_h      = 1.7
    top_pad    = 1.1
    bottom_pad = 0.4
    fig_h_cas  = top_pad + row_h * n_rows + bottom_pad

    fig_cas, ax_cas = plt.subplots(figsize=(fig_w, fig_h_cas))
    ax_cas.set_xlim(0, fig_w); ax_cas.set_ylim(0, fig_h_cas)
    ax_cas.axis('off'); ax_cas.grid(False)

    ax_cas.text(fig_w/2, fig_h_cas - 0.30,
                f'Treatment Pathway — {COUNTRY}, {REPORT_DATE}',
                ha='center', va='top', fontweight='bold', fontsize=12)
    ax_cas.text(fig_w/2, fig_h_cas - 0.62,
                'ESPEN cascade: Total → Received drug → Swallowed. Arrows show drop-off (Δ pp; n).',
                ha='center', va='top', fontsize=9, color='#555', style='italic')

    for ri, (_, row) in enumerate(casc_df.iterrows()):
        y_center = fig_h_cas - top_pad - (ri + 0.5) * row_h
        color    = dist_palette[ri % len(dist_palette)]
        counts   = row['_counts']
        ax_cas.text(left_pad - 0.25, y_center, row['district'],
                    ha='right', va='center', fontweight='bold', fontsize=11, color=color)
        for si, step in enumerate(CASCADE_STEPS):
            x   = left_pad + si * (box_w + gap)
            pct = row[step]
            n_at = counts[si]
            rect = mpatches.FancyBboxPatch(
                (x, y_center - box_h/2), box_w, box_h,
                boxstyle="round,pad=0.02,rounding_size=0.08",
                linewidth=1.3, edgecolor=color, facecolor=color, alpha=0.18)
            ax_cas.add_patch(rect)
            ax_cas.text(x + box_w/2, y_center + 0.22, step,
                        ha='center', va='center', fontsize=9.5, fontweight='bold')
            ax_cas.text(x + box_w/2, y_center - 0.02, f"{pct:.1f}%",
                        ha='center', va='center', fontsize=12, fontweight='bold', color=color)
            ax_cas.text(x + box_w/2, y_center - 0.27, f"n = {n_at:,}",
                        ha='center', va='center', fontsize=8, color='#444')
            if si < n_steps - 1:
                xs = x + box_w; xe = x + box_w + gap
                ax_cas.annotate('', xy=(xe, y_center), xytext=(xs, y_center),
                    arrowprops=dict(arrowstyle='->,head_width=0.35,head_length=0.55',
                                    color='#666', lw=1.4))
                d_pct = row[CASCADE_STEPS[si+1]] - row[step]
                d_n   = counts[si+1] - counts[si]
                mid   = (xs + xe) / 2
                if d_n != 0:
                    ax_cas.text(mid, y_center + 0.34, f"{d_pct:+.1f} pp",
                                ha='center', va='bottom', fontsize=8, color='#B33', fontweight='bold')
                    ax_cas.text(mid, y_center - 0.36, f"({d_n:+,})",
                                ha='center', va='top', fontsize=7.5, color='#B33')
                else:
                    ax_cas.text(mid, y_center + 0.34, '— no loss —',
                                ha='center', va='bottom', fontsize=8, color='#888')

    ax_cas.text(fig_w - right_pad - 0.05, 0.18,
                f'Epi coverage threshold: {THRESH_EPI}%',
                ha='right', va='bottom', fontsize=8.5, color=RED, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.25', facecolor='white', edgecolor=RED, lw=0.8))
    plt.tight_layout()
    FIGS['fig_cascade'] = fig_to_b64(fig_cas)
    print("  ✓ Treatment pathway cascade (3-step ESPEN)")

    # 6. DRUG RECEIVED (replaces "Offered" section)
    recv_rows = []
    for d in districts:
        sub = df[df['district']==d]
        valid = sub[sub['offered'].isin(['yes','no'])]
        n = len(valid)
        if n == 0: continue
        pct_y = 100*(valid['offered']=='yes').sum()/n
        pct_n = 100*(valid['offered']=='no').sum()/n
        recv_rows.append({'district':d, 'n':n,
                          'pct_yes':round(pct_y,1), 'pct_no':round(pct_n,1)})
    recv_df = pd.DataFrame(recv_rows)
    STATS['offered'] = {r['district']: {'pct_yes':r['pct_yes'],'pct_no':r['pct_no'],'n':int(r['n'])}
                        for _, r in recv_df.iterrows()}

    if not recv_df.empty:
        fig_o, ax_o = plt.subplots(figsize=(max(7, n_districts*1.6), 4.5))
        x = np.arange(len(recv_df)); w = 0.35
        ax_o.bar(x - w/2, recv_df['pct_yes'], w, color=GREEN,  edgecolor='black', lw=0.4, label='Received')
        ax_o.bar(x + w/2, recv_df['pct_no'],  w, color=RED,    edgecolor='black', lw=0.4, label='Not received')
        ax_o.set_xticks(x); ax_o.set_xticklabels(recv_df['district'].tolist(), fontsize=10)
        ax_o.set_ylabel('% of sampled respondents', fontweight='bold')
        ax_o.set_title(f'{DRUG} Received by District', fontweight='bold')
        ax_o.set_ylim(0, 105); ax_o.legend(fontsize=9, loc='upper right')
        for i, r in recv_df.iterrows():
            ax_o.text(i-w/2, r['pct_yes']+1, f"{r['pct_yes']:.0f}%", ha='center', fontsize=9)
            ax_o.text(i+w/2, r['pct_no']+1,  f"{r['pct_no']:.0f}%",  ha='center', fontsize=9)
        plt.tight_layout()
        FIGS['fig_offered'] = fig_to_b64(fig_o)
    print("  ✓ Drug received")

    # 7. REASONS NOT SWALLOWED (replaces "Reasons Not Offered")
    # ESPEN: p_reason_not_swallowed_{drug} — applies to those who received but didn't swallow
    nsw_sub = df[(df['offered']=='yes') & (df['swallow']=='no')]
    nsw_overall = parse_multiselect(nsw_sub['reasons_nsw'], separator=r'\s+')
    nsw_labelled = relabel(nsw_overall.head(8), REASONS_NSW_LABELS)
    STATS['reasons_not_swallowed'] = nsw_labelled.to_dict()

    nsw_by_dist = {}
    for d in districts:
        sub = df[(df['district']==d) & (df['offered']=='yes') & (df['swallow']=='no')]
        nsw_by_dist[d] = {'n': len(sub)}
        if len(sub):
            cts = parse_multiselect(sub['reasons_nsw'])
            top = cts.head(4)
            nsw_by_dist[d]['reasons'] = {REASONS_NSW_LABELS.get(k, k): int(v) for k, v in top.items()}
    STATS['not_offered'] = nsw_by_dist

    # Also compute STATS entries for backward compat with HTML table
    for d in districts:
        STATS['not_offered'].setdefault(d, {})
        STATS['not_offered'][d].setdefault('n_not_offered', nsw_by_dist.get(d, {}).get('n', 0))
        STATS['not_offered'][d].setdefault('pct_cdd_not_visit', 0)
        STATS['not_offered'][d].setdefault('pct_absent', 0)
    STATS['absent_reasons'] = []

    if nsw_labelled.any():
        fig_nsw, ax_nsw = plt.subplots(figsize=(8, max(3, 0.55*len(nsw_labelled))))
        labels_nsw = list(nsw_labelled.index)
        vals_nsw   = list(nsw_labelled.values)
        ax_nsw.barh(range(len(nsw_labelled)), vals_nsw, color=ORANGE, edgecolor='black', lw=0.4)
        ax_nsw.set_yticks(range(len(nsw_labelled)))
        ax_nsw.set_yticklabels(labels_nsw, fontsize=10)
        ax_nsw.invert_yaxis()
        ax_nsw.set_xlabel('Number of mentions', fontweight='bold')
        ax_nsw.set_title(f'Reasons for Not Swallowing {DRUG} (among those offered)', fontweight='bold')
        for i, v in enumerate(vals_nsw):
            ax_nsw.text(v + 0.3, i, f" {v}", va='center', fontsize=9)
        plt.tight_layout()
        FIGS['fig_reasons_not_offered'] = fig_to_b64(fig_nsw)
    print("  ✓ Reasons not swallowed")

    # 8. MDA DELIVERY LOCATION (new — from p_mda_location)
    loc_rows = []
    for d in districts:
        sub = df[(df['district']==d) & (df['mda_location'].isin(['House','Fixed.post','School']))]
        n = len(sub)
        if n == 0: continue
        row = {'district': d, 'n': n}
        for loc in ['House','Fixed.post','School']:
            row[MDA_LOCATION_LABELS[loc]] = round(100*(sub['mda_location']==loc).sum()/n, 1)
        loc_rows.append(row)
    loc_df = pd.DataFrame(loc_rows)
    STATS['mda_location'] = loc_df.to_dict(orient='records') if not loc_df.empty else []

    if not loc_df.empty:
        loc_cats = list(MDA_LOCATION_LABELS.values())
        loc_colors = [BLUE, ORANGE, GREEN]
        fig_loc, ax_loc = plt.subplots(figsize=(max(7, n_districts*2), 5))
        x_loc = np.arange(len(loc_df)); w_loc = 0.26
        for i, (cat, col) in enumerate(zip(loc_cats, loc_colors)):
            if cat in loc_df.columns:
                vals = loc_df[cat].fillna(0).values
                ax_loc.bar(x_loc + (i-1)*w_loc, vals, w_loc,
                           color=col, edgecolor='black', lw=0.4, label=cat)
        ax_loc.set_xticks(x_loc)
        ax_loc.set_xticklabels(loc_df['district'].tolist(), fontsize=10)
        ax_loc.set_ylabel('% of treated respondents', fontweight='bold')
        ax_loc.set_title(f'MDA Delivery Location — {DRUG}', fontweight='bold')
        ax_loc.set_ylim(0, 105); ax_loc.legend(fontsize=9, loc='upper right')
        plt.tight_layout()
        FIGS['fig_mda_location'] = fig_to_b64(fig_loc)
    print("  ✓ MDA delivery location")

    # 9. COMPLIANCE (swallowed if received)
    comp_rows = []
    for d in districts:
        sub = df[(df['district']==d) & (df['offered']=='yes')]
        n = len(sub)
        if n == 0: continue
        pct = 100*(sub['swallow']=='yes').sum()/n
        comp_rows.append({'district':d,'n':n,'pct':round(pct,1)})
    comp_df = pd.DataFrame(comp_rows)
    STATS['compliance'] = {r['district']: {'pct':r['pct'],'n':int(r['n'])}
                           for _, r in comp_df.iterrows()}
    print(f"  ✓ Compliance ({len(comp_df)} districts)")

    # 10. NEVER-TREATED
    nt_rows = []
    nt_village_rows = []
    for d in districts:
        sub = df[(df['district']==d)
                 & (df['nt_treat'].isin(['never','once','twice or more']))
                 & (df['age'] >= 8)
                 & (df['person_responding'] == 'themselves')]
        n = len(sub)
        if n == 0: continue
        n_never = int((sub['nt_treat']=='never').sum())
        pct  = 100*n_never/n
        lo, hi = wilson_ci(n_never, n)
        never_sub = sub[sub['nt_treat']=='never']
        pct_f   = round(100*(never_sub['sex'].str.lower()=='female').mean(),1) if len(never_sub) else 0
        age_med = round(float(never_sub['age'].median()),0) if len(never_sub) and never_sub['age'].notna().any() else None
        nt_rows.append({'district':d,'n':n,'n_never':n_never,
                        'pct':round(pct,1),'lo':round(lo,1),'hi':round(hi,1),
                        'pct_female_nt':pct_f,'age_median_nt':age_med})
        for v, vsub in sub.groupby('site'):
            nv = len(vsub)
            if nv == 0: continue
            nv_never = int((vsub['nt_treat']=='never').sum())
            v_pct    = 100*nv_never/nv
            v_lo, v_hi = wilson_ci(nv_never, nv)
            nt_village_rows.append({'district':d,'village':str(v),'n':nv,'n_never':nv_never,
                                    'pct':round(v_pct,1),'lo':round(v_lo,1),'hi':round(v_hi,1)})
    nt_df  = pd.DataFrame(nt_rows)
    ntv_df = pd.DataFrame(nt_village_rows)

    STATS['never_treated'] = {r['district']: {'n':int(r['n']),'n_never':int(r['n_never']),
                                               'pct':r['pct'],'lo':r['lo'],'hi':r['hi'],
                                               'pct_female':r['pct_female_nt'],
                                               'age_median':r['age_median_nt']}
                               for _, r in nt_df.iterrows()}
    STATS['never_treated_village'] = ntv_df.to_dict(orient='records')

    if not ntv_df.empty:
        dist_order   = [d for d in districts if d in set(ntv_df['district'].unique())]
        dist_color   = {d: dist_palette[i%len(dist_palette)] for i,d in enumerate(dist_order)}
        ntv_df['_dist_order'] = ntv_df['district'].map({d:i for i,d in enumerate(dist_order)})
        ntv_df = ntv_df.sort_values(['_dist_order','pct'], ascending=[True,False]).reset_index(drop=True)
        n_villages = len(ntv_df)
        fig_h_nt = max(4.0, 0.32*n_villages + 1.6)
        fig_nt, ax_nt = plt.subplots(figsize=(10, fig_h_nt))
        y      = np.arange(n_villages)
        colors = [dist_color[d] for d in ntv_df['district']]
        ax_nt.barh(y, ntv_df['pct'], color=colors, edgecolor='black', lw=0.4, alpha=0.85)
        ylabels = [f"{r['village']}  (n={int(r['n'])})" for _, r in ntv_df.iterrows()]
        ax_nt.set_yticks(y); ax_nt.set_yticklabels(ylabels, fontsize=8.5)
        ax_nt.invert_yaxis()
        x_pad = max(float(ntv_df['pct'].max())*0.015, 0.25)
        for i, r in ntv_df.iterrows():
            ax_nt.text(r['pct']+x_pad, i, f"{r['pct']:.1f}%", va='center', ha='left', fontsize=8.5, fontweight='bold')
        for d2 in dist_order:
            if d2 not in STATS['never_treated']: continue
            dpct = STATS['never_treated'][d2]['pct']
            rows_idx = ntv_df.index[ntv_df['district']==d2].tolist()
            if not rows_idx: continue
            y_lo, y_hi = min(rows_idx)-0.45, max(rows_idx)+0.45
            ax_nt.vlines(dpct, y_lo, y_hi, colors=dist_color[d2], linestyles='dashed', lw=1.6, alpha=0.9)
        legend_handles = [mpatches.Patch(facecolor=dist_color[d2], edgecolor='black', lw=0.4, alpha=0.85,
                            label=f"{d2} (district mean {STATS['never_treated'][d2]['pct']:.1f}%)")
                          for d2 in dist_order if d2 in STATS['never_treated']]
        legend_handles.append(Line2D([0],[0], color='#555', lw=1.6, linestyle='dashed', label='District mean'))
        ax_nt.legend(handles=legend_handles, fontsize=8.5, loc='lower right', frameon=True, framealpha=0.95)
        x_max = max(float(ntv_df['pct'].max())*1.20, 5.0)
        ax_nt.set_xlim(0, x_max)
        ax_nt.set_xlabel('% never treated — one bar per site; n shown in label', fontweight='bold', fontsize=10)
        ax_nt.set_title('Never-Treated Population by Site\nAmong respondents aged 8+ answering for themselves',
                        fontweight='bold', fontsize=11)
        ax_nt.grid(axis='x', alpha=0.3)
        plt.tight_layout()
        FIGS['fig_never_treated'] = fig_to_b64(fig_nt)

    # Never-treated by age group
    NT_AGE_ORDER = [g for g in AGE_ORDER if g != 'under 5']
    nt_age_rows = []
    for d in districts:
        sub_d = df[(df['district']==d)
                   & (df['nt_treat'].isin(['never','once','twice or more']))
                   & (df['age']>=8)
                   & (df['person_responding']=='themselves')]
        for g in NT_AGE_ORDER:
            sub = sub_d[sub_d['age_grp']==g]
            n   = len(sub)
            n_never = int((sub['nt_treat']=='never').sum()) if n else 0
            pct = round(100*n_never/n,1) if n else None
            nt_age_rows.append({'district':d,'age_grp':g,'n':n,'n_never':n_never,'pct':pct})
    sub_all = df[(df['nt_treat'].isin(['never','once','twice or more']))
                 & (df['age']>=8) & (df['person_responding']=='themselves')]
    nt_age_overall = []
    for g in NT_AGE_ORDER:
        sub = sub_all[sub_all['age_grp']==g]
        n   = len(sub)
        n_never = int((sub['nt_treat']=='never').sum()) if n else 0
        nt_age_overall.append({'age_grp':g,'n':n,'n_never':n_never,'pct':round(100*n_never/n,1) if n else None})
    STATS['never_treated_age'] = {'by_district':nt_age_rows,'overall':nt_age_overall,'age_order':NT_AGE_ORDER}
    print("  ✓ Never-treated")

    # 11. SIDE EFFECTS — ESPEN provides real labels (Vomiting/Nausea/Dizziness/Headache)
    se_rows = []
    for d in districts:
        sub = df[(df['district']==d) & (df['side_eff'].isin(['yes','no']))]
        n = len(sub)
        if n == 0: continue
        pct = 100*(sub['side_eff']=='yes').sum()/n
        se_rows.append({'district':d,'n':n,'pct':round(pct,1)})
    se_df = pd.DataFrame(se_rows)
    STATS['side_effects'] = {r['district']: {'pct':r['pct'],'n':int(r['n'])} for _,r in se_df.iterrows()}

    # Real side-effect type labels from ESPEN — no codebook mapping needed
    se_types_overall = parse_multiselect(df['side_type'])
    STATS['side_effect_top_types'] = se_types_overall.head(8).to_dict()

    if not se_df.empty:
        fig_se, ax_se = plt.subplots(figsize=(max(6, n_districts*1.4), 4))
        bar_c = [GREEN if v < 10 else YELLOW if v < 20 else RED for v in se_df['pct']]
        ax_se.bar(se_df['district'], se_df['pct'], color=bar_c, edgecolor='black', lw=0.4)
        ax_se.set_ylabel('% reporting side effects', fontweight='bold')
        ax_se.set_title('Side Effects Reported (any) — by District', fontweight='bold')
        ax_se.set_ylim(0, max(20, se_df['pct'].max()*1.3))
        for i, r in se_df.iterrows():
            ax_se.text(i, r['pct']+0.3, f"{r['pct']:.1f}%", ha='center', fontsize=10, fontweight='bold')
        plt.tight_layout()
        FIGS['fig_side_effects'] = fig_to_b64(fig_se)
    print("  ✓ Side effects")

    # 12. COMMUNICATION — heard about IVM campaign
    heard_rows = []
    for d in districts:
        sub = df[(df['district']==d) & (df['heard'].isin(['yes','no','do not know']))]
        n = len(sub)
        if n == 0: continue
        pct = 100*(sub['heard']=='yes').sum()/n
        heard_rows.append({'district':d,'n':n,'pct':round(pct,1)})
    heard_df = pd.DataFrame(heard_rows)
    STATS['heard'] = {r['district']: {'pct':r['pct'],'n':int(r['n'])} for _,r in heard_df.iterrows()}

    # Real channel labels from ESPEN — no codebook needed
    heard_channels = parse_multiselect(df['info_channel'])
    heard_channels_labelled = relabel(heard_channels.head(8), INFO_CHANNEL_LABELS)
    STATS['heard_channels'] = heard_channels_labelled.to_dict()

    if not heard_df.empty:
        fig_h2, ax_h2 = plt.subplots(figsize=(max(6, n_districts*1.4), 4))
        ax_h2.bar(heard_df['district'], heard_df['pct'], color=BLUE, edgecolor='black', lw=0.4)
        ax_h2.set_ylabel('% heard about campaign before MDA', fontweight='bold')
        ax_h2.set_title('Pre-Campaign Communication Reach', fontweight='bold')
        ax_h2.set_ylim(0, 105)
        for i, r in heard_df.iterrows():
            ax_h2.text(i, r['pct']+1, f"{r['pct']:.1f}%", ha='center', fontsize=10, fontweight='bold')
        plt.tight_layout()
        FIGS['fig_heard'] = fig_to_b64(fig_h2)
    print("  ✓ Communication")

    # 13. SATISFACTION — ESPEN: Yes / No / Do not know (not 0-3 scale)
    sat_rows = []
    for d in districts:
        sub = df[df['district']==d]
        s_valid = sub[sub['satisfaction'].isin(['yes','no','do not know'])]
        n = len(s_valid)
        if n == 0: continue
        row = {'district':d, 'n':n}
        for code, label in SATISFACTION_LABELS.items():
            row[label] = round(100*(s_valid['satisfaction']==code).sum()/n, 1)
        sat_rows.append(row)
    sat_df = pd.DataFrame(sat_rows)
    STATS['satisfaction'] = sat_df.to_dict(orient='records')

    if not sat_df.empty:
        sat_stack_order = ['Not satisfied', 'Do not know', 'Satisfied']
        colors_sat = {'Not satisfied': RED, 'Do not know': GREY, 'Satisfied': GREEN}
        fig_sat, ax_sat = plt.subplots(figsize=(max(7, n_districts*2), 5.5))
        x_sat  = np.arange(len(sat_df))
        bottom = np.zeros(len(sat_df))
        for cat in sat_stack_order:
            if cat not in sat_df.columns: continue
            vals = sat_df[cat].fillna(0).values
            ax_sat.bar(x_sat, vals, bottom=bottom, color=colors_sat[cat],
                       edgecolor='white', lw=0.8, label=cat)
            for xi, (pct, bot, n_tot) in enumerate(zip(vals, bottom, sat_df['n'].values)):
                if pct < 6: continue  # skip label; bottom advance handled by bottom += vals
                n_seg = round(pct/100*n_tot)
                mid_y = bot + pct/2
                ax_sat.text(xi, mid_y, f"{pct:.1f}%\n(n={n_seg})", ha='center', va='center',
                            fontsize=8, fontweight='bold', color='white',
                            bbox=dict(boxstyle='round,pad=0.15', facecolor='none', edgecolor='none'))
            bottom += vals
        ax_sat.set_xticks(x_sat); ax_sat.set_xticklabels(sat_df['district'].tolist(), fontsize=11)
        ax_sat.set_ylabel('% of respondents', fontweight='bold')
        ax_sat.set_title(f'Satisfaction with MDA Campaign — {COUNTRY}, {REPORT_DATE}', fontweight='bold')
        ax_sat.set_ylim(0, 105); ax_sat.legend(fontsize=9, loc='upper right', framealpha=0.9)
        plt.tight_layout()
        FIGS['fig_satisfaction'] = fig_to_b64(fig_sat)
    print("  ✓ Satisfaction")
    print(f"\n✅  Analysis complete: {len(FIGS)} figures, {len(TABLES)} tables, {len(STATS)} stat blocks")
    # ── Narratives ───────────────────────────────────────────────────────────────
    NARRATIVES = {}
    ce = STATS['coverage_epi']; ct = STATS['coverage_thera']; sa = STATS['sample']
    print("Generating narratives…")

    ed_lines = [f"- {d}: Epi {v['pct']}% ({v['lo']}-{v['hi']}); Thera {STATS['coverage_thera']['by_district'][d]['pct']}%"
                for d, v in ce['by_district'].items()]

    exec_tpl = (
        f"The {DRUG} mass drug administration evaluated in this survey achieved an overall "
        f"epidemiological coverage of {ce['overall_pct']}% (95% CI {ce['lower_95']}%–{ce['upper_95']}%) "
        f"and a therapeutic coverage of {ct['overall_pct']}% (95% CI {ct['lower_95']}%–{ct['upper_95']}%). "
        f"The WHO epidemiological threshold of {ce['threshold']}% is "
        f"{'met' if ce['meets_threshold'] else 'NOT met'} at the overall level, and the "
        f"therapeutic threshold of {ct['threshold']}% is "
        f"{'met' if ct['meets_threshold'] else 'NOT met'}. Coverage met the {ce['threshold']}% epi "
        f"target in {ce['n_good_districts']} of {len(districts)} districts surveyed."
    )
    NARRATIVES['executive_summary'] = ai_write('exec_summary', '', exec_tpl)

    meth_tpl = (
        f"The survey followed WHO/ESPEN cluster-sampling methodology. A total of {sa['n_total']:,} "
        f"individuals were interviewed across {sa['n_sites']} sites in {len(districts)} "
        f"district{'s' if len(districts)!=1 else ''}: {', '.join(districts)}. "
        f"Coverage point estimates are accompanied by 95% confidence intervals computed with a logit "
        f"transformation under cluster-robust variance to account for the design effect of clustered sampling.\n\n"
        f"Two coverage metrics are reported throughout, in line with WHO guidance. Epidemiological "
        f"coverage uses the entire sampled population as the denominator and is judged against the "
        f"{THRESH_EPI}% threshold. Therapeutic coverage restricts the denominator to those who received "
        f"the drug and is judged against the {THRESH_THERA}% threshold. "
        f"Geographic coverage is reported as the proportion of sampled sites in which at least one person "
        f"received the drug. Data were collected using the ESPEN Coverage Evaluation Survey instrument."
    )
    NARRATIVES['methodology'] = ai_write('methodology', '', meth_tpl)

    def _casc_largest_drop(d):
        c = STATS['cascade'][d]
        pairs = [
            ('Total to Received',    c['Total']    - c['Received']),
            ('Received to Swallowed',c['Received'] - c['Swallowed']),
        ]
        return max(pairs, key=lambda x: x[1])

    _casc_tpl_parts = []
    for d, v in STATS['cascade'].items():
        step, drop = _casc_largest_drop(d)
        _casc_tpl_parts.append(
            f"In {d}, the largest attrition was {step} ({drop:.1f} pp lost), "
            f"with {v['Swallowed']}% of the total sampled population ultimately confirmed as treated."
        )
    casc_tpl = (
        "The treatment pathway cascade traces attrition from the total sampled population through "
        "drug receipt to confirmed swallowing — two distinct and addressable programmatic levers. "
        + " ".join(_casc_tpl_parts)
        + " Where the largest drop is at the receipt stage, the priority is reaching missed "
        "households; where it is at the swallowing stage, targeted messaging on benefit and safety "
        "is warranted."
    )
    NARRATIVES['cascade'] = ai_write('cascade', '', casc_tpl)

    u5_flag = STATS['age']['u5_under_dhs_districts']
    u5_by_dist_str = ', '.join(f"{d}: {STATS['age']['u5_by_district'][d]}%" for d in districts)
    ds_tpl = (
        f"The CES drew {sa['n_total']:,} interviews across {sa['n_sites']} sites in "
        f"{', '.join(districts)}. Women made up {sa['pct_female']}% of respondents and the median age was "
        f"{sa['age_median']} years. Children under five represented {u5_by_dist_str} of the sample, "
        f"compared with a national DHS reference of {DHS_U5_PCT}%. "
        + (f"Districts {', '.join(u5_flag)} appear to have under-represented this age group."
           if u5_flag else "Under-5 representation broadly tracks the national reference.")
    )
    NARRATIVES['dataset'] = ai_write('dataset', '', ds_tpl)

    MANUAL_EXCERPT = """\
    Validation of reported coverage (Coverage Evaluation Guidelines excerpt):
    The survey coverage can be compared with reported coverage; if similar, the reported
    coverage is considered validated. A more objective rule is whether reported coverage falls
    within the 95% CI of the survey estimate. Reported >> Survey (>25 pp) suggests denominator
    problems or reporting-chain errors. Reported << Survey suggests undercounting of treatments.
    """
    _cov_rep_lines = []
    for d, v in ce['by_district'].items():
        if REPORTED_COVERAGE and d in REPORTED_COVERAGE:
            rep = REPORTED_COVERAGE[d] * 100
            within = (v['lo'] <= rep <= v['hi'])
            _cov_rep_lines.append(
                f"  - {d}: survey {v['pct']}% (95% CI {v['lo']}-{v['hi']}); "
                f"reported {rep:.1f}%; within_CI={within}")
        else:
            _cov_rep_lines.append(
                f"  - {d}: survey {v['pct']}% (95% CI {v['lo']}-{v['hi']}); reported not supplied")

    cov_tpl_lines = []
    for d in districts:
        e = ce['by_district'][d]; t = ct['by_district'][d]
        cov_tpl_lines.append(
            f"In {d}, epidemiological coverage was {e['pct']}% (95% CI {e['lo']}–{e['hi']}) "
            f"and therapeutic coverage was {t['pct']}% (95% CI {t['lo']}–{t['hi']}). "
            f"The {ce['threshold']}% epi threshold is {'met' if e['meets'] else 'NOT met'}."
        )
    cov_tpl = (
        f"Overall, epidemiological coverage was {ce['overall_pct']}% (95% CI {ce['lower_95']}–{ce['upper_95']}) "
        f"and therapeutic coverage was {ct['overall_pct']}% (95% CI {ct['lower_95']}–{ct['upper_95']}). "
        + " ".join(cov_tpl_lines)
    )
    NARRATIVES['coverage'] = ai_write('coverage', '', cov_tpl)

    missed_str = (', '.join(f"{m['site']} ({m['district']})" for m in STATS['missed_villages'])
                  if STATS['missed_villages'] else 'none')
    geo_tpl_lines = [f"{d}: {v['n_treated']} of {v['n_villages']} sites had ≥1 treated person ({v['pct']}%)."
                     for d, v in STATS['coverage_geo'].items()]
    geo_tpl = (
        "Geographic coverage measures whether MDA reached every sampled site. "
        + " ".join(geo_tpl_lines)
        + (f" Sites with no treated respondents: {missed_str}."
           if STATS['missed_villages'] else " No sampled site was completely missed.")
    )
    NARRATIVES['geographic'] = ai_write('geographic', '', geo_tpl)

    off_tpl_lines = []
    for d, v in STATS['offered'].items():
        off_tpl_lines.append(
            f"In {d}, {v['pct_yes']}% of respondents received the drug and {v['pct_no']}% did not.")
    nsw_top = list(STATS['reasons_not_swallowed'].items())[:3]
    nsw_str = '; '.join(f"{k}: {v}" for k,v in nsw_top) if nsw_top else 'no data'
    off_tpl = (
        "Drug receipt measures the proportion of the sampled population reached by the MDA distributor. "
        + " ".join(off_tpl_lines)
        + f" Among those who received the drug but did not swallow it, the leading reasons were: {nsw_str}."
    )
    NARRATIVES['offered'] = ai_write('offered', '', off_tpl)

    comp_lines = [f"{d}: {v['pct']}% (n={v['n']})" for d, v in STATS['compliance'].items()]
    comp_tpl = (
        f"Among individuals who received the drug, compliance (swallowing) was: {'; '.join(comp_lines)}. "
        "High compliance rates indicate that refusal is a minor driver of overall coverage shortfall; "
        "the dominant gap is failure to reach households in the first place."
    )
    NARRATIVES['compliance'] = ai_write('compliance', '', comp_tpl)

    nt = STATS['never_treated']
    nt_tpl_lines = [
        f"In {d}, {v['pct']}% (95% CI {v['lo']}–{v['hi']}) of respondents aged 8+ answering for "
        f"themselves reported never having received {DRUG}. Among the never-treated, "
        f"{v['pct_female']}% were female with a median age of {v['age_median']} years."
        for d, v in nt.items()
    ]
    nt_tpl = (
        "A persistently never-treated subpopulation is the primary barrier to elimination. "
        + " ".join(nt_tpl_lines)
        + " Districts with elevated never-treated rates should be prioritised for targeted "
        "social mobilisation."
    )
    NARRATIVES['never_treated'] = ai_write('never_treated', '', nt_tpl)

    se_lines = [f"{d}: {v['pct']}% (n={v['n']})" for d, v in STATS['side_effects'].items()]
    top_se = list(STATS['side_effect_top_types'].items())[:4]
    top_se_str = ', '.join(f"{k} ({v})" for k,v in top_se) if top_se else 'no data'
    se_tpl = (
        f"Side effects were reported by: {'; '.join(se_lines)}. "
        f"The most commonly reported types were: {top_se_str}. "
        "These rates are within the expected range for this class of MDA."
    )
    NARRATIVES['side_effects'] = ai_write('side_effects', '', se_tpl)

    hd_lines = [f"{d}: {v['pct']}% (n={v['n']})" for d, v in STATS['heard'].items()]
    top_ch = list(STATS['heard_channels'].items())[:3]
    top_ch_str = ', '.join(f"{k} ({v})" for k,v in top_ch) if top_ch else 'no data'
    hd_tpl = (
        f"Pre-campaign communication reached: {'; '.join(hd_lines)} of respondents. "
        f"The strongest channels were: {top_ch_str}. "
        "Where awareness was lower, additional investment in social mobilisation ahead of the next round is warranted."
    )
    NARRATIVES['communication'] = ai_write('communication', '', hd_tpl)

    sat_lines = []
    for s in STATS['satisfaction']:
        pos = s.get('Satisfied', 0)
        neg = s.get('Not satisfied', 0)
        sat_lines.append(f"{s['district']}: {pos:.0f}% satisfied, {neg:.0f}% not satisfied")
    sat_tpl = (
        "Community satisfaction with the MDA campaign was generally high. "
        + "; ".join(sat_lines)
        + ". Elevated dissatisfaction warrants follow-up on CDD conduct, timing, or messaging."
    )
    NARRATIVES['satisfaction'] = ai_write('satisfaction', '', sat_tpl)

    rec_tpl = (
        "First, strengthen household reach in districts where the dominant gap is drug non-receipt. "
        "Supervision, retraining, and household coverage tracking are the most direct levers.\n\n"
        "Second, prioritise the never-treated subpopulation through targeted social mobilisation "
        "in sites with the highest never-treated rates. Persistent miss patterns undermine the "
        "elimination target even when overall coverage appears acceptable.\n\n"
        "Third, validate reported versus survey coverage where the two diverge significantly, "
        "and follow up with a Data Quality Assessment.\n\n"
        "Fourth, investigate the missed sites identified in this survey and confirm whether "
        "MDA actually took place or whether the site was inadvertently skipped.\n\n"
        "Fifth, sustain the strong compliance and satisfaction signals — high compliance among "
        "those reached is the foundation that further programmatic improvements can build on."
    )
    NARRATIVES['recommendations'] = ai_write('recommendations', '', rec_tpl)

    con_tpl = (
        f"The {REPORT_DATE} MDA round achieved overall epidemiological coverage of {ce['overall_pct']}% "
        f"and therapeutic coverage of {ct['overall_pct']}% across {len(districts)} surveyed districts. "
        f"{'Both' if ce['meets_threshold'] and ct['meets_threshold'] else 'Not all'} WHO thresholds "
        "were met at the overall level. The main programmatic lever is improved household-level reach, "
        "with secondary attention to the persistently never-treated population. Findings should be "
        "triangulated with reported coverage and disease prevalence data before final decisions on the next round."
    )
    NARRATIVES['conclusion'] = ai_write('conclusion', '', con_tpl)

    print(f"✅  Narratives ready: {list(NARRATIVES.keys())}")
    # ── HTML assembly ─────────────────────────────────────────────────────────────
    ce = STATS['coverage_epi']; ct = STATS['coverage_thera']; sa = STATS['sample']

    def embed_fig(key, caption='', width='100%'):
        if key not in FIGS: return ''
        return (f'<figure style="margin:24px 0;text-align:center">'
                f'<img src="data:image/png;base64,{FIGS[key]}" '
                f'style="width:{width};max-width:900px;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.15)">'
                f'<figcaption style="margin-top:8px;font-style:italic;color:#555;font-size:13px">{caption}</figcaption>'
                f'</figure>')

    def verdict_banner(meets_epi, meets_thera, t_epi, t_thera):
        if meets_epi and meets_thera:
            msg = f"BOTH WHO thresholds met (epi {t_epi}%, therapeutic {t_thera}%)"
            bg = '#0DCF00'; col = '#0a3300'
        elif meets_epi or meets_thera:
            msg = (f"Epi threshold ({t_epi}%) MET — therapeutic NOT met" if meets_epi
                   else f"Therapeutic threshold ({t_thera}%) MET — epi NOT met")
            bg = '#F7B500'; col = '#3a2700'
        else:
            msg = f"NEITHER WHO threshold met (epi {t_epi}%, therapeutic {t_thera}%)"
            bg = '#FF4444'; col = 'white'
        return f'<div style="background:{bg};color:{col};padding:14px 20px;border-radius:8px;font-weight:bold;font-size:16px;margin:16px 0">{msg}</div>'

    def headline_banner(text, status):
        if status is True:   bg, col = '#0DCF00', '#0a3300'
        elif status is False: bg, col = '#FF4444', 'white'
        else:                 bg, col = '#e8ecf0', '#1a1a1a'
        return (f'<div style="background:{bg};color:{col};padding:16px 22px;border-radius:8px;'
                f'font-weight:700;font-size:18px;margin:10px 0;letter-spacing:0.2px;line-height:1.35">{text}</div>')

    def narrative_block(text):
        paras = [p.strip() for p in text.split('\n') if p.strip()]
        return ''.join(f'<p style="margin-bottom:14px;line-height:1.7">{p}</p>' for p in paras)

    scorecard_items = (
        scorecard_metric('Epi Coverage', f"{ce['overall_pct']}%", 'good' if ce['meets_threshold'] else 'low') +
        scorecard_metric('Epi 95% CI', f"{ce['lower_95']}%–{ce['upper_95']}%", 'neutral') +
        scorecard_metric(f'WHO Epi Target ({ce["threshold"]}%)', 'MET' if ce['meets_threshold'] else 'NOT MET',
                         'good' if ce['meets_threshold'] else 'low') +
        scorecard_metric('Therapeutic Coverage', f"{ct['overall_pct']}%", 'good' if ct['meets_threshold'] else 'low') +
        scorecard_metric('Therapeutic 95% CI', f"{ct['lower_95']}%–{ct['upper_95']}%", 'neutral') +
        scorecard_metric(f'WHO Thera Target ({ct["threshold"]}%)', 'MET' if ct['meets_threshold'] else 'NOT MET',
                         'good' if ct['meets_threshold'] else 'low') +
        scorecard_metric('Districts Meeting Epi', f"{ce['n_good_districts']} / {len(districts)}",
                         'good' if ce['n_low_districts']==0 else 'warn') +
        scorecard_metric('Interviews', f"{sa['n_total']:,}", 'neutral') +
        scorecard_metric('Districts', str(len(districts)), 'neutral') +
        scorecard_metric('Sites (clusters)', str(sa['n_sites']), 'neutral') +
        scorecard_metric('Design Effect (epi)', str(ce['deff']), 'neutral')
    )

    # Admin breakdown table — district/IU → sites
    _ab = STATS['admin_breakdown']
    _admin_rows = ''
    for d in districts:
        info = _ab['by_district'][d]
        _admin_rows += (f'<tr><td style="text-align:left">{d}</td>'
                        f'<td>{info["n_sites"]}</td>'
                        f'<td>{info["n_interviews"]:,}</td></tr>')
    _admin_rows += (f'<tr style="font-weight:bold;background:#f0f0f0">'
                    f'<td style="text-align:left">All districts ({len(districts)})</td>'
                    f'<td>{_ab["overall"]["sites"]}</td>'
                    f'<td>{sa["n_total"]:,}</td></tr>')
    admin_breakdown_html = f"""
    <h3 class="subsection" style="margin-top:18px">Administrative-unit breakdown</h3>
    <p class="table-caption" style="margin:4px 0 8px">
      Table. Sampled administrative units by district/IU.
    </p>
    <table class="ces-table" style="max-width:560px">
      <tr><th style="text-align:left">District / IU</th><th>Sites (clusters)</th><th>Interviews</th></tr>
      {_admin_rows}
    </table>"""

    # District coverage table
    iu_rows = ''
    for d in districts:
        e = ce['by_district'][d]; t = ct['by_district'][d]
        bg_e = '#d4f7d4' if e['meets'] else '#ffd4d4'
        bg_t = '#d4f7d4' if t['meets'] else '#ffd4d4'
        geo = STATS['coverage_geo'][d]
        iu_rows += f'''<tr>
          <td>{d}</td><td>{sa["by_district"][d]:,}</td><td>{geo['n_villages']}</td>
          <td>{e['pct']}%</td><td>{e['lo']}%–{e['hi']}%</td><td>{e.get('deff','—')}</td>
          <td style="background:{bg_e}"><span class="meet-txt">{"MET" if e['meets'] else "NOT MET"}</span></td>
          <td>{t['pct']}%</td><td>{t['lo']}%–{t['hi']}%</td>
          <td style="background:{bg_t}"><span class="meet-txt">{"MET" if t['meets'] else "NOT MET"}</span></td>
          <td>{geo['n_treated']}/{geo['n_villages']} ({geo['pct']}%)</td></tr>'''
    iu_rows += f'''<tr style="font-weight:bold;background:#f0f0f0">
      <td>OVERALL</td><td>{sa["n_total"]:,}</td><td>{sa["n_sites"]}</td>
      <td>{ce['overall_pct']}%</td><td>{ce['lower_95']}%–{ce['upper_95']}%</td><td>{ce.get('deff','—')}</td>
      <td style="background:{'#d4f7d4' if ce['meets_threshold'] else '#ffd4d4'}">{"MET" if ce['meets_threshold'] else "NOT MET"}</td>
      <td>{ct['overall_pct']}%</td><td>{ct['lower_95']}%–{ct['upper_95']}%</td>
      <td style="background:{'#d4f7d4' if ct['meets_threshold'] else '#ffd4d4'}">{"MET" if ct['meets_threshold'] else "NOT MET"}</td>
      <td>—</td></tr>'''

    # Received table
    off_rows = ''
    for d, v in STATS['offered'].items():
        off_rows += f'<tr><td>{d}</td><td>{v["n"]:,}</td><td>{v["pct_yes"]}%</td><td>{v["pct_no"]}%</td></tr>'

    # Compliance table
    comp_rows = ''
    for d in districts:
        cp = STATS['compliance'].get(d, {'pct':'—','n':'—'})
        se = STATS['side_effects'].get(d, {'pct':'—','n':'—'})
        hd = STATS['heard'].get(d, {'pct':'—','n':'—'})
        comp_rows += f'<tr><td>{d}</td><td>{cp["pct"]}%</td><td>{cp["n"]}</td><td>{se["pct"]}%</td><td>{hd["pct"]}%</td></tr>'

    # Never-treated table
    nt_rows_html = ''
    for d, v in STATS['never_treated'].items():
        nt_rows_html += f'<tr><td>{d}</td><td>{v["n"]:,}</td><td>{v["n_never"]}</td><td>{v["pct"]}% ({v["lo"]}–{v["hi"]}%)</td><td>{v["pct_female"]}%</td><td>{v["age_median"]}</td></tr>'

    # Missed sites list
    missed_html = ''
    if STATS['missed_villages']:
        missed_html = '<ul style="font-family:Arial,sans-serif;font-size:14px;margin:10px 0 0 20px">' + \
            ''.join(f"<li><strong>{m['site']}</strong> ({m['district']}) — {m['n_interviews']} interviews, 0 treated</li>"
                    for m in STATS['missed_villages']) + '</ul>'

    def kv_list(d):
        if not d: return '<em>No data</em>'
        return ('<ul style="font-family:Arial,sans-serif;font-size:14px">' +
                ''.join(f"<li>{k}: {v}</li>" for k, v in d.items()) + '</ul>')

    ai_badge = '<span class="badge">✍️ Template Narrative</span>'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CES Report — {COUNTRY} {REPORT_DATE}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:Georgia,'Times New Roman',serif;color:#1a1a1a;background:#f8f8f8;font-size:15px}}
  .page{{max-width:960px;margin:0 auto;background:#fff;padding:0}}
  .cover{{background:linear-gradient(135deg,#0b2a5e 0%,#1a4dbd 60%,#2e6fe0 100%);color:#fff;padding:80px 60px;min-height:340px}}
  .cover h1{{font-size:32px;line-height:1.3;margin-bottom:16px;font-weight:700}}
  .cover h2{{font-size:20px;font-weight:400;opacity:.85;margin-bottom:32px}}
  .badge-row{{display:flex;flex-wrap:wrap;gap:10px;margin-top:24px}}
  .badge{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.35);border-radius:20px;padding:6px 16px;font-size:13px;font-weight:600}}
  .section{{padding:40px 60px;border-bottom:1px solid #e8e8e8}}
  .section:last-child{{border-bottom:none}}
  h2.section-title{{font-size:22px;color:#0b2a5e;border-left:4px solid #1a4dbd;padding-left:14px;margin-bottom:20px;font-family:Arial,sans-serif}}
  h3.subsection{{font-size:17px;color:#1a4dbd;margin:28px 0 12px;font-family:Arial,sans-serif}}
  p{{margin-bottom:14px;line-height:1.75;color:#2a2a2a}}
  .scorecard{{display:flex;flex-wrap:wrap;gap:12px;margin:20px 0}}
  .scorecard-item{{flex:1 1 160px;border-radius:10px;padding:16px;background:#e8ecf0;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
  .sc-label{{font-size:12px;font-weight:600;font-family:Arial,sans-serif;text-transform:uppercase;letter-spacing:.5px;opacity:.8;margin-bottom:6px}}
  .sc-value{{font-size:24px;font-weight:700;font-family:Arial,sans-serif}}
  .ces-table{{border-collapse:collapse;width:100%;font-size:13px;font-family:Arial,sans-serif;margin:16px 0}}
  .ces-table th{{background:#1a4080;color:#fff;padding:9px 12px;text-align:center;font-weight:600}}
  .ces-table td{{padding:8px 12px;text-align:center;border-bottom:1px solid #e0e0e0}}
  .ces-table tr:nth-child(even){{background:#f5f7fa}}
  .ces-table tr:last-child td{{font-weight:bold;background:#f0f0f0}}
  .table-caption{{font-style:italic;color:#555;font-size:13px;margin-bottom:6px}}
  .toc{{background:#f0f4ff;border-radius:8px;padding:20px 28px;margin:20px 0}}
  .toc h3{{color:#0b2a5e;margin-bottom:10px;font-family:Arial,sans-serif}}
  .toc a{{color:#1a4dbd;text-decoration:none;line-height:2}}
  .toc a:hover{{text-decoration:underline}}
  .appendix-table{{width:100%;border-collapse:collapse;font-size:13px;font-family:Arial,sans-serif;margin:12px 0}}
  .appendix-table th{{padding:8px;text-align:left;color:#fff;font-weight:600}}
  .appendix-table td{{padding:8px;border:1px solid #ddd;vertical-align:top}}
  .callout{{background:#fff3cd;border-left:4px solid #F7B500;padding:14px 18px;border-radius:6px;margin:16px 0;font-family:Arial,sans-serif;font-size:14px}}
  .callout-red{{background:#ffe4e4;border-left-color:#FF4444}}
  .callout-info{{background:#e8f4fd;border-left-color:#1a4dbd}}
  .meet-txt{{font-weight:bold;font-size:12px;letter-spacing:.3px}}
  @page{{size:A4;margin:1.6cm 1.4cm}}
  @media print{{
    body{{background:#fff;font-size:13px}}
    .page{{max-width:100%;box-shadow:none}}
    .cover{{padding:28px 32px;min-height:auto}}
    .section{{padding:18px 22px;page-break-inside:avoid}}
    h2.section-title{{font-size:17px;margin-bottom:12px}}
    .ces-table{{font-size:10.5px}}
    .scorecard-item{{flex:1 1 110px;padding:10px}}
    .sc-value{{font-size:18px}}
  }}
</style>
</head>
<body>
<div class="page">

<!-- COVER -->
<div class="cover">
  <h1>Coverage Evaluation Survey (CES)<br>Post-MDA Independent Coverage Report</h1>
  <h2>WHO/ESPEN Cluster-Sample Analysis</h2>
  <div class="badge-row">
    <span class="badge">🌍 {COUNTRY}</span>
    <span class="badge">🦠 {DISEASE}</span>
    <span class="badge">💊 {DRUG}</span>
    <span class="badge">📅 {REPORT_DATE}</span>
    <span class="badge">📋 {len(districts)} district{'s' if len(districts)!=1 else ''}</span>
    <span class="badge">{'EPI: THRESHOLD MET' if ce['meets_threshold'] else 'EPI: BELOW THRESHOLD'}</span>
    <span class="badge">{'THERA: THRESHOLD MET' if ct['meets_threshold'] else 'THERA: BELOW THRESHOLD'}</span>
    {ai_badge}
  </div>
  <p style="margin-top:28px;font-size:13px;opacity:.75;color:white;">
    Generated by CSAT v10 · ESPEN data format · Logit CI method · Cluster-robust variance
  </p>
</div>

<!-- TOC -->
<div class="section">
  <div class="toc">
    <h3>Table of Contents</h3>
    <div><a href="#executive-summary">1. Executive Summary</a></div>
    <div><a href="#methodology">2. Methodology</a></div>
    <div><a href="#dataset">3. Dataset Overview &amp; Age Breakdown</a></div>
    <div><a href="#coverage">4. Coverage Results — Epidemiological &amp; Therapeutic</a></div>
    <div><a href="#geographic">5. Spatial Coverage</a></div>
    <div><a href="#mda-location">6. MDA Delivery Location</a></div>
    <div><a href="#offered">7. Drug Receipt and Compliance Barriers</a></div>
    <div><a href="#compliance">8. Compliance Among Reached</a></div>
    <div><a href="#never-treated">9. Never-Treated Population</a></div>
    <div><a href="#side-effects">10. Side Effects</a></div>
    <div><a href="#communication">11. Communication Reach</a></div>
    <div><a href="#satisfaction">12. Community Satisfaction</a></div>
    <div><a href="#recommendations">13. Recommended Next Steps</a></div>
    <div><a href="#conclusion">14. Conclusion</a></div>
    <div><a href="#appendix">Appendix A — WHO Interpretation Framework</a></div>
  </div>
</div>

<!-- 1. EXECUTIVE SUMMARY -->
<div id="executive-summary" class="section">
  <h2 class="section-title">1. Executive Summary</h2>
  {headline_banner(STATS['headlines']['effective']['text'], STATS['headlines']['effective']['all_ok'])}
  {headline_banner(STATS['headlines']['alignment']['text'], STATS['headlines']['alignment']['all_ok'])}
  {verdict_banner(ce['meets_threshold'], ct['meets_threshold'], ce['threshold'], ct['threshold'])}
  <div class="scorecard">{scorecard_items}</div>
  {admin_breakdown_html}
  {narrative_block(NARRATIVES['executive_summary'])}
</div>

<!-- 2. METHODOLOGY -->
<div id="methodology" class="section">
  <h2 class="section-title">2. Methodology</h2>
  {narrative_block(NARRATIVES['methodology'])}
  <h3 class="subsection">Treatment Pathway</h3>
  {narrative_block(NARRATIVES['cascade'])}
  {embed_fig('fig_cascade',
      'Figure. Treatment pathway cascade (ESPEN). Total sampled → Received drug → Swallowed. '
      'Arrows show percentage-point and absolute drop-off at each step. '
      'Note: CDD visit and respondent-presence steps are not collected in ESPEN.')}
</div>

<!-- 3. DATASET -->
<div id="dataset" class="section">
  <h2 class="section-title">3. Dataset Overview &amp; Age Breakdown</h2>
  {narrative_block(NARRATIVES['dataset'])}
  <table class="ces-table" style="max-width:600px">
    <tr><th>Parameter</th><th>Value</th></tr>
    <tr><td>Country</td><td>{COUNTRY}</td></tr>
    <tr><td>Disease</td><td>{DISEASE}</td></tr>
    <tr><td>Drug</td><td>{DRUG}</td></tr>
    <tr><td>MDA Round</td><td>{MDA_ROUND}</td></tr>
    <tr><td>Survey Date</td><td>{REPORT_DATE}</td></tr>
    <tr><td>Total Interviews</td><td>{sa['n_total']:,}</td></tr>
    <tr><td>Districts</td><td>{', '.join(districts)}</td></tr>
    <tr><td>Sites (clusters)</td><td>{sa['n_sites']}</td></tr>
    <tr><td>Female respondents</td><td>{sa['n_female']:,} ({sa['pct_female']}%)</td></tr>
    <tr><td>Median age</td><td>{sa['age_median']} years</td></tr>
  </table>
  <h3 class="subsection">Age Breakdown of Sampled Population</h3>
  {TABLES['age_breakdown']}
  {embed_fig('fig_age_breakdown', f'Figure. Age breakdown by district. Red dashed = DHS u5 reference ({DHS_U5_PCT}%).')}
</div>

<!-- 4. COVERAGE -->
<div id="coverage" class="section">
  <h2 class="section-title">4. Coverage Results — Epidemiological &amp; Therapeutic</h2>
  {narrative_block(NARRATIVES['coverage'])}
  <h3 class="subsection">Figure — Coverage by District</h3>
  {embed_fig('fig_coverage_district', f'Figure. Epi (left, WHO {THRESH_EPI}%) and therapeutic (right, WHO {THRESH_THERA}%) coverage by district, 95% CI.')}
  <h3 class="subsection">Table — District-Level Coverage Results</h3>
  <table class="ces-table">
    <tr>
      <th rowspan="2">District / IU</th><th rowspan="2">Interviews</th><th rowspan="2">Sites</th>
      <th colspan="4">Epidemiological (WHO {THRESH_EPI}%)</th>
      <th colspan="3">Therapeutic (WHO {THRESH_THERA}%)</th>
      <th rowspan="2">Geographic</th>
    </tr>
    <tr>
      <th>Coverage</th><th>95% CI</th><th>DEFF</th><th>Meets?</th>
      <th>Coverage</th><th>95% CI</th><th>Meets?</th>
    </tr>
    {iu_rows}
  </table>
  <h3 class="subsection">Figure — Survey Coverage by Site</h3>
  {''.join(embed_fig('fig_village_'+d, f'Figure. Epi coverage by site — {d}. Green = meets WHO target; red = below. Wilson 95% CI shown.')
           for d in districts if ('fig_village_'+d) in FIGS)}
</div>

<!-- 5. SPATIAL COVERAGE -->
<div id="geographic" class="section">
  <h2 class="section-title">5. Spatial Coverage</h2>
  {narrative_block(NARRATIVES['geographic'])}
  {('<div class="callout callout-red"><strong>Sites with zero treated respondents:</strong>'
    + missed_html + '</div>') if STATS['missed_villages'] else ''}
</div>

<!-- 6. MDA DELIVERY LOCATION -->
<div id="mda-location" class="section">
  <h2 class="section-title">6. MDA Delivery Location</h2>
  <div class="callout callout-info"><strong>ESPEN:</strong> The ESPEN instrument records
  where the drug was distributed — at the household, a fixed post, or a school.</div>
  {embed_fig('fig_mda_location', 'Figure. MDA delivery location breakdown by district (% of treated respondents).')}
</div>

<!-- 7. DRUG RECEIPT AND COMPLIANCE BARRIERS -->
<div id="offered" class="section">
  <h2 class="section-title">7. Drug Receipt and Compliance Barriers</h2>
  {narrative_block(NARRATIVES['offered'])}
  {embed_fig('fig_offered', 'Figure. % of sampled respondents who received the drug, by district.')}
  <table class="ces-table" style="max-width:500px">
    <tr><th>District</th><th>n sampled</th><th>% Received</th><th>% Not received</th></tr>
    {off_rows}
  </table>
  <h3 class="subsection">Reasons for Not Swallowing (among those who received)</h3>
  <div class="callout"><strong>Note:</strong> ESPEN records reasons for not swallowing among those
  who received the drug. Reasons for not being reached at all (CDD not visiting / respondent absent)
  are not collected in the ESPEN instrument — this is a key data gap flagged for review.</div>
  {embed_fig('fig_reasons_not_offered', f'Figure. Top reasons for not swallowing {DRUG} among those who received it.')}
  {kv_list(STATS['reasons_not_swallowed'])}
</div>

<!-- 8. COMPLIANCE -->
<div id="compliance" class="section">
  <h2 class="section-title">8. Compliance Among Reached</h2>
  {narrative_block(NARRATIVES['compliance'])}
  <table class="ces-table" style="max-width:700px">
    <tr><th>District</th><th>% Swallowed (if received)</th><th>n received</th>
        <th>% Side effects</th><th>% Heard pre-MDA</th></tr>
    {comp_rows}
  </table>
</div>

<!-- 9. NEVER TREATED -->
<div id="never-treated" class="section">
  <h2 class="section-title">9. Never-Treated Population</h2>
  <div class="callout"><strong>Why this matters:</strong> The truly never-treated subpopulation
  is the primary barrier to elimination. Even programmes meeting WHO thresholds
  can fail if a persistent core never receives treatment.</div>
  {narrative_block(NARRATIVES['never_treated'])}
  {embed_fig('fig_never_treated', 'Figure. % never treated (aged 8+, self-respondents) by site, grouped by district.')}
  <table class="ces-table" style="max-width:800px">
    <tr><th>District</th><th>Eligible respondents</th><th>n never treated</th>
        <th>% (95% CI)</th><th>% female (of never)</th><th>Median age (of never)</th></tr>
    {nt_rows_html}
  </table>
</div>

<!-- 10. SIDE EFFECTS -->
<div id="side-effects" class="section">
  <h2 class="section-title">10. Side Effects</h2>
  {narrative_block(NARRATIVES['side_effects'])}
  {embed_fig('fig_side_effects', 'Figure. % reporting any side effect, by district.')}
  <h3 class="subsection">Top reported side-effect types (overall)</h3>
  {kv_list(STATS['side_effect_top_types'])}
</div>

<!-- 11. COMMUNICATION -->
<div id="communication" class="section">
  <h2 class="section-title">11. Communication Reach</h2>
  {narrative_block(NARRATIVES['communication'])}
  {embed_fig('fig_heard', 'Figure. % who heard about the campaign before MDA started.')}
  <h3 class="subsection">Top information channels</h3>
  {kv_list(STATS['heard_channels'])}
</div>

<!-- 12. SATISFACTION -->
<div id="satisfaction" class="section">
  <h2 class="section-title">12. Community Satisfaction</h2>
  {narrative_block(NARRATIVES['satisfaction'])}
  {embed_fig('fig_satisfaction', f'Figure. Respondent satisfaction with the MDA — {COUNTRY}, {REPORT_DATE}. ESPEN uses Yes / No / Do not know.')}
</div>

<!-- 13. RECOMMENDATIONS -->
<div id="recommendations" class="section">
  <h2 class="section-title">13. Recommended Next Steps</h2>
  {narrative_block(NARRATIVES['recommendations'])}
</div>

<!-- 14. CONCLUSION -->
<div id="conclusion" class="section">
  <h2 class="section-title">14. Conclusion</h2>
  {narrative_block(NARRATIVES['conclusion'])}
</div>

<!-- APPENDIX -->
<div id="appendix" class="section">
  <h2 class="section-title">Appendix A — WHO Interpretation Framework</h2>
  <p style="font-style:italic;margin-bottom:16px">
    Adapted from: WHO Coverage Evaluation Surveys for Preventive Chemotherapy — Field Guide (2016)
  </p>
  <h3 class="subsection">1. Survey Coverage vs. WHO Target Threshold</h3>
  <table class="appendix-table">
    <tr style="background:#6a3d9a"><th>Finding</th><th>Potential Causes</th><th>Corrective Action</th></tr>
    <tr><td>Survey coverage <strong>below</strong> threshold</td>
        <td>Check sub-population coverage; check reasons for non-receipt and non-swallowing</td>
        <td>Targeted mobilisation; strengthen delivery; consider mop-up MDA</td></tr>
    <tr><td>Survey coverage <strong>above</strong> threshold</td>
        <td>Programme functioning well</td>
        <td>Sustain momentum for next MDA round</td></tr>
  </table>
  <h3 class="subsection">2. Survey Coverage vs. Reported Coverage</h3>
  <table class="appendix-table">
    <tr style="background:#c0392b"><th>Finding</th><th>Potential Causes</th><th>Corrective Action</th></tr>
    <tr><td>Reported &gt;&gt; Survey (&gt;25 pp)</td>
        <td>Distributor mis-reporting; incorrect population denominator</td>
        <td>Data Quality Assessment; improve distributor training; review tally sheets</td></tr>
    <tr><td>Reported &lt;&lt; Survey (&gt;25 pp)</td>
        <td>Incorrect denominator; data not aggregated correctly</td>
        <td>Locate more accurate population estimates; improve data aggregation</td></tr>
    <tr><td>Similar (≤10 pp)</td>
        <td>Reporting system functioning well</td>
        <td>Continue current system</td></tr>
  </table>
  <h3 class="subsection">3. ESPEN Data Gaps (v10 notes)</h3>
  <table class="appendix-table">
    <tr style="background:#555"><th>Missing element</th><th>Previous dataset</th><th>Impact on report</th></tr>
    <tr><td>Sub-district / admin3</td><td>admin3 column</td><td>Sub-district section replaced by MDA Delivery Location (p_mda_location).</td></tr>
    <tr><td>CDD visit flag</td><td>cddvisit_ov</td><td>Cascade simplified to 3 steps (Total → Received → Swallowed).</td></tr>
    <tr><td>Absent flag</td><td>absent_ov</td><td>Cannot distinguish "CDD didn't come" vs "respondent was absent".</td></tr>
  </table>
</div>

<!-- FOOTER -->
<div style="background:#0b2a5e;color:#aac;padding:20px 60px;font-size:12px;font-family:Arial,sans-serif">
  <p>CSAT v10 · {COUNTRY} · {DISEASE} · {DRUG} · {REPORT_DATE} · ESPEN Data Format · WHO/ESPEN Logit CI Method</p>
</div>

</div>
</body>
</html>"""

    return html
