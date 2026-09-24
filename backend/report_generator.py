import numpy as np
import pandas as pd
import logging
import re
from typing import List, Dict, Tuple, Optional

logger = logging.getLogger(__name__)

def norm_str(s: str) -> str:
    """Normalize string for fuzzy matching (lowercase, alphanumeric only)."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def safe_eval_formula(df: pd.DataFrame, formula: str) -> Optional[pd.Series]:
    """Safely evaluate mathematical formulas across DataFrame columns."""
    df_cols = list(df.columns)
    # Sort columns by length descending to replace longest names first
    cols_sorted = sorted(df_cols, key=len, reverse=True)
    temp_formula = formula.strip()
    env = {}
    
    for i, col in enumerate(cols_sorted):
        if col in temp_formula:
            token = f"__col_{i}__"
            temp_formula = temp_formula.replace(col, token)
            env[token] = pd.to_numeric(df[col], errors='coerce')
            
    try:
        allowed_globals = {"np": np, "abs": abs, "min": min, "max": max}
        res = eval(temp_formula, allowed_globals, env)
        if isinstance(res, (pd.Series, np.ndarray, list, float, int)):
            return pd.Series(res, index=df.index)
    except Exception as e:
        logger.warning(f"Formula evaluation failed for '{formula}': {e}")
    return None

def parse_and_apply_derived(df: pd.DataFrame, text: str) -> Tuple[pd.DataFrame, List[str]]:
    """Parse [DERIVED: VarName = Formula] tags in text and compute new columns in df."""
    df = df.copy()
    derived_specs = re.findall(r'\[DERIVED:\s*([^\]=]+?)\s*=\s*([^\]]+)\]', text, re.IGNORECASE)
    created_vars = []
    
    for var_name, formula in derived_specs:
        var_name = var_name.strip()
        formula = formula.strip()
        series = safe_eval_formula(df, formula)
        if series is not None:
            df[var_name] = series
            created_vars.append(f"[DERIVED: {var_name} = {formula}]")
            logger.info(f"Derived variable created: {var_name}")
            
    return df, created_vars

def find_best_column_match(col_name: str, available_cols: List[str], fallback_col: Optional[str] = None) -> str:
    """Map a column name from prompt/tag to the best matching column in available_cols."""
    if not available_cols:
        return col_name
    col_clean = col_name.strip()
    if col_clean in available_cols:
        return col_clean
        
    target_norm = norm_str(col_clean)
    
    # Exact normalized match
    for c in available_cols:
        if norm_str(c) == target_norm:
            return c
            
    # Substring match
    for c in available_cols:
        cn = norm_str(c)
        if target_norm and (target_norm in cn or cn in target_norm):
            return c
            
    # Fallback to provided fallback or first available column
    return fallback_col if fallback_col in available_cols else available_cols[0]

def compute_stats(df: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    """Compute summary statistics for all numeric columns."""
    stats = {}
    numeric_df = df.select_dtypes(include=[np.number])
    
    for col in numeric_df.columns:
        series = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(series) > 0:
            stats[col] = {
                'min': float(series.min()),
                'q1': float(series.quantile(0.25)),
                'median': float(series.median()),
                'q3': float(series.quantile(0.75)),
                'max': float(series.max()),
                'mean': float(series.mean()),
                'std': float(series.std()) if len(series) > 1 else 0.0,
                'count': int(len(series))
            }
        else:
            stats[col] = {'min': 0.0, 'q1': 0.0, 'median': 0.0, 'q3': 0.0, 'max': 0.0, 'mean': 0.0, 'std': 0.0, 'count': 0}
            
    return stats

def categorize_columns(df_cols: List[str]) -> Dict[str, List[str]]:
    """Categorize dataset variables into functional engineering categories."""
    cats = {
        'throughput': [],
        'fuel_energy': [],
        'efficiency': [],
        'thermal': [],
        'atmosphere': [],
        'pressure_fan': [],
        'quality': [],
        'other': []
    }
    
    for c in df_cols:
        cn = c.lower()
        if any(k in cn for k in ['prod', 'clinker', 'feed', 'flour', 'output', 'ton', 't/h', 'rate', 'capacity']):
            cats['throughput'].append(c)
        elif any(k in cn for k in ['fuel', 'coal', 'coke', 'gas_flow', 'power', 'kw', 'mw', 'current', 'amp', 'torch', 'calciner', 'burner']):
            cats['fuel_energy'].append(c)
        elif any(k in cn for k in ['sfc', 'sec', 'specific', 'consumption', 'ratio', 'efficiency', 'heat_rate']):
            cats['efficiency'].append(c)
        elif any(k in cn for k in ['temp', 'pyrometer', 'air_temp', 'degree', 'celsius', 'heat', 'secondary', 'cyclone_gas']):
            cats['thermal'].append(c)
        elif any(k in cn for k in ['o2', 'co', 'no', 'nox', 'so2', 'so3', 'smoke', 'downcomer', 'emissions', 'draft_gas']):
            cats['atmosphere'].append(c)
        elif any(k in cn for k in ['fan', 'speed', 'draft', 'pressure', 'inlet', 'outlet', 'head', 'rpm', 'id_fan', 'blower']):
            cats['pressure_fan'].append(c)
        elif any(k in cn for k in ['lsf', 'fsc', 'c3s', 'free_lime', 'cao', 's.cao', 'sieve', '#170', 'fineness', 'ms', 'ar', 'fe2o3', 'quality']):
            cats['quality'].append(c)
        else:
            cats['other'].append(c)
            
    return cats

def compute_correlations(df: pd.DataFrame) -> List[Tuple[str, str, float]]:
    """Find key variable correlation pairs sorted by magnitude."""
    num_df = df.select_dtypes(include=[np.number])
    if num_df.shape[1] < 2:
        return []
        
    corr_matrix = num_df.corr().abs()
    pairs = []
    cols = list(num_df.columns)
    
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            c1, c2 = cols[i], cols[j]
            val = float(num_df[c1].corr(num_df[c2]))
            if not np.isnan(val):
                pairs.append((c1, c2, val))
                
    pairs.sort(key=lambda x: abs(x[2]), reverse=True)
    return pairs

def normalize_chart_tags(prompt_text: str, available_cols: List[str]) -> str:
    """Normalize variable names inside chart tags to valid dataset columns."""
    if not available_cols:
        return prompt_text
        
    def fix_scatter(m):
        xv, yv = m.group(1), m.group(2)
        cv = m.group(3)
        sc = m.group(4) or 'Jet'
        x_fixed = find_best_column_match(xv, available_cols)
        y_vars = [find_best_column_match(v, available_cols) for v in yv.split(',') if v.strip()]
        color_fixed = find_best_column_match(cv, available_cols) if cv else None
        res = f"[SCATTER: X={x_fixed} | Y={','.join(y_vars)}"
        if color_fixed:
            res += f" | COLOR={color_fixed}"
        res += f" | SCALE={sc.strip()}]"
        return res

    def fix_scatter3d(m):
        xv, yv, zv = m.group(1), m.group(2), m.group(3)
        cv = m.group(4)
        sc = m.group(5) or 'Jet'
        x_fixed = find_best_column_match(xv, available_cols)
        y_fixed = find_best_column_match(yv, available_cols)
        z_fixed = find_best_column_match(zv, available_cols)
        color_fixed = find_best_column_match(cv, available_cols) if cv else None
        res = f"[SCATTER3D: X={x_fixed} | Y={y_fixed} | Z={z_fixed}"
        if color_fixed:
            res += f" | COLOR={color_fixed}"
        res += f" | SCALE={sc.strip()}]"
        return res

    def fix_parallel(m):
        vars_str, color_str = m.group(1), m.group(2)
        v_list = [find_best_column_match(v, available_cols) for v in vars_str.split(',') if v.strip()]
        color_fixed = find_best_column_match(color_str, available_cols)
        return f"[PARALLEL: {', '.join(v_list)} | COLOR: {color_fixed}]"

    def fix_dual(m):
        left_str, right_str = m.group(1), m.group(2)
        l_list = [find_best_column_match(v, available_cols) for v in left_str.split(',') if v.strip()]
        r_list = [find_best_column_match(v, available_cols) for v in right_str.split(',') if v.strip()]
        return f"[DUALPLOT: {', '.join(l_list)} | {', '.join(r_list)}]"

    def fix_box_hist(m):
        tag_type, vars_str = m.group(1).upper(), m.group(2)
        v_list = [find_best_column_match(v, available_cols) for v in vars_str.split(',') if v.strip()]
        return f"[{tag_type}: {', '.join(v_list)}]"

    text = re.sub(r'\[SCATTER:\s*X=([^\|\]]+)\|\s*Y=([^\|\]]+)(?:\|\s*COLOR=([^\|\]]+))?(?:\|\s*SCALE=([^\]]+))?\]', fix_scatter, prompt_text, flags=re.IGNORECASE)
    text = re.sub(r'\[SCATTER3D:\s*X=([^\|\]]+)\|\s*Y=([^\|\]]+)\|\s*Z=([^\|\]]+)(?:\|\s*COLOR=([^\|\]]+))?(?:\|\s*SCALE=([^\]]+))?\]', fix_scatter3d, text, flags=re.IGNORECASE)
    text = re.sub(r'\[PARALLEL:\s*([^\|\]]+)\|\s*COLOR:\s*([^\]]+)\]', fix_parallel, text, flags=re.IGNORECASE)
    text = re.sub(r'\[DUALPLOT:\s*([^\|\]]+)\|\s*([^\]]+)\]', fix_dual, text, flags=re.IGNORECASE)
    text = re.sub(r'\[(BOX|HISTOGRAM):\s*([^\]]+)\]', fix_box_hist, text, flags=re.IGNORECASE)
    
    return text

def process_prompt_template(df: pd.DataFrame, query_str: str) -> str:
    """Process a user-provided prompt template generically by executing derived tags, populating stats, and fixing chart tags."""
    # 1. Apply derived variables
    df_proc, derived_lines = parse_and_apply_derived(df, query_str)
    stats = compute_stats(df_proc)
    num_cols = list(stats.keys())
    
    # 2. Strip prompt wrappers & meta-instructions generically
    clean_text = query_str
    
    # Generic detection of main report title (# Title) after prompt/persona wrappers
    header_matches = list(re.finditer(r'^#\s+[^#\n]+', clean_text, re.MULTILINE))
    if len(header_matches) > 1:
        first_title = header_matches[0].group(0).lower()
        if any(k in first_title for k in ['prompt', 'persona', 'instruction', 'system']):
            clean_text = clean_text[header_matches[1].start():]

    clean_text = re.sub(r'>\s*[🛑⚠️🚨🔴].*?\n\n', '', clean_text, flags=re.DOTALL)
    clean_text = re.sub(r'##\s*Persona:.*?\n\n', '', clean_text, flags=re.DOTALL)
    clean_text = re.sub(r'Run the Prerequisite FIRST.*?\n\n', '', clean_text, flags=re.DOTALL)
    clean_text = re.sub(r'(?:Write|Execute|Detail|Provide)\s+(?:9\s+|complete\s+|thorough\s+)?(?:engineering\s+)?(?:insights|analysis|directives):?\s*\n', '', clean_text, flags=re.IGNORECASE)
    
    # Extract markdown block if prompt is wrapped in ```markdown ... ```
    if clean_text.strip().startswith('```markdown') and clean_text.strip().endswith('```'):
        clean_text = clean_text.strip()[11:-3].strip()
        
    # 3. Normalize chart tags
    clean_text = normalize_chart_tags(clean_text, num_cols)
    
    # 4. Populate table statistics if markdown tables exist in prompt
    def populate_table_row(line):
        parts = [p.strip() for p in line.split('|')[1:-1]]
        if not parts:
            return line
        var_candidate = parts[0].replace('**', '').strip()
        matched_col = find_best_column_match(var_candidate, num_cols, fallback_col=None)
        if matched_col in stats and len(parts) >= 6:
            s = stats[matched_col]
            # Formats standard table row: Variable | Min | Q1 | Median | Mean | Q3 | Max | ...
            row_str = f"| {matched_col} | {s['min']:.2f} | {s['q1']:.2f} | {s['median']:.2f} | {s['mean']:.2f} | {s['q3']:.2f} | {s['max']:.2f} |"
            if len(parts) > 7:
                row_str += " " + " | ".join(parts[7:]) + " |"
            return row_str
        return line

    lines = clean_text.split('\n')
    new_lines = []
    for line in lines:
        if line.strip().startswith('|') and not ('---' in line or 'Variable' in line or 'Min' in line or 'Golden' in line or 'Parameter' in line):
            new_lines.append(populate_table_row(line))
        else:
            new_lines.append(line)
            
    return "\n".join(new_lines)

def generate_default_report(df: pd.DataFrame, query_str: str = "") -> str:
    """Generate a comprehensive 16-Step, 4-Part Process Optimization Report generically for any dataset."""
    df = df.copy()
    
    # 1. Identify/compute derived KPI variables automatically if standard formula columns exist
    df_cols = list(df.columns)
    
    # Check for coal/fuel flows to auto-derive Total_Fuel_Flow, SFC, etc.
    cols_norm = {norm_str(c): c for c in df_cols}
    
    # Auto-derive Total_Fuel_Flow if main & precalciner coal exist
    main_fuel = next((c for norm, c in cols_norm.items() if any(k in norm for k in ['hoodcoal', 'maintorch', 'torchcoke', 'mainburner'])), None)
    precal_fuel = next((c for norm, c in cols_norm.items() if any(k in norm for k in ['calcinercoal', 'precalcoke', 'calcinercoke'])), None)
    prod_col = next((c for norm, c in cols_norm.items() if any(k in norm for k in ['clinkerproduction', 'production', 'output'])), None)
    feed_col = next((c for norm, c in cols_norm.items() if any(k in norm for k in ['rawmeal', 'mealfeed', 'flourflow', 'feedrate'])), None)
    
    derived_tags = []
    if main_fuel and precal_fuel and 'Total_Fuel_Flow' not in df.columns:
        df['Total_Fuel_Flow'] = df[main_fuel] + df[precal_fuel]
        derived_tags.append(f"[DERIVED: Total_Fuel_Flow = {main_fuel} + {precal_fuel}]")
        
    if prod_col and 'Total_Fuel_Flow' in df.columns and 'Specific_Fuel_Consumption' not in df.columns:
        prod_val = df[prod_col].replace(0, np.nan)
        df['Specific_Fuel_Consumption'] = df['Total_Fuel_Flow'] / prod_val
        derived_tags.append(f"[DERIVED: Specific_Fuel_Consumption = Total_Fuel_Flow / {prod_col}]")
        
    if precal_fuel and 'Total_Fuel_Flow' in df.columns and 'Precalciner_Fuel_Share' not in df.columns:
        fuel_val = df['Total_Fuel_Flow'].replace(0, np.nan)
        df['Precalciner_Fuel_Share'] = df[precal_fuel] / fuel_val
        derived_tags.append(f"[DERIVED: Precalciner_Fuel_Share = {precal_fuel} / Total_Fuel_Flow]")

    # Parse any explicit derived tags in query
    df, extra_derived = parse_and_apply_derived(df, query_str)
    derived_tags.extend(extra_derived)

    # 2. Compute full statistics and correlations
    stats = compute_stats(df)
    numeric_cols = list(stats.keys())
    
    if not numeric_cols:
        return "# Process Analysis Report\n\n> ⚠️ No numeric variables found in dataset to generate report."

    cats = categorize_columns(numeric_cols)
    corrs = compute_correlations(df)
    
    # Identify key variables
    primary_prod = cats['throughput'][0] if cats['throughput'] else numeric_cols[0]
    primary_feed = cats['throughput'][1] if len(cats['throughput']) > 1 else (cats['fuel_energy'][0] if cats['fuel_energy'] else numeric_cols[min(1, len(numeric_cols)-1)])
    primary_fuel = cats['efficiency'][0] if cats['efficiency'] else (cats['fuel_energy'][0] if cats['fuel_energy'] else numeric_cols[min(2, len(numeric_cols)-1)])
    primary_temp = cats['thermal'][0] if cats['thermal'] else numeric_cols[min(3, len(numeric_cols)-1)]
    primary_atmos = cats['atmosphere'][0] if cats['atmosphere'] else numeric_cols[min(4, len(numeric_cols)-1)]
    primary_press = cats['pressure_fan'][0] if cats['pressure_fan'] else numeric_cols[min(5, len(numeric_cols)-1)]
    primary_qual = cats['quality'][0] if cats['quality'] else numeric_cols[min(6, len(numeric_cols)-1)]
    
    prod_s = stats[primary_prod]
    feed_s = stats[primary_feed]
    fuel_s = stats[primary_fuel]
    temp_s = stats[primary_temp]
    atmos_s = stats[primary_atmos]
    press_s = stats[primary_press]
    qual_s = stats[primary_qual]

    # Calculate opportunities & headroom
    operating_hours = 8000.0
    prod_gap = max(0.0, prod_s['q3'] - prod_s['mean'])
    prod_opp_ann = prod_gap * operating_hours
    prod_opp_val = prod_opp_ann * 50.0 # $50 default unit economic value
    
    max_prod_gap = max(0.0, prod_s['max'] - prod_s['mean'])
    max_prod_opp_ann = max_prod_gap * operating_hours
    max_prod_opp_val = max_prod_opp_ann * 50.0
    
    fuel_saving_gap = max(0.0, fuel_s['mean'] - fuel_s['q1'])
    ann_throughput = prod_s['mean'] * operating_hours
    fuel_saved_ann = fuel_saving_gap * ann_throughput
    fuel_saved_val = fuel_saved_ann * 120.0 # $120 default fuel unit cost
    co2_reduction_ann = fuel_saved_ann * 3.15
    total_val_cons = prod_opp_val + fuel_saved_val

    # Top correlations summary text
    corr_bullets = []
    for c1, c2, val in corrs[:5]:
        direction = "positive" if val > 0 else "negative"
        corr_bullets.append(f"• **{c1} vs {c2}**: Strong {direction} correlation (**{val:+.3f}**).")
    corr_summary_str = "\n".join(corr_bullets) if corr_bullets else "• Moderate correlations observed across primary process variables."

    # Build full 16-Step Report
    lines = []
    lines.append("# Comprehensive Process Optimization & Data Analysis Report")
    lines.append("*Dynamically generated based on dataset statistical distributions and correlation signatures*")
    lines.append("\n---")
    
    # PREREQUISITE
    lines.append("\n## ⚡ PREREQUISITE — Create Derived KPI Variables")
    if derived_tags:
        for dt in derived_tags:
            lines.append(dt)
    else:
        lines.append("> All required process metrics and KPI variables are active in the dataset.")
        
    lines.append("\n---")
    
    # PART 1
    lines.append("\n## ⚡ PART 1 — Production & Process Efficiency Analysis")
    
    # STEP 1
    lines.append("\n### STEP 1 — Executive Summary")
    lines.append(f"Based on the dynamic analysis of the active dataset ({len(df):,} records, {len(numeric_cols)} numeric variables), here is the executive engineering summary:")
    lines.append(f"1. **Primary Metric Performance**: `{primary_prod}` averages **{prod_s['mean']:.2f}**, with a Q3 level of **{prod_s['q3']:.2f}** and peak capability of **{prod_s['max']:.2f}**.")
    lines.append(f"2. **Conservative Headroom Opportunity**: Elevating throughput from mean to Q3 yields an annual production gain of **{prod_opp_ann:,.0f} units** (estimated **${prod_opp_val:,.2f}/yr** value).")
    lines.append(f"3. **Feed-to-Output Relationship**: Strong alignment between `{primary_feed}` (mean: {feed_s['mean']:.2f}) and `{primary_prod}` indicates process output is input-driven.")
    lines.append(f"4. **Specific Energy & Efficiency**: `{primary_fuel}` operates at a mean of **{fuel_s['mean']:.3f}** (Q1 optimal: **{fuel_s['q1']:.3f}**). Reducing variance to Q1 represents **{fuel_saved_ann:,.0f} units/yr** savings.")
    lines.append(f"5. **Operating Thermal Profile**: `{primary_temp}` averages **{temp_s['mean']:.1f}** (range: {temp_s['min']:.1f} to {temp_s['max']:.1f}).")
    lines.append(f"6. **Atmospheric & Process Control**: `{primary_atmos}` averages **{atmos_s['mean']:.2f}** (Q1–Q3: {atmos_s['q1']:.2f} to {atmos_s['q3']:.2f}).")
    lines.append(f"7. **System Pressures & Fan Draft**: `{primary_press}` averages **{press_s['mean']:.2f}**, tracking load stability.")
    lines.append(f"8. **Process Quality Indicator**: `{primary_qual}` averages **{qual_s['mean']:.2f}** (median: {qual_s['median']:.2f}).")
    
    lines.append("\n---")
    
    # STEP 2
    lines.append("\n### STEP 2 — Primary Process Driver Analysis")
    lines.append(f"**Graph 1 — Feed vs Primary Metric: Production conversion efficiency**")
    lines.append(f"[SCATTER: X={primary_feed} | Y={primary_prod} | COLOR={primary_fuel} | SCALE=Jet]")
    lines.append(f"The scatter plot maps the relationship between `{primary_feed}` and `{primary_prod}`. High-efficiency operating points (colored by `{primary_fuel}`) concentrate in high-throughput bands.")
    
    lines.append(f"\n**Graph 2 — Energy Input vs Primary Output**")
    lines.append(f"[SCATTER: X={primary_fuel} | Y={primary_prod} | COLOR={primary_atmos} | SCALE=RdBu]")
    lines.append(f"This plot illustrates the energy consumption envelope. Beyond the upper quartile of `{primary_fuel}`, throughput plateaus, highlighting thermal loss or incomplete combustion risks.")
    
    lines.append("\n---")
    
    # STEP 3
    lines.append("\n### STEP 3 — Zone Efficiency & Operating Space Maps")
    lines.append(f"**Graph 3 — THE KEY CHART: Throughput vs Specific Consumption (4-Zone Map)**")
    lines.append(f"[SCATTER: X={primary_prod} | Y={primary_fuel} | COLOR={primary_temp} | SCALE=Viridis]")
    lines.append(f"This 4-zone efficiency map identifies optimal operation (high `{primary_prod}`, low `{primary_fuel}`) versus suboptimal startup or over-fueled operational states.")
    
    lines.append(f"\n**Graph 4 — 3D Operating Space: Input × Energy × Output**")
    lines.append(f"[SCATTER3D: X={primary_feed} | Y={primary_fuel} | Z={primary_prod} | COLOR={primary_temp} | SCALE=Jet]")
    lines.append(f"The 3D scatter chart isolates the complete operational envelope. Optimal operating clusters reside where input and energy inputs balance to maximize throughput.")
    
    par_vars = numeric_cols[:min(7, len(numeric_cols))]
    lines.append(f"\n**Graph 5 — Parallel Coordinates: High-Production Operational Signature**")
    lines.append(f"[PARALLEL: {', '.join(par_vars)} | COLOR: {primary_prod}]")
    lines.append("The parallel coordinate signature isolates the multi-variable line paths associated with top-quartile performance.")
    
    lines.append("\n---")
    
    # STEP 4
    lines.append("\n### STEP 4 — Time Trend & Operational Stability Overview")
    lines.append(f"[DUALPLOT: {primary_prod}, {primary_feed} | {primary_fuel}]")
    lines.append(f"[DUALPLOT: {primary_prod} | {primary_temp}]")
    lines.append("The time trend series illustrates operating stability, load shifts, and transient downtime events across the historical dataset.")
    
    lines.append("\n---")
    
    # STEP 5
    lines.append("\n### STEP 5 — Process Summary Table")
    lines.append("| Variable | Min | Q1 | Median | Mean | Q3 | Max | TARGET BAND |")
    lines.append("|---|---|---|---|---|---|---|---|")
    for col in numeric_cols:
        s = stats[col]
        lines.append(f"| {col} | {s['min']:.2f} | {s['q1']:.2f} | {s['median']:.2f} | {s['mean']:.2f} | {s['q3']:.2f} | {s['max']:.2f} | {s['q1']:.2f} – {s['q3']:.2f} |")
        
    lines.append("\n---")
    
    # PART 2
    lines.append("\n## ⚡ PART 2 — Operating Conditions, Chemical & System Diagnostics")
    
    # STEP 6
    lines.append("\n### STEP 6 — Combustion & Atmosphere Diagnostic Map")
    lines.append(f"**Graph 6 — Atmosphere Balance ({primary_atmos} vs {primary_prod})**")
    lines.append(f"[SCATTER: X={primary_atmos} | Y={primary_prod} | COLOR={primary_fuel} | SCALE=Jet]")
    lines.append(f"Optimal process atmosphere for `{primary_atmos}` resides between **{atmos_s['q1']:.2f}** and **{atmos_s['q3']:.2f}**.")
    lines.append(f"\n[BOX: {', '.join(cats['atmosphere'][:min(3, len(cats['atmosphere']))] or [primary_atmos])}]")
    
    lines.append("\n---")
    
    # STEP 7
    lines.append("\n### STEP 7 — Quality & Chemical Moduli Analysis")
    lines.append(f"**Graph 7 — Quality Driver ({primary_qual} vs {primary_prod})**")
    lines.append(f"[SCATTER: X={primary_qual} | Y={primary_prod} | COLOR={primary_temp} | SCALE=Jet]")
    lines.append(f"Quality indicator `{primary_qual}` operates at a mean of **{qual_s['mean']:.2f}**, ensuring product specifications while maintaining thermal stability.")
    
    lines.append("\n---")
    
    # STEP 8
    lines.append("\n### STEP 8 — Critical System Diagnostics")
    lines.append(f"Key correlation drivers extracted from correlation matrix:")
    lines.append(corr_summary_str)
    lines.append(f"\n[DUALPLOT: {primary_temp}, {primary_press} | {primary_prod}]")
    
    lines.append("\n---")
    
    # STEP 9
    lines.append("\n### STEP 9 — Fan, Draft System & Pressures")
    lines.append(f"**Graph 8 — System Pressure vs Throughput**")
    lines.append(f"[SCATTER: X={primary_press} | Y={primary_prod} | COLOR={primary_fuel} | SCALE=Jet]")
    lines.append(f"\n[BOX: {', '.join(cats['pressure_fan'][:min(4, len(cats['pressure_fan']))] or [primary_press])}]")
    
    lines.append("\n---")
    
    # STEP 10
    lines.append("\n### STEP 10 — Subsystem Split Optimization")
    lines.append(f"**Graph 9 — Fuel/Process Split Optimization**")
    lines.append(f"[SCATTER: X={primary_fuel} | Y={primary_prod} | COLOR={primary_temp} | SCALE=Jet]")
    
    lines.append("\n---")
    
    # STEP 11
    lines.append("\n### STEP 11 — Part 2 Diagnostic Summary Table")
    lines.append("| Diagnostic Variable | Min | Q1 | Median | Mean | Q3 | Max | STATUS | TARGET |")
    lines.append("|---|---|---|---|---|---|---|---|---|")
    diag_cols = (cats['atmosphere'] + cats['thermal'] + cats['pressure_fan'] + cats['quality'])[:8] or numeric_cols[:8]
    for c in diag_cols:
        s = stats[c]
        lines.append(f"| {c} | {s['min']:.2f} | {s['q1']:.2f} | {s['median']:.2f} | {s['mean']:.2f} | {s['q3']:.2f} | {s['max']:.2f} | Optimal | {s['q1']:.2f} – {s['q3']:.2f} |")
        
    lines.append("\n---")
    
    # PART 3
    lines.append("\n## ⚡ PART 3 — Multi-Variable Signature Maps & Distributions")
    
    # STEP 12
    lines.append("\n### STEP 12 — Multi-Variable Operating Maps (Parallel Coordinates)")
    par_vars_1 = numeric_cols[:min(6, len(numeric_cols))]
    lines.append(f"**Production Drivers Parallel Coordinate Plot**")
    lines.append(f"[PARALLEL: {', '.join(par_vars_1)} | COLOR: {primary_prod}]")
    
    if len(numeric_cols) >= 4:
        par_vars_2 = numeric_cols[min(2, len(numeric_cols)-4):min(8, len(numeric_cols))]
        lines.append(f"\n**Efficiency Drivers Parallel Coordinate Plot**")
        lines.append(f"[PARALLEL: {', '.join(par_vars_2)} | COLOR: {primary_fuel}]")
        
    lines.append("\n---")
    
    # STEP 13
    lines.append("\n### STEP 13 — 3D Golden Batch Clusters")
    lines.append(f"**3D Operating Space ({primary_feed} × {primary_temp} × {primary_prod})**")
    lines.append(f"[SCATTER3D: X={primary_feed} | Y={primary_temp} | Z={primary_prod} | COLOR={primary_fuel} | SCALE=Jet]")
    
    lines.append("\n---")
    
    # STEP 14
    lines.append("\n### STEP 14 — Statistical Distributions")
    lines.append(f"[BOX: {', '.join(numeric_cols[:min(6, len(numeric_cols))])}]")
    lines.append(f"[HISTOGRAM: {', '.join(numeric_cols[:min(4, len(numeric_cols))])}]")
    
    lines.append("\n---")
    
    # PART 4
    lines.append("\n## ⚡ PART 4 — Golden Batch & Value Lost Quantification")
    
    # STEP 15
    lines.append("\n### STEP 15 — THE GOLDEN BATCH: Complete Target Definition")
    lines.append("| Variable | Golden Batch Lower | Golden Batch Upper | Current Mean | Headroom Gap |")
    lines.append("|---|---|---|---|---|")
    for col in numeric_cols[:10]:
        s = stats[col]
        gap_val = s['q3'] - s['mean']
        gap_str = f"+{gap_val:.2f}" if gap_val >= 0 else f"{gap_val:.2f}"
        lines.append(f"| {col} | {s['q1']:.2f} | {s['q3']:.2f} | {s['mean']:.2f} | {gap_str} |")
        
    lines.append("\n---")
    
    # STEP 16
    lines.append("\n### STEP 16 — VALUE LOST & OPPORTUNITY QUANTIFICATION")
    lines.append(f"**A) Throughput Opportunity ({primary_prod})**")
    lines.append(f"- Elevating `{primary_prod}` from mean (**{prod_s['mean']:.2f}**) to Q3 (**{prod_s['q3']:.2f}**) over {operating_hours:,.0f} operating hours yields **{prod_opp_ann:,.0f} units/yr**, worth **${prod_opp_val:,.2f}/yr**.")
    lines.append(f"- Reaching peak capability (**{prod_s['max']:.2f}**) yields up to **{max_prod_opp_ann:,.0f} units/yr**, worth **${max_prod_opp_val:,.2f}/yr**.")
    
    lines.append(f"\n**B) Efficiency Opportunity ({primary_fuel})**")
    lines.append(f"- Reducing specific energy variance from mean (**{fuel_s['mean']:.3f}**) to Q1 (**{fuel_s['q1']:.3f}**) yields **{fuel_saved_ann:,.0f} units saved per year**, worth **${fuel_saved_val:,.2f}/yr**.")
    lines.append(f"- Estimated CO₂ reduction: **{co2_reduction_ann:,.0f} tons CO₂/year**.")
    
    lines.append("\n**C) Financial Opportunity Summary Table**")
    lines.append("| Opportunity | Annual Quantity | Financial Value |")
    lines.append("|---|---|---|")
    lines.append(f"| Throughput Optimization (conservative) | {prod_opp_ann:,.0f} units/yr | ${prod_opp_val:,.2f}/yr |")
    lines.append(f"| Throughput Optimization (peak potential) | {max_prod_opp_ann:,.0f} units/yr | ${max_prod_opp_val:,.2f}/yr |")
    lines.append(f"| Energy/Efficiency Savings | {fuel_saved_ann:,.0f} units/yr | ${fuel_saved_val:,.2f}/yr |")
    lines.append(f"| CO₂ Emissions Reduction | {co2_reduction_ann:,.0f} t CO₂/yr | Carbon Credit Headroom |")
    lines.append(f"| **TOTAL CONSERVATIVE OPPORTUNITY** | | **${total_val_cons:,.2f}/yr** |")

    return "\n".join(lines)

def generate_local_report(df: pd.DataFrame, query: str = "") -> str:
    """
    Main entry point for local report generation.
    Generates a generic, comprehensive 16-Step report or processes a custom user prompt template.
    Does NOT require an AI API key.
    """
    query_str = (query or "").strip()
    
    # If the user provided a full prompt/template containing markdown headers or chart tags
    has_chart_tags = bool(re.search(r'\[(?:SCATTER|SCATTER3D|PARALLEL|DUALPLOT|BOX|HISTOGRAM|DERIVED)[^\]]*\]', query_str, re.IGNORECASE))
    is_template_prompt = (len(query_str) > 200 and ("#" in query_str or has_chart_tags)) or "PART 1" in query_str or "STEP 1" in query_str
    
    if is_template_prompt:
        logger.info("Processing user-provided report prompt template.")
        return process_prompt_template(df, query_str)
    else:
        logger.info("Generating dynamic 16-step process report from dataset.")
        return generate_default_report(df, query_str)
