import pandas as pd
import numpy as np
import io
import logging

logger = logging.getLogger(__name__)

def load_data(file_content: bytes, filename: str) -> pd.DataFrame:
    """
    Load data from CSV or Excel file content.
    """
    try:
        if filename.endswith('.csv'):
            # Read CSV with flexible settings
            try:
                # Try reading with default settings first
                df = pd.read_csv(io.BytesIO(file_content), sep=None, engine='python')
            except Exception:
                # Fallback to specific encodings if needed
                df = pd.read_csv(io.BytesIO(file_content), sep=';', encoding='latin1')
        elif filename.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(io.BytesIO(file_content))
            # Fix for CSV data saved as Excel in a single column
            if df.shape[1] == 1:
                col_name = str(df.columns[0])
                if ',' in col_name or ';' in col_name:
                    sep = ';' if ';' in col_name else ','
                    try:
                        import io as _io
                        stream = _io.StringIO()
                        stream.write(col_name + '\n')
                        for val in df.iloc[:, 0]:
                            stream.write(str(val) + '\n')
                        stream.seek(0)
                        df = pd.read_csv(stream, sep=sep, engine='python')
                    except Exception as ex:
                        logger.warning(f"Attempted to split single-column Excel file but failed: {ex}")
        else:
            raise ValueError("Unsupported file format")
        
        # Standardize column names (strip whitespace and replace spaces with underscores)
        df.columns = df.columns.astype(str).str.strip().str.replace(' ', '_')
        
        # Try to identify timestamp column
        # Heuristic: look for 'time' or 'date' in column name, or check first column
        time_col = None
        for col in df.columns:
            if 'time' in col.lower() or 'date' in col.lower():
                time_col = col
                break
        
        if not time_col:
            # Fallback to first column
            time_col = df.columns[0]
            
        # Dynamically detect date format (DD/MM/YYYY vs MM/DD/YYYY vs YYYY-MM-DD)
        sample = df[time_col].dropna().astype(str).head(1000)
        is_day_first = True  # Default assumption for European format
        import re
        for s in sample:
            nums = re.findall(r'\d+', s)
            if len(nums) >= 2:
                if len(nums[0]) == 4:  # YYYY-MM-DD
                    is_day_first = False
                    break
                first, second = int(nums[0]), int(nums[1])
                if first > 12:  # DD/MM/YYYY
                    is_day_first = True
                    break
                elif second > 12:  # MM/DD/YYYY
                    is_day_first = False
                    break

        df[time_col] = pd.to_datetime(df[time_col], errors='coerce', dayfirst=is_day_first)
        df = df.dropna(subset=[time_col])
        df = df.sort_values(by=time_col)
        df = df.rename(columns={time_col: 'timestamp'})
        
        # Ensure numeric columns are numeric
        for col in df.columns:
            if col != 'timestamp':
                # Replace comma with dot for European formats
                if pd.api.types.is_string_dtype(df[col]):
                     df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
        # Remove extreme outliers (sensor glitches)
        for col in df.columns:
            if col != 'timestamp' and pd.api.types.is_numeric_dtype(df[col]):
                q_low = df[col].quantile(0.01)
                q_high = df[col].quantile(0.99)
                iqr = q_high - q_low
                if iqr > 0:
                    lower_bound = q_low - 1.5 * iqr
                    upper_bound = q_high + 1.5 * iqr
                    df.loc[(df[col] < lower_bound) | (df[col] > upper_bound), col] = np.nan
                
        return df
    except Exception as e:
        logger.error(f"Error loading file {filename}: {e}")
        raise

def parse_legend(file_content: bytes) -> dict:
    """
    Parse LEGEND.csv file to extract metadata (min, max, units).
    Expected format: Variable; Description; Unit; Scale min; Scale max
    Returns a dict: { 'VariableName': { 'Unit': '...', 'Scale min': val, 'Scale max': val } }
    """
    try:
        # Try semicolon first (common in these kiln exports), then fallback to auto-detect
        try:
            df = pd.read_csv(io.BytesIO(file_content), sep=';', engine='python')
        except:
            df = pd.read_csv(io.BytesIO(file_content), sep=None, engine='python')
            
        legend = {}
        
        # Normalize column names for easier lookup
        df.columns = df.columns.astype(str).str.strip()
        cols_lower = [c.lower() for c in df.columns]
        
        # Identify key columns by looking for common substrings
        def find_col(substrings):
            for i, c in enumerate(cols_lower):
                if any(sub in c for sub in substrings):
                    return df.columns[i]
            return None

        var_col = find_col(['variable', 'tag', 'name']) or df.columns[0]
        min_col = find_col(['scale min', 'scalemin', 'min']) 
        max_col = find_col(['scale max', 'scalemax', 'max'])
        unit_col = find_col(['unit'])
        desc_col = find_col(['description', 'desc'])
        
        for _, row in df.iterrows():
            var_name = str(row[var_col]).strip()
            if not var_name or pd.isna(var_name) or var_name.lower() == 'nan':
                continue
                
            entry = {}
            if min_col and pd.notna(row[min_col]):
                try:
                    entry['Scale min'] = str(row[min_col]).replace(',', '.')
                except: pass
            
            if max_col and pd.notna(row[max_col]):
                try:
                    entry['Scale max'] = str(row[max_col]).replace(',', '.')
                except: pass
                
            if unit_col and pd.notna(row[unit_col]):
                entry['Unit'] = str(row[unit_col]).strip()
            
            if desc_col and pd.notna(row[desc_col]):
                entry['Description'] = str(row[desc_col]).strip()
                
            legend[var_name] = entry
            
        return legend
    except Exception as e:
        logger.error(f"Error parsing LEGEND file: {e}")
        return {}

def merge_datasets(dfs: list[pd.DataFrame], interpolation_method: str = 'linear') -> pd.DataFrame:
    """
    Merge multiple dataframes on timestamp.
    interpolation_method: 'linear', 'nearest', 'ffill', 'bfill'
    """
    if not dfs:
        return pd.DataFrame()
    
    # Base dataframe is the first one
    merged_df = dfs[0].copy()
    merged_df = merged_df.set_index('timestamp')
    
    for i in range(1, len(dfs)):
        other_df = dfs[i].copy()
        other_df = other_df.set_index('timestamp')
        
        # Join on timestamp (outer join to keep all time points initially, or use reindex?)
        # For process data, usually we want to align to the main dataset (left join)
        # or union of all times. Let's do outer join and sort.
        merged_df = merged_df.join(other_df, how='outer', rsuffix=f'_file_{i}')
        
    merged_df = merged_df.sort_index()
    
    # Interpolate
    # Ensure all columns are numeric for interpolation
    # (Strings like descriptions shouldn't be interpolated, but we filtered those in load_data)
    if interpolation_method == 'linear':
         merged_df = merged_df.interpolate(method='time')
    elif interpolation_method == 'nearest':
         merged_df = merged_df.interpolate(method='nearest')
    elif interpolation_method == 'ffill':
         merged_df = merged_df.ffill()
    elif interpolation_method == 'bfill':
         merged_df = merged_df.bfill()
         
    # Reset index to make timestamp a column again
    merged_df = merged_df.reset_index()
    
    return merged_df

def filter_columns(df: pd.DataFrame, min_nonzero_pct: float = 0.0) -> pd.DataFrame:
    """
    Remove columns that have too many zeros or NaNs.
    min_nonzero_pct: 0-100. Keep col if (non_zero_count / total_rows) * 100 >= min_nonzero_pct
    """
    if min_nonzero_pct <= 0:
        return df
        
    keep_cols = ['timestamp']
    n_rows = len(df)
    if n_rows == 0:
        return df
        
    for col in df.columns:
        if col == 'timestamp':
            continue
            
        # Count non-zero and non-nan values
        # Assume numeric.
        series = df[col]
        non_zero = ((series != 0) & (series.notna())).sum()
        pct = (non_zero / n_rows) * 100
        
        if pct >= min_nonzero_pct:
            keep_cols.append(col)
            
    return df[keep_cols]


def calculate_derived_var(df: pd.DataFrame, name: str, formula: str) -> pd.DataFrame:
    """
    Add a new column based on a formula with exact and fuzzy column matching.
    Handles column names with parentheses, hyphens, and spaces cleanly.
    """
    import re
    df = df.copy()
    df_cols = list(df.columns)
    cols_sorted = sorted(df_cols, key=len, reverse=True)
    
    env = {}
    temp_formula = formula.strip()
    
    # 1. Match exact columns first (e.g. KILN_FEED_SET_POINT_RF_(Line_A_-_Pen_5))
    for i, col in enumerate(cols_sorted):
        if col in temp_formula:
            token_id = f"__col_{i}__"
            temp_formula = temp_formula.replace(col, token_id)
            env[token_id] = pd.to_numeric(df[col], errors='coerce')

    # 2. Fuzzy match any remaining word tokens if needed
    if any(c.isalpha() for c in temp_formula if c not in env):
        def norm(s):
            return re.sub(r'[^a-z0-9]', '', str(s).lower())

        tokens = re.findall(r'[a-zA-Z_#][a-zA-Z0-9_#\-\.]*', temp_formula)
        reserved = {'np', 'abs', 'min', 'max', 'sum', 'mean', 'std', 'log', 'exp', 'sqrt'}
        var_tokens = [t for t in tokens if t not in reserved and not t.startswith('__col_')]

        for token in sorted(set(var_tokens), key=len, reverse=True):
            token_norm = norm(token)
            matched = None
            for c in df_cols:
                if norm(c) == token_norm:
                    matched = c
                    break
            if matched and matched in df.columns:
                token_id = f"__col_{len(env)}__"
                temp_formula = temp_formula.replace(token, token_id)
                env[token_id] = pd.to_numeric(df[matched], errors='coerce')

    try:
        allowed_globals = {"np": np, "abs": abs, "min": min, "max": max, "sum": sum}
        res = eval(temp_formula, allowed_globals, env)
        if isinstance(res, (pd.Series, np.ndarray, list, float, int)):
            df[name] = pd.Series(res, index=df.index)
            logger.info(f"Successfully calculated derived variable {name} = {formula}")
            return df
    except Exception as err:
        logger.error(f"Error calculating derived var {name} = {formula}: {err}")
        raise ValueError(f"Failed to calculate derived variable '{name}': {err}")

    return df

def process_extra_dataset(
    current_df: pd.DataFrame, 
    file_content: bytes, 
    filename: str, 
    method: str = 'linear', 
    gap_limit: float = None,
    suffix: str = ''
) -> pd.DataFrame:
    """
    Load an extra dataset and merge it into current_df using merge_asof-like logic.
    """
    # Load the new data
    new_df = load_data(file_content, filename)
    if new_df.empty or 'timestamp' not in new_df.columns:
        raise ValueError(f"File {filename} has no valid data or timestamp")
        
    new_df = new_df.sort_values('timestamp')
    
    # Rename columns with suffix if provided
    if suffix:
        rename_map = {c: f"{c}{suffix}" for c in new_df.columns if c != 'timestamp'}
        new_df = new_df.rename(columns=rename_map)
        
    # Prepare for merge
    current_df = current_df.sort_values('timestamp')
    
    new_df_indexed = new_df.set_index('timestamp')
    # Remove duplicates in new_df index
    new_df_indexed = new_df_indexed[~new_df_indexed.index.duplicated(keep='first')]
    
    target_index = current_df['timestamp']
    
    if method == 'linear':
        # combine indexes, interpolate, then select target index
        # We perform a sort of outer join on index, interpolate, then reindex to target
        combined = pd.concat([new_df_indexed, pd.DataFrame(index=target_index)]).sort_index()
        combined = combined[~combined.index.duplicated(keep='first')]
        combined = combined.interpolate(method='time', limit_direction='both') # limit_direction both to extrapolate if needed? or strictly 'time'
        # Requirement says "Linear" or "Linear align + extrapolate"
        # Standard interpolate 'time' handles inside. 
        aligned_df = combined.reindex(target_index)
        
    else:
        # methods: 'nearest', 'ffill', 'bfill'
        pandas_method = method if method in ['nearest', 'ffill', 'bfill'] else None
        
        # Use reindex with method and tolerance
        tol = pd.Timedelta(minutes=gap_limit) if gap_limit and gap_limit > 0 else None
        aligned_df = new_df_indexed.reindex(target_index, method=pandas_method, tolerance=tol)
        
    aligned_df = aligned_df.reset_index()
    
    # Now merge columns. unique timestamp join.
    # aligned_df has same timestamps as current_df.
    merged = pd.merge(current_df, aligned_df, on='timestamp', how='left')
    
    return merged


def get_dataset_summary(df: pd.DataFrame) -> dict:
    """
    Calculate global metrics for the dataset.
    """
    if df.empty or 'timestamp' not in df.columns:
        return {}
        
    start_time = df['timestamp'].min()
    end_time = df['timestamp'].max()
    duration = end_time - start_time
    
    total_points = len(df)
    
    # Calculate average frequency
    if total_points > 1:
        diffs = df['timestamp'].diff().dropna()
        avg_freq = diffs.mean().total_seconds()
        median_freq = diffs.median().total_seconds()
    else:
        avg_freq = 0
        median_freq = 0
        
    return {
        "data_points": total_points,
        "start_time": str(start_time),
        "end_time": str(end_time),
        "duration_str": str(duration),
        "avg_sampling_seconds": avg_freq,
        "median_sampling_seconds": median_freq
    }

def detect_shutdowns(df: pd.DataFrame) -> list[dict]:
    """
    Identify gaps and periods of inactivity.
    """
    if df.empty or 'timestamp' not in df.columns:
        return []
        
    shutdowns = []
    
    # 1. Detect Time Gaps
    diffs = df['timestamp'].diff().dropna()
    if not diffs.empty:
        median_freq = diffs.median()
        threshold = median_freq * 10 # 10x median frequency is a gap
        
        gaps = df[df['timestamp'].diff() > threshold]
        for _, row in gaps.iterrows():
            # Find the row before this one to get the gap start
            idx = df.index.get_loc(row.name)
            prev_row = df.iloc[idx - 1]
            
            gap_duration = row['timestamp'] - prev_row['timestamp']
            
            shutdowns.append({
                "type": "GAP",
                "start": str(prev_row['timestamp']),
                "end": str(row['timestamp']),
                "duration": str(gap_duration)
            })
            
    # 2. Detect Zero-Production Periods (Inactivity)
    # Heuristic: look for columns like RPM, Speed, or Current that are stayed at 0
    prod_cols = [c for c in df.columns if any(sub in c.lower() for sub in ['rpm', 'speed', 'prod', 'feed', 'current'])]
    
    # For now, let's keep it simple and just do gaps. 
    # Real systems might have a dedicated "Status" or "Running" tag.
    
    return shutdowns


def to_native(val):
    if pd.isna(val):
        return None
    try:
        fval = float(val)
        if not np.isfinite(fval):
            return None
        return fval
    except:
        return None

def calculate_stats(df: pd.DataFrame) -> dict:
    """
    Calculate summary statistics for all numeric columns, plus global insights.
    Returns: { "stats": [...], "summary": {...}, "shutdowns": [...] }
    """
    stats_list = []
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    
    for col in numeric_cols:
        if col == 'timestamp':
            continue
            
        series = df[col].dropna()
        if series.empty:
            continue
            
        desc = series.describe()
        q1 = desc['25%']
        q3 = desc['75%']
        median = desc['50%']
        iqr = q3 - q1
        
        # Ranges
        opt_low = median - 0.5 * iqr
        opt_high = median + 0.5 * iqr
        
        # Stability Percentages
        n = len(series)
        pct_below_stable = (series < q1).sum() / n * 100
        pct_above_stable = (series > q3).sum() / n * 100
        pct_below_opt = (series < opt_low).sum() / n * 100
        pct_above_opt = (series > opt_high).sum() / n * 100
        
        # Analytical
        skew = series.skew()
        kurt = series.kurt()
        
        # Zeros
        zeros_count = (series == 0).sum()
        zeros_pct = (zeros_count / n) * 100
        
        stats_list.append({
            'variable': col,
            'min': to_native(desc['min']),
            'max': to_native(desc['max']),
            'mean': to_native(desc['mean']),
            'median': to_native(median),
            'q1': to_native(q1),
            'q3': to_native(q3),
            'std': to_native(desc['std']),
            'count': int(count_non_nan(series)),
            # Advanced metrics
            'opt_low': to_native(opt_low),
            'opt_high': to_native(opt_high),
            'pct_below_stable': to_native(pct_below_stable),
            'pct_above_stable': to_native(pct_above_stable),
            'pct_below_opt': to_native(pct_below_opt),
            'pct_above_opt': to_native(pct_above_opt),
            'skewness': to_native(skew),
            'kurtosis': to_native(kurt),
            'zeros_count': int(zeros_count),
            'zeros_pct': to_native(zeros_pct)
        })
        
    return {
        "stats": stats_list,
        "summary": get_dataset_summary(df),
        "shutdowns": detect_shutdowns(df)
    }

def count_non_nan(series):
    return series.count()

def get_stability_bands(df: pd.DataFrame, var_name: str) -> dict:
    """
    Return stable (Q1-Q3) and optimal (Median +/- 0.5 IQR) bands.
    """
    if var_name not in df.columns:
        return {}
        
    series = df[var_name].dropna()
    if series.empty:
        return {}
        
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    median = series.median()
    iqr = q3 - q1
    
    return {
        'stable_low': to_native(q1),
        'stable_high': to_native(q3),
        'opt_low': to_native(median - 0.5 * iqr),
        'opt_high': to_native(median + 0.5 * iqr)
    }

def downsample_for_plot(df: pd.DataFrame, max_points: int = 5000) -> pd.DataFrame:
    """
    Downsample the dataframe to a maximum number of points for plotting.
    Uses simple periodic sampling.
    """
    if len(df) <= max_points:
        return df
    
    # Calculate sampling interval
    interval = len(df) // max_points
    return df.iloc[::interval].copy()

def calculate_correlations(df: pd.DataFrame, method='pearson') -> dict:
    """
    Calculate correlation matrix.
    Returns format suitable for Plotly Heatmap:
    {
       'z': [[1, 0.5], [0.5, 1]],
       'x': ['VarA', 'VarB'],
       'y': ['VarA', 'VarB']
    }
    """
    numeric_df = df.select_dtypes(include='number')
    if numeric_df.empty:
        return {}
        
    corr_matrix = numeric_df.corr(method=method)
    
    # Replace NaN with None for JSON serialization
    # We need to ensure we cast to object explicitly to hold None
    corr_matrix = corr_matrix.astype(object).where(pd.notna(corr_matrix), None)
    
    return {
        'z': corr_matrix.values.tolist(),
        'x': corr_matrix.columns.tolist(),
        'y': corr_matrix.index.tolist(),
        # Also return pairwise list for "Top Correlations" table
        'pairs': get_top_correlations(corr_matrix)
    }

def get_top_correlations(corr_matrix: pd.DataFrame, top_n: int = 50) -> list[dict]:
    # Unstack and sort
    # mask the diagonal and lower triangle to avoid duplicates
    # values in corr_matrix might be None now, so we need to handle that
    
    # Convert back to numeric for masking/sorting if possible, or handle object
    # It's easier to re-calculate numeric matrix for sorting
    # But we passed in the one with Nones.
    
    # Let's handle it gracefully.
    try:
        # Cast to float, coercing None to NaN
        cm_float = corr_matrix.fillna(np.nan).astype(float)
        
        mask = np.triu(np.ones(cm_float.shape), k=1).astype(bool)
        upper = cm_float.where(mask)
        
        stacked = upper.stack().reset_index()
        stacked.columns = ['var1', 'var2', 'value']
        
        # Sort by absolute value
        stacked['abs_value'] = stacked['value'].abs()
        sorted_corr = stacked.sort_values(by='abs_value', ascending=False)
        
        records = sorted_corr.head(top_n).to_dict(orient='records')
        
        # Clean records (ensure no NaNs)
        clean_records = []
        for r in records:
            clean_records.append({
                'var1': r['var1'],
                'var2': r['var2'],
                'value': to_native(r['value'])
            })
            
        return clean_records
    except Exception as e:
        logger.error(f"Error in top correlations: {e}")
        return []
