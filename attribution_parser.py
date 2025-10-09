import re
import pandas as pd


def parse_attribution(cell_value):
    """Parse a single attribution string and return (completions, influence).

    Expected formats (examples):
      - "3 - 10 sales"
      - "1"

    Non-strings and unparsable values return (0, 0).
    """
    if not isinstance(cell_value, str):
        return 0, 0
    tokens = cell_value.split()
    if not tokens:
        return 0, 0
    try:
        completions = int(re.sub(r'\D', '', tokens[0])) if tokens[0] else 0
    except Exception:
        completions = 0

    influence = 0
    # look for pattern like "- 10 sales" after a dash
    if '-' in tokens:
        try:
            dash_index = tokens.index('-')
            if dash_index + 1 < len(tokens):
                count_str = tokens[dash_index + 1]
                count = int(re.sub(r'\D', '', count_str)) if count_str else 0
                # next token may describe unit; treat any numeric after dash as influence
                influence = count
        except Exception:
            influence = 0

    return completions, influence


def parse_series(series: pd.Series):
    """Apply parse_attribution across a pandas Series and return a DataFrame with
    columns ['Completions', 'Influence'].

    Will gracefully handle missing columns and non-string cells.
    """
    parsed = series.fillna('').apply(parse_attribution).apply(pd.Series)
    parsed.columns = ['Completions', 'Influence']
    return parsed


def parse_dataframe_attributions(df: pd.DataFrame, column_name_contains: str = 'Attribution'):
    """Search for a column that contains `column_name_contains` and parse it.

    Returns a DataFrame with parsed columns merged in, and the name of the
    column parsed (or None).
    """
    for col in df.columns:
        if column_name_contains.lower() in str(col).lower():
            parsed = parse_series(df[col])
            parsed.index = df.index
            return pd.concat([df, parsed], axis=1), col
    # nothing found -> empty parsed DataFrame
    empty = pd.DataFrame({'Completions': [0] * len(df), 'Influence': [0] * len(df)}, index=df.index)
    return pd.concat([df, empty], axis=1), None
