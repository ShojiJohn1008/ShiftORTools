"""Parsers for the two spreadsheet types.

Functions accept pandas.DataFrame (uploaded CSV/XLSX parsed to DataFrame).
"""
from typing import Dict, List, Tuple, Any
from .utils import normalize_name, normalize_date_input, get_month_dates, is_holiday
from .schema import Resident, ShiftJSON
import re


ROTATION_MAP = {
    "院外-救急しない": "OFFSITE_NO_ER",
    "院外-救急希望": "OFFSITE_ER_OK",
    "大学病院でローテート": "UNIV_ROTATION",
}


def parse_sheet1(df, month: str, col_map: Dict[str, Any] = None) -> Tuple[List[Resident], List[Dict[str, Any]]]:
    """Parse individual NG sheet.

    df: pandas.DataFrame
    month: 'YYYY-MM'
    col_map: {'name': colname, 'rotation': colname, 'ng_cols':[...]} (optional)

    Returns (residents, parse_errors)
    """
    import pandas as pd
    if col_map is None:
        col_map = {
            'name': 'C',
            'rotation': 'D',
            'ng_cols': ['G', 'H', 'I', 'J']
        }

    residents: List[Resident] = []
    parse_errors: List[Dict[str, Any]] = []
    year, mon = [int(x) for x in month.split('-')]

    for idx, row in df.iterrows():
        # name
        raw_name = row.get(col_map['name']) if col_map['name'] in row else row.iloc[2] if len(row) > 2 else None
        name = normalize_name(raw_name)
        if name == "":
            continue

        raw_rot = row.get(col_map['rotation']) if col_map['rotation'] in row else row.iloc[3] if len(row) > 3 else None
        rot = str(raw_rot).strip() if raw_rot is not None else ""
        rotation_type = ROTATION_MAP.get(rot, 'UNIV_ROTATION')

        ng_dates_set = set()
        ng_reasons = {}

        # Apply rotation-based NG additions
        if rotation_type == 'OFFSITE_NO_ER':
            # all dates in month
            for d in get_month_dates(year, mon):
                ng_dates_set.add(d.isoformat())
                ng_reasons.setdefault(d.isoformat(), []).append('rotation:院外-救急しない')
        elif rotation_type == 'OFFSITE_ER_OK':
            # add all weekdays (Mon-Fri) but exclude jpholiday
            for d in get_month_dates(year, mon):
                if d.weekday() < 5 and not is_holiday(d):
                    ng_dates_set.add(d.isoformat())
                    ng_reasons.setdefault(d.isoformat(), []).append('rotation:院外-救急希望')

        # Manual NG columns
        for col in col_map.get('ng_cols', []):
            raw = row.get(col) if col in row else None
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                continue
            try:
                dates = normalize_date_input(str(raw), month)
                for dd in dates:
                    ng_dates_set.add(dd)
                    ng_reasons.setdefault(dd, []).append(f'manual:{col}')
            except Exception as e:
                parse_errors.append({'row': int(idx)+1, 'col': col, 'text': raw, 'error': str(e)})

        resident = Resident(name=name, rotation_type=rotation_type, ng_dates=sorted(list(ng_dates_set)), ng_reasons=ng_reasons, source_rows=[{'row_index': int(idx)+1}])
        residents.append(resident)

    return residents, parse_errors


def parse_sheet2(df, month: str, resident_names: List[str], col_map: Dict[str, Any] = None) -> Tuple[List[Tuple[str,str]], List[str], List[Dict[str,Any]]]:
    """Parse offsite training schedule.

    Returns: list of (date_iso, resident_name) assignments, unknown_names, parse_errors
    """
    import pandas as pd
    if col_map is None:
        col_map = {'date': 'A', 'weekday': 'B', 'info': 'C'}

    assignments = []
    unknown_names = []
    parse_errors = []
    known_set = set([normalize_name(n) for n in resident_names])

    # Date inheritance: if a row's date cell is empty, inherit the most recent non-empty date cell above it
    last_date_token = None
    for idx, row in df.iterrows():
        raw_date = row.get(col_map['date']) if col_map['date'] in row else (row.iloc[0] if len(row) > 0 else None)
        raw_info = row.get(col_map['info']) if col_map['info'] in row else (row.iloc[2] if len(row) > 2 else None)

        # If both date and info are empty/NA, skip
        try:
            empty_date = raw_date is None or pd.isna(raw_date) or str(raw_date).strip() == ""
        except Exception:
            empty_date = (raw_date is None) or str(raw_date).strip() == ""
        try:
            empty_info = raw_info is None or pd.isna(raw_info) or str(raw_info).strip() == ""
        except Exception:
            empty_info = (raw_info is None) or str(raw_info).strip() == ""

        if empty_date and empty_info:
            continue

        # Determine which date token to use: inherit if necessary
        if empty_date:
            # try to inherit from last seen non-empty token
            if last_date_token is None:
                # If we don't have a last_date_token yet, search upward in the DataFrame
                found = False
                # iterate previous rows (idx-1 .. 0)
                for j in range(int(idx)-1, -1, -1):
                    prow = df.iloc[j]
                    try:
                        p_raw = prow.get(col_map['date']) if col_map['date'] in prow else (prow.iloc[0] if len(prow) > 0 else None)
                    except Exception:
                        p_raw = None
                    try:
                        if p_raw is None or pd.isna(p_raw):
                            continue
                    except Exception:
                        if p_raw is None:
                            continue
                    if str(p_raw).strip() == "":
                        continue
                    # found non-empty date above
                    date_token_to_use = str(p_raw)
                    last_date_token = date_token_to_use
                    found = True
                    break
                if not found:
                    # No prior date to inherit; skip this row and record parse issue
                    parse_errors.append({'row': int(idx)+1, 'col': col_map['date'], 'text': raw_date, 'error': 'no prior date to inherit'})
                    continue
            else:
                date_token_to_use = last_date_token
        else:
            date_token_to_use = str(raw_date)
            last_date_token = date_token_to_use

        # If info is empty, skip
        if empty_info:
            continue

        # Extract date string
        try:
            dates = normalize_date_input(date_token_to_use, month)
            if not dates:
                parse_errors.append({'row': int(idx)+1, 'col': col_map['date'], 'text': date_token_to_use, 'error': 'date parse empty'})
                continue
            date_iso = dates[0]
        except Exception as e:
            parse_errors.append({'row': int(idx)+1, 'col': col_map['date'], 'text': date_token_to_use, 'error': str(e)})
            continue

        info = str(raw_info).strip()
        # split on ： or :
        parts = re.split(r"：|:", info, maxsplit=1)
        if len(parts) == 1:
            # no colon, try to find name token by regexp of kanji/hiragana/katakana and spaces
            name_token = parts[0]
        else:
            name_token = parts[1]

        # remove digits, parentheses and grade markers like ①
        name_token = re.sub(r"[\d①-⑳\(\)\s]+", "", name_token)
        name_token = normalize_name(name_token)
        if name_token == "":
            parse_errors.append({'row': int(idx)+1, 'col': col_map['info'], 'text': raw_info, 'error': 'name parse empty'})
            continue

        if name_token in known_set:
            assignments.append((date_iso, name_token))
        else:
            unknown_names.append(name_token)

    # deduplicate unknown names
    unknown_names = sorted(list(set(unknown_names)))
    return assignments, unknown_names, parse_errors
