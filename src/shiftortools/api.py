"""FastAPI app to get/update hospital_weekday_slots.json configuration."""
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, validator
from typing import Dict
from datetime import datetime
from pathlib import Path
import json
from datetime import date, timedelta
from src.shiftortools import solver
import pandas as pd
import tempfile
from fastapi import UploadFile, File, Form
from fastapi.responses import FileResponse
from src.shiftortools.output import write_excel

CFG_PATH = Path('config/hospital_weekday_slots.json')

app = FastAPI(title='ShiftORTools Config API')

# Serve frontend static files from the frontend/ directory
app.mount("/static", StaticFiles(directory="frontend"), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# We'll validate payloads manually in endpoints to avoid Pydantic RootModel complexity


def validate_config(payload: dict) -> Dict[str, Dict[str, int]]:
    """
    Validate payload where each hospital maps to either:
      - weekday keys: string/int 0..6 -> int slots
      - date keys: 'YYYY-MM-DD' -> int slots

    Returns normalized dict: {hospital: {key_str: int}}
    where key_str is either '0'..'6' or 'YYYY-MM-DD'.
    """
    if not isinstance(payload, dict):
        raise ValueError('payload must be a dict')
    normalized = {}
    for h, mapping in payload.items():
        if not isinstance(mapping, dict):
            raise ValueError(f'config for {h} must be a dict')
        nm = {}
        for k, v in mapping.items():
            # accept weekday ints or date strings
            key_str = None
            # try integer weekday
            try:
                wk = int(k)
                if 0 <= wk <= 6:
                    key_str = str(wk)
                else:
                    raise ValueError('weekday keys must be between 0 and 6')
            except Exception:
                # try date string YYYY-MM-DD
                try:
                    # ensure valid date
                    if isinstance(k, str):
                        datetime.strptime(k, '%Y-%m-%d')
                        key_str = k
                    else:
                        raise
                except Exception:
                    raise ValueError('keys must be weekday integer 0..6 or date string YYYY-MM-DD')

            if not isinstance(v, int):
                try:
                    vv = int(v)
                except Exception:
                    raise ValueError('slot values must be integer-like')
            else:
                vv = v
            if vv < 0:
                raise ValueError('slot values must be non-negative')
            nm[key_str] = vv
        normalized[h] = nm
    return normalized


def read_config() -> Dict[str, Dict[int, int]]:
    if not CFG_PATH.exists():
        return {}
    with CFG_PATH.open('r', encoding='utf-8') as f:
        raw = json.load(f)
    # normalize values to ints but keep keys as strings (they may be '0'..'6' or 'YYYY-MM-DD')
    out = {}
    for h, m in raw.items():
        nm = {}
        for k, v in m.items():
            try:
                nm[str(k)] = int(v)
            except Exception:
                nm[str(k)] = 0
        out[h] = nm
    return out


def read_residents_for_month(month: str):
    """Try to read parsed residents from output/{month}-shift.json"""
    outp = Path('output') / f'{month}-shift.json'
    if not outp.exists():
        return None
    with outp.open('r', encoding='utf-8') as f:
        raw = json.load(f)
    # Expect key 'residents' as list of dicts
    return raw.get('residents')


def write_solver_output(month: str, solver_result: dict):
    outp = Path('output')
    outp.mkdir(parents=True, exist_ok=True)
    p = outp / f'{month}-solver.json'
    with p.open('w', encoding='utf-8') as f:
        json.dump(solver_result, f, ensure_ascii=False, indent=2)
    return str(p)


def _read_upload_to_df(upload: UploadFile):
    """Read UploadFile into pandas.DataFrame. Supports xlsx/xls/csv."""
    suffix = Path(upload.filename).suffix.lower() if upload.filename else ''
    # read into temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        content = upload.file.read()
        tmp.write(content)
        tmp_path = tmp.name
    try:
        if suffix in ('.xls', '.xlsx'):
            df = pd.read_excel(tmp_path, dtype=object)
        else:
            # try csv
            df = pd.read_csv(tmp_path, dtype=object)
    finally:
        try:
            Path(tmp_path).unlink()
        except Exception:
            pass
    return df


def write_config(payload: Dict[str, Dict[int, int]]):
    # write with string keys for JSON compatibility
    raw = {h: {str(k): int(v) for k, v in m.items()} for h, m in payload.items()}
    CFG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with CFG_PATH.open('w', encoding='utf-8') as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)


@app.get('/api/config')
def get_config():
    return read_config()


@app.put('/api/config')
def put_config(payload: dict):
    # Validate payload (weekday or date keys supported)
    try:
        normalized = validate_config(payload)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    # backup existing
    if CFG_PATH.exists():
        bak = CFG_PATH.with_suffix('.json.bak')
        try:
            CFG_PATH.replace(bak)
        except Exception:
            # if replace fails, ignore backup
            pass

    write_config(normalized)
    return {'status': 'ok', 'message': 'config saved'}


@app.post('/api/config')
def post_config(payload: dict):
    # alias for put
    return put_config(payload)


@app.get('/')
def root():
    index = Path('frontend/index.html')
    if not index.exists():
        raise HTTPException(status_code=404, detail='frontend not found')
    return FileResponse(index)


@app.get('/api/schedule')
def get_schedule(month: str = None):
    """Return a schedule preview for the given month (YYYY-MM).

    Uses `output/{month}-shift.json` for resident definitions (name, ng_dates).
    Falls back to error if not found.
    """
    if month is None:
        today = date.today()
        month = f"{today.year}-{str(today.month).zfill(2)}"

    cfg = read_config()
    if not cfg:
        raise HTTPException(status_code=400, detail='no config found; please set /api/config')

    residents = read_residents_for_month(month)
    if residents is None:
        raise HTTPException(status_code=400, detail=f'no resident data found for {month}; run parser/demo to generate output/{month}-shift.json')

    # call solver.assign_shifts_by_date which accepts date keys and weekday fallback
    res = solver.assign_shifts_by_date(residents, month, cfg)
    return res


@app.get('/api/residents')
def get_residents(month: str = None):
    """Return parsed residents and offsite entries persisted in output/{month}-shift.json."""
    if month is None:
        today = date.today()
        month = f"{today.year}-{str(today.month).zfill(2)}"

    shift_path = Path('output') / f'{month}-shift.json'
    if not shift_path.exists():
        raise HTTPException(status_code=400, detail=f'no parsed resident data for {month}; upload sheets first')
    try:
        with shift_path.open('r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot read resident data: {e}')
    return data


@app.post('/api/manual_assign')
def manual_assign(payload: dict):
    """Apply a manual assignment and persist it to solver and shift JSON.

    Expected payload: {"month":"YYYY-MM","date":"YYYY-MM-DD","resident":"Name","hospital":"Hospital Name"}
    """
    try:
        month = payload.get('month')
        date_iso = payload.get('date')
        resident = payload.get('resident')
        hospital = payload.get('hospital')
    except Exception:
        raise HTTPException(status_code=400, detail='invalid payload')
    if not (month and date_iso and resident and hospital):
        raise HTTPException(status_code=400, detail='month,date,resident,hospital are required')

    solver_path = Path('output') / f'{month}-solver.json'
    shift_path = Path('output') / f'{month}-shift.json'
    if not solver_path.exists():
        raise HTTPException(status_code=400, detail=f'no solver output found for {month}; run solver first')

    # load solver result
    try:
        with solver_path.open('r', encoding='utf-8') as f:
            solver_result = json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot read solver file: {e}')

    # Determine max allowed assignments for this resident (priority: payload, solver per_res_required, default 2)
    try:
        payload_limit = int(payload.get('max_assignments')) if payload.get('max_assignments') is not None else None
    except Exception:
        payload_limit = None

    per_res_required = solver_result.get('per_res_required', {}) if isinstance(solver_result.get('per_res_required', {}), dict) else {}
    def _get_limit_for(name):
        if payload_limit is not None:
            return payload_limit
        try:
            return int(per_res_required.get(name))
        except Exception:
            return 2


    # ensure structure
    assignments = solver_result.get('assignments') or {}
    dates = solver_result.get('dates') or []
    hospitals = solver_result.get('hospitals') or list(assignments.get(next(iter(assignments), ''), {}).keys() if assignments else [])

    # ensure date exists in assignments; if not, create date entry with hospitals from current config order
    if date_iso not in assignments:
        assignments[date_iso] = {h: [] for h in hospitals or list(read_config().keys())}
    # compute current assignment counts and whether resident is already assigned on this date
    current_count = 0
    resident_assigned_on_date = False
    for d, entry in assignments.items():
        for h, arr in entry.items():
            for name in arr:
                if name == resident:
                    current_count += 1
                    if d == date_iso:
                        resident_assigned_on_date = True

    # determine limit for this resident
    limit = _get_limit_for(resident)
    # if after removing existing assignment on this date and adding new one we would exceed limit, reject
    if (current_count - (1 if resident_assigned_on_date else 0) + 1) > limit:
        raise HTTPException(status_code=400, detail=f'上限回数（{limit}回）に達しています')

    # remove resident from any hospital on that date (prevent duplicates)
    for h in list(assignments.get(date_iso, {}).keys()):
        arr = assignments[date_iso].get(h) or []
        assignments[date_iso][h] = [n for n in arr if n != resident]

    # add resident to requested hospital if not present
    assignments.setdefault(date_iso, {})
    assignments[date_iso].setdefault(hospital, [])
    if resident not in assignments[date_iso][hospital]:
        assignments[date_iso][hospital].append(resident)

    # recompute per_res_counts and total_assigned from assignments to keep consistency
    per_res_counts = {}
    total_assigned = 0
    # collect list of residents from shift file if available to seed per_res_required
    residents_list = []
    if shift_path.exists():
        try:
            with shift_path.open('r', encoding='utf-8') as f:
                shiftjson = json.load(f)
                residents_list = [r.get('name') for r in shiftjson.get('residents', []) if 'name' in r]
        except Exception:
            residents_list = []

    # count
    for d, entry in assignments.items():
        for h, arr in entry.items():
            for name in arr:
                per_res_counts[name] = per_res_counts.get(name, 0) + 1
                total_assigned += 1

    solver_result['assignments'] = assignments
    if 'per_res_counts' in solver_result or True:
        solver_result['per_res_counts'] = per_res_counts
    solver_result['total_assigned'] = total_assigned
    # ensure hospitals list is present for front-end rendering
    if not solver_result.get('hospitals'):
        # preserve order from existing variable hospitals if available
        if hospitals:
            solver_result['hospitals'] = hospitals
        else:
            solver_result['hospitals'] = list(read_config().keys())

    # persist solver_result
    try:
        with solver_path.open('w', encoding='utf-8') as f:
            json.dump(solver_result, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot write solver file: {e}')

    # record manual assignment in shiftjson under key 'manual_assignments' for traceability
    try:
        shiftjson = {}
        if shift_path.exists():
            with shift_path.open('r', encoding='utf-8') as f:
                shiftjson = json.load(f)
        ma = shiftjson.get('manual_assignments', {})
        ma.setdefault(date_iso, {}).setdefault(hospital, [])
        if resident not in ma[date_iso][hospital]:
            ma[date_iso][hospital].append(resident)
        shiftjson['manual_assignments'] = ma
        with shift_path.open('w', encoding='utf-8') as f:
            json.dump(shiftjson, f, ensure_ascii=False, indent=2)
    except Exception:
        # non-fatal
        pass

    return {'status': 'ok', 'result': solver_result}


@app.post('/api/manual_move')
def manual_move(payload: dict):
    """Move a resident from one date/hospital to another atomically.

    Expected payload: {
      'month':'YYYY-MM', 'resident':'Name',
      'from_date':'YYYY-MM-DD', 'from_hospital': optional,
      'to_date':'YYYY-MM-DD', 'to_hospital':'Hospital Name',
      'max_assignments': optional int
    }
    """
    try:
        month = payload.get('month')
        resident = payload.get('resident')
        from_date = payload.get('from_date')
        from_hospital = payload.get('from_hospital')
        to_date = payload.get('to_date')
        to_hospital = payload.get('to_hospital')
    except Exception:
        raise HTTPException(status_code=400, detail='invalid payload')
    if not (month and resident and from_date and to_date and to_hospital):
        raise HTTPException(status_code=400, detail='month,resident,from_date,to_date,to_hospital are required')

    solver_path = Path('output') / f'{month}-solver.json'
    shift_path = Path('output') / f'{month}-shift.json'
    if not solver_path.exists():
        raise HTTPException(status_code=400, detail=f'no solver output found for {month}; run solver first')

    try:
        with solver_path.open('r', encoding='utf-8') as f:
            solver_result = json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot read solver file: {e}')

    assignments = solver_result.get('assignments') or {}

    # remove resident from source (either specific hospital or all hospitals on that date)
    removed = False
    if from_date in assignments:
        if from_hospital:
            arr = assignments[from_date].get(from_hospital) or []
            if resident in arr:
                assignments[from_date][from_hospital] = [n for n in arr if n != resident]
                removed = True
        else:
            for h in list(assignments[from_date].keys()):
                arr = assignments[from_date].get(h) or []
                if resident in arr:
                    assignments[from_date][h] = [n for n in arr if n != resident]
                    removed = True

    if not removed:
        raise HTTPException(status_code=400, detail='resident not assigned on from_date')

    # Determine max allowed assignments for this resident (priority: payload, solver per_res_required, default 2)
    try:
        payload_limit = int(payload.get('max_assignments')) if payload.get('max_assignments') is not None else None
    except Exception:
        payload_limit = None

    per_res_required = solver_result.get('per_res_required', {}) if isinstance(solver_result.get('per_res_required', {}), dict) else {}
    def _get_limit_for(name):
        if payload_limit is not None:
            return payload_limit
        try:
            return int(per_res_required.get(name))
        except Exception:
            return 2

    # compute current count after removal
    current_count = 0
    for d, entry in assignments.items():
        for h, arr in entry.items():
            for name in arr:
                if name == resident:
                    current_count += 1

    limit = _get_limit_for(resident)
    # if adding to target would exceed limit, reject
    if (current_count + 1) > limit:
        raise HTTPException(status_code=400, detail=f'上限回数（{limit}回）に達しています')

    # ensure to_date exists
    if to_date not in assignments:
        hospitals = solver_result.get('hospitals') or list(read_config().keys())
        assignments[to_date] = {h: [] for h in hospitals}

    # add resident to target hospital if not already present
    assignments[to_date].setdefault(to_hospital, [])
    if resident not in assignments[to_date][to_hospital]:
        assignments[to_date][to_hospital].append(resident)

    # recompute counts and totals
    per_res_counts = {}
    total_assigned = 0
    for d, entry in assignments.items():
        for h, arr in entry.items():
            for name in arr:
                per_res_counts[name] = per_res_counts.get(name, 0) + 1
                total_assigned += 1

    solver_result['assignments'] = assignments
    solver_result['per_res_counts'] = per_res_counts
    solver_result['total_assigned'] = total_assigned

    # persist solver_result
    try:
        with solver_path.open('w', encoding='utf-8') as f:
            json.dump(solver_result, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot write solver file: {e}')

    # update shiftjson manual_assignments: remove from old, add to new
    try:
        if shift_path.exists():
            with shift_path.open('r', encoding='utf-8') as f:
                shiftjson = json.load(f)
            ma = shiftjson.get('manual_assignments', {})
            # remove from old
            if from_date in ma:
                for h in list(ma[from_date].keys()):
                    arr = ma[from_date].get(h) or []
                    if resident in arr:
                        ma[from_date][h] = [n for n in arr if n != resident]
                        if not ma[from_date][h]:
                            ma[from_date].pop(h, None)
                if not ma.get(from_date):
                    ma.pop(from_date, None)
            # add to new
            ma.setdefault(to_date, {}).setdefault(to_hospital, [])
            if resident not in ma[to_date][to_hospital]:
                ma[to_date][to_hospital].append(resident)
            shiftjson['manual_assignments'] = ma
            with shift_path.open('w', encoding='utf-8') as f:
                json.dump(shiftjson, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return {'status': 'ok', 'result': solver_result}


@app.post('/api/manual_unassign')
def manual_unassign(payload: dict):
    """Remove a manual assignment for a resident on a given date.

    Expected payload: {"month":"YYYY-MM","date":"YYYY-MM-DD","resident":"Name"}
    If the resident is assigned in multiple hospitals on that date, all occurrences will be removed.
    """
    try:
        month = payload.get('month')
        date_iso = payload.get('date')
        resident = payload.get('resident')
    except Exception:
        raise HTTPException(status_code=400, detail='invalid payload')
    if not (month and date_iso and resident):
        raise HTTPException(status_code=400, detail='month,date,resident are required')

    solver_path = Path('output') / f'{month}-solver.json'
    shift_path = Path('output') / f'{month}-shift.json'
    if not solver_path.exists():
        raise HTTPException(status_code=400, detail=f'no solver output found for {month}; run solver first')

    try:
        with solver_path.open('r', encoding='utf-8') as f:
            solver_result = json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot read solver file: {e}')

    assignments = solver_result.get('assignments') or {}
    if date_iso not in assignments:
        raise HTTPException(status_code=400, detail=f'no assignments for {date_iso}')

    removed = False
    for h in list(assignments.get(date_iso, {}).keys()):
        arr = assignments[date_iso].get(h) or []
        if resident in arr:
            assignments[date_iso][h] = [n for n in arr if n != resident]
            removed = True

    if not removed:
        raise HTTPException(status_code=400, detail='resident not assigned on that date')

    # recompute per_res_counts and total_assigned
    per_res_counts = {}
    total_assigned = 0
    for d, entry in assignments.items():
        for h, arr in entry.items():
            for name in arr:
                per_res_counts[name] = per_res_counts.get(name, 0) + 1
                total_assigned += 1

    solver_result['assignments'] = assignments
    solver_result['per_res_counts'] = per_res_counts
    solver_result['total_assigned'] = total_assigned

    # persist solver_result
    try:
        with solver_path.open('w', encoding='utf-8') as f:
            json.dump(solver_result, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'cannot write solver file: {e}')

    # remove from shiftjson manual_assignments if present
    try:
        if shift_path.exists():
            with shift_path.open('r', encoding='utf-8') as f:
                shiftjson = json.load(f)
            ma = shiftjson.get('manual_assignments', {})
            if date_iso in ma:
                for h in list(ma[date_iso].keys()):
                    arr = ma[date_iso].get(h) or []
                    if resident in arr:
                        ma[date_iso][h] = [n for n in arr if n != resident]
                        # remove empty lists/keys
                        if not ma[date_iso][h]:
                            ma[date_iso].pop(h, None)
                if not ma.get(date_iso):
                    ma.pop(date_iso, None)
            shiftjson['manual_assignments'] = ma
            with shift_path.open('w', encoding='utf-8') as f:
                json.dump(shiftjson, f, ensure_ascii=False, indent=2)
    except Exception:
        # non-fatal
        pass

    return {'status': 'ok', 'result': solver_result}


@app.get('/api/is_holiday')
def api_is_holiday(date: str = None):
    """Return whether given ISO date string is a national holiday (Japan).

    Query param: date=YYYY-MM-DD
    """
    if not date:
        raise HTTPException(status_code=400, detail='date required')
    try:
        from datetime import date as _date
        d = _date.fromisoformat(date)
    except Exception:
        raise HTTPException(status_code=400, detail='invalid date')
    try:
        from src.shiftortools.utils import is_holiday
        res = bool(is_holiday(d))
    except Exception:
        res = False
    return {'date': date, 'is_holiday': res}


@app.post('/api/run')
def run_solver(month: str = None):
    """Run solver for month and save solver output to output/{month}-solver.json."""
    if month is None:
        today = date.today()
        month = f"{today.year}-{str(today.month).zfill(2)}"

    cfg = read_config()
    if not cfg:
        raise HTTPException(status_code=400, detail='no config found; please set /api/config')

    residents = read_residents_for_month(month)
    if residents is None:
        raise HTTPException(status_code=400, detail=f'no resident data found for {month}; run parser/demo to generate output/{month}-shift.json')

    res = solver.assign_shifts_by_date(residents, month, cfg)
    if res.get('status') == 'ok':
        # include hospitals ordering from config
        hospitals = list(cfg.keys())
        out = {'month': month, 'hospitals': hospitals, 'dates': res.get('dates'), 'assignments': res.get('assignments')}
        # include solver diagnostics and per-res assignment counts if present
        if 'per_res_counts' in res:
            out['per_res_counts'] = res.get('per_res_counts')
        if 'per_res_required' in res:
            out['per_res_required'] = res.get('per_res_required')
        if 'total_assigned' in res:
            out['total_assigned'] = res.get('total_assigned')
        if 'total_required' in res:
            out['total_required'] = res.get('total_required')
        if 'diagnostics' in res:
            out['diagnostics'] = res.get('diagnostics')
        write_solver_output(month, out)
        return {'status': 'ok', 'path': f'output/{month}-solver.json', 'result': out}
    else:
        return res


@app.post('/api/upload_both')
def upload_both(month: str = Form(...), sheet1: UploadFile = File(...), sheet2: UploadFile = File(...)):
    """Upload two files: sheet1 (研修医NG日) and sheet2 (院外研修日).

    Returns parsed residents (from sheet1) and assignments/unknowns from sheet2.
    """
    try:
        df1 = _read_upload_to_df(sheet1)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f'cannot read sheet1: {e}')
    try:
        df2 = _read_upload_to_df(sheet2)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f'cannot read sheet2: {e}')

    # parse sheet1
    from src.shiftortools.parsers import parse_sheet1, parse_sheet2

    residents, errors1 = parse_sheet1(df1, month)
    resident_names = [r.name for r in residents]

    assignments, unknowns, errors2 = parse_sheet2(df2, month, resident_names)

    # Build offsite raw-info mapping from uploaded sheet2: date_iso -> list of raw info strings
    from src.shiftortools.utils import normalize_date_input
    offsite_map = {}
    # replicate date inheritance logic similar to parse_sheet2
    last_date_token = None
    import pandas as _pd
    for idx, row in df2.iterrows():
        raw_date = row.get('A') if 'A' in row else (row.iloc[0] if len(row) > 0 else None)
        raw_info = row.get('C') if 'C' in row else (row.iloc[2] if len(row) > 2 else None)
        try:
            empty_date = raw_date is None or _pd.isna(raw_date) or str(raw_date).strip() == ""
        except Exception:
            empty_date = (raw_date is None) or str(raw_date).strip() == ""
        try:
            empty_info = raw_info is None or _pd.isna(raw_info) or str(raw_info).strip() == ""
        except Exception:
            empty_info = (raw_info is None) or str(raw_info).strip() == ""

        if empty_date and empty_info:
            continue

        if empty_date:
            if last_date_token is None:
                found = False
                for j in range(int(idx)-1, -1, -1):
                    prow = df2.iloc[j]
                    try:
                        p_raw = prow.get('A') if 'A' in prow else (prow.iloc[0] if len(prow) > 0 else None)
                    except Exception:
                        p_raw = None
                    try:
                        if p_raw is None or _pd.isna(p_raw):
                            continue
                    except Exception:
                        if p_raw is None:
                            continue
                    if str(p_raw).strip() == "":
                        continue
                    date_token_to_use = str(p_raw)
                    last_date_token = date_token_to_use
                    found = True
                    break
                if not found:
                    continue
            else:
                date_token_to_use = last_date_token
        else:
            date_token_to_use = str(raw_date)
            last_date_token = date_token_to_use

        if empty_info:
            continue

        try:
            date_isos = normalize_date_input(date_token_to_use, month)
            for di in date_isos:
                offsite_map.setdefault(di, []).append(str(raw_info).strip())
        except Exception:
            continue

    # merge assignments into residents as additional info (not modifying persisted files)
    assign_map = {}
    for date_iso, name in assignments:
        assign_map.setdefault(date_iso, []).append(name)

    # serialize residents
    res_data = []
    for r in residents:
        res_data.append({'name': r.name, 'rotation_type': r.rotation_type, 'ng_dates': r.ng_dates, 'ng_reasons': r.ng_reasons})

    # For any assignments (院外研修), treat the assignment day and the previous day as NG
    if assignments:
        # build lookup by resident name
        name_map = {rd['name']: rd for rd in res_data}
        for date_iso, name in assignments:
            if name not in name_map:
                continue
            rd = name_map[name]
            # ensure ng_dates is a set for easy addition
            existing = set(rd.get('ng_dates', []))
            # add the assignment day
            if date_iso not in existing:
                existing.add(date_iso)
                rd.setdefault('ng_reasons', {}).setdefault(date_iso, []).append('offsite:研修当日')
            # add previous day
            try:
                d = date.fromisoformat(date_iso)
                prev = (d - timedelta(days=1)).isoformat()
                if prev not in existing:
                    existing.add(prev)
                    rd.setdefault('ng_reasons', {}).setdefault(prev, []).append('offsite:前日')
            except Exception:
                # ignore date parsing errors
                pass
            rd['ng_dates'] = sorted(list(existing))

    # persist parsed residents for this month so /api/run and /api/schedule can use them
    outp = Path('output')
    outp.mkdir(parents=True, exist_ok=True)
    shift_file = outp / f'{month}-shift.json'
    try:
        with shift_file.open('w', encoding='utf-8') as f:
            # persist residents and the original offsite entries so Excel export can rebuild C列
            json.dump({'residents': res_data, 'offsite_entries': offsite_map}, f, ensure_ascii=False, indent=2)
    except Exception:
        # non-fatal: proceed but include a warning in response
        pass

    return {'status': 'ok', 'residents': res_data, 'assignments': assign_map, 'unknown_names': unknowns, 'errors': {'sheet1': errors1, 'sheet2': errors2}, 'persisted_path': str(shift_file)}


@app.get('/api/download')
def download_schedule(month: str = None):
    """Return an Excel file for the given month if solver output exists.

    Uses `output/{month}-solver.json` and `output/{month}-shift.json`.
    """
    if month is None:
        today = date.today()
        month = f"{today.year}-{str(today.month).zfill(2)}"

    solver_path = Path('output') / f'{month}-solver.json'
    shift_path = Path('output') / f'{month}-shift.json'
    if not solver_path.exists():
        raise HTTPException(status_code=400, detail=f'no solver output found for {month}; run solver first')
    # read both
    with solver_path.open('r', encoding='utf-8') as f:
        solver_result = json.load(f)
    shiftjson = {}
    if shift_path.exists():
        with shift_path.open('r', encoding='utf-8') as f:
            shiftjson = json.load(f)

    # create temp xlsx
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tf:
        tmp_path = tf.name
    try:
        write_excel(shiftjson, solver_result, tmp_path)
        return FileResponse(tmp_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=f'{month}-schedule.xlsx')
    finally:
        # Note: we don't unlink immediately because FileResponse will stream the file.
        pass


@app.post('/api/upload_sheet1')
def upload_sheet1(month: str = Form(...), sheet1: UploadFile = File(...)):
    try:
        df1 = _read_upload_to_df(sheet1)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f'cannot read sheet1: {e}')
    from src.shiftortools.parsers import parse_sheet1
    residents, errors1 = parse_sheet1(df1, month)
    res_data = [{'name': r.name, 'rotation_type': r.rotation_type, 'ng_dates': r.ng_dates, 'ng_reasons': r.ng_reasons} for r in residents]
    return {'status': 'ok', 'residents': res_data, 'errors': errors1}


@app.post('/api/upload_sheet2')
def upload_sheet2(month: str = Form(...), sheet2: UploadFile = File(...), resident_names: str = Form(None)):
    try:
        df2 = _read_upload_to_df(sheet2)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f'cannot read sheet2: {e}')
    from src.shiftortools.parsers import parse_sheet2
    names = []
    if resident_names:
        try:
            import json as _json
            names = _json.loads(resident_names)
        except Exception:
            # treat as comma-separated
            names = [n.strip() for n in resident_names.split(',') if n.strip()]
    assignments, unknowns, errors2 = parse_sheet2(df2, month, names)
    assign_map = {}
    for date_iso, name in assignments:
        assign_map.setdefault(date_iso, []).append(name)
    return {'status': 'ok', 'assignments': assign_map, 'unknown_names': unknowns, 'errors': errors2}
