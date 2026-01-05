"""Simple OR-Tools solver for hospital slot assignment.

Each resident must be assigned exactly `total_assignments_per_resident` slots (2).
Per-resident per-hospital maximums:
- '大学病院': up to 2
- other hospitals: up to 1 each

Hospital capacities (slots) must sum to number_of_residents * total_assignments_per_resident.
"""
from typing import List, Dict, Tuple, Any
from ortools.sat.python import cp_model
from datetime import date, timedelta
from typing import Set


HOSPITALS = ["大学病院", "永井病院", "遠山病院", "岩崎病院"]

NO_ASSIGNMENT_ROTATIONS = {
    'OFFSITE_OFFSITE_ONLY',
    'OFFSITE_NO_PREF',
    'OFFSITE_NO_ER',
    '院外-救急しない',
    '大学外‐院外のみ希望',
    '大学外-院外のみ希望',
    '大学外‐院外も救急も希望しない',
    '大学外-院外も救急も希望しない',
}


def assign_shifts(resident_names: List[str], hospital_slots: Dict[str, int], total_assignments_per_resident: int = 2) -> Dict[str, Any]:
    n_res = len(resident_names)
    total_needed = sum(hospital_slots.get(h, 0) for h in HOSPITALS)
    expected = n_res * total_assignments_per_resident
    if total_needed != expected:
        return {"status": "error", "message": f"Total hospital slots ({total_needed}) != expected assignments ({expected}). Adjust slots or resident count."}

    model = cp_model.CpModel()

    # variables x[r,h] integer: assignments of resident r to hospital h
    x = {}
    for i, r in enumerate(resident_names):
        for h in HOSPITALS:
            if h == "大学病院":
                ub = 2
            else:
                ub = 1
            x[(i, h)] = model.NewIntVar(0, ub, f'x_{i}_{h}')

    # Resident total assignments
    for i in range(n_res):
        model.Add(sum(x[(i, h)] for h in HOSPITALS) == total_assignments_per_resident)

    # Hospital capacity constraints
    for h in HOSPITALS:
        cap = hospital_slots.get(h, 0)
        model.Add(sum(x[(i, h)] for i in range(n_res)) == cap)

    # Solve
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10
    solver.parameters.num_search_workers = 8
    res = solver.Solve(model)
    if res != cp_model.OPTIMAL and res != cp_model.FEASIBLE:
        return {"status": "infeasible", "message": "No feasible assignment found"}

    # Build assignment lists
    assignments: Dict[str, List[str]] = {h: [] for h in HOSPITALS}
    per_res_counts: Dict[str, Dict[str, int]] = {}
    for i, r in enumerate(resident_names):
        per_res_counts[r] = {}
        for h in HOSPITALS:
            val = int(solver.Value(x[(i, h)]))
            if val > 0:
                # append resident name `val` times (slots identical)
                assignments[h].extend([r] * val)
            per_res_counts[r][h] = val

    return {"status": "ok", "assignments": assignments, "per_res_counts": per_res_counts}


def assign_shifts_by_day(residents: List[Dict[str, Any]], month: str, hospital_weekday_slots: Dict[str, Dict[int, int]], total_assignments_per_resident: int = 2) -> Dict[str, Any]:
    """Assign residents to hospital slots on specific dates.

    residents: list of dicts with keys: 'name' and 'ng_dates' (list of 'YYYY-MM-DD')
    month: 'YYYY-MM'
    hospital_weekday_slots: {hospital: {weekday_int(0=Mon..6=Sun): slots}}

    Returns assignment per date and per hospital.
    """
    year, mon = [int(x) for x in month.split('-')]
    # build date list
    dates = []
    d = date(year, mon, 1)
    while d.month == mon:
        dates.append(d)
        d = d + timedelta(days=1)

    n_res = len(residents)
    model = cp_model.CpModel()

    # map resident index and date index
    name_to_idx = {residents[i]['name']: i for i in range(n_res)}

    # capacity per (date_idx, hospital)
    cap = {}
    for di, dd in enumerate(dates):
        wd = dd.weekday()
        for h in hospital_weekday_slots:
            cap[(di, h)] = hospital_weekday_slots[h].get(wd, 0)

    # Quick feasibility diagnostics: total capacity vs required assignments
    total_capacity = sum(cap.values())
    # determine required assignments per resident (same logic used later)
    req_per_res = []
    for r in residents:
        if 'required' in r:
            req_per_res.append(int(r['required']))
            continue
        rot = r.get('rotation_type', '')
        if rot in NO_ASSIGNMENT_ROTATIONS:
            req_per_res.append(0)
        else:
            req_per_res.append(int(total_assignments_per_resident))
    total_required = sum(req_per_res)
    diagnostics = None
    if total_capacity < total_required:
        # build per-date totals and per-resident available-day counts for debugging
        per_date_totals = {}
        for di, dd in enumerate(dates):
            per_date_totals[dd.isoformat()] = sum(cap[(di, h)] for h in hospital_weekday_slots)
        per_res_avail = {}
        for idx, r in enumerate(residents):
            ng_set = set(r.get('ng_dates', []))
            avail_days = [dd for dd in dates if dd.isoformat() not in ng_set]
            per_res_avail[r.get('name', f'res_{idx}')] = len(avail_days)
        diagnostics = {
            'total_capacity': total_capacity,
            'total_required': total_required,
            'per_date_capacity': per_date_totals,
            'per_res_avail_days': per_res_avail
        }

    # variables x[r,di,h] in {0,1}
    x = {}
    for r in range(n_res):
        for di in range(len(dates)):
            for h in hospital_weekday_slots:
                x[(r, di, h)] = model.NewBoolVar(f'x_{r}_{di}_{h}')

    # Each resident total assignments == required per resident (allow 0 for some)
    # Determine required assignments per resident: prefer explicit 'required' key, else derive from 'rotation_type'
    req_per_res = []
    for r in residents:
        if 'required' in r:
            req_per_res.append(int(r['required']))
            continue
        rot = r.get('rotation_type', '')
        if rot in NO_ASSIGNMENT_ROTATIONS:
            req_per_res.append(0)
        else:
            # default behavior
            req_per_res.append(int(total_assignments_per_resident))

    # Allow up to the required assignments per resident, but don't force exact equality.
    # We'll maximize total assignments to get the best partial solution when full assignment is impossible.
    for r in range(n_res):
        model.Add(sum(x[(r, di, h)] for di in range(len(dates)) for h in hospital_weekday_slots) <= req_per_res[r])

    # Each (date,hospital) capacity constraint
    for di in range(len(dates)):
        for h in hospital_weekday_slots:
            model.Add(sum(x[(r, di, h)] for r in range(n_res)) <= cap[(di, h)])

    # Per-resident per-hospital max: university 2, others 1
    for r in range(n_res):
        for h in hospital_weekday_slots:
            ub = 2 if h == '大学病院' else 1
            model.Add(sum(x[(r, di, h)] for di in range(len(dates))) <= ub)

    # At most one assignment per resident per day
    for r in range(n_res):
        for di in range(len(dates)):
            model.Add(sum(x[(r, di, h)] for h in hospital_weekday_slots) <= 1)

    # NG dates: prohibit assignments on those dates
    for r_idx, r in enumerate(residents):
        ng_set: Set[str] = set(r.get('ng_dates', []))
        for di, dd in enumerate(dates):
            if dd.isoformat() in ng_set:
                for h in hospital_weekday_slots:
                    model.Add(x[(r_idx, di, h)] == 0)

    # Objective: maximize total assigned slots
    # Build auxiliary variables to prioritize filling "primary" slots per (date,hospital).
    # primary = ceil(cap/2). For each (di,h) and k in 1..primary, create y_k boolean indicating
    # whether at least k assignments exist for that (di,h). Constraint: sum_x >= k * y_k.
    primary_vars = []
    sum_primary = 0
    for di in range(len(dates)):
        for h in hospital_weekday_slots:
            cap_dh = cap[(di, h)]
            if cap_dh <= 0:
                continue
            primary = (cap_dh + 1) // 2  # ceil(cap/2)
            sum_primary += primary
            # create y variables y[(di,h,k)] for k=1..primary
            prev_y = None
            for k in range(1, primary + 1):
                y = model.NewBoolVar(f'y_{di}_{h}_{k}')
                # if y==1 then sum_x >= k
                model.Add(sum(x[(r, di, h)] for r in range(n_res)) >= k * y)
                # monotonicity: y_k+1 <= y_k
                if prev_y is not None:
                    model.Add(y <= prev_y)
                prev_y = y
                primary_vars.append(y)

    # Build auxiliary variables to prioritize having at least one assignment in non-university hospitals
    nonuniv_vars = []
    nonuniv_targets = 0
    for di in range(len(dates)):
        for h in hospital_weekday_slots:
            if h == '大学病院':
                continue
            if cap[(di, h)] <= 0:
                continue
            # y_nu indicates whether at least one assignment exists for (di,h)
            y_nu = model.NewBoolVar(f'nu_{di}_{h}')
            model.Add(sum(x[(r, di, h)] for r in range(n_res)) >= 1 * y_nu)
            nonuniv_vars.append(y_nu)
            nonuniv_targets += 1

    # total assigned variables (sum of all x)
    total_assigned_vars = [x[(r, di, h)] for r in range(n_res) for di in range(len(dates)) for h in hospital_weekday_slots]

    # Compose weighted objective to emulate lexicographic priorities:
    # 1) maximize number of non-university (date,h) that have >=1 assigned
    # 2) maximize number of primary slots filled (as before)
    # 3) maximize total assignments
    max_total_assigned = sum(cap.values())
    S_primary = len(primary_vars) if len(primary_vars) > 0 else 0
    S_nonuniv = nonuniv_targets if nonuniv_targets > 0 else 0
    # weights chosen to ensure lexicographic ordering
    W_total = max_total_assigned + 1
    W_primary = W_total * (S_primary + 1)
    W_nonuniv = W_primary * (S_nonuniv + 1)

    if S_nonuniv > 0 or S_primary > 0:
        model.Maximize(sum(nonuniv_vars) * W_nonuniv + sum(primary_vars) * W_primary + sum(total_assigned_vars))
    else:
        model.Maximize(sum(total_assigned_vars))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        return {"status": "infeasible", "message": "No feasible assignment found"}

    # Collect assignments
    assignments = {}
    for di, dd in enumerate(dates):
        date_str = dd.isoformat()
        assignments[date_str] = {h: [] for h in hospital_weekday_slots}
        for h in hospital_weekday_slots:
            for r in range(n_res):
                if solver.Value(x[(r, di, h)]) == 1:
                    assignments[date_str][h].append(residents[r]['name'])

    # compute per-res counts and total assigned
    per_res_counts = {}
    total_assigned_val = 0
    for r_idx, r in enumerate(residents):
        cnt = 0
        for di in range(len(dates)):
            for h in hospital_weekday_slots:
                if solver.Value(x[(r_idx, di, h)]) == 1:
                    cnt += 1
        per_res_counts[r['name']] = cnt
        total_assigned_val += cnt
    # per-res required mapping
    per_res_required = {residents[i]['name']: req_per_res[i] for i in range(n_res)}

    result = {"status": "ok", "dates": [d.isoformat() for d in dates], "assignments": assignments, "per_res_counts": per_res_counts, "per_res_required": per_res_required, "total_assigned": total_assigned_val, "total_required": sum(req_per_res)}
    if diagnostics is not None:
        result['diagnostics'] = diagnostics
    return result


def assign_shifts_by_date(residents: List[Dict[str, Any]], month: str, hospital_config: Dict[str, Dict[str, int]], total_assignments_per_resident: int = 2) -> Dict[str, Any]:
    """
    Similar to assign_shifts_by_day but hospital_config may contain date keys 'YYYY-MM-DD'
    or weekday keys '0'..'6'. For each date in the month, capacity for hospital h is determined by:
      hospital_config[h].get(date_str) else hospital_config[h].get(str(weekday)) else 0
    """
    year, mon = [int(x) for x in month.split('-')]
    # build date list
    dates = []
    d = date(year, mon, 1)
    while d.month == mon:
        dates.append(d)
        d = d + timedelta(days=1)

    n_res = len(residents)
    model = cp_model.CpModel()

    # map resident index and date index
    name_to_idx = {residents[i]['name']: i for i in range(n_res)}

    # capacity per (date_idx, hospital)
    cap = {}
    for di, dd in enumerate(dates):
        wd = dd.weekday()
        dstr = dd.isoformat()
        for h in hospital_config:
            # prefer explicit date key
            v = None
            if dstr in hospital_config[h]:
                v = hospital_config[h][dstr]
            elif str(wd) in hospital_config[h]:
                v = hospital_config[h][str(wd)]
            else:
                v = 0
            cap[(di, h)] = v

    # Quick feasibility diagnostics: total capacity vs required assignments
    total_capacity = sum(cap.values())
    req_per_res = []
    for r in residents:
        if 'required' in r:
            req_per_res.append(int(r['required']))
            continue
        rot = r.get('rotation_type', '')
        if rot in NO_ASSIGNMENT_ROTATIONS:
            req_per_res.append(0)
        else:
            req_per_res.append(int(total_assignments_per_resident))
    total_required = sum(req_per_res)
    diagnostics = None
    if total_capacity < total_required:
        per_date_totals = {}
        for di, dd in enumerate(dates):
            per_date_totals[dd.isoformat()] = sum(cap[(di, h)] for h in hospital_config)
        per_res_avail = {}
        for idx, r in enumerate(residents):
            ng_set = set(r.get('ng_dates', []))
            avail_days = [dd for dd in dates if dd.isoformat() not in ng_set]
            per_res_avail[r.get('name', f'res_{idx}')] = len(avail_days)
        diagnostics = {
            'total_capacity': total_capacity,
            'total_required': total_required,
            'per_date_capacity': per_date_totals,
            'per_res_avail_days': per_res_avail
        }

    # variables x[r,di,h] in {0,1}
    x = {}
    for r in range(n_res):
        for di in range(len(dates)):
            for h in hospital_config:
                x[(r, di, h)] = model.NewBoolVar(f'x_{r}_{di}_{h}')

    # required per resident
    req_per_res = []
    for r in residents:
        if 'required' in r:
            req_per_res.append(int(r['required']))
            continue
        rot = r.get('rotation_type', '')
        if rot in NO_ASSIGNMENT_ROTATIONS:
            req_per_res.append(0)
        else:
            req_per_res.append(int(total_assignments_per_resident))

    # Allow up to required per resident; we'll maximize total assignments instead of forcing equality
    for r in range(n_res):
        model.Add(sum(x[(r, di, h)] for di in range(len(dates)) for h in hospital_config) <= req_per_res[r])

    # capacity per date/hospital
    for di in range(len(dates)):
        for h in hospital_config:
            model.Add(sum(x[(r, di, h)] for r in range(n_res)) <= cap[(di, h)])

    # per-resident per-hospital max
    for r in range(n_res):
        for h in hospital_config:
            ub = 2 if h == '大学病院' else 1
            model.Add(sum(x[(r, di, h)] for di in range(len(dates))) <= ub)

    # at most one assignment per resident per day
    for r in range(n_res):
        for di in range(len(dates)):
            model.Add(sum(x[(r, di, h)] for h in hospital_config) <= 1)

    # NG dates
    for r_idx, r in enumerate(residents):
        ng_set = set(r.get('ng_dates', []))
        for di, dd in enumerate(dates):
            if dd.isoformat() in ng_set:
                for h in hospital_config:
                    model.Add(x[(r_idx, di, h)] == 0)

    # Objective: maximize total assignments
    # Build auxiliary variables for primary slots per (date,hospital)
    primary_vars = []
    sum_primary = 0
    for di in range(len(dates)):
        for h in hospital_config:
            cap_dh = cap[(di, h)]
            if cap_dh <= 0:
                continue
            primary = (cap_dh + 1) // 2
            sum_primary += primary
            prev_y = None
            for k in range(1, primary + 1):
                y = model.NewBoolVar(f'y_{di}_{h}_{k}')
                model.Add(sum(x[(r, di, h)] for r in range(n_res)) >= k * y)
                if prev_y is not None:
                    model.Add(y <= prev_y)
                prev_y = y
                primary_vars.append(y)

    # non-university priority: prefer at least one assignment in non-univ hospitals per date
    nonuniv_vars = []
    nonuniv_targets = 0
    for di in range(len(dates)):
        for h in hospital_config:
            if h == '大学病院':
                continue
            if cap[(di, h)] <= 0:
                continue
            y_nu = model.NewBoolVar(f'nu_{di}_{h}')
            model.Add(sum(x[(r, di, h)] for r in range(n_res)) >= 1 * y_nu)
            nonuniv_vars.append(y_nu)
            nonuniv_targets += 1

    total_assigned_vars = [x[(r, di, h)] for r in range(n_res) for di in range(len(dates)) for h in hospital_config]
    max_total_assigned = sum(cap.values())
    S_primary = len(primary_vars) if len(primary_vars) > 0 else 0
    S_nonuniv = nonuniv_targets if nonuniv_targets > 0 else 0
    W_total = max_total_assigned + 1
    W_primary = W_total * (S_primary + 1)
    W_nonuniv = W_primary * (S_nonuniv + 1)
    if S_nonuniv > 0 or S_primary > 0:
        model.Maximize(sum(nonuniv_vars) * W_nonuniv + sum(primary_vars) * W_primary + sum(total_assigned_vars))
    else:
        model.Maximize(sum(total_assigned_vars))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        return {"status": "infeasible", "message": "No feasible assignment found"}

    assignments = {}
    for di, dd in enumerate(dates):
        date_str = dd.isoformat()
        assignments[date_str] = {h: [] for h in hospital_config}
        for h in hospital_config:
            for r in range(n_res):
                if solver.Value(x[(r, di, h)]) == 1:
                    assignments[date_str][h].append(residents[r]['name'])

    # compute per-res counts and totals
    per_res_counts = {}
    total_assigned_val = 0
    for r_idx, r in enumerate(residents):
        cnt = 0
        for di in range(len(dates)):
            for h in hospital_config:
                if solver.Value(x[(r_idx, di, h)]) == 1:
                    cnt += 1
        per_res_counts[r['name']] = cnt
        total_assigned_val += cnt

    # per-res required mapping
    per_res_required = {residents[i]['name']: req_per_res[i] for i in range(n_res)}

    result = {"status": "ok", "dates": [d.isoformat() for d in dates], "assignments": assignments, "per_res_counts": per_res_counts, "per_res_required": per_res_required, "total_assigned": total_assigned_val, "total_required": sum(req_per_res)}
    if diagnostics is not None:
        result['diagnostics'] = diagnostics
    return result

