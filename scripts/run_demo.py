"""Simple demo: load CSV/XLSX and run parsers to produce ShiftJSON skeleton.

Usage:
  Adjust file paths below or pass via environment/args in future.
"""
import sys
from pathlib import Path
import pandas as pd
from shiftortools.parsers import parse_sheet1, parse_sheet2
from shiftortools.schema import ShiftJSON, Resident
from shiftortools.solver import assign_shifts, assign_shifts_by_day
from shiftortools.output import write_json, write_excel
import json as _json
from pathlib import Path
import os


def main():
    # User should replace these with actual upload paths
    sheet1_path = Path('sample_sheet1.csv')
    sheet2_path = Path('sample_sheet2.csv')
    target_month = '2026-01'

    if not sheet1_path.exists() or not sheet2_path.exists():
        print('Place sample_sheet1.csv and sample_sheet2.csv in the repository root and re-run.')
        sys.exit(1)

    df1 = pd.read_csv(sheet1_path, dtype=str, keep_default_na=False)
    df2 = pd.read_csv(sheet2_path, dtype=str, keep_default_na=False)

    residents, errors1 = parse_sheet1(df1, target_month)
    resident_names = [r.name for r in residents]
    assignments, unknown_names, errors2 = parse_sheet2(df2, target_month, resident_names)

    # Apply assignments to residents' ng_dates
    name_to_res = {r.name: r for r in residents}
    for date_iso, name in assignments:
        if name in name_to_res:
            r = name_to_res[name]
            if date_iso not in r.ng_dates:
                r.ng_dates.append(date_iso)
                r.ng_reasons.setdefault(date_iso, []).append('sheet2:assignment')

    shiftjson = ShiftJSON(month=target_month, residents=residents, unknown_names=unknown_names, parse_errors=errors1+errors2)

    print(shiftjson.to_json())

    # Save JSON output for audit
    out_dir = Path('output')
    out_dir.mkdir(exist_ok=True)
    out_json_path = out_dir / f"{target_month}-shift.json"
    write_json(shiftjson.to_dict(), str(out_json_path))
    print(f"Wrote JSON to {out_json_path}")

    # Demo for day-level solver: define weekday slots per hospital
    # weekday: 0=Mon .. 6=Sun
    hospital_weekday_slots = {
        "大学病院": {0:2, 1:2, 2:2, 3:2, 4:2, 5:0, 6:0},
        "永井病院": {0:0, 1:2, 2:0, 3:0, 4:0, 5:0, 6:0},
        "遠山病院": {0:0, 1:0, 2:1, 3:1, 4:1, 5:0, 6:0},
        "岩崎病院": {0:0, 1:0, 2:0, 3:0, 4:1, 5:0, 6:0},
    }

    residents_for_solver = []
    for r in residents:
        residents_for_solver.append({'name': r.name, 'ng_dates': r.ng_dates, 'rotation_type': r.rotation_type})

    # Allow overriding hospital weekday slots via config file
    cfg_path = Path('config/hospital_weekday_slots.json')
    if cfg_path.exists():
        try:
            with cfg_path.open('r', encoding='utf-8') as cf:
                hospital_weekday_slots = _json.load(cf)
            # convert keys to int for weekdays
            for h in list(hospital_weekday_slots.keys()):
                hospital_weekday_slots[h] = {int(k): int(v) for k, v in hospital_weekday_slots[h].items()}
            print(f"Loaded hospital_weekday_slots from {cfg_path}")
        except Exception as e:
            print(f"Failed to load config {cfg_path}: {e}")

    sol2 = assign_shifts_by_day(residents_for_solver, target_month, hospital_weekday_slots, total_assignments_per_resident=2)
    print("\nDay-level solver result:")
    import json
    print(json.dumps(sol2, ensure_ascii=False, indent=2))
    # If infeasible, produce simple diagnostics
    if sol2.get('status') != 'ok':
        print('\nDiagnosing infeasibility...')
        # build date list
        from datetime import date, timedelta
        year, mon = [int(x) for x in target_month.split('-')]
        dates = []
        d = date(year, mon, 1)
        while d.month == mon:
            dates.append(d)
            d = d + timedelta(days=1)

        # per-resident capacity estimate
        diag = {}
        for r in residents_for_solver:
            name = r['name']
            ng = set(r.get('ng_dates', []))
            possible = 0
            per_h = {}
            for h in hospital_weekday_slots:
                ub = 2 if h == '大学病院' else 1
                avail_days = 0
                for di, dd in enumerate(dates):
                    if dd.isoformat() in ng:
                        continue
                    wd = dd.weekday()
                    if hospital_weekday_slots[h].get(wd, 0) > 0:
                        avail_days += 1
                per_h[h] = {'avail_days': avail_days, 'max_assignable': min(ub, avail_days)}
                possible += per_h[h]['max_assignable']
            diag[name] = {'possible_total': possible, 'per_hospital': per_h}

        print(json.dumps({'diagnostics': diag}, ensure_ascii=False, indent=2))
    else:
        # write solver result and excel output
        out_solver_path = out_dir / f"{target_month}-solver.json"
        write_json(sol2, str(out_solver_path))
        print(f"Wrote solver JSON to {out_solver_path}")

        excel_path = out_dir / f"{target_month}-shift.xlsx"
        write_excel(shiftjson.to_dict(), sol2, str(excel_path))
        print(f"Wrote Excel to {excel_path}")


if __name__ == '__main__':
    main()
