"""
Workforce Assignment CP Solver
================================
Implements the CSOP model from the thesis using OR-Tools CP-SAT.

Designed to run both as a standalone CLI script and as a module imported
by a Streamlit app (app.py).

Input files (defaults):
  - Staff_data_availability.xlsx   (no is_stagiaire column required)
  - structured_projects.xlsx

CLI usage:
  pip install ortools pandas openpyxl
  python workforce_cp_solver.py
  python workforce_cp_solver.py --staff my_staff.xlsx --projects my_projects.xlsx
  python workforce_cp_solver.py --no-interactive        # test mode

Streamlit usage:
  See app.py — call run_pipeline() with file-like objects and month params.
"""

import argparse
import ast
import io
import sys
import pandas as pd
from ortools.sat.python import cp_model

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_STAFF    = "Staff_data_availability.xlsx"
DEFAULT_PROJECTS = "structured_projects.xlsx"

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
NOT_COLS  = {
    "Monday":    "Not_Mon",
    "Tuesday":   "Not_Tue",
    "Wednesday": "Not_Wed",
    "Thursday":  "Not_Thu",
    "Friday":    "Not_Fri",
}
COMBIWORLD_MDT_TYPES = {
    "Combiworld", "MDT", "combiworld", "mdt",
    "Combiworld/MDT", "MDT/Combiworld",
}

# ---------------------------------------------------------------------------
# 1. Data loading
# ---------------------------------------------------------------------------

def load_data(staff_source, projects_source):
    """
    Accept file paths (str/Path) or file-like objects (e.g. Streamlit UploadedFile).
    Returns (staff_df, projects_df).
    """
    staff_df    = pd.read_excel(staff_source)
    projects_df = pd.read_excel(projects_source)
    return staff_df, projects_df


# ---------------------------------------------------------------------------
# 2. Month calendar helpers
# ---------------------------------------------------------------------------

WEEKDAY_MAP = {
    "mon": "Monday", "tue": "Tuesday", "wed": "Wednesday",
    "thu": "Thursday", "fri": "Friday", "sat": "Saturday", "sun": "Sunday",
}
FULL_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday",
             "Friday", "Saturday", "Sunday"]


def build_calendar(first_weekday: str, num_days: int) -> list[dict]:
    """
    Return a list of {day: int, weekday: str} dicts for every day of the month.
    first_weekday must be a full name, e.g. 'Monday'.
    """
    start_idx = FULL_WEEK.index(first_weekday)
    return [
        {"day": d, "weekday": FULL_WEEK[(start_idx + d - 1) % 7]}
        for d in range(1, num_days + 1)
    ]


def build_working_days(calendar: list[dict]) -> list[dict]:
    """Filter calendar to Mon-Fri only."""
    return [c for c in calendar if c["weekday"] in DAY_NAMES]


def ask_planning_info() -> tuple[str, int]:
    """
    Interactive CLI prompt for an arbitrary planning period.
    Returns (first_weekday_full_name, n_days).
    """
    print("\n=== Planning Period Setup ===")
    while True:
        raw = input(
            "Enter the weekday of the 1st day of the planning period "
            "(Mon/Tue/Wed/Thu/Fri/Sat/Sun): "
        ).strip().lower()
        if raw in WEEKDAY_MAP:
            first_weekday = WEEKDAY_MAP[raw]
            break
        print("  → Please enter a valid abbreviation (Mon/Tue/Wed/Thu/Fri/Sat/Sun).")

    while True:
        raw = input("Enter the length of the planning period in days (e.g. 1–365): ").strip()
        if raw.isdigit() and 1 <= int(raw) <= 365:
            n_days = int(raw)
            break
        print("  → Please enter an integer between 1 and 365.")

    return first_weekday, n_days


# ---------------------------------------------------------------------------
# 3. Parameter extraction
# ---------------------------------------------------------------------------

def parse_project_cell(cell) -> tuple[int, float]:
    """
    Project day cells are strings like '(1, 5.0)' or '(3, 1.5)'.
    Returns (demand: int, duration: float).  Zero means not scheduled.
    """
    s = str(cell).strip()
    if pd.isna(cell) or s in ("", "nan", "(0, 0.0)", "(0, nan)"):
        return 0, 0.0
    try:
        tup      = ast.literal_eval(s)
        demand   = int(tup[0])
        duration = float(tup[1]) if not pd.isna(tup[1]) else 0.0
        return demand, duration
    except Exception:
        return 0, 0.0


def build_project_list(projects_df: pd.DataFrame, working_days: list[dict]) -> list[dict]:
    """
    Returns a list of project dicts:
      { name, type, demand, duration, occurs: {weekday: bool} }

    demand and duration are taken as the max across all scheduled weekdays
    (they should be consistent within a row; max is a safe fallback).
    """
    projects = []
    for _, row in projects_df.iterrows():
        p_name = str(row["project_name"])
        p_type = str(row["project_type"])

        occurs_on = {}
        demand    = 0
        duration  = 0.0

        for day in DAY_NAMES:
            occ, dur = parse_project_cell(row.get(day, "(0, 0.0)"))
            if occ > 0:
                occurs_on[day] = True
                demand   = max(demand, occ)
                duration = max(duration, dur)
            else:
                occurs_on[day] = False

        if demand == 0:
            continue

        projects.append({
            "name":     p_name,
            "type":     p_type,
            "demand":   demand,
            "duration": duration,
            "occurs":   occurs_on,
        })

    return projects


def build_worker_list(staff_df: pd.DataFrame) -> list[dict]:
    """
    Returns a list of worker dicts.
    Note: is_stagiaire column is no longer expected.
    Zero-contract workers play the role previously assigned to stagiaires.
    """
    workers = []
    for _, row in staff_df.iterrows():
        not_avail = {day: int(row.get(NOT_COLS[day], 0) or 0) for day in DAY_NAMES}
        avail_on  = {day: 1 - not_avail[day] for day in DAY_NAMES}

        contract_hours = float(row.get("Contract_hours_per_week", 0) or 0)

        workers.append({
            "id":             int(row["ID"]),
            "contract_hours": contract_hours,
            "is_permanent":   1 if contract_hours > 0 else 0,  # permanent = has contract hours
            "is_dreammaker":  int(row.get("dreammaker", 0) or 0),
            "avail_on":       avail_on,
            "BSC":            int(row.get("BSC", 0) or 0),
            "CC":             int(row.get("CC", 0) or 0),
            "Combiworld":     int(row.get("Combiworld", 0) or 0),
            "MDT":            int(row.get("MDT", 0) or 0),
        })
    return workers


# ---------------------------------------------------------------------------
# 4. Derived parameter: AvailableProject[w][p]
# ---------------------------------------------------------------------------

def compute_available_projects(
    workers: list[dict],
    projects: list[dict],
    working_days: list[dict],
) -> list[list[int]]:
    """
    AvailableProject[w][p] = 1 iff worker w is available on every weekday
    on which project p occurs within the given month's working days.
    """
    month_weekdays = {wd["weekday"] for wd in working_days}
    avail = []
    for w in workers:
        row = []
        for p in projects:
            ok = all(
                w["avail_on"][day]
                for day, scheduled in p["occurs"].items()
                if scheduled and day in month_weekdays
            )
            row.append(1 if ok else 0)
        avail.append(row)
    return avail


# ---------------------------------------------------------------------------
# 5. CP-SAT model
# ---------------------------------------------------------------------------

def solve(
    workers: list[dict],
    projects: list[dict],
    avail_project: list[list[int]],
    working_days: list[dict],
) -> tuple[str, dict, object]:
    """
    Build and solve the CP-SAT model.

    Unary constraints (preprocessing / domain reduction):
      A. Availability Filtering  — X[w,p] = 0 if AvailableProject[w,p] = 0
      B. Short-Shift Staffing    — X[w,p] = 0 if Duration_p < 1.5 AND ContractHours_w > 0

    Hard constraints:
      A. Project Demand          — Σ_w X[w,p] = Demand_p
      B. 4-Eye Principle         — Σ_w X[w,p]·IsDreammaker_w ≥ 1  (Combiworld/MDT)
      C. One Project per Day     — Σ_p Occurs[p,d]·X[w,p] ≤ 1

    Objective:
      Minimize Σ_w |ContractHours_w - Σ_p Duration_p·X[w,p]|
      (absolute deviation between contracted and assigned hours)

    Returns (status_str, assignments {(w_idx, p_idx): 0/1}, solver).
    """
    num_w = len(workers)
    num_p = len(projects)
    model = cp_model.CpModel()

    # ------------------------------------------------------------------
    # Decision variables X[w, p] ∈ {0, 1}
    # ------------------------------------------------------------------
    X = {
        (w, p): model.NewBoolVar(f"X_w{w}_p{p}")
        for w in range(num_w)
        for p in range(num_p)
    }

    # ------------------------------------------------------------------
    # Unary Constraint A: Availability Filtering
    # Fix X[w,p] = 0 whenever worker w cannot cover all days of project p.
    # ------------------------------------------------------------------
    for w in range(num_w):
        for p in range(num_p):
            if avail_project[w][p] == 0:
                model.Add(X[w, p] == 0)

    # ------------------------------------------------------------------
    # Unary Constraint B: Short-Shift Staffing
    # Duration_p < 1.5 h AND ContractHours_w > 0  =>  X[w,p] = 0
    # (Only zero-contract / non-permanent workers may take short shifts.)
    # ------------------------------------------------------------------
    for p_idx, p in enumerate(projects):
        if p["duration"] < 1.5:
            for w_idx, w in enumerate(workers):
                if w["is_permanent"] == 1:
                    model.Add(X[w_idx, p_idx] == 0)

    # ------------------------------------------------------------------
    # Hard Constraint A: Project Demand
    # Σ_w X[w,p] = Demand_p   ∀ p
    # ------------------------------------------------------------------
    for p_idx, p in enumerate(projects):
        model.Add(sum(X[w, p_idx] for w in range(num_w)) == p["demand"])

    # ------------------------------------------------------------------
    # Hard Constraint B: 4-Eye Principle (Combiworld / MDT)
    # Σ_w X[w,p]·IsDreammaker_w ≥ 1   ∀ p of type Combiworld or MDT
    # ------------------------------------------------------------------
    for p_idx, p in enumerate(projects):
        if p["type"] in COMBIWORLD_MDT_TYPES:
            model.Add(
                sum(X[w, p_idx] * workers[w]["is_dreammaker"]
                    for w in range(num_w)) >= 1
            )

    # ------------------------------------------------------------------
    # Hard Constraint C: One Project per Day per Worker
    # Σ_p Occurs[p,d]·X[w,p] ≤ 1   ∀ w, ∀ d
    # ------------------------------------------------------------------
    month_weekdays = {wd["weekday"] for wd in working_days}
    for day in DAY_NAMES:
        if day not in month_weekdays:
            continue
        day_projects = [p_idx for p_idx, p in enumerate(projects)
                        if p["occurs"].get(day, False)]
        if not day_projects:
            continue
        for w in range(num_w):
            model.Add(sum(X[w, p_idx] for p_idx in day_projects) <= 1)

    # ------------------------------------------------------------------
    # Objective: Minimize Σ_w |ContractHours_w - H_w|
    # where H_w = Σ_p Duration_p · X[w,p]
    #
    # CP-SAT requires integer arithmetic, so scale hours by SCALE=100.
    # |a - b|  is linearised with an auxiliary variable dev_w ≥ 0 via:
    #   dev_w ≥  (contract_w - H_w)
    #   dev_w ≥ -(contract_w - H_w)
    # Then minimize Σ_w dev_w.
    # ------------------------------------------------------------------
    SCALE = 100
    deviation_vars = []

    for w_idx, w in enumerate(workers):
        contract_scaled = int(round(w["contract_hours"] * SCALE))

        # H_w in scaled integer units
        H_w = sum(
            X[w_idx, p_idx] * int(round(p["duration"] * SCALE))
            for p_idx, p in enumerate(projects)
        )

        # Upper bound for deviation: max possible assigned hours (all projects)
        max_hours_scaled = sum(
            int(round(p["duration"] * SCALE)) for p in projects
        )
        max_dev = max(contract_scaled, max_hours_scaled)

        dev = model.NewIntVar(0, max_dev, f"dev_w{w_idx}")
        model.Add(dev >= contract_scaled - H_w)
        model.Add(dev >= H_w - contract_scaled)
        deviation_vars.append(dev)

    model.Minimize(sum(deviation_vars))

    # ------------------------------------------------------------------
    # Solve
    # ------------------------------------------------------------------
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    solver.parameters.log_search_progress = False

    status     = solver.Solve(model)
    status_str = solver.StatusName(status)

    assignments = {}
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for w in range(num_w):
            for p in range(num_p):
                assignments[w, p] = solver.Value(X[w, p])

    return status_str, assignments, solver


# ---------------------------------------------------------------------------
# 6. Output table
# ---------------------------------------------------------------------------

def build_output_table(
    workers: list[dict],
    projects: list[dict],
    assignments: dict,
) -> pd.DataFrame:
    """
    Returns a DataFrame with one row per worker:
      Worker_ID | Contract_hrs_per_wk | Assigned_projects | Assigned_hours
      | Hours_deviation | Is_permanent
    """
    rows = []
    for w_idx, w in enumerate(workers):
        assigned_projects = []
        total_hours = 0.0

        for p_idx, p in enumerate(projects):
            if assignments.get((w_idx, p_idx), 0) == 1:
                assigned_projects.append(f"{p['name']} ({p['duration']}h)")
                total_hours += p["duration"]

        deviation = round(w["contract_hours"] - total_hours, 2)

        rows.append({
            "Worker_ID":           w["id"],
            "Is_permanent":        "Yes" if w["is_permanent"] else "No (0h contract)",
            "Contract_hrs_per_wk": w["contract_hours"],
            "Assigned_projects":   "; ".join(assigned_projects) if assigned_projects else "—",
            "Assigned_hours":      round(total_hours, 2),
            "Hours_deviation":     deviation,   # negative = over-assigned
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 7. Public pipeline function (called by Streamlit app or CLI)
# ---------------------------------------------------------------------------

def run_pipeline(
    staff_source,
    projects_source,
    first_weekday: str,
    num_days: int,
    output_path: str | None = None,
) -> tuple[pd.DataFrame, dict, str]:
    """
    Full pipeline: load → build params → solve → build output table.

    Parameters
    ----------
    staff_source    : file path or file-like object for staff Excel
    projects_source : file path or file-like object for projects Excel
    first_weekday   : full weekday name of the 1st of the month, e.g. 'Monday'
    num_days        : total days in the month (28-31)
    output_path     : if given, save result table as Excel file here

    Returns
    -------
    result_df  : assignment results DataFrame
    summary    : dict with aggregate stats
    status_str : CP-SAT solver status string
    """
    staff_df, projects_df = load_data(staff_source, projects_source)

    calendar     = build_calendar(first_weekday, num_days)
    working_days = build_working_days(calendar)

    workers  = build_worker_list(staff_df)
    projects = build_project_list(projects_df, working_days)

    # Keep only projects active in this month
    month_weekdays = {wd["weekday"] for wd in working_days}
    projects = [
        p for p in projects
        if any(p["occurs"].get(d, False) for d in month_weekdays)
    ]

    avail_project = compute_available_projects(workers, projects, working_days)
    status_str, assignments, solver = solve(workers, projects, avail_project, working_days)

    result_df = build_output_table(workers, projects, assignments)

    if output_path:
        result_df.to_excel(output_path, index=False)

    # Aggregate stats
    permanent   = result_df[result_df["Is_permanent"] == "Yes"]
    total_contr = permanent["Contract_hrs_per_wk"].sum()
    total_asgnd = permanent["Assigned_hours"].sum()
    total_dev   = permanent["Hours_deviation"].abs().sum()
    summary = {
        "n_workers":           len(workers),
        "n_projects":          len(projects),
        "n_working_days":      len(working_days),
        "solver_status":       status_str,
        "permanent_contract_hrs":  round(total_contr, 2),
        "permanent_assigned_hrs":  round(total_asgnd, 2),
        "total_abs_deviation":     round(total_dev, 2),
        "total_assigned_hrs":      round(result_df["Assigned_hours"].sum(), 2),
    }

    return result_df, summary, status_str


# ---------------------------------------------------------------------------
# 8. CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Workforce Assignment CP Solver")
    parser.add_argument("--staff",    default=DEFAULT_STAFF)
    parser.add_argument("--projects", default=DEFAULT_PROJECTS)
    parser.add_argument("--output",   default="assignment_results.xlsx")
    parser.add_argument(
        "--no-interactive", action="store_true",
        help="Skip month prompt; use a default 20-day Mon-Fri month for testing",
    )
    args = parser.parse_args()

    print(f"\nLoading staff data from:    {args.staff}")
    print(f"Loading project data from:  {args.projects}")

    parser.add_argument(
        "--no-interactive", action="store_true",
        help="Skip planning-period prompt; use a default 20-day Mon–Fri period for testing",
    )

    result_df, summary, status_str = run_pipeline(
        staff_source    = args.staff,
        projects_source = args.projects,
        first_weekday   = first_weekday,
        num_days        = num_days,
        output_path     = args.output,
    )

    print(f"\n  Solver status:   {status_str}")
    if status_str not in ("OPTIMAL", "FEASIBLE"):
        print("\n⚠  No feasible solution found. Check demand vs worker availability.")
        sys.exit(1)

    print(f"  Workers:         {summary['n_workers']}")
    print(f"  Active projects: {summary['n_projects']}")
    print(f"  Working days:    {summary['n_working_days']}")

    print("\n=== Assignment Results ===")
    pd.set_option("display.max_colwidth", 80)
    pd.set_option("display.width", 220)
    print(result_df.to_string(index=False))

    print(f"\n--- Summary (permanent staff) ---")
    print(f"  Total contract hrs/wk:       {summary['permanent_contract_hrs']:.1f}")
    print(f"  Total assigned hrs:           {summary['permanent_assigned_hrs']:.2f}")
    print(f"  Total absolute deviation:     {summary['total_abs_deviation']:.2f}")
    print(f"\nResults saved to: {args.output}")


if __name__ == "__main__":
    main()
