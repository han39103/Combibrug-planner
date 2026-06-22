## `workforce_cp_solver.py` (new)

import argparse
import ast
import sys
from typing import Dict, List, Tuple

import pandas as pd
from ortools.sat.python import cp_model


# #########################
# Constants
# #########################

DEFAULT_STAFF = "Staff_data_availability.xlsx"
DEFAULT_PROJECTS = "structured_projects.xlsx"

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

NOT_COLS = {
    "Monday": "Not_Mon",
    "Tuesday": "Not_Tue",
    "Wednesday": "Not_Wed",
    "Thursday": "Not_Thu",
    "Friday": "Not_Fri",
}

COMBIWORLD_MDT_TYPES = {
    "Combiworld",
    "MDT",
    "combiworld",
    "mdt",
    "Combiworld/MDT",
    "MDT/Combiworld",
}

WEEKDAY_MAP = {
    "mon": "Monday",
    "tue": "Tuesday",
    "wed": "Wednesday",
    "thu": "Thursday",
    "fri": "Friday",
    "sat": "Saturday",
    "sun": "Sunday",
}
FULL_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday",
             "Friday", "Saturday", "Sunday"]


# #########################
# 1. Data loading
# #########################

def load_data(staff_source, projects_source):
    ###
 #   Accept file paths (str/Path) or file-like objects.
  #  Returns (staff_df, projects_df).
    ###
    staff_df = pd.read_excel(staff_source)
    projects_df = pd.read_excel(projects_source)
    return staff_df, projects_df



# #########################
# 2. Calendar helpers  (kept mainly for compatibility with app)
# #########################

def build_calendar(first_weekday: str, num_days: int) -> List[dict]:

#    Return a list of {"day": int, "weekday": str} for each day.
 #   first_weekday must be a full name ("Monday", etc.).

    start_idx = FULL_WEEK.index(first_weekday)
    return [
        {"day": d, "weekday": FULL_WEEK[(start_idx + d - 1) % 7]}
        for d in range(1, num_days + 1)
    ]


def build_working_days(calendar: List[dict]) -> List[dict]:
    #Filter calendar to Mon–Fri only
    return [c for c in calendar if c["weekday"] in DAY_NAMES]


def ask_planning_info() -> Tuple[str, int]:
    ###Interactive CLI prompt (kept for CLI compatibility).###
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


# #########################
# 3. Parsing and project/worker structures
# #########################

def parse_project_cell(cell) -> Tuple[int, float]:
    ###
#    Project day cells are strings like '(1, 5.0)' or '(3, 1.5)'.
 #   Returns (demand: int, duration_hours: float).
    ###
    if pd.isna(cell):
        return 0, 0.0
    s = str(cell).strip()
    if s in ("", "nan", "(0, 0.0)", "(0,0.0)", "(0, nan)"):
        return 0, 0.0
    try:
        tup = ast.literal_eval(s)
        demand = int(tup[0])
        duration = float(tup[1]) if not pd.isna(tup[1]) else 0.0
        return demand, duration
    except Exception:
        return 0, 0.0


def build_project_list(projects_df: pd.DataFrame) -> List[dict]:

    projects = []

    for _, row in projects_df.iterrows():

        p_name = str(row["project_name"])
        p_type = str(row["project_type"])

        demand_by_day = {}
        hours_by_day = {}

        total_week_hours = 0.0
        max_daily_demand = 0

        for day in DAY_NAMES:

            demand, hours = parse_project_cell(row.get(day))

            demand_by_day[day] = demand
            hours_by_day[day] = hours

            total_week_hours += demand * hours
            max_daily_demand = max(max_daily_demand, demand)

        if max_daily_demand == 0:
            continue

        projects.append(
            {
                "name": p_name,
                "type": p_type,
                "demand_by_day": demand_by_day,
                "hours_by_day": hours_by_day,
                "weekly_hours": total_week_hours,
                "max_daily_demand": max_daily_demand,
            }
        )

    return projects

def split_large_projects(projects):

    result = []

    for p in projects:

        if p["weekly_hours"] <= 40:
            result.append(p)
            continue

        max_demand = p["max_daily_demand"]

        core_demands = {}
        core_hours = {}

        for day in DAY_NAMES:

            demand = p["demand_by_day"][day]
            hours = p["hours_by_day"][day]

            if demand == max_demand:
                core_demands[day] = demand
                core_hours[day] = hours
            else:
                core_demands[day] = 0
                core_hours[day] = 0.0

        core_project = {
            "name": p["name"] + "_CORE",
            "type": p["type"],
            "demand_by_day": core_demands,
            "hours_by_day": core_hours,
            "weekly_hours": sum(
                core_demands[d] * core_hours[d]
                for d in DAY_NAMES
            ),
            "max_daily_demand": max_demand,
        }

        result.append(core_project)

        for day in DAY_NAMES:

            demand = p["demand_by_day"][day]

            if demand == 0:
                continue

            if demand == max_demand:
                continue

            sub_demands = {d: 0 for d in DAY_NAMES}
            sub_hours = {d: 0.0 for d in DAY_NAMES}

            sub_demands[day] = demand
            sub_hours[day] = p["hours_by_day"][day]

            result.append(
                {
                    "name": f"{p['name']}_{day}",
                    "type": p["type"],
                    "demand_by_day": sub_demands,
                    "hours_by_day": sub_hours,
                    "weekly_hours": demand * p["hours_by_day"][day],
                    "max_daily_demand": demand,
                }
            )

    return result


def build_worker_list(staff_df: pd.DataFrame) -> List[dict]:
    staff_df = staff_df.drop_duplicates(
        subset=["ID"],
        keep="first"
    )

 #   Note: contract is per week. Workers with 0 hours are non-permanent.
  #  Duplicate IDs are kept as separate workers (based on row index).
    workers = []

    for _, row in staff_df.iterrows():
        # Build availability
        avail_on = {}
        for day in DAY_NAMES:
            col = NOT_COLS[day]
            not_flag = int(row.get(col, 0) or 0)
            avail_on[day] = 0 if not_flag == 1 else 1

        contract_hours_week = float(row.get("Contract_hours_per_week", 0) or 0)

        workers.append(
            {
                "id": int(row["ID"]),
                "contract_hours_week": contract_hours_week,
                "is_permanent": 1 if contract_hours_week > 0 else 0,
                "is_dreammaker": int(row.get("dreammaker", 0) or 0),
                "avail_on": avail_on,
                "BSC": int(row.get("BSC", 0) or 0),
                "CC": int(row.get("CC", 0) or 0),
                "Combiworld": int(row.get("Combiworld", 0) or 0),
                "MDT": int(row.get("MDT", 0) or 0),
            }
        )
    return workers


# #########################
# 4. Qualification filter per project
# #########################

def worker_can_do_project(w: dict, p: dict) -> bool:
    ###
 #   Simple qualification rule:
  #    - If project_type contains 'BSC' → require BSC=1
   #   - If project_type contains 'CC'  → require CC=1
    #  - If project_type contains 'Combiworld' → require Combiworld=1
     # - If project_type contains 'MDT' → require MDT=1
    p_type = str(p["type"]).upper()

    if "BSC" in p_type and w["BSC"] != 1:
        return False

    if "CC" in p_type and w["CC"] != 1:
        return False

    if "COMBIWORLD" in p_type and w["Combiworld"] != 1:
        return False

    if "MDT" in p_type and w["MDT"] != 1:
        return False

    return True

# 5. CP-SAT model (new: X[w,p,day], weekly hours, slack)

def solve(
    workers: List[dict],
    projects: List[dict],
) -> Tuple[str, Dict[Tuple[int, int, str], int],
           Dict[Tuple[int, int, str], int],
           cp_model.CpSolver]:
    ###
#    Build and solve CP-SAT.

 #   Variables:
  #    X[w,p,d] ∈ {0,1}  -> worker w covers project p on day d
   #   Y[w,p]   ∈ {0,1}  -> worker w belongs to project p core team (for continuity)
    #  shortage[p,d] ∈ {0..demand[p,d]}  -> unmet demand on that project-day

#    Constraints (main ones):
 #     - Availability: X[w,p,d] = 0 if not available that day
  #    - Qualification: X[w,p,d] = 0 if worker not qualified
   #   - One project per worker per day
    #  - Project demand per day with slack: sum_w X[w,p,d] + shortage[p,d] = demand[p,d]
     # - 4-eye principle for Combiworld/MDT: per project (any day with demand)
      #- Weekly hours per worker: sum_{p,d} X[w,p,d] * hours[p,d] <= min(40, contract_hours_week)
      #- For projects with weekly_hours <= 40: continuity via Y[w,p]:
 #         * X[w,p,d] <= Y[w,p]
  #        * sum_w Y[w,p] = max_daily_demand(p)

   # Objective:
    #  Minimize  BigM * total_shortage  + sum_w |AssignedHours_w - ContractHours_w|
###
    num_w = len(workers)
    num_p = len(projects)

    model = cp_model.CpModel()

    # ##########
    # Decision variables
    # ##########
    # X[w,p,day]
    X = {}
    for w in range(num_w):
        for p in range(num_p):
            for day in DAY_NAMES:
                X[w, p, day] = model.NewBoolVar(f"X_w{w}_p{p}_{day}")

    # Y[w,p] for continuity (for projects <= 40h/week)
    Y = {}
    for w in range(num_w):
        for p in range(num_p):
            Y[w, p] = model.NewBoolVar(f"Y_w{w}_p{p}")

    # Shortage per project-day
    shortage = {}
    for p_idx, p in enumerate(projects):
        for day in DAY_NAMES:
            demand = p["demand_by_day"][day]
            if demand > 0:
                shortage[p_idx, day] = model.NewIntVar(
                    0, demand, f"shortage_p{p_idx}_{day}"
                )

    # ##########
    # Constraints
    # ##########

    # 1) Availability & qualification
    for w_idx, w in enumerate(workers):
        for p_idx, p in enumerate(projects):
            can_do = worker_can_do_project(w, p)
            for day in DAY_NAMES:
                if w["avail_on"][day] == 0 or not can_do:
                    # not available or not qualified
                    model.Add(X[w_idx, p_idx, day] == 0)

    # 2) One project per worker per day
    for w in range(num_w):
        for day in DAY_NAMES:
            day_projects = [
                p_idx for p_idx, p in enumerate(projects)
                if p["demand_by_day"][day] > 0
            ]
            if day_projects:
                model.Add(
                    sum(X[w, p_idx, day] for p_idx in day_projects) <= 1
                )

    # 3) Project demand per day (with shortage)
    for p_idx, p in enumerate(projects):
        for day in DAY_NAMES:
            demand = p["demand_by_day"][day]
            if demand > 0:
                model.Add(
                    sum(X[w, p_idx, day] for w in range(num_w))
                    + shortage[p_idx, day]
                    == demand
                )

    # 4) 4-eye principle for Combiworld/MDT:
    # at least one dreammaker assigned at some day where demand>0
    for p_idx, p in enumerate(projects):
        if p["type"] in COMBIWORLD_MDT_TYPES:
            # Build a list of X for all days (only where demand>0)
            for day in DAY_NAMES:
        
                if p["demand_by_day"][day] <= 0:
                    continue
            
                model.Add(
                    sum(
                        X[w_idx, p_idx, day]
                        for w_idx, w in enumerate(workers)
                        if w["is_dreammaker"] == 1
                    )
                    >= 1
                )

    # 5) Continuity for projects with total weekly hours <= 40
    for p_idx, p in enumerate(projects):
        if p["weekly_hours"] <= 40.0:
            # X[w,p,d] <= Y[w,p] and sum_w Y[w,p] = max_daily_demand(p)
            for w in range(num_w):
                for day in DAY_NAMES:
                    if p["demand_by_day"][day] > 0:
                        model.Add(X[w, p_idx, day] <= Y[w, p_idx])

            model.Add(
                sum(Y[w, p_idx] for w in range(num_w)) == p["max_daily_demand"]
            )
        else:
            # For projects >40h/week, Y[w,p] is not used to constrain X.
            for w in range(num_w):
                model.Add(Y[w, p_idx] == 0)

    # 6) Weekly hours per worker:
    #    AssignedHours_w = sum_{p,d} X[w,p,d] * hours[p,d]
    #    AssignedHours_w <= min(40, contract_hours_week)
    SCALE = 100  # integer scaling for hours
    deviation_vars = []
    total_shortage_vars = []

    for p_idx, p in enumerate(projects):
        for day in DAY_NAMES:
            if p["demand_by_day"][day] > 0:
                total_shortage_vars.append(shortage[p_idx, day])

    for w_idx, w in enumerate(workers):
        contract = w["contract_hours_week"]
        
        # Contract is only used in objective.
        # Workers may exceed contract hours.
        # We only impose a generic safety cap.
        
        MAX_WORKER_HOURS = 40.0

        # H_w scaled
        terms = []
        for p_idx, p in enumerate(projects):
            for day in DAY_NAMES:
                hours = p["hours_by_day"][day]
                if hours > 0 and p["demand_by_day"][day] > 0:
                    coef = int(round(hours * SCALE))
                    terms.append((X[w_idx, p_idx, day], coef))

        if terms:
            H_w = model.NewIntVar(0, int(round(MAX_WORKER_HOURS  * SCALE * 2)), f"H_w{w_idx}")
            model.Add(H_w == sum(var * coef for (var, coef) in terms))
        else:
            H_w = model.NewIntVar(0, 0, f"H_w{w_idx}")
            model.Add(H_w == 0)

        # Hard cap on weekly hours
        model.Add(H_w <= int(round(MAX_WORKER_HOURS * SCALE)))

        # Deviation from contract
        contract_scaled = int(round(contract * SCALE))
        max_dev = int(round(max(40.0, contract) * SCALE)) * 2
        dev = model.NewIntVar(0, max_dev, f"dev_w{w_idx}")
        model.Add(dev >= contract_scaled - H_w)
        model.Add(dev >= H_w - contract_scaled)
        deviation_vars.append(dev)

    # ##########
    # Objective
    # ##########
    BIG_M = 100000
    model.Minimize(
        BIG_M * sum(total_shortage_vars)
        + sum(deviation_vars)
    )

    # ##########
    # Solve
    # ##########
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    solver.parameters.log_search_progress = False

    status = solver.Solve(model)
    status_str = solver.StatusName(status)

    assignments = {}
    shortage_results = {}

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for w in range(num_w):
            for p in range(num_p):
                for day in DAY_NAMES:
                    assignments[w, p, day] = solver.Value(X[w, p, day])
        for key in shortage:
            shortage_results[key] = solver.Value(shortage[key])

    return status_str, assignments, shortage_results, solver


# #########################
# 6. Output tables
# #########################

def build_output_table(
    workers: List[dict],
    projects: List[dict],
    assignments: Dict[Tuple[int, int, str], int],
) -> pd.DataFrame:
    ###
#    Per-worker summary table:
 #     Worker_ID
  #    Contract_hrs_per_week
   #   Assigned_projects (list, with per-day detail)
    #  Assigned_hours_week
     # Deviation_from_contract (= assigned - contract, can be negative)
      #Is_permanent
    ###
    rows = []

    for w_idx, w in enumerate(workers):
        assigned_projects_desc = []
        total_week_hours = 0.0

        # Collect per project and day
        for p_idx, p in enumerate(projects):
            per_day_strs = []
            proj_total = 0.0
            for day in DAY_NAMES:
                if assignments.get((w_idx, p_idx, day), 0) == 1:
                    h = p["hours_by_day"][day]
                    proj_total += h
                    per_day_strs.append(f"{day} ({h}h)")
            if per_day_strs:
                assigned_projects_desc.append(
                    f"{p['name']} [{', '.join(per_day_strs)}] "
                    f"= {proj_total:.1f}h/week"
                )
                total_week_hours += proj_total

        deviation = round(total_week_hours - w["contract_hours_week"], 2)

        rows.append(
            {
                "Worker_ID": w["id"],
                "Is_permanent": "Yes" if w["is_permanent"] else "No (0h contract)",
                "Contract_hrs_per_week": w["contract_hours_week"],
                "Assigned_projects": "; ".join(assigned_projects_desc) if assigned_projects_desc else "—",
                "Assigned_hours_week": round(total_week_hours, 2),
                "Deviation_from_contract": deviation,  # negative = under-assigned
            }
        )

    return pd.DataFrame(rows)


def build_shortage_table(
    projects: List[dict],
    shortages: Dict[Tuple[int, int, str], int],
) -> pd.DataFrame:
    ###
#    Project-day shortage table:
 #     Project
  #    Project_Type
   #   Day
#      Demand
 #     Filled
  #    Missing
   #   Hours_per_worker
    #  Unfilled_worker_hours
    ###
    rows = []

    for p_idx, p in enumerate(projects):
        for day in DAY_NAMES:
            demand = p["demand_by_day"][day]
            if demand <= 0:
                continue
            missing = shortages.get((p_idx, day), 0)
            if missing > 0:
                hours = p["hours_by_day"][day]
                filled = demand - missing
                rows.append(
                    {
                        "Project": p["name"],
                        "Project_Type": p["type"],
                        "Day": day,
                        "Demand": demand,
                        "Filled": filled,
                        "Missing": missing,
                        "Hours_per_worker": hours,
                        "Unfilled_worker_hours": missing * hours,
                    }
                )

    return pd.DataFrame(rows)


# #########################
# 7. Public pipeline function (used by Streamlit + CLI)
# #########################

def run_pipeline(staff_source, projects_source, first_weekday, num_days, output_path=None):
    staff_df, projects_df = load_data(staff_source, projects_source)

#    Full pipeline: load → build params → solve → build output tables.

 #   Note: first_weekday & num_days are kept for interface compatibility,
  #  but the model is built for a single generic week (Mon–Fri).

    workers = build_worker_list(staff_df)
    projects = build_project_list(projects_df)
    projects = split_large_projects(projects)

    status_str, assignments, shortages, solver = solve(workers, projects)

    result_df = build_output_table(workers, projects, assignments)
    shortage_df = build_shortage_table(projects, shortages)

    if output_path:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            result_df.to_excel(writer, sheet_name="Assignments", index=False)
            shortage_df.to_excel(writer, sheet_name="Open_Shifts", index=False)

    # Aggregate stats (permanent only)
    permanent = result_df[result_df["Is_permanent"] == "Yes"]
    total_contr = permanent["Contract_hrs_per_week"].sum()
    total_asgnd = permanent["Assigned_hours_week"].sum()
    total_dev = permanent["Deviation_from_contract"].abs().sum()

    non_permanent = result_df[
        result_df["Is_permanent"] == "No (0h contract)"
    ]
    
    non_perm_assigned = non_permanent["Assigned_hours_week"].sum()
    non_perm_workers_used = (
        non_permanent["Assigned_hours_week"] > 0
    ).sum()

    
    summary = {
        "n_workers": len(workers),
        "n_projects": len(projects),
        "n_working_days": "Mon–Fri",  # Mon–Fri
        "solver_status": status_str,
        "permanent_contract_hrs": round(total_contr, 2),
        "permanent_assigned_hrs": round(total_asgnd, 2),
        "total_abs_deviation": round(total_dev, 2),
        "total_assigned_hrs": round(result_df["Assigned_hours_week"].sum(), 2),
        "non_perm_assigned_hrs": round(non_perm_assigned, 2),
        "non_perm_workers_used": int(non_perm_workers_used),

    }

    return result_df, shortage_df, summary, status_str


# #########################
# 8. CLI entry point
# #########################

def main():
    parser = argparse.ArgumentParser(description="Workforce Assignment CP Solver")
    parser.add_argument("--staff", default=DEFAULT_STAFF)
    parser.add_argument("--projects", default=DEFAULT_PROJECTS)
    parser.add_argument("--output", default="assignment_results.xlsx")
    parser.add_argument(
        "--no-interactive",
        action="store_true",
        help="Skip month prompt; planning params are ignored by the weekly model anyway.",
    )
    args = parser.parse_args()

    print(f"\nLoading staff data from: {args.staff}")
    print(f"Loading project data from: {args.projects}")

    if args.no_interactive:
        first_weekday = "Monday"
        num_days = 30
    else:
        first_weekday, num_days = ask_planning_info()

    result_df, shortage_df, summary, status_str = run_pipeline(
        staff_source=args.staff,
        projects_source=args.projects,
        first_weekday=first_weekday,
        num_days=num_days,
        output_path=args.output,
    )

    print(f"\nSolver status:   {status_str}")
    if status_str not in ("OPTIMAL", "FEASIBLE"):
        print("\n⚠  No feasible solution found. Check demand vs worker availability.")
        sys.exit(1)

    print(f"  Workers:         {summary['n_workers']}")
    print(f"  Active projects: {summary['n_projects']}")
    print(f"  Working days:    {summary['n_working_days']}")

    print("\n=== Assignment Results (weekly) ===")
    pd.set_option("display.max_colwidth", 120)
    pd.set_option("display.width", 220)
    print(result_df.to_string(index=False))

    print("\n# Summary (permanent staff, weekly) #")
    print(f"  Total contract hrs/week:  {summary['permanent_contract_hrs']:.1f}")
    print(f"  Total assigned hrs/week:  {summary['permanent_assigned_hrs']:.1f}")
    print(f"  Total absolute deviation: {summary['total_abs_deviation']:.1f}")

    print(f"\nResults saved to: {args.output}")


if __name__ == "__main__":
    main()

#
