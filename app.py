## `app.py` (updated to match weekly model + fix f‑strings)

import io
import streamlit as st
import pandas as pd
import traceback
import base64

from workforce_cp_solver import run_pipeline, WEEKDAY_MAP, DEFAULT_STAFF, DEFAULT_PROJECTS
def add_bg_image():

    with open("assets/20260621_234523245_iOS.jpg", "rb") as f:
        data = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:image/jpg;base64,{data}");
            background-size: cover;
            background-position: center;
        }}

        .stApp::before {{
            content: "";
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255,255,255,0.20);
            z-index: -1;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )
st.set_page_config(page_title="Workforce Assignment Solver", layout="wide")

st.markdown("""
<style>
div[data-testid="stVerticalBlock"] {
    background-color: rgba(255,255,255,0.8);
    padding: 15px;
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

add_bg_image()
st.markdown(
    """
    <h1 style='text-align:center;'>
    </h1>
    """,
    unsafe_allow_html=True
)

st.image(
    "assets/Combibrug logo.png",
    width=280
)
st.markdown("""
<div style="
    background-color: rgba(255,255,255,0.85);
    padding: 20px;
    border-radius: 15px;
    text-align: center;
    margin-bottom: 20px;
">
<h2>Workforce Assignment CP Solver</h2>
<p>Constraint Programming model — OR-Tools CP-SAT</p>
</div>
""", unsafe_allow_html=True)


# ##################################################
# Sidebar: inputs
# ##################################################
with st.sidebar:
    st.header("Input Files")
    staff_file = st.file_uploader("Staff availability file (.xlsx)", type=["xlsx"])
    projects_file = st.file_uploader("Projects file (.xlsx)", type=["xlsx"])

    st.divider()
    st.header("Planning Period (for display only)")
    weekday_options = ["Monday", "Tuesday", "Wednesday", "Thursday",
                       "Friday", "Saturday", "Sunday"]
    first_weekday = st.selectbox(
        "Weekday of the 1st day of the planning period",
        weekday_options,
        index=0,
    )
    num_days = st.number_input(
        "Length of planning period (days)",
        min_value=1,
        max_value=365,
        value=30,
        step=1,
    )

    st.divider()
    run_btn = st.button("▶  Run Solver", type="primary", use_container_width=True)


# ##################################################
# Main panel
# ##################################################
if run_btn:
    # Use uploaded files or fall back to defaults on disk
    staff_source = staff_file if staff_file is not None else DEFAULT_STAFF
    projects_source = projects_file if projects_file is not None else DEFAULT_PROJECTS

    with st.spinner("Running CP-SAT solver …"):
        try:
            result_df, shortage_df, summary, status_str = run_pipeline(
                staff_source=staff_source,
                projects_source=projects_source,
                first_weekday=first_weekday,
                num_days=int(num_days),
            )
        
        except Exception:
            st.error(traceback.format_exc())
            st.stop()
    # Status badge
    if status_str in ("OPTIMAL", "FEASIBLE"):
        st.success(f"Solver status: **{status_str}**")
    else:
        st.error(
            f"Solver status: **{status_str}** — no feasible solution found. "
            "Check demand vs worker availability."
        )
        st.stop()

    # Summary metrics
    st.subheader("Summary (weekly)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Workers", summary["n_workers"])
    c2.metric("Active projects", summary["n_projects"])
    c3.metric("Working days", summary["n_working_days"])
    c4.metric("Total |deviation| (permanent)", f"{summary['total_abs_deviation']:.1f} h")

    c5, c6, c7 = st.columns(3)
    
    c5.metric("Permanent contract hrs/week",f"{summary['permanent_contract_hrs']:.1f} h")
    c6.metric("Permanent assigned hrs/week",f"{summary['permanent_assigned_hrs']:.1f} h")
    c7.metric("Non-permanent assigned hrs",f"{summary['non_perm_assigned_hrs']:.1f} h")

    # Full results table
    st.subheader("Assignment Results (per worker, weekly)")
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    st.subheader("Open Shifts / Uncovered Demand (per project-day)")
    if shortage_df.empty:
        st.success("All project demand was covered.")
    else:
        st.warning(
            "Some project demand remains uncovered. "
            "These positions can be opened for additional registrations."
        )
        st.dataframe(shortage_df, use_container_width=True, hide_index=True)

    # Download button
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Assignments", index=False)
        shortage_df.to_excel(writer, sheet_name="Open_Shifts", index=False)
    buf.seek(0)

    st.download_button(
        label="⬇  Download results as Excel",
        data=buf,
        file_name="assignment_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info(
        "Upload your staff and projects files in the sidebar (or use the "
        "default sample files), set the planning period parameters (for reference), "
        "then click **Run Solver**.\n\n"
        "Note: The optimization itself is based on a single week (Mon–Fri)."
    )

