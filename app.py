import io
from datetime import date
from typing import Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DEFAULT_FILE_PATH = "/mnt/data/CNHI_Silver_Master_Tracker_standardized.xlsx"
MASTER_SHEET_NAME = "Silver_Master_Tracker"
REFERENCE_SHEETS = [
    "Domain_Reference",
    "Activity_Reference",
    "Batch_Reference",
    "Status_Reference",
    "Owner_Capacity",
    "Status_History",
]
STATUS_ORDER_FALLBACK = ["To Do", "In Progress", "Blocked", "Done"]
EXPECTED_COLUMNS = [
    "Task_ID",
    "Domain_ID",
    "Domain_Name",
    "Batch_ID",
    "Activity_ID",
    "Activity_Name",
    "Task_Title",
    "Owner",
    "Priority",
    "Status",
    "Planned_Start",
    "Planned_End",
    "Actual_Start",
    "Actual_End",
    "Effort_Points",
    "RAG",
    "Dependency_IDs",
    "Spec_Link",
    "Last_Updated",
    "Reopen_Flag",
    "Reopen_Count",
]
DATE_COLUMNS = [
    "Planned_Start",
    "Planned_End",
    "Actual_Start",
    "Actual_End",
    "Last_Updated",
]
STRING_COLUMNS = [
    "Task_ID",
    "Domain_ID",
    "Domain_Name",
    "Batch_ID",
    "Activity_ID",
    "Activity_Name",
    "Task_Title",
    "Owner",
    "Priority",
    "Status",
    "RAG",
    "Dependency_IDs",
    "Spec_Link",
    "Reopen_Flag",
]
NUMERIC_COLUMNS = ["Effort_Points", "Reopen_Count"]
ACTIVE_STATUSES = {"In Progress", "Blocked"}
OVERDUE_STATUSES = {"To Do", "In Progress", "Blocked"}


@st.cache_data(show_spinner=False)
def load_data(source: Union[str, io.BytesIO]):
    warnings: List[str] = []
    try:
        excel_file = pd.ExcelFile(source, engine="openpyxl")
    except FileNotFoundError:
        return pd.DataFrame(), {}, [
            "Default tracker file not found at the configured path.",
        ]
    except ValueError:
        return pd.DataFrame(), {}, [
            "Unable to read the tracker. Please upload a valid Excel file.",
        ]
    except Exception as exc:  # pragma: no cover - defensive
        return pd.DataFrame(), {}, [f"Failed to load workbook: {exc}"]

    if MASTER_SHEET_NAME not in excel_file.sheet_names:
        return pd.DataFrame(), {}, [
            f"Workbook missing required sheet '{MASTER_SHEET_NAME}'.",
        ]

    df_tasks = excel_file.parse(MASTER_SHEET_NAME)
    df_tasks.columns = [str(col).strip() for col in df_tasks.columns]

    for column in EXPECTED_COLUMNS:
        if column not in df_tasks.columns:
            if column in DATE_COLUMNS:
                df_tasks[column] = pd.NaT
            elif column in NUMERIC_COLUMNS:
                df_tasks[column] = np.nan
            else:
                df_tasks[column] = ""
            warnings.append(f"Column '{column}' missing in tracker. Filled with blanks.")

    for column in DATE_COLUMNS:
        if column in df_tasks.columns:
            df_tasks[column] = pd.to_datetime(df_tasks[column], errors="coerce")

    refs: Dict[str, pd.DataFrame] = {}
    for sheet in REFERENCE_SHEETS:
        if sheet in excel_file.sheet_names:
            refs[sheet] = excel_file.parse(sheet)

    status_order = STATUS_ORDER_FALLBACK
    status_reference = refs.get("Status_Reference")
    if status_reference is not None and not status_reference.empty:
        ref_cols = {col.lower(): col for col in status_reference.columns}
        status_col = ref_cols.get("status") or ref_cols.get("status_name")
        sort_col = ref_cols.get("sort_order")
        if status_col:
            status_reference = status_reference.copy()
            status_reference[status_col] = status_reference[status_col].astype(str).str.strip()
            if sort_col and sort_col in status_reference.columns:
                status_reference = status_reference.sort_values(sort_col)
            status_order = (
                status_reference[status_col]
                .dropna()
                .tolist()
            ) or STATUS_ORDER_FALLBACK

    df_tasks["Status"] = (
        df_tasks["Status"].fillna("").astype(str).str.strip().replace({"": "To Do"})
    )
    df_tasks["Status"] = pd.Categorical(
        df_tasks["Status"],
        categories=status_order,
        ordered=True,
    )

    for column in STRING_COLUMNS:
        if column in df_tasks.columns:
            df_tasks[column] = df_tasks[column].fillna("").astype(str).str.strip()

    return df_tasks, refs, warnings


def get_data_source(uploaded_file: Optional[object]) -> Union[str, io.BytesIO]:
    if uploaded_file is not None:
        return io.BytesIO(uploaded_file.getvalue())
    return DEFAULT_FILE_PATH


def apply_filters(
    df: pd.DataFrame,
    *,
    selected_batches: List[str],
    selected_domains: List[str],
    selected_activities: List[str],
    selected_owners: List[str],
    selected_priorities: List[str],
    selected_statuses: List[str],
    planned_start: Optional[pd.Timestamp],
    planned_end: Optional[pd.Timestamp],
    actual_start: Optional[pd.Timestamp],
    actual_end: Optional[pd.Timestamp],
    show_overdue: bool,
    show_blocked_only: bool,
    include_no_dates: bool,
) -> pd.DataFrame:
    filtered = df.copy()

    if selected_batches:
        filtered = filtered[filtered["Batch_ID"].isin(selected_batches)]
    if selected_domains:
        filtered = filtered[filtered["Domain_Name"].isin(selected_domains)]
    if selected_activities:
        filtered = filtered[filtered["Activity_Name"].isin(selected_activities)]
    if selected_owners:
        filtered = filtered[filtered["Owner"].isin(selected_owners)]
    if selected_priorities:
        filtered = filtered[filtered["Priority"].isin(selected_priorities)]
    if selected_statuses:
        filtered = filtered[filtered["Status"].isin(selected_statuses)]

    if show_overdue:
        filtered = filtered[filtered["is_overdue"]]
    if show_blocked_only:
        filtered = filtered[filtered["Status"] == "Blocked"]

    def in_range(series: pd.Series, start: Optional[pd.Timestamp], end: Optional[pd.Timestamp]):
        if start is None and end is None:
            return pd.Series(True, index=series.index)
        mask = pd.Series(True, index=series.index)
        if start is not None:
            mask &= series.ge(start)
        if end is not None:
            mask &= series.le(end)
        if include_no_dates:
            mask |= series.isna()
        return mask

    if "Planned_End" in filtered.columns:
        mask = in_range(filtered["Planned_End"], planned_start, planned_end)
        filtered = filtered[mask]
    if "Actual_End" in filtered.columns:
        mask = in_range(filtered["Actual_End"], actual_start, actual_end)
        filtered = filtered[mask]

    return filtered


def compute_kpis(df: pd.DataFrame, today: pd.Timestamp) -> Dict[str, float]:
    total = len(df)
    done_df = df[df["Status"] == "Done"]
    done_count = len(done_df)
    blocked_count = int((df["Status"] == "Blocked").sum())

    percent_complete = (done_count / total * 100) if total else 0.0
    on_time_numerator = done_df[
        (done_df["Actual_End"].notna())
        & (done_df["Planned_End"].notna())
        & (done_df["Actual_End"] <= done_df["Planned_End"])
    ]
    on_time_delivery = (len(on_time_numerator) / done_count * 100) if done_count else 0.0

    aging_df = df[df["Status"].isin(ACTIVE_STATUSES) & df["Actual_Start"].notna()].copy()
    if not aging_df.empty:
        aging_df["aging_days"] = (today - aging_df["Actual_Start"].dt.normalize()).dt.days
        aging_wip = float(aging_df["aging_days"].median())
    else:
        aging_wip = 0.0

    throughput_window = today - pd.Timedelta(days=14)
    throughput = int(
        done_df[
            (done_df["Actual_End"].notna())
            & (done_df["Actual_End"] >= throughput_window)
            & (done_df["Actual_End"] <= today)
        ].shape[0]
    )

    return {
        "total": total,
        "percent_complete": round(percent_complete, 1),
        "on_time_delivery": round(on_time_delivery, 1),
        "blocked": blocked_count,
        "aging_wip": round(aging_wip, 1),
        "throughput": throughput,
    }


def render_kpi_cards(kpis: Dict[str, float]):
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Total Tasks", f"{kpis['total']:,}")
    col2.metric("% Complete", f"{kpis['percent_complete']:.0f}%")
    col3.metric("On-Time Delivery", f"{kpis['on_time_delivery']:.0f}%")
    col4.metric("Blocked", f"{kpis['blocked']:,}")
    col5.metric("Aging WIP (days)", f"{kpis['aging_wip']:.1f}")
    col6.metric("Throughput (14d)", f"{kpis['throughput']:,}")


def progress_by_domain_chart(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()

    working = df.copy()
    working["Domain_Name"] = working["Domain_Name"].replace({"": "Unspecified"})

    status_categories = (
        [str(cat) for cat in working["Status"].cat.categories]
        if hasattr(working["Status"], "cat")
        else STATUS_ORDER_FALLBACK
    )

    domain_counts = (
        working.groupby(["Domain_Name", "Status"], dropna=False)
        .size()
        .reset_index(name="Task Count")
    )
    totals = (
        working.groupby("Domain_Name")
        .agg(
            total_tasks=("Task_ID", "size"),
            done_tasks=("Status", lambda x: (x == "Done").sum()),
            overdue_tasks=("is_overdue", "sum"),
        )
        .reset_index()
    )
    domain_counts = domain_counts.merge(totals, on="Domain_Name", how="left")
    domain_counts["Percent Complete"] = np.where(
        domain_counts["total_tasks"] > 0,
        (domain_counts["done_tasks"] / domain_counts["total_tasks"]) * 100,
        0,
    )
    domain_counts["Overdue"] = domain_counts["overdue_tasks"]

    fig = px.bar(
        domain_counts,
        x="Domain_Name",
        y="Task Count",
        color="Status",
        category_orders={"Status": status_categories},
        hover_data={
            "Percent Complete": ":.1f",
            "Overdue": True,
            "total_tasks": False,
            "done_tasks": False,
            "overdue_tasks": False,
        },
    )
    fig.update_layout(
        legend_title="Status",
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis_title="Domain",
        yaxis_title="Tasks",
    )
    return fig


def owner_workload_chart(df: pd.DataFrame) -> go.Figure:
    active_df = df[df["Status"].isin(ACTIVE_STATUSES)].copy()
    if active_df.empty:
        return go.Figure()

    active_df["Owner"] = active_df["Owner"].replace({"": "Unassigned"})
    active_df["Priority"] = active_df["Priority"].replace({"": "Unspecified"})

    workload = (
        active_df.groupby(["Owner", "Priority"], dropna=False)
        .size()
        .reset_index(name="Task Count")
    )
    fig = px.bar(
        workload,
        x="Owner",
        y="Task Count",
        color="Priority",
        barmode="stack",
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis_title="Owner",
        yaxis_title="Active Tasks",
    )
    return fig


def mini_gantt_chart(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()

    planned_df = df.dropna(subset=["Planned_Start", "Planned_End"])
    actual_df = df.dropna(subset=["Actual_Start", "Actual_End"])

    if planned_df.empty and actual_df.empty:
        return go.Figure()

    def label_series(frame: pd.DataFrame) -> pd.Series:
        return frame["Task_Title"].replace({"": np.nan}).fillna(frame["Task_ID"])

    rows = []
    if not planned_df.empty:
        planned_block = planned_df.copy()
        planned_block["Timeline_Type"] = "Planned"
        planned_block["Start"] = planned_block["Planned_Start"]
        planned_block["Finish"] = planned_block["Planned_End"]
        planned_block["ColorKey"] = "Planned"
        planned_block["Label"] = label_series(planned_block)
        rows.append(planned_block)
    if not actual_df.empty:
        actual_block = actual_df.copy()
        actual_block["Timeline_Type"] = "Actual"
        actual_block["Start"] = actual_block["Actual_Start"]
        actual_block["Finish"] = actual_block["Actual_End"]
        actual_block["ColorKey"] = actual_block["Status"].astype(str)
        actual_block["Label"] = label_series(actual_block)
        rows.append(actual_block)

    combined = pd.concat(rows, ignore_index=True)
    if combined.empty:
        return go.Figure()

    fig = px.timeline(
        combined,
        x_start="Start",
        x_end="Finish",
        y="Label",
        color="ColorKey",
        hover_data={"Domain_Name": True, "Status": True, "Timeline_Type": True},
        color_discrete_map={"Planned": "#d3d3d3"},
    )

    overdue_actual = actual_df[actual_df["is_overdue"]]
    if not overdue_actual.empty:
        overdue_trace = go.Scatter(
            x=overdue_actual["Actual_End"],
            y=label_series(overdue_actual),
            mode="markers",
            marker=dict(symbol="x", size=10, color="red"),
            name="Overdue",
            hoverinfo="text",
            hovertext=[
                f"{row.Task_ID} overdue ({row.Actual_End.date() if pd.notnull(row.Actual_End) else 'n/a'})"
                for row in overdue_actual.itertuples()
            ],
        )
        fig.add_trace(overdue_trace)

    fig.update_layout(
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis_title="Timeline",
        yaxis_title="Tasks",
        showlegend=True,
    )
    fig.update_yaxes(autorange="reversed")
    return fig


def throughput_trend_chart(df: pd.DataFrame) -> go.Figure:
    done_df = df[(df["Status"] == "Done") & df["Actual_End"].notna()]
    if done_df.empty:
        return go.Figure()

    done_df = done_df.copy()
    done_df["Actual_End_Date"] = done_df["Actual_End"].dt.date
    counts = (
        done_df.groupby("Actual_End_Date")
        .size()
        .sort_index()
        .rename("Count")
        .to_frame()
    )
    counts["Rolling_7"] = counts["Count"].rolling(7, min_periods=1).mean()
    counts = counts.reset_index()

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=counts["Actual_End_Date"],
            y=counts["Count"],
            name="Completed",
            marker_color="#4c78a8",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=counts["Actual_End_Date"],
            y=counts["Rolling_7"],
            name="7d Avg",
            mode="lines",
            line=dict(color="#f58518", width=3),
        )
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis_title="Date",
        yaxis_title="Tasks Completed",
        barmode="overlay",
    )
    return fig


def overdue_heatmap(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()

    overdue_df = df[df["is_overdue"]].copy()
    if overdue_df.empty:
        return go.Figure()

    overdue_df["Domain_Display"] = overdue_df["Domain_Name"].replace({"": "Unspecified"})
    overdue_df["Activity_Display"] = overdue_df["Activity_Name"].replace({"": "Unspecified"})

    matrix = overdue_df.pivot_table(
        index="Domain_Display",
        columns="Activity_Display",
        values="Task_ID",
        aggfunc="count",
        fill_value=0,
    )
    if matrix.empty:
        return go.Figure()

    fig = px.imshow(
        matrix,
        aspect="auto",
        color_continuous_scale="Reds",
        labels=dict(color="Overdue Tasks"),
    )
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10))
    return fig


def compute_dependency_risk(
    full_df: pd.DataFrame, filtered_df: pd.DataFrame, today: pd.Timestamp
) -> Tuple[int, List[str]]:
    if filtered_df.empty:
        return 0, []

    status_map = full_df.set_index("Task_ID")["Status"].astype(str).to_dict()
    risky_tasks: List[str] = []
    for row in filtered_df.itertuples():
        deps_raw = getattr(row, "Dependency_IDs", "")
        if not deps_raw:
            continue
        dependencies = [dep.strip() for dep in str(deps_raw).split(";") if dep.strip()]
        if not dependencies:
            continue
        planned_start = getattr(row, "Planned_Start", pd.NaT)
        actual_start = getattr(row, "Actual_Start", pd.NaT)
        should_start = False
        if pd.notna(planned_start) and planned_start <= today:
            should_start = True
        if pd.notna(actual_start):
            should_start = True
        if not should_start:
            continue
        for dep in dependencies:
            status = status_map.get(dep)
            if status != "Done":
                risky_tasks.append(row.Task_ID or row.Task_Title or str(row.Index))
                break
    return len(risky_tasks), risky_tasks[:10]


def compute_queue_health(df: pd.DataFrame) -> Tuple[str, float]:
    to_do = int((df["Status"] == "To Do").sum())
    in_progress = int((df["Status"] == "In Progress").sum())
    ratio = in_progress if in_progress else 1
    numeric = to_do / ratio
    return f"{to_do} : {in_progress}", round(numeric, 2)


def render_task_table(df: pd.DataFrame):
    if df.empty:
        st.info("No tasks match the current filters.")
        return

    table_cols = [
        "Task_ID",
        "Task_Title",
        "Domain_Name",
        "Activity_Name",
        "Owner",
        "Priority",
        "Status",
        "Planned_End",
        "Actual_End",
    ]
    available_cols = [col for col in table_cols if col in df.columns]
    display_df = df[available_cols].copy()
    display_df["Overdue"] = np.where(df["is_overdue"], "⚠️ Overdue", "✅ On Track")
    if "Spec_Link" in df.columns:
        display_df["Spec_Link"] = df["Spec_Link"].apply(
            lambda link: f"[Link]({link})" if isinstance(link, str) and link else ""
        )

    for date_col in ["Planned_End", "Actual_End"]:
        if date_col in display_df.columns:
            display_df[date_col] = display_df[date_col].dt.date

    st.dataframe(display_df, use_container_width=True, hide_index=True)


def to_timestamp(value: Optional[Union[date, pd.Timestamp]]) -> Optional[pd.Timestamp]:
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        return value
    return pd.to_datetime(value)


def main():
    st.set_page_config(page_title="Workstream KPI Dashboard", layout="wide")
    st.title("Workstream KPI Dashboard")
    st.caption("Live view of workstream throughput, blockers, and delivery performance.")

    today = pd.Timestamp.today().normalize()

    with st.sidebar:
        st.header("Filters")
        uploaded = st.file_uploader("Upload tracker", type=["xlsx"])
        data_source = get_data_source(uploaded)

        df_tasks, refs, warnings = load_data(data_source)

        if warnings:
            for warn in warnings:
                st.warning(warn)

        if df_tasks.empty:
            st.error("No tasks available. Please provide a valid tracker file.")
            return

        batches = sorted([val for val in df_tasks["Batch_ID"].unique() if val])
        domains = sorted([val for val in df_tasks["Domain_Name"].unique() if val])
        activities = sorted([val for val in df_tasks["Activity_Name"].unique() if val])
        owners = sorted([val for val in df_tasks["Owner"].unique() if val])
        priorities = sorted([val for val in df_tasks["Priority"].unique() if val])
        statuses = [str(status) for status in df_tasks["Status"].cat.categories]

        selected_batches = st.multiselect("Batch", batches)
        selected_domains = st.multiselect("Domain", domains)
        selected_activities = st.multiselect("Activity", activities)
        selected_owners = st.multiselect("Owner", owners)
        selected_priorities = st.multiselect("Priority", priorities)
        selected_statuses = st.multiselect("Status", statuses, default=statuses)

        st.markdown("### Date Filters")

        if "actual_range" not in st.session_state:
            st.session_state.actual_range = (
                (today - pd.Timedelta(days=30)).date(),
                today.date(),
            )

        quick_cols = st.columns(3)
        if quick_cols[0].button("Last 7"):
            st.session_state.actual_range = (
                (today - pd.Timedelta(days=7)).date(),
                today.date(),
            )
        if quick_cols[1].button("Last 14"):
            st.session_state.actual_range = (
                (today - pd.Timedelta(days=14)).date(),
                today.date(),
            )
        if quick_cols[2].button("Last 30"):
            st.session_state.actual_range = (
                (today - pd.Timedelta(days=30)).date(),
                today.date(),
            )

        apply_planned_range = st.checkbox("Apply Planned End range", value=False)
        planned_start_default = (today - pd.Timedelta(days=30)).date()
        planned_end_default = today.date()
        planned_start_input = st.date_input(
            "Planned End from",
            value=planned_start_default,
            key="planned_start_date",
        )
        planned_end_input = st.date_input(
            "Planned End to",
            value=planned_end_default,
            key="planned_end_date",
        )

        actual_start_default, actual_end_default = st.session_state.actual_range
        actual_start_input = st.date_input(
            "Actual End from",
            value=actual_start_default,
            key="actual_start_date",
        )
        actual_end_input = st.date_input(
            "Actual End to",
            value=actual_end_default,
            key="actual_end_date",
        )
        if isinstance(actual_start_input, date) and isinstance(actual_end_input, date):
            st.session_state.actual_range = (actual_start_input, actual_end_input)

        apply_actual_range = st.checkbox("Apply Actual End range", value=True)

        show_overdue = st.checkbox("Show only overdue", value=False)
        show_blocked_only = st.checkbox("Show only blocked", value=False)
        include_no_dates = st.checkbox("Include tasks without dates", value=True)

    df_tasks = df_tasks.copy()
    df_tasks["is_overdue"] = (
        df_tasks["Planned_End"].lt(today)
        & df_tasks["Status"].isin(OVERDUE_STATUSES)
    )
    df_tasks["is_active"] = df_tasks["Status"].isin(ACTIVE_STATUSES)

    planned_start = to_timestamp(planned_start_input) if apply_planned_range else None
    planned_end = to_timestamp(planned_end_input) if apply_planned_range else None
    actual_start = to_timestamp(actual_start_input) if apply_actual_range else None
    actual_end = to_timestamp(actual_end_input) if apply_actual_range else None

    filtered_df = apply_filters(
        df_tasks,
        selected_batches=selected_batches,
        selected_domains=selected_domains,
        selected_activities=selected_activities,
        selected_owners=selected_owners,
        selected_priorities=selected_priorities,
        selected_statuses=selected_statuses,
        planned_start=planned_start,
        planned_end=planned_end,
        actual_start=actual_start,
        actual_end=actual_end,
        show_overdue=show_overdue,
        show_blocked_only=show_blocked_only,
        include_no_dates=include_no_dates,
    )

    kpis = compute_kpis(filtered_df, today)
    render_kpi_cards(kpis)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.subheader("Progress by Domain")
        fig_domain = progress_by_domain_chart(filtered_df)
        st.plotly_chart(fig_domain, use_container_width=True)

    with chart_col2:
        st.subheader("Owner Workload")
        fig_owner = owner_workload_chart(filtered_df)
        st.plotly_chart(fig_owner, use_container_width=True)

    chart_col3, chart_col4 = st.columns(2)
    with chart_col3:
        st.subheader("Plan vs Actual Timeline")
        fig_gantt = mini_gantt_chart(filtered_df)
        st.plotly_chart(fig_gantt, use_container_width=True)

    with chart_col4:
        st.subheader("Throughput Trend")
        fig_throughput = throughput_trend_chart(filtered_df)
        st.plotly_chart(fig_throughput, use_container_width=True)

    st.subheader("Risk & Quality")
    risk_col1, risk_col2, risk_col3 = st.columns(3)

    heatmap_fig = overdue_heatmap(filtered_df)
    with risk_col1:
        st.caption("Overdue by Domain & Activity")
        if heatmap_fig.data:
            st.plotly_chart(heatmap_fig, use_container_width=True)
        else:
            st.info("No overdue tasks in the current view.")

    dependency_count, risky_samples = compute_dependency_risk(df_tasks, filtered_df, today)
    queue_ratio, queue_numeric = compute_queue_health(filtered_df)

    with risk_col2:
        st.metric("Dependency Risks", dependency_count)
        if risky_samples:
            with st.expander("At-risk tasks", expanded=False):
                for task in risky_samples:
                    st.write(task)
        else:
            st.caption("All dependencies appear clear for the filtered tasks.")

    with risk_col3:
        st.metric("Queue Health (To Do : In Progress)", queue_ratio)
        st.caption(f"Ratio value: {queue_numeric}")

    st.subheader("Task Details")
    render_task_table(filtered_df)

    csv_data = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered tasks",
        data=csv_data,
        file_name="filtered_tasks.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
