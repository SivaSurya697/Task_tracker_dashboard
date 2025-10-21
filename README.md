# Workstream KPI Dashboard

A single-page Streamlit application that visualizes project tasks from an Excel tracker. The dashboard is designed for program and project managers who need fast insight into delivery progress, workload distribution, and risks across domains, activities, and owners.

## Features

- **Excel ingestion with caching** via `@st.cache_data`, including graceful handling of missing columns and optional reference sheets.
- **Sidebar controls** for file upload, multi-select filters, date ranges, overdue/blocked toggles, and quick lookback buttons.
- **KPI summary cards** that surface completion, on-time delivery, blocked load, WIP aging, and recent throughput.
- **Interactive Plotly visuals**: stacked status by domain, owner workload, plan vs. actual mini-gantt, throughput trend, and an overdue heatmap.
- **Risk diagnostics** covering dependency risk counts and queue health ratios.
- **Downloadable task table** that reflects all active filters.

## Requirements

- Python 3.10+
- Streamlit
- pandas
- numpy
- plotly
- openpyxl (for Excel parsing)

Install the dependencies with pip:

```bash
pip install -r requirements.txt
```

If a `requirements.txt` file is not available, install the core libraries directly:

```bash
pip install streamlit pandas numpy plotly openpyxl
```

## Getting Started

1. Place the standardized tracker in the default location (`/mnt/data/CNHI_Silver_Master_Tracker_standardized.xlsx`) or prepare an Excel file with the same structure.
2. Launch the app from the repository root:

   ```bash
   streamlit run app.py
   ```

3. Use the sidebar to upload an alternate tracker if needed and adjust filters to explore the dataset.

## Data Expectations

The application expects a `Silver_Master_Tracker` worksheet with one row per task. Missing columns are automatically added with blank values so that the dashboard remains responsive. Optional reference sheets—`Domain_Reference`, `Activity_Reference`, `Batch_Reference`, `Status_Reference`, `Owner_Capacity`, and `Status_History`—are read when present to enhance status ordering and capacity metrics.

Status values are normalized to the order `To Do → In Progress → Blocked → Done`, unless a `Status_Reference` sheet specifies an alternate order via a `Sort_Order` column.

## Testing

To perform a lightweight syntax check before deployment, run:

```bash
python -m compileall app.py
```

## Project Structure

```
app.py          # Streamlit dashboard entry point
README.md       # Project overview and usage instructions
```

## Support

If you encounter issues loading the default tracker, the sidebar uploader can be used to provide a local copy. For further enhancements or bug fixes, please open an issue or submit a pull request.
