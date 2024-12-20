import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import calendar
import openpyxl

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import calendar
import openpyxl
import os

# Get the directory where the code file resides
current_directory = os.path.dirname(os.path.abspath(__file__))

# Function to load data from an Excel file
def load_data(file_path, sheet_name):
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except ValueError:  # Handle case where the sheet does not exist
        st.warning(f"Sheet '{sheet_name}' not found in {file_path}. Proceeding without it.")
        return pd.DataFrame()  # Return an empty DataFrame

# List all Excel files in the current directory
files = [f for f in os.listdir(current_directory) if f.endswith(".xlsx")]
months_available = [os.path.splitext(f)[0] for f in files]  # Extract month names

# Streamlit section for month-wise comparison
st.header("Month-Wise Progress Comparison")

if months_available:
    # Multi-select for months
    selected_months = st.multiselect(
        "Select Months to Compare:",
        options=months_available,
        default=months_available[:2]  # Default to first two months for comparison
    )

    if selected_months:
        # Load data for selected months
        monthly_data = {}
        for month in selected_months:
            file_path = os.path.join(current_directory, f"{month}.xlsx")
            sprint_data = load_data(file_path, "Sprint Tasks")
            loop_data = load_data(file_path, "Loop Tasks")

            # Preprocess and combine the data
            if not sprint_data.empty:
                sprint_data["Tasks Type"] = "Sprint"
            if not loop_data.empty:
                loop_data["Tasks Type"] = "Loop"

            combined_data = pd.concat([sprint_data, loop_data], ignore_index=True)
            combined_data["Month"] = month
            monthly_data[month] = combined_data

        # Combine data from all selected months
        all_months_data = pd.concat(monthly_data.values(), ignore_index=True)

        # Clean column names
        all_months_data.columns = all_months_data.columns.str.strip()

        # Drop rows where 'Status' is missing
        if "Status" in all_months_data.columns:
            all_months_data = all_months_data.dropna(subset=["Status"])
        else:
            st.warning("The 'Status' column is missing from the data.")
            all_months_data["Status"] = "Unknown"

        # Task Status Distribution
        st.subheader("Task Status Distribution Across Months")
        try:
            status_distribution = all_months_data.groupby(["Month", "Status"]).size().unstack(fill_value=0)
            fig_status_dist = px.bar(
                status_distribution,
                barmode="stack",
                title="Task Status Distribution Across Months",
                labels={"value": "Count", "Month": "Month", "Status": "Task Status"}
            )
            st.plotly_chart(fig_status_dist)
        except ValueError as e:
            st.error(f"Error in grouping data by Status: {e}")

        # Task Type Distribution
        st.subheader("Task Type Distribution Across Months")
        task_types_dist = all_months_data.groupby(["Month", "Tasks Type"]).size().unstack(fill_value=0)
        fig_task_types = px.bar(
            task_types_dist,
            barmode="stack",
            title="Task Type Distribution Across Months",
            labels={"value": "Count", "Month": "Month", "Tasks Type": "Task Type"}
        )
        st.plotly_chart(fig_task_types)

        # Task Completion Trend
        st.subheader("Task Completion Trend Across Months")
        if "Date" in all_months_data.columns:
            all_months_data["Date"] = pd.to_datetime(all_months_data["Date"], errors="coerce")
            task_trend = all_months_data.groupby([all_months_data["Date"].dt.to_period("M"), "Status"]).size().reset_index(name="Task Count")
            task_trend["Date"] = task_trend["Date"].dt.strftime("%Y-%m")
            fig_task_trend = px.line(
                task_trend,
                x="Date",
                y="Task Count",
                color="Status",
                title="Task Completion Trend Across Months",
                labels={"Task Count": "Number of Tasks", "Date": "Month"}
            )
            st.plotly_chart(fig_task_trend)
        else:
            st.warning("Date column not found in data. Task trend analysis skipped.")

        # Assignee Performance Across Months
        st.subheader("Assignee Performance Comparison Across Months")
        if "Assignee" in all_months_data.columns:
            assignee_performance = all_months_data.groupby(["Month", "Assignee"]).size().unstack(fill_value=0)
            fig_assignee_perf = px.bar(
                assignee_performance,
                barmode="stack",
                title="Assignee Performance Across Months",
                labels={"value": "Tasks Completed", "Month": "Month", "Assignee": "Resource"}
            )
            st.plotly_chart(fig_assignee_perf)
        else:
            st.warning("Assignee column not found in data. Performance comparison skipped.")

    else:
        st.warning("No months selected for comparison.")
else:
    st.warning("No valid Excel files found in the current directory.")
