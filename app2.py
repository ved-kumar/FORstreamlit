import streamlit as st
import pandas as pd
import plotly.express as px
import calendar
import openpyxl
import matplotlib.pyplot as plt

# Function to load data from an Excel file
def load_data(file_path, sheet_name):
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except ValueError:  # Handle case where the sheet does not exist
        st.warning(f"Sheet '{sheet_name}' not found in {file_path}. Proceeding without it.")
        return pd.DataFrame()  # Return an empty DataFrame

# Sidebar navigation
st.sidebar.title("Navigation")
app_mode = st.sidebar.radio("Choose an option", ["Month-wise Analytics", "Compare All Months"])

# Month-wise Analytics Section
if app_mode == "Month-wise Analytics":
    selected_month = st.selectbox(
        "Select Month:",
        options=calendar.month_name[1:],
        index=4  # Default to 'May'
    )

    # Define file path based on selected month
    file_path = f'{selected_month}.xlsx'

    sheet_name_loop = 'Loop Tasks'
    sheet_name_sprint = 'Sprint Tasks'

    # Load data from sheets
    df_sprint = load_data(file_path, sheet_name_sprint)
    df_loop = load_data(file_path, sheet_name_loop)

    # Ensure there are no leading or trailing spaces in column names
    if not df_loop.empty:
        df_loop.columns = df_loop.columns.str.strip()
    if not df_sprint.empty:
        df_sprint.columns = df_sprint.columns.str.strip()

    # Filter completed tasks
    if not df_sprint.empty:
        completed_tasks_sprint = df_sprint[df_sprint['Status'] == 'Done']
    if not df_loop.empty:
        completed_tasks_loop = df_loop[df_loop['Status'] == 'Done']

    # Combine both Sprint and Loop data for overall task management
    df = pd.concat([df_loop, df_sprint], ignore_index=True)

    # Proceed with the rest of the processing only if df is not empty
    if not df.empty:
        # Summary Metrics Section
        st.write("### Summary Metrics")
        total_tasks = len(df)
        completed_count = len(df[df['Status'] == 'Done'])
        pending_count = len(df[df['Status'] != 'Done'])

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Tasks", total_tasks)
        col2.metric("Completed Tasks", completed_count)
        col3.metric("Pending Tasks", pending_count)


        # Completed Task Types Section
        st.write("### Completed Task Types")
        task_types = df[df["Status"] == "Done"]["Issue Type"].value_counts().reset_index()
        task_types.columns = ["Task Type", "Count"]

        fig = px.bar(task_types, x="Task Type", y="Count", 
                    title="Completed Task Types", 
                    labels={"Count": "Number of Tasks", "Task Type": "Type"},
                    color="Count",
                    color_continuous_scale="Blues")
        st.plotly_chart(fig)


        # Top Contributors Section
        st.write("### Top Contributors")
        top_contributors = df[df["Status"] == "Done"]["Assignee"].value_counts().reset_index()
        top_contributors.columns = ["Assignee", "Count"]

        fig3 = px.bar(top_contributors, x="Assignee", y="Count", 
                    title="Top Contributors (Completed Tasks)", 
                    labels={"Count": "Number of Tasks", "Assignee": "Contributor"},
                    color="Count",
                    color_continuous_scale="Greens")
        st.plotly_chart(fig3)

        # Filter by Assignee Section
        st.write("### Filter by Assignee")
        assignee = st.selectbox("Select an Assignee", options=df["Assignee"].dropna().unique())

        assignee_tasks = df[df["Assignee"] == assignee]
        assignee_tasks_completed = assignee_tasks[assignee_tasks["Status"] == "Done"]
        assignee_tasks_pending = assignee_tasks[assignee_tasks["Status"] != "Done"]

        # Display task counts for the selected assignee
        completed_count = len(assignee_tasks_completed)
        pending_count = len(assignee_tasks_pending)
        total_count = len(assignee_tasks)

        st.write(f"**Summary for {assignee}:**")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Tasks", total_count)
        col2.metric("Completed Tasks", completed_count)
        col3.metric("Pending Tasks", pending_count)

        # Task Breakdown by Status for the Selected Assignee
        st.subheader(f"Task Status Breakdown for {assignee}")
        if not assignee_tasks.empty:
            task_status_breakdown = assignee_tasks["Status"].value_counts().reset_index()
            task_status_breakdown.columns = ["Status", "Count"]

            fig = px.bar(task_status_breakdown, x="Status", y="Count",
                        title=f"Task Status Breakdown for {assignee}",
                        labels={"Count": "Number of Tasks", "Status": "Task Status"},
                        color="Count",
                        color_continuous_scale="Reds")
            st.plotly_chart(fig)


        # Pending Task Details Section
        st.write("### Pending Task Details")
        if not assignee_tasks_pending.empty:
            st.write("#### Pending Tasks")
            st.dataframe(assignee_tasks_pending[["Summary", "Issue Type", "Status", "Date"]])
        else:
            st.write(f"No pending tasks for {assignee}.")

        # Completed Task Details Section
        st.write("#### Completed Task Details")
        if not assignee_tasks_completed.empty:
            st.write("### Completed Tasks")
            st.dataframe(assignee_tasks_completed[["Summary", "Issue Type", "Status", "Date"]])
        else:
            st.write(f"No completed tasks for {assignee}.")

else:  # Compare All Months Section
    st.write("#### Compare All Months Progress")

    uploaded_files = st.file_uploader("Upload monthly data files", accept_multiple_files=True, type=["xlsx"])

    if uploaded_files:
        all_month_data = []
        month_names = []

        for uploaded_file in uploaded_files:
            # Extract month name from file name (assuming it's named as '<Month>.xlsx')
            month_name = uploaded_file.name.split(".")[0]
            month_names.append(month_name)

            # Load data for the uploaded file
            monthly_data_sprint = load_data(uploaded_file, sheet_name_sprint)
            monthly_data_loop = load_data(uploaded_file, sheet_name_loop)

            # Preprocess and combine monthly data
            if not monthly_data_sprint.empty:
                monthly_data_sprint['Tasks Type'] = 'Sprint'
            if not monthly_data_loop.empty:
                monthly_data_loop['Tasks Type'] = 'Loop'

            monthly_data = pd.concat([monthly_data_sprint, monthly_data_loop], ignore_index=True)
            monthly_data['Month'] = month_name  # Add month name for identification
            all_month_data.append(monthly_data)

        # Combine all months' data into one DataFrame
        combined_data = pd.concat(all_month_data, ignore_index=True)

        # Clean and preprocess data (similar to earlier steps)
        combined_data['Resource Name'] = combined_data['Resource Name'].replace(replace_dict)
        combined_data['Resource Name'] = combined_data['Resource Name'].str.strip()
        combined_data['Status'] = combined_data['Status'].str.strip()

        # Aggregate data by month
        monthly_summary = combined_data.groupby('Month').agg(
            Total_Tasks=('Tasks List', 'count'),
            Completed_Tasks=('Status', lambda x: (x == 'Done').sum()),
            Pending_Tasks=('Status', lambda x: (x != 'Done').sum())
        ).reset_index()

        # Display summary metrics in a table
        st.write("#### Monthly Task Summary")
        st.dataframe(monthly_summary)

        # Visualize month-wise progress
        st.write("#### Month-Wise Progress Charts")
        fig_monthly_progress = px.bar(
            monthly_summary,
            x='Month',
            y=['Completed_Tasks', 'Pending_Tasks'],
            barmode='group',
            title="Monthly Progress Overview",
            labels={'value': 'Task Count', 'variable': 'Task Status'}
        )
        st.plotly_chart(fig_monthly_progress)

        # Allow filtering by specific months for deeper analysis
        selected_months = st.multiselect(
            "Filter by Month(s)",
            options=month_names,
            default=month_names
        )
        filtered_data = combined_data[combined_data['Month'].isin(selected_months)]

        # Visualize filtered data
        st.write("#### Filtered Data Analysis")
        fig_filtered_status = px.bar(
            filtered_data.groupby(['Month', 'Status']).size().reset_index(name='Count'),
            x='Month',
            y='Count',
            color='Status',
            title="Task Status by Month (Filtered)",
            barmode='stack'
        )
        st.plotly_chart(fig_filtered_status)

        # Resource-wise progress across selected months
        st.write("#### Resource-Wise Task Completion")
        fig_resource_month = px.line(
            filtered_data[filtered_data['Status'] == 'Done'].groupby(['Month', 'Resource Name']).size().reset_index(name='Completed_Tasks'),
            x='Month',
            y='Completed_Tasks',
            color='Resource Name',
            title="Resource-Wise Task Completion Across Months"
        )
        st.plotly_chart(fig_resource_month)

    else:
        st.info("Upload multiple Excel files to compare monthly progress.")
