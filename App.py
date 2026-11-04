
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# -------------------------
# Page configuration
# -------------------------
st.set_page_config(
    page_title="Mobile Usage Analysis",
    page_icon="📱",
    layout="wide"
)

st.title("📱 Mobile Phone Usage Pattern Analysis")
st.caption("Interactive dashboard for analyzing student mobile usage patterns")

DEFAULT_FILE = "Mobile_Phone_Usage_Pattern_Analysis_For_60_Students.xlsx"
DATA_FILE = "student_data.csv"

TIME_OPTIONS = [
    "<1H", "1-2H", "2-3H", "2-4H", "4-6H",
    "6-8H", "8-9H", "9-10H", "10H+", "None"
]

SCORE_OPTIONS = [
    "<40%", "40-60%", "61-80%", "80-85%", "85-90%", ">90%"
]


# -------------------------
# Load and clean data
# -------------------------
@st.cache_data
def load_data(file):
    if file is not None:
        raw = pd.read_excel(file, header=5)
    else:
        if not Path(DEFAULT_FILE).exists():
            return pd.DataFrame()
        raw = pd.read_excel(DEFAULT_FILE, header=5)

    raw.columns = [str(col).strip() for col in raw.columns]

    rename_map = {
        "No": "Student_ID",
        "Student Name": "Student_Name",
        "Semister": "Semester",
        "Faculty": "Faculty",
        "Age": "Age",
        "Daily Screen Time": "Screen_Time",
        "Socail Media Time": "Social_Media_Time",
        "Social Media Time": "Social_Media_Time",
        "Gameing Time": "Gaming_Time",
        "Gaming Time": "Gaming_Time",
        "Study Time": "Study_Time",
        "Sleep Time": "Sleep_Time",
        "Number of Apps": "Number_of_Apps",
        "Academic Score": "Academic_Score",
    }

    raw = raw.rename(columns=rename_map)

    required_columns = [
        "Student_ID", "Student_Name", "Semester", "Faculty", "Age",
        "Screen_Time", "Social_Media_Time", "Gaming_Time",
        "Study_Time", "Sleep_Time", "Number_of_Apps", "Academic_Score"
    ]

    available_columns = [
        col for col in required_columns if col in raw.columns
    ]

    data = raw[available_columns].copy()

    if "Student_ID" in data.columns:
        data = data.dropna(subset=["Student_ID"])

    for col in data.select_dtypes(include="object").columns:
        data[col] = data[col].astype("string").str.strip()
        data[col] = data[col].replace("", pd.NA)

    return data.reset_index(drop=True)


# -------------------------
# Persistent data loading
# -------------------------
if "student_data" not in st.session_state:
    if Path(DATA_FILE).exists():
        st.session_state.student_data = pd.read_csv(DATA_FILE)
    else:
        st.session_state.student_data = load_data(None)

if "next_student_id" not in st.session_state:
    existing_ids = pd.to_numeric(
        st.session_state.student_data.get(
            "Student_ID", pd.Series(dtype="float")
        ),
        errors="coerce"
    )
    st.session_state.next_student_id = (
        int(existing_ids.max()) + 1
        if existing_ids.notna().any()
        else 1
    )


# -------------------------
# Sidebar: upload data
# -------------------------
st.sidebar.header("Data Source")

uploaded_file = st.sidebar.file_uploader(
    "Upload an Excel file",
    type=["xlsx"]
)

if uploaded_file is not None:
    uploaded_data = load_data(uploaded_file)

    if not uploaded_data.empty:
        st.session_state.student_data = uploaded_data

        ids = pd.to_numeric(
            uploaded_data.get("Student_ID", pd.Series(dtype="float")),
            errors="coerce"
        )
        st.session_state.next_student_id = (
            int(ids.max()) + 1 if ids.notna().any() else 1
        )

        # Save uploaded data as the new persistent dataset
        uploaded_data.to_csv(
            DATA_FILE,
            index=False,
            encoding="utf-8-sig"
        )

data = st.session_state.student_data.copy()
st.sidebar.success(f"{len(data)} student records loaded")


# -------------------------
# Add student form
# -------------------------
with st.sidebar.expander("➕ Add Student", expanded=False):
    with st.form("add_student_form", clear_on_submit=True):
        student_name = st.text_input("Student Name")
        semester = st.text_input("Semester")
        faculty = st.text_input("Faculty")
        age = st.number_input("Age", min_value=10, max_value=100, value=20)

        screen_time = st.selectbox("Daily Screen Time", TIME_OPTIONS)
        social_time = st.selectbox("Social Media Time", TIME_OPTIONS)
        gaming_time = st.selectbox("Gaming Time", TIME_OPTIONS)
        study_time = st.selectbox("Study Time", TIME_OPTIONS)
        sleep_time = st.selectbox("Sleep Time", TIME_OPTIONS)

        number_of_apps = st.number_input(
            "Number of Apps",
            min_value=0,
            max_value=1000,
            value=10
        )

        academic_score = st.selectbox("Academic Score", SCORE_OPTIONS)

        submitted = st.form_submit_button(
            "Add Student",
            use_container_width=True
        )

        if submitted:
            if not student_name.strip():
                st.error("Please enter the student's name.")
            elif not semester.strip():
                st.error("Please enter the semester.")
            elif not faculty.strip():
                st.error("Please enter the faculty.")
            else:
                new_student = {
                    "Student_ID": st.session_state.next_student_id,
                    "Student_Name": student_name.strip(),
                    "Semester": semester.strip(),
                    "Faculty": faculty.strip(),
                    "Age": age,
                    "Screen_Time": screen_time,
                    "Social_Media_Time": social_time,
                    "Gaming_Time": gaming_time,
                    "Study_Time": study_time,
                    "Sleep_Time": sleep_time,
                    "Number_of_Apps": number_of_apps,
                    "Academic_Score": academic_score,
                }

                st.session_state.student_data = pd.concat(
                    [
                        st.session_state.student_data,
                        pd.DataFrame([new_student])
                    ],
                    ignore_index=True
                )

                # Persist data to CSV
                st.session_state.student_data.to_csv(
                    DATA_FILE,
                    index=False,
                    encoding="utf-8-sig"
                )

                st.session_state.next_student_id += 1
                st.success("Student added and saved successfully!")
                st.rerun()


# -------------------------
# KPI cards
# -------------------------
data = st.session_state.student_data.copy()

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Students", len(data))

if "Age" in data.columns:
    ages = pd.to_numeric(data["Age"], errors="coerce")
    col2.metric(
        "Average Age",
        f"{ages.mean():.1f}" if ages.notna().any() else "N/A"
    )
else:
    col2.metric("Average Age", "N/A")

if "Faculty" in data.columns:
    col3.metric("Faculties", data["Faculty"].nunique())
else:
    col3.metric("Faculties", "N/A")

col4.metric("Missing Cells", int(data.isna().sum().sum()))

st.divider()


# -------------------------
# Chart helper
# -------------------------
def show_bar_chart(dataframe, column, title, color):
    if column not in dataframe.columns:
        st.info(f"{column} is not available.")
        return

    counts = dataframe[column].value_counts(dropna=False)

    fig, ax = plt.subplots(figsize=(7, 4))
    counts.plot(kind="bar", ax=ax, color=color)

    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("Number of Students")
    ax.tick_params(axis="x", rotation=25)

    fig.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


# -------------------------
# Charts
# -------------------------
left, right = st.columns(2)

with left:
    st.subheader("Daily Screen Time")
    show_bar_chart(data, "Screen_Time", "Daily Screen Time", "#2878B5")

with right:
    st.subheader("Academic Score Distribution")
    show_bar_chart(
        data, "Academic_Score", "Academic Score", "#2A9D8F"
    )

left, right = st.columns(2)

with left:
    st.subheader("Social Media Time")
    show_bar_chart(
        data, "Social_Media_Time", "Social Media Time", "#E9A23B"
    )

with right:
    st.subheader("Gaming Time")
    show_bar_chart(data, "Gaming_Time", "Gaming Time", "#8E6CCF")


# -------------------------
# Student records
# -------------------------
st.divider()
st.subheader("Student Records")

st.dataframe(
    data,
    use_container_width=True,
    hide_index=True
)

csv_data = data.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    label="Download Current Data as CSV",
    data=csv_data,
    file_name="mobile_usage_data.csv",
    mime="text/csv"
)
