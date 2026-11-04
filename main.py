
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ==========================================
# 1. PROJECT CONFIGURATION
# ==========================================

INPUT_FILE = "Mobile_Phone_Usage_Pattern_Analysis_For_60_Students.xlsx"

OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# ==========================================
# 2. LOAD EXCEL DATA
# ==========================================

raw = pd.read_excel(INPUT_FILE, header=None)

print("=" * 50)
print("MOBILE PHONE USAGE PATTERN ANALYSIS")
print("=" * 50)

print("\nOriginal Excel shape:", raw.shape)
print("\nFirst rows of the Excel file:")
print(raw.head(10).to_string(index=False, header=False))

# ==========================================
# 3. FIND THE REAL HEADER ROW
# ==========================================

header_row = None

for i in range(min(15, len(raw))):
    row = [
        str(value).strip().lower()
        for value in raw.iloc[i].tolist()
    ]

    has_student_name = any(
        "student" in value and "name" in value
        for value in row
    )

    has_age = "age" in row

    if has_student_name and has_age:
        header_row = i
        break

if header_row is None:
    raise ValueError(
        "Could not find the student data header in Excel."
    )

print("\nHeader row index:", header_row)

# ==========================================
# 4. CREATE DATAFRAME
# ==========================================

data = raw.iloc[header_row + 1:, :12].copy()

data.columns = [
    "Student_ID",
    "Student_Name",
    "Semester",
    "Faculty",
    "Age",
    "Daily_Screen_Time",
    "Social_Media_Time",
    "Gaming_Time",
    "Study_Time",
    "Sleep_Time",
    "Number_of_Apps",
    "Academic_Score"
]

# ==========================================
# 5. CLEAN BASIC DATA
# ==========================================

for column in data.columns:
    data[column] = data[column].apply(
        lambda value: value.strip()
        if isinstance(value, str)
        else value
    )

# Convert ID and Age to numeric values
data["Student_ID"] = pd.to_numeric(
    data["Student_ID"], errors="coerce"
)

data["Age"] = pd.to_numeric(
    data["Age"], errors="coerce"
)

# Remove non-data rows, including repeated headers
data = data.dropna(
    subset=["Student_ID", "Student_Name", "Age"]
).copy()

data = data[
    (data["Student_ID"] % 1 == 0)
    & (data["Age"] % 1 == 0)
].copy()

data["Student_ID"] = data["Student_ID"].astype(int)
data["Age"] = data["Age"].astype(int)

data = data.reset_index(drop=True)

# ==========================================
# 6. STANDARDIZE CATEGORY LABELS
# ==========================================

def normalize_time(value):
    if pd.isna(value):
        return pd.NA

    key = str(value).strip().upper().replace(" ", "")

    labels = {
        "LESS1H": "<1 hour",
        "1-2H": "1-2 hours",
        "2-3H": "2-3 hours",
        "2-4H": "2-4 hours",
        "3-4H": "3-4 hours",
        "4-5H": "4-5 hours",
        "5-6H": "5-6 hours",
        "6-7H": "6-7 hours",
        "7-8H": "7-8 hours",
        "8-9H": "8-9 hours",
        "NONE": "None"
    }

    return labels.get(key, str(value).strip())


time_columns = [
    "Daily_Screen_Time",
    "Social_Media_Time",
    "Gaming_Time",
    "Study_Time",
    "Sleep_Time"
]

for column in time_columns:
    data[column] = data[column].apply(normalize_time)


def normalize_apps(value):
    if pd.isna(value):
        return pd.NA

    key = str(value).strip().upper().replace(" ", "")

    labels = {
        "LESS10": "<10 apps",
        "LESS20": "<20 apps",
        "LESS30": "<30 apps",
        "MORE20": ">20 apps",
        "MORE54": ">54 apps"
    }

    return labels.get(key, str(value).strip())


data["Number_of_Apps"] = data["Number_of_Apps"].apply(
    normalize_apps
)


def normalize_score(value):
    if pd.isna(value):
        return pd.NA

    key = str(value).strip().upper().replace(" ", "")

    labels = {
        "61-80%": "61-80%",
        "80-85%": "80-85%",
        "85-90%": "85-90%",
        "MORE90%": ">90%"
    }

    return labels.get(key, str(value).strip())


data["Academic_Score"] = data["Academic_Score"].apply(
    normalize_score
)

# ==========================================
# 7. DATA QUALITY REPORT
# ==========================================

print("\n" + "=" * 50)
print("DATA QUALITY REPORT")
print("=" * 50)

print("\nValid student records:", len(data))

print("\nMissing values per column:")
print(data.isnull().sum())

print("\nDuplicate student IDs:")
print(data["Student_ID"].duplicated().sum())

print("\nData types:")
print(data.dtypes)

print("\nCleaned data preview:")
print(data.head(10).to_string(index=False))

# Save cleaned data
data.to_csv(
    OUTPUT_DIR / "cleaned_mobile_usage_data.csv",
    index=False,
    encoding="utf-8-sig"
)

# ==========================================
# 8. DESCRIPTIVE STATISTICS
# ==========================================

summary = data.describe(include="all").transpose()

summary.to_csv(
    OUTPUT_DIR / "descriptive_summary.csv",
    encoding="utf-8-sig"
)

print("\nDescriptive statistics:")
print(summary)

# ==========================================
# 9. FREQUENCY TABLES
# ==========================================

categorical_columns = time_columns + [
    "Number_of_Apps",
    "Academic_Score"
]

with pd.ExcelWriter(
    OUTPUT_DIR / "analysis_tables.xlsx",
    engine="openpyxl"
) as writer:

    data.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

    summary.to_excel(
        writer,
        sheet_name="Summary"
    )

    for column in categorical_columns:

        frequency = (
            data[column]
            .value_counts(dropna=False)
            .rename_axis(column)
            .reset_index(name="Count")
        )

        frequency["Percent"] = (
            frequency["Count"] / len(data) * 100
        ).round(1)

        frequency.to_excel(
            writer,
            sheet_name=column[:31],
            index=False
        )

# ==========================================
# 10. VISUALIZATION
# ==========================================

sns.set_theme(style="whitegrid")

chart_columns = [
    ("Daily_Screen_Time", "Daily Screen Time"),
    ("Study_Time", "Daily Study Time"),
    ("Sleep_Time", "Sleep Duration"),
    ("Academic_Score", "Academic Score")
]

for column, title in chart_columns:

    plt.figure(figsize=(9, 5))

    order = data[column].value_counts().index

    sns.countplot(
        data=data,
        x=column,
        order=order
    )

    plt.title(title)
    plt.xlabel(column.replace("_", " "))
    plt.ylabel("Number of Students")
    plt.xticks(rotation=25, ha="right")
    plt.tight_layout()

    filename = column.lower() + "_distribution.png"

    plt.savefig(
        OUTPUT_DIR / filename,
        dpi=150
    )

    plt.close()

print("\nCharts saved successfully.")

# ==========================================
# 11. FINISH
# ==========================================

print("\n" + "=" * 50)
print("INITIAL ANALYSIS COMPLETED")
print("=" * 50)

print("Output folder:", OUTPUT_DIR.resolve())
print("Cleaned CSV and Excel analysis tables are ready.")
print("Next: correlation analysis and interpretation.")

# ==========================================
# STEP 2: CORRELATION ANALYSIS
# ==========================================

print("\n" + "=" * 50)
print("CORRELATION ANALYSIS")
print("=" * 50)


# Convert the standardized time categories to estimates.
# These values are approximate midpoints, not exact measurements.

time_mapping = {
    "NONE": 0,
    "<1HOUR": 0.5,
    "1-2HOURS": 1.5,
    "2-3HOURS": 2.5,
    "2-4HOURS": 3.0,
    "3-4HOURS": 3.5,
    "4-5HOURS": 4.5,
    "5-6HOURS": 5.5,
    "6-7HOURS": 6.5,
    "7-8HOURS": 7.5,
    "8-9HOURS": 8.5
}

score_mapping = {
    "61-80%": 70.5,
    "80-85%": 82.5,
    "85-90%": 87.5,
    ">90%": 95
}


def normalize_label(value):
    if pd.isna(value):
        return None

    return str(value).strip().upper().replace(" ", "")


# Create a separate numeric dataset for analysis.
analysis_data = pd.DataFrame(index=data.index)

time_columns_for_analysis = {
    "Screen_Time": "Daily_Screen_Time",
    "Social_Media_Time": "Social_Media_Time",
    "Gaming_Time": "Gaming_Time",
    "Study_Time": "Study_Time",
    "Sleep_Time": "Sleep_Time"
}

normalized_time_mapping = {
    normalize_label(key): value
    for key, value in time_mapping.items()
}

normalized_score_mapping = {
    normalize_label(key): value
    for key, value in score_mapping.items()
}

for new_column, original_column in time_columns_for_analysis.items():
    analysis_data[new_column] = data[original_column].map(
        lambda value: normalized_time_mapping.get(
            normalize_label(value)
        )
    )

analysis_data["Academic_Score"] = data["Academic_Score"].map(
    lambda value: normalized_score_mapping.get(
        normalize_label(value)
    )
)

# Show how many values could be converted.
print("\nValid numeric values per variable:")
print(analysis_data.notna().sum())

print("\nMissing numeric values per variable:")
print(analysis_data.isna().sum())

# Calculate Pearson correlation.
correlation_matrix = analysis_data.corr(method="pearson")

print("\nCorrelation Matrix:")
print(correlation_matrix.round(2))

# Save correlation results using the existing output folder.
correlation_matrix.to_csv(
    OUTPUT_DIR / "correlation_matrix.csv",
    encoding="utf-8-sig"
)

# Draw and save the heatmap.
sns.set_theme(style="whitegrid")

plt.figure(figsize=(11, 8))

sns.heatmap(
    correlation_matrix,
    annot=True,
    cmap="coolwarm",
    vmin=-1,
    vmax=1,
    center=0,
    fmt=".2f"
)

plt.title("Correlation Between Mobile Usage and Academic Factors")
plt.tight_layout()

plt.savefig(
    OUTPUT_DIR / "correlation_heatmap.png",
    dpi=300
)

plt.close()

print("\nCorrelation analysis completed.")
print("Saved:", OUTPUT_DIR / "correlation_matrix.csv")
print("Saved:", OUTPUT_DIR / "correlation_heatmap.png")


# ==========================================
# STEP 3: PATTERN DISCOVERY
# ==========================================

print("\n" + "=" * 50)
print("PATTERN DISCOVERY")
print("=" * 50)

# 1. Most common mobile screen-time category
print("\n1. Daily Screen Time Distribution:")
print(
    data["Daily_Screen_Time"]
    .value_counts()
    .to_string()
)

# 2. Most common social media category
print("\n2. Social Media Time Distribution:")
print(
    data["Social_Media_Time"]
    .value_counts()
    .to_string()
)

# 3. Gaming-time distribution
print("\n3. Gaming Time Distribution:")
print(
    data["Gaming_Time"]
    .value_counts(dropna=False)
    .to_string()
)

# 4. Academic-score distribution
print("\n4. Academic Score Distribution:")
print(
    data["Academic_Score"]
    .value_counts()
    .to_string()
)

# 5. Screen-time category compared with academic-score category
print("\n5. Screen Time vs Academic Score:")

screen_score_table = pd.crosstab(
    data["Daily_Screen_Time"],
    data["Academic_Score"],
    margins=True
)

print(screen_score_table)

# Save the cross-tabulation table
screen_score_table.to_csv(
    OUTPUT_DIR / "screen_time_vs_academic_score.csv",
    encoding="utf-8-sig"
)

print("\nPattern discovery tables saved.")


# ==========================================
# STEP 4: DATA VALIDATION
# ==========================================

print("\n" + "=" * 50)
print("DATA VALIDATION")
print("=" * 50)

# Find unusual gaming-time categories
print("\nGaming-time categories:")
print(data["Gaming_Time"].value_counts(dropna=False))

print("\nRecords with Gaming_Time = '1H':")
print(
    data.loc[
        data["Gaming_Time"].astype(str).str.strip().str.upper() == "1H",
        ["Student_ID", "Gaming_Time", "Daily_Screen_Time",
         "Study_Time", "Academic_Score"]
    ].to_string(index=False)
)

print("\nRecords with missing gaming time:")
print(
    data.loc[
        data["Gaming_Time"].isna(),
        ["Student_ID", "Gaming_Time", "Academic_Score"]
    ].to_string(index=False)
)



# ==========================================
# STEP 5: CHECK UNUSUAL TIME CATEGORIES
# ==========================================

print("\n" + "=" * 50)
print("CHECK UNUSUAL TIME CATEGORIES")
print("=" * 50)

time_columns = [
    "Daily_Screen_Time",
    "Social_Media_Time",
    "Gaming_Time",
    "Study_Time",
    "Sleep_Time"
]

for column in time_columns:
    print(f"\n{column}:")
    print(data[column].value_counts(dropna=False).to_string())
