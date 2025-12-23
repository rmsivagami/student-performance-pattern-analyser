import pandas as pd
import numpy as np

# -------------------------------
# 1. READ INPUT EXCEL
# -------------------------------
input_file = r"data/student_tests_input.xlsx"
df = pd.read_excel(input_file, sheet_name="Student_Test_Data")

TOTAL_QUESTIONS = 10

# -------------------------------
# 2. FUNCTION TO ANALYZE ONE STUDENT
# -------------------------------
def analyze_student_robust(student_row):
    scores = np.array([student_row[f"Test{i}_Score"] for i in range(1, 6)])
    times = np.array([student_row[f"Test{i}_Time"] for i in range(1, 6)])
    x = np.arange(1, 6)  # test numbers

    # Score trend (slope of linear regression)
    score_slope = np.polyfit(x, scores, 1)[0]
    # Time trend (negative slope = improvement)
    time_slope = -np.polyfit(x, times, 1)[0]
    
    # Score consistency
    score_std = np.std(scores, ddof=1)

    # Average score and maximum score
    avg_score = np.mean(scores)
    max_score = np.max(scores)


    # Learning classification (robust logic)
    if score_slope > 0 and time_slope > 0:
        learning_state = "learning in progress"
        category = 1
    elif score_slope > 0 and time_slope <= 0:
        learning_state = "Concepts are clear but needs time improvement"
        category = 2
    elif score_slope <= 0 and time_slope <= 0 and score_std <= 2:
        learning_state = "Needs both score and time improvement"
        category = 3
    elif score_std > 2:
        learning_state = "Answering randomly, may be luck"
        category = 4
    else:
        learning_state = "Concepts are clear but needs time improvement"
        category = 2  

    # Next Action based on category and max score
    if category == 1 and max_score > 8:
        next_action = "Ready to Move"
    elif category == 1 and max_score <= 8:
        next_action = "Needs Practice"
    elif category == 2:
        next_action = "Needs speed"
    elif category == 3:
        next_action = "Needs Revision"
    elif category == 4:
        next_action = "Guessing Behaviour"
    else:
        next_action = "Review individually"

    return pd.Series({
        "Average_Score": avg_score,
        "Score_Std": score_std,
        "Score_Trend_Slope": round(score_slope, 2),
        "Time_Trend_Slope": round(time_slope, 2),
        "Learning_State": learning_state,
        "Category": category,
        "Next_Action": next_action
    })

# -------------------------------
# 3. APPLY TO ALL STUDENTS
# -------------------------------
analysis_df = df.apply(analyze_student_robust, axis=1)

# Combine Student_ID with analysis
final_df = pd.concat([df["Student_ID"], analysis_df], axis=1)

# -------------------------------
# 4. RANK STUDENTS BY UNDERSTANDING INDEX
# -------------------------------
final_df1 = final_df[["Student_ID", "Category", "Learning_State","Next_Action"]]


# -------------------------------
# 5. SAVE TO EXCEL
# -------------------------------
output_file = r"output/student_learning_analysis_output.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Raw_Data", index=False)
    final_df.to_excel(writer, sheet_name="calculation", index=False)
    final_df1.to_excel(writer, sheet_name="Analysis_and_Rank", index=False)


