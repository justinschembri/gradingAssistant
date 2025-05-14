"""Export a structured grading excel to HTML"""

#standard
from pathlib import Path
import csv
from typing import Tuple
import math
#external
import pandas as pd
#internal

def load_grading_sheet(sheet: Path, group: int) -> dict:
    output = {"Question":[],
              "Marks Available":[],
              "Mark":[],
              "Feedback":[]
              }
    question_columns: list[int] = []
    marks_available_columns: list[int] = []
    mark_columns: list[int] = []
    feedback_columns: list[int] = []
    with open(sheet, "r") as f:
        reader = csv.reader(f)
        for idx, row in enumerate(reader):
            if idx == 0:
                for id, header in enumerate(row):
                    match header:
                        case "question":
                            question_columns.append(id)
                        case "marks available":
                            marks_available_columns.append(id)
                        case "mark":
                            mark_columns.append(id)
                        case "feedback":
                            feedback_columns.append(id)
            if idx == group:
                for id, data in enumerate(row):
                    if id in question_columns:
                        output["Question"].append(data)
                    elif id in marks_available_columns:
                        output["Marks Available"].append(int(data))
                    elif id in mark_columns:
                        output["Mark"].append(int(data))
                    elif id in feedback_columns:
                        output["Feedback"].append(data)
    return output

def convert_to_dataframe(data:dict) -> Tuple[pd.DataFrame, int]:
    # return on this is not so nice.
    df = pd.DataFrame(data)
    total = df["Mark"].sum()
    totals_row = {"Question": "",
             "Marks Available": "Your Total:",
             "Mark":total,
             "Feedback":""
             }
    df = pd.concat([df, pd.DataFrame([totals_row])], ignore_index=True)
    return (df, total)

def combine_dataframes(sheet_1: Path, sheet_2: Path, group: int, labs:list[int]) -> pd.DataFrame:
    df_1 = convert_to_dataframe(load_grading_sheet(sheet_1, group))
    df_2 = convert_to_dataframe(load_grading_sheet(sheet_2, group))
    grade_total = math.ceil(((df_1[1] + df_2[1]) / 2)) 
    df_1_header = {"Question":f"Lab{labs[0]}",
                   "Marks Available":"",
                   "Mark":"",
                   "Feedback":""}
    df_2_header = {"Question":f"Lab{labs[1]}",
                   "Marks Available":"",
                   "Mark":"",
                   "Feedback":""}
    grand_total = {"Question": "",
             "Marks Available": "Grand Total",
             "Mark":grade_total,
             "Feedback":""
             }
    combined_df = pd.concat(
            [
                pd.DataFrame([df_1_header]),
                df_1[0],
                pd.DataFrame([df_2_header]),
                df_2[0],
                pd.DataFrame([grand_total]),
                ],
            ignore_index = True
    )
    return combined_df


def to_html(combined_df:pd.DataFrame) -> None:
    html_table = combined_df.to_html(index=False, escape=True, classes="my-table")
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Styled Table</title>
        <style>
            .my-table {{
                font-family: Arial, sans-serif;
                border-collapse: collapse;
                width: 100%;
            }}
            .my-table th, .my-table td {{
                border: 1px solid #ddd;
                padding: 8px;
            }}
            .my-table tr:nth-child(even) {{
                background-color: #f2f2f2;
            }}
            .my-table th {{
                background-color: #f4f4f4;
            }}
        </style>
    </head>
    <body>
        {html_table}
    </body>
    </html>
    """

    with open("output.html", "w") as f:
        f.write(html_content)

if __name__ == "__main__":
    sheet_1 = Path("/Users/jschembri/Library/CloudStorage/GoogleDrive-justin@compound-engineering.com/My Drive/1-phd/2-archives+admin/teaching/l-geoweb2025/grading/assignment1/grading-lab1.csv")
    sheet_2 = Path("/Users/jschembri/Library/CloudStorage/GoogleDrive-justin@compound-engineering.com/My Drive/1-phd/2-archives+admin/teaching/l-geoweb2025/grading/assignment1/grading-lab2.csv")
    data = load_grading_sheet(sheet_1, group=1)
    print(data)
    df = convert_to_dataframe(data)
    print(df)
    combined_df = combine_dataframes(sheet_1, sheet_2, 1, [1,2])
    print(combined_df)
    to_html(combined_df)
    print("Saved to HTML!")
