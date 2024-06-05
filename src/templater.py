from config import ROOT_DIR, OUTPUT_PATH
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension
import json
import pandas as pd
from pathlib import Path
from typing import Tuple, List

def _set_constants(metadata_path:Path) -> Tuple[dict[str, List], int, dict[str, List]]:
    """
    Reads the metadata file and returns the total number of groups and the
    lab-assignment relationships. 

    Returns:
        tuple: A tuple containing the total number of groups and the
        lab-assignment relationships. 
    """
    with open(METADATA_PATH) as f:
        metadata = json.load(f)
        lab_assignment_relationships = metadata["assignments"]
        total_groups = metadata["groups"]
        num_questions = metadata["numQuestions"]
        return lab_assignment_relationships, total_groups, num_questions

METADATA_PATH = ROOT_DIR / "src" / "metadata.json"
LAB_ASSIGNMENT_REL, TTL_GRPS, QUESTION_ARRAYS = _set_constants(METADATA_PATH)

def construct_schema_sheet(lab:str) -> None:
    """Construct a schema template for a given lab.

    Constructed from the constants (TTL_GRPS, LAB_ASSIGNMENT_REL, NUM_QUESTIONS)
    defined in this module.

    Columns : 
    - 'question' (numbering constructed by this function)
    - 'marks available' (left empty, fill in the excel sheet)
    - 'question summary' (left empty, fill in the excel sheet)
    - 'schema' (left empty, fill in the excel sheet)

    Args:
        lab: Lab reference number.

    Returns:
        Template schema dataframe for a given lab.
    """
    num_questions = QUESTION_ARRAYS[lab]
    questions = pd.DataFrame(
        {"question": [i for i in num_questions]}
        )
    marks_available = pd.DataFrame(
        {"marks available":["" for n in num_questions]}
        )
    summary = pd.DataFrame(
        {"question summary":["" for n in num_questions]}
        )
    schema = pd.DataFrame(
        {"schema":["" for n in num_questions]}
        )
    output = pd.concat(
        [questions, marks_available, summary, schema], axis=1
        )
    save_path = OUTPUT_PATH / (lab + ".xlsx")
    output.to_excel(save_path, sheet_name="schema", index=False)

def construct_grade_template(lab:str, repeat_column:int, spacing:int) -> None:
    lab_questions = QUESTION_ARRAYS[lab]
    load_path = OUTPUT_PATH / (lab + ".xlsx")
    workbook = load_workbook(load_path)
    schema_sheet = workbook['schema']
    try:
        grading_sheet = workbook['grading']
    except:
        grading_sheet = workbook.create_sheet("grading")
    # column A: groups (group 1 → group n)
    def _repeated_columns() -> None:
        column_spacing = range(
            repeat_column, 
            len(lab_questions)*spacing, 
            spacing)
        name_map = {
            1:"question",
            2: "marks available",
            3: "question summary",
            4: "schema"}
        repeat_col_letter = get_column_letter(repeat_column)
        # first column, just groups:
        #header:
        grading_sheet['A1'] = 'groups' # type: ignore
        for i in range(1, TTL_GRPS+1):
            #contents
            grading_sheet["A" + str((i+1))] = "group " + str((i)) # type: ignore
        # repeated columns:
        for i in column_spacing:
            input_cols = [get_column_letter(i+1) for i in column_spacing]
            for idx, j in enumerate(input_cols):
                for g in range(TTL_GRPS):
                    grading_sheet[j + "1"] = name_map[repeat_column] # type: ignore
                    grading_sheet[j + str(g+2)] = '=schema!$'+ repeat_col_letter +'$' + str(idx+2) # type: ignore
        # mark and feedback columns:
        for i in range(spacing-1, len(lab_questions)*spacing, spacing):
            mark_column = get_column_letter(i+1)
            feedback_column = get_column_letter(i+2)
            grading_sheet[mark_column + "1"] = "mark" # type: ignore
            grading_sheet[feedback_column + "1"] = "feedback" # type: ignore
    
    def _styling() -> None:
        # HEADER TEXT:
        for c in grading_sheet['A']: # type: ignore
            c.font = Font(size=15, bold=True)
            c.alignment = Alignment(wrap_text=True, shrink_to_fit=False)
        for c in grading_sheet['1']: # type: ignore
            c.font = Font(size=15, bold=True)
            c.alignment = Alignment(wrap_text=True, shrink_to_fit=False)
        # COLUMN SPECIFIC / GRADING SHEET
        grading_sheet.column_dimensions["A"].width = 20 # type: ignore
        for column in grading_sheet.iter_cols(min_row=1, max_row=1): # type: ignore
            for cell in column:
                if cell.value == "question":
                    grading_sheet.column_dimensions[cell.column_letter].width = 13 # type: ignore
                elif cell.value == "marks available":
                    grading_sheet.column_dimensions[cell.column_letter].width = 13 # type: ignore
                elif cell.value == "question summary":
                    grading_sheet.column_dimensions[cell.column_letter].width = 40 # type: ignore
                    for c in grading_sheet[cell.column_letter]: # type: ignore
                        c.alignment = Alignment(wrap_text=True)
                elif cell.value == "schema":
                    grading_sheet.column_dimensions[cell.column_letter].width = 40 # type: ignore
                    for c in grading_sheet[cell.column_letter]: # type: ignore
                        c.alignment = Alignment(wrap_text=True)
                elif cell.value == "feedback":
                    grading_sheet.column_dimensions[cell.column_letter].width = 40 # type: ignore
                    for c in grading_sheet[cell.column_letter]: # type: ignore
                        c.alignment = Alignment(wrap_text=True)
        # SCHEMA SHEET
        schema_sheet.column_dimensions["A"].width = 20 # type: ignore
        schema_sheet.column_dimensions["B"].width = 40 # type: ignore
        schema_sheet.column_dimensions["C"].width = 40 # type: ignore
        schema_sheet.column_dimensions["D"].width = 40 # type: ignore
        for column in schema_sheet.iter_rows(): # type: ignore
            for c in column:
                c.font = Font(size= 15)
                c.alignment = Alignment(wrap_text=True)
        # ALTERNATING COLOURS
        fill_light = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        fill_dark = PatternFill(start_color="A9A9A9", end_color="A9A9A9", fill_type="solid")
        for i, row in enumerate(grading_sheet.iter_rows(min_row=2), start=2): # type: ignore
            fill = fill_light if i % 2 == 0 else fill_dark
            for cell in row:
                cell.fill = fill
        #BORDERS
        thin_border_side = Side(border_style="thick", color="000000")
        border_right = Border(right=thin_border_side)
        for column in grading_sheet.iter_cols(min_row=1, max_row=1): # type: ignore
            for cell in column:
                if cell.value == 'feedback' or cell.value == 'groups':
                    for row in grading_sheet.iter_rows(min_row=2, min_col=cell.column, max_col=cell.column): # type: ignore
                        for cell in row:
                            cell.border = border_right
        
    _repeated_columns()
    _styling()
    workbook.save(load_path)

if __name__ == "__main__":
    x = input("Which lab number do you want to generate a template for? ")
    lab_name = "lab" + x
    construct_schema_sheet(lab_name)
    construct_grade_template(lab=lab_name, repeat_column=1, spacing=6)
    construct_grade_template(lab=lab_name, repeat_column=2, spacing=6)
    construct_grade_template(lab=lab_name, repeat_column=3, spacing=6)
    construct_grade_template(lab=lab_name, repeat_column=4, spacing=6)
