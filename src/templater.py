from config import ROOT_DIR
import json
import pandas as pd
from pathlib import Path
from typing import Tuple

def _set_constants() -> Tuple[int, dict]:
    """
    Reads the metadata file and returns the total number of groups and the
    lab-assignment relationships. 

    Returns:
        tuple: A tuple containing the total number of groups and the
        lab-assignment relationships. 
    """
    with open(METADATA_PATH) as f:
        metadata = json.load(f)
        total_groups = metadata["groups"]
        lab_assignment_relationships = metadata["assignments"]
        return total_groups, lab_assignment_relationships

METADATA_PATH = ROOT_DIR / "src" / "metadata.json"
TOTAL_GROUPS, LAB_ASSIGNMENT_RELATIONSHIPS = _set_constants()

def _construct_df(lab_info:dict) -> "pd.DataFrame":
    """Create and export an xlsx grading template for labs.

    Template includes 4 initial columns: questions, marks available,
    question summary and grading schema. The remaining columns are reserved
    for feedback and the mark awarded on a group by group basis.

    Args:
        lab_a: _description_
        lab_b: _description_

    Returns:
        _description_
    """
    questions = pd.DataFrame(
        {"question":[str(i['number']) for i in lab_info['questions']]}
        )
    marks_available = pd.DataFrame(
        {"marks available":[i['marksAvailable'] for i in lab_info['questions']]}
        )
    summary = pd.DataFrame(
        {"question summary":[i['summary'] for i in lab_info['questions']]}
        )
    schema = pd.DataFrame(
        {"schema":[i['schema'] for i in lab_info['questions']]}
        )
    output = pd.concat(
        [questions, marks_available, summary, schema], axis=1
        )
    # Add empty columns per group:
    num_questions = range(len(output))
    for i in range(TOTAL_GROUPS):
        df = pd.DataFrame(
            {"Group " + str(i+1) + " Mark":["" for n in num_questions]}
            )
        output = pd.concat([output, df], axis=1)
        df = pd.DataFrame(
            {"Group " + str(i+1) + " Feedback":["" for n in num_questions]}
            )
        output = pd.concat([output, df], axis=1)
    save_path = "src/outputs/" + "lab" + str(lab_info["lab"]) + ".xlsx"
    output.to_excel(save_path, index=False)

if __name__ == "__main__":
    x = input("Which assignment templates do you want to generate?")
    lab_a, lab_b = LAB_ASSIGNMENT_RELATIONSHIPS[x]

    with open("src/inputs/"+ lab_a + ".json") as file_1, \
      open("src/inputs/"+ lab_b + ".json") as file_2:
        lab_a = json.load(file_1)
        lab_b = json.load(file_2)
        _construct_df(lab_a)
        _construct_df(lab_b)

