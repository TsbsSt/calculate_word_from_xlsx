import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import numpy as np
from itertools import islice
import os
import json
import re

def main():
    print("____calculate_word_from_xlsx____")

    # input config file
    config = input_config()

    # input workbook
    workbook, df, src = input_workbook(config)

    # build project
    write_to_workbook(workbook, df, config, src)

    # complete build
    print("***complete build***")


def input_config():
    config = {}
    src = "config.jsonc"

    # if is not config file, make config file to default_config

    if not os.path.isfile(src):
        print(f"{src} is not found")
        print(f"create {src}")
        make_config(src)

    # input config file

    with open(src, "r", encoding="utf-8") as f:
        file_content = f.read()
        
        # remove comments if the file is JSONC
        if src.endswith(".jsonc"):
            file_content = re.sub(r"\/\/.*|\/\*[\s\S]*?\*\/", "", file_content)
        
        config = json.loads(file_content)

    return config


def input_workbook(config):
    src = ""
    ext = ""

    # extract supported extensions from config
    supported_ext = config["file"]["extension"]

    # input workbook

    while True:
        print(f"Enter the synopsis file (supported_ext: {supported_ext})")
        src = input(">>> ").strip(r'"')
        ext = os.path.splitext(src)[1][1:]
        
        if ext not in supported_ext:
            print(f"Please enter Files with any extension ({supported_ext})")
            continue
        
        if not os.path.exists(src):
            print(f"The file '{src}' does not exist. Please enter again.")
            continue

        # if file is open, ask to enter it again
        if xlsx_is_open(src):
            print("The file is open. Please close the file before executing.")
            continue

        break

    # load workbook
    workbook = pd.ExcelFile(src)

    target_sheet = get_target_sheet(workbook, config)

    # load default sheet
    df = pd.read_excel(src, sheet_name=target_sheet, header=0, index_col=None)

    return workbook, df, src


def write_to_workbook(workbook, df, config, src):
    # specify column name
    heads = config["sheet"]["headers"].copy()

    # input config
    text_line = int(config["words"]["text_line"])
    detail_line = int(config["words"]["detail_line"])

    # 0 division measures
    df[heads["text"]] = df[heads["text"]].fillna("")
    df[heads["detail"]] = df[heads["detail"]].fillna("")
    df[heads["words"]] = df[heads["words"]].fillna(0).astype(int)
    df[heads["size"]] = df[heads["size"]].fillna(0).astype(int)

    # calculations and inputs
    df[heads["words"]] += np.ceil(df[heads["text"]].str.len() / text_line).astype(int) * text_line * df[heads["size"]]
    df[heads["words"]] += np.ceil(df[heads["detail"]].str.len() / detail_line).astype(int) * detail_line * df[heads["size"]]

    # assign total to the first row, and fill the other cells with None.
    total_sum = df[heads["words"]].sum()
    df[heads["total"]] = pd.Series([total_sum] + [None] * (len(df) - 1))

    # write and save worksheets
    target_sheet = get_target_sheet(workbook, config)
    
    wb_op = openpyxl.load_workbook(src)

    try:
        sheet = wb_op[target_sheet]
    except KeyError:
        sheet = wb_op.worksheets[target_sheet]
    except Exception:
        sheet = wb_op.worksheets[0]

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    wb_op.save(src)


def make_config(file_name):
    default_config = {
        "file": {
            "extension": [
                "xlsx"
            ]
        },
        "sheet": {
            "target": 0,
            "headers": {
                "text": "text",
                "detail": "detail",
                "size": "size",
                "words": "words",
                "total": "total"
            }
        },
        "words": {
            "text_line": 20,
            "detail_line": 10
        }
    }

    with open(file_name, "w") as f:

        json.dump(default_config, f, indent=4)


def xlsx_is_open(src):
    try:
        f = open(src, 'a')
        f.close()

    except:
        return True
    
    else:
        return False


def get_target_sheet(workbook, config):

    # extract target sheet from config
    target_sheet = config["sheet"]["target"]

    if type(target_sheet) is int:
        # if specified by number, load the sheet with the corresponding number
        if target_sheet < len(workbook.sheet_names) - 1:
            target_sheet = int(target_sheet)

    elif type(target_sheet) is str:
        # if specified by name, load the sheet with the corresponding name
        if any(target_sheet in name for name in workbook.sheet_names):
            target_sheet = target_sheet
    else:
        target_sheet = 0

    return target_sheet


if __name__ == "__main__":
    main()

