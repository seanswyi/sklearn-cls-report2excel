from argparse import ArgumentParser, Namespace
import os
from typing import Union

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import (
    Border,
    Color,
    Font,
    PatternFill,
    Side
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from tqdm import tqdm


# Formatting for header row and column.
#   Header color is light gray here.
HEADER_COLOR = PatternFill(
    start_color="D9D9D9",
    end_color="D9D9D9",
    fill_type="solid")
HEADER_FONT = Font(
    name="Arial",
    bold=True
)
DEFAULT_FONT = Font(
    name="Arial",
    bold=False
)


# For conditional formatting.
#   Higher = greener.
RED = Color(rgb="E67C73")
WHITE = Color(rgb="FFFFFF")
GREEN = Color(rgb="57BB8A")
COLOR_SCALE_RULE = ColorScaleRule(
    start_type="num",
    start_value=0.0,
    start_color=RED,
    mid_type="num",
    mid_value=0.6,
    mid_color=WHITE,
    end_type="num",
    end_value=1.0,
    end_color=GREEN,
)


# Floating point precision.
FLOAT_PREC = "0.0000"


def get_grid_coordinates(start_row: int, start_col: int, num_rows: int, num_cols: int):
    for i in range(num_rows):
        for j in range(num_cols):
            yield start_row + i, start_col + j


def convert_report2excel(
    workbook: Workbook,
    report: Union[pd.DataFrame, dict[str, float]],
    sheet_name: str=""
) -> Workbook:
    """Function to convert classification report to formatted Excel file.

    An openpyxl.Workbook object must first be created outside of the func and provided.
    The func will create a formatted sheet, add it to the provided Workbook, and return it.
    """
    if isinstance(report, dict):
        report = pd.DataFrame(report).T
        report.reset_index(inplace=True)
        report.rename(columns={"index": "class"})

    df_xl = dataframe_to_rows(
                df=report,
                index=False
            )

    worksheet = workbook.create_sheet(title=sheet_name)

    for row in df_xl:
        worksheet.append(row)

    # Outer boundary formatting.
    outer_border_top_row = 1
    outer_border_top_col = 1
    outer_border_bottom_row = report.shape[0] + 1
    outer_border_bottom_col = report.shape[1]

    top_left_corner = Border(
        top=Side(style="thin"),
        left=Side(style="thin")
    )
    top_right_corner = Border(
        top=Side(style="thin"),
        right=Side(style="thin")
    )
    bottom_left_corner = Border(
        left=Side(style="thin"),
        bottom=Side(style="thin")
    )
    bottom_right_corner = Border(
        right=Side(style="thin"),
        bottom=Side(style="thin")
    )
    top = Border(top=Side(style="thin"))
    bottom = Border(bottom=Side(style="thin"))
    left = Border(left=Side(style="thin"))
    right = Border(right=Side(style="thin"))

    top_left = (
        f"{get_column_letter(outer_border_top_col)}"
        f"{outer_border_top_row}"
    )
    bottom_right = (
        f"{get_column_letter(outer_border_bottom_col)}"
        f"{outer_border_bottom_row}"
    )
    for row in worksheet[top_left:bottom_right]:
        for cell in row:
            if (
                cell.row == 1 and
                cell.column == 1
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = top_left_corner
            elif (
                cell.row == 1 and
                cell.column == report.shape[0]
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = top_right_corner
            elif (
                cell.row == report.shape[0] + 1 and
                cell.column == 1
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = bottom_left_corner
            elif (
                cell.row == report.shape[0] + 1 and
                cell.column == report.shape[1]
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = bottom_right_corner
            elif (
                cell.row in range(1, report.shape[0] + 1) and
                cell.column == 1
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = left
            elif (
                cell.row in range(1, report.shape[0] + 1) and
                cell.column == report.shape[1]
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = right
            elif (
                cell.row == 1 and
                cell.column in range(1, report.shape[1] + 1)
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = top
            elif (
                cell.row == report.shape[0] + 1 and
                cell.column in range(1, report.shape[1] + 1)
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).border = bottom
            else:
                continue

    # Header boundary.
    row = 1
    start_col = 1
    end_col = report.shape[1]

    left = f"{get_column_letter(start_col)}{start_col}"
    right = f"{get_column_letter(end_col)}{end_col}"

    for cell in worksheet[left:right][0]:
        existing_border = cell.border
        new_border = Border(
            left=existing_border.left,
            right=existing_border.right,
            top=existing_border.top,
            bottom=Side(style="thin")
        )
        worksheet.cell(row=cell.row, column=cell.column).border = new_border

    # Support boundary.
    support_col = report.columns.get_loc("support") + 1

    start_row = 1
    end_row = report.shape[0] + 1

    top = f"{get_column_letter(support_col)}{start_row}"
    bottom = f"{get_column_letter(support_col)}{end_row}"

    for row in worksheet[top:bottom]:
        for cell in row:
            existing_border = cell.border
            new_border = Border(
                left=Side(style="thin"),
                right=existing_border.right,
                top=existing_border.top,
                bottom=existing_border.bottom
            )
            worksheet.cell(row=cell.row, column=cell.column).border = new_border

    # Label boundary.
    end_row = report.shape[0] + 1

    for row in worksheet["A1":f"A{end_row}"]:
        for cell in row:
            existing_border = cell.border
            new_border = Border(
                left=existing_border.left,
                right=Side(style="thin"),
                top=existing_border.top,
                bottom=existing_border.bottom
            )
            worksheet.cell(row=cell.row, column=cell.column).border = new_border

    # Thick avg divider.
    avg_row = report.shape[0] + 2

    try:
        if "accuracy" in report.iloc[:, 0].values:
            avg_row = report[report.iloc[:, 0] == "accuracy"].index[0] + 2
        else:
            avg_row = report[report.iloc[:, 0] == "micro avg"].index[0] + 2

        start = f"A{avg_row}"
        end = f"{get_column_letter(report.shape[1])}{avg_row}"

        for row in worksheet[start:end]:
            for cell in row:
                existing_border = cell.border
                new_border = Border(
                    left=existing_border.left,
                    right=existing_border.right,
                    top=Side(style="thick"),
                    bottom=existing_border.bottom
                )
                worksheet.cell(row=cell.row, column=cell.column).border = new_border
    except IndexError:
        pass

    # Conditional formatting.
    gradient_grid_start_row = 1
    gradient_grid_start_col = report.columns.get_loc("precision") + 1
    gradient_grid_end_row = avg_row + 2
    gradient_grid_end_col = support_col - 1

    grid_range = (
        f"{get_column_letter(gradient_grid_start_col)}{gradient_grid_start_row}:"
        f"{get_column_letter(gradient_grid_end_col)}{gradient_grid_end_row}"
    )
    worksheet.conditional_formatting.add(range_string=grid_range, cfRule=COLOR_SCALE_RULE)

    # Format fonts.
    for row in worksheet[top_left:bottom_right]:
        for cell in row:
            if (
                cell.row == 1 or
                cell.column == 1
            ):
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).font = HEADER_FONT
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).fill = HEADER_COLOR
            else:
                worksheet.cell(
                    row=cell.row,
                    column=cell.column
                ).font = DEFAULT_FONT

    # Format floating point precision.
    top_left = "B2"
    bottom_right = (
        f"{get_column_letter(support_col - 1)}"
        f"{report.shape[0] + 1}"
    )

    for row in worksheet[top_left:bottom_right]:
        for cell in row:
            worksheet.cell(
                row=cell.row,
                column=cell.column
            ).number_format = FLOAT_PREC

    return workbook


def main(args: Namespace) -> None:
    report_files: str = os.listdir(args.report_dir)
    report_filepaths: list[str] = [os.path.join(args.report_dir, f) for f in report_files]

    workbook = Workbook()
    workbook.remove(workbook.active)  # Remove default sheet.

    pbar = tqdm(
        iterable=report_filepaths,
        desc="Converting classification reports to formatted Excel files",
        total=len(report_filepaths)
    )
    for file in pbar:
        report = pd.read_csv(file)
        sheet_name = os.path.splitext(file.split('/')[-1])[0]
        workbook = convert_report2excel(workbook, report, sheet_name)

    workbook.save(args.report_dir)


if __name__ == "__main__":
    parser = ArgumentParser()

    parser.add_argument(
        "--report_dir",
        default="",
        type=str,
        help="Directory where all of the individual reports are.",
    )
    parser.add_argument(
        "--report_filename",
        default="",
        type=str,
        help="Path to single report file."
    )

    args = parser.parse_args()

    main(args)
