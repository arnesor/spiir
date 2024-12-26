import numpy as np
import openpyxl
import pandas as pd
from decimal import Decimal
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

year = 2024
in_filename = "alle-poster-2024-12-26.csv"
out_filename = f"spiir-accounting-{year}.xlsx"


def format_spiir_sheet(filename: str) -> None:
    wb = openpyxl.load_workbook(filename)
    ws = wb["Sheet1"]
    max_row = ws.max_row

    my_format = "# ##0;-# ##0;0;@"
    for row in ws.iter_rows(min_row=2, max_row=70, min_col=2, max_col=14):
        for cell in row:
            cell.number_format = my_format

    dim_holder = DimensionHolder(worksheet=ws)
    dim_holder["A"] = ColumnDimension(ws, min=1, max=1, width=30)
    for col in range(ws.min_column + 1, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(
            ws, min=col, max=col, width=9
        )
    ws.column_dimensions = dim_holder

    for col in range(ws.min_column + 1, ws.max_column):
        col_letter = get_column_letter(col)
        sum_string = f"=SUM({col_letter}2:{col_letter}{max_row})"
        ws.cell(row=max_row + 1, column=col).value = sum_string

    wb.save(f"formatted-{year}.xlsx")


def main():
    df = pd.read_csv(
        in_filename,
        index_col=0,
        sep=";",
        decimal=",",
        parse_dates=["Date", "CustomDate"],
        dayfirst=True,
        true_values=["Yes"],
        false_values=["No"],
        usecols=[
            "Id",
            "Date",
            "Description",
            "MainCategoryName",
            "CategoryName",
            "CategoryType",
            "ExpenseType",
            "Amount",
            "Extraordinary",
            "SplitGroupId",
            "CustomDate",
        ],
        dtype={
            "Id": "string",
            "Description": "string",
            "MainCategoryName": "string",
            "CategoryName": "string",
            "CategoryType": "string",
            "ExpenseType": "string",
            "SplitGroupId": "string",
        },
    )
#    df["Amount"] = df["Amount"].apply(Decimal)
    df["Amount"] = df["Amount"].astype(float)

    # There is a bug in Spiir which sometimes makes the original transaction in a
    # split transaction visible. It should be hidden. Solve this be removing the
    # original transaction (the first one) in each split group. Then add the non-split
    # transactions.
    split_group_df = df.groupby("SplitGroupId", as_index=False, group_keys=False).apply(
        lambda group: group.iloc[1:],
        include_groups = False
    )
    no_split_group_df = df[df.SplitGroupId.isnull()]
    df2 = pd.concat([split_group_df, no_split_group_df])

    df2 = df2[(df2["CategoryType"] != "Exclude") & (df2["Extraordinary"] == False)]
    df2["CorrectedDate"] = df2["CustomDate"].where(
        df2["CustomDate"].notnull(), df2["Date"]
    )
    df_year = df2[df2["CorrectedDate"].dt.year == year]
    print(f"Shape after fixing: {df_year.shape}")

    category_table = pd.pivot_table(
        df_year,
        values="Amount",
        index="CategoryName",
        columns=pd.Grouper(key="CorrectedDate", freq="ME"),
        aggfunc="sum",
        fill_value=0,
    )
    category_table.columns = pd.to_datetime(category_table.columns).strftime("%b %Y")

    category_table.to_excel(out_filename)
    print("Finished writing spreadsheet.")


if __name__ == "__main__":
    main()
    format_spiir_sheet(out_filename)
