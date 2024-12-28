from decimal import Decimal

import openpyxl
import pandas as pd
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


def fix_split_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """Fix errors in split transactions.

    There is a bug in Spiir which sometimes makes the original transaction in a
    split transaction visible. It should be hidden. Solve this be removing the
    original transaction (the first one) in each split group. Then add the non-split
    transactions.

    Args:
        df: DataFrame with raw Spiir transaction data. Expected to have a column called
            "SplitGroupId" which indicates split groups. Transactions with null values
            in "SplitGroupId" are considered as non-split transactions.

    Returns:
        A DatFrame with the fixed list of transactions.
    """
    split_group_df = df.groupby("SplitGroupId", as_index=False, group_keys=False).apply(
        lambda group: group.iloc[1:], include_groups=False
    )
    no_split_group_df = df[df.SplitGroupId.isnull()]
    return pd.concat([split_group_df, no_split_group_df])


def remove_excluded_and_extraordinary(df: pd.DataFrame) -> pd.DataFrame:
    """Filter out excluded and extraordinary rows in the dataframe.

    Args:
        df: The dataframe that contains a "CategoryType" column to specify
            category types and an "Extraordinary" column to indicate extraordinary entries.

    Returns:
        pd.DataFrame: A filtered dataframe excluding the "Exclude" and "Extraordinary" rows.
    """
    return df[(df["CategoryType"] != "Exclude") & (df["Extraordinary"] == False)]


def correct_dates_by_year(df: pd.DataFrame, year: int) -> pd.DataFrame:
    df = df.copy()  # Ensure the input DataFrame is not a slice
    df["CorrectedDate"] = df["CustomDate"].combine_first(df["Date"])
    return df[df["CorrectedDate"].dt.year == year]


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

    df_corrected = fix_split_transactions(df)
    df_base = remove_excluded_and_extraordinary(df_corrected)
    df_year = correct_dates_by_year(df_base, year)
    print(f"Shape after fixing: {df_year.shape}")

    df_year["Amount"] = df["Amount"].apply(Decimal)
    category_table = pd.pivot_table(
        df_year,
        values="Amount",
        index="CategoryName",
        columns=pd.Grouper(key="CorrectedDate", freq="ME"),
        aggfunc="sum",
        fill_value=0,
    )
    category_table.columns = pd.to_datetime(category_table.columns).strftime("%b %Y")

    # Convert Decimal columns to float
    for col in category_table.columns:
        category_table[col] = category_table[col].astype(float)

    print("Detailed Description:")
    print("Number of rows:", len(category_table))
    print("Number of columns:", len(category_table.columns))
    print("Column names:", category_table.columns.tolist())
    print("Index type:", type(category_table.index))
    print(
        "Index values:",
        (
            category_table.index.tolist()
            if len(category_table.index) < 200
            else "Too many to display"
        ),
    )
    print("Index name:", df.index.name)

    category_table.to_excel(out_filename)
    print("Finished writing spreadsheet.")


if __name__ == "__main__":
    main()
#    format_spiir_sheet(out_filename)
