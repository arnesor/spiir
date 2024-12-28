from pathlib import Path

import pandas as pd
import pytest
from pandas import testing as tm
from spiir.main import fix_split_transactions
from spiir.main import main
from spiir.main import read_transactions_file

from .crypto import decrypt_file
from .crypto import encrypt_file

transactions_file = Path(__file__).parent / "testdata" / "transactions-2024.csv"
facit_file = Path(__file__).parent / "testdata" / "month_facit.parquet"


@pytest.fixture(scope="session", autouse=True)
def decrypt_files() -> None:
    if not transactions_file.is_file():
        decrypt_file(transactions_file.with_suffix(".enc"), transactions_file)
    if not facit_file.is_file():
        decrypt_file(facit_file.with_suffix(".enc"), facit_file)


def encrypt_files() -> None:
    if transactions_file.is_file():
        encrypt_file(transactions_file, transactions_file.with_suffix(".enc"))
    if facit_file.is_file():
        encrypt_file(facit_file, facit_file.with_suffix(".enc"))


@pytest.fixture(scope="module")
def real_transactions() -> pd.DataFrame:
    return read_transactions_file(transactions_file)


def test_fix_split_transactions(real_transactions: pd.DataFrame) -> None:
    result = fix_split_transactions(real_transactions)
    assert len(result) < len(real_transactions)


def test_main() -> None:
    # encrypt_files()
    facit_df = pd.read_parquet(facit_file)
    facit_df.index = facit_df.index.astype(
        "string[python]"
    )  # The index dtype is lost when writing to parquet
    result_df = main(transactions_file)
    tm.assert_frame_equal(result_df, facit_df)
