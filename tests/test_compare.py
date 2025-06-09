import pandas as pd
from TableComparer.muhasebe import compare_tables


def test_compare_tables_basic():
    df1 = pd.DataFrame(
        {
            "TARİH": ["01/01/2020", "02/01/2020"],
            "BORÇ": ["10.00", "20.00"],
            "ALACAK": ["0.00", "0.00"],
            "OTHER": [1, 2],
        }
    )
    df2 = pd.DataFrame(
        {
            "TARİH": ["02/01/2020", "03/01/2020"],
            "BORÇ": ["20.00", "30.00"],
            "ALACAK": ["0.00", "0.00"],
            "OTHER": [2, 3],
        }
    )

    left_only, right_only = compare_tables(df1, df2)

    expected_left = df1.iloc[[0]].astype({"OTHER": float}).reset_index(drop=True)
    expected_right = df2.iloc[[1]].astype({"OTHER": float}).reset_index(drop=True)

    pd.testing.assert_frame_equal(left_only.reset_index(drop=True), expected_left)
    pd.testing.assert_frame_equal(right_only.reset_index(drop=True), expected_right)


def test_compare_tables_no_difference():
    df1 = pd.DataFrame(
        {
            "TARİH": ["01/01/2020"],
            "BORÇ": ["10.00"],
            "ALACAK": ["0.00"],
            "OTHER": [1],
        }
    )
    df2 = df1.copy()

    left_only, right_only = compare_tables(df1, df2)

    assert left_only.empty
    assert right_only.empty
