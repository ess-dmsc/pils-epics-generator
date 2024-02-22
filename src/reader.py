import pandas as pd
from typing import List, Tuple


COL_NAMES = ["axis_description", "fbs_description", "pv_name", "mc_unit", "mc_axis_nc", "mc_axis_pn", "has_temp", "temp_units", "extra_dev", "extra_name", "extra_type", "extra_desc"]


class ExcelReader:
    """
    A class for reading data from multi-sheet Excel files.

    Attributes:
        file_path (str): The path to the Excel file to be read.
    """

    def __init__(self, file_path: str):
        """
        Initializes the ExcelReader with the path to an Excel file.

        :param file_path: The path to the Excel file.
        """
        self.file_path = file_path

    def read_sheet_by_index(self, sheet_index: int, columns: List[int]) -> Tuple[pd.DataFrame, str]:
        """
        Reads specified columns from a sheet given by its index.

        :param sheet_index: The index of the sheet to read.
        :param columns: A list of column indices to read.
        :return: A pandas DataFrame containing the specified columns from the sheet.
        """
        # Load the specific sheet
        sheet_name = self._get_sheet_name_by_index(sheet_index)
        if sheet_name is None:
            raise ValueError(f"Sheet index {sheet_index} is out of range.")

        if len(columns) != len(COL_NAMES):
            raise ValueError(f"Number of columns must be {len(COL_NAMES)}")

        # Read instrument name from the sheet
        df_name = pd.read_excel(self.file_path, sheet_name=sheet_name, nrows=1)
        instrument_name = df_name.columns[2]

        # Read specified columns from the sheet
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, usecols=columns)
        df.columns = COL_NAMES

        df = self._fill_mc_unit(df)
        df = self._filter_dataframe(df)
        return df, instrument_name

    def _fill_mc_unit(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Fill in missing mc_unit values in the DataFrame.

        :param df: The DataFrame to fill.
        :return: The filled DataFrame.
        """
        # Fill in missing mc_unit values for all zeroes
        df['mc_unit'] = df['mc_unit'].replace(0, pd.NA)
        df['mc_unit'] = df['mc_unit'].fillna(method='ffill')
        return df

    def _filter_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filters the DataFrame to remove rows with missing values.

        :param df: The DataFrame to filter.
        :return: The filtered DataFrame.
        """

        # remove the third row because it is an example
        df = df.drop(3)
        # remove all rows where mc_axis is not an integer or is missing or there is no extra dev
        df = df[(df['mc_axis_nc'].apply(lambda x: isinstance(x, int) and x > 0)) | (
            df['mc_axis_pn'].apply(lambda x: isinstance(x, int) and x > 0)) | (
            df['extra_dev'].apply(lambda x: isinstance(x, str) and x == 'x'))
        ]
        # remove all rows where the pv_name is missing
        # df = df[df['pv_name'].apply(lambda x: isinstance(x, str))]
        return df


    def _get_sheet_name_by_index(self, sheet_index: int) -> str:
        """
        Retrieves the sheet name given its index.

        :param sheet_index: The index of the sheet.
        :return: The name of the sheet.
        """
        # Load the Excel file to get the sheet names
        xls = pd.ExcelFile(self.file_path)
        sheets = xls.sheet_names
        if 0 <= sheet_index < len(sheets):
            return sheets[sheet_index]
        else:
            return None

