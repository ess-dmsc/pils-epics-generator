import pandas as pd
from typing import List, Tuple


# COL_NAMES = ["axis_description", "fbs_description", "pv_name", "mc_unit", "ptp", "mc_axis_nc", "mc_axis_pn", "pils_name", "pils_unit", "has_temp", "temp_units", "extra_dev", "extra_name", "extra_type", "extra_desc"]
# COLUMNS_INDEX = [1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 17, 18]


COLUMN_INFO = [
    (1, "axis_description"),
    (2, "pv_name"),
    (3, "fbs_description"),
    (4, "mc_unit"),
    (6, "pv_root"),
    (7, "ptp"),
    (8, "axis_index"),
    (9, "actuator_type"),
    (10, "pils_name"),
    (11, "pils_unit"),
    (12, "has_temp"),
    (13, "temp_units"),
    (14, "extra_dev"),
    (15, "extra_name"),
    (16, "extra_type"),
    (17, "extra_desc")
]

COLUMNS_INDEX = [info[0] for info in COLUMN_INFO]
COL_NAMES = [info[1] for info in COLUMN_INFO]


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

        df = self._prep_ptp(df)
        df = self._fill_mc_unit(df)
        df = self._fill_ptp(df)
        df = self._fill_pv_root(df)
        df = self._filter_dataframe(df)
        df = self._build_nc_pn(df)
        # df = self._filter_non_axis_rows(df)
        df.index = range(len(df.index))
        print(df.to_string())
        return df, instrument_name

    def _fill_mc_unit(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Fill in missing mc_unit values in the DataFrame.

        :param df: The DataFrame to fill.
        :return: The filled DataFrame.
        """
        # Fill in missing mc_unit values for all zeroes
        df['mc_unit'] = df['mc_unit'].replace(0, pd.NA)
        df['mc_unit'] = df['mc_unit'].ffill()
        return df

    def _prep_ptp(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Prepare the ptp column in the DataFrame.

        If a row has an `mc_unit` value (not NaN and not 0), and `ptp` is missing, it should be set to 'no'.
        If `ptp` is explicitly 'yes', it remains unchanged.

        :param df: The DataFrame to prepare.
        :return: The updated DataFrame.
        """
        valid_mc_unit = df['mc_unit'].apply(lambda x: pd.notna(x) and x != 0)
        df.loc[valid_mc_unit & df['ptp'].isna(), 'ptp'] = 'no'

        return df

    def _fill_ptp(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Fill in missing ptp values in the DataFrame.

        :param df: The DataFrame to fill.
        :return: The filled DataFrame.
        """
        # Fill in missing ptp values for all zeroes
        df['ptp'] = df['ptp'].replace(0, pd.NA)
        df['ptp'] = df['ptp'].ffill()
        return df

    def _fill_pv_root(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Fill in missing pv_root values in the DataFrame.

        :param df: The DataFrame to fill.
        :return: The filled DataFrame.
        """
        # Fill in missing pv_root values for all zeroes
        df['pv_root'] = df['pv_root'].replace(0, pd.NA)
        df['pv_root'] = df['pv_root'].ffill()
        return df

    def _build_nc_pn(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Build the nc and pn columns in the DataFrame.
        If the actuator_type is "Electrical", then the nc column should be set
        to the axis_index and the pn column should be set to None.
        If the actuator_type is "Pneumatic", then the pn column should be set
        to the axis_index and the nc column should be set to None.

        :param df: The DataFrame to fill.
        :return: The filled DataFrame.
        """
        df['mc_axis_nc'] = df['axis_index']
        df['mc_axis_pn'] = df['axis_index']
        df.loc[df['actuator_type'] == 'Electrical', 'mc_axis_pn'] = pd.NA
        df.loc[df['actuator_type'] == 'Pneumatic', 'mc_axis_nc'] = pd.NA
        df.loc[df['actuator_type'] == 0, 'mc_axis_nc'] = pd.NA
        df.loc[df['actuator_type'] == 0, 'mc_axis_pn'] = pd.NA
        return df

    def _filter_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filters the DataFrame to remove rows with missing values.

        :param df: The DataFrame to filter.
        :return: The filtered DataFrame.
        """
        df = df.drop([0, 1, 2, 3, 4])
        df['axis_index'] = df['axis_index'].replace(0, pd.NA)
        df = df.dropna(subset=['axis_index'])
        return df

    def _filter_non_axis_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filters the DataFrame to remove rows that are not associated with a valid NC or PN axis.
        Keeps only rows where either 'mc_axis_nc' or 'mc_axis_pn' is present (not NaN).

        :param df: The DataFrame to filter.
        :return: The filtered DataFrame.
        """
        df = df.dropna(subset=['mc_axis_nc', 'mc_axis_pn'], how='all')
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

