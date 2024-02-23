import os

import pytest
from src.reader import ExcelReader


def data_file_path(filename):
    return os.path.join(os.path.dirname(__file__), filename)


@pytest.fixture
def reader():
    file_path = data_file_path("test.xlsx")
    return ExcelReader(file_path)


def test_can_read_excel_file(reader):
    file_path = data_file_path("test.xlsx")

    result = reader.file_path

    assert result == file_path


def test_can_read_sheet_by_index(reader):
    # Given
    sheet_index = 0
    columns = [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15]

    # When
    result, instrument_name = reader.read_sheet_by_index(sheet_index, columns)
    row_0 = result.iloc[0]

    # Then
    assert instrument_name == 'YMIR'
    assert row_0["axis_description"] == "Heavy Shutter"
    assert row_0["pv_name"] == "HvSht:MC-Pne-01"
    assert row_0["fbs_description"] == "Heavy Shutter"
    assert row_0["mc_unit"] == 1
    assert row_0["mc_axis_pn"] == 1


def test_dataframe_filtered(reader):
    # Given
    sheet_index = 0
    columns = [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15]

    # When
    result, _ = reader.read_sheet_by_index(sheet_index, columns)
    num_row = len(result)

    # Then
    assert num_row == 29



