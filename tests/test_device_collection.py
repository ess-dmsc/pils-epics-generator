from src.reader import ExcelReader
from src.device import Device, DeviceCollection

import pytest


@pytest.fixture
def reader():
    return ExcelReader("test.xlsx").read_sheet_by_index(0, [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15])


@pytest.fixture
def reader_nm():
    return ExcelReader("no_motor_test.xlsx").read_sheet_by_index(0, [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15])


def test_can_create_device_collection_xml(reader):
    df, instrument_name = reader
    device_collection = DeviceCollection(instrument_name)
    device_collection.from_dataframe(df)
    assert True


def test_can_create_without_motor_and_extra_devices(reader_nm):
    df, instrument_name = reader_nm
    device_collection = DeviceCollection(instrument_name)
    device_collection.from_dataframe(df)
    assert True
