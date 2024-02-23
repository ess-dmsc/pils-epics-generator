import pandas as pd
import os
from src.reader import ExcelReader
from src.device import Device, DeviceCollection

import pytest


def data_file_path(filename):
    return os.path.join(os.path.dirname(__file__), filename)

@pytest.fixture
def reader():
    filepath = data_file_path("test.xlsx")
    return ExcelReader(filepath).read_sheet_by_index(0, [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15])

@pytest.fixture
def reader_nm():
    filepath = data_file_path("no_motor_test.xlsx")
    return ExcelReader(filepath).read_sheet_by_index(0, [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15])



@pytest.fixture
def device_collection():
    collection = DeviceCollection("TestInstrument")
    # Add mock devices to collection here
    return collection

@pytest.fixture
def mock_dataframe():
    # Define the columns and mock data for the DataFrame
    data = {
        'axis_description': ['Motor Axis 1', 'Pneumatic Axis 1', 'Temperature Sensor'],
        'pv_name': ['IOC:MOTOR1', 'IOC:PNEU1', 'IOC:TEMP1'],
        'mc_unit': [1, 1, 1],
        'mc_axis_nc': [1, None, None],
        'mc_axis_pn': [None, 1, None],
        'has_temp': ['x', '', 'x'],
        'temp_units': ['c', '', 'c'],
        'extra_dev': ['', '', 'x'],
        'extra_name': ['', '', 'stTemp#1Vacuum'],
        'extra_type': ['', '', '1302'],
        'extra_desc': ['', '', 'Temp#1Vacuum'],
    }
    df = pd.DataFrame(data)
    return df


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


def test_device_collection_add_device(device_collection):
    device = Device(
        description="Test Device",
        pv_name="PV001",
        mc_unit=1,
        mc_axis_nc=1,
        mc_axis_pn=None,
        device_type='5010',
        has_temp=False,
        temp_units='',
        has_extra=False,
        extra_name='',
        extra_type='',
        extra_desc=''
    )
    device_collection.add_device(device)
    assert device in device_collection.devices_by_unit[1]


def test_device_collection_from_dataframe(mock_dataframe):
    device_collection = DeviceCollection("TestInstrument")
    device_collection.from_dataframe(mock_dataframe)

    motor = device_collection.devices_by_unit[1][0]
    pneumatic = device_collection.devices_by_unit[1][1]
    temp_sensor = device_collection.devices_by_unit[1][2]

    assert len(device_collection.devices_by_unit) > 0
    assert motor.device_type == '5010'
    assert pneumatic.device_type == '1E04'
    assert temp_sensor.device_type == '1302'


def test_xml_define_5010(device_collection):
    device = Device(
        description="Motor Device",
        pv_name="PV_Motor",
        mc_unit=1,
        mc_axis_nc=1,
        mc_axis_pn=None,
        device_type='5010',
        has_temp=True,
        temp_units='c',
        has_extra=False,
        extra_name='',
        extra_type='',
        extra_desc=''
    )
    expected_xml_list = [
        'stMotorM1 AT %MB128: ST_5010;',
        'fbMotorM1: FB_5010_Axis := (nPILSDeviceNumber := 1);',
        'stMotorM1Param: ST_AxisParameters;',
        'stMotorM1Temp AT %MB160: ST_1302;'
    ]
    device_collection.add_device(device)
    xml_output, idx, current_offset, num_devices = device_collection.xml_define_5010(device, 1, 128)

    assert idx == 3
    assert current_offset == 164
    assert num_devices == 2
    assert xml_output == expected_xml_list


def test_xml_describe_5010(device_collection):
    device = Device(
        description="Motor Device",
        pv_name="PV_Motor",
        mc_unit=1,
        mc_axis_nc=1,
        mc_axis_pn=None,
        device_type='5010',
        has_temp=True,
        temp_units='c',
        has_extra=False,
        extra_name='',
        extra_type='',
        extra_desc=''
    )
    device_collection.add_device(device)
    expected_xml_list = [
        "(nTypCode := 16#5010, sName := GVL.astAxes[1].stDescription.sAxisName, nOffset := 128, nUnit := GVL.astAxes[1].stDescription.eUnit, asAUX := [(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),('InterlockFwd'),('InterlockBwd'),('localMode'),('inTargetPos'),('homeSensor'),('notHomed'),('enabled')], nFlags := 1),",
        "(nTypCode := 16#1302, sName := 'Temp#1', nOffset := 160, nUnit := 16#0009),"
    ]
    xml_output, current_offset = device_collection.xml_describe_5010(device, 128)

    assert current_offset == 164
    assert xml_output == expected_xml_list
