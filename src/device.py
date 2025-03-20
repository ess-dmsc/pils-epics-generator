import os

import pandas as pd
from xml.etree.ElementTree import Element, SubElement, ElementTree
import xml.dom.minidom

VERSION = "1.0.0"

pils_device_byte_aligments = {
    '1201':  2,  # Simple discrete input, 16 bit signed integer
    '1202':  4,  # Simple discrete input, 32 bit signed integer
    '1204':  8,  # Simple discrete input, 64 bit signed integer
    '1302':  4,  # Simple analog input, 32 bit floating point, real
    '1304':  8,  # Simple analog input, 64 bit floating point, double
    '1602':  2,  # Simple discrete output, 16 bit signed integer
    '1604':  4,  # Simple discrete output, 32 bit signed integer
    '1608':  8,  # Simple discrete output, 64 bit signed integer
    '1704':  4,  # Simple analog output, 32 bit floating point, real
    '1708':  8,  # Simple analog output, 64 bit floating point, double
    '1802':  4,  # Extended status word that has 24 AUX bits
    '1A04':  4,  # discrete input, 32 bit signed integer + extended status word
    '1A08':  8,  # discrete input, 64 bit signed integer + extended status word + errorID
    '1B04':  4,  # analog input, 32 bit floating point + status word
    '1B08':  8,  # analog input, 64 bit floating point + extended status word + errorID
    '1E04':  4,  # discrete output, 16 bit signed integer + extended status word
    '1E06':  4,  # discrete output, 32 bit signed integer + extended status word
    '1E0C':  8,  # discrete output, 64 bit signed integer + extended status word + errorID
    '1F06':  4,  # analog output, 32 bit signed integer + extended status word
    '1F0C':  8,  # analog output, 64 bit signed integer + extended status word + errorID
    '5010':  8   # param device with 64 bit float, motor
}

pils_device_byte_lengths = {
    '1201':  2,
    '1202':  4,
    '1204':  8,
    '1302':  4,
    '1304':  8,
    '1602':  4,
    '1604':  8,
    '1608':  16,
    '1704':  4,
    '1708':  8,
    '1802':  4,
    '1A04':  8,
    '1A08':  16,
    '1B04':  8,
    '1B08':  16,
    '1E04':  8,
    '1E06':  12,
    '1E0C':  24,
    '1F06':  12,
    '1F0C':  24,
    '5010':  32
}

pils_temp_units = {
    'c': '16#0009',
    'k': '16#0008'
}

pils_units = {
    'mm': '16#FD04',
    'degree': '16#000C',
}

BASE_CSS = """
  <widget typeId="org.csstudio.opibuilder.widgets.ActionButton" version="2.0.0">
    <actions hook="false" hook_all="false">
      <action type="OPEN_DISPLAY">
        <path>motor-6.opi</path>
        <macros>
          <include_parent_macros>true</include_parent_macros>
          <PREFIX>YMIR-MCS1:MC-MCU-01:</PREFIX>
          <P>YMIR-</P>
          <M1>mcs1:MC-Spare-01:Mtr</M1>
          <M2>mcs1:MC-Spare-02:Mtr</M2>
          <M3>ColSl1:MC-SlYp-01:Mtr</M3>
          <M4>ColSl1:MC-SlYm-01:Mtr</M4>
          <M5>ColSl1:MC-SlZp-01:Mtr</M5>
          <M6>ColSl1:MC-SlZm-01:Mtr</M6>
          <R1>mcs1:MC-Spare-01:Mtr-</R1>
          <R2>mcs1:MC-Spare-02:Mtr-</R2>
          <R3>ColSl1:MC-SlYp-01:Mtr-</R3>
          <R4>ColSl1:MC-SlYm-01:Mtr-</R4>
          <R5>ColSl1:MC-SlZp-01:Mtr-</R5>
          <R6>ColSl1:MC-SlZm-01:Mtr-</R6>
        </macros>
        <mode>1</mode>
        <description></description>
      </action>
    </actions>
    <alarm_pulsing>false</alarm_pulsing>
    <backcolor_alarm_sensitive>false</backcolor_alarm_sensitive>
    <background_color>
      <color red="240" green="240" blue="240" />
    </background_color>
    <border_alarm_sensitive>false</border_alarm_sensitive>
    <border_color>
      <color red="0" green="128" blue="255" />
    </border_color>
    <border_style>0</border_style>
    <border_width>1</border_width>
    <enabled>true</enabled>
    <font>
      <opifont.name fontName=".SF NS Text" height="11" style="0" pixels="false">Default</opifont.name>
    </font>
    <forecolor_alarm_sensitive>false</forecolor_alarm_sensitive>
    <foreground_color>
      <color red="0" green="0" blue="0" />
    </foreground_color>
    <height>31</height>
    <image></image>
    <name>Action Button_15</name>
    <push_action_index>0</push_action_index>
    <pv_name></pv_name>
    <pv_value />
    <rules />
    <scale_options>
      <width_scalable>true</width_scalable>
      <height_scalable>true</height_scalable>
      <keep_wh_ratio>false</keep_wh_ratio>
    </scale_options>
    <scripts />
    <style>0</style>
    <text>YMIR-MCS1</text>
    <toggle_button>false</toggle_button>
    <tooltip>$(pv_name)
$(pv_value)</tooltip>
    <visible>true</visible>
    <widget_type>Action Button</widget_type>
    <width>180</width>
    <wuid>-23224cdf:15ca5e28b9e:-7fd5</wuid>
    <x>0</x>
    <y>0</y>
  </widget>
"""


def get_next_mb(memory_offset: int, device_type: str) -> int:
    """
    Calculate the next memory offset for a given device type.

    :param memory_offset: The current memory offset.
    :param device_type: The type of the device.
    :return: The next memory offset.
    """

    return memory_offset + pils_device_byte_lengths[device_type]


def align_mb(memory_offset: int, device_type: str) -> int:
    """
    Aligns the memory offset based on the device's byte alignment requirements.

    :param memory_offset: The current memory offset.
    :param device_type: The type of the device.
    :return: The aligned memory offset.
    """

    pils_device_byte_aligment = pils_device_byte_aligments[device_type]
    misaligned = memory_offset % pils_device_byte_aligment
    if misaligned != 0:
        memory_offset = memory_offset - misaligned + pils_device_byte_aligment

    return memory_offset


class Device:
    """
    Represents a device with its configurations and properties.
    """

    def __init__(self, description: str, pv_name: str, pv_root : str, mc_unit: int, ptp: bool, mc_axis_nc: int, mc_axis_pn: int,
                 device_type: str, pils_name: str, pils_unit: str, has_temp: bool = False, temp_units: str = 'c',
                 has_extra: bool = False, extra_name: str = '', extra_type: str = '', extra_desc: str = '') -> None:
        """
        Initializes a new instance of the Device class.

        :param description: Description of the device.
        :param pv_name: Process variable name of the device.
        :param mc_unit: Motion control unit number.
        :param mc_axis_nc: Motion control axis number (normal control).
        :param mc_axis_pn: Motion control axis number (pneumatic).
        :param device_type: The type identifier of the device.
        :param has_temp: Indicates if the device has a temperature sensor.
        :param temp_units: The unit of temperature measurement.
        :param has_extra: Indicates if the device has an extra device.
        :param extra_name: Name of the extra device.
        :param extra_type: Type identifier of the extra device.
        :param extra_desc: Description of the extra device.
        """
        self.description = description
        self.pv_name = pv_name
        self.pv_root = pv_root
        self.mc_unit = mc_unit
        self.ptp = ptp
        self.mc_axis_nc = mc_axis_nc
        self.mc_axis_pn = mc_axis_pn
        self.pils_name = pils_name
        self.pils_unit = pils_unit
        self.device_type = device_type
        self.has_temp = has_temp
        self.temp_units = temp_units
        self.has_extra = has_extra
        self.extra_name = extra_name
        self.extra_type = extra_type
        self.extra_desc = extra_desc

    @classmethod
    def from_dataframe_row(cls, row: pd.Series) -> 'Device':
        """
        Factory method to create a Device instance from a DataFrame row.

        :param row: A Series object representing a row from the DataFrame.
        :return: A Device instance.
        """
        is_pneumatic = not pd.isna(row['mc_axis_pn'])
        is_motor = not pd.isna(row['mc_axis_nc'])
        device_type = '1E04' if is_pneumatic else '5010'

        if not is_motor and not is_pneumatic:
            device_type = str(row['extra_type']) if not pd.isna(row['extra_type']) else None
            if device_type is None:
                device_type = '1302' if not pd.isna(row['has_temp']) else None

            if device_type is None:
                raise ValueError(f"Device type {row['extra_type']} not defined")

        return cls(
            description=row['axis_description'],
            pv_name=row['pv_name'] if not row['pv_name'] == 0 else None,
            pv_root=row['pv_root'] if not pd.isna(row['pv_root']) else '',
            mc_unit=int(row['mc_unit']),
            ptp=True if str(row['ptp']).lower() == 'yes' else False,
            mc_axis_nc=int(row['mc_axis_nc']) if not pd.isna(row['mc_axis_nc']) else None,
            mc_axis_pn=int(row['mc_axis_pn']) if not pd.isna(row['mc_axis_pn']) else None,
            device_type=device_type,
            pils_name=row['pils_name'],
            pils_unit=row['pils_unit'] if not pd.isna(row['pils_unit']) else 'mm',
            has_temp=True if not pd.isna(row['has_temp']) else False,
            temp_units=row['temp_units'] if not pd.isna(row['temp_units']) else 'c',
            has_extra=True if not pd.isna(row['extra_dev']) else False,
            extra_name=row['extra_name'] if not pd.isna(row['extra_name']) else '',
            extra_type=str(row['extra_type']) if not pd.isna(row['extra_type']) else '',
            extra_desc=row['extra_desc'] if not pd.isna(row['extra_desc']) else ''
        )


class DeviceCollection:
    """
    Represents a collection of devices grouped by their motion control unit.
    """

    def __init__(self, instrument: str = '') -> None:
        """
        Initializes a new instance of the DeviceCollection class.

        :param instrument: The name of the instrument the devices belong to.
        """

        # Change to a dictionary to group devices by mc_unit
        self.devices_by_unit = {}
        self.instrument = instrument

    def add_device(self, device: Device) -> None:
        """
        Adds a device to the collection.

        :param device: The device to be added.
        """

        if device.mc_unit not in self.devices_by_unit:
            self.devices_by_unit[device.mc_unit] = [device]
        else:
            self.devices_by_unit[device.mc_unit].append(device)

    def from_dataframe(self, df: pd.DataFrame) -> None:
        """
        Populates the device collection from a pandas DataFrame.

        :param df: The DataFrame containing device information.
        """
        for _, row in df.iterrows():
            self.add_device(Device.from_dataframe_row(row))

    def xml_define_5010(self, device, idx, current_offset):
        device_info = []
        num_devices = 0
        current_offset = align_mb(current_offset, device.device_type)
        motor_name = f"MotorM{device.mc_axis_nc}"
        device_str = f"st{motor_name} AT %MB{current_offset}: ST_{device.device_type};"
        fb_device_str = f"fb{motor_name}: FB_{device.device_type}_Axis := (nPILSDeviceNumber := {idx});"
        param_device_str = f"st{motor_name}Param: ST_AxisParameters;"
        current_offset = get_next_mb(current_offset, device.device_type)
        device_info.append(device_str)
        device_info.append(fb_device_str)
        device_info.append(param_device_str)
        idx += 1
        num_devices += 1

        if device.has_temp:
            current_offset = align_mb(current_offset, '1302')
            device_info.append(f"st{motor_name}Temp AT %MB{current_offset}: ST_1302;")
            current_offset = get_next_mb(current_offset, '1302')
            idx += 1
            num_devices += 1
        if device.has_extra:
            current_offset = align_mb(current_offset, device.extra_type)
            device_info.append(f"{device.extra_name} AT %MB{current_offset}: ST_{device.extra_type};")
            current_offset = get_next_mb(current_offset, device.extra_type)
            idx += 1
            num_devices += 1
        return device_info, idx, current_offset, num_devices

    def xml_describe_5010(self, device, current_offset):
        device_info = []
        current_offset = align_mb(current_offset, device.device_type)
        device_info_str = f"(nTypCode := 16#{device.device_type}, sName := '{device.pils_name}', nOffset := {current_offset}, nUnit := {pils_units[device.pils_unit]}, asAUX := [(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),(''),('InterlockFwd'),('InterlockBwd'),('localMode'),('inTargetPos'),('homeSensor'),('notHomed'),('enabled')], nFlags := 1),"
        current_offset = get_next_mb(current_offset, device.device_type)
        device_info.append(device_info_str)
        if device.has_temp:
            current_offset = align_mb(current_offset, '1302')
            temp_info_str = f"(nTypCode := 16#1302, sName := 'Temp#{device.mc_axis_nc}', nOffset := {current_offset}, nUnit := {pils_temp_units[device.temp_units]}),"
            current_offset = get_next_mb(current_offset, '1302')
            device_info.append(temp_info_str)
        if device.has_extra:
            current_offset = align_mb(current_offset, device.extra_type)
            extra_info_str = f"(nTypCode := 16#{device.extra_type}, sName := '{device.extra_desc}', nOffset := {current_offset}),"
            current_offset = get_next_mb(current_offset, device.extra_type)
            device_info.append(extra_info_str)
        return device_info, current_offset

    def xml_define_1E04(self, device, idx, current_offset):
        device_info = []
        num_devices = 0
        current_offset = align_mb(current_offset, device.device_type)
        motor_name = f"PneumaticP{device.mc_axis_pn}"
        device_str = f"st{motor_name} AT %MB{current_offset}: ST_{device.device_type};"
        fb_device_str = f"fb{motor_name}: FB_{device.device_type}_Pneumatic := (nPILSDeviceNumber := {idx});"
        current_offset = get_next_mb(current_offset, device.device_type)
        device_info.append(device_str)
        device_info.append(fb_device_str)
        idx += 1
        num_devices += 1

        if device.has_temp:
            current_offset = align_mb(current_offset, '1302')
            device_info.append(f"st{motor_name}Temp AT %MB{current_offset}: ST_1302;")
            current_offset = get_next_mb(current_offset, '1302')
            idx += 1
            num_devices += 1

        return device_info, idx, current_offset, num_devices

    def xml_describe_1E04(self, device, current_offset):
        device_info = []
        current_offset = align_mb(current_offset, device.device_type)
        device_info_str = f"(nTypCode := 16#{device.device_type}, sName := '{device.pils_name}', nOffset := {current_offset}, asAux := [('Closed'),('Closing'),('Opening'),('Opened'),('InTheMiddle'),(''),(''),(''),(''),(''),(''),(''),(''),(''),('Interlocked'),('PSSPermitDenied'),('SolenoidActive'),('Retracted'),('Extended'),('Retracting'),('Extending'),(''),(''),('')]),"
        current_offset = get_next_mb(current_offset, device.device_type)
        device_info.append(device_info_str)
        if device.has_temp:
            current_offset = align_mb(current_offset, '1302')
            temp_info_str = f"(nTypCode := 16#1302, sName := 'Temp#{device.mc_axis_pn}', nOffset := {current_offset}, nUnit := {pils_temp_units[device.temp_units]}),"
            current_offset = get_next_mb(current_offset, '1302')
            device_info.append(temp_info_str)
        return device_info, current_offset

    def xml_define_1302(self, device, idx, current_offset):
        current_offset = align_mb(current_offset, '1302')
        device_info = f"{device.extra_name} AT %MB{current_offset}: ST_1302;"
        current_offset = get_next_mb(current_offset, '1302')
        idx += 1
        return [device_info], idx, current_offset, 1

    def xml_describe_1302(self, device, current_offset):
        current_offset = align_mb(current_offset, '1302')
        device_info = f"(nTypCode := 16#1302, sName := '{device.extra_desc}', nOffset := {current_offset}, nUnit := {pils_temp_units[device.temp_units]}),"
        current_offset = get_next_mb(current_offset, '1302')
        return [device_info], current_offset

    def xml_define_1A04(self, name, idx, current_offset):
        current_offset = align_mb(current_offset, '1A04')
        device_info = f"{name} AT %MB{current_offset}: ST_1A04;"
        current_offset = get_next_mb(current_offset, '1A04')
        idx += 1
        return [device_info], idx, current_offset, 1

    def xml_describe_1A04(self, name, current_offset, asAux=''):
        current_offset = align_mb(current_offset, '1A04')
        if asAux:
            device_info = f"(nTypCode := 16#1A04, sName := '{name}', nOffset := {current_offset}, asAux := {asAux}),"
        else:
            device_info = f"(nTypCode := 16#1A04, sName := '{name}', nOffset := {current_offset}),"
        current_offset = get_next_mb(current_offset, '1A04')
        return [device_info], current_offset

    def xml_define_1201(self, name, idx, current_offset):
        current_offset = align_mb(current_offset, '1201')
        device_info = f"{name} AT %MB{current_offset}: ST_1201;"
        current_offset = get_next_mb(current_offset, '1201')
        idx += 1
        return [device_info], idx, current_offset, 1

    def xml_describe_1201(self, name, current_offset):
        current_offset = align_mb(current_offset, '1201')
        device_info = f"(nTypCode := 16#1201, sName := '{name}', nOffset := {current_offset}),"
        current_offset = get_next_mb(current_offset, '1201')
        return [device_info], current_offset

    def xml_define_1204(self, name, idx, current_offset):
        current_offset = align_mb(current_offset, '1204')
        device_info = f"{name} AT %MB{current_offset}: ST_1204;"
        current_offset = get_next_mb(current_offset, '1204')
        idx += 1
        return [device_info], idx, current_offset, 1

    def xml_describe_1204(self, name, current_offset):
        current_offset = align_mb(current_offset, '1204')
        device_info = f"(nTypCode := 16#1204, sName := '{name}', nOffset := {current_offset}, nUnit := 16#F711),"
        current_offset = get_next_mb(current_offset, '1204')
        return [device_info], current_offset

    def xml_define_1B08(self, current_offset):
        current_offset = align_mb(current_offset, '1B08')
        device_info = f"stPressureSensor AT %MB{current_offset}: ST_1B08;\n"
        current_offset = get_next_mb(current_offset, '1B08')
        return device_info, current_offset, 1

    def xml_describe_1B08(self, current_offset):
        current_offset = align_mb(current_offset, '1B08')
        device_info = f"(nTypCode := 16#1B08, sName := 'SysPressureValue', nOffset := {current_offset}),"
        current_offset = get_next_mb(current_offset, '1B08')
        return device_info, current_offset

    def xml_define_1802(self, current_offset):
        current_offset = align_mb(current_offset, '1802')
        device_info = f"stCabinetStatus AT %MB{current_offset}: ST_1802;\n"
        current_offset = get_next_mb(current_offset, '1802')
        return device_info, current_offset, 1

    def xml_describe_1802(self, current_offset, is_last=False):
        current_offset = align_mb(current_offset, '1802')
        device_info = f"(nTypCode := 16#1802, sName := 'Cabinet#0', nOffset := {current_offset}, asAux := [('24VPSFailed'), ('48VPSFailed'), ('MCBError'), ('SPDError'), ('DoorOpen'), ('TempHigh'), ('FuseTripped'), ('EStop'), ('ECMasterError'), ('SlaveNotOP'), ('SlaveMissing'), ('CPULoadHigh'),(''), (''), (''), (''), (''), (''), (''),(''), (''), (''), (''), ('')])"
        if is_last:
            device_info += "];"
        current_offset = get_next_mb(current_offset, '1802')
        return device_info, current_offset

    def xml_define_extra(self, device, index, current_offset):
        if device.device_type == '1302':
            return self.xml_define_1302(device, index, current_offset)
        else:
            raise NotImplementedError

    def xml_describe_extra(self, device, current_offset):
        if device.device_type == '1302':
            return self.xml_describe_1302(device, current_offset)
        else:
            raise NotImplementedError

    def build_ptp_define(self, idx, current_offset):
        device_info = []
        ptp_offset, idx, current_offset, num_new_devices = self.xml_define_1A04("stPTPOffset", idx, current_offset)
        ptp_state, idx, current_offset, num_new_devices = self.xml_define_1A04("stPTPState", idx, current_offset)
        ptp_sync, idx, current_offset, num_new_devices = self.xml_define_1201("stPTPSyncSeqNum", idx, current_offset)
        ptp_error, idx, current_offset, num_new_devices = self.xml_define_1A04("stPTPErrorStatus", idx, current_offset)
        sys_time, idx, current_offset, num_new_devices = self.xml_define_1204("stSystemUTCtime", idx, current_offset)
        device_info.extend([ptp_offset[0], ptp_state[0], ptp_sync[0], ptp_error[0], sys_time[0], ""])
        return device_info, idx, current_offset

    def build_ptp_describe(self, current_offset):
        device_info = []
        ptp_offset, current_offset = self.xml_describe_1A04("PTPOffset#0", current_offset)
        ptp_state, current_offset = self.xml_describe_1A04("PTPState#0", current_offset)
        ptp_sync, current_offset = self.xml_describe_1201("PTPSyncSeqNum#0", current_offset)
        ptp_error, current_offset = self.xml_describe_1A04("PTPErrorStatus#0", current_offset,
            asAux="[(''), (''), (''), (''), (''), (''), (''), (''), (''), (''), ('NotFullySynched'), ('NotSynchronized'),('NotPTPslave'), ('NotPTPv2'), ('RdDiagError'), ('CableNotConnected_PTPnotStarted'), (''), (''), (''),(''), (''), (''), (''), ('')]"
        )
        sys_time, current_offset = self.xml_describe_1204("SystemUTCtime#0", current_offset)
        device_info.extend([ptp_offset[0], ptp_state[0], ptp_sync[0], ptp_error[0], sys_time[0]])
        return device_info, current_offset

    def build_definition(self, devices):
        device_definitions = []
        idx = 1
        num_devices = 0
        curr_offset = 128
        pneumatic_exists = False
        for device in devices:
            if device.mc_axis_nc is not None:
                device_info, new_idx, curr_offset, num_new_devices = self.xml_define_5010(device, idx, curr_offset)
            elif device.mc_axis_pn is not None:
                device_info, new_idx, curr_offset, num_new_devices = self.xml_define_1E04(device, idx, curr_offset)
                pneumatic_exists = True
            elif device.mc_axis_pn is None and device.mc_axis_nc is None:
                device_info, new_idx, curr_offset, num_new_devices = self.xml_define_extra(device, idx, curr_offset)
            else:
                raise ValueError("undefined device")
            idx = new_idx
            num_devices += num_new_devices
            device_definitions.extend(device_info + [""])

        if pneumatic_exists:
            device_info, curr_offset, num_new_devices = self.xml_define_1B08(curr_offset)
            num_devices += num_new_devices
            device_definitions.append(device_info)

        if devices[0].ptp:
            ptp_info, new_idx, curr_offset = self.build_ptp_define(idx, curr_offset)
            idx = new_idx
            num_devices += 5
            device_definitions.extend(ptp_info)

        device_info, curr_offset, num_new_devices = self.xml_define_1802(curr_offset)
        num_devices += num_new_devices
        device_definitions.append(device_info)

        return device_definitions, num_devices, pneumatic_exists

    def build_description(self, devices, num_devices, pneumatic_exists):
        # Add the devices array part
        device_description = []
        device_description.append("// Array of Devices")
        device_array_str = f"astDevices: ARRAY [1..{num_devices}] OF ST_DeviceInfo :=["
        device_description.append(device_array_str)

        curr_offset = 128
        for device in devices:
            if device.mc_axis_nc is not None:
                device_info, curr_offset = self.xml_describe_5010(device, curr_offset)
                device_description.extend(device_info)
            elif device.mc_axis_pn is not None:
                device_info, curr_offset = self.xml_describe_1E04(device, curr_offset)
                device_description.extend(device_info)
            elif device.mc_axis_pn is None and device.mc_axis_nc is None:
                device_info, curr_offset = self.xml_describe_extra(device, curr_offset)
                device_description.extend(device_info)
            else:
                raise ValueError("mc_axis_nc or mc_axis_pn must be set")

        if pneumatic_exists:
            device_info, curr_offset = self.xml_describe_1B08(curr_offset)
            device_description.append(device_info)

        if devices[0].ptp:
            ptp_info, curr_offset = self.build_ptp_describe(curr_offset)
            device_description.extend(ptp_info)

        device_info, curr_offset = self.xml_describe_1802(curr_offset, is_last=True)
        device_description.append(device_info)

        device_description[-1] = device_description[-1].rstrip(',')  # Remove the comma from the last device entry
        return device_description

    def get_xml_start(self, mc_unit):
        root = Element('TcPlcObject')
        root.set('Version', '1.1.0.1')
        root.set('ProductVersion', '3.1.4024.5')

        gvl = SubElement(root, 'GVL')
        gvl.set('Name', 'GVL_PILS')
        gvl.set('Id', '{ace53d5e-03a7-4e79-9a33-72de579eb8fd}')

        declaration = SubElement(gvl, 'Declaration')
        cdata_content = [
            'VAR_GLOBAL',
            f"sPLCName: STRING[34] := '{self.instrument.lower()}-mcs{mc_unit}';",  # TODO: Replace with actual PLC name
            f"sPLCVersion: STRING[34] := '{VERSION}';",
            "sPLCAuthor1: STRING[34] := 'ESS Lund';",
            "sPLCAuthor2: STRING[34] := 'ESS Lund';\n",
        ]
        return root, declaration, cdata_content

    def add_xml_end(self, declaration, cdata_content):
        declaration.text = "<![CDATA[" + "\n\t".join(cdata_content)
        declaration.text += "\nEND_VAR\n"
        declaration.text += "]]>"

    def finish_xml(self, root, declaration, cdata_content):
        self.add_xml_end(declaration, cdata_content)

        # Generate rough string with ElementTree
        rough_string = xml.etree.ElementTree.tostring(root, encoding='UTF-8')

        # Pretty print with minidom
        reparsed = xml.dom.minidom.parseString(rough_string)
        pretty_xml_str = reparsed.toprettyxml(indent="  ")

        # Manually prepend the desired XML declaration
        xml_declaration = '<?xml version="1.0" encoding="utf-8"?>'
        pretty_xml_str = pretty_xml_str.replace('<?xml version="1.0" ?>', xml_declaration)
        pretty_xml_str = pretty_xml_str.replace("&lt;", "<")
        pretty_xml_str = pretty_xml_str.replace("&gt;", ">").strip()
        return pretty_xml_str

    def to_xml(self) -> None:
        """
        Generates an XML file per motion control unit from the device collection.
        """
        for mc_unit, devices in self.devices_by_unit.items():
            xml_file_path = f"mc_unit_{mc_unit}.TcGVL"
            root, declaration, cdata_content = self.get_xml_start(mc_unit)

            device_definitions, num_devices, pneumatic_exists = self.build_definition(devices)
            cdata_content.extend(device_definitions)

            device_description = self.build_description(devices, num_devices, pneumatic_exists)
            cdata_content.extend(device_description)

            pretty_xml_str = self.finish_xml(root, declaration, cdata_content)

            # Ensure UTF-8 encoding is preserved by writing as bytes
            with open(xml_file_path, 'w', encoding='utf-8') as file:
                file.write(pretty_xml_str)

    def format_spare_motor(self, mc_unit, idx):
        if idx < 10:
            return f"MC-Spare-0{idx}"
        return f"MC-Spare-{idx}"

    def format_spare_pneumatic(self, mc_unit, idx):
        if idx < 10:
            return f"MC-Spare-0{idx}"
        return f"MC-Spare-{idx}"

    def to_st_cmd(self, ioc_ip, plc_ip, return_it=False):
        for mc_unit, devices in self.devices_by_unit.items():
            st_cmd_file_path = f"st.{self.instrument.lower()}-mcs{mc_unit}.iocsh"


            # Use a list to collect command lines
            commands = []

            # Filter devices with pv_name and mc_axis_nc not None
            # devices = [device for device in devices if device.pv_name is not None and device.mc_axis_nc is not None]
            num_devices = len(devices)

            # find the maximum number of devices per axis type
            # pn_devs = [device for device in devices if device.mc_axis_pn is not None]
            # nc_devs = [device for device in devices if device.mc_axis_nc is not None]
            #
            # num_devices = max(len(pn_devs), len(nc_devs))


            # Add commands for EPICS environment setup
            commands.extend([
                'require essioc',
                'require calc',
                'require ethercatmc',
                '',
                f'epicsEnvSet("MOTOR_PORT",    "MCU1")',
                f'epicsEnvSet("IPADDR",        "{plc_ip}")',
                f'epicsEnvSet("IPPORT",        "48898")',
                f'epicsEnvSet("AMSNETIDIOC",   "{ioc_ip}.1.1")',
                f'epicsEnvSet("ASYN_PORT",     "MC_CPU1")',
                # '# prefix for all, system in ESS naming convention',
                f'epicsEnvSet("SYSPFX",        "{self.instrument.upper()}-")',
                # '# prefix for all MCU-ish records like PTP',
                f'epicsEnvSet("REG_NAME",      "MCS{mc_unit}:MC-MCU-0{mc_unit}:")',
                'epicsEnvSet("PREC",          "3")',
                f'epicsEnvSet("ECM_NUMAXES",   "{num_devices}")',
                f'epicsEnvSet("ECM_OPTIONS",   "adsPort=852;amsNetIdRemote={plc_ip}.1.1;amsNetIdLocal=$(AMSNETIDIOC)")',
                '',
                'epicsEnvSet("ECM_MOVINGPOLLPERIOD", "0")',
                'epicsEnvSet("ECM_IDLEPOLLPERIOD",   "0")',
                '',
                # '< ethercatmcController.iocsh',
                'iocshLoad("$(ethercatmc_DIR)ethercatmcController.iocsh")',
                ''
            ])

            # Add commands for each device
            spare_nc_idx = 1
            spare_pn_idx = 1
            idx = 1
            for device in devices:
                if device.mc_axis_nc is not None:
                    commands.extend([
                        '#',
                        f'# AXIS {idx}',
                        '#',
                        'epicsEnvSet("AXISCONFIG",      "")',
                        f'epicsEnvSet("AXIS_NAME",       "{device.pv_name if device.pv_name is not None else self.format_spare_motor(mc_unit, spare_nc_idx)}:Mtr")',
                        f'epicsEnvSet("AXIS_NO",         "{idx}")',
                        'epicsEnvSet("RAWENCSTEP_ADEL", "0")',
                        'epicsEnvSet("RAWENCSTEP_MDEL", "0")',
                        # '< ethercatmcIndexerAxis.iocsh',
                        # '< ethercatmcAxisdebug.iocsh',
                        'iocshLoad("$(ethercatmc_DIR)ethercatmcIndexerAxis.iocsh")',
                        'iocshLoad("$(ethercatmc_DIR)ethercatmcAxisdebug.iocsh")',
                        ''
                    ])
                    idx += 1
                    if device.pv_name is None:
                        spare_nc_idx += 1
                elif device.mc_axis_pn is not None:
                    commands.extend([
                        '#',
                        f'# AXIS {idx}',
                        '#',
                        'epicsEnvSet("AXISCONFIG",      "")',
                        f'epicsEnvSet("AXIS_NAME",       "{device.pv_name if device.pv_name is not None else self.format_spare_pneumatic(mc_unit, spare_pn_idx)}:Sht")',
                        f'epicsEnvSet("AXIS_NO",         "{idx}")',
                        # '< ethercatmcShutter.iocsh',
                        'iocshLoad("$(ethercatmc_DIR)ethercatmcShutter.iocsh")',
                        ''
                    ])
                    idx += 1
                    if device.pv_name is None:
                        spare_pn_idx += 1

            # Add commands to start the poller
            commands.extend([
                'epicsEnvSet("MOVINGPOLLPERIOD", "9")',
                'epicsEnvSet("IDLEPOLLPERIOD",   "100")',
                'ethercatmcStartPoller("$(MOTOR_PORT)", "$(MOVINGPOLLPERIOD)", "$(IDLEPOLLPERIOD)")',
                '',
                'iocinit()'
            ])

            # Join the command lines with newline characters and write to file
            command_string = "\n".join(commands)

            if return_it:
                return command_string

            with open(st_cmd_file_path, 'w') as file:
                file.write(command_string)

    def to_opi(self):
        for mc_unit, devices in self.devices_by_unit.items():
            devices = [device for device in devices if device.mc_axis_nc is not None]
            num_devices = len(devices)

            file_name = 'base_motor.mid'
            file_path = os.path.join(os.path.dirname(__file__), file_name)

            macros = [
                f'<PREFIX>{self.instrument.upper()}-MCS{mc_unit}:MC-MCU-0{mc_unit}:</PREFIX>',
                f'<P>YMIR-</P>'
            ]
            for idx, device in enumerate(devices, start=1):
                name = f"{device.pv_name}:Mtr" if device.pv_name is not None else f"{self.format_spare_motor(mc_unit, idx)}:Mtr"
                macros.append(f'<M{idx}>{name}</M{idx}>')
            for idx, device in enumerate(devices, start=1):
                name = f"{device.pv_name}:Mtr" if device.pv_name is not None else f"{self.format_spare_motor(mc_unit, idx)}:Mtr"
                macros.append(f'<R{idx}>{name}-</R{idx}>')

            with open(file_path, 'r') as file:
                content = file.readlines()

            for i, line in enumerate(content):
                indent_length = len(line) - len(line.lstrip())
                if "$MACROS$" in line:
                    line = line.lstrip()
                    line = line.replace("$MACROS$", "\n".join([f"{' '*indent_length}{macro}" for macro in macros]))
                    content[i] = line

                if "$NUM_DEVS$" in line:
                    line = line.replace("$NUM_DEVS$", str(num_devices))
                    content[i] = line

                if "$MCU_NAME$" in line:
                    line = line.replace("$MCU_NAME$", f"{self.instrument.upper()}-MCS{mc_unit}")
                    content[i] = line

            final_string = "".join(content)

            with open(f"IOC-{self.instrument.upper()}-MCS{mc_unit}.mid", 'w') as file:
                file.write(final_string)
