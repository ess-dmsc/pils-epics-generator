import argparse
import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(script_dir, '..')
sys.path.append(project_root)

from src.device import DeviceCollection
from src.reader import ExcelReader


def main():
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Generate XML configuration for devices from an Excel file.")
    parser.add_argument("-p", "--path", help="Path to the Excel file containing device data.", required=True)

    # Parse arguments from the command line
    args = parser.parse_args()

    # Read devices from the Excel file
    excel_reader = ExcelReader(args.path)
    df, instrument_name = excel_reader.read_sheet_by_index(0, [1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15])

    # Create a DeviceCollection and populate it from the DataFrame
    device_collection = DeviceCollection(instrument_name)
    device_collection.from_dataframe(df)

    # Generate XML from the device collection
    device_collection.to_xml()


if __name__ == "__main__":
    main()
