# PILS Generator

The PILS Generator is a tool designed to convert device data from Excel files into XML PILS tables. It reads specified Excel files containing device configurations and generates corresponding XML files that describe each device's properties and settings.

## Features

- Reads device data from Excel files.
- Generates PILS table configuration files for devices.
- Supports various device types and properties, including motors and temperature sensors.

## Getting Started

### Prerequisites

Before running the PILS EPICS Generator, ensure you have the following installed:

- Python 3.11 (only one I have tested
- Dependencies listed in `requirements.txt`

### Installation

1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/ess-dmsc/pils-epics-generator.git
   ```

   or with ssh

   ```bash
   git clone git@github.com:ess-dmsc/pils-epics-generator.git
   ```

2. Navigate to the cloned directory:

   ```bash
   cd pils-epics-generator
   ```

3. **Create a virtual environment** to isolate the project dependencies. You can use either the built-in Python `venv` module (for a virtual environment) or Conda, if you prefer.

   - **Using `venv`**:

     If you don't have Python installed, install it first. Then, create a virtual environment by running:

     ```bash
     python -m venv venv
     ```

     This command creates a new directory `venv` in your project directory, containing the virtual environment.

     Activate the virtual environment with:

     - On **Windows**:

       ```cmd
       .\venv\Scripts\activate
       ```

     - On **macOS and Linux**:

       ```bash
       source venv/bin/activate
       ```

     Your command line will change to show the name of the activated environment. While activated, any Python or pip commands will use the versions in the virtual environment, not the global ones.

   - **Using Conda**:

     If you have Anaconda or Miniconda installed, create a Conda environment by running:

     ```bash
     conda create --name pils-env python=3.11
     ```

     Replace `python=3.11` with the version of Python you want to use.

     Activate the Conda environment with:

     ```bash
     conda activate pils-env
     ```

     You'll see the environment name `(pils-env)` in your terminal prompt, indicating that the environment is activated.

4. Install the required Python libraries:

   ```bash
   pip install -r requirements.txt
   ```

### Usage

To generate an XML configuration file from an Excel file, use the `generate_pils_table.py` script located in the `bin` directory. You must specify the path to the Excel file containing the device data using the `-p` or `--path` argument.

```bash
python bin/generate_pils_table.py -p /path/to/your/excel_file.xlsx --pils 1
```

If you want to generate the EPICS st.cmd files run:

```bash
python bin/generate_pils_table.py -p tests/test.xlsx --ioc 1 --ioc-ip '10.102.10.49' --plc-ip '10.102.10.44'
```

however, it currently puts the same plc ip in all the st.cmd files, so if you have multiple PLCs you will need to manually change the IP in the st.cmd files. This is TODO.

### Excel File Format

Look at the file tests/test.xlsx files to see the expected Excel file structure

Each row in the Excel file represents a device with its configurations and properties.

### Output

The script generates XML files in the same directory as the script is run. Each XML file corresponds to a motion control unit and includes the configuration for all devices associated with that unit.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues to improve the project.
