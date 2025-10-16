# BODs XLS Export

![BOD Counter Interface](image.png)

Automated system for counting and exporting BODs (Bulk Order Deeds) from Ultima Online to Excel spreadsheets.

## Description

This project automates the process of counting BODs organized in specific bags in Ultima Online, generating detailed reports in Excel format (.xlsx) with visual formatting based on material type (Verite, Agapite, Gold, etc.).

## Features

- **Modern Dark Theme Interface** - Professional and easy on the eyes
- **Automatic BOD counting** in configured bags
- **Excel export** with color formatting by material
- **Custom export directory** selection
- **Real-time status updates** with progress indicators
- **One-click folder access** to view generated reports
- Support for different material types (Verite, Agapite, Gold, Valorite, Bronze, Copper)
- Reports organized by item type (LBOD, COIF, LEGS, TUNIC, ARMS, GLOVES, GORGET, HELM)
- Empty cells highlighted in red for easy identification

## Prerequisites

- **Python 3.8 or higher**
- **UO Stealth** installed and configured
- **Ultima Online** running

## Installation

### 1. Clone or download the project

```bash
git clone <your-repository>
cd bodsxlsexport
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

**Or install manually:**

```bash
pip install pandas openpyxl
```

## ⚙️ Configuration

### 1. Configure bag IDs in `bs_config.py` file

Open the `bs_config.py` file and edit the bag IDs according to your in-game setup:

```python
veritePackList = {
    '20veritelbod': 0x43ab4b98,  # Replace with your bag ID
    '20veritelegs': 0x43ab4b96,
    # ... add or remove as needed
}

agapitePackList = {
    '10agapitelbod': 0x43a6b818,
    # ... configure your bags
}

goldPackList = {
    'copperlbod': 0x43a37e25,
    # ... configure your bags
}
```

**How to get a bag ID:**
1. In UO Stealth, use the command to get object IDs
2. Click on the desired bag
3. Copy the hexadecimal ID and paste it in the configuration file

### 2. Configure export path

In the `xlsGenerator.py` file, line 13, adjust the path where reports will be saved:

```python
full_path = os.path.join("D:\\UOSTEALTH2025\\Scripts\\BodCollector", filename)
```

Replace with your desired path.

## How to Use

### 1. Start UO Stealth

Make sure UO Stealth is running and your character is logged into the game.

### 2. Load the script in UO Stealth

1. Open **UO Stealth**
2. Go to **Script** menu
3. Click **Open** and select `countBodsGenXLS.py`
4. Click the **Play** button to start the script

### 3. Use the interface

1. **Select export directory** - Click "Browse" to choose where to save reports
2. **Select the BOD collection** you want to count (Verite, Agapite or Gold)
3. Click ** Start** to begin counting
4. **Monitor the status** - Watch real-time progress updates
5. **Open reports** - Click ** Open** to view generated Excel files

### 4. Locate the report

The file will be saved in the configured directory with the name:
```
bod_report_YYYYMMDD_HHMMSS.xlsx
```

Example: `bod_report_20251016_143025.xlsx`

## Report Format

The generated Excel report contains:

- **Columns**: Organized by material and quantity (e.g., VERITE 20e, AGAPITE 15e)
- **Rows**: Item types (LBOD, COIF, LEGS, etc.)
- **Colors**:
  - 🟢 Green: Verite
  - 🔵 Blue: Valorite
  - 🟣 Purple: Agapite
  - 🟡 Yellow: Gold
  - 🔴 Red: Zero quantity (missing)

## Project Structure

```
bodsxlsexport/
│
├── countBodsGenXLS.py      # Main script with modern GUI
├── xlsGenerator.py         # Excel spreadsheet generator
├── bs_config.py            # Bag configuration and IDs
├── requirements.txt        # Project dependencies
├── README.md              # This documentation
├── image.png              # Interface screenshot
│
└── modules/
    ├── common_utils.py    # Utility functions
    └── connection.py      # UO Stealth connection functions
```


**Note:** This project is an automation tool for personal use in Ultima Online. Make sure you comply with the rules of the server where you play.
