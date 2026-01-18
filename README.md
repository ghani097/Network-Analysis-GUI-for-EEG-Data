# PLI Network Analysis GUI

A graphical interface for analyzing Phase Lag Index (PLI) data across brain networks using mixed-effects models.

## Features

- Load PLI data from Excel files
- Baseline adjustment (ANCOVA) with Pre-session as covariate
- Mixed-effects modeling with statsmodels
- Between-group and within-session contrasts
- Automatic visualization with significance markers
- Export results to Excel and CSV

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/PLI-Network-Analysis-GUI.git
   cd PLI-Network-Analysis-GUI
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Windows
Double-click `run_gui.bat` or run:
```bash
python network_analysis_gui.py
```

### Command Line
```bash
python network_analysis.py --input data.xlsx --output results
python network_analysis.py --no-baseline  # Disable baseline adjustment
```

## Input Data Format

The input Excel file should contain the following columns:
- `Participant` - Participant ID
- `Group` - Group label (e.g., "Chiro", "Control")
- `Session` - Session name (e.g., "Pre", "Post", "Post4W")
- `Network` - Brain network (e.g., "DMN", "SN", "CEN")
- `FrequencyTag` - Frequency band (e.g., "Alpha", "Beta")
- `MeanPLI` - Mean Phase Lag Index value

A test data file is included in `test_data/PLI-UK-Both-Groups-UP3.xlsx`.

## Output

The analysis generates:
- Individual plots for each network/frequency combination
- Combined results figure (`combined_results.png`)
- Statistics Excel file (`analysis_statistics.xlsx`)
- Summary CSV (`summary.csv`)

## Requirements

- Python 3.8+
- pandas
- numpy
- matplotlib
- scipy
- statsmodels
- PyQt5
- openpyxl

## License

MIT License
