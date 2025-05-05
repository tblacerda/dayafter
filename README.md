# Day After Report Generator

This project is a Python-based tool for generating detailed reports on 4G and 5G network performance metrics. It processes Excel files containing network data, calculates various metrics, and generates a comprehensive PDF report with visualizations and summaries.

## Features

- **Data Processing**: Load and process Excel files for 4G and 5G technologies.
- **Metrics Calculation**: Calculate key performance indicators such as data volume, throughput, user peaks, and accessibility.
- **Visualization**: Generate time series plots, boxplots, and facet plots for network metrics.
- **PDF Report**: Create a detailed PDF report with summaries and visualizations for each group and technology.

## Project Structure

```
__DAY_AFTER/
├── 4G/
│   └── (2025-05-05)-r4g_cell-r4g_cell_cellname.xlsx
├── 5G/
│   └── (2025-05-05)-r5g_cell-r5g_cell_cellname.xlsx
├── dayafterRev5.py
├── report_GARANHUNS_MAI25.pdf
```

## Requirements

- Python 3.8+
- Required Python libraries:
  - pandas
  - matplotlib
  - openpyxl

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   ```
2. Navigate to the project directory:
   ```bash
   cd __DAY_AFTER
   ```
3. Install the required Python libraries:
   ```bash
   pip install pandas matplotlib openpyxl
   ```

## Usage

1. Place the 4G and 5G Excel files in their respective folders (`4G/` and `5G/`).
2. Run the script:
   ```bash
   python dayafterRev5.py
   ```
3. The generated PDF report will be saved as `report_GARANHUNS_MAI25.pdf` in the project directory.

## Configuration

The script uses predefined configurations for processing 4G and 5G data. These configurations include column renaming, dropping unnecessary columns, and applying unit conversions. You can modify these configurations in the `tech_configs` dictionary in `dayafterRev5.py`.

## Output

The output is a PDF report containing:
- Summary metrics for each group.
- Time series plots for key metrics.
- Boxplots for throughput and user distribution.
- Facet plots for user trends per cell.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Author

Developed by Tiago Beltrão Lacerda