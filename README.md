# Column_Matching

## Overview
The `Column_Matching` is a Python-based tool designed to create matched pairs of IDs based on a numerical matching column and gender (if applicable). This package is useful for researchers needing structured participant matching based on predefined criteria.

### Author
**Michael Khela**  
Email: [michael.khela99@gmail.com](mailto:michael.khela99@gmail.com)

## Requirements
- **Python** 3.12.1
- Required Python libraries:
  - `pandas` (2.2.0)
  - `openpyxl` (3.1.2)

To install dependencies, run:
```sh
pip install pandas openpyxl
```

## Installation
1. Clone or download the `Column_Matching` repository.
2. Copy the `Column_Matching` folder to your working directory.
3. Ensure your CSV input file is structured correctly with the required columns.

## Usage
Run the `Group_Matching.py` script to process participant data:
```sh
python Group_Matching.py
```

### Input File Format
The input CSV file must contain the following columns:
- **ID** (Unique participant identifier)
- **cohort** (Grouping category)
- **matching** (Numerical value for matching)
- **sex** (Optional; used for gender-based matching)

### Output
- The script generates an output file with matched pairs in the `Outputs/` directory.

## Contact
For issues or inquiries, contact **Michael Khela** at [michael.khela99@gmail.com](mailto:michael.khela99@gmail.com).

