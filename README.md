# CSV File Merger and Plotter

## Overview

`FileSelector` is a Python-based GUI application that allows users to select a folder containing CSV files, automatically merge them into a single Excel file, and generate a plot based on the data. The application is designed to handle CSV files with a specific structure and outputs both a combined Excel file and a visual plot of the data.

## Features

- **Folder Selection:** Allows the user to select a folder containing CSV files.
- **CSV Merging:** Merges all CSV files in the selected folder into a single Excel file.
- **Data Plotting:** Generates a combined plot from the data in the CSV files and saves it as an image.
- **Excel Export:** Saves the combined data and the generated plot into an Excel file.

## Requirements

- Python 3.x
- Required Python libraries:
  - `numpy`
  - `pandas`
  - `xlsxwriter`
  - `matplotlib`
  - `PyQt6`

You can install the required libraries using the following command:

```bash
pip install numpy pandas xlsxwriter matplotlib PyQt6
