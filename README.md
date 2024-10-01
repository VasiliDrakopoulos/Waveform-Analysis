# Photometry Data Analysis Tool

This Python tool allows for the analysis of photometry data, enabling the user to upload Excel files, select baseline and comparison periods, and calculate significant differences between groups. The tool performs both bootstrapped confidence intervals (CIs) and permutation tests to evaluate the data.

## Features

- **Bootstrap Confidence Interval Calculation:** Computes bootstrapped confidence intervals for data using the `bootstrap` function from SciPy.
- **Permutation Test:** Permutation tests are used to compare differences between groups and determine significance.
- **Excel Integration:** Load photometry data from Excel, analyze it, and save the results back to an Excel file.
- **Graphical User Interface (GUI):** The application uses a simple GUI for file uploads, setting periods, and saving results.

## How to Use

1. **Load Excel Data:** Click the "Load Excel File" button to load your photometry data in `.xlsx` format.
2. **Set Baseline and Comparison Periods:** Enter the row indices for the baseline and comparison periods.
3. **Analyze Data:** The application will analyze the loaded data, calculating bootstrapped confidence intervals and performing permutation tests between groups.
4. **Save Results:** Save the analyzed results to an Excel file, which includes confidence intervals, significance maps, and significant timings.

### Excel File Structure

- **Sheet 1:** Timings, where each row corresponds to a time point.
- **Subsequent Sheets:** Each sheet corresponds to a group of animals, where columns represent individual animals and rows represent z-scores at different time points.

## Dependencies

- **Python 3.x**
- **NumPy**
- **Pandas**
- **SciPy**
- **Tkinter** (for GUI)

## Installation

Ensure you have the following dependencies installed before running the application. You can install them by running the following command:

```bash
pip install -r requirements.txt
