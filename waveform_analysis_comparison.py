import numpy as np
import pandas as pd
from scipy.stats import bootstrap
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.simpledialog import askinteger

# Functions for analysis
def boot_CI(data, n_resamples=10000):
    if data.size == 0 or np.all(np.isnan(data)):
        return np.nan, np.nan 
    bootstrapped_means = bootstrap((data,), np.mean, n_resamples=n_resamples, random_state=42)
    return bootstrapped_means.confidence_interval.low, bootstrapped_means.confidence_interval.high

def permTest_array(data1, data2, num_permutations=10000):
    observed_diff = np.mean(data1) - np.mean(data2)
    combined_data = np.concatenate([data1, data2])
    count = 0

    for _ in range(num_permutations):
        np.random.shuffle(combined_data)
        new_data1 = combined_data[:len(data1)]
        new_data2 = combined_data[len(data1):]
        new_diff = np.mean(new_data1) - np.mean(new_data2)
        if abs(new_diff) >= abs(observed_diff):
            count += 1

    p_value = count / num_permutations
    return observed_diff, p_value

def calculate_differences(data, baseline_period, comparison_period):
    baseline_mean = np.nanmean(data[baseline_period, :], axis=0)
    comparison_mean = np.nanmean(data[comparison_period, :], axis=0)
    return comparison_mean - baseline_mean

def get_significant_ranges(timings, significance_map):
    significant_times = timings[significance_map == 1]
    if len(significant_times) == 0:
        return []

    ranges = []
    start = significant_times[0]
    for i in range(1, len(significant_times)):
        if significant_times[i] > significant_times[i-1] + 1:
            ranges.append((start, significant_times[i-1]))
            start = significant_times[i]
    ranges.append((start, significant_times[-1]))

    return ranges

def analyze_photometry_data(groups, baseline_period, comparison_period, timings):
    results = {}
    significance_map = np.zeros(groups[0].shape[0])
    baseline_diff_map = np.zeros((len(groups), groups[0].shape[0]))

    # Analyze baseline differences within groups
    for i in range(len(groups)):
        diff = calculate_differences(groups[i], baseline_period, comparison_period)
        observed_diff, p_value = permTest_array(groups[i][baseline_period, :].flatten(), groups[i][comparison_period, :].flatten())
        if p_value < 0.05:
            baseline_diff_map[i, comparison_period] = 1

    # Compare groups
    for i in range(len(groups)):
        for j in range(i + 1, len(groups)):
            group1, group2 = groups[i], groups[j]
            diff1 = calculate_differences(group1, baseline_period, comparison_period)
            diff2 = calculate_differences(group2, baseline_period, comparison_period)

            if diff1.size == 0 or diff2.size == 0 or np.all(np.isnan(diff1)) or np.all(np.isnan(diff2)):
                print(f"Skipping comparison between Group {i+1} and Group {j+1} due to empty or invalid differences.")
                continue

            observed_diff, p_value = permTest_array(diff1, diff2)
            if p_value < 0.05:
                significance_map[comparison_period] = 1

            ci1 = boot_CI(diff1)
            ci2 = boot_CI(diff2)
            results[f'Group {i+1} vs Group {j+1}'] = {
                'Bootstrapped CI 1': ci1,
                'Bootstrapped CI 2': ci2,
                'Observed Difference': observed_diff,
                'p-value': p_value
            }

    group_sig_timings = get_significant_ranges(timings[comparison_period], significance_map[comparison_period])
    baseline_sig_timings = []
    for i in range(len(groups)):
        sig_times = get_significant_ranges(timings[comparison_period], baseline_diff_map[i, comparison_period])
        baseline_sig_timings.append(sig_times)

    return results, significance_map, baseline_diff_map, group_sig_timings, baseline_sig_timings

# Reading the Excel data
def read_excel_data(file_path):
    timings_df = pd.read_excel(file_path, sheet_name=0)
    timings = timings_df.iloc[:, 0].values

    groups = []
    sheet_names = pd.ExcelFile(file_path).sheet_names[1:]

    for sheet in sheet_names:
        group_df = pd.read_excel(file_path, sheet_name=sheet)
        print(f"Group {sheet} shape: {group_df.shape}") 
        groups.append(group_df.values)

    return groups, timings

# Saving the results to an Excel file
def save_results_to_excel(file_path, results, significance_map, group_sig_timings, baseline_sig_timings, timings):
    with pd.ExcelWriter(file_path) as writer:
        pd.DataFrame(timings, columns=['Timings']).to_excel(writer, sheet_name='Timings')

        for key, res in results.items():
            pd.DataFrame(res).to_excel(writer, sheet_name=key)

        pd.DataFrame(significance_map, columns=['Significance']).to_excel(writer, sheet_name='Significance Map')

        group_sig_df = pd.DataFrame(group_sig_timings, columns=['Start', 'End'])
        group_sig_df.to_excel(writer, sheet_name='Group Sig Timings', index=False)

        for i, sig_times in enumerate(baseline_sig_timings):
            if len(sig_times) > 0:
                baseline_sig_df = pd.DataFrame(sig_times, columns=['Start', 'End'])
                baseline_sig_df.to_excel(writer, sheet_name=f'Baseline Sig Timings G{i+1}', index=False)
            else:
                pd.DataFrame(columns=['Start', 'End']).to_excel(writer, sheet_name=f'Baseline Sig Timings G{i+1}', index=False)

# GUI for file upload and selecting baseline and comparison periods
def open_file():
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        global groups, timings
        groups, timings = read_excel_data(file_path)
        messagebox.showinfo("File Loaded", "Excel file loaded successfully.")
    else:
        messagebox.showerror("Error", "No file selected.")

def save_file():
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return save_path

def analyze_and_save():
    if groups is None or timings is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    baseline_start = askinteger("Input", "Enter the start of baseline period (row index):", minvalue=0, maxvalue=len(timings)-1)
    baseline_end = askinteger("Input", "Enter the end of baseline period (row index):", minvalue=0, maxvalue=len(timings)-1)
    comparison_start = askinteger("Input", "Enter the start of comparison period (row index):", minvalue=0, maxvalue=len(timings)-1)
    comparison_end = askinteger("Input", "Enter the end of comparison period (row index):", minvalue=0, maxvalue=len(timings)-1)

    baseline_period = slice(baseline_start, baseline_end + 1)
    comparison_period = slice(comparison_start, comparison_end + 1)

    results, significance_map, baseline_diff_map, group_sig_timings, baseline_sig_timings = analyze_photometry_data(groups, baseline_period, comparison_period, timings)

    save_path = save_file()
    if save_path:
        save_results_to_excel(save_path, results, significance_map, group_sig_timings, baseline_sig_timings, timings)
        messagebox.showinfo("Results Saved", "Results saved successfully.")

# Main GUI setup
if __name__ == "__main__":
    global groups, timings
    groups = None
    timings = None

    root = tk.Tk()
    root.title("Photometry Data Analysis")
    root.geometry("400x200")

    load_button = tk.Button(root, text="Load Excel File", command=open_file)
    load_button.pack(pady=10)

    analyze_button = tk.Button(root, text="Analyze and Save Results", command=analyze_and_save)
    analyze_button.pack(pady=10)

    root.mainloop()
