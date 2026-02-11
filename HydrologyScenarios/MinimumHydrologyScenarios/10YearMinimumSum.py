# MinimumHydrologyScenariosCode.py
#
#################

# Purpose
# This code iterates through all ensembles and traces in HydrologyScenarios.xlsx
#   by Homa Salehabadi (2023). During the iterations, the code finds the most minimum
#   consecutive years. The user chooses the window.

# Please report bugs/feedback to: Anabelle Myers A02369941@aggies.usu.edu


# Anabelle G. Myers
# August 11, 2025

# Utah State University
#A02369941@aggies.usu.edu
####################

# Interpreters
import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from pathlib import Path
from openpyxl import Workbook


year = int(input("Enter the number of consecutive years over which to calculate the minimum of each ensemble: "))

# Input
code_file = Path(__file__).parent # Locates code
MinimumHydrologyScenarios = code_file.parent # Locates parent folder
input_file = MinimumHydrologyScenarios / 'HydrologyScenarios.xlsx' # Shows where the input is relative to where the code is

# Output
output_file = f"{year}yearsMinimumSumHydrologyResults.xlsx" # Names output a unique name
output_path = code_file / 'Results' / output_file   # Shows where the output is relative to where the code is

sheet_names = pd.ExcelFile(input_file).sheet_names  # Variable for ease of access to all the sheets in HydrologyScenarios.xlsx

# Specifies which sheets not to read
excluded_sheets = ["ReadMe", "AvailableHydrologyScenarios", "ScenarioListForAnalysis", "AvailableMetrics", "MetricsForAnalysis", "Heatmap", "Hist"]

# Empty list to collect the minimum years window of trace calculated from the for loops
all_results = []

# Counter initializers and storages
count_ensembles = 0
all_traces = []

# START ENSEMBLE LOOP
for ensemble_input in sheet_names:  # Iterates through ensembles
    if ensemble_input in excluded_sheets:  # Skips over sheets listed in excluded_sheets
        continue

    count_ensembles += 1 # Increases count by one each time the ensemble for loop is iterated through

    # Reads and converts values in ensembles to numeric values
    ensemble = pd.read_excel(input_file, sheet_name=ensemble_input)  # Selects a specific sheet and converts it into a DataFrame
    ensemble = ensemble.apply(pd.to_numeric, errors='coerce')  # Coverts all values into numeric if not already and turns values that cannot be converted int 'NaN'

    # START TRACE LOOP
    # Start ISM trace loop
    if 'ISM' in ensemble_input: # Selects ensembles with 'ISM' in the title

        count_ismtraces = 0 # Count initializer for ISM traces

        # Searches for the overall minimum sum of consecutive years
        for ism_trace in ensemble.columns[1:2]:  # Iterates through the first trace in each ISM ensemble
            series = ensemble[ism_trace]  # Isolates the trace and attaches it to a variable
            doubled = pd.concat([series, series], ignore_index = True) # Duplicates and vertically stacks the ISM trace
            rollingSum = series.rolling(window=year).sum()  # Calculates sum of a rolling window
            valid = rollingSum.dropna()  # Drops NaN values to calculate only full windows

            end_idx = int(valid.idxmin())  # Finds the row of minimum rolling window and converts to an integer
            start_idx = end_idx - (year - 1)  # Finds the position of the start of the window based off of the end

            window = series.iloc[start_idx: start_idx + year].reset_index(drop=True)  # Extracts called window
            sum = round(window.sum(), 1) # Calculated average

            # Stores results into a dictionary
            ism_result = {
                'Ensemble': ensemble_input,
                'Trace': ism_trace,
                'Start Row': start_idx + 1,
                'Sum': sum
            }
            for i in range(len(window)):  # Assigns results to correct title to store into Excel columns
                ism_result[f'Year{i + 1}'] = round(window.iloc[i], 1)

            all_results.append(ism_result) # Adds ISM results to all results
            count_ismtraces += 1 # Increases count of iterations through ISM traces

        all_traces.append(count_ismtraces) # Adds ISM iterations count to the total iteration trace count
        continue # Forces to move to the next ensemble
    # End ISM trace loop

    # Start all other traces loop
    count_traces = 0 # Count initializer for iterations through the rest of the traces
    # Searches for the overall minimum sum of consecutive years
    for trace in ensemble.columns[1:]:  # Iterates through each trace in the ensemble
        series = ensemble[trace]  # Isolates one column to be iterated through
        rollingSum = series.rolling(window=year).sum()  # Calculates sum of a rolling window
        valid = rollingSum.dropna()  # Drops NaN values to calculate only full windows

        end_idx = int(valid.idxmin())  # Finds end position of minimum rolling window
        start_idx = end_idx - (year - 1)  # Finds the position of the start of the window based on the end

        window = series.iloc[start_idx: start_idx + year].reset_index(drop=True)  # Extracts called window
        sum = round(window.sum(), 1)

        # Stores results into a dictionary
        result = {
            'Ensemble': ensemble_input,
            'Trace': trace,
            'Start Row': start_idx + 1,
            'Sum': sum
            }

        for i in range(len(window)):  # Assigns results to correct title to store into Excel columns
            result[f'Year{i + 1}'] = round(window.iloc[i], 1)

        all_results.append(result)  # Adds results to the list of all results
        count_traces += 1 # Increases count of iterations through all other traces
        # End all other traces loop
        # END TRACE LOOP
    all_traces.append(count_traces) # Adds the count of iterations through all other traces to the total iteration count
# END ENSEMBLE LOOP

# Convert to DataFrame
df = pd.DataFrame(all_results) # Puts all_results into a DataFrame
df = df.sort_values(by='Sum', ascending=True).reset_index(drop=True) # Sorts the sums and outputs to Excel in this order. Resets the index, so rows are correct.

# Write results to a single-sheet Excel file with a table
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ensemble Minimums', index=False) # Adds DataFrame to a named sheet
    workbook = writer.book # Accesses the workbook
    worksheet = writer.sheets['Ensemble Minimums'] # Accesses the worksheet

    end_col = chr(65 + len(df.columns) - 1) # Calculates column letter
    end_row = len(df) + 1 # Calculate row number
    table_range = f"A1:{end_col}{end_row}" # Creates Excel range

    table = Table(displayName="MinPerEnsemble", ref=table_range) # Creates table
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True) # Creates table style
    table.tableStyleInfo = style
    worksheet.add_table(table) # adds table to the worksheet


print(f"\nResults saved to:\n{output_path}") # Displays output path

#print("\nThe code iterated through", count_ensembles, "ensembles and", sum(all_traces), "traces.") # Displays number of iterations through ensembles and traces
