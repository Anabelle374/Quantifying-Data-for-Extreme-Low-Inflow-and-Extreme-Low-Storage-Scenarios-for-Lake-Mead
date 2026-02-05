# This code answers the question: What is the change in annual flow over all Reclamation's Post-2026 ensembles and traces.

# Tools
import pandas as pd
from pathlib import Path
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import sys


# ==== Input and Output Setup ===
# Input
code_file = Path(__file__).parent # Locates code
input_file = '10yearsMinimumHydrologyResults.xlsx' # Identifies input
# Output
output_file = "Differences.xlsx" # Identifies output
output_path = code_file / 'Results' / output_file # Locates output path

# === Input Data Preparation to be Used for Calculations ===
# Read input file
ensemble = pd.read_excel(input_file, header=0) # Reads first row as column names

# Identifies the Year, Start Row, and Average columns and changes them to iterable integers
year_cols = [c for c in ensemble.columns if c.startswith("Year")]   # Identifies columns that start with 'Year'
ensemble.loc[:, year_cols] = ensemble.loc[:, year_cols].apply(pd.to_numeric, errors='coerce')   # Converts Year columns to numerics and forces errors to 'NaN'
ensemble['Start Row'] = pd.to_numeric(ensemble['Start Row'], errors='coerce').astype('Int64')   # Converts start row to numerics
ensemble['Average'] = pd.to_numeric(ensemble['Average'], errors='coerce')   # Converts Average column to numerics

# Filter sequences by Average to keep sequences where average is <= 7.5
filtered = ensemble[ensemble['Average'] <= 7.5]

# Creates narrow form
narrow_flow = filtered.melt(
    id_vars=['Ensemble', 'Trace', 'Start Row', 'Average'], # Keeps these columns the same
    value_vars=year_cols, # Turns the Year rows into a column
    var_name='YearCol', # Names column above
    value_name='Flow' # Names new column for flow values
)

# Creating Indices
narrow_flow['YearOffset'] = narrow_flow['YearCol'].str.replace("Year", "", regex=False).astype(int) # Indexes Years into numerics
narrow_flow['Row'] = (narrow_flow['Start Row'].astype(int) + (narrow_flow['YearOffset'] - 1)).astype(int) # Calculates row index in wide data

# Keeps sequences in order then sorts the sequences by average, then ensemble, trace, start row, then row
narrow_flow = narrow_flow.sort_values(by=['Average', 'Ensemble', 'Trace', 'Start Row', 'Row']).reset_index(drop=True)

# === Calculations ===
# Calculates annual difference by current - previous year in each 10 year sequence
narrow_flow['Difference'] = (
    narrow_flow
    .groupby(['Ensemble', 'Trace', 'Start Row'])['Flow'] # Only calculates annual difference in same 10 year sequence
    .transform(lambda s: s.diff().round(3)) # First value in sequence is blank
)

# Creates final dataframe with columns identified below
result = narrow_flow[['Ensemble', 'Trace', 'Row', 'Difference']].reset_index(drop=True)

# Count
num_sequences = narrow_flow.groupby(['Ensemble', 'Trace', 'Start Row']).ngroups
num_ensembles = result['Ensemble'].nunique()
num_traces = result['Trace'].nunique()
print(f"\nSequences of 10 consecutive years in results: {num_sequences}")
print(f"Ensembles in results: {num_ensembles}")
print(f"Traces in results: {num_traces}")

# Write to Excel with a table
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    result.to_excel(writer, sheet_name='Differences', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Differences']

    # Table range
    num_cols = len(result.columns)
    end_col = get_column_letter(num_cols)
    end_row = len(result) + 1
    table_range = f"A1:{end_col}{end_row}"

    # Table style
    table = Table(displayName="DifferencesByRow", ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    worksheet.add_table(table)

print(f"\nResults saved to:\n {output_path}")


# Creating Histogram
hist_df = pd.read_excel(output_path) # Reads file created above
hist_data = -(hist_df['Difference']).dropna() # Mirrors data and drops NaN's
fig, ax = plt.subplots() # Prepares the plot below

edges = np.arange(-0.5, 8, 1) # Defines bin boundaries
ax.hist(hist_data, bins=edges, density=True, edgecolor='black') # Creates histogram

ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1, decimals=0)) # Y-axis is percentage

# Further definition of bin boundaries
ax.set_xticks(np.arange(0, 8))
ax.set_xlim(left=-0.5, right=7.5)

# Axis labels
ax.set_xlabel("Annual decrease in flow\n (million acre-feet per year)", fontsize=24)
ax.set_ylabel("Percentage", fontsize=26)
ax.tick_params(axis='both', labelsize=20)
fig.tight_layout()

# Saves the histogram
hist_png_path = output_path.parent / 'AnnualDecreaseInFlow.png'
plt.savefig(hist_png_path, dpi=300, bbox_inches='tight')

print(f"\nHistogram image saved to:\n {hist_png_path}")

plt.show()



print("\nComplete")
sys.exit()





