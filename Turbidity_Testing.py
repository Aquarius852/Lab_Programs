from sklearn.preprocessing import MinMaxScaler
import pandas as pd
from tkinter.filedialog import askopenfilename
import os
from openpyxl.chart import LineChart, Reference, Series
from getch import pause

print("Please select the CSV file to open:")

csv_file_path = askopenfilename(
    title="Select CSV file",
    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
)

if not csv_file_path:
    print("No file selected. Exiting...")
    exit()

title = os.path.splitext(os.path.basename(csv_file_path))[0]
excel_file_path = os.path.join(os.path.dirname(csv_file_path), f"Corrected {title}.xlsx")

try:
    turbidity_test = pd.read_csv(csv_file_path)
except UnicodeDecodeError:
    print("Couldn't load as UTF-8 Encoded CSV")
    turbidity_test = pd.read_csv(csv_file_path, encoding="ANSI")

data = []
replicates = int(input("How many replicates?"))
samples = int(input("How many samples?"))
first_column = 2
last_column = replicates + 2
time_column = turbidity_test.iloc[:, 0]

for i in range(samples):
    sample = turbidity_test.iloc[:, first_column:last_column]
    scaler = MinMaxScaler()
    scaled_sample = pd.DataFrame(scaler.fit_transform(sample), columns=sample.columns)
    scaled_sample["Average"] = scaled_sample.mean(axis=1)
    scaled_sample["Final Correction"] = scaler.fit_transform(scaled_sample["Average"].values.reshape(-1, 1))
    data.append(scaled_sample)
    first_column += replicates
    last_column += replicates

with pd.ExcelWriter(excel_file_path) as writer:
    for idx, DataFrame in enumerate(data):
        DataFrame.insert(0, "Time", time_column)
        sheet_name = f'Sample_{idx + 1}'  # Naming each sheet
        DataFrame.to_excel(writer, sheet_name=sheet_name, index=False)

with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
    workbook = writer.book
    worksheet_results = workbook.create_sheet('Results')

    chart = LineChart()
    chart.legend = None
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1

    for idx in range(samples):
        sheet_name = f'Sample_{idx + 1}'
        data_ref = Reference(workbook[sheet_name], min_col=replicates+3, min_row=2, max_row=len(data[idx]) + 1)
        time_ref = Reference(workbook[sheet_name], min_col=1, min_row=2, max_row=len(data[idx]) + 1)

        series = Series(data_ref, title=sheet_name)
        chart.set_categories(time_ref)
        chart.series.append(series)

    worksheet_results.add_chart(chart, "A1")

    workbook.save(excel_file_path)

print(f'Corrected {title} added to {excel_file_path}.')
pause('Press any key to exit...')
