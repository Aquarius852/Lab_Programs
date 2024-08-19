import pandas as pd
from tkinter.filedialog import askopenfilename
from getch import pause

print("Please select your Excel file:")

excel_file_path = askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
)

if not excel_file_path:
    print("You did not select a file. Exiting program...")
    exit()

dfs = pd.read_excel(excel_file_path, sheet_name=None)
all_organoid_data = []

for sheet_name, df in dfs.items():

    label_columns = [col for col in df.columns if 'Label' in col]
    area_columns = [col for col in df.columns if 'Area' in col]

    results = []
    organoid_areas = []
    lumen_counts = []

    for label_col, area_col in zip(label_columns, area_columns):
        temp_df = df[[label_col, area_col]].copy()
        temp_df['rank'] = temp_df.groupby(label_col).cumcount(ascending=False)

        organoid_area = temp_df[temp_df['rank'] == 0][area_col].values
        organoid_areas.append(organoid_area[0])

        lumen_count = temp_df[area_col].count()
        lumen_counts.append(lumen_count)

        temp_df = temp_df[temp_df['rank'] != 0].drop(columns='rank')
        grouped_sum = temp_df.groupby(label_col)[area_col].sum().reset_index()
        grouped_sum.columns = ['Sample', 'Lumen Area']
        results.append(grouped_sum[['Sample', 'Lumen Area']])

    results_df = pd.concat(results, ignore_index=True)
    results_df['Lumen Count'] = lumen_counts
    results_df['Organoid Areas'] = organoid_areas
    results_df['Percent Lumen Area'] = results_df['Lumen Area'] / results_df['Organoid Areas']

    all_organoid_data.append(results_df)

df = pd.concat(all_organoid_data)

with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='Results', index=False)

print(f"Sheet 'Results' added to {excel_file_path}.")

pause("Press any key to exit...")

