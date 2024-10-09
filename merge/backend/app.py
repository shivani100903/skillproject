from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    if 'files' not in request.files:
        return "No file part", 400

    files = request.files.getlist('files')
    dataframes = []

    for file in files:
        if file.filename.endswith('.xlsx'):
            try:
                # Skip the first few rows for the header (adjust as needed)
                df = pd.read_excel(file, skiprows=5)  # assuming the main column headers start on row 6
                df.columns = df.columns.str.strip()
                dataframes.append(df)
            except Exception as e:
                return f"Error reading {file.filename}: {e}", 400
        else:
            return "Only Excel files are supported.", 400

    if len(dataframes) != 2:
        return "Please upload exactly two Excel files.", 400

    df1, df2 = dataframes
    common_columns = set(df1.columns) & set(df2.columns)

    if not common_columns:
        return "The two files do not have any common columns.", 400

    # Merge the DataFrames on common columns
    merged_df = pd.merge(df1, df2, on=list(common_columns), how='inner')

    if merged_df.empty:
        return "The merged file is empty.", 400

    # Define your custom multi-row header
    header_rows = [
        ["SHRI G.S. INSTITUTE OF TECHNOLOGY AND SCIENCE"],  # Empty first row
        ["SESSION : Jan-June 2023; Semester 'B'"],
        ["III YEAR ATTENDANCE SHEET SECTION B (till 31/07/2024)"],
        ["Note: Please fill the attendance till 31/07/2024. Please add the name of any student whose name is not present in the sheet and who is attending the lectures at the bottom of this sheet."],
        ["", "Total Classes", "Total Present", "Percentage", "Total Labs", "Total Present", "Percentage"]
    ]

    # Prepare to write the merged data with headers
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write the merged DataFrame to the Excel writer without headers
        merged_df.to_excel(writer, sheet_name='Merged Data', startrow=5, index=False, header=True)
        
        workbook = writer.book
        worksheet = writer.sheets['Merged Data']
        
        # Write each custom header row
        for row_num, header_row in enumerate(header_rows):
            worksheet.write_row(row_num, 0, header_row + [""] * (len(merged_df.columns) - len(header_row)))

        # Adjust column widths (optional)
        for idx, col in enumerate(merged_df.columns):
            max_len = max(merged_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)

    output.seek(0)

    return send_file(output, as_attachment=True, download_name="merged_output.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)
