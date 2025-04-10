import streamlit as st
import pandas as pd
import io
import numpy as np
from datetime import datetime

# Function to strip leading zeros and convert to numeric if applicable
def strip_leading_zeros(val):
    try:
        if isinstance(val, str):
            val = val.strip().replace(',', '')
            if val.replace('.', '', 1).isdigit() or \
               (val.startswith('-') and val[1:].replace('.', '', 1).isdigit()):
                return float(val)
        return val
    except:
        return val

# Apply numeric cleaning only on likely numeric columns
def convert_possible_numeric(df):
    for col in df.columns:
        df[col] = df[col].apply(strip_leading_zeros)
        try:
            df[col] = pd.to_numeric(df[col], errors='ignore')
        except:
            pass
    return df

# Define the checklist data as a DataFrame
checklist_data = {
    "S.No": range(1, 18),
    "Checklist": [
        "Database & Warehouse is parameterized (In case of DESQL Reports)",
        "All the columns of Cognos replicated in PBI (No extra columns)",
        "All the filters of Cognos replicated in PBI",
        "Filters working as expected (single/multi select as usual)",
        "Column names matching with Cognos",
        "Currency symbols to be replicated",
        "Filters need to be aligned vertically/horizontally",
        "Report Name & Package name to be written",
        "Entire model to be refreshed before publishing to PBI service",
        "Date Last refreshed to be removed from filter/table",
        "Table's column header to be bold",
        "Table style to not have grey bars",
        "Pre-applied filters while generating validation report?",
        "Dateformat to be YYYY-MM-DD [hh:mm:ss] in refresh date as well",
        "Sorting is replicated",
        "Filter pane to be hidden before publishing to PBI service",
        "Mentioned the exception in our validation document like numbers/columns/values mismatch (if any)"
    ],
    "Status - Level1": ["" for _ in range(17)],
    "Status - Level2": ["" for _ in range(17)]
}
checklist_df = pd.DataFrame(checklist_data)

# Function to generate validation report
def generate_validation_report(cognos_df, pbi_df):
    dims = [col for col in cognos_df.columns if col in pbi_df.columns and 
            (cognos_df[col].dtype == 'object' or '_id' in col.lower() or '_key' in col.lower() or
             '_ID' in col or '_KEY' in col)]

    cognos_df[dims] = cognos_df[dims].fillna('NAN')
    pbi_df[dims] = pbi_df[dims].fillna('NAN')

    cognos_measures = [col for col in cognos_df.columns if col not in dims and np.issubdtype(cognos_df[col].dtype, np.number)]
    pbi_measures = [col for col in pbi_df.columns if col not in dims and np.issubdtype(pbi_df[col].dtype, np.number)]
    all_measures = list(set(cognos_measures) & set(pbi_measures))

    cognos_agg = cognos_df.groupby(dims)[all_measures].sum().reset_index()
    pbi_agg = pbi_df.groupby(dims)[all_measures].sum().reset_index()

    cognos_agg['unique_key'] = cognos_agg[dims].astype(str).agg('-'.join, axis=1).str.upper()
    pbi_agg['unique_key'] = pbi_agg[dims].astype(str).agg('-'.join, axis=1).str.upper()

    cognos_agg = cognos_agg[['unique_key'] + [col for col in cognos_agg.columns if col != 'unique_key']]
    pbi_agg = pbi_agg[['unique_key'] + [col for col in pbi_agg.columns if col != 'unique_key']]

    validation_report = pd.DataFrame({'unique_key': list(set(cognos_agg['unique_key']) | set(pbi_agg['unique_key']))})

    for dim in dims:
        validation_report[dim] = validation_report['unique_key'].map(dict(zip(cognos_agg['unique_key'], cognos_agg[dim])))
        validation_report[dim].fillna(validation_report['unique_key'].map(dict(zip(pbi_agg['unique_key'], pbi_agg[dim]))), inplace=True)

    validation_report['presence'] = validation_report['unique_key'].apply(
        lambda key: 'Present in Both' if key in cognos_agg['unique_key'].values and key in pbi_agg['unique_key'].values
        else ('Present in Cognos' if key in cognos_agg['unique_key'].values
              else 'Present in PBI')
    )

    for measure in all_measures:
        validation_report[f'{measure}_Cognos'] = validation_report['unique_key'].map(dict(zip(cognos_agg['unique_key'], cognos_agg[measure])))
        validation_report[f'{measure}_PBI'] = validation_report['unique_key'].map(dict(zip(pbi_agg['unique_key'], pbi_agg[measure])))

        validation_report[f'{measure}_Diff'] = validation_report[f'{measure}_PBI'].fillna(0) - validation_report[f'{measure}_Cognos'].fillna(0)

    column_order = ['unique_key'] + dims + ['presence'] + \
                   [col for measure in all_measures for col in 
                    [f'{measure}_Cognos', f'{measure}_PBI', f'{measure}_Diff']]
    validation_report = validation_report[column_order]

    return validation_report, cognos_agg, pbi_agg

# Function to create column checklist
def column_checklist(cognos_df, pbi_df):
    cognos_columns = cognos_df.columns.tolist()
    pbi_columns = pbi_df.columns.tolist()

    checklist_df = pd.DataFrame({
        'Cognos Columns': cognos_columns + [''] * (max(len(pbi_columns), len(cognos_columns)) - len(cognos_columns)),
        'PowerBI Columns': pbi_columns + [''] * (max(len(pbi_columns), len(cognos_columns)) - len(pbi_columns))
    })

    checklist_df['Match'] = checklist_df.apply(lambda row: row['Cognos Columns'] == row['PowerBI Columns'], axis=1)
    
    return checklist_df

# Function to create diff checker
def generate_diff_checker(validation_report):
    diff_columns = [col for col in validation_report.columns if col.endswith('_Diff')]

    diff_checker = pd.DataFrame({
        'Diff Column Name': diff_columns,
        'Sum of Difference': [validation_report[col].sum() for col in diff_columns]
    })

    presence_summary = {
        'Diff Column Name': 'All rows present in both',
        'Sum of Difference': 'Yes' if all(validation_report['presence'] == 'Present in Both') else 'No'
    }
    diff_checker = pd.concat([diff_checker, pd.DataFrame([presence_summary])], ignore_index=True)

    return diff_checker

# Main function
def main():
    st.title("Validation Report Generator")

    st.markdown("""
    **Important Assumptions:**
    1. Upload the Excel file with two sheets: "Cognos" and "PBI".
    2. Make sure the column names are similar in both sheets.
    3. If there are ID/Key/Code columns, make sure the ID or Key columns contains "_ID" or "_KEY" (case insensitive)
    """)

    model_name = st.text_input("Enter the model name:")
    report_name = st.text_input("Enter the report name:")

    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            cognos_df = pd.read_excel(xls, 'Cognos')
            pbi_df = pd.read_excel(xls, 'PBI')

            # Clean numeric-looking strings
            cognos_df = cognos_df.applymap(strip_leading_zeros)
            pbi_df = pbi_df.applymap(strip_leading_zeros)

            # Standardize text columns
            cognos_df = cognos_df.apply(lambda x: x.str.upper().str.strip() if x.dtype == "object" else x)
            pbi_df = pbi_df.apply(lambda x: x.str.upper().str.strip() if x.dtype == "object" else x)

            option = st.radio("Select Option", ["Data Present", "Only Column Names Present"])

            if option == "Only Column Names Present":
                column_checklist_df = column_checklist(cognos_df, pbi_df)

                st.subheader("Column Checklist Preview")
                st.dataframe(column_checklist_df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    checklist_df.to_excel(writer, sheet_name='Checklist', index=False)
                    cognos_df.to_excel(writer, sheet_name='Cognos', index=False)
                    pbi_df.to_excel(writer, sheet_name='PBI', index=False)
                    column_checklist_df.to_excel(writer, sheet_name='Column Checklist', index=False)

                output.seek(0)
                today_date = datetime.today().strftime('%Y-%m-%d')
                dynamic_filename = f"{model_name}_{report_name}_ColumnCheck_Report_{today_date}.xlsx" if model_name and report_name else f"ColumnCheck_Report_{today_date}.xlsx"

                st.download_button(
                    label="Download Column Check Excel Report",
                    data=output,
                    file_name=dynamic_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            elif option == "Data Present":
                validation_report, cognos_agg, pbi_agg = generate_validation_report(cognos_df, pbi_df)
                column_checklist_df = column_checklist(cognos_df, pbi_df)
                diff_checker_df = generate_diff_checker(validation_report)

                st.subheader("Validation Report Preview")
                st.dataframe(validation_report)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    checklist_df.to_excel(writer, sheet_name='Checklist', index=False)
                    cognos_agg.to_excel(writer, sheet_name='Cognos', index=False)
                    pbi_agg.to_excel(writer, sheet_name='PBI', index=False)
                    validation_report.to_excel(writer, sheet_name='Validation_Report', index=False)
                    column_checklist_df.to_excel(writer, sheet_name='Column Checklist', index=False)
                    diff_checker_df.to_excel(writer, sheet_name='Diff Checker', index=False)

                output.seek(0)
                today_date = datetime.today().strftime('%Y-%m-%d')
                dynamic_filename = f"{model_name}_{report_name}_ValidationReport_{today_date}.xlsx" if model_name and report_name else f"ValidationReport_{today_date}.xlsx"

                st.download_button(
                    label="Download Excel Report",
                    data=output,
                    file_name=dynamic_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
