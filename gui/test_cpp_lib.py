import excel_vba

# Define the path to your .bas VBA script and the list of Excel file paths.
vba_script = r"C:\path\to\your_script.bas"
excel_files = [
    r"C:\path\to\file1.xlsx",
    r"C:\path\to\file2.xlsx",
    # add other files as needed
]

results = excel_vba.run_vba_on_all_files(vba_script, excel_files)
for result in results:
    print("Macro returned:", result)
