import csv
import os
import win32com.client as win32

# === Configuration ===
csv_file = os.getcwd() + "\\source.csv"
template_path = os.getcwd() + "\\Master_Volume_Comtest.xlsx"
output_folder = os.getcwd() + "\\output"

os.makedirs(output_folder, exist_ok=True)

def process_document(template_path, output_folder, service_link, site_name):
    # Launch Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    # Open the master document
    wb = excel.Workbooks.Open(template_path)
    sheet = wb.Sheets("BOQ")

    # Modify intended value
    sheet.Range("C4").Value = service_link
    sheet.Range("C5").Value = site_name

    # Save the document under a new name (keeps macros!)
    new_name = site_name + "_" + site_name + ".xlsx"
    output_path = os.path.join(output_folder, new_name)
    wb.SaveAs(output_path)


    # Close the new doc
    wb.Close(SaveChanges=False)
    excel.Quit()
    print(f"✅ Saved new excel: {output_folder}")
    
# === MAIN LOOP ===
with open(csv_file, encoding="utf-8-sig") as f:
    reader = csv.DictReader(f, delimiter=';')
    for row in reader:
        service_link = row['service_link'].strip()
        process_document(template_path, output_folder, service_link, service_link)
print("🎉 All files generated successfully!")
