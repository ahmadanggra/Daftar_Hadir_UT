import csv
import os
import win32com.client

# === Configuration ===
csv_file = os.getcwd() + "\\source.csv"
template_path = os.getcwd() + "\\Master.docm"
output_folder = os.getcwd() + "\\output"

os.makedirs(output_folder, exist_ok=True)

def process_document(template_path, output_folder, site_id, site_name, site_ref):
    # Launch Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Open the master document
    doc = word.Documents.Open(template_path)

    # Run a macro inside it (with arguments)
    word.Run("LookupFromExcel", site_id, site_name, site_ref)

    # Save the document under a new name (keeps macros!)
    new_name = "ATP UT PS-BDW OTM FSA " + site_name + ".docm"
    output_path = os.path.join(output_folder, new_name)
    doc.SaveAs2(output_path, FileFormat=13)

    # Close the new doc
    doc.Close(SaveChanges=False)
    word.Quit()
    print(f"✅ Saved new docm: {output_folder}")
    
# === MAIN LOOP ===
with open(csv_file, encoding="utf-8-sig") as f:
    reader = csv.DictReader(f, delimiter=';')
    for row in reader:
        site_id = row['site_id'].strip()
        site_name = row['site_name'].strip()
        site_ref = row['site_ref'].strip()
        process_document(template_path, output_folder, site_id, site_name, site_ref)
print("🎉 All files generated successfully!")
