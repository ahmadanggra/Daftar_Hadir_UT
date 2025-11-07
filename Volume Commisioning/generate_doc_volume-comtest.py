import os
import pandas as pd
from openpyxl import load_workbook

# === Configuration ===
csv_file = os.path.join(os.getcwd(), "source.csv")
template_path = os.path.join(os.getcwd(), "Master_Volume_Comtest.xlsx")
output_folder = os.path.join(os.getcwd(), "output")

os.makedirs(output_folder, exist_ok=True)

def process_document(template_path, output_folder, service_link, site_name, single_site ):
    # Load workbook using openpyxl
    wb = load_workbook(template_path)
    sheet = wb["BOQ"]

    # Modify intended cells
    sheet["C4"] = f": {service_link}"
    sheet["C5"] = f": {site_name}"

    # Save the document under a new name
    if single_site:
        new_name = f"{service_link}.xlsx"
    else:
        new_name = f"{service_link}_{site_name}.xlsx"
    output_path = os.path.join(output_folder, new_name)

    wb.save(output_path)
    wb.close()

    print(f"✅ Saved new Excel: {output_path}")

# === MAIN LOOP ===
df = pd.read_csv(csv_file, delimiter=";", encoding="utf-8-sig")

for _, row in df.iterrows():
    service_link = row["service_link"].strip()
    site_name = service_link.split()
    is_single = True if len(site_name) == 2 else False
    if len(site_name) == 4:
        process_document(template_path, output_folder, service_link, site_name[1], is_single)
        process_document(template_path, output_folder, service_link, site_name[3], is_single)
    else:
        process_document(template_path, output_folder, service_link, site_name[1], is_single)

print("🎉 All files generated successfully!")
