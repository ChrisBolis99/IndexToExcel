import os
from openpyxl import Workbook

def create_index(folder_path, output_excel):
    wb = Workbook()
    ws = wb.active
    ws.title = "Index"

    ws.column_dimensions['A'].width = 300

    def add_to_sheet(level, name, item_type):
        indent = "----" * level
        ws.append([f"{indent}{name}", item_type])


    for root, dirs, files in os.walk(folder_path):
        depth = root.replace(folder_path, "").count(os.sep)

        if depth == 0 or os.path.basename(root):
            folder_name = os.path.basename(root) if depth > 0 else os.path.basename(folder_path)
            add_to_sheet(depth, folder_name, "Folder")

        for file_name in files:
            add_to_sheet(depth + 1, file_name, "File")

    wb.save(output_excel)
    print(f"File saved at {output_excel}")


folder_to_scan = r"C:\Users\Christopher Bolis\Documents\Unreal Projects\SpaceSurvival"
output_file = "index.xlsx"

create_index(folder_to_scan, output_file)
