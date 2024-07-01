import yaml
import re
import openpyxl
import os
import argparse

def read_yaml(file_path):
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)
    return data

def extract_and_pair_columns(data):
    extracted_data = []
    for index in range(0, len(data), 2):
        match_mei = re.search(r'mei\s+\d+\s+\d+\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)', data[index])
        match_shi = re.search(r'shi\s+\d+\s+\d+\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)', data[index + 1])

        if match_mei and match_shi:
            row = [f'epoch_{index // 2 + 1}']
            row.extend([float(match_mei.group(2)), float(match_shi.group(2))])
            row.extend([float(match_mei.group(1)), float(match_shi.group(1))])
            row.extend([float(match_mei.group(3)), float(match_shi.group(3))])
            row.extend([float(match_mei.group(4)), float(match_shi.group(4))])
            extracted_data.append(row)
        else:
            print(f"No matching pattern found in {data[index]} or {data[index + 1]}")
    return extracted_data

def save_to_excel(data, headers, excel_path, overwrite=False):
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    if not overwrite:
        base, ext = os.path.splitext(excel_path)
        i = 1
        new_excel_path = excel_path
        while os.path.exists(new_excel_path):
            new_excel_path = f"{base}_{i}{ext}"
            i += 1
        excel_path = new_excel_path

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Extracted Data'

    sheet.append(headers)

    for row in data:
        sheet.append(row)

    workbook.save(excel_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract data from YAML and save to Excel")
    parser.add_argument('--folder', type=str, default='T_MIX_', help='Path to the parent folder')
    parser.add_argument('--yaml_name', type=str, default='val_get_inf.yaml', help='YAML file name to look for in each subfolder')
    parser.add_argument('--save_folder', type=str, default='excel', help='Folder to save the Excel files')
    parser.add_argument('--overwrite', action='store_true', help='Whether to overwrite the existing file')

    args = parser.parse_args()

    parent_folder = args.folder
    yaml_name = args.yaml_name
    save_folder = args.save_folder
    overwrite = args.overwrite

    headers = ['epoch_name', 'mei_Recall', 'shi_Recall', 'mei_Precision', 'shi_Precision', 'mei_mAP50', 'shi_mAP50', 'mei_mAP50-95', 'shi_mAP50-95']

    for subfolder in os.listdir(parent_folder):
        subfolder_path = os.path.join(parent_folder, subfolder)
        if os.path.isdir(subfolder_path):
            yaml_file_path = os.path.join(subfolder_path, yaml_name)
            if os.path.exists(yaml_file_path):
                data = read_yaml(yaml_file_path)
                extracted_data = extract_and_pair_columns(data)
                excel_file_path = os.path.join(save_folder, f"{subfolder}.xlsx")
                save_to_excel(extracted_data, headers, excel_file_path, overwrite=overwrite)
                print(f'数据已保存到: {excel_file_path}')
            else:
                print(f"YAML 文件 {yaml_name} 在 {subfolder_path} 中不存在")