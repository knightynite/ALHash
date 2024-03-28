import zipfile
import os
import xml.etree.ElementTree as ET

def list_excel_files(directory):
    """List all .xlsx files in the specified directory."""
    return [file for file in os.listdir(directory) if file.endswith('.xlsx')]

def user_select_files(files):
    """Prompt user to select files from the list."""
    print("Select the Excel file(s) to process by entering their numbers separated by commas (e.g., 1,3):")
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")
    selected_indices = input("Enter selection: ").split(',')
    selected_files = [files[int(i)-1] for i in selected_indices if i.isdigit() and 1 <= int(i) <= len(files)]
    return selected_files

def convert_format(input_str):
    parts = input_str.split('$')
    hashValue = parts[4]
    saltValue = parts[6]
    spinCount = parts[8]
    output_str = f"$office$2016$0${spinCount}${saltValue}${hashValue}"
    return output_str

def extract_sheet_protection_data_for_all_sheets(excel_file_path):
    excel_base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    output_directory = os.path.join(os.getcwd(), excel_base_name + "_SheetProtectionData")

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    try:
        with zipfile.ZipFile(excel_file_path, 'r') as z:
            sheet_files = [f for f in z.namelist() if f.startswith('xl/worksheets/')]

            for sheet_file in sheet_files:
                with z.open(sheet_file) as f:
                    xml_content = f.read()

                    tree = ET.fromstring(xml_content)
                    namespace = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    sheet_protection = tree.find('.//main:sheetProtection', namespace)

                    output_file_name = os.path.basename(sheet_file).split('.')[0] + '_sheetProtectionData.txt'
                    output_file_path = os.path.join(output_directory, output_file_name)

                    if sheet_protection is not None:
                        # Format attribute-value pairs and prepend a '$'
                        protection_data_str = "$" + "$".join([f'{attr}${value}' for attr, value in sheet_protection.attrib.items()])
                        converted_data_str = convert_format(protection_data_str)
                    else:
                        converted_data_str = "No sheetProtection data found."

                    with open(output_file_path, 'w', encoding='utf-8') as out_file:
                        out_file.write(converted_data_str)
    except zipfile.BadZipFile:
        print("Error: The file is not a valid zip file or is corrupted.")
    except FileNotFoundError:
        print("Error: The file does not exist at the specified path.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    directory = os.getcwd()  # Use the current directory, or modify as needed
    excel_files = list_excel_files(directory)
    
    if not excel_files:
        print("No Excel files found in the directory.")
    else:
        selected_files = user_select_files(excel_files)
        for file_path in selected_files:
            full_path = os.path.join(directory, file_path)
            extract_sheet_protection_data_for_all_sheets(full_path)
