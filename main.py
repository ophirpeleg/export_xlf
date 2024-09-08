import os
import openpyxl
from xml.etree import ElementTree as ET
from xml.dom import minidom
from tkinter import filedialog, messagebox
import tkinter as tk
import logging
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Function to style the Excel sheet
def style_excel_sheet(ws):
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color="D3D3D3"),
                         right=Side(style='thin', color="D3D3D3"),
                         top=Side(style='thin', color="D3D3D3"),
                         bottom=Side(style='thin', color="D3D3D3"))

    for col in ws.columns:
        for cell in col:
            cell.border = thin_border
            if cell.row == 1:  # Apply header style
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")

    ws.sheet_view.showGridLines = False



def create_output_directory(output_path):
    try:
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        logging.info("Output directory created.")
    except Exception as e:
        logging.error(f"Error creating output directory: {e}")
        print(f"Error creating output directory: {e}")
        raise


def excel_to_xliff(excel_file):
    try:
        wb = openpyxl.load_workbook(excel_file)
        logging.info("Excel workbook loaded successfully.")

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]

            root = ET.Element("xliff", version="1.2")
            file_attributes = {
                "original": "Salesforce",
                "source-language": "en_US",
                "target-language": sheet_name,
                "translation-type": "metadata",
                "datatype": "xml"
            }
            file_element = ET.SubElement(root, "file", **file_attributes)
            body_element = ET.SubElement(file_element, "body")

            seen_ids = set()  # To keep track of seen Ids and avoid duplicates
            for row in sheet.iter_rows(min_row=2, values_only=True):
                id_value = str(row[0])  # Assuming the ID is in the first column
                if id_value not in seen_ids:  # Check if the ID is already processed
                    seen_ids.add(id_value)  # Mark this ID as seen
                    trans_unit_element = ET.SubElement(body_element, "trans-unit", id=id_value, maxwidth=str(row[1]), size_unit=str(row[2]))
                    source_element = ET.SubElement(trans_unit_element, "source")
                    source_element.text = f"{str(row[3])}"
                    target_element = ET.SubElement(trans_unit_element, "target")
                    target_element.text = f"{str(row[4])}"

                    if len(row) > 5 and row[5]:
                        note_element = ET.SubElement(trans_unit_element, "note")
                        note_element.text = f"{str(row[5])}"

            xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>\n'
            xml_string = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
            xml_string = xml_string.split('\n', 1)[1] if xml_string.startswith('<?xml') else xml_string

            full_xml = xml_declaration + xml_string

            output_file_path = filedialog.asksaveasfilename(
                title="Save File As",
                defaultextension=".xlf",
                filetypes=[("XLIFF files", "*.xlf"), ("All Files", "*.*")],
                initialfile=f"{sheet_name}_output.xlf"
            )

            if output_file_path:
                with open(output_file_path, "w", encoding="utf-8") as f:
                    f.write(full_xml)
                logging.info(f"File saved successfully at {output_file_path}")
                print(f"File saved successfully at {output_file_path}")
            else:
                logging.info("File save operation was cancelled.")
                print("File save operation was cancelled.")

        logging.info("Conversion complete.")
        print("Conversion complete.")
    except Exception as e:
        logging.error(f"An error occurred during the conversion: {e}")
        print(f"An error occurred during the conversion: {e}")
        raise


# Function to convert XLIFF to Excel
def xliff_to_excel(xliff_file):
    try:
        tree = ET.parse(xliff_file)
        root = tree.getroot()

        # Extract target-language value
        target_language = root.find(".//file").attrib.get("target-language", "translations")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = target_language  # Set the worksheet title to the target-language

        # Create headers
        ws.append(["ID", "Max Width", "Size Unit", "Source", "Target", "Note"])

        # Add data to the sheet
        for file_element in root.findall("file"):
            for trans_unit in file_element.find("body").findall("trans-unit"):
                id_value = trans_unit.get("id", "")
                max_width = trans_unit.get("maxwidth", "")
                size_unit = trans_unit.get("size-unit", "")
                source_text = trans_unit.find("source").text if trans_unit.find("source") is not None else ""
                target_text = trans_unit.find("target").text if trans_unit.find("target") is not None else ""
                note_text = trans_unit.find("note").text if trans_unit.find("note") is not None else ""

                ws.append([id_value, max_width, size_unit, source_text, target_text, note_text])

        # Apply styling to the worksheet
        style_excel_sheet(ws)

        # Set default output file name using target-language
        default_output_filename = f"Excel to xlf {target_language}.xlsx"

        # Save the file
        output_file_path = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")],
            initialfile=default_output_filename
        )

        if output_file_path:
            wb.save(output_file_path)
            logging.info(f"XLIFF converted to Excel and saved at {output_file_path}")
            print(f"XLIFF converted to Excel and saved at {output_file_path}")
        else:
            logging.info("File save operation was cancelled.")
            print("File save operation was cancelled.")

    except Exception as e:
        logging.error(f"An error occurred during the XLIFF to Excel conversion: {e}")
        print(f"An error occurred during the XLIFF to Excel conversion: {e}")
        raise

def select_excel_to_xliff():
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls"),]
    )
    if excel_file_path:
        logging.info("Excel file selected.")
        print("Excel file selected.")
        excel_to_xliff(excel_file_path)
    else:
        logging.info("No Excel file selected. Exiting.")
        print("No Excel file selected. Exiting.")

def select_xliff_to_excel():
    xliff_file_path = filedialog.askopenfilename(
        title="Select XLIFF File",
        filetypes=[("XLIFF files", "*.xlf"), ("All Files", "*.*")]
    )
    if xliff_file_path:
        logging.info("XLIFF file selected.")
        print("XLIFF file selected.")
        xliff_to_excel(xliff_file_path)
    else:
        logging.info("No XLIFF file selected. Exiting.")
        print("No XLIFF file selected. Exiting.")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel to XLIFF Converter")
    root.geometry("300x150")

    btn_excel_to_xliff = tk.Button(root, text="Excel to XLIFF", command=select_excel_to_xliff, width=20)
    btn_excel_to_xliff.pack(pady=10)

    btn_xliff_to_excel = tk.Button(root, text="XLIFF to Excel", command=select_xliff_to_excel, width=20)
    btn_xliff_to_excel.pack(pady=10)

    root.mainloop()
