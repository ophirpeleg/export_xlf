import os
import openpyxl
from xml.etree import ElementTree as ET
from xml.dom import minidom
from tkinter import filedialog
import tkinter as tk
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
                "target-language": sheet_name.lower().replace("_", "-"),
                "translation-type": "metadata",
                "datatype": "xml"
            }
            file_element = ET.SubElement(root, "file", **file_attributes)
            body_element = ET.SubElement(file_element, "body")

            for row in sheet.iter_rows(min_row=2, values_only=True):
                trans_unit_element = ET.SubElement(body_element, "trans-unit", id=str(row[0]), maxwidth=str(row[1]), size_unit=str(row[2]))
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
                initialfile=f"{sheet_name.lower().replace('_', '-')}_output.xlf"
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

if __name__ == "__main__":
    try:
        root = tk.Tk()
        root.withdraw()

        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if excel_file_path:
            logging.info("Excel file selected.")
            print("Excel file selected.")
            excel_to_xliff(excel_file_path)
        else:
            logging.info("No Excel file selected. Exiting.")
            print("No Excel file selected. Exiting.")
    except Exception as e:
        logging.error(f"An error occurred in the main execution: {e}")
        print(f"An error occurred in the main execution: {e}")
