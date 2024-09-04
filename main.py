import argparse
import logging
import os
import zipfile
import re


def remove_protection(xml_text, sheet=True):
    """
    Remove the sheetProtection element from the provided XML text using a regular expression.

    Parameters:
        xml_text (str): The XML content as a string.

    Returns:
        str: The modified XML content with the sheetProtection element removed.
    """
    # Define the regular expression pattern to match the <sheetProtection> element
    pattern = r"<workbookProtection[^>]*\s*/>"
    if sheet:
        pattern = r"<sheetProtection[^>]*\s*/>"

    # Replace the matched pattern with an empty string to remove the element
    modified_xml_text = re.sub(pattern, "", str(xml_text), flags=re.IGNORECASE)

    return modified_xml_text

def process_zip_file(zip_file_path):
    """
    Process a zip file, remove sheet protection from worksheets, and save the modified contents to a new zip archive.

    Parameters:
        zip_file_path (str): The path to the zip file to process.
    """
    try:
        # Open the original zip file for reading
        with zipfile.ZipFile(zip_file_path, "r") as zfile:

            # Get the folder and file name
            folder_path = os.path.dirname(zip_file_path)
            new_file_name = os.path.basename(zip_file_path)
            new_file_name = "modified_" + new_file_name

            # Create a new zip file for writing the modified contents
            with zipfile.ZipFile(
                os.path.join(folder_path, new_file_name), "w"
            ) as modified_zfile:

                # Loop through each file in the original zip archive
                for file_info in zfile.infolist():
                    file_name = file_info.filename

                    # Read the content of the current file
                    with zfile.open(file_name) as file:
                        content = file.read()

                    # Check if the file is in the 'xl\worksheets\' folder
                    worksheet_folder = "xl/worksheets/"
                    workbook_file = "xl/workbook.xml"
                    if file_name.startswith(worksheet_folder):
                        # Remove sheet protection from the content
                        modified_content = remove_protection(str(content.decode()))

                        # Save the modified content into the new zip archive
                        modified_zfile.writestr(file_name, modified_content)
                    elif file_name == workbook_file:
                        modified_content = remove_protection(str(content.decode()), sheet=False)

                        # Save the modified content into the new zip archive
                        modified_zfile.writestr(file_name, modified_content)

                    else:
                        # If the file is not in the 'xl\worksheets\' folder, save it as it is
                        modified_zfile.writestr(file_name, content)

    except zipfile.BadZipFile as e:
        print(f"Failed to open zip file: {e}")
    except Exception as e:
        print(f"Error occurred: {e}")


def setup_logging():
    # Set up logging configuration
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main():
    # Configure logging
    setup_logging()

    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description="Remove sheet protection from a Excel file."
    )
    parser.add_argument("excel_file", help="Path to the Excel file to process")
    args = parser.parse_args()

    excel_file_path = args.excel_file

    # Check if the specified zip file exists
    if not os.path.exists(excel_file_path):
        logging.error("The specified Excel file does not exist.")
        return

    try:
        # Process the zip file to remove sheet protection
        process_zip_file(excel_file_path)
        logging.info(
            "Sheet protection removed successfully. Modified file saved as 'modified_{}'.".format(
                os.path.basename(excel_file_path)
            )
        )
    except Exception as e:
        # Log any error that occurs during the process
        logging.error(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
