import os
import comtypes.client
import argparse
import logging
from logging.handlers import RotatingFileHandler

# Set up logging
logger = logging.getLogger("")
logger.setLevel(logging.INFO)
# Output logs to a file
file_handler = RotatingFileHandler(
    "doc_convert.log", maxBytes=1024 * 1024 * 100, backupCount=10
)
file_handler.setLevel(logging.INFO)
# Also output logs to console
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
# Set logging format
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
# Add the handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)


# Define function to convert files
def convert_to_pdf(input_file_path, output_file_path, file_type):
    try:
        if file_type == "word":
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(input_file_path)
            doc.SaveAs(output_file_path, FileFormat=17)  # 17 corresponds to PDF in Word
            doc.Close()
            word.Quit()
        elif file_type == "excel":
            excel = comtypes.client.CreateObject("Excel.Application")
            excel.Visible = False
            doc = excel.Workbooks.Open(input_file_path)
            doc.SaveAs(
                output_file_path, FileFormat=57
            )  # 57 corresponds to PDF in Excel
            doc.Close()
            excel.Quit()
        elif file_type == "powerpoint":
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            doc = powerpoint.Presentations.Open(input_file_path)
            doc.SaveAs(
                output_file_path, FileFormat=32
            )  # 32 corresponds to PDF in PowerPoint
            doc.Close()
            powerpoint.Quit()
        logger.info(f"Successfully converted {input_file_path} to {output_file_path}")
    except Exception as e:
        logger.error(f"Failed to convert {input_file_path} with error: {e}")


def main(input_dir, output_dir):
    # Traverse through each file in the input directory
    for file in os.listdir(input_dir):
        input_file_path = os.path.join(input_dir, file)
        output_file_path = os.path.join(output_dir, file)

        # Convert based on file extension
        if file.endswith(".doc") or file.endswith(".docx"):
            output_file_path = output_file_path.replace(".doc", ".pdf").replace(
                ".docx", ".pdf"
            )
            convert_to_pdf(input_file_path, output_file_path, "word")
        elif file.endswith(".xls") or file.endswith(".xlsx"):
            output_file_path = output_file_path.replace(".xls", ".pdf").replace(
                ".xlsx", ".pdf"
            )
            convert_to_pdf(input_file_path, output_file_path, "excel")
        elif file.endswith(".ppt") or file.endswith(".pptx"):
            output_file_path = output_file_path.replace(".ppt", ".pdf").replace(
                ".pptx", ".pdf"
            )
            convert_to_pdf(input_file_path, output_file_path, "powerpoint")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert Word, Excel, and PowerPoint documents to PDF."
    )
    parser.add_argument(
        "-i",
        "--input_dir",
        type=str,
        required=True,
        help="Input directory containing the documents",
    )
    parser.add_argument(
        "-o",
        "--output_dir",
        type=str,
        required=True,
        help="Output directory to save the PDFs",
    )
    args = parser.parse_args()
    main(args.input_dir, args.output_dir)
