import re
import shutil
import argparse
import logging
import comtypes.client
from time import sleep
from logging.handlers import RotatingFileHandler
from pathlib import Path

# Set up logging
logger = logging.getLogger("")
logger.setLevel(logging.INFO)
file_handler = RotatingFileHandler(
    "doc_convert.log", maxBytes=1024 * 1024 * 100, backupCount=10
)
file_handler.setLevel(logging.INFO)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(console_handler)


# Define function to convert files
def convert_to_pdf(input_file_path, output_file_path, file_type):
    try:
        if not output_file_path.exists():
            if file_type == "word":
                word = comtypes.client.CreateObject("Word.Application")
                doc = word.Documents.Open(str(input_file_path))
                doc.SaveAs(
                    str(output_file_path), FileFormat=17
                )  # 17 corresponds to PDF in Word
                doc.Close()
                word.Quit()
            elif file_type == "excel":
                excel = comtypes.client.CreateObject("Excel.Application")
                doc = excel.Workbooks.Open(str(input_file_path))
                doc.SaveAs(
                    str(output_file_path), FileFormat=57
                )  # 57 corresponds to PDF in Excel
                doc.Close()
                excel.Quit()
            elif file_type == "powerpoint":
                powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                powerpoint.Visible = 1
                doc = powerpoint.Presentations.Open(str(input_file_path))
                sleep(1)  # This is necessary to allow powerpoint to open the file
                doc.SaveAs(
                    str(output_file_path), FileFormat=32
                )  # 32 corresponds to PDF in PowerPoint
                doc.Close()
                powerpoint.Quit()
            logger.info(
                f"Successfully converted {input_file_path} to {output_file_path}"
            )
        else:
            logger.info(
                f"Skipping {input_file_path} as {output_file_path} already exists"
            )
    except Exception as e:
        logger.error(f"Failed to convert {input_file_path} with error: {e}")


def main(input_dir, output_dir):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    # Create temporary directory
    temp_dir = Path.cwd() / "temp"
    temp_dir.mkdir(exist_ok=True)

    for file in input_dir.iterdir():
        if file.suffix.lower() in [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"]:
            new_name = re.sub(r"\W+", "_", file.stem) + file.suffix.lower()
            new_path = temp_dir / new_name
            if not new_path.exists():
                shutil.copy(file, new_path)

    # Traverse through each file in the temporary directory
    for file in temp_dir.iterdir():
        input_file_path = temp_dir / file.name
        output_file_path = output_dir / file.name

        # Convert based on file extension
        if file.suffix.lower() in [".doc", ".docx"]:
            output_file_path = output_file_path.with_suffix(".pdf")
            convert_to_pdf(input_file_path, output_file_path, "word")
        elif file.suffix.lower() in [".xls", ".xlsx"]:
            output_file_path = output_file_path.with_suffix(".pdf")
            convert_to_pdf(input_file_path, output_file_path, "excel")
        elif file.suffix.lower() in [".ppt", ".pptx"]:
            output_file_path = output_file_path.with_suffix(".pdf")
            convert_to_pdf(input_file_path, output_file_path, "powerpoint")

    # Delete temporary directory
    shutil.rmtree(temp_dir)


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
