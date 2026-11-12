import argparse
import configparser
import csv
import logging
import re
import sys
import zipfile
from email.message import Message
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse

import requests
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

logger = logging.getLogger(__name__)


def load_config(config_path):
    config = configparser.ConfigParser()
    config_file = Path(config_path)

    if not config_file.is_file():
        raise FileNotFoundError(
            f"Configuration file not found: {config_file}"
        )

    config.read(config_file, encoding="utf-8")

    timeout = config.getint(
        "DOWNLOAD",
        "timeout",
        fallback=30
    )

    encoding = config.get(
        "CSV",
        "encoding",
        fallback="utf-8"
    )

    log_level = config.get(
        "LOGGING",
        "level",
        fallback="INFO"
    ).upper()

    log_file = config.get(
        "LOGGING",
        "file",
        fallback="converter.log"
    )

    if timeout <= 0:
        raise ValueError("Download timeout must be positive.")

    if not encoding.strip():
        raise ValueError("CSV encoding cannot be empty.")

    valid_log_levels = {
        "DEBUG",
        "INFO",
        "WARNING",
        "ERROR",
        "CRITICAL"
    }

    if log_level not in valid_log_levels:
        raise ValueError(f"Invalid logging level: {log_level}")

    return {
        "timeout": timeout,
        "encoding": encoding,
        "log_level": log_level,
        "log_file": log_file
    }


def configure_logging(config):
    log_level = getattr(logging, config["log_level"])

    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(
                config["log_file"],
                encoding="utf-8"
            ),
            logging.StreamHandler()
        ],
        force=True
    )


def is_google_sheet_url(input_source):
    parsed_url = urlparse(input_source)

    return (
        parsed_url.scheme in {"http", "https"}
        and parsed_url.netloc.lower() == "docs.google.com"
        and "/spreadsheets/d/" in parsed_url.path
    )


def extract_spreadsheet_id(url):
    match = re.search(
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
        url
    )

    if not match:
        raise ValueError("Invalid Google Sheets URL.")

    return match.group(1)


def download_google_sheet(url, timeout):
    spreadsheet_id = extract_spreadsheet_id(url)

    download_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/export?format=xlsx"
    )

    response = requests.get(
        download_url,
        timeout=timeout
    )

    response.raise_for_status()

    workbook_content = response.content

    if not workbook_content.startswith(b"PK"):
        raise ValueError(
            "Google Sheets did not return a valid Excel file. "
            "Check whether the spreadsheet is publicly accessible."
        )

    content_disposition = response.headers.get(
        "Content-Disposition",
        ""
    )

    message = Message()
    message["Content-Disposition"] = content_disposition

    downloaded_filename = message.get_filename()

    if downloaded_filename:
        workbook_name = Path(downloaded_filename).stem
    else:
        workbook_name = spreadsheet_id

    return workbook_content, workbook_name


def read_workbook(input_source, timeout):
    parsed_url = urlparse(input_source)

    if parsed_url.scheme in {"http", "https"}:
        if not is_google_sheet_url(input_source):
            raise ValueError(
                "Only Google Sheets URLs are supported."
            )

        workbook_content, workbook_name = download_google_sheet(
            input_source,
            timeout
        )

        workbook_source = BytesIO(workbook_content)

    else:
        input_path = Path(input_source)

        if not input_path.suffix:
            input_path = input_path.with_suffix(".xlsx")

        if input_path.suffix.lower() != ".xlsx":
            raise ValueError(
                "Input file must have an .xlsx extension."
            )

        if not input_path.is_file():
            raise FileNotFoundError(
                f"Excel file not found: {input_path}"
            )

        workbook_name = input_path.stem
        workbook_source = input_path

    workbook = load_workbook(
        workbook_source,
        read_only=True,
        data_only=True
    )

    try:
        if not workbook.sheetnames:
            raise ValueError(
                "The workbook does not contain any worksheets."
            )

        worksheets = {}

        for worksheet in workbook.worksheets:
            rows = list(
                worksheet.iter_rows(values_only=True)
            )

            worksheets[worksheet.title] = rows

    finally:
        workbook.close()

    return worksheets, workbook_name


def sanitize_filename(filename):
    safe_name = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        str(filename)
    )

    safe_name = safe_name.strip(" .")
    safe_name = safe_name[:150]

    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10))
    }

    if safe_name.upper() in reserved_names:
        safe_name = f"_{safe_name}"

    return safe_name or "Sheet"


def get_unique_filename(sheet_name, used_names):
    base_name = sanitize_filename(sheet_name)
    filename = base_name
    counter = 2

    while filename.casefold() in used_names:
        filename = f"{base_name}_{counter}"
        counter += 1

    used_names.add(filename.casefold())

    return filename


def normalize_row(row):
    normalized_values = []

    for value in row:
        if isinstance(value, float) and value.is_integer():
            value = int(value)

        normalized_values.append(value)

    return normalized_values


def convert_workbook(worksheets, workbook_name, encoding):
    safe_workbook_name = sanitize_filename(workbook_name)

    output_directory = Path.cwd() / safe_workbook_name

    output_directory.mkdir(
        parents=True,
        exist_ok=True
    )

    used_names = set()
    converted_files = []

    for sheet_name, rows in worksheets.items():
        csv_filename = get_unique_filename(
            sheet_name,
            used_names
        )

        csv_path = output_directory / f"{csv_filename}.csv"

        with csv_path.open(
            "w",
            newline="",
            encoding=encoding
        ) as csv_file:

            writer = csv.writer(csv_file)
            writer.writerows(
                normalize_row(row)
                for row in rows
            )

        converted_files.append(csv_path)

        logger.info(
            "Converted worksheet '%s' to '%s'",
            sheet_name,
            csv_path
        )

    return output_directory, converted_files


def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Convert every worksheet in an Excel workbook "
            "or Google Sheet into separate CSV files."
        )
    )

    parser.add_argument(
        "input",
        help="Path to an .xlsx file or a Google Sheets URL."
    )

    parser.add_argument(
        "--config",
        default="config.ini",
        help="Path to the application configuration file."
    )

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    try:
        config = load_config(args.config)
        configure_logging(config)

        logger.info("Starting Excel to CSV conversion.")

        worksheets, workbook_name = read_workbook(
            args.input,
            config["timeout"]
        )

        output_directory, converted_files = convert_workbook(
            worksheets,
            workbook_name,
            config["encoding"]
        )

        logger.info(
            "Conversion completed. %d CSV file(s) created.",
            len(converted_files)
        )

        print("\nConversion completed successfully.")
        print(f"Output directory: {output_directory}")

        for csv_file in converted_files:
            print(f"Created: {csv_file.name}")

        return 0

    except FileNotFoundError as error:
        logger.exception("File not found")
        print(f"Error: {error}")
        return 1

    except requests.RequestException as error:
        logger.exception("Google Sheets download failed")
        print(f"Download error: {error}")
        return 1

    except InvalidFileException as error:
        logger.exception("Invalid Excel workbook")
        print(f"Invalid Excel file: {error}")
        return 1

    except zipfile.BadZipFile as error:
        logger.exception("Corrupted Excel workbook")
        print(f"Corrupted Excel file: {error}")
        return 1

    except configparser.Error as error:
        logger.exception("Configuration error")
        print(f"Configuration error: {error}")
        return 1

    except UnicodeError as error:
        logger.exception("Encoding error")
        print(f"Encoding error: {error}")
        return 1

    except (csv.Error, ValueError, OSError) as error:
        logger.exception("Conversion failed")
        print(f"Conversion error: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
