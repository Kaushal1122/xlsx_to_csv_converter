import configparser
import csv
import logging
import zipfile
from io import BytesIO
from unittest.mock import Mock

import pytest
import requests
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

import converter
from converter import (
    build_parser,
    configure_logging,
    convert_workbook,
    download_google_sheet,
    extract_spreadsheet_id,
    get_unique_filename,
    is_google_sheet_url,
    load_config,
    normalize_row,
    read_workbook,
    sanitize_filename,
)


def create_test_workbook(file_path):
    workbook = Workbook()

    first_sheet = workbook.active
    first_sheet.title = "Office Supply Sales"

    first_sheet.append(["Product", "Quantity", "Price"])
    first_sheet.append(["Pen", 10, 5])
    first_sheet.append(["Notebook", 5, 40])

    second_sheet = workbook.create_sheet("Food Sales")

    second_sheet.append(["Food", "Quantity", "Price"])
    second_sheet.append(["Rice", 20, 60])
    second_sheet.append(["Wheat", 15, 45])

    workbook.save(file_path)
    workbook.close()


def create_workbook_bytes():
    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "Online Sales"

    worksheet.append(["Product", "Quantity"])
    worksheet.append(["Pen", 10])

    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()

    return buffer.getvalue()


def test_load_config(tmp_path):
    config_file = tmp_path / "config.ini"

    config_file.write_text(
        "[DOWNLOAD]\n"
        "timeout = 45\n"
        "\n"
        "[CSV]\n"
        "encoding = utf-8\n"
        "\n"
        "[LOGGING]\n"
        "level = DEBUG\n"
        "file = test.log\n",
        encoding="utf-8"
    )

    config = load_config(config_file)

    assert config["timeout"] == 45
    assert config["encoding"] == "utf-8"
    assert config["log_level"] == "DEBUG"
    assert config["log_file"] == "test.log"


def test_load_config_defaults(tmp_path):
    config_file = tmp_path / "config.ini"
    config_file.write_text("", encoding="utf-8")

    config = load_config(config_file)

    assert config["timeout"] == 30
    assert config["encoding"] == "utf-8"
    assert config["log_level"] == "INFO"
    assert config["log_file"] == "converter.log"


def test_load_config_missing_file(tmp_path):
    config_file = tmp_path / "missing.ini"

    with pytest.raises(FileNotFoundError):
        load_config(config_file)


def test_load_config_invalid_timeout(tmp_path):
    config_file = tmp_path / "config.ini"

    config_file.write_text(
        "[DOWNLOAD]\n"
        "timeout = 0\n",
        encoding="utf-8"
    )

    with pytest.raises(ValueError, match="timeout must be positive"):
        load_config(config_file)


def test_load_config_empty_encoding(tmp_path):
    config_file = tmp_path / "config.ini"

    config_file.write_text(
        "[CSV]\n"
        "encoding =   \n",
        encoding="utf-8"
    )

    with pytest.raises(ValueError, match="encoding cannot be empty"):
        load_config(config_file)


def test_load_config_invalid_log_level(tmp_path):
    config_file = tmp_path / "config.ini"

    config_file.write_text(
        "[LOGGING]\n"
        "level = INVALID\n",
        encoding="utf-8"
    )

    with pytest.raises(ValueError, match="Invalid logging level"):
        load_config(config_file)


def test_configure_logging(monkeypatch, tmp_path):
    basic_config = Mock()

    monkeypatch.setattr(
        converter.logging,
        "basicConfig",
        basic_config
    )

    config = {
        "log_level": "DEBUG",
        "log_file": str(tmp_path / "test.log")
    }

    configure_logging(config)

    basic_config.assert_called_once()

    logging_arguments = basic_config.call_args.kwargs

    assert logging_arguments["level"] == logging.DEBUG
    assert logging_arguments["force"] is True
    assert len(logging_arguments["handlers"]) == 2


def test_google_sheet_url_validation():
    valid_url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    invalid_url = "https://example.com/spreadsheet"

    assert is_google_sheet_url(valid_url) is True
    assert is_google_sheet_url(invalid_url) is False


def test_google_sheet_url_http():
    url = (
        "http://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    assert is_google_sheet_url(url) is True


def test_google_sheet_url_invalid_domain():
    url = (
        "https://docs.google.com.example.com/"
        "spreadsheets/d/ABC123456/edit"
    )

    assert is_google_sheet_url(url) is False


def test_extract_spreadsheet_id():
    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    spreadsheet_id = extract_spreadsheet_id(url)

    assert spreadsheet_id == "ABC123456"


def test_extract_spreadsheet_id_invalid_url():
    with pytest.raises(ValueError):
        extract_spreadsheet_id(
            "https://docs.google.com/spreadsheets/"
        )


def test_download_google_sheet(monkeypatch):
    response = Mock()
    response.content = b"PK\x03\x04test"
    response.headers = {
        "Content-Disposition": (
            'attachment; filename="Quarterly Report.xlsx"'
        )
    }

    get_mock = Mock(return_value=response)

    monkeypatch.setattr(
        converter.requests,
        "get",
        get_mock
    )

    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    content, workbook_name = download_google_sheet(
        url,
        timeout=45
    )

    assert content == b"PK\x03\x04test"
    assert workbook_name == "Quarterly Report"

    get_mock.assert_called_once_with(
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/export?format=xlsx",
        timeout=45
    )

    response.raise_for_status.assert_called_once()


def test_download_google_sheet_without_filename(monkeypatch):
    response = Mock()
    response.content = b"PK\x03\x04test"
    response.headers = {}

    monkeypatch.setattr(
        converter.requests,
        "get",
        Mock(return_value=response)
    )

    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    content, workbook_name = download_google_sheet(
        url,
        timeout=30
    )

    assert content == b"PK\x03\x04test"
    assert workbook_name == "ABC123456"


def test_download_google_sheet_invalid_content(monkeypatch):
    response = Mock()
    response.content = b"Not an Excel file"
    response.headers = {}

    monkeypatch.setattr(
        converter.requests,
        "get",
        Mock(return_value=response)
    )

    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    with pytest.raises(ValueError, match="valid Excel file"):
        download_google_sheet(url, timeout=30)


def test_download_google_sheet_http_error(monkeypatch):
    response = Mock()
    response.raise_for_status.side_effect = requests.HTTPError(
        "403 Forbidden"
    )

    monkeypatch.setattr(
        converter.requests,
        "get",
        Mock(return_value=response)
    )

    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    with pytest.raises(requests.HTTPError):
        download_google_sheet(url, timeout=30)


def test_read_workbook(tmp_path):
    workbook_path = tmp_path / "sales.xlsx"

    create_test_workbook(workbook_path)

    worksheets, workbook_name = read_workbook(
        str(workbook_path),
        timeout=30
    )

    assert workbook_name == "sales"
    assert len(worksheets) == 2

    assert "Office Supply Sales" in worksheets
    assert "Food Sales" in worksheets

    assert worksheets["Office Supply Sales"][0] == (
        "Product",
        "Quantity",
        "Price"
    )

    assert worksheets["Office Supply Sales"][1] == (
        "Pen",
        10,
        5
    )


def test_read_workbook_without_extension(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    workbook_path = tmp_path / "sales.xlsx"
    create_test_workbook(workbook_path)

    worksheets, workbook_name = read_workbook(
        "sales",
        timeout=30
    )

    assert workbook_name == "sales"
    assert len(worksheets) == 2


def test_read_workbook_google_sheet(monkeypatch):
    workbook_content = create_workbook_bytes()

    download_mock = Mock(
        return_value=(workbook_content, "Online Report")
    )

    monkeypatch.setattr(
        converter,
        "download_google_sheet",
        download_mock
    )

    url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    worksheets, workbook_name = read_workbook(
        url,
        timeout=45
    )

    assert workbook_name == "Online Report"
    assert "Online Sales" in worksheets

    download_mock.assert_called_once_with(url, 45)


def test_read_workbook_unsupported_url():
    with pytest.raises(
        ValueError,
        match="Only Google Sheets URLs are supported"
    ):
        read_workbook(
            "https://example.com/file.xlsx",
            timeout=30
        )


def test_read_workbook_missing_file(tmp_path):
    workbook_path = tmp_path / "missing.xlsx"

    with pytest.raises(FileNotFoundError):
        read_workbook(str(workbook_path), timeout=30)


def test_read_workbook_invalid_extension(tmp_path):
    workbook_path = tmp_path / "sales.csv"
    workbook_path.touch()

    with pytest.raises(ValueError):
        read_workbook(str(workbook_path), timeout=30)


def test_read_workbook_corrupted_file(tmp_path):
    workbook_path = tmp_path / "corrupted.xlsx"
    workbook_path.write_bytes(b"Not a valid Excel workbook")

    with pytest.raises(zipfile.BadZipFile):
        read_workbook(str(workbook_path), timeout=30)


def test_read_workbook_empty_workbook(tmp_path, monkeypatch):
    workbook_path = tmp_path / "empty.xlsx"
    workbook_path.touch()

    fake_workbook = Mock()
    fake_workbook.sheetnames = []
    fake_workbook.worksheets = []

    monkeypatch.setattr(
        converter,
        "load_workbook",
        Mock(return_value=fake_workbook)
    )

    with pytest.raises(
        ValueError,
        match="does not contain any worksheets"
    ):
        read_workbook(str(workbook_path), timeout=30)

    fake_workbook.close.assert_called_once()


@pytest.mark.parametrize(
    "original, expected",
    [
        ("Sales/2026", "Sales_2026"),
        ("Food:Sales", "Food_Sales"),
        ("Office*Supplies", "Office_Supplies"),
        ("  Sales  ", "Sales"),
        ("", "Sheet"),
        ("CON", "_CON"),
        ("PRN", "_PRN"),
        ("COM1", "_COM1"),
        ("LPT9", "_LPT9"),
        ("Sales?2026", "Sales_2026"),
        ("Sales|2026", "Sales_2026"),
    ]
)
def test_sanitize_filename(original, expected):
    assert sanitize_filename(original) == expected


def test_sanitize_filename_maximum_length():
    filename = "A" * 200

    result = sanitize_filename(filename)

    assert len(result) == 150


def test_get_unique_filename():
    used_names = set()

    first_name = get_unique_filename(
        "Sales",
        used_names
    )

    second_name = get_unique_filename(
        "Sales",
        used_names
    )

    assert first_name == "Sales"
    assert second_name == "Sales_2"


def test_get_unique_filename_case_insensitive():
    used_names = set()

    first_name = get_unique_filename(
        "Sales",
        used_names
    )

    second_name = get_unique_filename(
        "sales",
        used_names
    )

    assert first_name == "Sales"
    assert second_name == "sales_2"


def test_normalize_row():
    row = (
        "Pen",
        2001.0,
        10.0,
        10.5,
        None,
        True
    )

    result = normalize_row(row)

    assert result == [
        "Pen",
        2001,
        10,
        10.5,
        None,
        True
    ]


def test_convert_workbook(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    worksheets = {
        "Office Supply Sales": [
            ("Product", "Quantity", "Price"),
            ("Pen", 10, 5),
            ("Notebook", 5, 40)
        ],
        "Food Sales": [
            ("Food", "Quantity", "Price"),
            ("Rice", 20, 60),
            ("Wheat", 15, 45)
        ]
    }

    output_directory, converted_files = convert_workbook(
        worksheets,
        "sales",
        "utf-8"
    )

    assert output_directory == tmp_path / "sales"
    assert output_directory.is_dir()
    assert len(converted_files) == 2

    office_file = output_directory / "Office Supply Sales.csv"
    food_file = output_directory / "Food Sales.csv"

    assert office_file.is_file()
    assert food_file.is_file()

    with office_file.open(
        "r",
        newline="",
        encoding="utf-8"
    ) as csv_file:
        rows = list(csv.reader(csv_file))

    assert rows == [
        ["Product", "Quantity", "Price"],
        ["Pen", "10", "5"],
        ["Notebook", "5", "40"]
    ]


def test_convert_workbook_duplicate_sheet_names(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    worksheets = {
        "Sales/2026": [
            ("Product",),
            ("Pen",)
        ],
        "Sales:2026": [
            ("Product",),
            ("Book",)
        ]
    }

    output_directory, converted_files = convert_workbook(
        worksheets,
        "sales",
        "utf-8"
    )

    assert len(converted_files) == 2

    assert (output_directory / "Sales_2026.csv").is_file()
    assert (output_directory / "Sales_2026_2.csv").is_file()


def test_convert_workbook_normalizes_numbers(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    worksheets = {
        "Numbers": [
            ("Year", "Quantity", "Price"),
            (2001.0, 10.0, 10.5)
        ]
    }

    _, converted_files = convert_workbook(
        worksheets,
        "numeric_data",
        "utf-8"
    )

    with converted_files[0].open(
        "r",
        newline="",
        encoding="utf-8"
    ) as csv_file:
        rows = list(csv.reader(csv_file))

    assert rows[1] == ["2001", "10", "10.5"]


def test_convert_workbook_empty_worksheets(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    output_directory, converted_files = convert_workbook(
        {},
        "empty",
        "utf-8"
    )

    assert output_directory.is_dir()
    assert converted_files == []


def test_convert_workbook_invalid_encoding(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    worksheets = {
        "Sales": [
            ("Product",),
            ("Café",)
        ]
    }

    with pytest.raises(UnicodeEncodeError):
        convert_workbook(
            worksheets,
            "sales",
            "ascii"
        )


def test_build_parser():
    parser = build_parser()

    args = parser.parse_args([
        "sales.xlsx",
        "--config",
        "custom.ini"
    ])

    assert args.input == "sales.xlsx"
    assert args.config == "custom.ini"


def test_build_parser_default_config():
    parser = build_parser()

    args = parser.parse_args(["sales.xlsx"])

    assert args.input == "sales.xlsx"
    assert args.config == "config.ini"


def test_build_parser_missing_input():
    parser = build_parser()

    with pytest.raises(SystemExit):
        parser.parse_args([])


def test_main_success(monkeypatch, tmp_path, capsys):
    output_file = tmp_path / "Sales.csv"
    output_file.touch()

    monkeypatch.setattr(
        "sys.argv",
        ["converter.py", "sales.xlsx"]
    )

    monkeypatch.setattr(
        converter,
        "load_config",
        Mock(return_value={
            "timeout": 30,
            "encoding": "utf-8",
            "log_level": "INFO",
            "log_file": "converter.log"
        })
    )

    configure_mock = Mock()
    monkeypatch.setattr(
        converter,
        "configure_logging",
        configure_mock
    )

    monkeypatch.setattr(
        converter,
        "read_workbook",
        Mock(return_value=({"Sales": []}, "sales"))
    )

    monkeypatch.setattr(
        converter,
        "convert_workbook",
        Mock(return_value=(tmp_path, [output_file]))
    )

    result = converter.main()

    captured = capsys.readouterr()

    assert result == 0
    assert "Conversion completed successfully." in captured.out
    assert "Created: Sales.csv" in captured.out
    configure_mock.assert_called_once()


@pytest.mark.parametrize(
    "error, expected_message",
    [
        (
            FileNotFoundError("missing file"),
            "Error: missing file"
        ),
        (
            requests.ConnectionError("connection failed"),
            "Download error: connection failed"
        ),
        (
            InvalidFileException("invalid workbook"),
            "Invalid Excel file: invalid workbook"
        ),
        (
            zipfile.BadZipFile("bad archive"),
            "Corrupted Excel file: bad archive"
        ),
        (
            configparser.Error("bad configuration"),
            "Configuration error: bad configuration"
        ),
        (
            UnicodeError("encoding failed"),
            "Encoding error: encoding failed"
        ),
        (
            csv.Error("CSV failed"),
            "Conversion error: CSV failed"
        ),
        (
            ValueError("invalid value"),
            "Conversion error: invalid value"
        ),
        (
            OSError("file operation failed"),
            "Conversion error: file operation failed"
        ),
    ]
)
def test_main_handles_errors(
    monkeypatch,
    capsys,
    error,
    expected_message
):
    monkeypatch.setattr(
        "sys.argv",
        ["converter.py", "sales.xlsx"]
    )

    def raise_error(config_path):
        raise error

    monkeypatch.setattr(
        converter,
        "load_config",
        raise_error
    )

    result = converter.main()

    captured = capsys.readouterr()

    assert result == 1
    assert expected_message in captured.out