import csv

import pytest
from openpyxl import Workbook

from converter import (
    convert_workbook,
    extract_spreadsheet_id,
    get_unique_filename,
    is_google_sheet_url,
    load_config,
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


def test_load_config_missing_file(tmp_path):
    config_file = tmp_path / "missing.ini"

    with pytest.raises(FileNotFoundError):
        load_config(config_file)


def test_google_sheet_url_validation():
    valid_url = (
        "https://docs.google.com/spreadsheets/d/"
        "ABC123456/edit"
    )

    invalid_url = "https://example.com/spreadsheet"

    assert is_google_sheet_url(valid_url) is True
    assert is_google_sheet_url(invalid_url) is False


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


def test_read_workbook_missing_file(tmp_path):
    workbook_path = tmp_path / "missing.xlsx"

    with pytest.raises(FileNotFoundError):
        read_workbook(str(workbook_path), timeout=30)


def test_read_workbook_invalid_extension(tmp_path):
    workbook_path = tmp_path / "sales.csv"
    workbook_path.touch()

    with pytest.raises(ValueError):
        read_workbook(str(workbook_path), timeout=30)


@pytest.mark.parametrize(
    "original, expected",
    [
        ("Sales/2026", "Sales_2026"),
        ("Food:Sales", "Food_Sales"),
        ("Office*Supplies", "Office_Supplies"),
        ("  Sales  ", "Sales"),
        ("", "Sheet"),
        ("CON", "_CON"),
    ]
)
def test_sanitize_filename(original, expected):
    assert sanitize_filename(original) == expected


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
