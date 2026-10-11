import xml.etree.ElementTree as ET

from openpyxl import Workbook

import main
from main import MainWindow


def test_quote_identifier():
    assert main.quote_identifier("users") == "users"
    assert main.quote_identifier("mi tabla") == '"mi tabla"'
    assert main.quote_identifier('a"b') == '"a""b"'


def test_sql_value():
    assert main.sql_value(None, "n") == "NULL"
    assert main.sql_value("admin", "s") == "'admin'"
    assert main.sql_value("O'Brien", "s") == "'O''Brien'"
    assert main.sql_value(1, "n") == "1"
    assert main.sql_value(1.5, "n") == "1.5"
    assert main.sql_value(True, "b") == "1"
    assert main.sql_value(False, "b") == "0"


def test_generate_insert_statements(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "users"
    sheet.append(["id", "name", "active"])
    sheet.append([1, "O'Brien", 1])
    sheet.append([2, "admin", 0])
    sheet.append([None, None, None])  # fila vacía, debe omitirse

    path = tmp_path / "test.xlsx"
    workbook.save(path)

    statements = main.generate_insert_statements(str(path))

    assert statements == [
        "INSERT INTO users (id, name, active) VALUES (1, 'O''Brien', 1);",
        "INSERT INTO users (id, name, active) VALUES (2, 'admin', 0);",
    ]


def test_list_sheet_names(tmp_path):
    workbook = Workbook()
    workbook.active.title = "users"
    workbook.create_sheet("data")
    path = tmp_path / "multi.xlsx"
    workbook.save(path)

    assert main.list_sheet_names(str(path)) == ["users", "data"]


def test_generate_insert_statements_filtered(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "users"
    sheet.append(["id", "name"])
    sheet.append([1, "admin"])

    other = workbook.create_sheet("data")
    other.append(["id", "value"])
    other.append([1, 42])

    path = tmp_path / "multi.xlsx"
    workbook.save(path)

    statements = main.generate_insert_statements(str(path), ["data"])

    assert statements == ["INSERT INTO data (id, value) VALUES (1, 42);"]


def test_generate_insert_statements_with_create(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "users"
    sheet.append(["id", "name", "active"])
    sheet.append([1, "admin", 1])

    path = tmp_path / "test.xlsx"
    workbook.save(path)

    statements = main.generate_insert_statements(str(path), include_create=True)

    assert statements == [
        "CREATE TABLE IF NOT EXISTS users (id INTEGER, name TEXT, active INTEGER);",
        "INSERT INTO users (id, name, active) VALUES (1, 'admin', 1);",
    ]


def test_sheet_selection_dialog(qapp):
    dialog = main.SheetSelectionDialog(["users", "data"])
    dialog.sheet_list.setCurrentRow(1)
    dialog._accept()

    assert dialog.selected_sheet == "data"
    assert dialog.result() == 1  # Accepted


def test_menus(qapp):
    window = MainWindow()
    menu_titles = [action.text() for action in window.menuBar().actions()]
    assert "Archivo" in menu_titles
    assert "Opciones" in menu_titles
    assert "Ver" in menu_titles
    assert "Ayuda" in menu_titles


def test_preferences_persisted(qapp):
    window = MainWindow()
    window.toggle_dark_mode()
    window._save_preferences()

    root = ET.parse(main.CONFIG_FILE).getroot()
    assert root.findtext("settings/dark_mode") == "true"

    restored = MainWindow()
    assert restored.dark_mode is True


def test_include_create_preference_persisted(qapp):
    window = MainWindow()
    window._set_include_create(True)
    window._save_preferences()

    root = ET.parse(main.CONFIG_FILE).getroot()
    assert root.findtext("settings/include_create") == "true"

    restored = MainWindow()
    assert restored.include_create is True
