import copy
import json
import queue
import threading

import openpyxl
import pytest

from upd_editor import UPDReconcilerApp, load_app_config


ORDER_LABEL = "Заказ №"
BASIS_LABEL = "Основание передачи"


def create_order(path, mappings, order_number="ORD-TEST-001"):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    if order_number is not None:
        worksheet.cell(row=1, column=2, value=ORDER_LABEL)
        worksheet.cell(row=1, column=3, value=order_number)
    for row_index, (internal_code, supplier_article) in enumerate(mappings, start=2):
        worksheet.cell(row=row_index, column=2, value=internal_code)
        worksheet.cell(row=row_index, column=3, value=supplier_article)
    workbook.save(path)


def create_upd(path, articles, basis_rows=(20,)):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    for row_index, article in enumerate(articles, start=17):
        worksheet.cell(row=row_index, column=2, value=article)
    for row_index in basis_rows:
        worksheet.cell(row=row_index, column=1, value=BASIS_LABEL)
    workbook.save(path)


def run_processing(order_path, upd_path, accept_preview=True, config=None):
    app = UPDReconcilerApp.__new__(UPDReconcilerApp)
    app._events = queue.Queue()
    worker = threading.Thread(
        target=app._process_files_worker,
        args=(order_path, upd_path, config),
        daemon=True,
    )
    worker.start()
    events = []
    preview = None
    result = None

    while result is None:
        try:
            event = app._events.get(timeout=5)
        except queue.Empty:
            continue
        if event[0] == "preview":
            preview = event[1]
            event[2].put(accept_preview)
        elif event[0] == "finished":
            result = event[1:]
        else:
            events.append(event)

    worker.join(timeout=5)
    assert not worker.is_alive()
    return result, events, preview


def test_preview_verifies_saved_copy_and_preserves_source(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [("INT-001", "SUP-001")])
    create_upd(upd_path, ["SUP-001"])

    result, _, preview = run_processing(order_path, upd_path)

    assert result == ("done", str(tmp_path / "upd_reconciled.xlsx"))
    assert "SUP-001 -> INT-001" in preview
    assert "ORD-TEST-001" in preview
    assert "T19" in preview

    output = openpyxl.load_workbook(result[1], data_only=True)
    source = openpyxl.load_workbook(upd_path, data_only=True)
    assert output.active["B17"].value == "INT-001"
    assert output.active["T19"].value == "ORD-TEST-001"
    assert source.active["B17"].value == "SUP-001"


def test_numeric_articles_use_zero_format_and_text_articles_are_unchanged(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [(45, 123)])
    workbook = openpyxl.load_workbook(order_path)
    workbook.active["B2"].number_format = "000000"
    workbook.active["C2"].number_format = "000000"
    workbook.save(order_path)
    create_upd(upd_path, [123])
    workbook = openpyxl.load_workbook(upd_path)
    workbook.active["B17"].number_format = "000000"
    workbook.save(upd_path)

    result, _, preview = run_processing(order_path, upd_path)

    assert result[0] == "done"
    assert "000123 -> 000045" in preview
    output = openpyxl.load_workbook(result[1], data_only=True)
    assert output.active["B17"].value == "000045"

    assert UPDReconcilerApp.normalize_article("00123.0") == "00123.0"


def test_ambiguous_mapping_is_not_applied(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(
        order_path,
        [("INT-001", "SUP-001"), ("INT-002", "SUP-001")],
    )
    create_upd(upd_path, ["SUP-001"])

    result, events, preview = run_processing(order_path, upd_path)

    assert result[0] == "error"
    assert "сопоставления" in result[1]
    assert preview is None
    assert any("ДУБЛИКАТ" in event[1] for event in events if event[0] == "log")
    assert not (tmp_path / "upd_reconciled.xlsx").exists()
    source = openpyxl.load_workbook(upd_path, data_only=True)
    assert source.active["B17"].value == "SUP-001"


def test_configured_layout_and_sheet_are_used(tmp_path):
    order_path = tmp_path / "custom_order.xlsx"
    upd_path = tmp_path / "custom_upd.xlsx"

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["D1"] = ORDER_LABEL
    worksheet["E1"] = "ORD-CUSTOM-1"
    worksheet["F4"] = "INT-CUSTOM-1"
    worksheet["G4"] = "SUP-CUSTOM-1"
    workbook.save(order_path)

    workbook = openpyxl.Workbook()
    workbook.create_sheet("Data")
    worksheet = workbook.worksheets[1]
    worksheet["D5"] = "SUP-CUSTOM-1"
    worksheet["C10"] = BASIS_LABEL
    workbook.save(upd_path)

    config = copy.deepcopy(load_app_config())
    config["order"].update(
        {
            "first_data_row": 4,
            "internal_code_column": "F",
            "supplier_article_column": "G",
            "order_number_label_column": "D",
            "order_number_value_column": "E",
        }
    )
    config["upd"].update(
        {
            "sheet_index": 2,
            "first_data_row": 5,
            "article_column": "D",
            "basis_label": BASIS_LABEL,
            "order_number_column": "J",
            "order_number_row_offset": -2,
        }
    )

    result, _, preview = run_processing(order_path, upd_path, config=config)

    assert result[0] == "done"
    assert "J8" in preview
    output = openpyxl.load_workbook(result[1], data_only=True)
    assert output.worksheets[1]["D5"].value == "INT-CUSTOM-1"
    assert output.worksheets[1]["J8"].value == "ORD-CUSTOM-1"


def test_invalid_config_column_is_rejected(tmp_path):
    config_path = tmp_path / "config.json"
    config = copy.deepcopy(load_app_config())
    config["upd"]["article_column"] = "not-a-column"
    config_path.write_text(json.dumps(config), encoding="utf-8")

    with pytest.raises(ValueError, match="Некорректный столбец"):
        load_app_config(config_path)


def test_missing_order_number_stops_before_preview(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [("INT-001", "SUP-001")], order_number=None)
    create_upd(upd_path, ["SUP-001"])

    result, _, preview = run_processing(order_path, upd_path)

    assert result[0] == "error"
    assert "Номер заказа" in result[1]
    assert preview is None
    assert not (tmp_path / "upd_reconciled.xlsx").exists()


@pytest.mark.parametrize(
    ("basis_rows", "expected_error"),
    [
        ((), "не найдено"),
        ((20, 22), "несколько"),
        ((1,), "первой строке"),
    ],
)
def test_invalid_order_number_destination_stops_before_preview(
    tmp_path, basis_rows, expected_error
):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [("INT-001", "SUP-001")])
    create_upd(upd_path, ["SUP-001"], basis_rows=basis_rows)

    result, _, preview = run_processing(order_path, upd_path)

    assert result[0] == "error"
    assert expected_error in result[1]
    assert preview is None
    assert not (tmp_path / "upd_reconciled.xlsx").exists()


def test_declining_preview_does_not_save_output(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [("INT-001", "SUP-001")])
    create_upd(upd_path, ["SUP-001"])

    result, _, preview = run_processing(order_path, upd_path, accept_preview=False)

    assert result == ("cancel", "")
    assert preview is not None
    assert not (tmp_path / "upd_reconciled.xlsx").exists()
    source = openpyxl.load_workbook(upd_path, data_only=True)
    assert source.active["B17"].value == "SUP-001"


def test_existing_result_is_not_overwritten(tmp_path):
    order_path = tmp_path / "order.xlsx"
    upd_path = tmp_path / "upd.xlsx"
    create_order(order_path, [("INT-001", "SUP-001")])
    create_upd(upd_path, ["SUP-001"])
    existing_output = tmp_path / "upd_reconciled.xlsx"
    existing_output.write_bytes(b"keep this existing result")

    result, _, _ = run_processing(order_path, upd_path)

    assert result == ("done", str(tmp_path / "upd_reconciled_1.xlsx"))
    assert existing_output.read_bytes() == b"keep this existing result"
