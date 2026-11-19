import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import json
import queue
import threading
import openpyxl
from openpyxl.utils.cell import column_index_from_string


CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.json"


def load_app_config(config_path=CONFIG_PATH):
    try:
        with Path(config_path).open(encoding="utf-8") as config_file:
            config = json.load(config_file)
        if not isinstance(config, dict):
            raise ValueError("Корень config.json должен быть объектом.")
        order_config = config["order"]
        upd_config = config["upd"]
    except (OSError, json.JSONDecodeError, KeyError, TypeError) as error:
        raise ValueError(f"Не удалось загрузить config.json: {error}") from error
    if not isinstance(order_config, dict) or not isinstance(upd_config, dict):
        raise ValueError("Секции order и upd в config.json должны быть объектами.")

    for section_name, section, row_keys in (
        ("order", order_config, ("sheet_index", "first_data_row")),
        ("upd", upd_config, ("sheet_index", "first_data_row")),
    ):
        for key in row_keys:
            value = section.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ValueError(
                    f"В config.json параметр {section_name}.{key} должен быть "
                    "положительным целым числом."
                )

    column_keys = (
        (order_config, "internal_code_column"),
        (order_config, "supplier_article_column"),
        (order_config, "order_number_label_column"),
        (order_config, "order_number_value_column"),
        (upd_config, "article_column"),
        (upd_config, "order_number_column"),
    )
    for section, key in column_keys:
        try:
            column_index_from_string(section[key])
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError(f"Некорректный столбец config.json: {key}.") from error

    for section, key in (
        (order_config, "order_number_label"),
        (upd_config, "basis_label"),
    ):
        if not isinstance(section.get(key), str) or not section[key].strip():
            raise ValueError(f"В config.json поле {key} не должно быть пустым.")

    row_offset = upd_config.get("order_number_row_offset")
    if isinstance(row_offset, bool) or not isinstance(row_offset, int):
        raise ValueError(
            "В config.json параметр upd.order_number_row_offset должен быть целым числом."
        )
    return config


class UPDReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Редактор УПД и Сверка Заказов (Python)")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.order_path = ""
        self.upd_path = ""
        self._events = queue.Queue()
        self._processing = False
        
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close_window)

    def setup_ui(self):
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.btn_order = tk.Button(btn_frame, text="1. Выбрать файл ЗАКАЗА", command=self.load_order, width=30)
        self.btn_order.grid(row=0, column=0, pady=5)
        
        self.btn_upd = tk.Button(btn_frame, text="2. Выбрать файл УПД", command=self.load_upd, width=30)
        self.btn_upd.grid(row=1, column=0, pady=5)
        
        self.btn_run = tk.Button(btn_frame, text="ЗАПУСТИТЬ ОБРАБОТКУ", command=self.process_files, bg="#90ee90", width=25, height=3)
        self.btn_run.grid(row=0, column=1, rowspan=2, padx=20)
        
        self.log_area = tk.Text(self.root, wrap=tk.WORD, height=18)
        self.log_area.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    def log(self, text):
        self.log_area.insert(tk.END, text + "\n")
        self.log_area.see(tk.END)

    def close_window(self):
        if self._processing:
            messagebox.showwarning(
                "Обработка",
                "Дождитесь завершения обработки перед закрытием окна."
            )
            return
        self.root.destroy()

    @staticmethod
    def normalize_article(value, number_format="General"):
        if value is None:
            return ""
        if isinstance(value, bool):
            return str(value).strip()
        if isinstance(value, (int, float)):
            if isinstance(value, float) and not value.is_integer():
                return str(value).strip()
            text = str(int(value))
            if isinstance(number_format, str) and number_format and set(number_format) == {"0"}:
                sign = "-" if text.startswith("-") else ""
                digits = text[1:] if sign else text
                return sign + digits.zfill(len(number_format))
            return text
        return str(value).strip()

    @staticmethod
    def unique_output_path(source):
        candidate = source.with_name(f"{source.stem}_reconciled{source.suffix}")
        counter = 1
        while candidate.exists():
            candidate = source.with_name(
                f"{source.stem}_reconciled_{counter}{source.suffix}"
            )
            counter += 1
        return candidate

    def load_order(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.order_path = file
            self.btn_order.config(text=f"Заказ: {Path(file).name}")

    def load_upd(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.upd_path = file
            self.btn_upd.config(text=f"УПД: {Path(file).name}")

    def process_files(self):
        if not self.order_path or not self.upd_path:
            messagebox.showwarning("Ошибка", "Выберите оба файла перед началом!")
            return

        order_path = Path(self.order_path)
        upd_path = Path(self.upd_path)

        if not order_path.is_file() or not upd_path.is_file():
            messagebox.showerror("Ошибка", "Один из входных файлов не существует.")
            return
        if order_path.resolve() == upd_path.resolve():
            messagebox.showerror("Ошибка", "Файл заказа и файл УПД должны быть разными.")
            return
        if order_path.suffix.lower() != ".xlsx" or upd_path.suffix.lower() != ".xlsx":
            messagebox.showerror("Ошибка", "Поддерживаются только файлы формата .xlsx.")
            return

        self._processing = True
        self.btn_order.config(state=tk.DISABLED)
        self.btn_upd.config(state=tk.DISABLED)
        self.btn_run.config(state=tk.DISABLED)
        try:
            worker = threading.Thread(
                target=self._process_files_worker,
                args=(order_path, upd_path),
                daemon=True
            )
            worker.start()
        except Exception as error:
            self._processing = False
            self.btn_order.config(state=tk.NORMAL)
            self.btn_upd.config(state=tk.NORMAL)
            self.btn_run.config(state=tk.NORMAL)
            messagebox.showerror("Ошибка", f"Не удалось запустить обработку:\n{error}")
            return
        self.root.after(50, self._poll_processing_events)

    def _process_files_worker(self, order_path, upd_path, config=None):
        try:
            result = self._run_processing(order_path, upd_path, config)
        except Exception as error:
            result = ("error", str(error))
        self._events.put(("finished", result[0], result[1]))

    def _run_processing(self, order_path, upd_path, config=None):
        wb_order = None
        wb_upd = None
        wb_verify = None

        try:
            if config is None:
                config = load_app_config()
            order_config = config["order"]
            upd_config = config["upd"]
            order_code_column = column_index_from_string(
                order_config["internal_code_column"]
            )
            order_article_column = column_index_from_string(
                order_config["supplier_article_column"]
            )
            order_label_column = column_index_from_string(
                order_config["order_number_label_column"]
            )
            order_value_column = column_index_from_string(
                order_config["order_number_value_column"]
            )
            upd_article_column = column_index_from_string(
                upd_config["article_column"]
            )
            upd_order_column = column_index_from_string(
                upd_config["order_number_column"]
            )

            self._events.put(("log", "--- Старт обработки ---"))

            # 1. Читаем файл заказа
            self._events.put(("log", "Чтение файла заказа..."))
            wb_order = openpyxl.load_workbook(order_path, data_only=True)
            order_sheet_index = order_config["sheet_index"] - 1
            if order_sheet_index >= len(wb_order.worksheets):
                raise ValueError("В заказе отсутствует лист, указанный в config.json.")
            ws_order = wb_order.worksheets[order_sheet_index]

            art_map = {}
            duplicate_articles = set()
            order_number = ""

            for row_idx in range(1, ws_order.max_row + 1):
                order_code_cell = ws_order.cell(
                    row=row_idx, column=order_code_column
                )
                order_article_cell = ws_order.cell(
                    row=row_idx, column=order_article_column
                )
                cell_b = self.normalize_article(
                    order_code_cell.value, order_code_cell.number_format
                )
                cell_c = self.normalize_article(
                    order_article_cell.value, order_article_cell.number_format
                )

                if row_idx >= order_config["first_data_row"] and cell_b and cell_c:
                    if cell_c in art_map and art_map[cell_c] != cell_b:
                        duplicate_articles.add(cell_c)
                    else:
                        art_map[cell_c] = cell_b

                label_cell = ws_order.cell(
                    row=row_idx, column=order_label_column
                )
                label = self.normalize_article(
                    label_cell.value, label_cell.number_format
                )
                if order_config["order_number_label"] in label:
                    value_cell = ws_order.cell(
                        row=row_idx, column=order_value_column
                    )
                    order_number = self.normalize_article(
                        value_cell.value, value_cell.number_format
                    )

            for article in duplicate_articles:
                art_map.pop(article, None)
                self._events.put(
                    ("log", f"[ДУБЛИКАТ] Артикул '{article}' имеет неоднозначное соответствие.")
                )

            if not order_number:
                return ("error", "Номер заказа не найден в файле.")
            if not art_map:
                return (
                    "error",
                    "Не найдены сопоставления по структуре, заданной в config.json."
                )

            self._events.put(
                ("log", f"Загружено позиций: {len(art_map)}. Номер заказа: {order_number}")
            )

            # 2. Обрабатываем УПД
            self._events.put(("log", "Применение изменений к УПД..."))
            wb_upd = openpyxl.load_workbook(upd_path)
            upd_sheet_index = upd_config["sheet_index"] - 1
            if upd_sheet_index >= len(wb_upd.worksheets):
                raise ValueError("В УПД отсутствует лист, указанный в config.json.")
            ws_upd = wb_upd.worksheets[upd_sheet_index]
            preview = []

            for row_idx in range(upd_config["first_data_row"], ws_upd.max_row + 1):
                article_cell = ws_upd.cell(row=row_idx, column=upd_article_column)
                current_art = self.normalize_article(
                    article_cell.value, article_cell.number_format
                )
                if not current_art:
                    continue

                if current_art in art_map:
                    preview.append((row_idx, current_art, art_map[current_art]))
                else:
                    self._events.put(
                        ("log", f"[НЕ НАЙДЕН] Строка {row_idx}: артикул '{current_art}'")
                    )

            if not preview:
                return (
                    "warning",
                    "Не найдено однозначных совпадений. Проверьте столбец и строку "
                    "начала данных в config.json."
                )

            # Validate the order-number destination before asking the user to apply changes.
            basis_matches = []
            for row in ws_upd.iter_rows():
                for cell in row:
                    if str(cell.value or "").strip() == upd_config["basis_label"]:
                        basis_matches.append(cell)

            if not basis_matches:
                return ("error", f'Поле "{upd_config["basis_label"]}" не найдено.')
            if len(basis_matches) > 1:
                return (
                    "error",
                    f'Найдено несколько полей "{upd_config["basis_label"]}" '
                    f"({len(basis_matches)})."
                )

            target_row = basis_matches[0].row + upd_config["order_number_row_offset"]
            if target_row <= 0:
                return (
                    "error",
                    f'Поле "{upd_config["basis_label"]}" находится в первой строке; '
                    "невозможно определить строку для номера заказа."
                )

            preview_text = "\n".join(
                f"Строка {row}: {old_article} -> {new_article}"
                for row, old_article, new_article in preview
            )
            target_column_name = upd_config["order_number_column"].upper()
            preview_text += (
                f"\n\n{upd_config['basis_label']}: номер заказа {order_number} "
                f"будет записан в ячейку {target_column_name}{target_row}."
            )
            response = queue.Queue(maxsize=1)
            self._events.put(("preview", preview_text, response))
            if not response.get():
                self._events.put(("log", "Операция отменена пользователем."))
                return ("cancel", "")

            for row_idx, current_art, new_article in preview:
                ws_upd.cell(row=row_idx, column=upd_article_column).value = new_article
                self._events.put(
                    ("log", f"[OK] Строка {row_idx}: заменён артикул '{current_art}' на '{new_article}'")
                )

            ws_upd.cell(row=target_row, column=upd_order_column).value = order_number
            self._events.put(("log", f"Номер заказа успешно вписан в строку {target_row}."))

            # 3. Сохранение
            output_path = self.unique_output_path(upd_path)
            wb_upd.save(output_path)
            wb_upd.close()
            wb_upd = None

            # Verify the serialized workbook rather than relying only on in-memory values.
            wb_verify = openpyxl.load_workbook(output_path, data_only=True, read_only=True)
            ws_verify = wb_verify.worksheets[upd_sheet_index]
            for row_idx, _, expected_article in preview:
                saved_article = ws_verify.cell(
                    row=row_idx, column=upd_article_column
                ).value
                if saved_article != expected_article:
                    raise ValueError(
                        f"Проверка результата не пройдена: ячейка "
                        f"{upd_config['article_column'].upper()}{row_idx} "
                        f"содержит {saved_article!r}, ожидалось {expected_article!r}."
                    )

            saved_order_number = ws_verify.cell(
                row=target_row, column=upd_order_column
            ).value
            if saved_order_number != order_number:
                raise ValueError(
                    f"Проверка результата не пройдена: ячейка "
                    f"{target_column_name}{target_row} "
                    f"содержит {saved_order_number!r}, ожидалось {order_number!r}."
                )
            wb_verify.close()
            wb_verify = None

            self._events.put(("log", "--- Успешно завершено! ---"))
            return ("done", str(output_path))

        finally:
            if wb_order:
                wb_order.close()
            if wb_upd:
                wb_upd.close()
            if wb_verify:
                wb_verify.close()

    def _poll_processing_events(self):
        while True:
            try:
                event = self._events.get_nowait()
            except queue.Empty:
                break

            if event[0] == "log":
                self.log(event[1])
            elif event[0] == "preview":
                accepted = messagebox.askyesno(
                    "Подтверждение изменений",
                    f"Будут внесены следующие изменения:\n\n{event[1]}\n\nПродолжить?"
                )
                event[2].put(accepted)
            elif event[0] == "finished":
                self._finish_processing(event[1], event[2])

        if self._processing:
            self.root.after(50, self._poll_processing_events)

    def _finish_processing(self, result, detail):
        self._processing = False
        self.btn_order.config(state=tk.NORMAL)
        self.btn_upd.config(state=tk.NORMAL)
        self.btn_run.config(state=tk.NORMAL)

        if result == "done":
            messagebox.showinfo("Успех", f"Обработка завершена!\nФайл сохранён:\n{detail}")
        elif result == "warning":
            messagebox.showwarning("Предпросмотр", detail)
        elif result == "error":
            self.log(f"ОШИБКА: {detail}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{detail}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UPDReconcilerApp(root)
    root.mainloop()