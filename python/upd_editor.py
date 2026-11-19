import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import openpyxl

class UPDReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Редактор УПД и Сверка Заказов (Python)")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.order_path = ""
        self.upd_path = ""
        
        self.setup_ui()

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
        self.root.update()

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

        wb_order = None
        wb_upd = None

        try:
            self.log("--- Старт обработки ---")

            # 1. Читаем файл заказа
            self.log("Чтение файла заказа...")
            wb_order = openpyxl.load_workbook(order_path, data_only=True)
            ws_order = wb_order.active

            art_map = {}
            duplicate_articles = set()
            order_number = ""

            for row_idx in range(1, ws_order.max_row + 1):
                order_code_cell = ws_order.cell(row=row_idx, column=2)
                order_article_cell = ws_order.cell(row=row_idx, column=3)
                cell_b = self.normalize_article(
                    order_code_cell.value, order_code_cell.number_format
                )
                cell_c = self.normalize_article(
                    order_article_cell.value, order_article_cell.number_format
                )

                if cell_b and cell_c:
                    if cell_c in art_map and art_map[cell_c] != cell_b:
                        duplicate_articles.add(cell_c)
                    else:
                        art_map[cell_c] = cell_b

                if "Заказ №" in cell_b:
                    order_number = cell_c

            for article in duplicate_articles:
                art_map.pop(article, None)
                self.log(f"[ДУБЛИКАТ] Артикул '{article}' имеет неоднозначное соответствие.")

            if not order_number:
                messagebox.showerror(
                    "Ошибка",
                    "Номер заказа не найден в файле. Обработка остановлена."
                )
                return

            self.log(f"Загружено позиций: {len(art_map)}. Номер заказа: {order_number}")

            # 2. Обрабатываем УПД
            self.log("Применение изменений к УПД...")
            wb_upd = openpyxl.load_workbook(upd_path)
            ws_upd = wb_upd.active
            preview = []

            # Замена артикулов начиная с 17 строки (Колонка B - это индекс 2 в openpyxl)
            for row_idx in range(17, ws_upd.max_row + 1):
                article_cell = ws_upd.cell(row=row_idx, column=2)
                current_art = self.normalize_article(
                    article_cell.value, article_cell.number_format
                )
                if not current_art:
                    continue

                if current_art in art_map:
                    preview.append((row_idx, current_art, art_map[current_art]))
                else:
                    self.log(f"[НЕ НАЙДЕН] Строка {row_idx}: артикул '{current_art}'")

            if not preview:
                messagebox.showwarning("Предпросмотр", "Не найдено ни одного однозначного артикула для замены.")
                return

            preview_text = "\n".join(
                f"Строка {row}: {old_article} -> {new_article}"
                for row, old_article, new_article in preview
            )
            if not messagebox.askyesno(
                "Подтверждение изменений",
                f"Будут внесены следующие изменения:\n\n{preview_text}\n\nПродолжить?"
            ):
                self.log("Операция отменена пользователем.")
                return

            for row_idx, current_art, new_article in preview:
                ws_upd.cell(row=row_idx, column=2).value = new_article
                self.log(f"[OK] Строка {row_idx}: заменён артикул '{current_art}' на '{new_article}'")

            # Поиск "Основание передачи": совпадение должно быть единственным
            basis_matches = []
            for row in ws_upd.iter_rows():
                for cell in row:
                    if str(cell.value or "").strip() == "Основание передачи":
                        basis_matches.append(cell)

            if not basis_matches:
                messagebox.showerror(
                    "Ошибка",
                    'Поле "Основание передачи" не найдено. Обработка остановлена.'
                )
                return
            if len(basis_matches) > 1:
                messagebox.showerror(
                    "Ошибка",
                    f'Найдено несколько полей "Основание передачи" ({len(basis_matches)}). '
                    "Обработка остановлена."
                )
                return

            target_row = basis_matches[0].row - 1
            if target_row <= 0:
                messagebox.showerror(
                    "Ошибка",
                    'Поле "Основание передачи" находится в первой строке; '
                    "невозможно определить строку для номера заказа."
                )
                return
            ws_upd.cell(row=target_row, column=20).value = order_number
            self.log(f"Номер заказа успешно вписан в строку {target_row}.")

            # 3. Сохранение
            output_path = self.unique_output_path(upd_path)
            wb_upd.save(output_path)
            self.log("--- Успешно завершено! ---")
            messagebox.showinfo("Успех", f"Обработка завершена!\nФайл сохранён:\n{output_path}")

        except Exception as e:
            self.log(f"ОШИБКА: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            if wb_order:
                wb_order.close()
            if wb_upd:
                wb_upd.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = UPDReconcilerApp(root)
    root.mainloop()