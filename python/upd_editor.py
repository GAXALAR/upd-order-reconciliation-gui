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

            for row in ws_order.iter_rows(min_row=1, max_row=ws_order.max_row):
                # Индексы 1 и 2 соответствуют колонкам B и C
                cell_b = str(row[1].value or "").strip()
                cell_c = str(row[2].value or "").strip()

                if cell_b and cell_c:
                    if cell_c in art_map and art_map[cell_c] != cell_b:
                        duplicate_articles.add(cell_c)
                    else:
                        art_map[cell_c] = cell_b

                if "Заказ №" in cell_b:
                    order_number = str(row[2].value or "").strip()

            for article in duplicate_articles:
                art_map.pop(article, None)
                self.log(f"[ДУБЛИКАТ] Артикул '{article}' имеет неоднозначное соответствие.")

            self.log(f"Загружено позиций: {len(art_map)}. Номер заказа: {order_number}")

            # 2. Обрабатываем УПД
            self.log("Применение изменений к УПД...")
            wb_upd = openpyxl.load_workbook(upd_path)
            ws_upd = wb_upd.active
            preview = []

            # Замена артикулов начиная с 17 строки (Колонка B - это индекс 2 в openpyxl)
            for row_idx in range(17, ws_upd.max_row + 1):
                current_art = str(ws_upd.cell(row=row_idx, column=2).value or "").strip()
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

            # Поиск "Основание передачи" и вставка номера заказа
            for row in ws_upd.iter_rows(min_row=1, max_row=ws_upd.max_row):
                for cell in row:
                    if str(cell.value or "").strip() == "Основание передачи":
                        target_row = cell.row - 1
                        if target_row > 0:
                            # Колонка T - это 20
                            ws_upd.cell(row=target_row, column=20).value = order_number
                            self.log(f"Номер заказа успешно вписан в строку {target_row}.")
                        break

            # 3. Сохранение
            output_path = upd_path.with_name(f"{upd_path.stem}_reconciled{upd_path.suffix}")
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