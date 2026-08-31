import tkinter as tk
from tkinter import filedialog, messagebox
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
            self.btn_order.config(text=f"Заказ: {file.split('/')[-1]}")

    def load_upd(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.upd_path = file
            self.btn_upd.config(text=f"УПД: {file.split('/')[-1]}")

    def process_files(self):
        if not self.order_path or not self.upd_path:
            messagebox.showwarning("Ошибка", "Выберите оба файла перед началом!")
            return
            
        try:
            self.log("--- Старт обработки ---")
            
            # 1. Читаем файл заказа
            self.log("Чтение файла заказа...")
            wb_order = openpyxl.load_workbook(self.order_path, data_only=True)
            ws_order = wb_order.active
            
            art_map = {}
            order_number = ""
            
            for row in ws_order.iter_rows(min_row=1, max_row=ws_order.max_row):
                # Индексы 1 и 2 соответствуют колонкам B и C
                cell_b = str(row[1].value or "").strip()
                cell_c = str(row[2].value or "").strip()
                
                if cell_b and cell_c:
                    art_map[cell_c] = cell_b
                    
                if "Заказ №" in cell_b:
                    order_number = str(row[2].value or "").strip()
                    
            self.log(f"Загружено позиций: {len(art_map)}. Номер заказа: {order_number}")
            wb_order.close()

            # 2. Обрабатываем УПД
            self.log("Применение изменений к УПД...")
            wb_upd = openpyxl.load_workbook(self.upd_path)
            ws_upd = wb_upd.active
            
            # Замена артикулов начиная с 17 строки (Колонка B - это индекс 2 в openpyxl)
            for row_idx in range(17, ws_upd.max_row + 1):
                current_art = str(ws_upd.cell(row=row_idx, column=2).value or "").strip()
                if not current_art:
                    continue
                    
                if current_art in art_map:
                    ws_upd.cell(row=row_idx, column=2).value = art_map[current_art]
                    self.log(f"[OK] Строка {row_idx}: Заменен номер '{current_art}'")

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
            wb_upd.save(self.upd_path)
            self.log("--- Успешно завершено! ---")
            messagebox.showinfo("Успех", "Обработка завершена!")
            
        except Exception as e:
            self.log(f"ОШИБКА: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UPDReconcilerApp(root)
    root.mainloop()