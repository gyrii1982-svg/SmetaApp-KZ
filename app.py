import os
import sys
import datetime
import openpyxl
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import threading
import random

# Безопасный импорт Pillow для логотипа
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(base_path)

def load_price_database():
    price_file = "prices.xlsx"
    db = {}
    if not os.path.exists(price_file):
        print(f"⚠️ Файл '{price_file}' не найден!")
        return db
    try:
        wb_p = openpyxl.load_workbook(price_file, data_only=True)
        ws_p = wb_p.active
        row = 2
        while row <= 5000:
            name = ws_p.cell(row=row, column=1).value
            if not name:
                row += 1
                continue
            price_val = ws_p.cell(row=row, column=2).value
            supplier_val = ws_p.cell(row=row, column=3).value
            unit_val = ws_p.cell(row=row, column=4).value  # Ед. изм. из prices.xlsx (столбец 4)
            
            try:
                price = float(price_val) if price_val is not None else 0.0
            except ValueError:
                price = 0.0
            
            supplier = str(supplier_val) if supplier_val is not None else "Не указан"
            unit = str(unit_val).strip() if unit_val is not None else "шт."
            
            db[str(name).strip().lower()] = (price, supplier, unit)
            row += 1
    except Exception as e:
        print(f"⚠️ Ошибка чтения prices.xlsx: {e}")
    return db

class SmetaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ультимативный СМР-комплекс v6.4 (С ценами из Excel + Логотип)")
        self.root.geometry("850x720")
        self.root.minsize(750, 600)

        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # --- Загрузка логотипа ---
        self.logo_img = None
        if HAS_PIL:
            logo_path = "logo.png"
            if os.path.exists(logo_path):
                try:
                    img = Image.open(logo_path)
                    img = img.resize((55, 55), Image.Resampling.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(img)
                except Exception as e:
                    print(f"⚠️ Ошибка загрузки логотипа: {e}")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Вкладка 1: Сметный расчет и Консоль-лог ---
        self.tab_smeta = tk.Frame(self.notebook, bg="#f0f4f1")
        self.notebook.add(self.tab_smeta, text=" 📊 Расчёт сметы и Консоль ")

        # --- Вкладка 2: Динамическая Аркада ---
        self.tab_game = tk.Frame(self.notebook, bg="#1e1e1e")
        self.notebook.add(self.tab_game, text=" 🕹️ Релакс-Аркада ")

        self.setup_smeta_tab()
        self.setup_game_tab()

    def setup_smeta_tab(self):
        top_frame = tk.Frame(self.tab_smeta, bg="#f0f4f1", pady=10, padx=10)
        top_frame.pack(fill=tk.X)

        if self.logo_img:
            lbl_logo = tk.Label(top_frame, image=self.logo_img, bg="#f0f4f1")
            lbl_logo.pack(side=tk.LEFT, padx=(0, 15))

        self.btn_run = tk.Button(
            top_frame, 
            text="🚀 Запустить расчёт сметы", 
            font=("Arial", 11, "bold"), 
            bg="#4CAF50", 
            fg="white", 
            padx=12, 
            pady=6,
            command=self.start_calculation_thread
        )
        self.btn_run.pack(side=tk.LEFT)

        self.lbl_status = tk.Label(top_frame, text="Статус: Ожидание запуска", font=("Arial", 10), bg="#f0f4f1", fg="gray")
        self.lbl_status.pack(side=tk.LEFT, padx=15)

        # Встроенная консоль
        self.log_area = scrolledtext.ScrolledText(self.tab_smeta, wrap=tk.WORD, font=("Consolas", 10), bg="#1e1e1e", fg="#00ff00")
        self.log_area.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        self.log("=" * 60)
        self.log("  Система готова к работе с файлом prices.xlsx.")
        self.log("  Нажмите кнопку 'Запустить расчёт сметы' для старта.")
        self.log("=" * 60)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def start_calculation_thread(self):
        self.btn_run.config(state=tk.DISABLED)
        self.lbl_status.config(text="Статус: Выполняется расчет...", fg="orange")
        threading.Thread(target=self.run_calculation, daemon=True).start()

    def run_calculation(self):
        report_lines = []
        
        def log_msg(msg):
            self.log(msg)
            report_lines.append(msg)

        log_msg("=" * 60)
        log_msg(f"  РАСЧЁТ СМР (prices.xlsx + Ед. изм. + Отчёт)")
        log_msg(f"  Дата запуска: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        log_msg("=" * 60)
        
        file_name = "SMR_smeta.xlsx"
        if not os.path.exists(file_name):
            log_msg(f"\n❌ Ошибка: Файл сметы '{file_name}' не найден!")
            self.finish_calculation(success=False)
            return

        price_db = load_price_database()
        log_msg(f"\n📚 Загружено позиций из файла prices.xlsx: {len(price_db)}")

        try:
            wb = openpyxl.load_workbook(file_name)
            ws = wb["Смета СМР"]
            ws.cell(row=12, column=8).value = "Статус / Поставщик"
            
            row_idx = 13
            total_python_sum = 0.0
            error_count = 0
            
            while row_idx <= 150:
                name_cell = ws.cell(row=row_idx, column=3).value
                if not name_cell:
                    ws.cell(row=row_idx, column=4).value = None
                    ws.cell(row=row_idx, column=6).value = None
                    ws.cell(row=row_idx, column=7).value = f"=E{row_idx}*F{row_idx}"
                    ws.cell(row=row_idx, column=8).value = None
                    row_idx += 1
                    continue
                    
                qty_cell = ws.cell(row=row_idx, column=5).value
                
                qty = 0.0
                qty_error = False
                if isinstance(qty_cell, (datetime.date, datetime.datetime)):
                    qty_error = True
                elif qty_cell is not None:
                    try:
                        qty = float(qty_cell)
                    except ValueError:
                        qty_error = True
                else:
                    qty = 0.0
                    
                if qty_error:
                    ws.cell(row=row_idx, column=4).value = "-"
                    ws.cell(row=row_idx, column=5).value = "ОШИБКА"
                    ws.cell(row=row_idx, column=6).value = 0.0
                    ws.cell(row=row_idx, column=7).value = 0.0
                    ws.cell(row=row_idx, column=8).value = "🔴 ОШИБКА: Дата в количестве!"
                    error_count += 1
                    row_idx += 1
                    continue
                else:
                    ws.cell(row=row_idx, column=5).value = qty
                    ws.cell(row=row_idx, column=5).number_format = '0.00'

                material_name = str(name_cell).strip()
                name_key = material_name.lower()
                
                price = 0.0
                supplier = ""
                unit = "шт."
                status_text = ""
                
                # Поиск по базе prices.xlsx
                if name_key in price_db:
                    price, supplier, unit = price_db[name_key]
                    status_text = f"🟢 100% | {supplier}"
                else:
                    best_match_score = 0.0
                    best_price = 0.0
                    best_supplier = ""
                    best_unit = "шт."
                    
                    for db_name, (p, s, u) in price_db.items():
                        score = SequenceMatcher(None, name_key, db_name).ratio()
                        if score > best_match_score:
                            best_match_score = score
                            best_price = p
                            best_supplier = s
                            best_unit = u
                    
                    if best_match_score >= 0.80:
                        price = best_price
                        supplier = best_supplier
                        unit = best_unit
                        percent_val = int(best_match_score * 100)
                        status_text = f"🟡 {percent_val}% схожесть | {supplier}"
                    else:
                        supplier = "🔴 ТМЦ НЕ НАЙДЕНО В БАЗЕ"
                        status_text = supplier
                        unit = "-"

                current_price = ws.cell(row=row_idx, column=6).value
                if current_price is not None and current_price != "" and float(current_price) > 0:
                    price = float(current_price)
                    if "НЕ НАЙДЕНО" in status_text:
                        status_text = "🟡 Ручная цена"

                # ЗАПИСЬ ВСЕХ ПАРАМЕТРОВ В ФАЙЛ SMR_smeta.xlsx (включая столбец D / колонка 4)
                ws.cell(row=row_idx, column=4).value = unit  # Ед. изм.
                ws.cell(row=row_idx, column=5).value = qty   # Количество
                ws.cell(row=row_idx, column=6).value = price if price > 0 else None  # Цена за ед.
                ws.cell(row=row_idx, column=7).value = f"=E{row_idx}*F{row_idx}"     # Сумма
                ws.cell(row=row_idx, column=8).value = status_text  # Статус / Поставщик
                
                total_python_sum += (qty * price)
                if "НЕ НАЙДЕНО" in status_text:
                    error_count += 1
                    
                log_msg(f"🔍 Строка {row_idx}: {material_name} (Ед: {unit}, Кол-во: {qty}, Цена: {price}) -> {status_text}")
                row_idx += 1

            # Итоговые формулы в файле сметы
            ws['H3'] = "=SUM(G13:G150)"
            ws['H5'] = "=H3*C7"
            ws['H6'] = "=H5*C8/100"
            ws['H7'] = "=(H5+H6)*C9/100"
            ws['H9'] = "=H5+H6+H7"

            wb.save(file_name)
            
            log_msg("\n" + "=" * 50)
            log_msg(f"  РАСЧЁТ ЗАВЕРШЕН. Предупреждений/ошибок: {error_count}")
            log_msg(f"  Контрольная сумма (Python): {total_python_sum:,.2f} тг")
            log_msg("=" * 50)

            with open("smr_report.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(report_lines))
            log_msg("\n📄 Отчёт успешно сохранён в файл: smr_report.txt")
            
            self.finish_calculation(success=True)

        except Exception as e:
            log_msg(f"\n❌ Ошибка: {e}")
            self.finish_calculation(success=False)

    def finish_calculation(self, success):
        self.btn_run.config(state=tk.NORMAL)
        if success:
            self.lbl_status.config(text="Статус: Готово ✅", fg="green")
            messagebox.showinfo("Успех", "Расчёт успешно завершён! Столбец Ед. изм. и сметы обновлены.")
        else:
            self.lbl_status.config(text="Статус: Ошибка ❌", fg="red")
            messagebox.showerror("Ошибка", "Произошла ошибка при расчёте. Смотрите консоль в окне.")

    # --- Настройка аркады ---
    def setup_game_tab(self):
        game_top = tk.Frame(self.tab_game, bg="#1e1e1e")
        game_top.pack(fill=tk.X, padx=10, pady=5)

        if self.logo_img:
            lbl_game_logo = tk.Label(game_top, image=self.logo_img, bg="#1e1e1e")
            lbl_game_logo.pack(side=tk.LEFT, padx=(0, 15))

        lbl_info = tk.Label(game_top, text="🕹️ ПРОБЕЛ — прыжок! Случайные препятствия и нарастающая скорость!", font=("Arial", 11, "bold"), bg="#1e1e1e", fg="#00ffcc")
        lbl_info.pack(side=tk.LEFT, pady=5)

        self.canvas = tk.Canvas(self.tab_game, width=780, height=320, bg="#111111", highlightthickness=0)
        self.canvas.pack(padx=10, pady=5)

        self.ground = self.canvas.create_rectangle(0, 270, 780, 280, fill="#00ff00")
        self.player = self.canvas.create_rectangle(50, 230, 80, 270, fill="#ffcc00")
        self.obstacle = self.canvas.create_rectangle(800, 230, 830, 270, fill="#ff3333")

        self.score_text = self.canvas.create_text(680, 30, text="Счёт: 0 | Скор: 6.0", fill="white", font=("Consolas", 12, "bold"))

        self.is_jumping = False
        self.jump_speed = 0
        self.game_score = 0
        self.obs_speed = 6.0

        self.root.bind("<space>", self.jump)
        self.update_game()

    def jump(self, event=None):
        if not self.is_jumping and self.notebook.index(self.notebook.select()) == 1:
            self.is_jumping = True
            self.jump_speed = -14

    def update_game(self):
        if self.notebook.index(self.notebook.select()) == 1:
            if self.is_jumping:
                self.canvas.move(self.player, 0, self.jump_speed)
                self.jump_speed += 1
                
                p_coords = self.canvas.coords(self.player)
                if p_coords[3] >= 270:
                    self.canvas.coords(self.player, p_coords[0], 230, p_coords[2], 270)
                    self.is_jumping = False

            self.canvas.move(self.obstacle, -self.obs_speed, 0)
            obs_coords = self.canvas.coords(self.obstacle)

            if obs_coords[2] < 0:
                random_offset = random.randint(150, 450)
                self.canvas.coords(self.obstacle, 800 + random_offset, 230, 830 + random_offset, 270)
                self.game_score += 1
                self.obs_speed += 0.25
                if self.obs_speed > 18.0:
                    self.obs_speed = 18.0
                self.canvas.itemconfig(self.score_text, text=f"Счёт: {self.game_score} | Скор: {self.obs_speed:.1f}")

            p_coords = self.canvas.coords(self.player)
            if (obs_coords[0] < p_coords[2] and obs_coords[2] > p_coords[0] and 
                obs_coords[1] < p_coords[3] and obs_coords[3] > p_coords[1]):
                self.game_score = 0
                self.obs_speed = 6.0
                self.canvas.itemconfig(self.score_text, text=f"Счёт: 0 | КРАШ! Сброс")
                self.canvas.coords(self.obstacle, 800, 230, 830, 270)

        self.root.after(20, self.update_game)

if __name__ == "__main__":
    root = tk.Tk()
    app = SmetaApp(root)
    root.mainloop()