import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkcalendar import Calendar, DateEntry
import datetime
import locale
import pyodbc
import os
import sys
import shutil
from final_handbook_fix import HandbookForm
from replacement_form import ReplacementForm
from db_manager import DatabaseManager
import re

# Константи для стилю
APP_BG_COLOR = "#f0f0f0"  # Світло-сірий фон
HEADER_BG_COLOR = "#4a6984"  # Темно-синій для заголовків
HEADER_FG_COLOR = "white"  # Білий текст для заголовків
BUTTON_BG_COLOR = "#1a365d"  # Темно-синій для кнопок (змінено на більш насичений)
BUTTON_FG_COLOR = "#ffffff"  # Білий текст для кнопок
HIGHLIGHT_COLOR = "#2ecc71"  # Яскраво-зелений для виділення
FONT_FAMILY = "Arial"  # Шрифт для всього додатку
DEFAULT_FONT = (FONT_FAMILY, 10)  # Звичайний шрифт
HEADER_FONT = (FONT_FAMILY, 12, "bold")  # Шрифт для заголовків
BUTTON_FONT = (FONT_FAMILY, 12, "bold")  # Шрифт для кнопок (збільшено розмір до 12)
PADDING = 10  # Стандартний відступ

# Set locale to Ukrainian
try:
    locale.setlocale(locale.LC_ALL, 'uk_UA.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Ukrainian_Ukraine.1251')
    except:
        print("Українська локаль недоступна, використовується стандартна")

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Програма для ведення замін")
        self.geometry("1000x700")  # Збільшено розмір вікна
        self.minsize(1000, 700)  # Збільшено мінімальний розмір
        
        # Застосовуємо стиль до всього додатку
        self.configure(bg=APP_BG_COLOR)
        self.style = ttk.Style()
        self.style.configure('TFrame', background=APP_BG_COLOR)
        self.style.configure('TLabel', background=APP_BG_COLOR, font=DEFAULT_FONT)
        self.style.configure('TButton', font=BUTTON_FONT)
        
        # Створюємо кастомні стилі
        self.style.configure('Header.TLabel', background=HEADER_BG_COLOR, foreground=HEADER_FG_COLOR, font=HEADER_FONT, padding=PADDING)
        self.style.configure('Header.TFrame', background=HEADER_BG_COLOR)
        self.style.configure('Action.TButton', background=BUTTON_BG_COLOR, foreground=BUTTON_FG_COLOR)
        
        # Встановлюємо початкову ширину вікна
        self.window_width = 1000  # Збільшено з 800 до 1000
        
        # Connect to database
        self.db_path = os.path.abspath("dataBase.mdb")
        self.conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path}'
        
        # Initialize data
        self.teachers_list = self.get_teachers_from_db()
        
        # Створюємо папку Zaminy, якщо вона не існує
        # Визначаємо шлях до папки програми
        try:
            # Спочатку пробуємо отримати шлях до EXE-файлу (для скомпільованої програми)
            if getattr(sys, 'frozen', False):
                # Якщо програма скомпільована в EXE
                application_path = os.path.dirname(sys.executable)
            else:
                # Якщо програма запущена з Python
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            self.replacements_dir = os.path.join(application_path, "Zaminy")
        except Exception as e:
            print(f"Помилка при визначенні шляху до папки програми: {e}")
            # Використовуємо старий метод як запасний варіант
            self.replacements_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zaminy")
        
        if not os.path.exists(self.replacements_dir):
            os.makedirs(self.replacements_dir)
            
        # Змінні для слайдера
        self.slider_visible = False
        self.slider_width = 350  # Збільшено ширину слайдера до 350
        self.animation_speed = 10  # Швидкість анімації (менше = швидше)
        self.slider_frame = None
        self.slider_container = None
        self.slider_position = 'right'  # Позиція слайдера (right)
        
        # Поточний навчальний рік (останній створений або поточний)
        self.current_academic_year = self.get_last_academic_year() or self.get_current_academic_year()
        self.academic_year_var = tk.StringVar(value=self.current_academic_year)
        
        self.create_widgets()
    
    def get_teachers_from_db(self):
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            # Try to find all tables in the database
            try:
                cursor.execute("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
                tables = [row[0] for row in cursor.fetchall()]
                print(f"Доступні таблиці: {tables}")
                
                # Try to find a table that might contain teachers
                teacher_table = None
                for table in tables:
                    table_str = str(table).lower()
                    if 'виклад' in table_str or 'teacher' in table_str or 'препод' in table_str:
                        teacher_table = table
                        break
                
                if teacher_table:
                    print(f"Знайдено таблицю викладачів: {teacher_table}")
                    cursor.execute(f"SELECT * FROM [{teacher_table}]")
                else:
                    # If no teacher table found, try common names with square brackets
                    try:
                        cursor.execute("SELECT * FROM [teachers]")
                    except:
                        try:
                            cursor.execute("SELECT * FROM [Викладачі]")
                        except:
                            try:
                                cursor.execute("SELECT * FROM [Викладач]")
                            except:
                                try:
                                    cursor.execute("SELECT * FROM [Преподаватели]")
                                except Exception as e:
                                    print(f"Не вдалося знайти таблицю викладачів: {e}")
                                    return []
            except Exception as e:
                print(f"Помилка при пошуку таблиць: {e}")
                # Try common names with square brackets as fallback
                try:
                    cursor.execute("SELECT * FROM [teachers]")
                except:
                    try:
                        cursor.execute("SELECT * FROM [Викладачі]")
                    except:
                        try:
                            cursor.execute("SELECT * FROM [Викладач]")
                        except Exception as e:
                            print(f"Не вдалося знайти таблицю викладачів: {e}")
                            return []
            rows = cursor.fetchall()
            # Get the column names from cursor description
            columns = [column[0] for column in cursor.description]
            print(f"Доступні колонки: {columns}")
            
            # Try to find the appropriate column for teacher names
            name_column = None
            for i, col in enumerate(columns):
                col_lower = col.lower() if isinstance(col, str) else str(col).lower()
                print(f"Перевірка колонки: {col} (індекс {i})")
                if col_lower in ['name', 'прізвище', 'pib', 'піб', 'викладач', 'назва']:
                    name_column = i
                    print(f"Знайдено колонку з іменами викладачів: {col} (індекс {i})")
                    break
            
            if name_column is not None:
                teachers = []
                for row in rows:
                    if len(row) > name_column:
                        teacher_name = row[name_column]
                        print(f"Додавання викладача: {teacher_name}")
                        teachers.append(teacher_name)
            else:
                # If we can't find a suitable column, use the first one
                print("Не знайдено відповідної колонки, використовуємо першу колонку")
                teachers = []
                for row in rows:
                    if len(row) > 0:
                        teacher_name = row[0]
                        print(f"Додавання викладача (перша колонка): {teacher_name}")
                        teachers.append(teacher_name)
            
            print(f"Загальна кількість викладачів: {len(teachers)}")
            print(f"Перші 5 викладачів: {teachers[:5] if len(teachers) >= 5 else teachers}")
                
            conn.close()
            return teachers
        except Exception as e:
            print(f"Помилка підключення до бази даних: {e}")
            return []
    
    def get_current_academic_year(self):
        # Визначаємо поточний навчальний рік
        now = datetime.datetime.now()
        if now.month >= 9:  # Якщо поточний місяць вересень або пізніше
            return f"{now.year}-{now.year + 1}"
        else:
            return f"{now.year - 1}-{now.year}"
    
    def get_last_academic_year(self):
        # Визначаємо останній створений навчальний рік
        if os.path.exists(self.replacements_dir):
            folders = [f for f in os.listdir(self.replacements_dir) 
                      if os.path.isdir(os.path.join(self.replacements_dir, f))]
            if folders:
                # Сортуємо папки за датою створення (найновіша перша)
                folders.sort(key=lambda x: os.path.getctime(os.path.join(self.replacements_dir, x)), reverse=True)
                return folders[0]
        return None
    
    def create_academic_year_folder(self):
        # Відкриваємо діалогове вікно для введення назви навчального року
        academic_year = simpledialog.askstring("Новий навчальний рік", 
                                            "Введіть назву навчального року (наприклад, 2025-2026):",
                                            initialvalue=self.get_current_academic_year())
        
        if academic_year:
            # Оновлюємо змінну з поточним навчальним роком
            self.academic_year_var.set(academic_year)
            
            # Створюємо папку для нового навчального року
            year_dir = os.path.join(self.replacements_dir, academic_year)
            
            if not os.path.exists(year_dir):
                os.makedirs(year_dir)
                messagebox.showinfo("Інформація", f"Створено папку для навчального року {academic_year}")
            else:
                messagebox.showinfo("Інформація", f"Папка для навчального року {academic_year} вже існує")
            
            # Оновлюємо TreeView і текст поточного року
            self.update_treeview()
            if self.year_label:
                self.year_label.config(text=f"Поточний навчальний рік: {academic_year}")
    
    def update_treeview(self):
        # Оновлюємо TreeView з вмістом папки Zaminy
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Додаємо навчальні роки
        if os.path.exists(self.replacements_dir):
            for year_folder in sorted(os.listdir(self.replacements_dir), reverse=True):
                year_path = os.path.join(self.replacements_dir, year_folder)
                if os.path.isdir(year_path):
                    year_id = self.tree.insert("", "end", text=year_folder, values=(year_path,))
                    
                    # Додаємо місяці в навчальному році
                    for month_folder in sorted(os.listdir(year_path)):
                        month_path = os.path.join(year_path, month_folder)
                        if os.path.isdir(month_path):
                            month_id = self.tree.insert(year_id, "end", text=month_folder, values=(month_path,))
                            
                            # Додаємо файли в місяці
                            for file in sorted(os.listdir(month_path)):
                                file_path = os.path.join(month_path, file)
                                if os.path.isfile(file_path):
                                    self.tree.insert(month_id, "end", text=file, values=(file_path,))
    
    def on_window_resize(self, event=None):
        # Обробка зміни розміру вікна
        if hasattr(self, 'slider_container') and self.slider_container:
            # Оновлюємо розмір слайдера
            self.window_width = self.winfo_width()
            
            if self.slider_visible:
                self.slider_container.place(x=self.window_width - self.slider_width, y=0, 
                                          width=self.slider_width, height=self.winfo_height())
                self.toggle_button.place(x=self.window_width - self.slider_width - 20, y=self.winfo_height()//2 - 15, width=20, height=30)
            else:
                self.slider_container.place(x=self.window_width, y=0, 
                                          width=self.slider_width, height=self.winfo_height())
                self.toggle_button.place(x=self.window_width - 20, y=self.winfo_height()//2 - 15, width=20, height=30)
    
    def animate_slider(self, start_x, end_x, show):
        # Анімація виїзду/заїзду слайдера
        current_x = start_x
        step = (end_x - start_x) / 20  # 20 кроків анімації (збільшено для плавнішої анімації)
        self.animation_speed = 10  # Зменшуємо час між кроками для плавнішої анімації
        print(f"Animation step: {step}")
        
        def move_slider():
            nonlocal current_x
            # Перевіряємо, чи досягнуто кінцевої позиції
            if (show and step < 0 and current_x <= end_x) or \
               (show and step > 0 and current_x >= end_x) or \
               (not show and step < 0 and current_x <= end_x) or \
               (not show and step > 0 and current_x >= end_x):
                # Досягнуто кінцевої позиції
                self.slider_container.place(x=end_x, y=0, width=self.slider_width, height=self.winfo_height())
                # Встановлюємо кнопку в кінцеву позицію
                button_x = end_x - 20 if show else self.window_width - 20
                self.toggle_button.place(x=button_x, y=self.winfo_height()//2 - 15, width=20, height=30)
                
                # Змінюємо текст кнопки
                self.toggle_button.config(text=">" if show else "<")
                print(f"Animation completed at position {end_x}")
            else:
                # Продовжуємо анімацію
                current_x += step
                print(f"Current position: {current_x}")
                self.slider_container.place(x=current_x, y=0, width=self.slider_width, height=self.winfo_height())
                # Переміщуємо кнопку разом зі слайдером
                button_x = current_x - 20 if show else self.window_width - 20
                self.toggle_button.place(x=button_x, y=self.winfo_height()//2 - 15, width=20, height=30)
                self.after(self.animation_speed, move_slider)
        
        move_slider()
    
    def create_slider_content(self):
        # Створюємо вміст слайдера
        # Очищуємо вміст слайдера перед створенням нового
        for widget in self.slider_container.winfo_children():
            widget.destroy()
        
        # Створюємо вміст слайдера
        slider_frame = tk.Frame(self.slider_container, bg=APP_BG_COLOR)
        slider_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок слайдера
        header_frame = tk.Frame(slider_frame, bg=HEADER_BG_COLOR)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        header_label = tk.Label(header_frame, text="Навчальний рік", font=HEADER_FONT, bg=HEADER_BG_COLOR, fg=HEADER_FG_COLOR)
        header_label.pack(pady=PADDING)
        
        # Фрейм для інформації про поточний навчальний рік
        info_frame = tk.Frame(slider_frame, bg=APP_BG_COLOR)
        info_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Мітка з поточним навчальним роком
        year_label_text = tk.Label(info_frame, text="Навч. рік:", bg=APP_BG_COLOR, font=BUTTON_FONT)
        year_label_text.pack(side=tk.LEFT, padx=(PADDING, 5))
        
        self.year_label = tk.Label(info_frame, text=self.academic_year_var.get(), 
                                  bg=APP_BG_COLOR, font=BUTTON_FONT, fg=BUTTON_BG_COLOR)
        self.year_label.pack(side=tk.LEFT, padx=0)
        
        # Фрейм для кнопки "Новий навч. рік"
        btn_frame = tk.Frame(slider_frame, bg=APP_BG_COLOR)
        btn_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Кнопка "Новий навчальний рік"
        new_year_btn = tk.Button(btn_frame, text="Новий навч. рік", command=self.create_academic_year_folder,
                               bg=HIGHLIGHT_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT, relief=tk.RAISED,
                               padx=PADDING, pady=PADDING//2, borderwidth=2)
        new_year_btn.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Заголовок для дерева файлів
        files_header = tk.Frame(slider_frame, bg=APP_BG_COLOR)
        files_header.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        tk.Label(files_header, text="Файли та папки", font=HEADER_FONT, bg=APP_BG_COLOR).pack(side=tk.LEFT)
        
        # Створюємо фрейм для TreeView
        tree_frame = tk.Frame(slider_frame, bg=APP_BG_COLOR)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=PADDING, pady=PADDING)
        
        # Налаштовуємо стиль для TreeView
        self.style.configure("Treeview", font=DEFAULT_FONT, rowheight=25)
        self.style.configure("Treeview.Heading", font=BUTTON_FONT)
        
        # Створюємо TreeView для відображення файлів та папок
        self.tree = ttk.Treeview(tree_frame, columns=("path",), displaycolumns=())
        self.tree.heading("#0", text="Файли та папки")
        self.tree.column("#0", width=self.slider_width - 30)  # Ширина стовпця з урахуванням скролбара
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Додаємо скролбар
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Прив'язуємо подію вибору елемента
        self.tree.bind("<<TreeviewSelect>>", self.on_treeview_select)
    
    def toggle_slider(self):
        # Показуємо/приховуємо слайдер для вибору навчального року
        self.window_width = self.winfo_width()
        print(f"Window width: {self.window_width}")
        
        if self.slider_visible:
            # Приховуємо слайдер
            start_x = self.window_width - self.slider_width
            end_x = self.window_width
            print(f"Hiding slider: from {start_x} to {end_x}")
            self.animate_slider(start_x, end_x, False)
            self.slider_visible = False
        else:
            # Показуємо слайдер
            # Спочатку створюємо вміст слайдера
            self.create_slider_content()
            # Оновлюємо дерево файлів
            self.update_treeview()
            
            # Запускаємо анімацію виїзду слайдера
            start_x = self.window_width  # Початкова позиція за межами вікна (справа)
            end_x = self.window_width - self.slider_width  # Кінцева позиція (зліва від правого краю)
            print(f"Showing slider: from {start_x} to {end_x}")
            self.animate_slider(start_x, end_x, True)
            self.slider_visible = True
    
    def select_academic_year(self, year):
        # Вибір навчального року зі слайдера
        if year:
            self.academic_year_var.set(year)
            if self.year_label:
                self.year_label.config(text=f"Поточний навчальний рік: {year}")
            self.update_treeview()
            # Закриваємо слайдер
            self.toggle_slider()
    
    def is_academic_year_format(self, text):
        # Перевіряємо, чи текст відповідає формату навчального року (YYYY-YYYY)
        return bool(re.match(r'^\d{4}-\d{4}$', text))
    
    def on_treeview_select(self, event):
        # Обробка вибору елемента в TreeView
        selected_item = self.tree.selection()
        if selected_item:
            item_text = self.tree.item(selected_item[0], "text")
            item_path = self.tree.item(selected_item[0], "values")[0] if self.tree.item(selected_item[0], "values") else None
            
            # Перевіряємо, чи це навчальний рік
            if self.is_academic_year_format(item_text):
                # Змінюємо поточний навчальний рік
                self.academic_year_var.set(item_text)
                self.current_academic_year = item_text
                
                # Оновлюємо мітку в слайдері
                self.year_label.config(text=item_text)
                
                # Повідомляємо про зміну навчального року
                messagebox.showinfo("Зміна навчального року", f"Поточний навчальний рік змінено на {item_text}")
                
                # Не закриваємо слайдер після вибору навчального року, щоб користувач міг продовжити перегляд файлів
                # self.toggle_slider() - видалено цей рядок
            elif os.path.isfile(item_path):
                # Якщо вибрано файл, можна додати код для його відкриття
                pass
            
            print(f"Selected: {item_text}, Path: {item_path}")
    
    def create_widgets(self):
        # Створюємо заголовок вікна
        header_frame = ttk.Frame(self, style='Header.TFrame')
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        header_label = ttk.Label(header_frame, text="Програма для ведення замін", style='Header.TLabel')
        header_label.pack(pady=PADDING)
        
        # Головний контейнер
        self.main_container = ttk.Frame(self)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=PADDING, pady=PADDING)
        
        # Розділяємо вікно на ліву та праву частини
        left_frame = ttk.Frame(self.main_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=PADDING)
        
        right_frame = ttk.Frame(self.main_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=PADDING)
        
        # Прив'язуємо подію зміни розміру вікна
        self.bind("<Configure>", self.on_window_resize)
        
        # Створюємо групу кнопок для швидкого доступу
        buttons_frame = ttk.LabelFrame(left_frame, text="Швидкий доступ")
        buttons_frame.pack(fill=tk.X, pady=PADDING)
        
        # Кнопка довідника з оновленим стилем
        handbook_btn = tk.Button(buttons_frame, text="Довідник", command=self.open_handbook, 
                               bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                               relief=tk.RAISED, borderwidth=2, padx=10, pady=5)
        handbook_btn.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Кнопка оновлення даних
        refresh_btn = tk.Button(buttons_frame, text="Оновити дані", command=self.refresh_database_data, 
                               bg="#e67e22", fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                               relief=tk.RAISED, borderwidth=2, padx=10, pady=5)
        refresh_btn.pack(fill=tk.X, padx=PADDING, pady=(0, PADDING//2))
        
        # Кнопка для переходу до форми замін (спочатку неактивна)
        self.replacement_btn = tk.Button(buttons_frame, text="Бланк замін", command=self.open_replacement_form, 
                                      bg="#a0a0a0", fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                                      relief=tk.RAISED, borderwidth=2, padx=10, pady=5,
                                      state=tk.DISABLED, cursor="arrow")
        self.replacement_btn.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Фрейм для вибору типу тижня
        week_frame = ttk.LabelFrame(left_frame, text="Навчання за")
        week_frame.pack(fill=tk.X, pady=PADDING)
        
        # Радіокнопки для вибору типу тижня
        self.week_type = tk.StringVar(value="чисельником")
        # Додаємо перевірку полів при зміні типу тижня
        self.week_type.trace_add("write", lambda *args: self.check_required_fields())
        ttk.Radiobutton(week_frame, text="Чисельником", variable=self.week_type, value="чисельником").pack(side=tk.LEFT, padx=PADDING)
        ttk.Radiobutton(week_frame, text="Знаменником", variable=self.week_type, value="знаменником").pack(side=tk.LEFT, padx=PADDING)
        
        # Фрейм для вибору дати
        date_frame = ttk.LabelFrame(left_frame, text="Дата заміни")
        date_frame.pack(fill=tk.X, pady=PADDING)
        
        date_container = ttk.Frame(date_frame)
        date_container.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Покращений вибір дати
        self.date_entry = DateEntry(date_container, width=15, background=BUTTON_BG_COLOR,
                                    foreground=BUTTON_FG_COLOR, borderwidth=2, 
                                    date_pattern='dd.MM.yyyy', font=DEFAULT_FONT)
        self.date_entry.pack(side=tk.LEFT, padx=PADDING)
        self.date_entry.bind("<<DateEntrySelected>>", self.update_date_info)
        
        # Фрейм для інформації про дату
        info_frame = ttk.Frame(date_frame)
        info_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        # Дата словами та день тижня
        ttk.Label(info_frame, text="Дата словами:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.date_text = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.date_text, state="readonly", width=30).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(info_frame, text="День тижня:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.weekday_text = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.weekday_text, state="readonly", width=30).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Оновлюємо інформацію про дату
        self.update_date_info()
        
        # Фрейм для чергування
        duty_frame = ttk.LabelFrame(left_frame, text="Чергування")
        duty_frame.pack(fill=tk.X, pady=PADDING)
        
        # Чергова група
        duty_group_frame = ttk.Frame(duty_frame)
        duty_group_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        ttk.Label(duty_group_frame, text="Чергова група:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.duty_group = tk.StringVar()
        # Додаємо перевірку полів при зміні чергової групи
        self.duty_group.trace_add("write", lambda *args: self.check_required_fields())
        ttk.Entry(duty_group_frame, textvariable=self.duty_group, width=20, font=DEFAULT_FONT).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Черговий викладач з автозаповненням
        duty_teacher_frame = ttk.Frame(duty_frame)
        duty_teacher_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        ttk.Label(duty_teacher_frame, text="Черговий викладач:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.duty_teacher = tk.StringVar()
        # Додаємо перевірку полів при зміні чергового викладача
        self.duty_teacher.trace_add("write", lambda *args: self.check_required_fields())
        self.duty_teacher_entry = ttk.Entry(duty_teacher_frame, textvariable=self.duty_teacher, width=20, font=DEFAULT_FONT)
        self.duty_teacher_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Створюємо список для автозаповнення
        self.teacher_listbox = tk.Listbox(duty_teacher_frame, width=30, height=5, font=DEFAULT_FONT)
        self.teacher_listbox.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.teacher_listbox.grid_remove()  # Спочатку приховуємо
        
        # Прив'язуємо події для автозаповнення
        self.duty_teacher.trace("w", lambda name, index, mode: self.update_teacher_list(
                                                                entry_var=self.duty_teacher, 
                                                                listbox=self.teacher_listbox))
        self.teacher_listbox.bind("<<ListboxSelect>>", lambda event: self.on_teacher_select(event))
        
        # Черговий викладач в гуртожитку з автозаповненням
        dorm_teacher_frame = ttk.Frame(duty_frame)
        dorm_teacher_frame.pack(fill=tk.X, padx=PADDING, pady=PADDING)
        
        ttk.Label(dorm_teacher_frame, text="Черговий викладач в гуртожитку:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.dorm_teacher = tk.StringVar()
        # Додаємо перевірку полів при зміні викладача гуртожитку
        self.dorm_teacher.trace_add("write", lambda *args: self.check_required_fields())
        self.dorm_teacher_entry = ttk.Entry(dorm_teacher_frame, textvariable=self.dorm_teacher, width=20, font=DEFAULT_FONT)
        self.dorm_teacher_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Створюємо список для автозаповнення
        self.dorm_teacher_listbox = tk.Listbox(dorm_teacher_frame, width=30, height=5, font=DEFAULT_FONT)
        self.dorm_teacher_listbox.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.dorm_teacher_listbox.grid_remove()  # Спочатку приховуємо
        
        # Прив'язуємо події для автозаповнення
        self.dorm_teacher.trace("w", lambda name, index, mode: self.update_teacher_list(
                                                                    entry_var=self.dorm_teacher, 
                                                                    listbox=self.dorm_teacher_listbox))
        self.dorm_teacher_listbox.bind("<<ListboxSelect>>", 
                                       lambda event: self.on_teacher_select(event, 
                                                                           self.dorm_teacher, 
                                                                           self.dorm_teacher_listbox))
        
        # Додаємо пустий фрейм для відступу внизу
        spacer_frame = ttk.Frame(left_frame, height=20)
        spacer_frame.pack(fill=tk.X, pady=10)
        
        # Права частина - порожній фрейм для балансу інтерфейсу
        right_frame.pack_propagate(False)  # Забороняємо зменшення фрейму
        
        # Створюємо контейнер для слайдера (початково невидимий)
        # Додаємо його в кінці, щоб він був поверх інших елементів
        self.slider_container = tk.Frame(self, relief=tk.RAISED, borderwidth=2, bg=APP_BG_COLOR)
        self.slider_container.place(x=self.window_width, y=0, width=self.slider_width, height=700)  # Збільшено висоту до 700
        
        # Кнопка зі стрілкою для відкриття/закриття слайдера
        self.toggle_button = tk.Button(self, text="<", width=2, font=BUTTON_FONT, bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, command=self.toggle_slider)
        self.toggle_button.place(x=self.window_width - 20, y=350 - 15, width=20, height=30)  # Змінено позицію кнопки по вертикалі
        
        # Створюємо вміст слайдера
        self.create_slider_content()
        
        # Створюємо дерево файлів
        self.update_treeview()
        
        # Перевіряємо стан полів для активації/деактивації кнопки "Бланк замін"
        self.check_required_fields()
    
    def update_date_info(self, event=None):
        # Отримуємо поточну дату
        selected_date = self.date_entry.get_date()
        
        # Форматуємо дату словами
        months = ["січня", "лютого", "березня", "квітня", "травня", "червня", 
                 "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]
        date_text_str = f"{selected_date.day} {months[selected_date.month-1]} {selected_date.year} року"
        self.date_text.set(date_text_str)
        
        # Визначаємо день тижня
        weekdays = ["понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота", "неділя"]
        weekday_idx = selected_date.weekday()  # 0 - понеділок, 6 - неділя
        self.weekday_text.set(weekdays[weekday_idx])
        
        # Перевіряємо чи всі поля заповнені
        self.check_required_fields()
    
    def update_teacher_list(self, name=None, index=None, mode=None, entry_var=None, listbox=None):
        if entry_var is None:
            entry_var = self.duty_teacher
            listbox = self.teacher_listbox
            
        typed = entry_var.get().lower()
        print(f"Пошук викладачів за текстом: '{typed}'")
        print(f"Кількість викладачів у списку: {len(self.teachers_list)}")
        
        if typed == '':
            listbox.grid_remove()
        else:
            listbox.grid()
            listbox.delete(0, tk.END)
            
            # Перевіряємо, що список викладачів не порожній
            if not self.teachers_list:
                print("Список викладачів порожній!")
                listbox.insert(tk.END, "Список викладачів порожній")
                return
            
            # Виводимо перші 5 викладачів для відлагодження
            print(f"Перші 5 викладачів: {self.teachers_list[:5] if len(self.teachers_list) >= 5 else self.teachers_list}")
                
            found_count = 0
            for teacher in self.teachers_list:
                try:
                    # Спробуємо перетворити на рядок та порівняти
                    teacher_str = str(teacher)
                    if typed in teacher_str.lower():
                        listbox.insert(tk.END, teacher_str)
                        found_count += 1
                except Exception as e:
                    print(f"Помилка при обробці викладача: {e}")
            
            print(f"Знайдено {found_count} викладачів, що відповідають пошуку")
    
    def on_teacher_select(self, event, entry_var=None, listbox=None):
        if entry_var is None:
            entry_var = self.duty_teacher
            listbox = self.teacher_listbox
            
        if listbox.curselection():
            selected = listbox.get(listbox.curselection())
            entry_var.set(selected)
            listbox.grid_remove()
    
    def check_required_fields(self):
        """Перевіряє чи всі обов'язкові поля заповнені"""
        # Перевіряємо, чи всі необхідні атрибути існують
        required_attrs = ['date_text', 'weekday_text', 'week_type', 'duty_group', 
                         'duty_teacher', 'dorm_teacher', 'academic_year_var', 'replacement_btn']
        
        # Перевіряємо, чи всі атрибути існують
        for attr in required_attrs:
            if not hasattr(self, attr):
                return False  # Якщо якийсь атрибут відсутній, повертаємо False
        
        # Отримуємо значення полів
        date_text = self.date_text.get()
        weekday = self.weekday_text.get()
        week_type = self.week_type.get()
        duty_group = self.duty_group.get()
        duty_teacher = self.duty_teacher.get()
        dorm_teacher = self.dorm_teacher.get()
        academic_year = self.academic_year_var.get()
        
        # Перевіряємо чи всі поля заповнені
        all_fields_filled = all([
            date_text, weekday, week_type, duty_group, 
            duty_teacher, dorm_teacher, academic_year
        ])
        
        # Оновлюємо стан кнопки відповідно до результату перевірки
        if all_fields_filled:
            self.replacement_btn.config(state=tk.NORMAL, bg=BUTTON_BG_COLOR)
            self.replacement_btn.config(cursor="hand2")
        else:
            self.replacement_btn.config(state=tk.DISABLED, bg="#a0a0a0")
            self.replacement_btn.config(cursor="arrow")
        
        return all_fields_filled
    
    def validate_duty_group(self, group_name):
        """Базова перевірка назви чергової групи"""
        if not group_name or not group_name.strip():
            return False, "Назва чергової групи не може бути порожньою"
        
        # Дозволяємо будь-яку назву групи, що не є порожньою
        return True, "Чергова група прийнята"
    
    def open_handbook(self):
        handbook = HandbookForm(self)
        handbook.grab_set()  # Make window modal
    
    def open_replacement_form(self):
        # Перевіряємо чи всі поля заповнені
        if not self.check_required_fields():
            messagebox.showwarning("Недостатньо даних", 
                                 "Будь ласка, заповніть усі обов'язкові поля перед створенням бланку замін.")
            return
        
        # Отримуємо значення з головної форми
        date_text = self.date_text.get()
        weekday = self.weekday_text.get()
        week_type = self.week_type.get()
        duty_group = self.duty_group.get()
        duty_teacher = self.duty_teacher.get()
        dorm_teacher = self.dorm_teacher.get()
        
        # Валідація чергової групи
        group_valid, group_message = self.validate_duty_group(duty_group)
        if not group_valid:
            messagebox.showerror("Помилка валідації чергової групи", group_message)
            return
        else:
            print(f"[OK] {group_message}")
        
        # Отримуємо навчальний рік
        academic_year = self.academic_year_var.get()
        
        # Створюємо та показуємо форму замін
        replacement_window = tk.Toplevel(self)
        replacement_window.title("Форма замін")
        replacement_window.geometry("1000x650")  # Зменшуємо висоту з 700 до 650
        
        # Додаємо стиль до вікна
        replacement_window.configure(bg=APP_BG_COLOR)
        
        # Створюємо форму замін як фрейм всередині вікна
        replacement_form = ReplacementForm(replacement_window, date_text, weekday, week_type, 
                                           duty_group, duty_teacher, dorm_teacher, 
                                           self.replacements_dir, academic_year)
        replacement_form.pack(fill=tk.BOTH, expand=True)
        
        replacement_window.grab_set()  # Make window modal
    
    def refresh_database_data(self):
        """Оновлення даних з бази даних без перезапуску програми"""
        try:
            # Показуємо повідомлення про початок оновлення
            messagebox.showinfo("Оновлення даних", "Оновлення даних з бази даних...")
            
            # Створюємо DatabaseManager для оновлення даних
            db = DatabaseManager()
            
            # Оновлюємо дані в базі даних
            if db.refresh_data():
                # Оновлюємо список викладачів
                old_teachers_count = len(self.teachers_list)
                self.teachers_list = db.get_teachers()
                
                # Оновлюємо віджети форми
                self.update_main_form_widgets()
                
                # Показуємо інформацію про зміни
                changes = []
                if len(self.teachers_list) != old_teachers_count:
                    changes.append(f"Викладачі: {old_teachers_count} → {len(self.teachers_list)}")
                
                # Отримуємо кількість відділень та груп
                departments = db.get_departments()
                all_groups = db.get_all_groups()
                total_groups = sum(len(groups) for groups in all_groups.values())
                
                changes.append(f"Відділення: {len(departments)}")
                changes.append(f"Групи: {total_groups}")
                
                if changes:
                    change_text = "Поточна структура даних:\n" + "\n".join(changes)
                else:
                    change_text = "Структура даних оновлена"
                
                messagebox.showinfo("Оновлення завершено", 
                                  f"Дані успішно оновлено з бази даних!\n\n{change_text}")
            else:
                messagebox.showwarning("Помилка оновлення", 
                                     "Не вдалося оновити дані з бази даних.\n"
                                     "Перевірте підключення до бази даних.")
        
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при оновленні даних: {e}")
    
    def update_main_form_widgets(self):
        """Оновлення віджетів головної форми після оновлення даних"""
        try:
            # Оновлюємо автозаповнення для чергового викладача
            current_teacher = self.duty_teacher.get()
            
            # Якщо поточний викладач більше не існує, очищуємо поле
            if current_teacher and current_teacher not in self.teachers_list:
                self.duty_teacher.set("")
            
            # Оновлюємо автозаповнення для чергового викладача в гуртожитку
            current_dorm_teacher = self.dorm_teacher.get()
            
            # Якщо поточний викладач більше не існує, очищуємо поле
            if current_dorm_teacher and current_dorm_teacher not in self.teachers_list:
                self.dorm_teacher.set("")
            
            print("Віджети головної форми успішно оновлено")
            
        except Exception as e:
            print(f"Помилка при оновленні віджетів головної форми: {e}")

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
