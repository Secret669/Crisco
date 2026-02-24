import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import os
from db_manager import DatabaseManager

# Константи для стилю
APP_BG_COLOR = "#f0f0f0"  # Світло-сірий фон
HEADER_BG_COLOR = "#4a6984"  # Темно-синій для заголовків
HEADER_FG_COLOR = "white"  # Білий текст для заголовків
BUTTON_BG_COLOR = "#1a365d"  # Темно-синій для кнопок
BUTTON_FG_COLOR = "#ffffff"  # Білий текст для кнопок
HIGHLIGHT_COLOR = "#2ecc71"  # Яскраво-зелений для виділення
FONT_FAMILY = "Arial"  # Шрифт для всього додатку
DEFAULT_FONT = (FONT_FAMILY, 10)  # Звичайний шрифт
HEADER_FONT = (FONT_FAMILY, 12, "bold")  # Шрифт для заголовків
BUTTON_FONT = (FONT_FAMILY, 12, "bold")  # Шрифт для кнопок
PADDING = 10  # Стандартний відступ

class HandbookForm(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Довідник")
        self.geometry("1000x700")  # Збільшено розмір вікна
        self.minsize(1000, 700)  # Збільшено мінімальний розмір
        
        # Застосовуємо стиль до всього вікна
        self.configure(bg=APP_BG_COLOR)
        self.style = ttk.Style()
        self.style.configure('TFrame', background=APP_BG_COLOR)
        self.style.configure('TLabel', background=APP_BG_COLOR, font=DEFAULT_FONT)
        self.style.configure('TButton', font=BUTTON_FONT)
        
        # Створюємо кастомні стилі
        self.style.configure('Header.TLabel', background=HEADER_BG_COLOR, foreground=HEADER_FG_COLOR, font=HEADER_FONT, padding=PADDING)
        self.style.configure('Header.TFrame', background=HEADER_BG_COLOR)
        self.style.configure('Treeview', font=DEFAULT_FONT, rowheight=25)
        self.style.configure('Treeview.Heading', font=BUTTON_FONT)
        
        # Connect to database
        self.db_path = os.path.abspath("dataBase.mdb")
        self.conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path}'
        
        # Current active table
        self.current_table = None
        self.current_data = []
        
        # Назви колонок для різних таблиць
        self.column_mappings = {
            "department": {"id": "ID", "name": "Name"},
            "discipline": {"id": "ID_discpline", "name": "Name"},
            "groups": {"id": "ID", "name": "Name"},
            "teachers": {"id": "ID", "name": "PIB"}
        }
        
        self.create_widgets()
    
    def create_widgets(self):
        # Створюємо заголовок вікна
        header_frame = ttk.Frame(self, style='Header.TFrame')
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        header_label = ttk.Label(header_frame, text="Довідник", style='Header.TLabel')
        header_label.pack(pady=PADDING)
        
        # Головний контейнер
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=PADDING, pady=PADDING)
        
        # Ліва панель з кнопками (зменшуємо ширину, щоб більше місця залишити для правої панелі)
        left_panel = ttk.Frame(main_frame, width=200)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=PADDING, pady=PADDING)
        
        # Створюємо фрейм для кнопок
        buttons_label = ttk.Label(left_panel, text="Таблиці", font=HEADER_FONT)
        buttons_label.pack(anchor=tk.W, pady=(0, PADDING))
        
        self.buttons_frame = ttk.Frame(left_panel)
        self.buttons_frame.pack(fill=tk.Y, expand=True)
        
        # Створюємо кнопки для основних таблиць з новим стилем
        dept_btn = tk.Button(self.buttons_frame, text="Відділення", 
                          bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                          relief=tk.RAISED, borderwidth=2, padx=10, pady=5,
                          command=lambda: self.load_table("department"))
        dept_btn.pack(fill=tk.X, pady=PADDING)
        
        teachers_btn = tk.Button(self.buttons_frame, text="Викладачі", 
                              bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                              relief=tk.RAISED, borderwidth=2, padx=10, pady=5,
                              command=lambda: self.load_table("teachers"))
        teachers_btn.pack(fill=tk.X, pady=PADDING)
        
        disc_btn = tk.Button(self.buttons_frame, text="Дисципліни", 
                           bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                           relief=tk.RAISED, borderwidth=2, padx=10, pady=5,
                           command=lambda: self.load_table("discipline"))
        disc_btn.pack(fill=tk.X, pady=PADDING)
        groups_btn = tk.Button(self.buttons_frame, text="Групи", 
                            bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                            relief=tk.RAISED, borderwidth=2, padx=10, pady=5,
                            command=lambda: self.load_table("groups"))
        groups_btn.pack(fill=tk.X, pady=PADDING)
        
        # Права панель з списком та елементами керування (збільшуємо розмір)
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=PADDING, pady=PADDING)
        
        # Заголовок правої панелі
        right_header = ttk.Label(right_panel, text="Записи", font=HEADER_FONT)
        right_header.pack(anchor=tk.W, pady=(0, PADDING))
        
        # Поле пошуку
        search_frame = ttk.Frame(right_panel)
        search_frame.pack(fill=tk.X, pady=PADDING)
        
        ttk.Label(search_frame, text="Пошук:", font=DEFAULT_FONT).pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_list)
        ttk.Entry(search_frame, textvariable=self.search_var, width=40, font=DEFAULT_FONT).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Список елементів
        list_frame = ttk.Frame(right_panel)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=PADDING)
        
        # Збільшуємо розмір шрифту для елементів списку
        list_font = (FONT_FAMILY, 12)  # Збільшений шрифт для списку
        self.item_listbox = tk.Listbox(list_frame, width=50, height=20, font=list_font, 
                                     bg='white', selectbackground=BUTTON_BG_COLOR, selectforeground='white')
        self.item_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Скролбар для списку
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.item_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.item_listbox.config(yscrollcommand=scrollbar.set)
        
        # Кнопки для операцій CRUD
        button_frame = ttk.Frame(right_panel)
        button_frame.pack(fill=tk.X, pady=PADDING)
        
        # Кнопка додавання
        add_btn = tk.Button(button_frame, text="Додати", 
                           bg=HIGHLIGHT_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                           relief=tk.RAISED, borderwidth=2, padx=15, pady=8,  # Збільшено відступи
                           command=self.add_item)
        add_btn.pack(side=tk.LEFT, padx=10)  # Збільшено відступ між кнопками
        
        # Кнопка редагування
        edit_btn = tk.Button(button_frame, text="Редагувати", 
                            bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                            relief=tk.RAISED, borderwidth=2, padx=15, pady=8,  # Збільшено відступи
                            command=self.edit_item)
        edit_btn.pack(side=tk.LEFT, padx=10)  # Збільшено відступ між кнопками
        
        # Кнопка видалення
        delete_btn = tk.Button(button_frame, text="Видалити", 
                              bg="#e74c3c", fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                              relief=tk.RAISED, borderwidth=2, padx=15, pady=8,  # Збільшено відступи
                              command=self.delete_item)
        delete_btn.pack(side=tk.LEFT, padx=10)  # Збільшено відступ між кнопками
    
    def load_table(self, table_name):
        self.current_table = table_name
        self.search_var.set("")  # Clear search field
        self.refresh_list()
    
    def refresh_list(self):
        if not self.current_table:
            return
            
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            
            # Очищаємо список перед завантаженням нових даних
            self.item_listbox.delete(0, tk.END)
            
            # Отримуємо назви колонок для поточної таблиці
            id_col_name = self.column_mappings.get(self.current_table, {}).get("id", "ID")
            name_col_name = self.column_mappings.get(self.current_table, {}).get("name", "Name")
            
            # Виконуємо запит
            cursor.execute(f"SELECT * FROM [{self.current_table}]")
            rows = cursor.fetchall()
            columns = [column[0] for column in cursor.description]
            
            print(f"Стовпці таблиці {self.current_table}: {columns}")
            
            # Знаходимо індекси колонок ID та Name
            id_col = None
            name_col = None
            
            for i, col in enumerate(columns):
                if col == id_col_name:
                    id_col = i
                elif col == name_col_name:
                    name_col = i
            
            # Використовуємо знайдені індекси або значення за замовчуванням
            if id_col is None:
                id_col = 0
            if name_col is None:
                name_col = 1
            
            # Друкуємо інформацію про колонки
            print(f"Використовуємо колонку ID: {id_col} ({columns[id_col] if id_col < len(columns) else 'unknown'})")
            print(f"Використовуємо колонку Name: {name_col} ({columns[name_col] if name_col < len(columns) else 'unknown'})")
            
            # Створюємо дані та заповнюємо список
            self.current_data = []
            self.filtered_data = []  # Очищаємо відфільтровані дані
            
            for row in rows:
                if len(row) > max(id_col, name_col):
                    item_id = row[id_col]
                    item_name = row[name_col] if row[name_col] else ""
                    self.current_data.append((item_id, item_name))
                    self.item_listbox.insert(tk.END, item_name)
            
            conn.close()
        except Exception as e:
            print(f"Помилка при завантаженні даних: {e}")
            messagebox.showerror("Помилка", f"Помилка при завантаженні даних: {e}")
    
    def filter_list(self, *args):
        search_text = self.search_var.get().lower()
        
        self.item_listbox.delete(0, tk.END)
        
        # Створюємо новий список для відфільтрованих даних
        self.filtered_data = []
        
        for item in self.current_data:
            # Для всіх таблиць просто шукаємо в назві
            if search_text in item[1].lower():
                self.filtered_data.append(item)
                self.item_listbox.insert(tk.END, item[1])
    
    def add_item(self):
        if not self.current_table:
            messagebox.showinfo("Інформація", "Спочатку виберіть категорію")
            return
            
        # Створюємо діалог для додавання запису
        dialog = ItemDialog(self, "Додати", self.current_table)
        
        if dialog.result:
            try:
                conn = pyodbc.connect(self.conn_str)
                cursor = conn.cursor()
                
                # Отримуємо назву колонки Name для поточної таблиці
                name_col_name = self.column_mappings.get(self.current_table, {}).get("name", "Name")
                
                if self.current_table == "groups":
                    # Отримуємо ID відділення
                    cursor.execute("SELECT ID FROM [department] WHERE Name = ?", (dialog.result[1],))
                    dept_row = cursor.fetchone()
                    if dept_row and len(dept_row) > 0:
                        dept_id = dept_row[0]
                        
                        # Додаємо запис з відділенням
                        cursor.execute(f"INSERT INTO [{self.current_table}] ({name_col_name}, [Number Of Department]) VALUES (?, ?)", 
                                      (dialog.result[0], dept_id))
                    else:
                        messagebox.showerror("Помилка", "Відділення не знайдено")
                else:
                    # Додаємо запис без відділення
                    cursor.execute(f"INSERT INTO [{self.current_table}] ({name_col_name}) VALUES (?)", 
                                  (dialog.result[0],))
                
                conn.commit()
                conn.close()
                
                # Оновлюємо список
                self.refresh_list()
                
                # Автоматично оновлюємо дані для всіх важливих таблиць
                if self.current_table in ["department", "groups", "teachers", "audiences", "discpline", "discipline", "disciplines"]:
                    self.refresh_main_app_data()
                    table_names = {
                        "department": "відділення",
                        "groups": "групу", 
                        "teachers": "викладача",
                        "audiences": "аудиторію",
                        "discpline": "дисципліну",
                        "discipline": "дисципліну",
                        "disciplines": "дисципліну"
                    }
                    item_name = table_names.get(self.current_table, "запис")
                    messagebox.showinfo("Інформація", f"Новий {item_name} додано та дані автоматично оновлено!")
                
            except Exception as e:
                print(f"Помилка при додаванні запису: {e}")
                messagebox.showerror("Помилка", f"Помилка при додаванні запису: {e}")
    
    def edit_item(self):
        if not self.current_table:
            return
            
        selection = self.item_listbox.curselection()
        if not selection:
            messagebox.showinfo("Інформація", "Виберіть запис для редагування")
            return
        
        # Отримуємо вибраний індекс
        selected_index = selection[0]
        
        # Використовуємо відфільтрований список, якщо він є
        data_source = self.filtered_data if hasattr(self, 'filtered_data') and self.filtered_data else self.current_data
        
        # Перевіряємо, чи є дані для цього індексу
        if selected_index >= len(data_source):
            messagebox.showinfo("Інформація", "Не вдалося знайти вибраний запис")
            return
            
        # Отримуємо ID та назву вибраного запису
        item_id = data_source[selected_index][0]
        item_name = data_source[selected_index][1]
        
        print(f"Редагування: вибрано запис №{selected_index}, ID={item_id}, назва='{item_name}'")
        
        # Створюємо діалог для редагування запису
        if self.current_table == "groups":
            try:
                # Отримуємо назву відділення для групи
                conn = pyodbc.connect(self.conn_str)
                cursor = conn.cursor()
                
                # Отримуємо дані про групу з бази даних
                cursor.execute(f"SELECT * FROM [{self.current_table}] WHERE ID = {item_id}")
                row = cursor.fetchone()
                
                # Отримуємо назву групи безпосередньо з бази даних
                # Це гарантує, що ми використовуємо правильну назву групи
                group_name = ""
                if row and len(row) > 1:  # Перевіряємо, що є колонка з назвою
                    # Знаходимо індекс колонки Name
                    columns = [column[0] for column in cursor.description]
                    name_col = None
                    for i, col in enumerate(columns):
                        if col == "Name":
                            name_col = i
                            break
                    
                    if name_col is not None and name_col < len(row):
                        group_name = row[name_col] if row[name_col] else ""
                    else:
                        group_name = row[1] if row[1] else ""  # За замовчуванням беремо другу колонку
                
                dept_name = ""
                if row:
                    # Знаходимо колонку з ID відділення (третя колонка)
                    dept_id = row[2] if len(row) > 2 else None
                    
                    # Якщо є ID відділення, отримуємо його назву
                    if dept_id:
                        cursor.execute(f"SELECT Name FROM [department] WHERE ID = {dept_id}")
                        dept_row = cursor.fetchone()
                        if dept_row and len(dept_row) > 0:
                            dept_name = dept_row[0]
                
                conn.close()
                
                # Використовуємо назву групи з бази даних, а не з фільтрованого списку
                print(f"Використовуємо назву групи з бази даних: '{group_name}'")
                dialog = ItemDialog(self, "Редагувати", self.current_table, initial_values=(group_name, dept_name))
            except Exception as e:
                print(f"Помилка при отриманні даних для редагування: {e}")
                dialog = ItemDialog(self, "Редагувати", self.current_table, initial_values=(item_name, ""))
        else:
            # Для інших таблиць просто передаємо назву
            dialog = ItemDialog(self, "Редагувати", self.current_table, initial_values=(item_name,))
        
        if dialog.result:
            try:
                conn = pyodbc.connect(self.conn_str)
                cursor = conn.cursor()
                
                # Отримуємо назву колонки Name для поточної таблиці
                name_col_name = self.column_mappings.get(self.current_table, {}).get("name", "Name")
                id_col_name = self.column_mappings.get(self.current_table, {}).get("id", "ID")
                
                print(f"Редагування запису: таблиця={self.current_table}, ID={item_id}")
                print(f"Нові дані: {dialog.result}")
                print(f"Назва колонки для імені: {name_col_name}")
                print(f"Назва колонки для ID: {id_col_name}")
                
                # Використовуємо прямий SQL-запит без параметрів
                if self.current_table == "groups" and len(dialog.result) > 1:
                    # Отримуємо ID відділення
                    cursor.execute(f"SELECT ID FROM [department] WHERE Name = '{dialog.result[1]}'")
                    dept_row = cursor.fetchone()
                    
                    if dept_row and len(dept_row) > 0:
                        dept_id = dept_row[0]
                        
                        # Оновлюємо запис з відділенням
                        query = f"UPDATE [{self.current_table}] SET {name_col_name} = '{dialog.result[0]}', [Number Of Department] = {dept_id} WHERE {id_col_name} = {item_id}"
                        print(f"SQL-запит: {query}")
                        cursor.execute(query)
                    else:
                        messagebox.showerror("Помилка", "Відділення не знайдено")
                else:
                    # Оновлюємо запис без відділення
                    query = f"UPDATE [{self.current_table}] SET {name_col_name} = '{dialog.result[0]}' WHERE {id_col_name} = {item_id}"
                    print(f"SQL-запит: {query}")
                    cursor.execute(query)
                
                conn.commit()
                conn.close()
                
                # Оновлюємо список
                self.refresh_list()
                
                # Автоматично оновлюємо дані для всіх важливих таблиць
                if self.current_table in ["department", "groups", "teachers", "audiences", "discpline", "discipline", "disciplines"]:
                    self.refresh_main_app_data()
                    table_names = {
                        "department": "відділення",
                        "groups": "групи", 
                        "teachers": "викладача",
                        "audiences": "аудиторії",
                        "discpline": "дисципліни",
                        "discipline": "дисципліни",
                        "disciplines": "дисципліни"
                    }
                    item_name = table_names.get(self.current_table, "запису")
                    messagebox.showinfo("Інформація", f"Дані {item_name} змінено та автоматично оновлено!")
                
            except Exception as e:
                print(f"Помилка при оновленні запису: {e}")
                messagebox.showerror("Помилка", f"Помилка при оновленні запису: {e}")
    
    def delete_item(self):
        if not self.current_table:
            return
            
        selection = self.item_listbox.curselection()
        if not selection:
            messagebox.showinfo("Інформація", "Виберіть запис для видалення")
            return
        
        # Отримуємо вибраний індекс
        selected_index = selection[0]
        
        # Використовуємо відфільтрований список, якщо він є
        data_source = self.filtered_data if hasattr(self, 'filtered_data') and self.filtered_data else self.current_data
        
        # Перевіряємо, чи є дані для цього індексу
        if selected_index >= len(data_source):
            messagebox.showinfo("Інформація", "Не вдалося знайти вибраний запис")
            return
            
        # Отримуємо ID та назву вибраного запису
        item_id = data_source[selected_index][0]
        item_name = data_source[selected_index][1]
        
        # Отримуємо назву колонки ID для поточної таблиці
        id_col_name = self.column_mappings.get(self.current_table, {}).get("id", "ID")
        
        print(f"Видалення: вибрано запис №{selected_index}, ID={item_id}, назва='{item_name}'")
        
        # Підтвердження видалення
        if not messagebox.askyesno("Підтвердження", f"Ви дійсно хочете видалити '{item_name}'?"):
            return
            
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            
            # Видаляємо запис
            query = f"DELETE FROM [{self.current_table}] WHERE {id_col_name} = {item_id}"
            print(f"SQL-запит: {query}")
            cursor.execute(query)
            
            conn.commit()
            conn.close()
            
            # Оновлюємо список
            self.refresh_list()
            
            # Автоматично оновлюємо дані для всіх важливих таблиць
            if self.current_table in ["department", "groups", "teachers", "audiences", "discpline", "discipline", "disciplines"]:
                self.refresh_main_app_data()
                table_names = {
                    "department": "відділення",
                    "groups": "групу", 
                    "teachers": "викладача",
                    "audiences": "аудиторію",
                    "discpline": "дисципліну",
                    "discipline": "дисципліну",
                    "disciplines": "дисципліну"
                }
                item_name = table_names.get(self.current_table, "запис")
                messagebox.showinfo("Інформація", f"{item_name.capitalize()} видалено та дані автоматично оновлено!")
            
        except Exception as e:
            print(f"Помилка при видаленні запису: {e}")
            messagebox.showerror("Помилка", f"Помилка при видаленні запису: {e}")
    
    def refresh_main_app_data(self):
        """Оновлення даних у головній програмі після змін у довіднику"""
        try:
            # Створюємо DatabaseManager для оновлення даних
            db = DatabaseManager()
            
            # Оновлюємо дані в базі даних
            if db.refresh_data():
                print("Дані головної програми успішно оновлено з довідника")
                
                # Якщо головна програма має метод оновлення, викликаємо його
                if hasattr(self.parent, 'refresh_database_data'):
                    # Оновлюємо дані у головній формі без показу повідомлень
                    try:
                        # Оновлюємо список викладачів
                        self.parent.teachers_list = db.get_teachers()
                        # Оновлюємо віджети головної форми
                        self.parent.update_main_form_widgets()
                        print("Віджети головної форми оновлено з довідника")
                    except Exception as e:
                        print(f"Помилка при оновленні віджетів головної форми: {e}")
            else:
                print("Не вдалося оновити дані головної програми")
                
        except Exception as e:
            print(f"Помилка при оновленні даних головної програми: {e}")

class ItemDialog(tk.Toplevel):
    def __init__(self, parent, action, table_name, initial_values=None):
        super().__init__(parent)
        self.parent = parent
        self.title(f"{action} запис")
        self.geometry("450x250")  # Збільшуємо розмір вікна
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Застосовуємо стиль до вікна
        self.configure(bg=APP_BG_COLOR)
        
        self.result = None
        self.table_name = table_name
        
        # Створюємо заголовок вікна
        header_frame = tk.Frame(self, bg=HEADER_BG_COLOR)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        header_label = tk.Label(header_frame, text=f"{action} запис", 
                               font=HEADER_FONT, bg=HEADER_BG_COLOR, fg=HEADER_FG_COLOR)
        header_label.pack(pady=PADDING)
        
        # Створюємо головний фрейм
        main_frame = tk.Frame(self, bg=APP_BG_COLOR, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Поле для назви
        name_label = tk.Label(main_frame, text="Назва:", font=DEFAULT_FONT, bg=APP_BG_COLOR)
        name_label.grid(row=0, column=0, sticky=tk.W, pady=10, padx=5)
        
        self.name_var = tk.StringVar()
        if initial_values and len(initial_values) > 0:
            self.name_var.set(initial_values[0])
        
        name_entry = tk.Entry(main_frame, textvariable=self.name_var, width=30, font=DEFAULT_FONT,
                             relief=tk.SOLID, borderwidth=1)
        name_entry.grid(row=0, column=1, sticky=tk.W, pady=10, padx=5)
        name_entry.focus_set()  # Фокус на поле вводу
        
        # Поле для відділення (тільки для груп)
        self.dept_var = None
        if table_name == "groups":
            dept_label = tk.Label(main_frame, text="Відділення:", font=DEFAULT_FONT, bg=APP_BG_COLOR)
            dept_label.grid(row=1, column=0, sticky=tk.W, pady=10, padx=5)
            
            self.dept_var = tk.StringVar()
            if initial_values and len(initial_values) > 1:
                self.dept_var.set(initial_values[1])
                
            # Отримуємо список відділень з бази даних
            departments = self.get_departments()
            
            # Створюємо випадаючий список
            dept_combo = ttk.Combobox(main_frame, textvariable=self.dept_var, values=departments, 
                                     width=28, font=DEFAULT_FONT)
            dept_combo.grid(row=1, column=1, sticky=tk.W, pady=10, padx=5)
            
            if departments and (not initial_values or not initial_values[1]):
                dept_combo.current(0)
        
        # Кнопки
        button_frame = tk.Frame(main_frame, bg=APP_BG_COLOR)
        button_frame.grid(row=3, column=0, columnspan=2, pady=15)
        
        # Кнопка OK
        ok_btn = tk.Button(button_frame, text="OK", command=self.ok_clicked,
                          bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                          relief=tk.RAISED, borderwidth=2, padx=20, pady=5)
        ok_btn.pack(side=tk.LEFT, padx=10)
        
        # Кнопка Скасувати
        cancel_btn = tk.Button(button_frame, text="Скасувати", command=self.cancel_clicked,
                              bg="#e74c3c", fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
                              relief=tk.RAISED, borderwidth=2, padx=10, pady=5)
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        self.wait_window(self)
    
    def get_departments(self):
        """Get list of departments from database"""
        departments = []
        try:
            conn = pyodbc.connect(self.parent.conn_str)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM [department]")
            rows = cursor.fetchall()
            columns = [column[0] for column in cursor.description]
            
            # Find Name column
            name_col = 1  # Default to second column
            for i, col in enumerate(columns):
                if col.lower() == 'name':
                    name_col = i
                    break
            
            departments = [row[name_col] for row in rows if row[name_col]]
            conn.close()
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при отриманні списку відділень: {e}")
        return departments
    
    def ok_clicked(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Попередження", "Назва не може бути порожньою")
            return
        
        if self.table_name == "groups" and self.dept_var:
            dept = self.dept_var.get().strip()
            if not dept:
                messagebox.showwarning("Попередження", "Виберіть відділення")
                return
            self.result = (name, dept)
        else:
            self.result = (name,)
        
        self.destroy()
    
    def cancel_clicked(self):
        self.result = None
        self.destroy()
