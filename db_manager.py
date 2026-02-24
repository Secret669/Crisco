import os
import sys
import pyodbc
import sqlite3
from tkinter import messagebox

class DatabaseManager:
    """
    Клас для управління підключенням до бази даних та виконання запитів.
    Реалізує патерн Singleton для забезпечення єдиного підключення до бази даних.
    """
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DatabaseManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
            
        self._initialized = True
        self.conn = None
        self.cursor = None
        
        # Кеш для часто використовуваних даних
        self._cache = {
            'departments': None,
            'groups': None,
            'teachers': None,
            'audiences': None,
            'disciplines': None,
            'department_structure': None
        }
        self._cache_valid = False
        
        # Визначаємо шлях до бази даних відносно EXE-файлу або скрипта
        try:
            # Спочатку пробуємо отримати шлях до EXE-файлу (для скомпільованої програми)
            if getattr(sys, 'frozen', False):
                # Якщо програма скомпільована в EXE
                application_path = os.path.dirname(sys.executable)
            else:
                # Якщо програма запущена з Python
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            self.db_path = os.path.join(application_path, "dataBase.mdb")
        except Exception as e:
            print(f"Помилка при визначенні шляху до бази даних: {e}")
            # Використовуємо старий метод як запасний варіант
            self.db_path = os.path.abspath("dataBase.mdb")
        
        print(f"Шлях до бази даних: {self.db_path}")
        self.conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path}'
        self.connect()
    
    def connect(self):
        """Підключення до бази даних"""
        # Спочатку перевіряємо, чи існує файл бази даних
        if not os.path.exists(self.db_path):
            # Спробуємо знайти базу даних в інших можливих місцях
            possible_paths = [
                os.path.join(os.path.dirname(self.db_path), "dataBase.accdb"),  # Спробуємо .accdb в тій же папці
                os.path.abspath("dataBase.mdb"),  # Спробуємо в поточній папці
                os.path.abspath("dataBase.accdb"),  # Спробуємо .accdb в поточній папці
                os.path.join(os.path.expanduser("~"), "Desktop", "Crisco", "dataBase.mdb"),  # На робочому столі
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    self.db_path = path
                    self.conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path}'
                    print(f"Знайдено базу даних за шляхом: {self.db_path}")
                    break
            else:
                # Якщо базу даних не знайдено, повідомляємо про це
                messagebox.showwarning("Попередження", 
                                      "Файл бази даних не знайдено. \n"
                                      "Програма буде працювати з тестовими даними.")
                print(f"Файл бази даних не знайдено: {self.db_path}")
        
        # Спробуємо підключитися до бази даних за допомогою pyodbc
        try:
            self.conn = pyodbc.connect(self.conn_str)
            self.cursor = self.conn.cursor()
            print("Успішне підключення до бази даних")
            return True
        except Exception as e:
            print(f"Помилка підключення до бази даних: {e}")
            
            # Якщо не вдалося підключитися до бази даних, повідомляємо про це тільки один раз
            if "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}" in str(e):
                messagebox.showerror("Помилка", 
                                   "Драйвер Microsoft Access не знайдено. \n"
                                   "Перевірте, чи встановлено Microsoft Office або драйвер ODBC для Access.\n"
                                   "Програма буде працювати з тестовими даними.")
            else:
                messagebox.showerror("Помилка", f"Помилка підключення до бази даних: {e}\n"
                                                    "Програма буде працювати з тестовими даними.")
            
            self.conn = None
            self.cursor = None
            return False
    
    def is_connected(self):
        """Перевірка, чи є активне підключення до бази даних"""
        return self.conn is not None and self.cursor is not None
    
    def execute_query(self, query, params=None):
        """Виконання запиту до бази даних"""
        if not self.is_connected():
            if not self.connect():
                return None
        
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            return self.cursor
        except Exception as e:
            print(f"Помилка виконання запиту: {e}")
            return None
    
    def fetch_all(self, query, params=None):
        """Виконання запиту та отримання всіх результатів"""
        cursor = self.execute_query(query, params)
        if cursor:
            try:
                return cursor.fetchall()
            except Exception as e:
                print(f"Помилка отримання результатів: {e}")
        return []
    
    def fetch_one(self, query, params=None):
        """Виконання запиту та отримання одного результату"""
        cursor = self.execute_query(query, params)
        if cursor:
            try:
                return cursor.fetchone()
            except Exception as e:
                print(f"Помилка отримання результату: {e}")
        return None
    
    def commit(self):
        """Збереження змін у базі даних"""
        if self.is_connected():
            try:
                self.conn.commit()
                return True
            except Exception as e:
                print(f"Помилка збереження змін: {e}")
        return False
    
    def close(self):
        """Закриття підключення до бази даних"""
        if self.is_connected():
            try:
                self.cursor.close()
                self.conn.close()
                self.conn = None
                self.cursor = None
                return True
            except Exception as e:
                print(f"Помилка закриття підключення: {e}")
        return False
    
    def _invalidate_cache(self):
        """Очищення кешу"""
        self._cache_valid = False
        for key in self._cache:
            self._cache[key] = None
    
    def _get_cached_or_fetch(self, cache_key, fetch_func):
        """Отримання даних з кешу або завантаження з БД"""
        if self._cache_valid and self._cache[cache_key] is not None:
            return self._cache[cache_key]
        
        data = fetch_func()
        self._cache[cache_key] = data
        self._cache_valid = True
        return data
    
    def get_departments(self):
        """Отримання списку відділень з таблиці department"""
        return self._get_cached_or_fetch('departments', self._fetch_departments)
    
    def _fetch_departments(self):
        """Завантаження відділень з БД"""
        # Стандартні назви відділень, які будуть використані, якщо не вдасться отримати дані з бази
        default_dept_names = {
            1: "Загальноосвітньої підготовки",
            2: "Економічне",
            3: "Інформаційних технологій",
            4: "Будівельне",
            5: "Земельно-правове"
        }
        
        print("Отримання списку відділень...")
        
        # Якщо немає підключення до бази даних, повертаємо стандартні відділення
        if not self.is_connected():
            return list(default_dept_names.values())
        
        try:
            # Спочатку отримуємо дані з таблиці department
            dept_names = {}
            try:
                # Отримуємо назви відділень з таблиці department
                rows = self.fetch_all("SELECT ID, Name FROM department ORDER BY ID")
                
                if rows and len(rows) > 0:
                    for row in rows:
                        dept_id = row[0]
                        dept_name = row[1]
                        dept_names[dept_id] = dept_name
                    print(f"Отримано {len(dept_names)} відділень з бази даних: {dept_names}")
            except Exception as e:
                # Якщо виникла помилка, використовуємо стандартні назви
                print(f"Помилка при отриманні назв відділень з таблиці department: {e}")
            
            # Якщо не вдалося отримати назви з бази даних, використовуємо стандартні назви
            if not dept_names:
                dept_names = default_dept_names
                print(f"Використовуємо стандартні назви відділень: {dept_names}")
            
            # Повертаємо ВСІ відділення з таблиці department
            if dept_names:
                departments = list(dept_names.values())
                print(f"Повертаємо всі відділення з таблиці department: {departments}")
                return departments
            
            # Додатково перевіряємо відділення з таблиці groups (для сумісності)
            try:
                rows = self.fetch_all("SELECT DISTINCT [Number Of Department] FROM groups ORDER BY [Number Of Department]")
                if rows and len(rows) > 0:
                    # Отримуємо унікальні ID відділень
                    dept_ids = [row[0] for row in rows]
                    
                    # Створюємо список відділень з назвами
                    departments = []
                    
                    # Додаємо відділення за ID
                    for dept_id in dept_ids:
                        if dept_id in dept_names:
                            departments.append(dept_names[dept_id])
                        else:
                            departments.append(f"Відділення {dept_id}")
                    
                    print(f"Повертаємо список відділень з таблиці groups: {departments}")
                    return departments
            except Exception as e:
                print(f"Помилка при отриманні відділень з таблиці groups: {e}")
        except Exception as e:
            print(f"Помилка при отриманні відділень: {e}")
        
        # Якщо виникла помилка або не вдалося отримати дані, повертаємо стандартні відділення
        print(f"Повертаємо стандартні відділення: {list(default_dept_names.values())}")
        return list(default_dept_names.values())
    
    def fix_group_name(self, group_name):
        """Виправлення кодування назви групи"""
        # Просто повертаємо оригінальну назву групи з бази даних
        return group_name
    
    def get_groups_by_department(self, department):
        """Отримання списку груп за відділенням"""
        try:
            # Спочатку отримуємо ID відділення з бази даних за його назвою
            dept_id = None
            
            # Спробуємо знайти ID відділення в базі даних
            try:
                rows = self.fetch_all("SELECT ID FROM department WHERE Name = ?", (department,))
                if rows and len(rows) > 0:
                    dept_id = rows[0][0]
                    print(f"Знайдено ID відділення '{department}': {dept_id}")
            except Exception as e:
                print(f"Помилка при пошуку ID відділення '{department}': {e}")
            
            # Якщо не знайшли в базі даних, використовуємо резервний словник
            if dept_id is None:
                dept_names = {
                    "Загальноосвітньої підготовки": 1,
                    "Економічне": 2,
                    "Інформаційних технологій": 3,
                    "Будівельне": 4,
                    "Земельно-правове": 5,
                    "Філологічне": 6,
                    "Медичне": 7,
                    "Фізико-математичне": 8
                }
                
                if department in dept_names:
                    dept_id = dept_names[department]
                else:
                    # Якщо відділення не знайдено в словнику, перевіряємо, чи це відділення з номером
                    if department.startswith("Відділення "):
                        try:
                            dept_id = int(department.replace("Відділення ", ""))
                        except:
                            pass
            
            if dept_id is not None:
                # Отримуємо групи для відділення
                rows = self.fetch_all("SELECT Name FROM groups WHERE [Number Of Department] = ?", (dept_id,))
                if rows:
                    # Використовуємо дані з бази даних, але виправляємо кодування
                    fixed_groups = []
                    for row in rows:
                        fixed_name = self.fix_group_name(row[0])
                        fixed_groups.append(fixed_name)
                    print(f"Знайдено {len(fixed_groups)} груп для відділення '{department}' (ID: {dept_id})")
                    return fixed_groups
        except Exception as e:
            print(f"Помилка при отриманні груп для відділення {department}: {e}")
        
        # Якщо не вдалося отримати дані, повертаємо порожній список
        return []
    
    def get_all_groups(self):
        """Отримання словника груп за відділеннями"""
        try:
            # Спочатку отримуємо актуальні назви відділень з бази даних
            dept_id_to_name = {}
            try:
                dept_rows = self.fetch_all("SELECT ID, Name FROM department")
                if dept_rows:
                    for row in dept_rows:
                        dept_id_to_name[row[0]] = row[1]
                    print(f"Отримано мапу відділень з бази даних: {dept_id_to_name}")
            except Exception as e:
                print(f"Помилка при отриманні відділень з бази даних: {e}")
            
            # Якщо не вдалося отримати з бази, використовуємо резервний словник
            if not dept_id_to_name:
                dept_id_to_name = {
                    1: "Загальноосвітньої підготовки",
                    2: "Економічне",
                    3: "Інформаційних технологій",
                    4: "Будівельне",
                    5: "Земельно-правове",
                    6: "Філологічне",
                    7: "Медичне",
                    8: "Фізико-математичне"
                }
                print(f"Використовуємо резервну мапу відділень: {dept_id_to_name}")
            
            # Тепер отримуємо групи з таблиці groups
            rows = self.fetch_all("SELECT Name, [Number Of Department] FROM groups")
            if rows:
                # Створюємо словник груп за відділеннями
                departments = self.get_departments()
                groups = {dept: [] for dept in departments}
                
                # Розподіляємо групи за відділеннями
                for row in rows:
                    group_name = row[0]
                    # Виправляємо кодування назви групи
                    group_name = self.fix_group_name(group_name)
                    dept_id = row[1]
                    
                    # Отримуємо назву відділення за ID
                    if dept_id in dept_id_to_name:
                        dept_name = dept_id_to_name[dept_id]
                    else:
                        # Якщо відділення невідоме, додаємо його з номером
                        dept_name = f"Відділення {dept_id}"
                    
                    # Виправляємо переплутані групи Будівельного та Земельно-правового відділень
                    # (тільки якщо використовуємо резервний словник)
                    if len(dept_id_to_name) <= 8:  # Це означає, що ми використовуємо резервний словник
                        # Якщо назва групи починається з "Б" або містить "Буд", вона належить до Будівельного відділення
                        if group_name.startswith("Б") or "Буд" in group_name or "буд" in group_name.lower():
                            dept_name = "Будівельне"
                        # Якщо назва групи починається з "З" або містить "Зем" або "Прав", вона належить до Земельно-правового відділення
                        elif group_name.startswith("З") or "Зем" in group_name or "Прав" in group_name or "зем" in group_name.lower() or "прав" in group_name.lower():
                            dept_name = "Земельно-правове"
                    
                    # Додаємо групу до відповідного відділення
                    if dept_name in groups:
                        groups[dept_name].append(group_name)
                    else:
                        # Якщо відділення не існує в списку, додаємо його
                        groups[dept_name] = [group_name]
                
                print(f"Створено словник груп: {groups}")
                # Якщо є групи, повертаємо їх
                if any(groups.values()):
                    return groups
        except Exception as e:
            print(f"Помилка при отриманні груп: {e}")
        
        # Якщо не вдалося отримати дані, використовуємо тестові дані
        return {
            "Загальноосвітньої підготовки": ["11-М", "11-Ф", "21-П", "31-О", "22-К"],
            "Економічне": ["11-Е", "21-Е", "31-Е"],
            "Земельно-правове": ["11-З", "21-З", "31-З"],
            "Будівельне": ["11-Б", "21-Б", "31-Б"]
        }
    
    def get_audiences(self):
        """Отримання списку аудиторій"""
        try:
            rows = self.fetch_all("SELECT Number FROM audiences")
            if rows:
                return [str(row[0]) for row in rows]
        except Exception as e:
            print(f"Помилка при отриманні аудиторій: {e}")
        return []
    
    def get_disciplines(self):
        """Отримання списку дисциплін з бази даних"""
        # Спробуємо різні варіанти таблиць і запитів для отримання дисциплін
        try:
            # Спочатку спробуємо таблицю discpline
            print("Спроба отримати дані з таблиці discpline...")
            cursor = self.execute_query("SELECT Name FROM discpline WHERE Name IS NOT NULL AND Name <> '' ORDER BY Name")
            if cursor:
                disciplines = [row.Name for row in cursor.fetchall()]
                if disciplines and len(disciplines) > 0:
                    print(f"Знайдено {len(disciplines)} дисциплін в таблиці discpline")
                    print(f"Перші 5 дисциплін: {disciplines[:5] if len(disciplines) >= 5 else disciplines}")
                    return sorted(disciplines)
        except Exception as e:
            print(f"Помилка при отриманні дисциплін з таблиці discpline: {e}")
        
        try:
            # Спробуємо таблицю discipline
            print("Спроба отримати дані з таблиці discipline...")
            cursor = self.execute_query("SELECT Name FROM discipline WHERE Name IS NOT NULL AND Name <> '' ORDER BY Name")
            if cursor:
                disciplines = [row.Name for row in cursor.fetchall()]
                if disciplines and len(disciplines) > 0:
                    print(f"Знайдено {len(disciplines)} дисциплін в таблиці discipline")
                    print(f"Перші 5 дисциплін: {disciplines[:5] if len(disciplines) >= 5 else disciplines}")
                    return sorted(disciplines)
        except Exception as e:
            print(f"Помилка при отриманні дисциплін з таблиці discipline: {e}")
        
        try:
            # Спробуємо таблицю disciplines
            print("Спроба отримати дані з таблиці disciplines...")
            cursor = self.execute_query("SELECT Name FROM disciplines WHERE Name IS NOT NULL AND Name <> '' ORDER BY Name")
            if cursor:
                disciplines = [row.Name for row in cursor.fetchall()]
                if disciplines and len(disciplines) > 0:
                    print(f"Знайдено {len(disciplines)} дисциплін в таблиці disciplines")
                    print(f"Перші 5 дисциплін: {disciplines[:5] if len(disciplines) >= 5 else disciplines}")
                    return sorted(disciplines)
        except Exception as e:
            print(f"Помилка при отриманні дисциплін з таблиці disciplines: {e}")
        
        try:
            # Спробуємо таблицю Дисципліни
            print("Спроба отримати дані з таблиці Дисципліни...")
            cursor = self.execute_query("SELECT Назва FROM Дисципліни WHERE Назва IS NOT NULL AND Назва <> '' ORDER BY Назва")
            if cursor:
                disciplines = [row.Назва for row in cursor.fetchall()]
                if disciplines and len(disciplines) > 0:
                    print(f"Знайдено {len(disciplines)} дисциплін в таблиці Дисципліни")
                    print(f"Перші 5 дисциплін: {disciplines[:5] if len(disciplines) >= 5 else disciplines}")
                    return sorted(disciplines)
        except Exception as e:
            print(f"Помилка при отриманні дисциплін з таблиці Дисципліни: {e}")
        
        # Якщо не вдалося отримати дисципліни з жодної таблиці, повертаємо стандартний список
        print("Не вдалося отримати дисципліни з бази даних, використовуємо стандартний список")
        return self.get_default_disciplines()
    
    def get_default_disciplines(self):
        """Отримання стандартного списку дисциплін"""
        return [
            "Алгебра", "Англійська мова", "Біологія", 
            "Географія", "Геометрія", 
            "Іноземна мова", "Інформатика", "Історія України", 
            "Математика", "Правознавство", 
            "Українська мова", "Фізика", "Хімія"
        ]
    
    def get_group_departments(self):
        """Отримання словника ID відділень груп"""
        try:
            # Спробуємо отримати дані з таблиці groups
            rows = self.fetch_all("SELECT Name, [Number Of Department] FROM groups")
            if rows:
                return {row[0]: row[1] for row in rows}
        except Exception as e:
            print(f"Помилка при отриманні груп та відділень: {e}")
        
        # Якщо не вдалося отримати дані, використовуємо тестові дані
        return {
            "11-М": 1, "11-Ф": 1, "21-П": 1, "31-О": 1, "22-К": 1,
            "11-Е": 2, "21-Е": 2, "31-Е": 2
        }
    
    def refresh_data(self):
        """Оновлення даних з бази даних (очищення кешу)"""
        print("Оновлення даних з бази даних...")
        
        # Очищаємо кеш
        self._invalidate_cache()
        
        # Перепідключаємося до бази даних для отримання свіжих даних
        if self.is_connected():
            try:
                # Закриваємо поточне підключення
                self.close()
                # Відкриваємо нове підключення
                self.connect()
                print("Дані успішно оновлено")
                return True
            except Exception as e:
                print(f"Помилка при оновленні даних: {e}")
                return False
        else:
            # Якщо не було підключення, спробуємо підключитися
            return self.connect()
    
    def get_department_structure(self):
        """Отримання повної структури відділень з ID та назвами для динамічної генерації бланку"""
        print("Отримання структури відділень...")
        
        # Стандартна структура відділень
        default_structure = [
            {"id": 1, "name": "Загальноосвітньої підготовки", "order": 1},
            {"id": 2, "name": "Економічне", "order": 2},
            {"id": 3, "name": "Інформаційних технологій", "order": 3},
            {"id": 4, "name": "Будівельне", "order": 4},
            {"id": 5, "name": "Земельно-правове", "order": 5}
        ]
        
        # Якщо немає підключення до бази даних, повертаємо стандартну структуру
        if not self.is_connected():
            print("Немає підключення до бази даних, використовуємо стандартну структуру")
            return default_structure
        
        try:
            # Отримуємо всі відділення з бази даних
            dept_rows = self.fetch_all("SELECT ID, Name FROM department ORDER BY ID")
            print(f"Знайдено відділень у таблиці department: {len(dept_rows) if dept_rows else 0}")
            if dept_rows:
                for row in dept_rows:
                    print(f"  - ID: {row[0]}, Назва: {row[1]}")
            
            group_rows = self.fetch_all("SELECT DISTINCT [Number Of Department] FROM groups ORDER BY [Number Of Department]")
            print(f"Знайдено унікальних відділень у таблиці groups: {len(group_rows) if group_rows else 0}")
            if group_rows:
                for row in group_rows:
                    print(f"  - Відділення ID: {row[0]}")
            
            # Використовуємо ВСІ відділення з таблиці department, навіть якщо у них немає груп
            if dept_rows:
                # Створюємо словник назв відділень
                dept_names = {row[0]: row[1] for row in dept_rows}
                
                # Отримуємо ID відділень, які мають групи
                active_dept_ids = [row[0] for row in group_rows] if group_rows else []
                
                # Створюємо структуру для ВСІХ відділень з таблиці department
                structure = []
                for i, (dept_id, dept_name) in enumerate(dept_names.items(), 1):
                    structure.append({
                        "id": dept_id,
                        "name": dept_name,
                        "order": i,
                        "has_groups": dept_id in active_dept_ids
                    })
                
                print(f"Створено структуру з бази даних: {len(structure)} відділень")
                for dept in structure:
                    groups_status = "з групами" if dept['has_groups'] else "без груп"
                    print(f"  - {dept['name']} (ID: {dept['id']}) - {groups_status}")
                
                return structure
            
        except Exception as e:
            print(f"Помилка при отриманні структури відділень: {e}")
        
        # Якщо не вдалося отримати дані, повертаємо стандартну структуру
        print("Використовуємо стандартну структуру відділень")
        return default_structure

    def get_teachers(self):
        """Отримання списку викладачів"""
        try:
            # Прямий запит до таблиці teachers
            rows = self.fetch_all("SELECT PIB FROM teachers")
            if rows:
                return [row[0] for row in rows]
        except Exception as e:
            print(f"Помилка отримання викладачів: {e}")
        
        # Якщо не вдалося отримати дані, повертаємо тестові дані
        return [
            "Петров П.П.", "Іванов І.І.", "Сидоров С.С.", 
            "Ковальчук О.В.", "Шевченко Т.Г.", "Мельник А.М."
        ]
