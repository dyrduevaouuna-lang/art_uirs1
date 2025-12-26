import pandas as pd
import random
import datetime
import os
import logging
from pathlib import Path
from config import Config

logger = logging.getLogger(__name__)

class ParticipantsHandler:
    """Класс для работы с участниками: загрузка, генерация, управление"""
    
    @staticmethod
    def create_directories():
        """Создание необходимых директорий"""
        Config.PDF_OUTPUT_DIR.mkdir(exist_ok=True)
        Config.QR_OUTPUT_DIR.mkdir(exist_ok=True)
        Config.TEMPLATES_DIR.mkdir(exist_ok=True)
        logger.info("Созданы необходимые директории")
    
    @staticmethod
    def import_from_csv(csv_path: str) -> list:
        """Импорт данных участников из CSV файла"""
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
            participants = []
            required_columns = ['Имя', 'Фамилия', 'Email']
            
            # Проверяем обязательные колонки
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"В CSV файле отсутствует обязательная колонка: {col}")
            
            # Обрабатываем каждую строку
            for idx, row in df.iterrows():
                # Формируем полное имя
                full_name = f"{row['Имя']} {row['Фамилия']}"
                if 'Отчество' in df.columns and pd.notna(row.get('Отчество')):
                    full_name = f"{row['Имя']} {row['Отчество']} {row['Фамилия']}"
                
                participant = {
                    "ID": idx + 1,
                    "Имя": str(row['Имя']).strip(),
                    "Фамилия": str(row['Фамилия']).strip(),
                    "full_name": full_name.strip(),
                    "Email": str(row['Email']).strip()
                }
                
                # Опциональные поля
                if 'Курс' in df.columns and pd.notna(row.get('Курс')):
                    participant['course_name'] = str(row['Курс']).strip()
                else:
                    participant['course_name'] = 'Основы Python'
                
                if 'Часы' in df.columns and pd.notna(row.get('Часы')):
                    try:
                        participant['hours'] = int(row['Часы'])
                    except (ValueError, TypeError):
                        participant['hours'] = Config.CERTIFICATE_CONFIG['default_hours']
                else:
                    participant['hours'] = Config.CERTIFICATE_CONFIG['default_hours']
                
                if 'Дата_завершения' in df.columns and pd.notna(row.get('Дата_завершения')):
                    participant['date_completed'] = str(row['Дата_завершения']).strip()
                else:
                    participant['date_completed'] = datetime.datetime.now().strftime('%Y-%m-%d')
                
                participants.append(participant)
            
            logger.info(f"Успешно импортировано {len(participants)} участников из {csv_path}")
            return participants
            
        except FileNotFoundError:
            logger.error(f"CSV файл не найден: {csv_path}")
            raise
        except pd.errors.EmptyDataError:
            logger.error(f"CSV файл пуст: {csv_path}")
            raise
        except Exception as e:
            logger.error(f"Ошибка при импорте из CSV: {e}")
            raise
    
    @staticmethod
    def generate_random_participants(num: int = 5) -> list:
        """Генерация случайных участников"""
        participants = []
        courses = ["Основы Python", "Машинное обучение", "Веб-разработка", "Анализ данных", "DevOps"]
        
        for i in range(1, num + 1):
            first_name = random.choice(["Иван", "Петр", "Сергей", "Анна", "Мария", "Елена", "Алексей"])
            last_name = random.choice(["Иванов", "Петров", "Сидоров", "Кузнецова", "Смирнова", "Попова", "Амажаев"])
            
            participants.append({
                "ID": i,
                "Имя": first_name,
                "Фамилия": last_name,
                "full_name": f"{first_name} {last_name}",
                "Email": f"{last_name.lower()}.{first_name.lower()}{i}@example.com",
                "course_name": random.choice(courses),
                "hours": random.randint(20, 100),
                "date_completed": (datetime.datetime.now() - datetime.timedelta(days=random.randint(1, 30))).strftime('%Y-%m-%d')
            })
        
        logger.info(f"Сгенерировано {len(participants)} случайных участников")
        return participants
    
    @staticmethod
    def get_test_participants() -> list:
        """Получение тестовых участников"""
        participants = [
            {
                "ID": 1,
                "Имя": "Мария",
                "Фамилия": "Дюрдуева",
                "full_name": "Мария Дюрдуева",
                "Email": "dyrduyeva@mail.ru",
                "course_name": "Основы Python",
                "hours": 40,
                "date_completed": datetime.datetime.now().strftime('%Y-%m-%d')
            },
            {
                "ID": 2,
                "Имя": "Алексей",
                "Фамилия": "Амажаев",
                "full_name": "Алексей Амажаев",
                "Email": "a.mazhaa1@mail.ru",
                "course_name": "Веб-разработка",
                "hours": 60,
                "date_completed": (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            }
        ]
        
        logger.info(f"Используются тестовые данные ({len(participants)} участников)")
        return participants
    
    @staticmethod
    def create_example_csv():
        """Создание примера CSV файла"""
        example_data = [
            {
                "Имя": "Диана",
                "Фамилия": "Полотебнова",
                "Отчество": "Алексеевна",
                "Email": "dyrduyeva@mail.ru",
                "Курс": "Основы Python",
                "Часы": 40,
                "Дата_завершения": datetime.datetime.now().strftime('%Y-%m-%d')
            },
            {
                "Имя": "Радик",
                "Фамилия": "Карасин",
                "Отчество": "Дмитриевич",
                "Email": "a.mazhaa1@mail.ru",
                "Курс": "Веб-разработка",
                "Часы": 60,
                "Дата_завершения": (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            }
        ]
        
        df = pd.DataFrame(example_data)
        df.to_csv("participants.csv", index=False, encoding='utf-8-sig')
        logger.info("Создан пример CSV файла: participants.csv")
        return "participants.csv"
    
    @staticmethod
    def load_participants(source_type: str = "csv", csv_path: str = None) -> list:
        """Универсальный метод загрузки участников"""
        if csv_path is None:
            csv_path = r"D:\popitka3\certification_system_tpu\src\participants.csv"
        
        if source_type == "csv":
            if not os.path.exists(csv_path):
                print(f"⚠️  Файл не найден: {csv_path}")
                print("Создаю пример CSV файла...")
                csv_path = ParticipantsHandler.create_example_csv()
            
            return ParticipantsHandler.import_from_csv(csv_path)
        
        elif source_type == "test":
            return ParticipantsHandler.get_test_participants()
        
        elif source_type == "random":
            num = input("Сколько участников сгенерировать? (по умолчанию 5): ").strip()
            num = int(num) if num else 5
            return ParticipantsHandler.generate_random_participants(num)
        
        else:
            raise ValueError(f"Неизвестный тип источника: {source_type}")