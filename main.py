import logging
import os
import sys
from pathlib import Path

# Добавляем текущую директорию в путь для импорта модулей
sys.path.append(str(Path(__file__).parent))

# Импортируем модули
from config import Config
from participants_handler import ParticipantsHandler
from certificate_generator import CertificateGenerator
from email_sender import EmailSender
from report_generator import ReportGenerator

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    """Основная функция программы - управляет всем процессом"""
    print("=" * 60)
    print("ГЕНЕРАЦИЯ СЕРТИФИКАТОВ С QR-КОДАМИ И РАССЫЛКА")
    print("=" * 60)
    
    try:
        # Шаг 1: Инициализация
        print("\n[1] ИНИЦИАЛИЗАЦИЯ")
        print("-" * 40)
        ParticipantsHandler.create_directories()
        
        # Шаг 2: Тестирование почты
        print("\n[2] ПРОВЕРКА ПОЧТОВОГО СОЕДИНЕНИЯ")
        print("-" * 40)
        smtp_connected = EmailSender.test_smtp_connection()
        
        if not smtp_connected:
            print("\n⚠️  Не удалось подключиться к почтовому серверу.")
            print("Проверьте настройки в config.py:")
            print(f"  Email: {Config.SENDER_EMAIL}")
            print(f"  Сервер: {Config.SMTP_SERVER}:{Config.SMTP_PORT}")
            print("\nПродолжить без отправки email? (y/n): ", end="")
            
            try:
                user_input = input().strip().lower()
                continue_without_email = (user_input == 'y')
            except:
                print("\n⚠️  Ошибка ввода, продолжаем без отправки email...")
                continue_without_email = True
            
            if not continue_without_email:
                print("\nЗавершение программы.")
                return
        
        # Шаг 3: Загрузка участников
        print("\n[3] ЗАГРУЗКА УЧАСТНИКОВ")
        print("-" * 40)
        print("Выберите источник данных:")
        print("1. CSV файл (по умолчанию)")
        print("2. Тестовые данные")
        print("3. Случайные данные")
        
        source_choice = input("\nВаш выбор (1/2/3): ").strip()
        
        participants = []
        if source_choice == "2":
            participants = ParticipantsHandler.get_test_participants()
        elif source_choice == "3":
            participants = ParticipantsHandler.generate_random_participants()
        else:
            # По умолчанию используем CSV
            csv_path = r"C:\Users\dyrdu\AppData\Local\Programs\Python\Python313\УИРС\art_uirs\participants.csv"
            participants = ParticipantsHandler.import_from_csv(csv_path)
        
        if not participants:
            print("❌ Нет данных для обработки. Завершение программы.")
            return
        
        print(f"✓ Загружено {len(participants)} участников")
        
        # Шаг 4: Генерация сертификатов
        print("\n[4] ГЕНЕРАЦИЯ СЕРТИФИКАТОВ")
        print("-" * 40)
        
        # Удаляем старый шаблон если есть
        template_path = Config.TEMPLATES_DIR / 'certificate_template.html'
        if template_path.exists():
            print("⚠️  Удаляю старый шаблон для создания нового с QR-кодом...")
            try:
                template_path.unlink()
            except:
                pass
        
        generator = CertificateGenerator()
        
        # Генерация сертификатов для всех участников
        successful = 0
        failed = 0
        
        print(f"\nГенерация сертификатов для {len(participants)} участников...")
        print("-" * 60)
        
        for participant in participants:
            print(f"\nУчастник: {participant['full_name']}")
            print(f"Email: {participant['Email']}")
            print(f"Курс: {participant['course_name']}")
            
            result = generator.create_certificate(participant)
            
            if result['status'] == 'success':
                pdf_path = result['pdf_path']
                print(f"✓ Сертификат создан: {pdf_path.name}")
                print(f"  ID сертификата: {participant.get('certificate_id', 'N/A')}")
                print(f"  QR-код сгенерирован и добавлен в сертификат")
                successful += 1
            else:
                print(f"✗ Ошибка создания сертификата: {result.get('error', 'Неизвестная ошибка')}")
                failed += 1
        
        # Шаг 5: Рассылка email
        print("\n[5] РАССЫЛКА EMAIL")
        print("-" * 40)
        
        email_successful = 0
        email_failed = 0
        
        if smtp_connected:
            email_successful, email_failed = EmailSender.send_emails_to_all(participants)
        else:
            print("⚠️  Пропуск отправки email (SMTP не подключен)")
        
        # Шаг 6: Генерация отчета
        print("\n[6] ГЕНЕРАЦИЯ ОТЧЕТА")
        print("-" * 40)
        ReportGenerator.save_report(participants)
        
        # Шаг 7: Вывод итогов
        print("\n" + "=" * 60)
        print("ИТОГОВАЯ СТАТИСТИКА")
        print("=" * 60)
        
        print(f"ОБРАБОТКА УЧАСТНИКОВ:")
        print(f"  Успешно: {successful}")
        print(f"  Не удалось: {failed}")
        print(f"  Всего: {len(participants)}")
        
        if smtp_connected:
            print(f"\nРАССЫЛКА EMAIL:")
            print(f"  Успешно отправлено: {email_successful}")
            print(f"  Не удалось отправить: {email_failed}")
        
        if not smtp_connected and successful > 0:
            print(f"\n⚠️  ВНИМАНИЕ: Сертификаты созданы, но email не отправлены!")
            print(f"   Файлы сертификатов сохранены в: {Config.PDF_OUTPUT_DIR.absolute()}")
        
        print(f"\nДИРЕКТОРИИ:")
        print(f"  Сертификаты: {Config.PDF_OUTPUT_DIR.absolute()}")
        print(f"  QR-коды: {Config.QR_OUTPUT_DIR.absolute()}")
        print(f"  Шаблоны: {Config.TEMPLATES_DIR.absolute()}")
        print("=" * 60)
        
        print("\n✅ Программа успешно завершена!")
        
    except KeyboardInterrupt:
        print("\n\n⚠️  Программа прервана пользователем.")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Запуск с GUI: python main.py --gui
    if "--gui" in sys.argv or "-g" in sys.argv:
        try:
            from gui import start_gui
            start_gui()
        except Exception as e:
            print(f"Ошибка запуска GUI: {e}")
            main()
    else:
        main() 