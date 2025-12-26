import warnings
import os
import sys
import io
from pathlib import Path

# Подавление всех предупреждений
warnings.filterwarnings("ignore")
os.environ['GIO_USE_VFS'] = 'local'
os.environ['GTK_USE_PORTAL'] = '0'

# Перенаправляем stderr для подавления предупреждений
sys.stderr = io.StringIO()

# Поддержка .env — читаем простые KEY=VALUE строки из файла .env
def load_dotenv(dotenv_path: Path = Path('.env')):
    if not dotenv_path.exists():
        return
    try:
        for line in dotenv_path.read_text(encoding='utf-8').splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' not in line:
                continue
            key, val = line.split('=', 1)
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key not in os.environ:
                os.environ[key] = val
    except Exception:
        pass

# Загружаем .env при импорте модуля
load_dotenv()

class Config:
    """Класс для хранения конфигурации программы"""
    PDF_OUTPUT_DIR = Path("certificates")
    QR_OUTPUT_DIR = Path("qr_codes")
    TEMPLATES_DIR = Path("templates")
    
    # Конфигурация сертификатов
    CERTIFICATE_CONFIG = {
        'certificate_prefix': 'CERT',
        'base_url': 'https://example.com/verify',
        'default_hours': 40,
        'organization': 'Образовательный центр'
    }
    
    # Настройки для отправки email — берем из окружения (без хардкода)
    SENDER_EMAIL = os.getenv('SENDER_EMAIL', '')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD', '')
    SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.mail.ru')
    try:
        SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
    except ValueError:
        SMTP_PORT = 587
    
    # Стиль сертификата
    FONT_PATH = os.getenv('FONT_PATH', 'arial.ttf')
    FONT_NAME = os.getenv('FONT_NAME', 'Arial')

    @staticmethod
    def save_to_env(updates: dict, dotenv_path: Path = Path('.env')):
        """Сохраняет (дописывает/перезаписывает) пары KEY=VALUE в .env.
        ВАЖНО: сохраняет секреты на диск — используйте осознанно.
        """
        existing = {}
        if dotenv_path.exists():
            try:
                for line in dotenv_path.read_text(encoding='utf-8').splitlines():
                    if not line or line.strip().startswith('#') or '=' not in line:
                        continue
                    k, v = line.split('=', 1)
                    existing[k.strip()] = v.strip()
            except Exception:
                existing = {}
        # Обновляем
        for k, v in updates.items():
            existing[k] = str(v)
        # Записываем
        try:
            lines = [f"{k}={v}" for k, v in existing.items()]
            dotenv_path.write_text('\n'.join(lines), encoding='utf-8')
            for k, v in updates.items():
                os.environ[k] = str(v)
        except Exception:
            pass