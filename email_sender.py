import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formatdate
from pathlib import Path
from config import Config

logger = logging.getLogger(__name__)

class EmailSender:
    """Класс для отправки email с сертификатами"""
    
    @staticmethod
    def test_smtp_connection() -> bool:
        """Тестирование подключения к SMTP серверу"""
        try:
            logger.info(f"Тестирование подключения к {Config.SMTP_SERVER}:{Config.SMTP_PORT}")
            
            if Config.SMTP_PORT == 587:
                server = smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT)
                server.starttls()
            elif Config.SMTP_PORT == 465:
                server = smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT)
            else:
                logger.error(f"Неподдерживаемый порт: {Config.SMTP_PORT}")
                return False
            
            server.login(Config.SENDER_EMAIL, Config.SENDER_PASSWORD)
            server.quit()
            
            logger.info("✓ Подключение успешно!")
            return True
            
        except Exception as e:
            logger.error(f"✗ Ошибка подключения: {e}")
            return False
    
    @staticmethod
    def send_email_with_attachment(recipient_email: str, subject: str, body: str, attachment_path: Path) -> bool:
        """Отправка email с вложением"""
        try:
            message = MIMEMultipart()
            message["From"] = Config.SENDER_EMAIL
            message["To"] = recipient_email
            message["Subject"] = subject
            message["Date"] = formatdate(localtime=True)
            message.attach(MIMEText(body, "plain", "utf-8"))

            # Добавляем вложение
            filename = attachment_path.name
            
            with open(attachment_path, "rb") as attachment:
                # Указываем правильный MIME-тип для PDF
                part = MIMEBase("application", "pdf")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                
                # Кодируем имя файла для поддержки русских символов
                encoded_filename = Header(filename, 'utf-8').encode()
                
                # Устанавливаем заголовки
                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=encoded_filename
                )
                part.add_header(
                    "Content-Type",
                    "application/pdf",
                    name=encoded_filename
                )
                message.attach(part)

            # Настройка SMTP и отправка
            if Config.SMTP_PORT == 587:
                server = smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT)
                server.starttls()
            elif Config.SMTP_PORT == 465:
                server = smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT)
            else:
                logger.error(f"Неподдерживаемый порт: {Config.SMTP_PORT}")
                return False

            server.login(Config.SENDER_EMAIL, Config.SENDER_PASSWORD)
            server.send_message(message)
            server.quit()
            
            logger.info(f"Email успешно отправлен на {recipient_email}")
            return True
            
        except smtplib.SMTPAuthenticationError:
            logger.error("Ошибка аутентификации. Проверьте email и пароль.")
            return False
        except Exception as e:
            logger.error(f"Ошибка при отправке email на {recipient_email}: {e}")
            return False
    
    @staticmethod
    def send_certificate_email(participant: dict) -> bool:
        """Отправка email с сертификатом для конкретного участника"""
        if 'pdf_path' not in participant:
            logger.error(f"У участника {participant['full_name']} нет сертификата для отправки")
            return False
        
        subject = f"Ваш сертификат: {participant['course_name']}"
        body = f"""Уважаемый(ая) {participant['full_name']}!

Поздравляем с успешным завершением курса «{participant['course_name']}»!

Ваш сертификат прикреплен к этому письму. 
Для проверки подлинности сертификата отсканируйте QR-код или перейдите по ссылке:
{participant.get('verification_url', '')}

ID вашего сертификата: {participant.get('certificate_id', 'N/A')}

С наилучшими пожеланиями,
{Config.CERTIFICATE_CONFIG['organization']}
"""
        
        return EmailSender.send_email_with_attachment(
            participant['Email'], 
            subject, 
            body, 
            participant['pdf_path']
        )
    
    @staticmethod
    def send_emails_to_all(participants: list) -> tuple:
        """Отправка email всем участникам с сертификатами"""
        successful = 0
        failed = 0
        
        print(f"\nОтправка email для {len(participants)} участников...")
        print("-" * 60)
        
        for participant in participants:
            if 'certificate_id' in participant and 'pdf_path' in participant:
                print(f"\nОтправка email для: {participant['full_name']}")
                
                if EmailSender.send_certificate_email(participant):
                    print(f"✓ Email отправлен")
                    successful += 1
                else:
                    print(f"✗ Ошибка отправки email")
                    failed += 1
            else:
                print(f"⚠️  У участника {participant['full_name']} нет сертификата")
                failed += 1
        
        return successful, failed