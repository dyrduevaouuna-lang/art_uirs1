import hashlib
import datetime
import logging
import base64
import io
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
import qrcode
from weasyprint import HTML
from config import Config

logger = logging.getLogger(__name__)

class CertificateGenerator:
    """Класс для генерации сертификатов с QR-кодами"""
    
    def __init__(self):
        # Создаем Jinja2 окружение
        self.env = Environment(loader=FileSystemLoader(Config.TEMPLATES_DIR))
        
        # Создаем шаблон по умолчанию, если он не существует
        self.create_default_template()
        self.template = self.env.get_template('certificate_template.html')
    
    def create_default_template(self):
        """Создание HTML-шаблона сертификата по умолчанию"""
        template_path = Config.TEMPLATES_DIR / 'certificate_template.html'
        
        if not template_path.exists():
            template_content = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Сертификат</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 0;
        }
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 2cm;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        }
        .certificate {
            background: white;
            border: 20px solid #2c3e50;
            border-radius: 15px;
            padding: 2cm;
            position: relative;
            box-shadow: 0 0 50px rgba(0,0,0,0.2);
            min-height: 15cm;
        }
        .header {
            text-align: center;
            margin-bottom: 1.5cm;
        }
        .title {
            font-size: 36px;
            color: #2c3e50;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .subtitle {
            font-size: 18px;
            color: #7f8c8d;
        }
        .content {
            text-align: center;
            margin: 1.5cm 0;
        }
        .name {
            font-size: 32px;
            color: #2980b9;
            font-weight: bold;
            margin: 20px 0;
            padding: 10px;
            border-bottom: 2px solid #3498db;
            display: inline-block;
        }
        .course {
            font-size: 20px;
            color: #34495e;
            margin: 10px 0;
        }
        .details {
            font-size: 16px;
            color: #7f8c8d;
            margin: 10px 0;
        }
        .footer {
            margin-top: 2cm;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }
        .qr-code {
            text-align: center;
            width: 150px;
        }
        .qr-code img {
            width: 120px;
            height: 120px;
            display: block;
            margin: 0 auto;
        }
        .qr-text {
            font-size: 10px;
            color: #7f8c8d;
            margin-top: 5px;
            text-align: center;
        }
        .signature {
            text-align: right;
        }
        .signature-line {
            width: 200px;
            border-top: 1px solid #2c3e50;
            margin: 40px 0 5px;
        }
        .cert-id {
            position: absolute;
            bottom: 10px;
            right: 10px;
            font-size: 10px;
            color: #95a5a6;
        }
        .decoration {
            position: absolute;
            width: 100px;
            height: 100px;
            opacity: 0.1;
        }
        .decoration-1 {
            top: 50px;
            left: 50px;
            border-radius: 50%;
            background: #3498db;
        }
        .decoration-2 {
            bottom: 50px;
            right: 50px;
            border-radius: 50%;
            background: #e74c3c;
        }
        .verification-info {
            font-size: 9px;
            color: #7f8c8d;
            margin-top: 3px;
            text-align: center;
            word-break: break-all;
        }
        .verification-url {
            font-size: 8px;
            color: #95a5a6;
            margin-top: 2px;
            text-align: center;
            word-break: break-all;
        }
    </style>
</head>
<body>
    <div class="certificate">
        <div class="decoration decoration-1"></div>
        <div class="decoration decoration-2"></div>
        
        <div class="header">
            <div class="title">СЕРТИФИКАТ</div>
            <div class="subtitle">о прохождении обучения</div>
        </div>
        
        <div class="content">
            <div class="subtitle">Настоящим удостоверяется, что</div>
            <div class="name">{{ full_name }}</div>
            <div class="course">успешно завершил(а) курс:</div>
            <div class="course" style="color: #e74c3c; font-size: 22px;">«{{ course_name }}»</div>
            <div class="details">Продолжительность: {{ hours }} академических часов</div>
            <div class="details">Дата завершения: {{ date_completed }}</div>
        </div>
        
        <div class="footer">
            <div class="qr-code">
                <img src="data:image/png;base64,{{ qr_code_base64 }}" alt="QR Code">
                <div class="qr-text">Отсканируйте для верификации</div>
                <div class="verification-info">ID: {{ certificate_id }}</div>
                <div class="verification-url">{{ verification_url|truncate(40) }}</div>
            </div>
            
            <div class="signature">
                <div class="signature-line"></div>
                <div>Директор образовательного центра</div>
                <div style="font-size: 14px; margin-top: 5px;">{{ organization }}</div>
            </div>
        </div>
        
        <div class="cert-id">
            ID: {{ certificate_id }}
        </div>
    </div>
</body>
</html>
            """
            template_path.write_text(template_content, encoding='utf-8')
            logger.info("Создан шаблон сертификата по умолчанию")
    
    def generate_qr_code(self, data: str, participant_name: str):
        """Генерация QR-кода с данными для верификации"""
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(data)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            
            safe_name = "".join(c if c.isalnum() else "_" for c in participant_name)
            qr_filename = f"qr_{safe_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            qr_path = Config.QR_OUTPUT_DIR / qr_filename
            
            img.save(qr_path)
            logger.debug(f"QR-код сохранен: {qr_path}")
            
            # Конвертируем изображение в base64 для HTML
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            qr_base64 = base64.b64encode(buffered.getvalue()).decode()
            
            return qr_path, qr_base64
            
        except Exception as e:
            logger.error(f"Ошибка при генерации QR-кода: {e}")
            raise
    
    def generate_certificate_id(self, participant: dict) -> str:
        """Генерация уникального ID сертификата"""
        data_string = f"{participant['full_name']}_{participant['course_name']}_{participant.get('email', '')}"
        hash_object = hashlib.sha256(data_string.encode())
        short_hash = hash_object.hexdigest()[:12].upper()
        
        return f"{Config.CERTIFICATE_CONFIG['certificate_prefix']}-{short_hash}"
    
    def generate_verification_url(self, certificate_id: str) -> str:
        """Генерация URL для верификации"""
        return f"{Config.CERTIFICATE_CONFIG['base_url']}/verify/{certificate_id}"
    
    def create_certificate(self, participant: dict) -> dict:
        """Создание сертификата для участника"""
        try:
            # Добавляем недостающие поля
            if 'full_name' not in participant:
                participant['full_name'] = f"{participant.get('Имя', '')} {participant.get('Фамилия', '')}"
            
            if 'course_name' not in participant:
                participant['course_name'] = participant.get('course', 'Основы Python')
            
            # Генерация ID сертификата
            certificate_id = self.generate_certificate_id(participant)
            
            # Генерация URL для верификации
            verification_url = self.generate_verification_url(certificate_id)
            
            # Генерация QR-кода (возвращает путь и base64)
            qr_path, qr_base64 = self.generate_qr_code(verification_url, participant['full_name'])
            
            # Подготовка данных для шаблона
            template_data = {
                'full_name': participant['full_name'],
                'course_name': participant['course_name'],
                'date_completed': participant.get('date_completed', datetime.datetime.now().strftime('%Y-%m-%d')),
                'hours': participant.get('hours', Config.CERTIFICATE_CONFIG['default_hours']),
                'certificate_id': certificate_id,
                'qr_code_base64': qr_base64,
                'verification_url': verification_url,
                'organization': Config.CERTIFICATE_CONFIG['organization']
            }
            
            # Рендеринг HTML
            html_content = self.template.render(**template_data)
            
            # Генерация PDF
            safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in participant['full_name'])
            pdf_filename = f"Сертификат_{safe_name}_{certificate_id}.pdf"
            pdf_path = Config.PDF_OUTPUT_DIR / pdf_filename
            
            HTML(string=html_content).write_pdf(pdf_path)
            
            logger.info(f"Сертификат создан: {pdf_filename}")
            
            # Добавляем информацию в объект участника
            participant['certificate_id'] = certificate_id
            participant['pdf_path'] = pdf_path
            participant['verification_url'] = verification_url
            
            return {
                'participant': participant,
                'pdf_path': pdf_path,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error(f"Ошибка при создании сертификата для {participant.get('full_name', 'Неизвестный')}: {e}")
            return {
                'participant': participant,
                'status': 'error',
                'error': str(e)
            }