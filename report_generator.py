import pandas as pd
import logging
from pathlib import Path
from config import Config

logger = logging.getLogger(__name__)

class ReportGenerator:
    """Класс для генерации отчетов"""
    
    @staticmethod
    def save_report(participants: list, output_path: str = "report.csv"):
        """Сохранение отчета в CSV"""
        try:
            report_data = []
            
            for p in participants:
                report_data.append({
                    'ID': p.get('ID', ''),
                    'Полное имя': p.get('full_name', ''),
                    'Email': p.get('Email', ''),
                    'Курс': p.get('course_name', ''),
                    'Часы': p.get('hours', ''),
                    'Дата завершения': p.get('date_completed', ''),
                    'ID сертификата': p.get('certificate_id', ''),
                    'Ссылка для верификации': p.get('verification_url', ''),
                    'Файл сертификата': p.get('pdf_path', Path('')).name if 'pdf_path' in p else '',
                    'Статус': 'Успешно' if 'certificate_id' in p else 'Ошибка'
                })
            
            df = pd.DataFrame(report_data)
            report_path = Path(output_path)
            df.to_csv(report_path, index=False, encoding='utf-8-sig')
            
            logger.info(f"Отчет сохранен: {report_path.absolute()}")
            print(f"\n✓ Отчет сохранен: {report_path.absolute()}")
            
            return report_path
            
        except Exception as e:
            logger.error(f"Ошибка при сохранении отчета: {e}")
            return None