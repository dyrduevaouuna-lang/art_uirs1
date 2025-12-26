import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from participants_handler import ParticipantsHandler
from certificate_generator import CertificateGenerator
from email_sender import EmailSender
from report_generator import ReportGenerator
from config import Config
import os
import traceback


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Certificate Generator — GUI")
        root.geometry("900x600")

        self.participants = []

        # --- Frames ---
        top_frame = ttk.Frame(root)
        top_frame.pack(fill="x", padx=8, pady=6)

        mid_frame = ttk.Frame(root)
        mid_frame.pack(fill="both", expand=True, padx=8, pady=6)

        bottom_frame = ttk.Frame(root)
        bottom_frame.pack(fill="x", padx=8, pady=6)

        # Notebook with tabs: Participants and Settings
        self.notebook = ttk.Notebook(mid_frame)
        self.notebook.pack(fill="both", expand=True)

        participants_tab = ttk.Frame(self.notebook)
        settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(participants_tab, text="Участники")
        self.notebook.add(settings_tab, text="Настройки")

        # --- Controls ---
        self.load_button = ttk.Button(top_frame, text="Загрузить CSV", command=self.load_csv)
        self.load_button.pack(side="left")

        self.sample_button = ttk.Button(top_frame, text="Тестовые данные", command=self.load_test_data)
        self.sample_button.pack(side="left", padx=(6, 0))

        self.generate_button = ttk.Button(top_frame, text="Генерировать сертификаты", command=self.generate_certificates)
        self.generate_button.pack(side="left", padx=(6, 0))

        self.send_button = ttk.Button(top_frame, text="Отправить email", command=self.send_emails)
        self.send_button.pack(side="left", padx=(6, 0))

        self.report_button = ttk.Button(top_frame, text="Сохранить отчёт", command=self.save_report)
        self.report_button.pack(side="left", padx=(6, 0))

        # --- Treeview for participants (in participants tab) ---
        columns = ("ID", "full_name", "Email", "course_name", "certificate_id", "status")
        self.tree = ttk.Treeview(participants_tab, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=120)

        self.tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(participants_tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # --- Settings tab ---
        smtp_frame = ttk.LabelFrame(settings_tab, text="SMTP настройки")
        smtp_frame.pack(fill="x", padx=8, pady=8)

        ttk.Label(smtp_frame, text="Отправитель (email):").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.email_var = tk.StringVar(value=Config.SENDER_EMAIL)
        self.email_entry = ttk.Entry(smtp_frame, textvariable=self.email_var, width=45)
        self.email_entry.grid(row=0, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(smtp_frame, text="Пароль:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.pwd_var = tk.StringVar(value="")
        self.pwd_entry = ttk.Entry(smtp_frame, textvariable=self.pwd_var, show="*", width=45)
        self.pwd_entry.grid(row=1, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(smtp_frame, text="SMTP сервер:").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        self.server_var = tk.StringVar(value=Config.SMTP_SERVER)
        self.server_entry = ttk.Entry(smtp_frame, textvariable=self.server_var, width=45)
        self.server_entry.grid(row=2, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(smtp_frame, text="Порт:").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        self.port_var = tk.StringVar(value=str(Config.SMTP_PORT))
        self.port_entry = ttk.Entry(smtp_frame, textvariable=self.port_var, width=10)
        self.port_entry.grid(row=3, column=1, sticky="w", padx=4, pady=2)

        self.save_pwd_var = tk.BooleanVar(value=False)
        self.save_checkbox = ttk.Checkbutton(smtp_frame, text="Сохранить в .env (включая пароль)", variable=self.save_pwd_var)
        self.save_checkbox.grid(row=4, column=1, sticky="w", padx=4, pady=6)

        btn_frame = ttk.Frame(settings_tab)
        btn_frame.pack(fill="x", padx=8, pady=(0,8))

        self.save_settings_button = ttk.Button(btn_frame, text="Сохранить настройки", command=self.save_settings)
        self.save_settings_button.pack(side="left")

        self.test_smtp_button = ttk.Button(btn_frame, text="Тест SMTP", command=self.test_smtp)
        self.test_smtp_button.pack(side="left", padx=(8,0))

        # --- Log / status ---
        self.log_text = tk.Text(bottom_frame, height=6, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        # Progress
        self.progress = ttk.Progressbar(bottom_frame, mode="determinate")
        self.progress.pack(fill="x", pady=(6, 0))

        # Initialize dirs
        ParticipantsHandler.create_directories()

    def append_log(self, message: str):
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _append)

    def load_csv(self):
        csv_path = filedialog.askopenfilename(title="Выберите CSV-файл", filetypes=[("CSV files", "*.csv"), ("All files", "*")])
        if not csv_path:
            return
        try:
            self.append_log(f"Загрузка участников из {csv_path}...")
            participants = ParticipantsHandler.import_from_csv(csv_path)
            self.participants = participants
            self.populate_tree()
            self.append_log(f"Загружено {len(participants)} участников")
        except Exception as e:
            self.append_log(f"Ошибка загрузки CSV: {e}")
            messagebox.showerror("Ошибка", str(e))

    def load_test_data(self):
        self.participants = ParticipantsHandler.get_test_participants()
        self.populate_tree()
        self.append_log("Загружены тестовые данные")

    def populate_tree(self):
        # Очистка
        for i in self.tree.get_children():
            self.tree.delete(i)

        for p in self.participants:
            self.tree.insert("", "end", values=(
                p.get('ID', ''),
                p.get('full_name', ''),
                p.get('Email', ''),
                p.get('course_name', ''),
                p.get('certificate_id', ''),
                'Готов' if 'certificate_id' in p else ''
            ))

    def generate_certificates(self):
        if not self.participants:
            messagebox.showwarning("Нет данных", "Нет участников для обработки")
            return

        def worker():
            try:
                generator = CertificateGenerator()
                total = len(self.participants)
                self.progress['maximum'] = total
                self.progress['value'] = 0

                for idx, p in enumerate(self.participants, start=1):
                    self.append_log(f"Генерация: {p.get('full_name')}")
                    result = generator.create_certificate(p)
                    if result.get('status') == 'success':
                        p.update(result.get('participant', {}))
                        self.append_log(f"✓ Создан: {p.get('pdf_path').name}")
                    else:
                        self.append_log(f"✗ Ошибка: {result.get('error')}")

                    # Обновляем строку в TreeView
                    self.root.after(0, self.populate_tree)
                    self.progress['value'] = idx

                self.append_log("Генерация завершена")
                messagebox.showinfo("Готово", "Генерация сертификатов завершена")
            except Exception as e:
                self.append_log(f"Ошибка генерации: {e}")
                traceback.print_exc()

        threading.Thread(target=worker, daemon=True).start()

    def send_emails(self):
        if not self.participants:
            messagebox.showwarning("Нет данных", "Нет участников для отправки")
            return

        # Берём значения из UI (если заполнены)
        email = self.email_var.get().strip() if hasattr(self, 'email_var') else ''
        pwd = self.pwd_var.get().strip() if hasattr(self, 'pwd_var') else ''
        server = self.server_var.get().strip() if hasattr(self, 'server_var') else ''
        port = self.port_var.get().strip() if hasattr(self, 'port_var') else ''

        if email:
            Config.SENDER_EMAIL = email
        if server:
            Config.SMTP_SERVER = server
        try:
            if port:
                Config.SMTP_PORT = int(port)
        except ValueError:
            pass

        # Если пароль введён в UI, используем его; иначе, если нет пароля, запрашиваем временно
        if pwd:
            Config.SENDER_PASSWORD = pwd
        elif not Config.SENDER_PASSWORD:
            pwd = tk.simpledialog.askstring("Пароль", "Введите пароль отправителя:", show='*')
            if pwd:
                Config.SENDER_PASSWORD = pwd
            else:
                return

        def worker():
            try:
                connected = EmailSender.test_smtp_connection()
                if not connected:
                    self.append_log("Ошибка подключения к SMTP. Проверьте настройки.")
                    messagebox.showerror("Ошибка", "Не удалось подключиться к SMTP-серверу")
                    return

                sent_success, sent_failed = EmailSender.send_emails_to_all(self.participants)
                self.append_log(f"Отправлено: {sent_success}, Ошибок: {sent_failed}")
                messagebox.showinfo("Готово", f"Отправлено: {sent_success}, Ошибок: {sent_failed}")
            except Exception as e:
                self.append_log(f"Ошибка отправки: {e}")
                traceback.print_exc()

        threading.Thread(target=worker, daemon=True).start()

    def save_settings(self):
        """Применить настройки из вкладки и опционально сохранить в .env"""
        Config.SENDER_EMAIL = self.email_var.get().strip()
        pwd = self.pwd_var.get().strip()
        if pwd:
            Config.SENDER_PASSWORD = pwd
        Config.SMTP_SERVER = self.server_var.get().strip()
        try:
            Config.SMTP_PORT = int(self.port_var.get().strip())
        except Exception:
            Config.SMTP_PORT = 587

        if self.save_pwd_var.get():
            Config.save_to_env({
                'SENDER_EMAIL': Config.SENDER_EMAIL,
                'SENDER_PASSWORD': Config.SENDER_PASSWORD,
                'SMTP_SERVER': Config.SMTP_SERVER,
                'SMTP_PORT': str(Config.SMTP_PORT)
            })
            self.append_log("Настройки сохранены в .env")
            messagebox.showinfo("Сохранено", "Настройки сохранены в .env")
        else:
            self.append_log("Настройки применены в сессии (не сохранены в .env)")
            messagebox.showinfo("Сохранено", "Настройки применены в текущей сессии (не сохранены на диск)")

    def test_smtp(self):
        """Тест соединения с SMTP с текущими значениями в UI (не сохраняет по умолчанию)"""
        # Временно применяем UI значения
        email = self.email_var.get().strip()
        pwd = self.pwd_var.get().strip()
        server = self.server_var.get().strip()
        port = self.port_var.get().strip()

        old = {
            'SENDER_EMAIL': Config.SENDER_EMAIL,
            'SENDER_PASSWORD': Config.SENDER_PASSWORD,
            'SMTP_SERVER': Config.SMTP_SERVER,
            'SMTP_PORT': Config.SMTP_PORT
        }

        try:
            if email:
                Config.SENDER_EMAIL = email
            if pwd:
                Config.SENDER_PASSWORD = pwd
            if server:
                Config.SMTP_SERVER = server
            try:
                if port:
                    Config.SMTP_PORT = int(port)
            except Exception:
                pass

            connected = EmailSender.test_smtp_connection()
            if connected:
                self.append_log("✓ SMTP: подключение установлено")
                messagebox.showinfo("OK", "Успешное подключение к SMTP")
            else:
                self.append_log("✗ SMTP: ошибка подключения")
                messagebox.showerror("Ошибка", "Не удалось подключиться к SMTP-серверу. Проверьте лог.")
        finally:
            # Восстанавливаем прежние значения (не сохраняем изменения, если пользователь не нажал Сохранить)
            Config.SENDER_EMAIL = old['SENDER_EMAIL']
            Config.SENDER_PASSWORD = old['SENDER_PASSWORD']
            Config.SMTP_SERVER = old['SMTP_SERVER']
            Config.SMTP_PORT = old['SMTP_PORT']
    def save_report(self):
        try:
            default_name = "report.csv"
            path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=default_name, filetypes=[("CSV files", "*.csv")])
            if not path:
                return
            ReportGenerator.save_report(self.participants, output_path=path)
            self.append_log(f"Отчёт сохранен: {path}")
            messagebox.showinfo("Готово", f"Отчёт сохранен: {path}")
        except Exception as e:
            self.append_log(f"Ошибка сохранения отчёта: {e}")
            messagebox.showerror("Ошибка", str(e))


def start_gui():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    start_gui()
