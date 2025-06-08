import sys
import os
import winsound
import json
import win32evtlog
import pywintypes
import ctypes
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QWidget, QGroupBox, QCheckBox, QLineEdit,
    QPushButton, QFileDialog, QTextEdit, QLabel, QHBoxLayout, QVBoxLayout
)
from PyQt6.QtGui import QPalette, QColor, QIntValidator, QIcon
from PyQt6.QtCore import Qt, QTimer
import subprocess
import tempfile

APPDATA_DIR = os.path.join(os.getenv("APPDATA"), "forgotten", "WHEAD")
os.makedirs(APPDATA_DIR, exist_ok=True)

LOG_FILE_PATH = os.path.join(APPDATA_DIR, "LTC-Logger.log")
CONFIG_PATH = os.path.join(APPDATA_DIR, "LTC-Cursed.forgotten")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def clear_log():
    try:
        with open(LOG_FILE_PATH, "w", encoding="utf-8") as f:
            f.write("")
    except Exception:
        pass

def write_log(text: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    line = f"[{timestamp}] {text}\n"
    try:
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

def event_to_dict(event):
    d = {}
    for attr in dir(event):
        if attr.startswith('_'):
            continue
        try:
            value = getattr(event, attr)
            if callable(value):
                continue
            if isinstance(value, (datetime, pywintypes.Time)):
                value = str(value)
            d[attr] = value
        except Exception:
            continue
    return d

def show_messagebox(text, title="WHEA Monitor"):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)  # MB_ICONINFORMATION

class TriggerSettingsForm(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Настройка триггера")

        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFixedSize(450, 520)

        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(220, 220, 220))
        self.setPalette(palette)

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1f26;
                color: #e0e0e0;
                font-family: 'Segoe UI';
                font-size: 14px;
            }
            QGroupBox {
                background-color: #323232;
                border: 1px solid gray;
                border-radius: 4px;
                margin-top: 4px;
                padding: 8px;
                min-height: 130px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #F0F0F0;
                font-weight: bold;
                font-size: 15px;
            }
            QCheckBox {
                color: #F0F0F0;
                spacing: 6px;
                padding: 2px 0;
                font-size: 14px;
                background-color: transparent;
            }
            QLabel {
                background-color: transparent;
                color: #F0F0F0;
                font-size: 14px;
            }
            QLineEdit {
                background-color: #2c2d35;
                border: 1px solid #444;
                border-radius: 6px;
                padding: 6px 8px;
                color: #ffffff;
                min-height: 20px;
                font-size: 14px;
            }
            QPushButton {
                background-color: #3f51b5;
                border-radius: 6px;
                color: white;
                padding: 4px 10px;
                min-height: 28px;
            }
            QPushButton:hover {
                background-color: #303f9f;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # Панель "Сообщение"
        self.group_message = QGroupBox("Сообщение")
        layout_msg = QVBoxLayout(self.group_message)

        self.checkbox_msg_enable = QCheckBox("Включить")
        self.checkbox_msg_enable.stateChanged.connect(self.update_message_controls)

        self.checkbox_msg_custom = QCheckBox("Кастомный текст")
        self.checkbox_msg_custom.setEnabled(False)
        self.checkbox_msg_custom.stateChanged.connect(self.update_message_controls)

        self.line_message = QLineEdit()
        self.line_message.setPlaceholderText("Сообщение")
        self.line_message.setEnabled(False)

        layout_msg.addWidget(self.checkbox_msg_enable)
        layout_msg.addWidget(self.checkbox_msg_custom)
        layout_msg.addWidget(self.line_message)
        layout_msg.addStretch()

        main_layout.addWidget(self.group_message)

        # Панель "Звук"
        self.group_audio = QGroupBox("Звук")
        layout_audio = QVBoxLayout(self.group_audio)

        self.checkbox_audio_enable = QCheckBox("Включить")
        self.checkbox_audio_enable.stateChanged.connect(self.update_audio_controls)

        layout_audio.addWidget(self.checkbox_audio_enable)

        self.sounds = [
            ("Windows Background", r"C://Windows//Media//Windows Background.wav"),
            ("System Notify", r"C://Windows//Media//Windows Notify System Generic.wav"),
            ("Critical Stop", r"C://Windows//Media//Windows Critical Stop.wav")
        ]
        self.sound_buttons = []
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        for idx, (name, path) in enumerate(self.sounds):
            btn = QPushButton(name)
            btn.setEnabled(False)
            btn.setFixedHeight(28)
            btn.clicked.connect(self.make_sound_button_handler(idx, path))
            self.sound_buttons.append(btn)
            buttons_layout.addWidget(btn)
        layout_audio.addLayout(buttons_layout)

        self.label_selected_sound = QLabel("Выбран звук: —")
        self.label_selected_sound.setStyleSheet("background-color: transparent; margin-top: 8px;")
        layout_audio.addWidget(self.label_selected_sound)
        layout_audio.addStretch()

        main_layout.addWidget(self.group_audio)

        # Панель "Запуск"
        self.group_execute = QGroupBox("Запуск")
        layout_exec = QVBoxLayout(self.group_execute)

        self.checkbox_exec_enable = QCheckBox("Включить")
        self.checkbox_exec_enable.stateChanged.connect(self.update_execute_controls)

        layout_exec.addWidget(self.checkbox_exec_enable)

        path_layout = QHBoxLayout()
        self.line_program = QLineEdit()
        self.line_program.setPlaceholderText("Путь к программе")
        self.btn_browse = QPushButton("Обзор")
        self.btn_browse.clicked.connect(self.browse_program)
        self.btn_browse.setEnabled(False)
        path_layout.addWidget(self.line_program)
        path_layout.addWidget(self.btn_browse)
        layout_exec.addLayout(path_layout)

        self.line_args = QLineEdit()
        self.line_args.setPlaceholderText("Аргумент запуска")
        self.line_args.setEnabled(False)
        layout_exec.addWidget(self.line_args)
        layout_exec.addStretch()

        main_layout.addWidget(self.group_execute)

        self.selected_audio_index = None

        self.load_config()
        self.update_message_controls()
        self.update_audio_controls()
        self.update_execute_controls()

    def update_message_controls(self):
        enabled = self.checkbox_msg_enable.isChecked()
        self.checkbox_msg_custom.setEnabled(enabled)
        custom_enabled = self.checkbox_msg_custom.isChecked()
        self.line_message.setEnabled(enabled and custom_enabled)

    def update_audio_controls(self):
        enabled = self.checkbox_audio_enable.isChecked()
        for btn in self.sound_buttons:
            btn.setEnabled(enabled)
        if not enabled:
            self.selected_audio_index = None
            self.label_selected_sound.setText("Выбран звук: —")
        else:
            if self.selected_audio_index is not None:
                self.label_selected_sound.setText(f"Выбран звук: {self.sounds[self.selected_audio_index][0]}")
            else:
                self.label_selected_sound.setText("Выбран звук: —")

    def update_execute_controls(self):
        enabled = self.checkbox_exec_enable.isChecked()
        self.line_program.setEnabled(enabled)
        self.btn_browse.setEnabled(enabled)
        self.line_args.setEnabled(enabled)

    def make_sound_button_handler(self, idx, path):
        def handler():
            if os.path.exists(path):
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                self.selected_audio_index = idx
                self.label_selected_sound.setText(f"Выбран звук: {self.sounds[idx][0]}")
        return handler

    def browse_program(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите .exe или .bat", "", "Executable Files (*.exe *.bat)")
        if path:
            self.line_program.setText(path)

    def get_current_config(self):
        cfg = {
            "message_enabled": self.checkbox_msg_enable.isChecked(),
            "message_custom_enabled": self.checkbox_msg_custom.isChecked(),
            "message_text": self.line_message.text(),

            "audio_enabled": self.checkbox_audio_enable.isChecked(),
            "audio_selected_index": self.selected_audio_index,

            "execute_enabled": self.checkbox_exec_enable.isChecked(),
            "execute_path": self.line_program.text(),
            "execute_args": self.line_args.text(),
        }
        return cfg

    def load_config(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    self.checkbox_msg_enable.setChecked(cfg.get("message_enabled", False))
                    self.checkbox_msg_custom.setChecked(cfg.get("message_custom_enabled", False))
                    self.line_message.setText(cfg.get("message_text", ""))
                    self.checkbox_audio_enable.setChecked(cfg.get("audio_enabled", False))

                    idx = cfg.get("audio_selected_index")
                    self.selected_audio_index = idx if idx in [0,1,2] else None
                    if self.selected_audio_index is not None:
                        self.label_selected_sound.setText(f"Выбран звук: {self.sounds[self.selected_audio_index][0]}")
                    else:
                        self.label_selected_sound.setText("Выбран звук: —")

                    self.checkbox_exec_enable.setChecked(cfg.get("execute_enabled", False))
                    self.line_program.setText(cfg.get("execute_path", ""))
                    self.line_args.setText(cfg.get("execute_args", ""))
            except Exception as e:
                write_log(f"Ошибка загрузки конфигурации: {e}")
        else:
            write_log(f"Файл конфигурации {CONFIG_PATH} не найден, загружены настройки по умолчанию")

    def save_config(self):
        cfg = self.get_current_config()
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
            write_log("Конфигурация триггера сохранена")
            if hasattr(self, 'save_callback') and self.save_callback:
                self.save_callback(cfg)
        except Exception as e:
            write_log(f"Ошибка сохранения конфигурации: {e}")

    def closeEvent(self, event):
        self.save_config()
        event.accept()

class WheaMonitorApp(QWidget):
    WHEA_EVENT_IDS = {17, 18, 19, 20, 41, 45, 46, 47}

    def __init__(self):
        super().__init__()
        clear_log()
        write_log(f"Приложение запущено. Логи: {LOG_FILE_PATH}")

        self.setWindowTitle("Монитор WHEA")

        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFixedSize(550, 380)

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1f26;
                color: #e0e0e0;
                font-family: 'Segoe UI';
                font-size: 14px;
            }
            QGroupBox {
                background-color: #323232;
                border: 1px solid gray;
                border-radius: 4px;
                margin-top: 6px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #F0F0F0;
            }
            QLabel, QCheckBox {
                color: #F0F0F0;
            }
            QLineEdit, QComboBox {
                background-color: #2c2d35;
                border: 1px solid #444;
                border-radius: 6px;
                padding: 4px 8px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #3f51b5;
                border-radius: 6px;
                color: white;
                padding: 4px 10px;
            }
            QPushButton:hover {
                background-color: #303f9f;
            }
        """)

        self.interval_label = QLabel("Интервал проверки (сек)", self)
        self.interval_label.setStyleSheet("background-color: transparent;")
        self.interval_label.adjustSize()
        self.interval_label.move(20, 20)

        self.interval_input = QLineEdit("30", self)
        self.interval_input.setValidator(QIntValidator(5, 3600))
        self.interval_input.setGeometry(200, 18, 60, 24)

        self.start_button = QPushButton("Старт", self)
        self.start_button.setGeometry(280, 16, 80, 28)
        self.stop_button = QPushButton("Стоп", self)
        self.stop_button.setGeometry(370, 16, 80, 28)
        self.stop_button.setEnabled(False)

        self.log_output = QTextEdit(self)
        self.log_output.setGeometry(20, 60, 510, 250)
        self.log_output.setReadOnly(True)

        self.trigger_button = QPushButton("⚡ Настроить триггер", self)
        self.trigger_button.setGeometry(20, 320, 180, 30)

        self.tools_button = QPushButton("WHEA Tools", self)
        self.tools_button.setGeometry(220, 320, 130, 30)
        self.tools_button.clicked.connect(self.run_whea_tools)

        self.trigger_form = TriggerSettingsForm()
        self.trigger_form.save_callback = self.update_trigger_config
        self.trigger_form.load_config()
        self.trigger_config = self.trigger_form.get_current_config()

        self.start_button.clicked.connect(self.start_monitor)
        self.stop_button.clicked.connect(self.stop_monitor)
        self.trigger_button.clicked.connect(lambda: self.trigger_form.show())

        self.timer = QTimer()
        self.timer.timeout.connect(self.check_whea_events)

        self.monitor_start_time = datetime.now()
        self.last_error_count = 0

    def run_whea_tools(self):
        bat_code = r"""@echo off
chcp 65001 >nul

:: Проверка прав администратора
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Admin rights required. Restarting as admin...
    powershell -Command "Start-Process -Verb runAs -FilePath '%~f0'"
    exit /b
)

:main_loop
cls
echo ┌─────┐ ┌─────┐ ┌─────┐  ┌───────────────────────────────────────────────────┐
echo │ ┌───┘ │ ┌─┐ │ │ ┌───┘  │ Forgotten Private Coalition                       │
echo │ └───┐ │ └─┘ │ │ │      │ WHEA trigger tools                                │
echo │ ┌───┘ │ ┌───┘ │ │      │ Private version                                   │
echo │ │     │ │     │ └───┐  │ License CC BY 4.0                                 │
echo └─┘     └─┘     └─────┘  └───────────────────────────────────────────────────┘
echo.
echo Available WHEA EventIDs and meanings:
echo  17 - General hardware error event
echo  18 - Machine Check Exception (MCE)
echo  19 - Corrected Machine Check error
echo  20 - PCI Express error
echo  41 - Detailed WHEA error report
echo  45 - Corrected memory error
echo  46 - Corrected processor error
echo  47 - Corrected PCI Express error
echo.
echo Команды:
echo  ex - Закрыть программу
echo  el - Открыть журнал событий Windows/System
echo  ec - Очистить все ошибки от источника TestWHEA
echo.

if defined last_message (
    echo %last_message%
    echo.
    set "last_message="
)

set /p input=Введите "EventID Level" (Level: 1=Warning, 2=Error) или команду (ex, el, ec): 

:: Проверка команд
if /i "%input%"=="ex" (
    exit /b
)
if /i "%input%"=="el" (
    start eventvwr.msc /s:"System"
    set last_message=Журнал событий Windows/System запущен.
    goto main_loop
)
if /i "%input%"=="ec" (
    echo Очистка всех ошибок от источника TestWHEA...
    wevtutil.exe cl System
    set last_message=Журнал System очищен.
    goto main_loop
)

:: Обработка ввода EventID и Level
for /f "tokens=1,2" %%a in ("%input%") do (
    set "code=%%a"
    set "level=%%b"
)

if not defined code (
    set last_message=Неверный ввод. Попробуйте снова.
    goto main_loop
)
if not defined level (
    set last_message=Неверный ввод. Попробуйте снова.
    goto main_loop
)

set executed=0

for %%E in (17 18 19 20 41 45 46 47) do (
    if "%code%"=="%%E" (
        if "%level%"=="1" (
            eventcreate /T WARNING /ID %%E /L SYSTEM /SO TestWHEA /D "Test warning WHEA (EventID %%E)"
            set executed=1
        ) else if "%level%"=="2" (
            eventcreate /T ERROR /ID %%E /L SYSTEM /SO TestWHEA /D "Test error WHEA (EventID %%E)"
            set executed=1
        )
    )
)

if "%executed%"=="1" (
    set last_message=Операция выполнена успешно.
    goto main_loop
) else (
    set last_message=Неверный ввод или неподдерживаемый EventID/Level. Попробуйте снова.
    goto main_loop
)
"""
        try:
            with tempfile.NamedTemporaryFile("w", delete=False, suffix=".bat", encoding="utf-8") as tmpfile:
                tmpfile.write(bat_code)
                tmp_path = tmpfile.name
            subprocess.Popen(
                ["cmd.exe", "/c", tmp_path],
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception as e:
            write_log(f"Ошибка запуска WHEA Tools: {e}")

    def update_trigger_config(self, config):
        self.trigger_config = config
        write_log(f"Обновлена конфигурация триггера: {json.dumps(config, ensure_ascii=False)}")

    def start_monitor(self):
        try:
            interval = int(self.interval_input.text())
            if interval < 5 or interval > 3600:
                raise ValueError()
        except ValueError:
            self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] Ошибка: интервал должен быть от 5 до 3600")
            write_log("Ошибка запуска мониторинга: неверный интервал.")
            return

        self.interval_input.setDisabled(True)
        self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] Мониторинг WHEA запущен")
        write_log(f"Мониторинг WHEA запущен с интервалом {interval} секунд")
        self.monitor_start_time = datetime.now()
        self.timer.start(interval * 1000)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.last_error_count = 0

    def stop_monitor(self):
        self.timer.stop()
        self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] Мониторинг WHEA остановлен")
        write_log("Мониторинг WHEA остановлен")
        self.interval_input.setDisabled(False)
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.last_error_count = 0

    def check_whea_events(self):
        server = 'localhost'
        log_type = 'System'
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

        try:
            hand = win32evtlog.OpenEventLog(server, log_type)
            write_log(f"Открыт журнал событий Windows: {log_type} на сервере {server}")
        except Exception as e:
            err = f"Ошибка открытия журнала событий: {e}"
            self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] {err}")
            write_log(err)
            return

        count = 0
        found_event_ids = set()
        sample_events_info = []

        while True:
            try:
                events = win32evtlog.ReadEventLog(hand, flags, 0)
                if events:
                    write_log(f"Прочитано {len(events)} событий из журнала")
                else:
                    write_log("Журнал событий пуст")
            except Exception as e:
                err = f"Ошибка чтения журнала событий: {e}"
                self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] {err}")
                write_log(err)
                break

            if not events:
                write_log("Достигнут конец журнала событий.")
                break

            for ev in events[:5]:
                ev_dict = event_to_dict(ev)
                sample_events_info.append(ev_dict)

            for event in events:
                try:
                    event_time = datetime.fromtimestamp(pywintypes.Time(event.TimeGenerated).timestamp())
                except Exception:
                    continue

                if event_time < self.monitor_start_time:
                    continue

                event_id = event.EventID & 0xFFFF

                if event_id not in self.WHEA_EVENT_IDS:
                    continue

                count += 1
                found_event_ids.add(event_id)

            break

        write_log(f"Структура первых прочитанных событий (макс 5): {json.dumps(sample_events_info, ensure_ascii=False, indent=2)}")
        write_log(f"Всего найдено ошибок/предупреждений WHEA: {count}, Коды ошибок: {sorted(found_event_ids)}")

        timestamp = datetime.now().strftime('%H:%M:%S')

        if count > 0 and count != self.last_error_count:
            codes_str = ', '.join(str(eid) for eid in sorted(found_event_ids))
            msg = f"[{timestamp}] Найдена ошибка WHEA ({count}). Коды ошибок: {codes_str}"
            self.log_output.append(msg)
            write_log(msg)
            self.handle_trigger()
            self.last_error_count = count
        elif count == 0 and self.last_error_count != 0:
            self.last_error_count = 0
            write_log(f"[{timestamp}] Ошибок WHEA не найдено, счетчик ошибок сброшен.")

    def handle_trigger(self):
        cfg = self.trigger_config

        if cfg.get("message_enabled", False):
            if cfg.get("message_custom_enabled", False) and cfg.get("message_text", "").strip():
                msg = cfg["message_text"].strip()
            else:
                msg = "Обнаружена ошибка WHEA"
            show_messagebox(msg)
            write_log(f"Отправлено сообщение: {msg}")

        if cfg.get("audio_enabled", False):
            idx = cfg.get("audio_selected_index")
            if idx in [0, 1, 2]:
                path = self.trigger_form.sounds[idx][1]
                if os.path.exists(path):
                    winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                    write_log(f"Воспроизведён звук: {path}")
                    return
            winsound.MessageBeep(winsound.MB_OK)
            write_log("Воспроизведён стандартный звуковой сигнал")

        if cfg.get("execute_enabled", False):
            path = cfg.get("execute_path")
            args = cfg.get("execute_args", "")
            if path and os.path.exists(path):
                try:
                    if args.strip():
                        os.startfile(f'"{path}" {args}')
                    else:
                        os.startfile(path)
                    winsound.MessageBeep(winsound.MB_OK)
                    write_log(f"Запущена программа: {path} с аргументами: {args}")
                except Exception as e:
                    err = f"Ошибка запуска программы: {e}"
                    self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] {err}")
                    write_log(err)
            else:
                write_log("Путь к программе не задан или не существует")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WheaMonitorApp()
    window.show()
    sys.exit(app.exec())
