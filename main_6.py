"""Путевые листы Менеджер - главный модуль"""

import sys
import os
import time

# Добавь в начало файла с импортами

# Логирование печати
try:
    from utils.print_logger import print_logger

    PRINT_LOGGER_AVAILABLE = True
    print("✅ PrintLogger импортирован")
except ImportError as e:
    PRINT_LOGGER_AVAILABLE = False
    print(f"⚠️ PrintLogger не найден: {e}")


    # Создаем заглушку
    class PrintLoggerStub:
        def log_operation(self, *args, **kwargs): pass

        def log_print_dialog_open(self, *args, **kwargs): pass

        def log_print_start(self, *args, **kwargs): pass

        def log_print_complete(self, *args, **kwargs): pass

        def log_file_created(self, *args, **kwargs): pass

        def log_excel_opened(self, *args, **kwargs): pass

        def log_print_dialog_closed(self, *args, **kwargs): pass


    print_logger = PrintLoggerStub()

# В вашей основной программе
from modules.simple_excel_printer_final import print_waybill_simple

# ===== ФИКС: Устанавливаем пути ПЕРВЫМ ДЕЛОМ =====
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)
sys.path.insert(0, os.path.join(script_dir, "utils"))
sys.path.insert(0, os.path.join(script_dir, "modules"))

# ===== ТЕПЕРЬ ВСЕ ОСТАЛЬНЫЕ ИМПОРТЫ =====
import logging
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QTableWidget, QTableWidgetItem, QPushButton,
    QLabel, QLineEdit, QComboBox, QDateEdit, QDateTimeEdit, QTextEdit,
    QMessageBox, QFileDialog, QGroupBox, QFormLayout,
    QSpinBox, QDoubleSpinBox, QCheckBox, QStatusBar,
    QHeaderView, QSplitter, QDialog, QDialogButtonBox, QGridLayout
)
from PySide6.QtCore import Qt, QDate, QSettings, QTimer, QDateTime
from PySide6.QtGui import QFont, QAction, QColor

try:
    from modules.smart_printer import SmartWaybillPrinter, print_waybill_smart, preview_waybill_smart
    from modules.mapping_editor import edit_mapping_dialog

    SMART_PRINTER_AVAILABLE = True
    print("✅ Умный принтер импортирован")
except ImportError as e:
    SMART_PRINTER_AVAILABLE = False
    print(f"⚠️ Умный принтер не найден: {e}")

# Настройка логирования
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

print("=== DEBUG: Запуск main.py ===")

# ===== ИМПОРТ ВСЕХ НЕОБХОДИМЫХ КЛАССОВ =====
try:
    from utils.waybill_exporter import Exporter

    print("✅ Exporter импортирован")
except ImportError as e:
    print(f"❌ Ошибка импорта Exporter: {e}")


    class Exporter:
        def __init__(self):
            print("⚠️ Заглушка Exporter")

try:
    from utils.template_manager import TemplateManager

    print("✅ TemplateManager импортирован")
except ImportError as e:
    print(f"❌ Ошибка импорта TemplateManager: {e}")


    class TemplateManager:
        def __init__(self):
            print("⚠️ Заглушка TemplateManager")

try:
    from utils.numbering_manager import NumberingSettingsDialog, NumberingInfoWidget

    print("✅ NumberingSettingsDialog, NumberingInfoWidget импортированы")
except ImportError as e:
    print(f"❌ Ошибка импорта NumberingManager: {e}")


    class NumberingSettingsDialog:
        pass


    class NumberingInfoWidget:
        pass

try:
    from modules.excel_printer import ExcelWaybillPrinter

    print("✅ ExcelWaybillPrinter импортирован")
except ImportError as e:
    print(f"❌ Ошибка импорта ExcelWaybillPrinter: {e}")


    class ExcelWaybillPrinter:
        def __init__(self):
            print("⚠️ Заглушка ExcelWaybillPrinter")

try:
    from modules.waybill_printer_1c import print_waybill_1c, preview_waybill_1c

    PRINT_1C_AVAILABLE = True
    print("✅ waybill_printer_1c импортирован")
except ImportError:
    PRINT_1C_AVAILABLE = False
    print("⚠️ waybill_printer_1c не найден")

try:
    import pandas as pd

    PANDAS_AVAILABLE = True
    print("✅ pandas импортирован")
except ImportError:
    PANDAS_AVAILABLE = False
    print("⚠️ pandas не установлен, импорт Excel будет ограничен")

# В начале main_6.py должны быть:
import matplotlib

matplotlib.use('Qt5Agg')  # или 'QtAgg'

# Или для Qt Charts:


# Импорт simple_excel_printer
try:
    from modules.simple_excel_printer import print_waybill_simple, get_simple_printer

    SIMPLE_PRINTER_AVAILABLE = True
    print("✅ simple_excel_printer импортирован")

    # Тестируем принтер
    printer = get_simple_printer()
    if printer.check_template():
        print("   ✅ Шаблон 'образецфинал.xlsx' найден")
    else:
        print("   ⚠️ Шаблон 'образецфинал.xlsx' не найден")

except ImportError as e:
    print(f"⚠️ simple_excel_printer не найден: {e}")
    SIMPLE_PRINTER_AVAILABLE = False

# Замените импорт отчетов на:
try:
    from modules.simple_reports_fixed import SimpleReportsDialogFixed as SimpleReportsDialog

    REPORTS_AVAILABLE = True
    print("✅ SimpleReportsDialogFixed импортирован")
except ImportError as e:
    print(f"⚠️ SimpleReportsDialogFixed не найден: {e}")
    REPORTS_AVAILABLE = False

# В секции импортов main_6.py добавьте:
try:
    from modules.reports_dialog import ReportsDialog

    REPORTS_AVAILABLE = True
    print("✅ ReportsDialog импортирован")
except ImportError as e:
    REPORTS_AVAILABLE = False
    print(f"⚠️ ReportsDialog не найден: {e}")

# В секции импортов main_6.py:
try:
    # Пробуем импортировать исправленную версию
    from modules.simple_excel_printer_fixed import print_waybill_simple, get_simple_printer, simple_printer

    SIMPLE_PRINTER_AVAILABLE = True
    print("✅ simple_excel_printer (исправленный) импортирован")

    # Тестируем
    printer = get_simple_printer()
    if printer and printer.check_template():
        print("   ✅ Шаблон 'образецфинал.xlsx' найден")
        print(f"   📁 Папка печати: {printer.prints_dir}")
    else:
        print("   ⚠️ Шаблон не найден, будет создан простой файл")

except ImportError as e:
    print(f"⚠️ simple_excel_printer не найден: {e}")
    SIMPLE_PRINTER_AVAILABLE = False


    # Заглушки
    def print_waybill_simple(waybill_data=None, **kwargs):
        print("⚠️ simple_excel_printer не доступен")
        return None


    def get_simple_printer():
        return None

        simple_printer = None


    # Создаём заглушки для совместимости
    def print_waybill_simple(data, output_path=None):
        print("⚠️ simple_excel_printer не доступен")
        return None


    def get_simple_printer():
        return None

try:
    from modules.advanced_printer import print_waybill_advanced, get_print_preview

    ADVANCED_PRINTER_AVAILABLE = True
    print("✅ advanced_printer импортирован")
except ImportError:
    ADVANCED_PRINTER_AVAILABLE = False
    print("⚠️ advanced_printer не найден")

try:
    from modules.excel_preview_printer import print_with_preview

    PREVIEW_PRINTER_AVAILABLE = True
    print("✅ excel_preview_printer импортирован")
except ImportError:
    PREVIEW_PRINTER_AVAILABLE = False
    print("⚠️ excel_preview_printer не найден")

try:
    from OksanaPutList.database_fixed import DatabaseManager, recreate_database

    print("✅ database_fixed импортирован")
except ImportError as e:
    print(f"❌ Ошибка импорта database_fixed: {e}")


    class DatabaseManager:
        def __init__(self):
            print("⚠️ Заглушка DatabaseManager")


    # def recreate_database():
    #     print("⚠️ Заглушка recreate_database")
    def recreate_database():
        """Пересоздание базы данных"""
        try:
            db_path = "waybills.db"

            # Удаляем старую БД если существует
            if os.path.exists(db_path):
                os.remove(db_path)
                print(f"🗑️ Удалена старая база данных: {db_path}")

            # Создаем новую БД
            db = DatabaseManager()

            # Создаем таблицы
            db.create_tables()

            # Вставляем начальные данные
            db._insert_initial_data()

            # Инициализируем настройки нумерации
            db.update_numbering_settings({
                'prefix': 'ПЛ-',
                'current_number': 1,
                'digits': 6,
                'auto_increment': True,
                'reset_period': 'never',
                'last_reset_date': None
            })

            print("✅ База данных пересоздана успешно")
            return db

        except Exception as e:
            print(f"❌ Ошибка пересоздания базы данных: {e}")
            raise

print("\n=== DEBUG: Все импорты завершены ===\n")

os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'

# Проверьте какие модули НАЙДЕНЫ, а какие НЕТ:

print("=== ПРОВЕРКА ИМПОРТОВ ПЕЧАТИ ===")
print(f"SMART_PRINTER_AVAILABLE: {SMART_PRINTER_AVAILABLE}")
print(f"PRINT_1C_AVAILABLE: {PRINT_1C_AVAILABLE}")
print(f"SIMPLE_PRINTER_AVAILABLE: {SIMPLE_PRINTER_AVAILABLE}")  # ← Скорее всего FALSE
print(f"ADVANCED_PRINTER_AVAILABLE: {ADVANCED_PRINTER_AVAILABLE}")
print(f"PREVIEW_PRINTER_AVAILABLE: {PREVIEW_PRINTER_AVAILABLE}")
print("================================\n")


class WaybillManagerApp(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()

        # Инициализация базы данных
        try:
            self.db = DatabaseManager()
        except Exception as e:
            print(f"Ошибка подключения к БД: {e}")
            print("Пытаемся пересоздать базу данных...")
            recreate_database()
            self.db = DatabaseManager()

        # Инициализация вспомогательных классов
        self.exporter = Exporter()
        self.printer = ExcelWaybillPrinter()
        self.template_manager = TemplateManager()

        # Текущий путевой лист
        self.current_waybill_id = None

        # Инициализация интерфейса
        self.init_ui()
        self.load_settings()
        self.load_data()

        # Настройка окна
        self.setMinimumSize(1200, 700)
        self.showMaximized()

        # Тестирование виджетов через 1 секунду
        QTimer.singleShot(1000, self.test_widgets_enabled)

    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Путевые листы Менеджер v1.0")
        self.setGeometry(100, 100, 1400, 800)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Главный layout
        main_layout = QVBoxLayout(central_widget)

        # Панель инструментов
        self.create_toolbar()

        # Разделитель для основной области
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая панель - список путевых листов
        left_panel = self.create_waybill_list_panel()
        splitter.addWidget(left_panel)

        # Правая панель - вкладки
        right_panel = self.create_tabs_panel()
        splitter.addWidget(right_panel)

        splitter.setSizes([400, 1000])
        main_layout.addWidget(splitter)
        main_layout.setStretchFactor(splitter, 1)

        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов")

        # Меню
        self.create_menu()

    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = self.addToolBar("Панель инструментов")

        # Кнопки
        new_action = QAction("Новый", self)
        new_action.triggered.connect(self.create_new_waybill)
        toolbar.addAction(new_action)

        save_action = QAction("Сохранить", self)
        save_action.triggered.connect(self.save_waybill)
        toolbar.addAction(save_action)

        print_action = QAction("Печать", self)
        print_action.triggered.connect(self.print_current_waybill)
        toolbar.addAction(print_action)

        toolbar.addSeparator()

        export_action = QAction("Экспорт", self)
        export_action.triggered.connect(self.export_data)
        toolbar.addAction(export_action)

        import_action = QAction("Импорт", self)
        import_action.triggered.connect(self.import_data)
        toolbar.addAction(import_action)

        preview_action = QAction("Предпросмотр", self)
        preview_action.triggered.connect(self.preview_current_waybill)
        toolbar.addAction(preview_action)

        return toolbar

    def create_waybill_list_panel(self):
        """Создание панели списка путевых листов"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)

        # Заголовок
        title = QLabel("📋 СПИСОК ПУТЕВЫХ ЛИСТОВ")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(title)

        # Панель фильтров
        filter_panel = QWidget()
        filter_layout = QVBoxLayout(filter_panel)
        filter_layout.setSpacing(5)

        # Строка 1: Быстрый поиск
        search_layout = QHBoxLayout()
        search_label = QLabel("🔍 Поиск:")
        search_label.setFixedWidth(50)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Номер, водитель, автомобиль...")
        self.search_edit.textChanged.connect(self.filter_waybills)

        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_edit)
        filter_layout.addLayout(search_layout)

        # Строка 2: Фильтры водителя и статуса
        filters_layout = QHBoxLayout()

        # Фильтр водителя
        driver_filter_layout = QHBoxLayout()
        driver_label = QLabel("Водитель:")
        driver_label.setFixedWidth(60)
        self.filter_driver = QComboBox()
        self.filter_driver.setFixedWidth(150)
        self.filter_driver.currentTextChanged.connect(self.filter_waybills)

        driver_filter_layout.addWidget(driver_label)
        driver_filter_layout.addWidget(self.filter_driver)
        filters_layout.addLayout(driver_filter_layout)

        # Фильтр механика - ДОБАВИМ ЭТОТ БЛОК
        mechanic_filter_layout = QHBoxLayout()
        mechanic_label = QLabel("Механик:")
        mechanic_label.setFixedWidth(60)
        self.filter_mechanic = QComboBox()
        self.filter_mechanic.setFixedWidth(150)
        self.filter_mechanic.currentTextChanged.connect(self.filter_waybills)

        mechanic_filter_layout.addWidget(mechanic_label)
        mechanic_filter_layout.addWidget(self.filter_mechanic)
        filters_layout.addLayout(mechanic_filter_layout)

        # Фильтр статуса
        status_filter_layout = QHBoxLayout()
        status_label = QLabel("Статус:")
        status_label.setFixedWidth(50)
        self.filter_status = QComboBox()
        self.filter_status.setFixedWidth(120)
        self.filter_status.addItems(["Все", "Черновик", "Активен", "Завершен", "Архив", "Проведен"])
        self.filter_status.currentTextChanged.connect(self.filter_waybills)

        status_filter_layout.addWidget(status_label)
        status_filter_layout.addWidget(self.filter_status)
        filters_layout.addLayout(status_filter_layout)

        filters_layout.addStretch()
        filter_layout.addLayout(filters_layout)
        layout.addWidget(filter_panel)

        # Таблица путевых листов
        self.waybill_table = QTableWidget()
        self.waybill_table.setColumnCount(8)  # Увеличил до 8 колонок
        self.waybill_table.setHorizontalHeaderLabels([
            "✅", "ID", "Дата", "Водитель", "Автомобиль", "Статус", "Механик", "Организация"  # Добавил "Механик"
        ])

        # Настройка ширины колонок
        header = self.waybill_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.waybill_table.setColumnWidth(0, 30)

        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.waybill_table.setColumnWidth(1, 50)

        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.waybill_table.setColumnWidth(2, 80)

        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.waybill_table.setColumnWidth(5, 100)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)  # Механик
        self.waybill_table.setColumnWidth(6, 120)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Stretch)

        # Настройка внешнего вида
        self.waybill_table.setAlternatingRowColors(True)
        self.waybill_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.waybill_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.waybill_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        # Подключение сигналов
        self.waybill_table.itemDoubleClicked.connect(self.load_waybill)
        self.waybill_table.cellClicked.connect(self.on_waybill_cell_clicked)
        self.waybill_table.itemSelectionChanged.connect(self.on_waybill_selected)

        layout.addWidget(self.waybill_table)

        # Панель кнопок управления
        button_panel = QWidget()
        button_layout = QHBoxLayout(button_panel)
        button_layout.setSpacing(5)

        # Кнопки с иконками
        buttons_config = [
            ("📝", "Новый", self.create_new_waybill, "Создать новый путевой лист", "#4CAF50"),
            ("📂", "Открыть", self.open_selected_waybill, "Открыть выбранный путевой лист", "#2196F3"),
            ("✅", "Провести", self.conduct_waybill, "Провести выбранный путевой лист", "#FF9800"),
            ("🗑️", "Удалить", self.delete_waybill, "Удалить выбранный путевой лист", "#f44336"),
            ("🔄", "Обновить", self.refresh_waybill_list, "Обновить список путевых листов", "#9E9E9E")
        ]

        for icon, text, slot, tooltip, color in buttons_config:
            btn = QPushButton(f"{icon} {text}")
            btn.clicked.connect(slot)
            btn.setToolTip(tooltip)
            btn.setStyleSheet(f"""
                QPushButton {{
                    padding: 6px 12px;
                    border-radius: 4px;
                    background-color: {color};
                    color: white;
                    font-weight: bold;
                    border: none;
                }}
                QPushButton:hover {{
                    background-color: {self.darken_color(color)};
                }}
                QPushButton:pressed {{
                    background-color: {self.darken_color(color, 40)};
                }}
            """)
            button_layout.addWidget(btn)

        button_layout.addStretch()
        layout.addWidget(button_panel)

        # Панель статистики
        stats_panel = QWidget()
        stats_layout = QHBoxLayout(stats_panel)
        stats_layout.setSpacing(15)

        self.stats_total = QLabel("📊 Всего: 0")
        self.stats_draft = QLabel("📋 Черновики: 0")
        self.stats_active = QLabel("🚗 Активные: 0")
        self.stats_completed = QLabel("✅ Завершены: 0")

        # Стили для статистики
        stats_style = """
            QLabel {
                padding: 5px 10px;
                border-radius: 3px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                font-weight: bold;
            }
        """

        for stat_label in [self.stats_total, self.stats_draft, self.stats_active, self.stats_completed]:
            stat_label.setStyleSheet(stats_style)
            stats_layout.addWidget(stat_label)

        stats_layout.addStretch()
        layout.addWidget(stats_panel)

        return panel

    def darken_color(self, hex_color, amount=20):
        """Затемнение цвета для эффекта hover"""
        hex_color = hex_color.lstrip('#')
        r = max(0, int(hex_color[0:2], 16) - amount)
        g = max(0, int(hex_color[2:4], 16) - amount)
        b = max(0, int(hex_color[4:6], 16) - amount)
        return f'#{r:02x}{g:02x}{b:02x}'

    def create_tabs_panel(self):
        """Создание панели с вкладками"""
        self.tab_widget = QTabWidget()

        # Вкладка редактирования путевого листа
        self.edit_tab = self.create_waybill_edit_tab()
        self.tab_widget.addTab(self.edit_tab, "Путевой лист")

        # Вкладка водителей
        self.drivers_tab = self.create_drivers_tab()
        self.tab_widget.addTab(self.drivers_tab, "Водители")

        # Вкладка механиков
        self.mechanics_tab = self.create_mechanics_tab()
        self.tab_widget.addTab(self.mechanics_tab, "Механики")

        # Вкладка автомобилей
        self.vehicles_tab = self.create_vehicles_tab()
        self.tab_widget.addTab(self.vehicles_tab, "Автомобили")

        # Вкладка предприятий
        self.companies_tab = self.create_companies_tab()
        self.tab_widget.addTab(self.companies_tab, "Предприятия")

        # Вкладка настроек
        self.settings_tab = self.create_settings_tab()
        self.tab_widget.addTab(self.settings_tab, "Настройки")

        return self.tab_widget

    def create_waybill_edit_tab(self):
        """Создание вкладки редактирования путевого листа"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # 1. ВЕРХНЯЯ ПАНЕЛЬ: Основная информация
        top_panel = QWidget()
        top_layout = QHBoxLayout(top_panel)

        # Номер и дата
        num_date_widget = QWidget()
        num_date_layout = QFormLayout(num_date_widget)

        self.waybill_id = QLineEdit()
        self.waybill_id.setReadOnly(True)
        self.waybill_id.setFixedWidth(150)

        self.waybill_date = QDateEdit()
        self.waybill_date.setDate(QDate.currentDate())
        self.waybill_date.setCalendarPopup(True)
        self.waybill_date.setDisplayFormat("dd.MM.yyyy")
        self.waybill_date.setFixedWidth(150)

        num_date_layout.addRow("Номер:", self.waybill_id)
        num_date_layout.addRow("Дата:", self.waybill_date)

        # Статус
        status_widget = QWidget()
        status_layout = QFormLayout(status_widget)

        self.waybill_status = QComboBox()
        self.waybill_status.addItems(["Черновик", "Активен", "Завершен", "Архив", "Проведен"])
        self.waybill_status.setFixedWidth(150)

        status_layout.addRow("Статус:", self.waybill_status)

        top_layout.addWidget(num_date_widget)
        top_layout.addWidget(status_widget)
        top_layout.addStretch()

        layout.addWidget(top_panel)

        # 2. ОСНОВНЫЕ ДАННЫЕ
        main_data_group = QGroupBox("Основные данные")
        main_data_layout = QGridLayout(main_data_group)
        main_data_layout.setSpacing(15)

        # Организация (строка 0)
        main_data_layout.addWidget(QLabel("Организация*:"), 0, 0)
        self.company_combo_waybill = QComboBox()
        self.company_combo_waybill.addItem("Выберите предприятие...")
        self.company_combo_waybill.setMinimumWidth(300)
        main_data_layout.addWidget(self.company_combo_waybill, 0, 1)

        # Кнопка просмотра данных организации (строка 0, колонка 2)
        self.btn_view_company = QPushButton("📋 Просмотреть")
        self.btn_view_company.clicked.connect(self.view_company_details)
        self.btn_view_company.setFixedWidth(120)
        main_data_layout.addWidget(self.btn_view_company, 0, 2)

        # Водитель (строка 1)
        main_data_layout.addWidget(QLabel("Водитель*:"), 1, 0)
        self.driver_combo = QComboBox()
        self.driver_combo.addItem("Выберите водителя...")
        self.driver_combo.setMinimumWidth(300)
        main_data_layout.addWidget(self.driver_combo, 1, 1)

        # Кнопка просмотра данных водителя (строка 1, колонка 2)
        self.btn_view_driver = QPushButton("👤 Просмотреть")
        self.btn_view_driver.clicked.connect(self.view_driver_details)
        self.btn_view_driver.setFixedWidth(120)
        main_data_layout.addWidget(self.btn_view_driver, 1, 2)

        # Автомобиль (строка 2)
        main_data_layout.addWidget(QLabel("Автомобиль*:"), 2, 0)
        self.vehicle_combo = QComboBox()
        self.vehicle_combo.addItem("Выберите автомобиль...")
        self.vehicle_combo.setMinimumWidth(300)
        main_data_layout.addWidget(self.vehicle_combo, 2, 1)

        # Кнопка просмотра данных автомобиля (строка 2, колонка 2)
        self.btn_view_vehicle = QPushButton("🚗 Просмотреть")
        self.btn_view_vehicle.clicked.connect(self.view_vehicle_details)
        self.btn_view_vehicle.setFixedWidth(120)
        main_data_layout.addWidget(self.btn_view_vehicle, 2, 2)

        # Механик (строка 3)
        main_data_layout.addWidget(QLabel("Механик:"), 3, 0)
        self.mechanic_combo = QComboBox()
        self.mechanic_combo.addItem("Выберите механика...")
        self.mechanic_combo.setMinimumWidth(300)
        main_data_layout.addWidget(self.mechanic_combo, 3, 1)

        # Кнопка просмотра данных механика (строка 3, колонка 2)
        self.btn_view_mechanic = QPushButton("🔧 Просмотреть")
        self.btn_view_mechanic.clicked.connect(self.view_mechanic_details)
        self.btn_view_mechanic.setFixedWidth(120)
        main_data_layout.addWidget(self.btn_view_mechanic, 3, 2)

        layout.addWidget(main_data_group)

        # 3. РАЗДЕЛИТЕЛЬ НА ДВЕ КОЛОНКИ: Технические данные
        columns_splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая колонка: Одометр и топливо
        left_column = QWidget()
        left_layout = QVBoxLayout(left_column)
        left_layout.setSpacing(10)

        # Одометр
        odo_group = QGroupBox("Одометр")
        odo_layout = QFormLayout(odo_group)

        self.vehicle_odo_start = QSpinBox()
        self.vehicle_odo_start.setMaximum(999999)
        self.vehicle_odo_start.setSuffix(" км")
        self.vehicle_odo_start.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.vehicle_odo_start.setValue(0)
        self.vehicle_odo_start.setMinimumWidth(150)

        self.vehicle_odo_end = QSpinBox()
        self.vehicle_odo_end.setMaximum(999999)
        self.vehicle_odo_end.setSuffix(" км")
        self.vehicle_odo_end.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.vehicle_odo_end.setValue(0)
        self.vehicle_odo_end.setMinimumWidth(150)

        odo_layout.addRow("При выезде:", self.vehicle_odo_start)
        odo_layout.addRow("При возврате:", self.vehicle_odo_end)

        left_layout.addWidget(odo_group)

        # Топливо
        fuel_group = QGroupBox("Топливо")
        fuel_layout = QFormLayout(fuel_group)

        self.fuel_type = QComboBox()
        self.fuel_type.addItems(["Бензин", "Дизельное топливо", "Газ", "Электричество", "Гибрид"])
        self.fuel_type.setCurrentIndex(0)

        self.fuel_start = QDoubleSpinBox()
        self.fuel_start.setMaximum(1000)
        self.fuel_start.setSuffix(" л")
        self.fuel_start.setDecimals(1)
        self.fuel_start.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.fuel_start.setValue(0.0)
        self.fuel_start.setMinimumWidth(120)

        self.fuel_issued = QDoubleSpinBox()
        self.fuel_issued.setMaximum(1000)
        self.fuel_issued.setSuffix(" л")
        self.fuel_issued.setDecimals(1)
        self.fuel_issued.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.fuel_issued.setValue(0.0)
        self.fuel_issued.setMinimumWidth(120)

        self.fuel_end = QDoubleSpinBox()
        self.fuel_end.setMaximum(1000)
        self.fuel_end.setSuffix(" л")
        self.fuel_end.setDecimals(1)
        self.fuel_end.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.fuel_end.setValue(0.0)
        self.fuel_end.setMinimumWidth(120)

        fuel_layout.addRow("Тип:", self.fuel_type)
        fuel_layout.addRow("Остаток (выезд):", self.fuel_start)
        fuel_layout.addRow("Выдано:", self.fuel_issued)
        fuel_layout.addRow("Остаток (возврат):", self.fuel_end)

        left_layout.addWidget(fuel_group)
        left_layout.addStretch()
        columns_splitter.addWidget(left_column)

        # Правая колонка: Медосмотры, Маршрут, Примечания
        right_column = QWidget()
        right_layout = QVBoxLayout(right_column)
        right_layout.setSpacing(8)  # Уменьшили spacing с 10 до 8

        # Медосмотры - делаем выше
        medical_group = QGroupBox("Медосмотры")
        medical_layout = QFormLayout(medical_group)
        medical_layout.setSpacing(8)  # Уменьшаем spacing

        self.medical_pre_date = QDateTimeEdit()
        self.medical_pre_date.setDateTime(datetime.now())
        self.medical_pre_date.setDisplayFormat("dd.MM.yyyy HH:mm")
        self.medical_pre_date.setCalendarPopup(True)
        self.medical_pre_date.setMinimumWidth(200)
        self.medical_pre_date.setFixedHeight(30)  # Фиксированная высота

        self.medical_pre_doctor = QLineEdit()
        self.medical_pre_doctor.setPlaceholderText("ФИО врача")
        self.medical_pre_doctor.setMinimumWidth(200)
        self.medical_pre_doctor.setFixedHeight(30)

        self.medical_post_date = QDateTimeEdit()
        self.medical_post_date.setDateTime(datetime.now())
        self.medical_post_date.setDisplayFormat("dd.MM.yyyy HH:mm")
        self.medical_post_date.setCalendarPopup(True)
        self.medical_post_date.setMinimumWidth(200)
        self.medical_post_date.setFixedHeight(30)

        self.medical_post_doctor = QLineEdit()
        self.medical_post_doctor.setPlaceholderText("ФИО врача")
        self.medical_post_doctor.setMinimumWidth(200)
        self.medical_post_doctor.setFixedHeight(30)

        medical_layout.addRow("Предрейсовый:", self.medical_pre_date)
        medical_layout.addRow("Врач:", self.medical_pre_doctor)
        medical_layout.addRow("Послерейсовый:", self.medical_post_date)
        medical_layout.addRow("Врач:", self.medical_post_doctor)

        right_layout.addWidget(medical_group, stretch=2)  # Медосмотры занимают больше места

        # Маршрут - делаем меньше
        route_group = QGroupBox("Маршрут")
        route_layout = QVBoxLayout(route_group)

        self.route_text = QTextEdit()
        self.route_text.setPlaceholderText("Укажите маршрут следования...")
        self.route_text.setFixedHeight(50)  # Фиксированная высота
        route_layout.addWidget(self.route_text)

        right_layout.addWidget(route_group, stretch=1)  # Меньший приоритет

        # Примечания - делаем еще меньше
        notes_group = QGroupBox("Примечания")
        notes_layout = QVBoxLayout(notes_group)

        self.notes_text = QTextEdit()
        self.notes_text.setPlaceholderText("Опоздания, простои, заезды в гараж...")
        self.notes_text.setFixedHeight(40)  # Фиксированная высота
        notes_layout.addWidget(self.notes_text)

        right_layout.addWidget(notes_group, stretch=1)  # Меньший приоритет

        right_layout.addStretch()

        columns_splitter.addWidget(right_column)
        # columns_splitter.setSizes([350, 650])
        columns_splitter.setSizes([250, 750])  # Больше места правой колонке
        layout.addWidget(columns_splitter)

        # 4. НИЖНЯЯ ПАНЕЛЬ: Кнопки действий
        bottom_panel = QWidget()
        bottom_layout = QHBoxLayout(bottom_panel)

        self.btn_save = QPushButton("💾 Сохранить")
        self.btn_save.clicked.connect(self.save_waybill)
        self.btn_save.setStyleSheet("font-weight: bold; padding: 8px 20px;")

        self.btn_print = QPushButton("🖨️ Печать")
        self.btn_print.clicked.connect(self.print_current_waybill)
        self.btn_print.setStyleSheet("padding: 8px 20px;")

        self.btn_preview = QPushButton("👁️ Предпросмотр")
        self.btn_preview.clicked.connect(self.preview_current_waybill)
        self.btn_preview.setStyleSheet("padding: 8px 20px;")

        self.btn_clear = QPushButton("🆕 Новый")
        self.btn_clear.clicked.connect(self.create_new_waybill)
        self.btn_clear.setStyleSheet("padding: 8px 20px;")

        bottom_layout.addWidget(self.btn_save)
        bottom_layout.addWidget(self.btn_print)
        bottom_layout.addWidget(self.btn_preview)
        bottom_layout.addWidget(self.btn_clear)
        bottom_layout.addStretch()

        layout.addWidget(bottom_panel)

        # Подключение сигналов
        self.company_combo_waybill.currentIndexChanged.connect(self.on_company_changed)
        self.driver_combo.currentIndexChanged.connect(self.on_driver_changed)
        self.vehicle_combo.currentIndexChanged.connect(self.on_vehicle_changed)

        return tab

    def view_company_details(self):
        """Просмотр деталей организации"""
        company_id = self.company_combo_waybill.currentData()
        if not company_id:
            QMessageBox.information(self, "Предприятие", "Выберите предприятие для просмотра")
            return

        company = self.db.get_company(company_id)
        if company:
            details = f"""
            <h3>{company['name']}</h3>
            <b>ИНН:</b> {company.get('inn', 'не указан')}<br>
            <b>КПП:</b> {company.get('kpp', 'не указан')}<br>
            <b>ОГРН:</b> {company.get('ogrn', 'не указан')}<br>
            <b>Адрес:</b> {company.get('address', 'не указан')}<br>
            <b>Телефон:</b> {company.get('phone', 'не указан')}<br>
            <b>E-mail:</b> {company.get('email', 'не указан')}<br>
            <b>Директор:</b> {company.get('director', 'не указан')}<br>
            """
            QMessageBox.information(self, "Данные предприятия", details)
        else:
            QMessageBox.warning(self, "Ошибка", "Предприятие не найдено в базе данных")

    def view_driver_details(self):
        """Просмотр деталей водителя"""
        driver_id = self.driver_combo.currentData()
        if not driver_id:
            QMessageBox.information(self, "Водитель", "Выберите водителя для просмотра")
            return

        driver = self.db.get_driver(driver_id)
        if driver:
            # Форматирование даты выдачи в/у
            license_issue_date = driver.get('license_issue_date', 'не указана')
            if license_issue_date and license_issue_date != 'не указана':
                try:
                    date_obj = datetime.strptime(license_issue_date, "%Y-%m-%d")
                    license_issue_date = date_obj.strftime("%d.%m.%Y")
                except:
                    pass

            details = f"""
            <h3>{driver['fio']}</h3>
            <b>СНИЛС:</b> {driver.get('snils', 'не указан')}<br>
            <b>Водительское удостоверение:</b> {driver.get('license', 'не указан')}<br>
            <b>Дата выдачи в/у:</b> {license_issue_date}<br>
            <b>Класс:</b> {driver.get('license_class', 'не указан')}<br>
            <b>Телефон:</b> {driver.get('phone', 'не указан')}<br>
            <b>Дата медосмотра:</b> {driver.get('medical_date', 'не указана')}<br>
            <b>Адрес:</b> {driver.get('address', 'не указан')}<br>
            <b>Предприятие:</b> {driver.get('company_name', 'не указано')}<br>
            """
            QMessageBox.information(self, "Данные водителя", details)
        else:
            QMessageBox.warning(self, "Ошибка", "Водитель не найден в базе данных")

    def view_vehicle_details(self):
        """Просмотр деталей автомобиля"""
        vehicle_id = self.vehicle_combo.currentData()
        if not vehicle_id:
            QMessageBox.information(self, "Автомобиль", "Выберите автомобиль для просмотра")
            return

        vehicle = self.db.get_vehicle(vehicle_id)
        if vehicle:
            # Получаем данные предприятия и водителя
            company_name = "Не указано"
            if vehicle.get('company_id'):
                company = self.db.get_company(vehicle['company_id'])
                if company:
                    company_name = company['name']

            driver_name = "Не назначен"
            if vehicle.get('driver_id'):
                driver = self.db.get_driver(vehicle['driver_id'])
                if driver:
                    driver_name = driver['fio']

            details = f"""
            <h3>{vehicle['brand']} {vehicle.get('model', '')}</h3>
            <b>Гос. номер:</b> {vehicle.get('plate', 'не указан')}<br>
            <b>VIN:</b> {vehicle.get('vin', 'не указан')}<br>
            <b>Год выпуска:</b> {vehicle.get('year', 'не указан')}<br>
            <b>Пробег:</b> {vehicle.get('mileage', 0)} км<br>
            <b>Цвет:</b> {vehicle.get('color', 'не указан')}<br>
            <b>Предприятие:</b> {company_name}<br>
            <b>Водитель:</b> {driver_name}<br>
            """
            QMessageBox.information(self, "Данные автомобиля", details)
        else:
            QMessageBox.warning(self, "Ошибка", "Автомобиль не найден в базе данных")

    def view_mechanic_details(self):
        """Просмотр деталей механика"""
        mechanic_id = self.mechanic_combo.currentData()
        if not mechanic_id:
            QMessageBox.information(self, "Механик", "Выберите механика для просмотра")
            return

        mechanic = self.db.get_mechanic(mechanic_id)
        if mechanic:
            # Форматирование даты выдачи удостоверения
            license_date = mechanic.get('license_date', 'не указана')
            if license_date and license_date != 'не указана':
                try:
                    date_obj = datetime.strptime(license_date, "%Y-%m-%d")
                    license_date = date_obj.strftime("%d.%m.%Y")
                except:
                    pass

            details = f"""
            <h3>{mechanic['fio']}</h3>
            <b>Должность:</b> {mechanic.get('position', 'не указана')}<br>
            <b>Телефон:</b> {mechanic.get('phone', 'не указан')}<br>
            <b>E-mail:</b> {mechanic.get('email', 'не указан')}<br>
            <b>Удостоверение №:</b> {mechanic.get('license_number', 'не указан')}<br>
            <b>Дата выдачи:</b> {license_date}<br>
            <b>Предприятие:</b> {mechanic.get('company_name', 'не указано')}<br>
            """
            QMessageBox.information(self, "Данные механика", details)
        else:
            QMessageBox.warning(self, "Ошибка", "Механик не найден в базе данных")

    def test_widgets_enabled(self):
        """Тестирование активности виджетов (для отладки)"""
        print("=== ТЕСТ АКТИВНОСТИ ВИДЖЕТОВ ===")
        print(f"company_combo_waybill enabled: {self.company_combo_waybill.isEnabled()}")
        print(f"driver_combo enabled: {self.driver_combo.isEnabled()}")
        print(f"vehicle_combo enabled: {self.vehicle_combo.isEnabled()}")
        print(f"mechanic_combo enabled: {self.mechanic_combo.isEnabled()}")
        print(f"fuel_type enabled: {self.fuel_type.isEnabled()}")
        print(f"medical_pre_date enabled: {self.medical_pre_date.isEnabled()}")

    def create_drivers_tab(self):
        """Создание вкладки управления водителями"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Таблица водителей
        self.drivers_table = QTableWidget()
        self.drivers_table.setColumnCount(9)  # Увеличил до 9 колонок
        self.drivers_table.setHorizontalHeaderLabels([
            "ID", "ФИО", "СНИЛС", "В/у №", "Класс", "Дата выдачи", "Телефон", "Медосмотр до", "Автомобили"
        ])
        self.drivers_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.drivers_table)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_driver)

        btn_edit = QPushButton("Редактировать")
        btn_edit.clicked.connect(self.edit_driver)

        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self.delete_driver)

        btn_import = QPushButton("Импорт из Excel")
        btn_import.clicked.connect(self.import_drivers)

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.clicked.connect(self.export_drivers)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_export)

        layout.addLayout(btn_layout)

        return tab

    def print_current_waybill(self):
        """Единая функция печати - открывает диалог печати/предпросмотра"""
        if not self.current_waybill_id:
            QMessageBox.warning(self, "Предупреждение", "Нет активного путевого листа")
            return

        try:
            # Получаем данные путевого листа
            waybill_data = self.db.get_waybill(self.current_waybill_id)
            if not waybill_data:
                QMessageBox.warning(self, "Ошибка", "Путевой лист не найден")
                return

            # Обновляем данные из формы
            self.update_waybill_data_from_form(waybill_data)

            print(f"DEBUG: Открываем диалог печати для ПЛ №{waybill_data.get('number', '')}")

            # Открываем диалог печати с предпросмотром
            dialog = PrintPreviewDialog(self, waybill_data, self.db)

            # Позиционируем окно рядом с главным
            main_rect = self.geometry()
            dialog.move(main_rect.right() - 850, main_rect.top() + 50)

            dialog.exec()

            self.status_bar.showMessage("Операция печати завершена")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при подготовке печати: {str(e)}")
            import traceback
            traceback.print_exc()

    def create_mechanics_tab(self):
        """Создание вкладки управления механиками"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Таблица механиков
        self.mechanics_table = QTableWidget()
        self.mechanics_table.setColumnCount(8)
        self.mechanics_table.setHorizontalHeaderLabels([
            "ID", "ФИО", "Должность", "Телефон", "E-mail", "Удостоверение", "Дата выдачи", "Предприятие"
        ])
        self.mechanics_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.mechanics_table)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_mechanic)

        btn_edit = QPushButton("Редактировать")
        btn_edit.clicked.connect(self.edit_mechanic)

        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self.delete_mechanic)

        btn_import = QPushButton("Импорт из Excel")
        btn_import.clicked.connect(self.import_mechanics)

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.clicked.connect(self.export_mechanics)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_export)

        layout.addLayout(btn_layout)

        return tab

    def create_vehicles_tab(self):
        """Создание вкладки управления автомобилями"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Таблица автомобилей
        self.vehicles_table = QTableWidget()
        self.vehicles_table.setColumnCount(9)
        self.vehicles_table.setHorizontalHeaderLabels([
            "ID", "Марка", "Модель", "Гос. номер", "VIN", "Год", "Пробег", "Водитель", "Предприятие"
        ])
        self.vehicles_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.vehicles_table)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_vehicle)

        btn_edit = QPushButton("Редактировать")
        btn_edit.clicked.connect(self.edit_vehicle)

        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self.delete_vehicle)

        btn_import = QPushButton("Импорт из Excel")
        btn_import.clicked.connect(self.import_vehicles)

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.clicked.connect(self.export_vehicles)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_export)

        layout.addLayout(btn_layout)

        return tab

    def create_companies_tab(self):
        """Создание вкладки управления предприятиями"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Таблица предприятий
        self.companies_table = QTableWidget()
        self.companies_table.setColumnCount(7)
        self.companies_table.setHorizontalHeaderLabels([
            "ID", "Наименование", "ИНН", "КПП", "ОГРН", "Адрес", "Телефон"
        ])
        self.companies_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.companies_table)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_company)

        btn_edit = QPushButton("Редактировать")
        btn_edit.clicked.connect(self.edit_company)

        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self.delete_company)

        btn_import = QPushButton("Импорт из Excel")
        btn_import.clicked.connect(self.import_companies)

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.clicked.connect(self.export_companies)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_export)

        layout.addLayout(btn_layout)

        return tab

    def create_settings_tab(self):
        """Создание вкладки настроек"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Настройки печати
        print_group = QGroupBox("Настройки печати")
        print_layout = QFormLayout()

        self.print_printer = QComboBox()
        self.print_printer.addItems(["По умолчанию", "HP LaserJet", "Canon Pixma"])

        self.print_copies = QSpinBox()
        self.print_copies.setMinimum(1)
        self.print_copies.setMaximum(10)
        self.print_copies.setValue(1)

        self.print_orientation = QComboBox()
        self.print_orientation.addItems(["Книжная", "Альбомная"])

        print_layout.addRow("Принтер:", self.print_printer)
        print_layout.addRow("Копий:", self.print_copies)
        print_layout.addRow("Ориентация:", self.print_orientation)

        print_group.setLayout(print_layout)
        layout.addWidget(print_group)

        # Настройки организации
        org_settings_group = QGroupBox("Настройки организации по умолчанию")
        org_settings_layout = QFormLayout()

        self.default_org_name = QLineEdit()
        self.default_org_address = QTextEdit()
        self.default_org_address.setMaximumHeight(60)
        self.default_org_phone = QLineEdit()
        self.default_org_inn = QLineEdit()

        org_settings_layout.addRow("Наименование:", self.default_org_name)
        org_settings_layout.addRow("Адрес:", self.default_org_address)
        org_settings_layout.addRow("Телефон:", self.default_org_phone)
        org_settings_layout.addRow("ИНН:", self.default_org_inn)

        org_settings_group.setLayout(org_settings_layout)
        layout.addWidget(org_settings_group)

        # Пути сохранения
        paths_group = QGroupBox("Пути сохранения")
        paths_layout = QFormLayout()

        self.save_path = QLineEdit()
        btn_browse = QPushButton("Обзор")
        btn_browse.clicked.connect(self.browse_save_path)

        path_layout = QHBoxLayout()
        path_layout.addWidget(self.save_path)
        path_layout.addWidget(btn_browse)

        self.auto_backup = QCheckBox("Автоматическое резервное копирование")
        self.backup_interval = QSpinBox()
        self.backup_interval.setMinimum(1)
        self.backup_interval.setMaximum(30)
        self.backup_interval.setSuffix(" дней")
        self.backup_interval.setValue(7)

        paths_layout.addRow("Папка сохранения:", path_layout)
        paths_layout.addRow(self.auto_backup)
        paths_layout.addRow("Интервал бэкапа:", self.backup_interval)

        paths_group.setLayout(paths_layout)
        layout.addWidget(paths_group)

        # Кнопки сохранения настроек
        btn_layout = QHBoxLayout()
        btn_save_settings = QPushButton("Сохранить настройки")
        btn_save_settings.clicked.connect(self.save_settings)

        btn_reset_settings = QPushButton("Сбросить настройки")
        btn_reset_settings.clicked.connect(self.reset_settings)

        btn_layout.addWidget(btn_save_settings)
        btn_layout.addWidget(btn_reset_settings)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        layout.addStretch()

        return tab

    def create_menu(self):
        """Создание меню"""
        menubar = self.menuBar()

        # Меню Файл
        file_menu = menubar.addMenu("Файл")

        new_action = QAction("Новый путевой лист", self)
        new_action.triggered.connect(self.create_new_waybill)
        file_menu.addAction(new_action)

        open_action = QAction("Открыть...", self)
        open_action.triggered.connect(self.open_waybill)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        import_action = QAction("Импорт данных...", self)
        import_action.triggered.connect(self.import_data)
        file_menu.addAction(import_action)

        export_action = QAction("Экспорт данных...", self)
        export_action.triggered.connect(self.export_data)
        file_menu.addAction(export_action)

        file_menu.addSeparator()

        print_action = QAction("Печать...", self)
        print_action.triggered.connect(self.print_current_waybill)
        file_menu.addAction(print_action)

        preview_action = QAction("Предварительный просмотр...", self)
        preview_action.triggered.connect(self.preview_current_waybill)
        file_menu.addAction(preview_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Меню Справочники
        ref_menu = menubar.addMenu("Справочники")

        drivers_action = QAction("Водители", self)
        drivers_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(1))
        ref_menu.addAction(drivers_action)

        mechanics_action = QAction("Механики", self)
        mechanics_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(2))
        ref_menu.addAction(mechanics_action)

        vehicles_action = QAction("Автомобили", self)
        vehicles_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(3))
        ref_menu.addAction(vehicles_action)

        companies_action = QAction("Предприятия", self)
        companies_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(4))
        ref_menu.addAction(companies_action)

        # Меню Сервис
        service_menu = menubar.addMenu("Сервис")

        backup_action = QAction("Создать резервную копию", self)
        backup_action.triggered.connect(self.create_backup)
        service_menu.addAction(backup_action)

        restore_action = QAction("Восстановить из копии", self)
        restore_action.triggered.connect(self.restore_backup)
        service_menu.addAction(restore_action)

        service_menu.addSeparator()

        # В меню Сервис:
        reports_action = QAction("📈 Отчёты с графиками", self)
        reports_action.triggered.connect(self.show_reports)
        service_menu.addAction(reports_action)

        # Или отдельное меню "Отчёты":
        reports_menu = menubar.addMenu("📈 Отчёты")
        reports_action = QAction("Открыть отчёты", self)
        reports_action.triggered.connect(self.show_reports)
        reports_menu.addAction(reports_action)

        # Быстрые отчёты
        quick_reports = reports_menu.addMenu("Быстрые отчёты")
        quick_month = QAction("За текущий месяц", self)
        quick_month.triggered.connect(lambda: self.generate_quick_report('month'))
        quick_reports.addAction(quick_month)

        quick_week = QAction("За текущую неделю", self)
        quick_week.triggered.connect(lambda: self.generate_quick_report('week'))
        quick_reports.addAction(quick_week)

        quick_year = QAction("За текущий год", self)
        quick_year.triggered.connect(lambda: self.generate_quick_report('year'))
        quick_reports.addAction(quick_year)

        numbering_action = QAction("Настройки нумерации...", self)
        numbering_action.triggered.connect(self.show_numbering_settings)
        service_menu.addAction(numbering_action)

        history_action = QAction("Экспорт истории статусов", self)
        history_action.triggered.connect(self.export_history)
        service_menu.addAction(history_action)

        # Меню Справка
        help_menu = menubar.addMenu("Справка")

        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # В меню Сервис добавьте:
        mapping_action = QAction("📋 Редактировать маппинг печати...", self)
        mapping_action.triggered.connect(self.edit_print_mapping)
        service_menu.addAction(mapping_action)

        return service_menu

    def load_data(self):
        """Загрузка данных из базы в выпадающие списки"""
        print("DEBUG: Загрузка данных из БД...")

        try:
            # Сохраняем текущие выбранные ID
            current_company_id = self.company_combo_waybill.currentData() if self.company_combo_waybill.count() > 0 else None
            current_driver_id = self.driver_combo.currentData() if self.driver_combo.count() > 0 else None
            current_vehicle_id = self.vehicle_combo.currentData() if self.vehicle_combo.count() > 0 else None
            current_mechanic_id = self.mechanic_combo.currentData() if hasattr(self,
                                                                               'mechanic_combo') and self.mechanic_combo.count() > 0 else None

            # Загрузка предприятий
            companies = self.db.get_companies()
            print(f"DEBUG: Получено {len(companies)} предприятий из БД")

            self.company_combo_waybill.blockSignals(True)
            self.company_combo_waybill.clear()
            self.company_combo_waybill.addItem("Выберите предприятие...", None)

            for company in companies:
                self.company_combo_waybill.addItem(company['name'], company['id'])

            # Восстанавливаем выбранное значение
            if current_company_id:
                index = self.company_combo_waybill.findData(current_company_id)
                if index >= 0:
                    self.company_combo_waybill.setCurrentIndex(index)
                    print(f"DEBUG: Восстановлено предприятие с ID: {current_company_id}")
            self.company_combo_waybill.blockSignals(False)

            # Загрузка водителей
            drivers = self.db.get_drivers()
            print(f"DEBUG: Получено {len(drivers)} водителей из БД")

            self.driver_combo.blockSignals(True)
            self.filter_driver.blockSignals(True)

            self.driver_combo.clear()
            self.filter_driver.clear()

            self.driver_combo.addItem("Выберите водителя...", None)
            self.filter_driver.addItem("Все водители", None)

            for driver in drivers:
                self.driver_combo.addItem(driver['fio'], driver['id'])
                self.filter_driver.addItem(driver['fio'], driver['id'])

            self.driver_combo.blockSignals(False)
            self.filter_driver.blockSignals(False)

            # Загрузка механиков для фильтра (ДОБАВЛЕНО)
            mechanics = self.db.get_mechanics()
            print(f"DEBUG: Получено {len(mechanics)} механиков из БД")

            self.filter_mechanic.blockSignals(True)  # Блокируем сигналы чтобы избежать рекурсии
            self.filter_mechanic.clear()
            self.filter_mechanic.addItem("Все механики", None)

            for mechanic in mechanics:
                self.filter_mechanic.addItem(mechanic['fio'], mechanic['id'])

            self.filter_mechanic.blockSignals(False)

            # Восстанавливаем выбранное значение
            if current_driver_id:
                index = self.driver_combo.findData(current_driver_id)
                if index >= 0:
                    self.driver_combo.setCurrentIndex(index)
            self.driver_combo.blockSignals(False)
            self.filter_driver.blockSignals(False)

            # Загрузка автомобилей
            vehicles = self.db.get_vehicles()
            print(f"DEBUG: Получено {len(vehicles)} автомобилей из БД")

            self.vehicle_combo.blockSignals(True)
            self.vehicle_combo.clear()
            self.vehicle_combo.addItem("Выберите автомобиль...", None)

            for vehicle in vehicles:
                display_text = f"{vehicle['brand']} {vehicle.get('model', '')} ({vehicle['plate']})"
                self.vehicle_combo.addItem(display_text, vehicle['id'])

            # Восстанавливаем выбранное значение
            if current_vehicle_id:
                index = self.vehicle_combo.findData(current_vehicle_id)
                if index >= 0:
                    self.vehicle_combo.setCurrentIndex(index)
            self.vehicle_combo.blockSignals(False)

            # Загрузка механиков
            mechanics = self.db.get_mechanics()
            print(f"DEBUG: Получено {len(mechanics)} механиков из БД")

            if hasattr(self, 'mechanic_combo'):
                self.mechanic_combo.blockSignals(True)
                self.mechanic_combo.clear()
                self.mechanic_combo.addItem("Выберите механика...", None)

                for mechanic in mechanics:
                    self.mechanic_combo.addItem(mechanic['fio'], mechanic['id'])

                # Восстанавливаем выбранное значение
                if current_mechanic_id:
                    index = self.mechanic_combo.findData(current_mechanic_id)
                    if index >= 0:
                        self.mechanic_combo.setCurrentIndex(index)
                self.mechanic_combo.blockSignals(False)

            # Обновление таблиц
            self.refresh_waybill_list()
            self.refresh_drivers_table()
            self.refresh_mechanics_table()
            self.refresh_vehicles_table()
            self.refresh_companies_table()

            print("DEBUG: Загрузка данных завершена")

        except Exception as e:
            print(f"ERROR: Ошибка в load_data: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки данных: {str(e)}")

    def generate_number_manually(self):
        """Ручная генерация номера путевого листа"""
        try:
            import re

            # Получаем последний номер из БД
            last_waybill = self.db.get_last_waybill_number()

            if last_waybill:
                # Извлекаем числовую часть
                match = re.search(r'ПЛ-(\d+)', last_waybill)
                if match:
                    next_num = int(match.group(1)) + 1
                else:
                    # Пробуем другие форматы
                    match = re.search(r'(\d+)', last_waybill)
                    if match:
                        next_num = int(match.group(1)) + 1
                    else:
                        next_num = 1
            else:
                next_num = 1

            return f"ПЛ-{next_num:06d}"

        except Exception as e:
            print(f"ERROR в generate_number_manually: {e}")
            # Возвращаем номер по умолчанию
            return "ПЛ-000001"

    def on_company_changed(self, index):
        """Обработка изменения предприятия"""
        company_id = self.company_combo_waybill.currentData()
        print(f"DEBUG: Выбрано предприятие с ID: {company_id}")

    def on_driver_changed(self, index):
        """Обработка изменения водителя"""
        print(f"DEBUG on_driver_changed START: index={index}")
        try:
            driver_id = self.driver_combo.currentData()
            print(f"DEBUG: Выбран водитель с ID: {driver_id}")

            if driver_id:
                # Получаем данные водителя только для информации
                driver = self.db.get_driver(driver_id)
                print(f"DEBUG: Данные водителя: {driver}")

                # НЕ устанавливаем дату медосмотра автоматически!
                # Врач сам установит нужную дату в полях medical_pre_date и medical_post_date

                # Только логируем информацию о последнем медосмотре
                if driver and driver.get('medical_date'):
                    print(f"INFO: Последний медосмотр водителя был: {driver['medical_date']}")
                else:
                    print(f"INFO: Дата медосмотра водителя не указана в БД")

            print(f"DEBUG on_driver_changed END")
        except Exception as e:
            print(f"ERROR in on_driver_changed: {e}")
            import traceback
            traceback.print_exc()

    def on_vehicle_changed(self, index):
        """Обработка изменения автомобиля"""
        vehicle_id = self.vehicle_combo.currentData()
        print(f"DEBUG: Выбран автомобиль с ID: {vehicle_id}")

        if vehicle_id:
            vehicle = self.db.get_vehicle(vehicle_id)
            if vehicle:
                # Устанавливаем текущий пробег
                current_mileage = vehicle.get('mileage', 0)
                self.vehicle_odo_start.setValue(current_mileage)
                self.vehicle_odo_end.setValue(current_mileage)
                print(f"DEBUG: Установлен пробег: {current_mileage} км")

    def save_waybill(self):
        """Сохранение путевого листа с проверкой обязательных полей"""
        try:
            # ПРОВЕРКА ОБЯЗАТЕЛЬНЫХ ПОЛЕЙ
            errors = []

            # Проверка предприятия
            company_id = self.company_combo_waybill.currentData()
            if not company_id:
                errors.append("Не выбрано предприятие")
                self.company_combo_waybill.setStyleSheet("border: 1px solid red;")
            else:
                self.company_combo_waybill.setStyleSheet("")

            # Проверка водителя
            driver_id = self.driver_combo.currentData()
            if not driver_id:
                errors.append("Не выбран водитель")
                self.driver_combo.setStyleSheet("border: 1px solid red;")
            else:
                self.driver_combo.setStyleSheet("")

            # Проверка автомобиля
            vehicle_id = self.vehicle_combo.currentData()
            if not vehicle_id:
                errors.append("Не выбран автомобиль")
                self.vehicle_combo.setStyleSheet("border: 1px solid red;")
            else:
                self.vehicle_combo.setStyleSheet("")

            # Проверка даты
            if not self.waybill_date.date().isValid():
                errors.append("Неверная дата")

            # Проверка одометра
            odo_start = self.vehicle_odo_start.value()
            odo_end = self.vehicle_odo_end.value()
            if odo_end < odo_start:
                errors.append("Показания одометра при возврате не могут быть меньше показаний при выезде")
                self.vehicle_odo_end.setStyleSheet("border: 1px solid red;")
            else:
                self.vehicle_odo_end.setStyleSheet("")

            if errors:
                error_msg = "Ошибки при сохранении:\n\n• " + "\n• ".join(errors)
                QMessageBox.warning(self, "Ошибка заполнения", error_msg)
                return

            # Получаем данные предприятия, водителя, автомобиля и механика из БД
            company = self.db.get_company(company_id) if company_id else None
            driver = self.db.get_driver(driver_id) if driver_id else None
            vehicle = self.db.get_vehicle(vehicle_id) if vehicle_id else None
            mechanic_id = self.mechanic_combo.currentData() if hasattr(self, 'mechanic_combo') else None
            mechanic = self.db.get_mechanic(mechanic_id) if mechanic_id else None

            # Если все проверки пройдены, собираем данные
            waybill_data = {
                'number': self.waybill_id.text(),
                'date': self.waybill_date.date().toString("yyyy-MM-dd"),
                'status': self.waybill_status.currentText(),

                # Данные организации
                'org_name': company['name'] if company else '',
                'org_address': company.get('address', '') if company else '',
                'org_phone': company.get('phone', '') if company else '',
                'org_inn': company.get('inn', '') if company else '',
                'org_ogrn': company.get('ogrn', '') if company else '',
                'company_id': company_id,

                # Данные водителя
                'driver_id': driver_id,
                'driver_fio': driver['fio'] if driver else '',
                'driver_snils': driver.get('snils', '') if driver else '',
                'driver_license': driver.get('license', '') if driver else '',
                'driver_license_class': driver.get('license_class', '') if driver else '',
                'driver_medical_date': driver.get('medical_date', '') if driver else '',
                'driver_license_issue_date': driver.get('license_issue_date', '') if driver else '',

                # Данные автомобиля
                'vehicle_id': vehicle_id,
                'vehicle_brand': vehicle['brand'] if vehicle else '',
                'vehicle_model': vehicle.get('model', '') if vehicle else '',
                'vehicle_plate': vehicle['plate'] if vehicle else '',

                # Данные механика
                'mechanic_id': mechanic_id,
                'mechanic_fio': mechanic['fio'] if mechanic else '',
                'mechanic_position': mechanic.get('position', '') if mechanic else '',
                'mechanic_license': mechanic.get('license_number', '') if mechanic else '',
                'mechanic_license_date': mechanic.get('license_date', '') if mechanic else '',

                # Одометр
                'odo_start': odo_start,
                'odo_end': odo_end,

                # Топливо
                'fuel_type': self.fuel_type.currentText(),
                'fuel_start': self.fuel_start.value(),
                'fuel_issued': self.fuel_issued.value(),
                'fuel_end': self.fuel_end.value(),

                # Медосмотры
                'medical_pre_date': self.medical_pre_date.dateTime().toString("yyyy-MM-dd HH:mm"),
                'medical_pre_doctor': self.medical_pre_doctor.text(),
                'medical_post_date': self.medical_post_date.dateTime().toString("yyyy-MM-dd HH:mm"),
                'medical_post_doctor': self.medical_post_doctor.text(),

                # Маршрут и примечания
                'route': self.route_text.toPlainText(),
                'notes': self.notes_text.toPlainText()
            }

            # Рассчитываем пробег
            mileage = odo_end - odo_start
            waybill_data['mileage'] = mileage if mileage > 0 else 0

            print(f"DEBUG: Данные для сохранения:")
            print(f"  Номер: {waybill_data['number']}")
            print(f"  Предприятие ID: {company_id}")
            print(f"  Водитель ID: {driver_id}")
            print(f"  Автомобиль ID: {vehicle_id}")
            print(f"  Механик ID: {mechanic_id}")
            print(f"  Пробег: {mileage} км")

            if self.current_waybill_id:
                # Обновление существующего
                self.db.update_waybill(self.current_waybill_id, waybill_data)
                message = f"Путевой лист {self.waybill_id.text()} обновлен"

                # Обновляем пробег автомобиля в БД
                if vehicle_id:
                    self.db.update_vehicle_mileage(vehicle_id, odo_end)
            else:
                # Создание нового
                self.current_waybill_id = self.db.create_waybill(waybill_data)
                message = f"Путевой лист {self.waybill_id.text()} создан"

                # Обновляем пробег автомобиля в БД
                if vehicle_id:
                    self.db.update_vehicle_mileage(vehicle_id, odo_end)

            self.status_bar.showMessage(message)
            QMessageBox.information(self, "Сохранение", message)

            # Обновление списка
            self.refresh_waybill_list()
            print("✅ Путевой лист сохранен успешно")

        except Exception as e:
            print(f"ERROR: Ошибка при сохранении: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении: {str(e)}")

    def create_new_waybill(self):
        """Создание нового путевого листа"""
        try:
            print("DEBUG: Создание нового путевого листа...")

            # Сброс текущего ID
            self.current_waybill_id = None

            # Пробуем получить номер через новую систему нумерации
            new_number = None

            # Проверяем, есть ли метод get_next_number в БД
            if hasattr(self.db, 'get_next_number'):
                try:
                    new_number = self.db.get_next_number()
                    print(f"DEBUG: Получен номер из БД: {new_number}")
                except Exception as e:
                    print(f"DEBUG: Ошибка получения номера из БД: {e}")
                    new_number = None

            # Если номер не сгенерировался через БД, создаем вручную
            if not new_number:
                print("DEBUG: Использую ручную генерацию номера...")
                import re
                from datetime import datetime

                # Получаем последний номер из БД
                last_waybill = self.db.get_last_waybill_number()
                if last_waybill:
                    # Извлекаем числовую часть
                    match = re.search(r'ПЛ-(\d+)', last_waybill)
                    if match:
                        next_num = int(match.group(1)) + 1
                    else:
                        next_num = 1
                else:
                    next_num = 1

                new_number = f"ПЛ-{next_num:06d}"
                print(f"DEBUG: Сгенерирован номер вручную: {new_number}")

            print(f"DEBUG: Итоговый номер: {new_number}")

            # Очистка формы
            self.clear_waybill_form()

            # Установка номера
            self.waybill_id.setText(new_number)

            # Установка даты по умолчанию
            self.waybill_date.setDate(QDate.currentDate())

            # Установка статуса по умолчанию
            self.waybill_status.setCurrentText("Черновик")

            # Сброс выбора в комбобоксах
            self.company_combo_waybill.setCurrentIndex(0)
            self.driver_combo.setCurrentIndex(0)
            self.vehicle_combo.setCurrentIndex(0)
            if hasattr(self, 'mechanic_combo'):
                self.mechanic_combo.setCurrentIndex(0)

            # Сброс одометра и топлива
            self.vehicle_odo_start.setValue(0)
            self.vehicle_odo_end.setValue(0)
            self.fuel_start.setValue(0.0)
            self.fuel_issued.setValue(0.0)
            self.fuel_end.setValue(0.0)

            # Установка текущей даты для медосмотров
            current_datetime = datetime.now()
            self.medical_pre_date.setDateTime(current_datetime)
            self.medical_post_date.setDateTime(current_datetime)

            # Очистка текстовых полей
            self.route_text.clear()
            self.notes_text.clear()

            # Сброс подсветки ошибок
            self.company_combo_waybill.setStyleSheet("")
            self.driver_combo.setStyleSheet("")
            self.vehicle_combo.setStyleSheet("")
            self.vehicle_odo_end.setStyleSheet("")

            self.status_bar.showMessage(f"Создан новый путевой лист №{new_number}")

            # Переключение на вкладку редактирования
            self.tab_widget.setCurrentIndex(0)

            print(f"DEBUG: Новый путевой лист создан: {new_number}")

        except Exception as e:
            print(f"ERROR: Ошибка при создании путевого листа: {e}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при создании путевого листа: {str(e)}")

    def clear_waybill_form(self):
        """Очистка формы путевого листа"""
        try:
            # Очищаем основные поля
            self.waybill_id.clear()
            self.waybill_date.setDate(QDate.currentDate())
            self.waybill_status.setCurrentIndex(0)

            # Сбрасываем комбобоксы
            self.company_combo_waybill.setCurrentIndex(0)
            self.driver_combo.setCurrentIndex(0)
            self.vehicle_combo.setCurrentIndex(0)
            if hasattr(self, 'mechanic_combo'):
                self.mechanic_combo.setCurrentIndex(0)

            # Сбрасываем числовые поля
            self.vehicle_odo_start.setValue(0)
            self.vehicle_odo_end.setValue(0)
            self.fuel_start.setValue(0.0)
            self.fuel_issued.setValue(0.0)
            self.fuel_end.setValue(0.0)

            # Сбрасываем даты
            current_datetime = datetime.now()
            self.medical_pre_date.setDateTime(current_datetime)
            self.medical_post_date.setDateTime(current_datetime)

            # Очищаем текстовые поля
            self.medical_pre_doctor.clear()
            self.medical_post_doctor.clear()
            self.route_text.clear()
            self.notes_text.clear()

            print("DEBUG: Форма путевого листа очищена")
        except Exception as e:
            print(f"ERROR in clear_waybill_form: {e}")

    def filter_waybills(self):
        """Фильтрация путевых листов"""
        print("DEBUG: filter_waybills вызван")

        try:
            # Получаем фильтры
            driver_filter = self.filter_driver.currentText()
            status_filter = self.filter_status.currentText()
            mechanic_filter = self.filter_mechanic.currentText()  # НОВЫЙ ФИЛЬТР
            search_text = self.search_edit.text().lower()

            # Получаем все путевые листы
            all_waybills = self.db.get_waybills()

            # Применяем фильтры
            filtered_waybills = []
            for waybill in all_waybills:
                # Фильтр по водителю
                if driver_filter != "Все водители":
                    driver_fio = waybill.get('driver_fio', '')
                    if driver_fio != driver_filter:
                        continue

                # Фильтр по статусу
                if status_filter != "Все":
                    status = waybill.get('status', '')
                    if status != status_filter:
                        continue

                # Фильтр по механику (ДОБАВЛЕНО)
                if mechanic_filter != "Все механики":
                    mechanic_fio = waybill.get('mechanic_fio', '')
                    if mechanic_fio != mechanic_filter:
                        continue

                # Быстрый поиск
                if search_text:
                    search_fields = [
                        str(waybill.get('number', '')),
                        str(waybill.get('date', '')),
                        str(waybill.get('driver_fio', '')),
                        str(waybill.get('vehicle_info', '')),
                        str(waybill.get('status', '')),
                        str(waybill.get('mechanic_fio', '')),  # Добавили механика в поиск
                        str(waybill.get('org_name', ''))
                    ]
                    if not any(search_text in str(field).lower() for field in search_fields):
                        continue

                filtered_waybills.append(waybill)

            # Обновляем таблицу с отфильтрованными данными
            self.update_waybill_table(filtered_waybills)

        except Exception as e:
            print(f"ERROR: Ошибка в filter_waybills: {e}")

    def update_waybill_table(self, waybills):
        """Обновление таблицы путевых листов"""
        self.waybill_table.setRowCount(len(waybills))

        for row, waybill in enumerate(waybills):
            waybill_id = waybill['id']
            status = waybill.get('status', 'Черновик')

            # Колонка 0: Чекбокс
            checkbox_item = QTableWidgetItem()
            checkbox_item.setFlags(checkbox_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            checkbox_item.setCheckState(
                Qt.CheckState.Checked if status == 'Проведен'
                else Qt.CheckState.Unchecked
            )
            checkbox_item.setData(Qt.ItemDataRole.UserRole, waybill_id)
            checkbox_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.waybill_table.setItem(row, 0, checkbox_item)

            # Колонка 1: ID
            id_item = QTableWidgetItem(str(waybill_id))
            id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.waybill_table.setItem(row, 1, id_item)

            # Колонка 2: Дата
            date_str = waybill.get('date', '')
            if date_str:
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                    date_str = date_obj.strftime('%d.%m.%Y')
                except:
                    pass
            self.waybill_table.setItem(row, 2, QTableWidgetItem(date_str))

            # Колонка 3: Водитель (только фамилия)
            driver_fio = waybill.get('driver_fio', '')
            if driver_fio:
                parts = driver_fio.split()
                driver_fio = parts[0] if parts else driver_fio
            self.waybill_table.setItem(row, 3, QTableWidgetItem(driver_fio))

            # Колонка 4: Автомобиль
            vehicle_info = waybill.get('vehicle_info', '')
            if not vehicle_info:
                vehicle_brand = waybill.get('vehicle_brand', '')
                vehicle_plate = waybill.get('vehicle_plate', '')
                vehicle_info = f"{vehicle_brand} {vehicle_plate}".strip()
            self.waybill_table.setItem(row, 4, QTableWidgetItem(vehicle_info))

            # Колонка 5: Статус с иконкой
            status_icons = {
                'Проведен': '📋 ',
                'Активен': '🚗 ',
                'Черновик': '📝 ',
                'Завершен': '✅ ',
                'Архив': '🗄️ '
            }
            icon = status_icons.get(status, '❓ ')
            status_item = QTableWidgetItem(f"{icon}{status}")
            self.waybill_table.setItem(row, 5, status_item)

            # Колонка 6: Механик (НОВАЯ КОЛОНКА)
            mechanic_fio = waybill.get('mechanic_fio', '')
            if mechanic_fio:
                parts = mechanic_fio.split()
                mechanic_fio = parts[0] if parts else mechanic_fio
            self.waybill_table.setItem(row, 6, QTableWidgetItem(mechanic_fio))

            # Колонка 7: Организация
            org_name = waybill.get('org_name', '')
            if not org_name and waybill.get('company_id'):
                company = self.db.get_company(waybill['company_id'])
                if company:
                    org_name = company.get('name', '')

            if len(org_name) > 30:
                org_name = org_name[:27] + "..."
            self.waybill_table.setItem(row, 7, QTableWidgetItem(org_name))

            # Цвет строки
            row_color = self.get_status_color(status)
            for col in range(8):
                item = self.waybill_table.item(row, col)
                if item:
                    item.setBackground(row_color)

    def get_status_color(self, status):
        """Возвращает цвет фона в зависимости от статуса"""
        colors = {
            'Проведен': QColor(220, 255, 220),  # Светло-зеленый
            'Активен': QColor(255, 255, 200),  # Светло-желтый
            'Черновик': QColor(245, 245, 245),  # Светло-серый
            'Завершен': QColor(220, 220, 255),  # Светло-синий
            'Архив': QColor(240, 240, 240)  # Серый
        }
        return colors.get(status, QColor(255, 255, 255))

    def on_waybill_cell_clicked(self, row, column):
        """Обработка клика по ячейке таблицы (чекбокс проведения)"""
        # Только для колонки чекбокса (0)
        if column == 0:
            item = self.waybill_table.item(row, column)
            if not item:
                return

            waybill_id = item.data(Qt.ItemDataRole.UserRole)
            if not waybill_id:
                # Получаем ID из колонки 1
                id_item = self.waybill_table.item(row, 1)
                if id_item:
                    waybill_id = int(id_item.text())

            # Получаем текущий статус из колонки 5
            status_item = self.waybill_table.item(row, 5)
            current_status = ""
            if status_item:
                status_text = status_item.text()
                # Извлекаем статус из текста (может быть с иконкой)
                for status in ["Черновик", "Активен", "Завершен", "Архив", "Проведен"]:
                    if status in status_text:
                        current_status = status
                        break

            new_check_state = item.checkState()

            # Если ставим галочку и статус не "Проведен"
            if new_check_state == Qt.CheckState.Checked and current_status != 'Проведен':
                reply = QMessageBox.question(
                    self, "Провести путевой лист",
                    "Вы уверены, что хотите провести этот путевой лист?\n"
                    "После проведения редактирование будет ограничено.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.change_waybill_status(waybill_id, current_status, 'Проведен', row)
                else:
                    # Отменяем изменение
                    item.setCheckState(Qt.CheckState.Unchecked)

            # Если снимаем галочку и статус "Проведен"
            elif new_check_state == Qt.CheckState.Unchecked and current_status == 'Проведен':
                reply = QMessageBox.question(
                    self, "Отменить проводку",
                    "Вы уверены, что хотите отменить проводку?\n"
                    "Путевой лист снова станет доступен для редактирования.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.change_waybill_status(waybill_id, current_status, 'Активен', row)
                else:
                    # Возвращаем галочку
                    item.setCheckState(Qt.CheckState.Checked)

    def change_waybill_status(self, waybill_id, old_status, new_status, row):
        """Изменение статуса путевого листа"""
        try:
            waybill = self.db.get_waybill(waybill_id)
            if not waybill:
                QMessageBox.warning(self, "Ошибка", "Путевой лист не найден")
                return

            # Обновляем только статус
            success = self.db.update_waybill(waybill_id, {'status': new_status})

            if success:
                # Обновляем отображение в таблице
                status_item = self.waybill_table.item(row, 5)
                if status_item:
                    # Обновляем текст статуса
                    status_icons = {
                        'Проведен': '📋 ',
                        'Активен': '🚗 ',
                        'Черновик': '📝 ',
                        'Завершен': '✅ ',
                        'Архив': '🗄️ '
                    }
                    icon = status_icons.get(new_status, '')
                    status_item.setText(f"{icon}{new_status}")

                # Обновляем чекбокс
                checkbox_item = self.waybill_table.item(row, 0)
                if checkbox_item:
                    checkbox_item.setCheckState(
                        Qt.CheckState.Checked if new_status == 'Проведен'
                        else Qt.CheckState.Unchecked
                    )

                # Обновляем цвет строки
                new_color = self.get_status_color(new_status)
                for col in range(self.waybill_table.columnCount()):
                    item = self.waybill_table.item(row, col)
                    if item:
                        item.setBackground(new_color)

                self.status_bar.showMessage(f"Статус изменен: {old_status} → {new_status}")
                print(f"Статус путевого листа {waybill_id} изменен: {old_status} → {new_status}")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось изменить статус")

        except Exception as e:
            print(f"ERROR: Ошибка при изменении статуса: {e}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при изменении статуса: {str(e)}")

    def refresh_waybill_list(self):
        """Обновление списка путевых листов"""
        try:
            waybills = self.db.get_waybills()

            # Обновляем статистику
            self._update_stats(waybills)

            # Обновляем таблицу
            self.update_waybill_table(waybills)

        except Exception as e:
            print(f"ERROR: Ошибка в refresh_waybill_list: {e}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось обновить список: {str(e)}")

    def _update_stats(self, waybills):
        """Обновление статистики (приватный метод)"""
        total = len(waybills)
        draft = len([w for w in waybills if w.get('status') == 'Черновик'])
        active = len([w for w in waybills if w.get('status') == 'Активен'])
        completed = len([w for w in waybills if w.get('status') in ['Завершен', 'Проведен']])

        self.stats_total.setText(f"📊 Всего: {total}")
        self.stats_draft.setText(f"📋 Черновики: {draft}")
        self.stats_active.setText(f"🚗 Активные: {active}")
        self.stats_completed.setText(f"✅ Завершены: {completed}")

    def refresh_companies_table(self):
        """Обновление таблицы предприятий"""
        try:
            companies = self.db.get_companies()
            self.companies_table.setRowCount(len(companies))

            for row, company in enumerate(companies):
                self.companies_table.setItem(row, 0, QTableWidgetItem(str(company['id'])))
                self.companies_table.setItem(row, 1, QTableWidgetItem(company['name']))
                self.companies_table.setItem(row, 2, QTableWidgetItem(company.get('inn', '')))
                self.companies_table.setItem(row, 3, QTableWidgetItem(company.get('kpp', '')))
                self.companies_table.setItem(row, 4, QTableWidgetItem(company.get('ogrn', '')))
                self.companies_table.setItem(row, 5, QTableWidgetItem(company.get('address', '')))
                self.companies_table.setItem(row, 6, QTableWidgetItem(company.get('phone', '')))

        except Exception as e:
            print(f"ERROR: Ошибка в refresh_companies_table: {e}")

    def add_company(self):
        """Добавление нового предприятия"""
        dialog = CompanyDialog(self)
        if dialog.exec():
            company_data = dialog.get_data()
            if not company_data.get('name'):
                QMessageBox.warning(self, "Ошибка", "Наименование предприятия обязательно!")
                return
            self.db.create_company(company_data)
            self.refresh_companies_table()
            self.load_data()

    def edit_company(self):
        """Редактирование предприятия"""
        selected_items = self.companies_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите предприятие")
            return

        row = selected_items[0].row()
        company_id = int(self.companies_table.item(row, 0).text())

        company = self.db.get_company(company_id)
        if company:
            dialog = CompanyDialog(self, company)
            if dialog.exec():
                company_data = dialog.get_data()
                self.db.update_company(company_id, company_data)
                self.refresh_companies_table()
                self.load_data()

    def delete_company(self):
        """Удаление предприятия"""
        selected_items = self.companies_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите предприятие")
            return

        row = selected_items[0].row()
        company_id = int(self.companies_table.item(row, 0).text())
        company_name = self.companies_table.item(row, 1).text()

        # Проверяем, есть ли автомобили у этого предприятия
        vehicles = self.db.get_vehicles_by_company(company_id)
        if vehicles:
            QMessageBox.warning(
                self, "Ошибка",
                f"Невозможно удалить предприятие '{company_name}'\n"
                f"К нему привязаны {len(vehicles)} автомобилей.\n"
                "Сначала удалите или перепривяжите автомобили."
            )
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить предприятие {company_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.db.delete_company(company_id)
            self.refresh_companies_table()
            self.load_data()

    def import_companies(self):
        """Импорт предприятий из Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            try:
                companies = self.exporter.import_companies_from_excel(file_path)
                for company in companies:
                    self.db.create_company(company)
                self.refresh_companies_table()
                self.load_data()
                QMessageBox.information(self, "Импорт", f"Импортировано {len(companies)} предприятий")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при импорте: {str(e)}")

    def export_companies(self):
        """Экспорт предприятий в Excel"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                companies = self.db.get_companies()
                self.exporter.export_companies_to_excel(file_path, companies)
                QMessageBox.information(self, "Экспорт", f"Экспортировано {len(companies)} предприятий")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    def refresh_drivers_table(self):
        """Обновление таблицы водителей"""
        try:
            drivers = self.db.get_drivers()
            self.drivers_table.setRowCount(len(drivers))

            for row, driver in enumerate(drivers):
                self.drivers_table.setItem(row, 0, QTableWidgetItem(str(driver['id'])))
                self.drivers_table.setItem(row, 1, QTableWidgetItem(driver['fio']))
                self.drivers_table.setItem(row, 2, QTableWidgetItem(driver.get('snils', '')))
                self.drivers_table.setItem(row, 3, QTableWidgetItem(driver.get('license', '')))
                self.drivers_table.setItem(row, 4, QTableWidgetItem(driver.get('license_class', '')))
                self.drivers_table.setItem(row, 5, QTableWidgetItem(driver.get('license_issue_date', '')))
                self.drivers_table.setItem(row, 6, QTableWidgetItem(driver.get('phone', '')))
                self.drivers_table.setItem(row, 7, QTableWidgetItem(driver.get('medical_date', '')))

                # Автомобили водителя
                try:
                    vehicles = self.db.get_vehicles_by_driver(driver['id'])
                    if vehicles:
                        first_vehicle = vehicles[0]
                        vehicle_text = f"{first_vehicle.get('brand', '')} ({first_vehicle.get('plate', '')})"
                        if len(vehicles) > 1:
                            vehicle_text += f" +{len(vehicles) - 1}"
                    else:
                        vehicle_text = "Не назначен"
                except:
                    vehicle_text = "Не назначен"

                self.drivers_table.setItem(row, 8, QTableWidgetItem(vehicle_text))

        except Exception as e:
            print(f"ERROR: Ошибка в refresh_drivers_table: {e}")
            import traceback
            traceback.print_exc()

    def refresh_mechanics_table(self):
        """Обновление таблицы механиков"""
        try:
            mechanics = self.db.get_mechanics()
            self.mechanics_table.setRowCount(len(mechanics))

            for row, mechanic in enumerate(mechanics):
                self.mechanics_table.setItem(row, 0, QTableWidgetItem(str(mechanic['id'])))
                self.mechanics_table.setItem(row, 1, QTableWidgetItem(mechanic['fio']))
                self.mechanics_table.setItem(row, 2, QTableWidgetItem(mechanic.get('position', '')))
                self.mechanics_table.setItem(row, 3, QTableWidgetItem(mechanic.get('phone', '')))
                self.mechanics_table.setItem(row, 4, QTableWidgetItem(mechanic.get('email', '')))
                self.mechanics_table.setItem(row, 5, QTableWidgetItem(mechanic.get('license_number', '')))

                # Дата выдачи удостоверения
                license_date = mechanic.get('license_date', '')
                self.mechanics_table.setItem(row, 6, QTableWidgetItem(license_date))

                # Предприятие
                company_name = mechanic.get('company_name', 'Не назначено')
                self.mechanics_table.setItem(row, 7, QTableWidgetItem(company_name))

        except Exception as e:
            print(f"ERROR: Ошибка в refresh_mechanics_table: {e}")
            import traceback
            traceback.print_exc()

    def refresh_vehicles_table(self):
        """Обновление таблицы автомобилей"""
        try:
            vehicles = self.db.get_vehicles()
            self.vehicles_table.setRowCount(len(vehicles))

            for row, vehicle in enumerate(vehicles):
                self.vehicles_table.setItem(row, 0, QTableWidgetItem(str(vehicle['id'])))
                self.vehicles_table.setItem(row, 1, QTableWidgetItem(vehicle['brand']))
                self.vehicles_table.setItem(row, 2, QTableWidgetItem(vehicle.get('model', '')))
                self.vehicles_table.setItem(row, 3, QTableWidgetItem(vehicle['plate']))
                self.vehicles_table.setItem(row, 4, QTableWidgetItem(vehicle.get('vin', '')))
                self.vehicles_table.setItem(row, 5, QTableWidgetItem(str(vehicle.get('year', ''))))
                self.vehicles_table.setItem(row, 6, QTableWidgetItem(str(vehicle.get('mileage', 0))))

                # Водитель
                driver_id = vehicle.get('driver_id')
                if driver_id:
                    driver = self.db.get_driver(driver_id)
                    driver_text = driver['fio'] if driver else "Не назначен"
                else:
                    driver_text = "Не назначен"
                self.vehicles_table.setItem(row, 7, QTableWidgetItem(driver_text))

                # Предприятие
                company_id = vehicle.get('company_id')
                if company_id:
                    company = self.db.get_company(company_id)
                    company_text = company['name'] if company else "Не назначено"
                else:
                    company_text = "Не назначено"
                self.vehicles_table.setItem(row, 8, QTableWidgetItem(company_text))

        except Exception as e:
            print(f"ERROR: Ошибка в refresh_vehicles_table: {e}")

    def load_waybill(self, item):
        """Загрузка путевого листа по двойному клику"""
        row = item.row()
        waybill_id = int(self.waybill_table.item(row, 1).text())
        self.open_waybill_by_id(waybill_id)

    def open_selected_waybill(self):
        """Открытие выбранного путевого листа"""
        selected_items = self.waybill_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите путевой лист")
            return

        row = selected_items[0].row()
        waybill_id = int(self.waybill_table.item(row, 1).text())
        self.open_waybill_by_id(waybill_id)

    def open_waybill_by_id(self, waybill_id):
        """Открытие путевого листа по ID"""
        waybill = self.db.get_waybill(waybill_id)
        if not waybill:
            QMessageBox.warning(self, "Ошибка", "Путевой лист не найден")
            return

        self.current_waybill_id = waybill_id

        # Заполнение формы
        self.waybill_id.setText(waybill['number'])
        self.waybill_date.setDate(QDate.fromString(waybill['date'], "yyyy-MM-dd"))
        self.waybill_status.setCurrentText(waybill['status'])

        # Установка предприятия
        company_id = waybill.get('company_id')
        if company_id:
            index = self.company_combo_waybill.findData(company_id)
            if index >= 0:
                self.company_combo_waybill.setCurrentIndex(index)

        # Установка водителя
        driver_id = waybill.get('driver_id')
        if driver_id:
            index = self.driver_combo.findData(driver_id)
            if index >= 0:
                self.driver_combo.setCurrentIndex(index)

        # Установка автомобиля
        vehicle_id = waybill.get('vehicle_id')
        if vehicle_id:
            index = self.vehicle_combo.findData(vehicle_id)
            if index >= 0:
                self.vehicle_combo.setCurrentIndex(index)

        # Установка механика
        mechanic_id = waybill.get('mechanic_id')
        if mechanic_id and hasattr(self, 'mechanic_combo'):
            index = self.mechanic_combo.findData(mechanic_id)
            if index >= 0:
                self.mechanic_combo.setCurrentIndex(index)

        # Одометр
        self.vehicle_odo_start.setValue(waybill.get('odo_start', 0))
        self.vehicle_odo_end.setValue(waybill.get('odo_end', 0))

        # Топливо
        fuel_type = waybill.get('fuel_type', 'Бензин')
        index = self.fuel_type.findText(fuel_type)
        if index >= 0:
            self.fuel_type.setCurrentIndex(index)

        self.fuel_start.setValue(waybill.get('fuel_start', 0))
        self.fuel_issued.setValue(waybill.get('fuel_issued', 0))
        self.fuel_end.setValue(waybill.get('fuel_end', 0))

        # Медосмотры
        medical_pre_date = waybill.get('medical_pre_date')
        if medical_pre_date:
            try:
                dt = QDateTime.fromString(medical_pre_date, "yyyy-MM-dd HH:mm")
                if dt.isValid():
                    self.medical_pre_date.setDateTime(dt)
            except:
                pass

        self.medical_pre_doctor.setText(waybill.get('medical_pre_doctor', ''))

        medical_post_date = waybill.get('medical_post_date')
        if medical_post_date:
            try:
                dt = QDateTime.fromString(medical_post_date, "yyyy-MM-dd HH:mm")
                if dt.isValid():
                    self.medical_post_date.setDateTime(dt)
            except:
                pass

        self.medical_post_doctor.setText(waybill.get('medical_post_doctor', ''))

        # Маршрут и примечания
        self.route_text.setText(waybill.get('route', ''))
        self.notes_text.setText(waybill.get('notes', ''))

        self.status_bar.showMessage(f"Загружен путевой лист {waybill['number']}")

    def on_waybill_selected(self):
        """Обработка выбора путевого листа"""
        selected = self.waybill_table.selectedItems()
        if selected:
            row = selected[0].row()
            status_item = self.waybill_table.item(row, 5)
            if status_item:
                status = status_item.text()
                self.status_bar.showMessage(f"Выбран путевой лист. Статус: {status}")

    def delete_waybill(self):
        """Удаление путевого листа"""
        selected_items = self.waybill_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите путевой лист для удаления")
            return

        row = selected_items[0].row()
        waybill_id_item = self.waybill_table.item(row, 1)
        if not waybill_id_item:
            QMessageBox.warning(self, "Ошибка", "Не удалось получить ID путевого листа")
            return

        waybill_id = int(waybill_id_item.text())
        waybill_number_item = self.waybill_table.item(row, 2)
        waybill_number = waybill_number_item.text() if waybill_number_item else f"ID:{waybill_id}"

        reply = QMessageBox.question(
            self, "Подтверждение удаления",
            f"Вы уверены, что хотите удалить путевой лист?\n\n"
            f"Номер: {waybill_number}\n"
            f"ID: {waybill_id}\n\n"
            "Это действие нельзя отменить!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                success = self.db.delete_waybill(waybill_id)
                if success:
                    self.status_bar.showMessage(f"Путевой лист {waybill_number} удален")
                    QMessageBox.information(self, "Успех", f"Путевой лист {waybill_number} удален")
                    self.refresh_waybill_list()
                    if self.current_waybill_id == waybill_id:
                        self.current_waybill_id = None
                        self.clear_waybill_form()
                        self.status_bar.showMessage("Открытый путевой лист удален")
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось удалить путевой лист")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении: {str(e)}")

    def conduct_waybill(self):
        """Провести путевой лист (закрыть)"""
        selected_items = self.waybill_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите путевой лист для проведения")
            return

        row = selected_items[0].row()
        waybill_id = int(self.waybill_table.item(row, 1).text())

        waybill = self.db.get_waybill(waybill_id)
        if not waybill:
            QMessageBox.warning(self, "Ошибка", "Путевой лист не найден")
            return

        waybill_number = waybill.get('number', '')

        reply = QMessageBox.question(
            self, "Провести путевой лист",
            f"Провести путевой лист №{waybill_number}?\n\n"
            f"Водитель: {waybill.get('driver_fio', 'Не указан')}\n"
            f"Механик: {waybill.get('mechanic_fio', 'Не указан')}\n"
            f"Дата: {waybill.get('date', 'Не указана')}\n\n"
            "После проведения путевой лист будет считаться закрытым.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                waybill_data = {
                    'status': 'Проведен',
                    'number': waybill.get('number', ''),
                    'date': waybill.get('date', ''),
                    'driver_id': waybill.get('driver_id'),
                    'vehicle_id': waybill.get('vehicle_id'),
                    'company_id': waybill.get('company_id'),
                    'mechanic_id': waybill.get('mechanic_id')
                }

                self.db.update_waybill(waybill_id, waybill_data)
                self.refresh_waybill_list()

                QMessageBox.information(self, "Успех",
                                        f"Путевой лист №{waybill_number} проведен успешно!")
                self.status_bar.showMessage(f"Путевой лист №{waybill_number} проведен")

            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при проведении: {str(e)}")

    def preview_current_waybill(self):
        """Предпросмотр текущего путевого листа"""
        if not self.current_waybill_id:
            QMessageBox.warning(self, "Предупреждение", "Нет активного путевого листа")
            return

        try:
            # Получаем данные путевого листа
            waybill_data = self.db.get_waybill(self.current_waybill_id)
            if not waybill_data:
                QMessageBox.warning(self, "Ошибка", "Путевой лист не найден")
                return

            # Обновляем данные из формы
            self.update_waybill_data_from_form(waybill_data)

            # Пробуем использовать доступные модули предпросмотра
            success = False

            if PRINT_1C_AVAILABLE:
                try:
                    result = preview_waybill_1c(waybill_data, self.db, self)
                    if result:
                        success = True
                        self.status_bar.showMessage(f"Предпросмотр создан (1С): {result}")
                except Exception as e:
                    print(f"Ошибка предпросмотра 1С: {e}")

            if not success and PREVIEW_PRINTER_AVAILABLE:
                try:
                    result = print_with_preview(waybill_data, self.db, self)
                    if result:
                        success = True
                        self.status_bar.showMessage(f"Предпросмотр создан")
                except Exception as e:
                    print(f"Ошибка предпросмотра: {e}")

            if success:
                QMessageBox.information(self, "Предпросмотр", "Предпросмотр создан успешно")
            else:
                QMessageBox.warning(self, "Предпросмотр",
                                    "Не удалось создать предпросмотр. Проверьте доступность модулей.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при создании предпросмотра: {str(e)}")
            import traceback
            traceback.print_exc()

    def update_waybill_data_from_form(self, waybill_data):
        """Обновление данных путевого листа из формы"""
        waybill_data.update({
            'odo_start': self.vehicle_odo_start.value(),
            'odo_end': self.vehicle_odo_end.value(),
            'medical_pre_date': self.medical_pre_date.dateTime().toString("yyyy-MM-dd HH:mm"),
            'medical_pre_doctor': self.medical_pre_doctor.text(),
            'medical_post_date': self.medical_post_date.dateTime().toString("yyyy-MM-dd HH:mm"),
            'medical_post_doctor': self.medical_post_doctor.text(),
            'route': self.route_text.toPlainText(),
            'notes': self.notes_text.toPlainText(),
            'fuel_start': self.fuel_start.value(),
            'fuel_issued': self.fuel_issued.value(),
            'fuel_end': self.fuel_end.value(),
            'fuel_type': self.fuel_type.currentText(),
            'mechanic_id': self.mechanic_combo.currentData() if hasattr(self, 'mechanic_combo') else None
        })

    def export_data(self):
        """Экспорт данных"""
        options = ["Путевые листы (Excel)", "Водители (Excel)", "Механики (Excel)", "Автомобили (Excel)",
                   "Предприятия (Excel)",
                   "Все данные (CSV)"]

        dialog = QDialog(self)
        dialog.setWindowTitle("Экспорт данных")
        layout = QVBoxLayout(dialog)

        combo = QComboBox()
        combo.addItems(options)
        layout.addWidget(combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec():
            export_type = combo.currentText()
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Сохранить файл", "", "Excel Files (*.xlsx);;CSV Files (*.csv)"
            )

            if file_path:
                try:
                    if export_type == "Путевые листы (Excel)":
                        self.exporter.export_waybills_to_excel(file_path, self.db.get_waybills())
                    elif export_type == "Водители (Excel)":
                        self.exporter.export_drivers_to_excel(file_path, self.db.get_drivers())
                    elif export_type == "Механики (Excel)":
                        self.exporter.export_mechanics_to_excel(file_path, self.db.get_mechanics())
                    elif export_type == "Автомобили (Excel)":
                        self.exporter.export_vehicles_to_excel(file_path, self.db.get_vehicles())
                    elif export_type == "Предприятия (Excel)":
                        self.exporter.export_companies_to_excel(file_path, self.db.get_companies())
                    elif export_type == "Все данные (CSV)":
                        self.exporter.export_all_to_csv(file_path, self.db)

                    self.status_bar.showMessage(f"Данные экспортированы в {file_path}")
                    QMessageBox.information(self, "Экспорт", "Данные успешно экспортированы")

                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    # =============== добавил для импорта путевых листов===========
    def import_waybills_from_excel_simple(self, file_path):
        """Импорт путевых листов из Excel/CSV файла с полной структурой"""
        try:
            import pandas as pd
            from datetime import datetime

            print(f"Импорт из файла: {file_path}")

            # Читаем файл с заголовками
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, sep='\t', encoding='utf-8')
            else:
                df = pd.read_excel(file_path)

            print(f"Найдено строк: {len(df)}")
            print(f"Колонки: {list(df.columns)}")

            imported_count = 0
            skipped_count = 0
            errors = []

            for idx, row in df.iterrows():
                try:
                    # Преобразуем данные из строки
                    waybill_data = {
                        # Основные данные
                        'number': str(row['Номер']).strip() if 'Номер' in row else f'ИМП-{idx + 1:06d}',
                        'date': self._parse_import_date(row.get('Дата')) if 'Дата' in row else datetime.now().strftime(
                            '%Y-%m-%d'),
                        'status': str(row.get('Статус', 'Черновик')).strip(),

                        # Данные организации
                        'org_name': str(row.get('Организация', '')).strip(),
                        'org_inn': str(row.get('ИНН организации', '')).strip(),

                        # Данные водителя
                        'driver_fio': str(row.get('Водитель', '')).strip(),
                        'driver_snils': str(row.get('СНИЛС', '')).strip(),
                        'driver_license': str(row.get('В/У', '')).strip(),
                        'driver_license_class': str(row.get('Класс', '')).strip(),

                        # Данные автомобиля
                        'vehicle_info': str(row.get('Автомобиль', '')).strip(),
                        'vehicle_brand': str(row.get('Марка', '')).strip(),
                        'vehicle_plate': str(row.get('Гос. номер', '')).strip(),

                        # Одометр и пробег
                        'odo_start': float(row.get('Пробег начальный', 0)),
                        'odo_end': float(row.get('Пробег конечный', 0)),
                        'mileage': float(row.get('Пробег по рейсу', 0)),

                        # Топливо
                        'fuel_type': str(row.get('Тип топлива', 'Бензин')).strip(),
                        'fuel_start': float(row.get('Топливо начальное', 0)),
                        'fuel_issued': float(row.get('Топливо выдано', 0)),
                        'fuel_end': float(row.get('Топливо конечное', 0)),

                        # Медосмотры
                        'medical_pre_date': self._parse_import_datetime(row.get('Медосмотр предрейсовый')),
                        'medical_pre_doctor': str(row.get('Врач предрейсовый', '')).strip(),
                        'medical_post_date': self._parse_import_datetime(row.get('Медосмотр послерейсовый')),
                        'medical_post_doctor': str(row.get('Врач послерейсовый', '')).strip(),

                        # Маршрут и примечания
                        'route': str(row.get('Маршрут', '')).strip(),
                        'notes': str(row.get('Примечания', '')).strip()
                    }

                    # Пропускаем пустые строки
                    if not waybill_data['number'] or waybill_data['number'] == 'nan':
                        print(f"Строка {idx + 1}: пропущена (нет номера)")
                        skipped_count += 1
                        continue

                    if not waybill_data['driver_fio'] or waybill_data['driver_fio'] == 'nan':
                        print(f"Строка {idx + 1}: пропущена (нет водителя)")
                        skipped_count += 1
                        continue

                    # Проверяем обязательные поля
                    required_fields = ['org_name', 'vehicle_plate']
                    missing_fields = []
                    for field in required_fields:
                        if not waybill_data.get(field) or waybill_data[field] == 'nan':
                            missing_fields.append(field)

                    if missing_fields:
                        print(f"Строка {idx + 1}: отсутствуют поля: {missing_fields}")
                        skipped_count += 1
                        continue

                    # 1. Работа с организацией
                    company_name = waybill_data['org_name']
                    company_id = None

                    # Ищем компанию по названию
                    companies = self.db.get_companies()
                    for company in companies:
                        if company['name'] == company_name:
                            company_id = company['id']
                            break

                    # Если компании нет - создаем
                    if not company_id:
                        new_company = {
                            'name': company_name,
                            'inn': waybill_data.get('org_inn', ''),
                            'kpp': '',
                            'ogrn': '',
                            'address': '',
                            'phone': '',
                            'email': '',
                            'director': ''
                        }
                        company_id = self.db.create_company(new_company)
                        print(f"Создана новая организация: {company_name}")

                    waybill_data['company_id'] = company_id

                    # 2. Работа с водителем
                    driver_fio = waybill_data['driver_fio']
                    driver_id = None

                    # Ищем водителя по ФИО
                    drivers = self.db.get_drivers()
                    for driver in drivers:
                        if driver['fio'] == driver_fio:
                            driver_id = driver['id']
                            break

                    # Если водителя нет - создаем
                    if not driver_id:
                        new_driver = {
                            'fio': driver_fio,
                            'snils': waybill_data.get('driver_snils', ''),
                            'license': waybill_data.get('driver_license', ''),
                            'license_class': waybill_data.get('driver_license_class', ''),
                            'license_issue_date': datetime.now().strftime('%Y-%m-%d'),
                            'phone': '',
                            'medical_date': datetime.now().strftime('%Y-%m-%d'),
                            'address': '',
                            'company_id': company_id
                        }
                        driver_id = self.db.create_driver(new_driver)
                        print(f"Создан новый водитель: {driver_fio}")

                    waybill_data['driver_id'] = driver_id

                    # 3. Работа с автомобилем
                    vehicle_plate = waybill_data['vehicle_plate']
                    vehicle_id = None

                    # Ищем автомобиль по госномеру
                    vehicles = self.db.get_vehicles()
                    for vehicle in vehicles:
                        if vehicle['plate'] == vehicle_plate:
                            vehicle_id = vehicle['id']
                            break

                    # Если автомобиля нет - создаем
                    if not vehicle_id:
                        new_vehicle = {
                            'brand': waybill_data.get('vehicle_brand', ''),
                            'model': '',
                            'plate': vehicle_plate,
                            'vin': '',
                            'year': datetime.now().year,
                            'mileage': waybill_data.get('odo_end', 0),
                            'color': '',
                            'driver_id': driver_id,
                            'company_id': company_id
                        }
                        vehicle_id = self.db.create_vehicle(new_vehicle)
                        print(f"Создан новый автомобиль: {vehicle_plate}")

                    waybill_data['vehicle_id'] = vehicle_id

                    # 4. Проверяем, не существует ли уже такой путевой лист
                    existing_waybills = self.db.get_waybills()
                    waybill_exists = False
                    for wb in existing_waybills:
                        if wb['number'] == waybill_data['number']:
                            waybill_exists = True
                            print(f"Путевой лист {waybill_data['number']} уже существует, пропускаем")
                            skipped_count += 1
                            break

                    if waybill_exists:
                        continue

                    # 5. Сохраняем путевой лист
                    result = self.db.create_waybill(waybill_data)
                    if result:
                        imported_count += 1
                        print(f"✓ Импортирован ПЛ: {waybill_data['number']}")
                    else:
                        print(f"✗ Ошибка сохранения ПЛ: {waybill_data['number']}")
                        skipped_count += 1

                except Exception as e:
                    error_msg = f"Строка {idx + 1}: {str(e)}"
                    print(error_msg)
                    errors.append(error_msg)
                    skipped_count += 1

            # Выводим итоги
            print(f"\n{'=' * 50}")
            print(f"ИМПОРТ ЗАВЕРШЕН")
            print(f"Импортировано: {imported_count}")
            print(f"Пропущено: {skipped_count}")

            if errors:
                print(f"\nОшибки ({len(errors)}):")
                for error in errors[:10]:
                    print(f"  {error}")
                if len(errors) > 10:
                    print(f"  ... и еще {len(errors) - 10} ошибок")

            return imported_count, skipped_count, errors

        except Exception as e:
            print(f"Ошибка чтения файла: {e}")
            import traceback
            traceback.print_exc()
            return 0, 0, [f"Ошибка чтения файла: {str(e)}"]

    def _parse_import_date(self, value):
        """Парсинг даты из различных форматов для импорта"""
        from datetime import datetime

        try:
            if pd.isnull(value):
                return datetime.now().strftime('%Y-%m-%d')

            if isinstance(value, datetime):
                return value.strftime('%Y-%m-%d')

            if isinstance(value, str):
                value = value.strip()
                if not value:
                    return datetime.now().strftime('%Y-%m-%d')

                # Пробуем разные форматы
                formats = [
                    '%Y-%m-%d',  # 2024-12-10
                    '%d.%m.%Y',  # 10.12.2024
                    '%d/%m/%Y',  # 10/12/2024
                    '%Y.%m.%d',  # 2024.12.10
                    '%d-%m-%Y',  # 10-12-2024
                ]

                for fmt in formats:
                    try:
                        dt = datetime.strptime(value, fmt)
                        return dt.strftime('%Y-%m-%d')
                    except:
                        continue

            return datetime.now().strftime('%Y-%m-%d')

        except Exception:
            return datetime.now().strftime('%Y-%m-%d')

    def _parse_import_datetime(self, value):
        """Парсинг даты-времени для импорта"""
        from datetime import datetime

        try:
            if pd.isnull(value):
                return datetime.now().strftime('%Y-%m-%d %H:%M')

            if isinstance(value, datetime):
                return value.strftime('%Y-%m-%d %H:%M')

            if isinstance(value, str):
                value = value.strip()
                if not value:
                    return datetime.now().strftime('%Y-%m-%d %H:%M')

                # Пробуем с временем
                formats_with_time = [
                    '%Y-%m-%d %H:%M',
                    '%d.%m.%Y %H:%M',
                    '%d/%m/%Y %H:%M',
                    '%Y-%m-%d %H:%M:%S',
                ]

                for fmt in formats_with_time:
                    try:
                        dt = datetime.strptime(value, fmt)
                        return dt.strftime('%Y-%m-%d %H:%M')
                    except:
                        continue

                # Пробуем только дату
                return self._parse_import_date(value) + ' 00:00'

            return datetime.now().strftime('%Y-%m-%d %H:%M')

        except Exception:
            return datetime.now().strftime('%Y-%m-%d %H:%M')

    def import_data(self):
        """Импорт данных"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Выберите файл для импорта", "",
                "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*.*)"
            )

            if not file_path:
                return  # Пользователь отменил выбор

            if not os.path.exists(file_path):
                QMessageBox.warning(self, "Ошибка", f"Файл не найден:\n{file_path}")
                return

            # Проверяем размер файла
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                QMessageBox.warning(self, "Ошибка", "Файл пустой")
                return

            # Диалог выбора типа данных
            dialog = QDialog(self)
            dialog.setWindowTitle("Выберите тип данных для импорта")
            layout = QVBoxLayout(dialog)

            combo = QComboBox()
            combo.addItems(["Водители", "Механики", "Автомобили", "Путевые листы", "Предприятия"])
            layout.addWidget(combo)

            # Информация о файле
            file_info = QLabel(f"Файл: {os.path.basename(file_path)}\nРазмер: {file_size:,} байт")
            file_info.setStyleSheet("font-style: italic; color: #666; padding: 5px;")
            layout.addWidget(file_info)

            buttons = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)

            if not dialog.exec():
                return  # Пользователь отменил

            data_type = combo.currentText()

            import_time_start = time.time()
            imported_count = 0
            error_messages = []

            try:
                if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                    if data_type == "Водители":
                        drivers = self.exporter.import_drivers_from_excel(file_path)
                        if drivers:
                            for driver in drivers:
                                try:
                                    if self.db.create_driver(driver):
                                        imported_count += 1
                                except Exception as e:
                                    error_messages.append(
                                        f"Ошибка создания водителя {driver.get('fio', 'N/A')}: {str(e)}")
                            self.refresh_drivers_table()
                        else:
                            QMessageBox.warning(self, "Предупреждение",
                                                "Не удалось прочитать данные водителей из файла")

                    elif data_type == "Механики":
                        mechanics = self.exporter.import_mechanics_from_excel(file_path)
                        if mechanics:
                            for mechanic in mechanics:
                                try:
                                    if self.db.create_mechanic(mechanic):
                                        imported_count += 1
                                except Exception as e:
                                    error_messages.append(
                                        f"Ошибка создания механика {mechanic.get('fio', 'N/A')}: {str(e)}")
                            self.refresh_mechanics_table()
                        else:
                            QMessageBox.warning(self, "Предупреждение",
                                                "Не удалось прочитать данные механиков из файла")

                    elif data_type == "Автомобили":
                        vehicles = self.exporter.import_vehicles_from_excel(file_path)
                        if vehicles:
                            for vehicle in vehicles:
                                try:
                                    if self.db.create_vehicle(vehicle):
                                        imported_count += 1
                                except Exception as e:
                                    error_messages.append(
                                        f"Ошибка создания автомобиля {vehicle.get('plate', 'N/A')}: {str(e)}")
                            self.refresh_vehicles_table()
                        else:
                            QMessageBox.warning(self, "Предупреждение",
                                                "Не удалось прочитать данные автомобилей из файла")


                    elif data_type == "Путевые листы":
                        # Полный импорт путевых листов
                        if not PANDAS_AVAILABLE:
                            QMessageBox.critical(self, "Ошибка",
                                                 "Для импорта путевых листов требуется библиотека pandas.\n"
                                                 "Установите её командой: pip install pandas")
                            return

                        imported_count, skipped_count, errors = self.import_waybills_from_excel_simple(file_path)

                        message = f"Импорт путевых листов завершен:\n\n"
                        message += f"✅ Импортировано: {imported_count}\n"
                        message += f"⏭️ Пропущено: {skipped_count}"

                        if errors:
                            message += f"\n\n❌ Ошибки ({len(errors)}):\n"
                            for i, error in enumerate(errors[:3]):
                                message += f"{i + 1}. {error}\n"
                            if len(errors) > 3:
                                message += f"... и еще {len(errors) - 3} ошибок"

                        QMessageBox.information(self, "Импорт путевых листов", message)
                        self.refresh_waybill_list()


                    elif data_type == "Предприятия":
                        companies = self.exporter.import_companies_from_excel(file_path)
                        if companies:
                            for company in companies:
                                try:
                                    if self.db.create_company(company):
                                        imported_count += 1
                                except Exception as e:
                                    error_messages.append(
                                        f"Ошибка создания предприятия {company.get('name', 'N/A')}: {str(e)}")
                            self.refresh_companies_table()
                        else:
                            QMessageBox.warning(self, "Предупреждение",
                                                "Не удалось прочитать данные предприятий из файла")

                elif file_path.endswith('.csv'):
                    # Для CSV файлов
                    result = self.exporter.import_from_csv(file_path, self.db)
                    if result:
                        imported_count = result.get('imported', 0)
                        error_messages = result.get('errors', [])
                    else:
                        QMessageBox.warning(self, "Предупреждение", "Не удалось импортировать данные из CSV файла")

                # Обновляем все данные
                self.load_data()

                import_time = time.time() - import_time_start

                # Формируем сообщение об итогах
                if imported_count > 0:
                    success_msg = f"✅ Импорт завершен успешно!\n"
                    success_msg += f"Импортировано записей: {imported_count}\n"
                    success_msg += f"Время импорта: {import_time:.2f} сек"

                    if error_messages:
                        success_msg += f"\n\n⚠️ Было {len(error_messages)} ошибок:\n"
                        for i, error in enumerate(error_messages[:3]):  # Показываем только 3 первые ошибки
                            success_msg += f"{i + 1}. {error}\n"
                        if len(error_messages) > 3:
                            success_msg += f"... и еще {len(error_messages) - 3} ошибок"

                    QMessageBox.information(self, "Импорт данных", success_msg)
                    self.status_bar.showMessage(
                        f"Импортировано {imported_count} записей из {os.path.basename(file_path)}", 5000)

                else:
                    if error_messages:
                        error_text = "❌ Импорт завершен с ошибками:\n\n"
                        for i, error in enumerate(error_messages[:5]):
                            error_text += f"{i + 1}. {error}\n"
                        if len(error_messages) > 5:
                            error_text += f"... и еще {len(error_messages) - 5} ошибок"
                    else:
                        error_text = "⚠️ Не удалось импортировать данные. Проверьте формат файла."

                    QMessageBox.warning(self, "Импорт данных", error_text)
                    self.status_bar.showMessage("Ошибка импорта данных", 5000)

            except Exception as e:
                QMessageBox.critical(self, "Ошибка импорта",
                                     f"Критическая ошибка при импорте данных:\n\n{str(e)}\n\n"
                                     f"Проверьте формат файла и соответствие данных.")
                import traceback
                traceback.print_exc()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при импорте: {str(e)}")

    def add_driver(self):
        """Добавление нового водителя"""
        dialog = DriverDialog(self)
        if dialog.exec():
            driver_data = dialog.get_data()
            self.db.create_driver(driver_data)
            self.refresh_drivers_table()
            self.load_data()

    def edit_driver(self):
        """Редактирование водителя"""
        selected_items = self.drivers_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите водителя")
            return

        row = selected_items[0].row()
        driver_id = int(self.drivers_table.item(row, 0).text())

        driver = self.db.get_driver(driver_id)
        if driver:
            dialog = DriverDialog(self, driver)
            if dialog.exec():
                driver_data = dialog.get_data()
                self.db.update_driver(driver_id, driver_data)
                self.refresh_drivers_table()
                self.load_data()

    def delete_driver(self):
        """Удаление водителя"""
        selected_items = self.drivers_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите водителя")
            return

        row = selected_items[0].row()
        driver_id = int(self.drivers_table.item(row, 0).text())
        driver_name = self.drivers_table.item(row, 1).text()

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить водителя {driver_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.db.delete_driver(driver_id)
            self.refresh_drivers_table()
            self.load_data()

    def add_mechanic(self):
        """Добавление нового механика"""
        dialog = MechanicDialog(self)
        if dialog.exec():
            mechanic_data = dialog.get_data()
            if not mechanic_data.get('fio'):
                QMessageBox.warning(self, "Ошибка", "ФИО механика обязательно!")
                return
            self.db.create_mechanic(mechanic_data)
            self.refresh_mechanics_table()
            self.load_data()

    def edit_mechanic(self):
        """Редактирование механика"""
        selected_items = self.mechanics_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите механика")
            return

        row = selected_items[0].row()
        mechanic_id = int(self.mechanics_table.item(row, 0).text())

        mechanic = self.db.get_mechanic(mechanic_id)
        if mechanic:
            dialog = MechanicDialog(self, mechanic)
            if dialog.exec():
                mechanic_data = dialog.get_data()
                self.db.update_mechanic(mechanic_id, mechanic_data)
                self.refresh_mechanics_table()
                self.load_data()

    def delete_mechanic(self):
        """Удаление механика"""
        selected_items = self.mechanics_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите механика")
            return

        row = selected_items[0].row()
        mechanic_id = int(self.mechanics_table.item(row, 0).text())
        mechanic_name = self.mechanics_table.item(row, 1).text()

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить механика {mechanic_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.db.delete_mechanic(mechanic_id)
            self.refresh_mechanics_table()
            self.load_data()

    def import_drivers(self):
        """Импорт водителей из Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            try:
                drivers = self.exporter.import_drivers_from_excel(file_path)
                for driver in drivers:
                    self.db.create_driver(driver)
                self.refresh_drivers_table()
                self.load_data()
                QMessageBox.information(self, "Импорт", f"Импортировано {len(drivers)} водителей")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при импорте: {str(e)}")

    def export_drivers(self):
        """Экспорт водителей в Excel"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                drivers = self.db.get_drivers()
                self.exporter.export_drivers_to_excel(file_path, drivers)
                QMessageBox.information(self, "Экспорт", f"Экспортировано {len(drivers)} водителей")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    def import_mechanics(self):
        """Импорт механиков из Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            try:
                mechanics = self.exporter.import_mechanics_from_excel(file_path)
                for mechanic in mechanics:
                    self.db.create_mechanic(mechanic)
                self.refresh_mechanics_table()
                self.load_data()
                QMessageBox.information(self, "Импорт", f"Импортировано {len(mechanics)} механиков")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при импорте: {str(e)}")

    def export_mechanics(self):
        """Экспорт механиков в Excel"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                mechanics = self.db.get_mechanics()
                self.exporter.export_mechanics_to_excel(file_path, mechanics)
                QMessageBox.information(self, "Экспорт", f"Экспортировано {len(mechanics)} механиков")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    def add_vehicle(self):
        """Добавление нового автомобиля"""
        dialog = VehicleDialog(self)
        if dialog.exec():
            vehicle_data = dialog.get_data()
            if not vehicle_data.get('company_id'):
                QMessageBox.warning(self, "Ошибка", "Автомобиль должен быть привязан к предприятию!")
                return
            self.db.create_vehicle(vehicle_data)
            self.refresh_vehicles_table()
            self.load_data()

    def edit_vehicle(self):
        """Редактирование автомобиля"""
        selected_items = self.vehicles_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите автомобиль")
            return

        row = selected_items[0].row()
        vehicle_id = int(self.vehicles_table.item(row, 0).text())

        vehicle = self.db.get_vehicle(vehicle_id)
        if vehicle:
            dialog = VehicleDialog(self, vehicle)
            if dialog.exec():
                vehicle_data = dialog.get_data()
                if not vehicle_data.get('company_id'):
                    QMessageBox.warning(self, "Ошибка", "Автомобиль должен быть привязан к предприятию!")
                    return
                self.db.update_vehicle(vehicle_id, vehicle_data)
                self.refresh_vehicles_table()
                self.load_data()

    def delete_vehicle(self):
        """Удаление автомобиля"""
        selected_items = self.vehicles_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите автомобиль")
            return

        row = selected_items[0].row()
        vehicle_id = int(self.vehicles_table.item(row, 0).text())
        vehicle_plate = self.vehicles_table.item(row, 3).text()

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить автомобиль {vehicle_plate}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.db.delete_vehicle(vehicle_id)
            self.refresh_vehicles_table()
            self.load_data()

    def import_vehicles(self):
        """Импорт автомобилей из Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            try:
                vehicles = self.exporter.import_vehicles_from_excel(file_path)
                for vehicle in vehicles:
                    self.db.create_vehicle(vehicle)
                self.refresh_vehicles_table()
                self.load_data()
                QMessageBox.information(self, "Импорт", f"Импортировано {len(vehicles)} автомобилей")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при импорте: {str(e)}")

    def export_vehicles(self):
        """Экспорт автомобилей в Excel"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                vehicles = self.db.get_vehicles()
                self.exporter.export_vehicles_to_excel(file_path, vehicles)
                QMessageBox.information(self, "Экспорт", f"Экспортировано {len(vehicles)} автомобилей")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    def browse_save_path(self):
        """Выбор папки сохранения"""
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку сохранения")
        if folder:
            self.save_path.setText(folder)

    def save_settings(self):
        """Сохранение настроек"""
        settings = QSettings("ПутевыеЛисты", "Менеджер")

        settings.setValue("printer", self.print_printer.currentText())
        settings.setValue("copies", self.print_copies.value())
        settings.setValue("orientation", self.print_orientation.currentText())

        settings.setValue("default_org_name", self.default_org_name.text())
        settings.setValue("default_org_address", self.default_org_address.toPlainText())
        settings.setValue("default_org_phone", self.default_org_phone.text())
        settings.setValue("default_org_inn", self.default_org_inn.text())

        settings.setValue("save_path", self.save_path.text())
        settings.setValue("auto_backup", self.auto_backup.isChecked())
        settings.setValue("backup_interval", self.backup_interval.value())

        self.status_bar.showMessage("Настройки сохранены")
        QMessageBox.information(self, "Настройки", "Настройки успешно сохранены")

    def load_settings(self):
        """Загрузка настроек"""
        settings = QSettings("ПутевыеЛисты", "Менеджер")

        printer = settings.value("printer", "По умолчанию")
        self.print_printer.setCurrentText(printer)

        copies = settings.value("copies", 1, type=int)
        self.print_copies.setValue(copies)

        orientation = settings.value("orientation", "Книжная")
        self.print_orientation.setCurrentText(orientation)

        self.default_org_name.setText(settings.value("default_org_name", ""))
        self.default_org_address.setText(settings.value("default_org_address", ""))
        self.default_org_phone.setText(settings.value("default_org_phone", ""))
        self.default_org_inn.setText(settings.value("default_org_inn", ""))

        self.save_path.setText(settings.value("save_path", str(Path.home() / "ПутевыеЛисты")))
        self.auto_backup.setChecked(settings.value("auto_backup", True, type=bool))
        self.backup_interval.setValue(settings.value("backup_interval", 7, type=int))

    def reset_settings(self):
        """Сброс настроек"""
        reply = QMessageBox.question(
            self, "Сброс настроек",
            "Сбросить все настройки к значениям по умолчанию?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            settings = QSettings("ПутевыеЛисты", "Менеджер")
            settings.clear()
            self.load_settings()

    def create_backup(self):
        """Создание резервной копии"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Создать резервную копию", "", "Backup Files (*.backup)"
        )

        if file_path:
            try:
                self.db.create_backup(file_path)
                self.status_bar.showMessage(f"Резервная копия создана: {file_path}")
                QMessageBox.information(self, "Резервное копирование", "Резервная копия успешно создана")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при создании резервной копии: {str(e)}")

    def restore_backup(self):
        """Восстановление из резервной копии"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Восстановить из резервной копии", "", "Backup Files (*.backup)"
        )

        if file_path:
            reply = QMessageBox.question(
                self, "Подтверждение",
                "Восстановление удалит все текущие данные. Продолжить?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                try:
                    self.db.restore_backup(file_path)
                    self.load_data()
                    self.status_bar.showMessage(f"Данные восстановлены из: {file_path}")
                    QMessageBox.information(self, "Восстановление", "Данные успешно восстановлены")
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Ошибка при восстановлении: {str(e)}")

    # В методе show_reports() замените на:
    def show_reports(self):
        """Показать отчёты с графиками"""
        if REPORTS_AVAILABLE:
            reports_dialog = ReportsDialog(self, self.db)
            reports_dialog.exec()
        else:
            QMessageBox.warning(self, "Ошибка", "Модуль отчётов недоступен")

    def generate_quick_report(self, period):
        """Генерация быстрого отчёта"""
        from datetime import datetime, timedelta

        today = datetime.now()

        if period == 'month':
            date_from = today.replace(day=1).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
            title = f"Отчёт за {today.strftime('%B %Y')}"

        elif period == 'week':
            date_from = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
            title = f"Отчёт за неделю {date_from} - {date_to}"

        elif period == 'year':
            date_from = today.replace(month=1, day=1).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
            title = f"Отчёт за {today.year} год"

        else:
            return

        if REPORTS_AVAILABLE:
            dialog = ReportsDialog(self, self.db)

            # Устанавливаем даты
            dialog.date_from.setDate(QDate.fromString(date_from, "yyyy-MM-dd"))
            dialog.date_to.setDate(QDate.fromString(date_to, "yyyy-MM-dd"))

            # Генерируем отчёт
            dialog.generate_report()
            dialog.exec()

    def show_about(self):
        """Показать информацию о программе"""
        about_text = """
        <h2>Путевые листы Менеджер v1.0</h2>
        <p>Программа для учета и печати путевых листов</p>
        <p>© 2024 Все права защищены</p>
        <p>Версия: 1.0.0</p>
        <p>Разработчик: Команда ПутевыеЛисты</p>
        """
        QMessageBox.about(self, "О программе", about_text)

    def show_numbering_settings(self):
        """Показать диалог настроек нумерации"""
        try:
            dialog = NumberingSettingsDialog(self.db, self)
            dialog.exec()
        except Exception as e:
            print(f"ERROR: Ошибка открытия настроек нумерации: {e}")
            QMessageBox.warning(self, "Ошибка",
                                f"Не удалось открыть настройки нумерации:\n{str(e)}")

    def export_history(self):
        """Экспорт истории изменений статусов"""
        try:
            # Диалог выбора периода
            dialog = QDialog(self)
            dialog.setWindowTitle("Экспорт истории")
            dialog.setModal(True)
            dialog.resize(400, 150)

            layout = QVBoxLayout(dialog)

            # Заголовок
            title = QLabel("Выберите период для экспорта истории:")
            title.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
            layout.addWidget(title)

            # Выбор периода
            period_layout = QHBoxLayout()
            period_layout.addWidget(QLabel("С:"))
            date_from = QDateEdit()
            date_from.setDate(QDate.currentDate().addMonths(-1))
            date_from.setCalendarPopup(True)
            date_from.setDisplayFormat("dd.MM.yyyy")
            period_layout.addWidget(date_from)

            period_layout.addWidget(QLabel("По:"))
            date_to = QDateEdit()
            date_to.setDate(QDate.currentDate())
            date_to.setCalendarPopup(True)
            date_to.setDisplayFormat("dd.MM.yyyy")
            period_layout.addWidget(date_to)

            layout.addLayout(period_layout)

            # Кнопки
            buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)

            if dialog.exec():
                from_date = date_from.date().toString("yyyy-MM-dd")
                to_date = date_to.date().toString("yyyy-MM-dd")

                # Получаем историю из БД
                history_data = self.db.get_status_history(from_date, to_date)

                if not history_data:
                    QMessageBox.information(self, "Экспорт", "Нет данных за выбранный период")
                    return

                # Диалог сохранения файла
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "Сохранить историю", "", "Excel Files (*.xlsx);;CSV Files (*.csv)"
                )

                if file_path:
                    if hasattr(self.exporter, 'export_history_to_excel'):
                        self.exporter.export_history_to_excel(file_path, history_data)
                    else:
                        self.export_history_to_csv(file_path, history_data)

                    QMessageBox.information(self, "Экспорт", f"История экспортирована в {file_path}")
                    self.status_bar.showMessage(f"История экспортирована")

        except Exception as e:
            print(f"ERROR: Ошибка при экспорте истории: {e}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте истории: {str(e)}")

    def export_history_to_csv(self, file_path, history_data):
        """Простой экспорт истории в CSV"""
        try:
            import csv

            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['ID', 'ID путевого листа', 'Номер путевого', 'Старый статус',
                              'Новый статус', 'Изменен пользователем', 'Дата изменения', 'Примечания']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')

                writer.writeheader()
                for item in history_data:
                    writer.writerow({
                        'ID': item.get('id', ''),
                        'ID путевого листа': item.get('waybill_id', ''),
                        'Номер путевого': item.get('waybill_number', ''),
                        'Старый статус': item.get('old_status', ''),
                        'Новый статус': item.get('new_status', ''),
                        'Изменен пользователем': item.get('changed_by', ''),
                        'Дата изменения': item.get('change_date', ''),
                        'Примечания': item.get('notes', '')
                    })

            print(f"INFO: История экспортирована в CSV: {file_path}")

        except Exception as e:
            print(f"ERROR: Ошибка при экспорте в CSV: {e}")
            raise

    def open_waybill(self):
        """Открытие путевого листа через диалог выбора"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Открыть путевой лист")
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)

        # Создаем список путевых листов
        waybills = self.db.get_waybills()

        list_widget = QTableWidget()
        list_widget.setColumnCount(5)
        list_widget.setHorizontalHeaderLabels(["Номер", "Дата", "Водитель", "Механик", "Автомобиль"])
        list_widget.setRowCount(len(waybills))

        for row, waybill in enumerate(waybills):
            list_widget.setItem(row, 0, QTableWidgetItem(str(waybill['number'])))
            list_widget.setItem(row, 1, QTableWidgetItem(str(waybill['date'])))
            list_widget.setItem(row, 2, QTableWidgetItem(str(waybill.get('driver_fio', ''))))
            list_widget.setItem(row, 3, QTableWidgetItem(str(waybill.get('mechanic_fio', ''))))
            list_widget.setItem(row, 4, QTableWidgetItem(str(waybill.get('vehicle_info', ''))))

        list_widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        list_widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        list_widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        list_widget.resizeColumnsToContents()

        layout.addWidget(list_widget)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec():
            selected_items = list_widget.selectedItems()
            if selected_items:
                row = selected_items[0].row()
                waybill_id = waybills[row]['id']
                self.open_waybill_by_id(waybill_id)

    def edit_print_mapping(self):
        """Редактирование маппинга полей печати"""
        if SMART_PRINTER_AVAILABLE:
            edit_mapping_dialog(self)
        else:
            QMessageBox.warning(self, "Ошибка", "Редактор маппинга недоступен")

    def closeEvent(self, event):
        """Обработка закрытия окна"""
        # Проверяем, есть ли несохраненные изменения
        if self.current_waybill_id:
            reply = QMessageBox.question(
                self, "Выход",
                "Сохранить изменения перед выходом?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.save_waybill()
                event.accept()
            elif reply == QMessageBox.StandardButton.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


class CompanyDialog(QDialog):
    """Диалог для редактирования предприятия"""

    def __init__(self, parent=None, company=None):
        super().__init__(parent)
        self.company = company
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Редактирование предприятия" if self.company else "Новое предприятие")
        self.setModal(True)
        self.resize(500, 400)

        layout = QFormLayout(self)

        self.name_edit = QLineEdit()
        self.inn_edit = QLineEdit()
        self.kpp_edit = QLineEdit()
        self.ogrn_edit = QLineEdit()
        self.address_edit = QTextEdit()
        self.address_edit.setMaximumHeight(80)
        self.phone_edit = QLineEdit()
        self.email_edit = QLineEdit()
        self.director_edit = QLineEdit()

        layout.addRow("Наименование*:", self.name_edit)
        layout.addRow("ИНН:", self.inn_edit)
        layout.addRow("КПП:", self.kpp_edit)
        layout.addRow("ОГРН:", self.ogrn_edit)
        layout.addRow("Адрес:", self.address_edit)
        layout.addRow("Телефон:", self.phone_edit)
        layout.addRow("E-mail:", self.email_edit)
        layout.addRow("Директор:", self.director_edit)

        if self.company:
            self.name_edit.setText(self.company.get('name', ''))
            self.inn_edit.setText(self.company.get('inn', ''))
            self.kpp_edit.setText(self.company.get('kpp', ''))
            self.ogrn_edit.setText(self.company.get('ogrn', ''))
            self.address_edit.setText(self.company.get('address', ''))
            self.phone_edit.setText(self.company.get('phone', ''))
            self.email_edit.setText(self.company.get('email', ''))
            self.director_edit.setText(self.company.get('director', ''))

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_data(self):
        """Получить данные из формы"""
        return {
            'name': self.name_edit.text(),
            'inn': self.inn_edit.text(),
            'kpp': self.kpp_edit.text(),
            'ogrn': self.ogrn_edit.text(),
            'address': self.address_edit.toPlainText(),
            'phone': self.phone_edit.text(),
            'email': self.email_edit.text(),
            'director': self.director_edit.text()
        }


class VehicleDialog(QDialog):
    """Диалог для редактирования автомобиля"""

    def __init__(self, parent=None, vehicle=None):
        super().__init__(parent)
        self.vehicle = vehicle
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Редактирование автомобиля" if self.vehicle else "Новый автомобиль")
        self.setModal(True)

        layout = QFormLayout(self)

        self.brand_edit = QLineEdit()
        self.model_edit = QLineEdit()
        self.plate_edit = QLineEdit()
        self.vin_edit = QLineEdit()
        self.year_edit = QSpinBox()
        self.year_edit.setMinimum(1900)
        self.year_edit.setMaximum(QDate.currentDate().year())
        self.year_edit.setValue(QDate.currentDate().year())
        self.mileage_edit = QSpinBox()
        self.mileage_edit.setMaximum(999999)
        self.color_edit = QLineEdit()
        self.driver_combo = QComboBox()
        self.company_combo = QComboBox()

        layout.addRow("Марка*:", self.brand_edit)
        layout.addRow("Модель:", self.model_edit)
        layout.addRow("Гос. номер*:", self.plate_edit)
        layout.addRow("VIN:", self.vin_edit)
        layout.addRow("Год выпуска:", self.year_edit)
        layout.addRow("Пробег (км):", self.mileage_edit)
        layout.addRow("Цвет:", self.color_edit)
        layout.addRow("Водитель:", self.driver_combo)
        layout.addRow("Предприятие*:", self.company_combo)

        # Загрузка водителей
        if self.parent():
            db = self.parent().db
            drivers = db.get_drivers()
            self.driver_combo.addItem("Не назначен", None)
            for driver in drivers:
                self.driver_combo.addItem(driver['fio'], driver['id'])

        # Загрузка предприятий
        if self.parent():
            db = self.parent().db
            companies = db.get_companies()
            self.company_combo.addItem("Выберите предприятие...", None)
            for company in companies:
                self.company_combo.addItem(company['name'], company['id'])

        if self.vehicle:
            self.brand_edit.setText(self.vehicle.get('brand', ''))
            self.model_edit.setText(self.vehicle.get('model', ''))
            self.plate_edit.setText(self.vehicle.get('plate', ''))
            self.vin_edit.setText(self.vehicle.get('vin', ''))
            self.year_edit.setValue(self.vehicle.get('year', QDate.currentDate().year()))
            self.mileage_edit.setValue(self.vehicle.get('mileage', 0))
            self.color_edit.setText(self.vehicle.get('color', ''))

            driver_id = self.vehicle.get('driver_id')
            if driver_id:
                index = self.driver_combo.findData(driver_id)
                if index >= 0:
                    self.driver_combo.setCurrentIndex(index)

            company_id = self.vehicle.get('company_id')
            if company_id:
                index = self.company_combo.findData(company_id)
                if index >= 0:
                    self.company_combo.setCurrentIndex(index)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_data(self):
        """Получить данные из формы"""
        return {
            'brand': self.brand_edit.text(),
            'model': self.model_edit.text(),
            'plate': self.plate_edit.text(),
            'vin': self.vin_edit.text(),
            'year': self.year_edit.value(),
            'mileage': self.mileage_edit.value(),
            'color': self.color_edit.text(),
            'driver_id': self.driver_combo.currentData(),
            'company_id': self.company_combo.currentData()
        }


class DriverDialog(QDialog):
    """Диалог для редактирования водителя"""

    def __init__(self, parent=None, driver=None):
        super().__init__(parent)
        self.driver = driver
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Редактирование водителя" if self.driver else "Новый водитель")
        self.setModal(True)
        self.resize(500, 450)

        layout = QFormLayout(self)

        self.fio_edit = QLineEdit()
        self.snils_edit = QLineEdit()
        self.license_edit = QLineEdit()
        self.license_class_edit = QLineEdit()
        self.license_issue_date_edit = QDateEdit()
        self.license_issue_date_edit.setDate(QDate.currentDate())
        self.license_issue_date_edit.setCalendarPopup(True)
        self.license_issue_date_edit.setDisplayFormat("dd.MM.yyyy")

        self.phone_edit = QLineEdit()
        self.medical_date_edit = QDateEdit()
        self.medical_date_edit.setDate(QDate.currentDate())
        self.medical_date_edit.setCalendarPopup(True)
        self.medical_date_edit.setDisplayFormat("dd.MM.yyyy")

        self.address_edit = QTextEdit()
        self.address_edit.setMaximumHeight(60)

        self.company_combo = QComboBox()
        self.company_combo.addItem("Выберите предприятие...", None)

        # Загрузка предприятий
        if self.parent():
            db = self.parent().db
            companies = db.get_companies()
            for company in companies:
                self.company_combo.addItem(company['name'], company['id'])

        layout.addRow("ФИО*:", self.fio_edit)
        layout.addRow("СНИЛС:", self.snils_edit)
        layout.addRow("В/у №:", self.license_edit)
        layout.addRow("Класс:", self.license_class_edit)
        layout.addRow("Дата выдачи в/у:", self.license_issue_date_edit)
        layout.addRow("Телефон:", self.phone_edit)
        layout.addRow("Медосмотр до:", self.medical_date_edit)
        layout.addRow("Адрес:", self.address_edit)
        layout.addRow("Предприятие:", self.company_combo)

        if self.driver:
            self.fio_edit.setText(self.driver.get('fio', ''))
            self.snils_edit.setText(self.driver.get('snils', ''))
            self.license_edit.setText(self.driver.get('license', ''))
            self.license_class_edit.setText(self.driver.get('license_class', ''))

            license_issue_date = self.driver.get('license_issue_date')
            if license_issue_date:
                self.license_issue_date_edit.setDate(QDate.fromString(license_issue_date, "yyyy-MM-dd"))

            self.phone_edit.setText(self.driver.get('phone', ''))

            medical_date = self.driver.get('medical_date')
            if medical_date:
                self.medical_date_edit.setDate(QDate.fromString(medical_date, "yyyy-MM-dd"))

            self.address_edit.setText(self.driver.get('address', ''))

            company_id = self.driver.get('company_id')
            if company_id:
                index = self.company_combo.findData(company_id)
                if index >= 0:
                    self.company_combo.setCurrentIndex(index)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_data(self):
        """Получить данные из формы"""
        return {
            'fio': self.fio_edit.text(),
            'snils': self.snils_edit.text(),
            'license': self.license_edit.text(),
            'license_class': self.license_class_edit.text(),
            'license_issue_date': self.license_issue_date_edit.date().toString("yyyy-MM-dd"),
            'phone': self.phone_edit.text(),
            'medical_date': self.medical_date_edit.date().toString("yyyy-MM-dd"),
            'address': self.address_edit.toPlainText(),
            'company_id': self.company_combo.currentData()
        }


class MechanicDialog(QDialog):
    """Диалог для редактирования механика"""

    def __init__(self, parent=None, mechanic=None):
        super().__init__(parent)
        self.mechanic = mechanic
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Редактирование механика" if self.mechanic else "Новый механик")
        self.setModal(True)
        self.resize(500, 400)

        layout = QFormLayout(self)

        self.fio_edit = QLineEdit()
        self.position_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.email_edit = QLineEdit()
        self.license_edit = QLineEdit()
        self.license_date_edit = QDateEdit()
        self.license_date_edit.setDate(QDate.currentDate())
        self.license_date_edit.setCalendarPopup(True)
        self.license_date_edit.setDisplayFormat("dd.MM.yyyy")

        self.company_combo = QComboBox()
        self.company_combo.addItem("Выберите предприятие...", None)

        # Загрузка предприятий
        if self.parent():
            db = self.parent().db
            companies = db.get_companies()
            for company in companies:
                self.company_combo.addItem(company['name'], company['id'])

        layout.addRow("ФИО*:", self.fio_edit)
        layout.addRow("Должность:", self.position_edit)
        layout.addRow("Телефон:", self.phone_edit)
        layout.addRow("E-mail:", self.email_edit)
        layout.addRow("Удостоверение №:", self.license_edit)
        layout.addRow("Дата выдачи:", self.license_date_edit)
        layout.addRow("Предприятие:", self.company_combo)

        if self.mechanic:
            self.fio_edit.setText(self.mechanic.get('fio', ''))
            self.position_edit.setText(self.mechanic.get('position', ''))
            self.phone_edit.setText(self.mechanic.get('phone', ''))
            self.email_edit.setText(self.mechanic.get('email', ''))
            self.license_edit.setText(self.mechanic.get('license_number', ''))

            license_date = self.mechanic.get('license_date')
            if license_date:
                self.license_date_edit.setDate(QDate.fromString(license_date, "yyyy-MM-dd"))

            company_id = self.mechanic.get('company_id')
            if company_id:
                index = self.company_combo.findData(company_id)
                if index >= 0:
                    self.company_combo.setCurrentIndex(index)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_data(self):
        """Получить данные из формы"""
        return {
            'fio': self.fio_edit.text(),
            'position': self.position_edit.text(),
            'phone': self.phone_edit.text(),
            'email': self.email_edit.text(),
            'license_number': self.license_edit.text(),
            'license_date': self.license_date_edit.date().toString("yyyy-MM-dd"),
            'company_id': self.company_combo.currentData()
        }


class ReportsDialog(QDialog):
    """Диалог отчетов"""

    def __init__(self, parent=None, db=None):
        super().__init__(parent)
        self.db = db
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Отчеты")
        self.setModal(True)
        self.resize(600, 400)

        layout = QVBoxLayout(self)

        # Выбор типа отчета
        self.report_type = QComboBox()
        self.report_type.addItems([
            "Пробег по автомобилям",
            "Расход топлива",
            "Активность водителей",
            "Активность механиков",
            "Путевые листы за период"
        ])
        layout.addWidget(self.report_type)

        # Период для отчета
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("С:"))
        self.date_from = QDateEdit()
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_from.setCalendarPopup(True)
        period_layout.addWidget(self.date_from)

        period_layout.addWidget(QLabel("По:"))
        self.date_to = QDateEdit()
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setCalendarPopup(True)
        period_layout.addWidget(self.date_to)

        layout.addLayout(period_layout)

        # Таблица результатов
        self.results_table = QTableWidget()
        layout.addWidget(self.results_table)

        # Кнопки
        button_layout = QHBoxLayout()

        btn_generate = QPushButton("Сформировать отчет")
        btn_generate.clicked.connect(self.generate_report)

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.clicked.connect(self.export_report)

        button_layout.addWidget(btn_generate)
        button_layout.addWidget(btn_export)
        button_layout.addStretch()

        layout.addLayout(button_layout)

    def generate_report(self):
        """Генерация отчета"""
        report_type = self.report_type.currentText()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")

        try:
            if report_type == "Пробег по автомобилям":
                data = self.db.get_mileage_report(date_from, date_to)
                self.display_report(data, ["Автомобиль", "Пробег (км)", "Количество рейсов"])

            elif report_type == "Расход топлива":
                data = self.db.get_fuel_report(date_from, date_to)
                self.display_report(data, ["Автомобиль", "Топливо", "Расход по норме", "Фактический расход"])

            elif report_type == "Активность водителей":
                data = self.db.get_driver_activity_report(date_from, date_to)
                self.display_report(data, ["Водитель", "Количество рейсов", "Общий пробег"])

            elif report_type == "Активность механиков":
                data = self.db.get_mechanic_activity_report(date_from, date_to)
                self.display_report(data, ["Механик", "Количество путевых листов", "Должность"])

            elif report_type == "Путевые листы за период":
                data = self.db.get_waybills_by_period(date_from, date_to)
                self.display_report(data, ["Номер", "Дата", "Водитель", "Механик", "Автомобиль", "Пробег"])

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при формировании отчета: {str(e)}")

    def display_report(self, data, headers):
        """Отображение отчета в таблице"""
        self.results_table.clear()
        self.results_table.setColumnCount(len(headers))
        self.results_table.setHorizontalHeaderLabels(headers)
        self.results_table.setRowCount(len(data))

        for row, item in enumerate(data):
            for col, value in enumerate(item.values()):
                self.results_table.setItem(row, col, QTableWidgetItem(str(value)))

    def export_report(self):
        """Экспорт отчета в Excel"""
        if self.results_table.rowCount() == 0:
            QMessageBox.warning(self, "Предупреждение", "Сначала сформируйте отчет")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчет", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                # Здесь будет код экспорта таблицы в Excel
                QMessageBox.information(self, "Экспорт", "Отчет успешно экспортирован")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")


class PrintPreviewDialog(QDialog):
    """Единый диалог для печати и предпросмотра"""

    def __init__(self, parent=None, waybill_data=None, db=None):
        super().__init__(parent)
        self.waybill_data = waybill_data
        self.db = db
        self.parent_app = parent
        self.init_ui()

    def __init__(self, parent=None, waybill_data=None, db=None):
        super().__init__(parent)
        self.waybill_data = waybill_data
        self.db = db
        self.parent_app = parent

        # Логируем открытие диалога
        if PRINT_LOGGER_AVAILABLE:
            print_logger.log_print_dialog_open(waybill_data)

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Печать путевого листа")
        self.setModal(True)
        self.resize(800, 600)

        layout = QVBoxLayout(self)

        # 1. Верхняя панель: Параметры печати
        params_group = QGroupBox("Параметры печати")
        params_layout = QGridLayout(params_group)

        # Выбор шаблона (строка 0)
        params_layout.addWidget(QLabel("Шаблон:"), 0, 0)
        self.template_combo = QComboBox()
        self.load_templates()
        params_layout.addWidget(self.template_combo, 0, 1, 1, 2)

        # Выбор принтера (строка 1)
        params_layout.addWidget(QLabel("Принтер:"), 1, 0)
        self.printer_combo = QComboBox()
        self.load_printers()
        params_layout.addWidget(self.printer_combo, 1, 1, 1, 2)

        # Количество копий и ориентация (строка 2)
        params_layout.addWidget(QLabel("Копий:"), 2, 0)
        self.copies_spin = QSpinBox()
        self.copies_spin.setMinimum(1)
        self.copies_spin.setMaximum(10)
        self.copies_spin.setValue(1)
        params_layout.addWidget(self.copies_spin, 2, 1)

        params_layout.addWidget(QLabel("Ориентация:"), 2, 2)
        self.orientation_combo = QComboBox()
        self.orientation_combo.addItems(["Книжная", "Альбомная"])
        params_layout.addWidget(self.orientation_combo, 2, 3)

        layout.addWidget(params_group)

        # 2. Середина: Предпросмотр данных
        preview_group = QGroupBox("Данные для печати")
        preview_layout = QVBoxLayout(preview_group)

        self.data_text = QTextEdit()
        self.data_text.setReadOnly(True)
        self.data_text.setMaximumHeight(200)
        preview_layout.addWidget(self.data_text)

        layout.addWidget(preview_group)

        # 3. Статус печати
        self.status_label = QLabel("Готов к печати")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                padding: 10px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                border-radius: 5px;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.status_label)

        # 4. Нижняя панель: Кнопки действий
        button_layout = QHBoxLayout()

        self.btn_preview = QPushButton("👁️ Создать предпросмотр")
        self.btn_preview.clicked.connect(self.open_preview)

        self.btn_print = QPushButton("🖨️ Печать")
        self.btn_print.clicked.connect(self.print_document)
        self.btn_print.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                padding: 10px 25px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        self.btn_print_direct = QPushButton("🖨️ Печать сразу")
        self.btn_print_direct.clicked.connect(self.print_direct)
        self.btn_print_direct.setToolTip("Печать без открытия Excel")
        self.btn_print_direct.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                padding: 10px 25px;
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)

        self.btn_save = QPushButton("💾 Сохранить")
        self.btn_save.clicked.connect(self.save_to_file)

        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)

        button_layout.addWidget(self.btn_preview)
        button_layout.addWidget(self.btn_print)
        button_layout.addWidget(self.btn_print_direct)
        button_layout.addWidget(self.btn_save)
        button_layout.addStretch()
        button_layout.addWidget(self.btn_cancel)

        layout.addLayout(button_layout)

        # Заполняем данные для предпросмотра
        self.update_preview_data()

    def load_templates(self):
        """Загрузка доступных шаблонов"""
        self.template_combo.clear()
        self.template_combo.addItem("По умолчанию", "default")

        # Проверяем наличие шаблонов
        script_dir = os.path.dirname(os.path.abspath(__file__))
        templates_dir = os.path.join(script_dir, "data", "templates")

        if os.path.exists(templates_dir):
            for file in os.listdir(templates_dir):
                if file.endswith('.xlsx') and not file.startswith('~$'):
                    display_name = file.replace('.xlsx', '').replace('_', ' ')
                    self.template_combo.addItem(display_name, os.path.join(templates_dir, file))

    def load_printers(self):
        """Загрузка доступных принтеров"""
        self.printer_combo.clear()
        self.printer_combo.addItem("По умолчанию", "")

        try:
            # Пробуем получить список принтеров
            import win32print
            printers = win32print.EnumPrinters(2)  # PRINTER_ENUM_LOCAL
            for printer in printers:
                self.printer_combo.addItem(printer[2], printer[2])
        except ImportError:
            # Если нет win32print, добавляем стандартные варианты
            self.printer_combo.addItems(["Microsoft Print to PDF", "Отправить в OneNote"])
        except Exception as e:
            print(f"Не удалось получить список принтеров: {e}")
            self.printer_combo.addItem("Системный принтер", "")

    def update_preview_data(self):
        """Обновление отображаемых данных"""
        if not self.waybill_data:
            return

        data_text = f"""
        <div style="font-family: Arial; line-height: 1.6;">
        <h3>ПУТЕВОЙ ЛИСТ №{self.waybill_data.get('number', '')}</h3>
        <b>Дата:</b> {self.waybill_data.get('date', '')}<br>
        <b>Водитель:</b> {self.waybill_data.get('driver_fio', '')}<br>
        <b>Автомобиль:</b> {self.waybill_data.get('vehicle_info', '')}<br>
        <b>Механик:</b> {self.waybill_data.get('mechanic_fio', 'Не указан')}<br>
        <b>Пробег:</b> {self.waybill_data.get('mileage', 0)} км<br>
        <b>Топливо:</b> {self.waybill_data.get('fuel_type', '')}<br>
        <b>Статус:</b> {self.waybill_data.get('status', '')}
        </div>
        """
        self.data_text.setHtml(data_text)

    def print_document(self):
        """Обычная печать (открывает в Excel)"""
        import time
        start_time = time.time()

        try:
            # Логируем начало
            if PRINT_LOGGER_AVAILABLE:
                print_logger.log_print_start(
                    self.waybill_data,
                    printer=self.printer_combo.currentText(),
                    copies=self.copies_spin.value(),
                    template=self.template_combo.currentText()
                )

            self.status_label.setText("⏳ Подготовка файла...")
            self.btn_print.setEnabled(False)
            self.btn_print_direct.setEnabled(False)

            # Создаем файл
            file_path = self.save_to_temp_file()
            if not file_path or not os.path.exists(file_path):
                raise Exception("Не удалось создать файл")

            # Логируем создание файла
            if PRINT_LOGGER_AVAILABLE:
                print_logger.log_file_created(self.waybill_data, file_path)

            # Открываем в Excel
            if sys.platform == "win32":
                os.startfile(file_path)
                self.status_label.setText("✅ Файл открыт в Excel. Нажмите Ctrl+P для печати")
            else:
                import subprocess
                subprocess.call(['xdg-open', file_path])
                self.status_label.setText("✅ Файл открыт. Нажмите Ctrl+P для печати")

            # Логируем открытие Excel
            if PRINT_LOGGER_AVAILABLE:
                print_logger.log_excel_opened(self.waybill_data, file_path)
                elapsed_ms = int((time.time() - start_time) * 1000)
                print_logger.log_print_complete(
                    self.waybill_data,
                    success=True,
                    file_path=file_path,
                    elapsed_ms=elapsed_ms
                )

            # Восстанавливаем кнопки через 3 секунды
            QTimer.singleShot(3000, lambda: self.enable_buttons())

        except Exception as e:
            error_msg = str(e)
            self.status_label.setText(f"❌ Ошибка: {error_msg[:50]}")
            self.enable_buttons()

            # Логируем ошибку
            if PRINT_LOGGER_AVAILABLE:
                elapsed_ms = int((time.time() - start_time) * 1000)
                print_logger.log_print_complete(
                    self.waybill_data,
                    success=False,
                    error=error_msg,
                    elapsed_ms=elapsed_ms
                )

            print(f"Ошибка печати: {e}")

    def print_direct(self):
        """Прямая печать без открытия Excel"""
        import time
        start_time = time.time()

        try:
            print("DEBUG: Нажата кнопка 'Печать сразу'")

            # Логируем начало
            if PRINT_LOGGER_AVAILABLE:
                print_logger.log_print_start(
                    self.waybill_data,
                    printer=self.printer_combo.currentText(),
                    copies=self.copies_spin.value(),
                    template=self.template_combo.currentText()
                )

            self.status_label.setText("⏳ Начинаю печать...")
            self.btn_print.setEnabled(False)
            self.btn_print_direct.setEnabled(False)

            # Пробуем прямую печать
            if self.try_direct_print():
                print("DEBUG: Печать успешно запущена")

                # Логируем успех
                if PRINT_LOGGER_AVAILABLE:
                    elapsed_ms = int((time.time() - start_time) * 1000)
                    print_logger.log_print_complete(
                        self.waybill_data,
                        success=True,
                        elapsed_ms=elapsed_ms
                    )

                # Закрываем диалог через 2 секунды
                QTimer.singleShot(2000, self.accept)
            else:
                self.status_label.setText("⚠️ Открываю Excel...")
                print("DEBUG: Прямая печать не сработала, открываю Excel")

                # Логируем fallback
                if PRINT_LOGGER_AVAILABLE:
                    print_logger.log_operation(
                        operation="Прямая печать не сработала, открываю Excel",
                        waybill_data=self.waybill_data,
                        success=False,
                        module="PrintPreviewDialog"
                    )

                # Открываем файл в Excel как запасной вариант
                self.open_preview()
                QTimer.singleShot(3000, lambda: self.enable_buttons())

        except Exception as e:
            error_msg = str(e)
            self.status_label.setText(f"❌ Ошибка: {error_msg[:50]}")
            self.enable_buttons()

            # Логируем ошибку
            if PRINT_LOGGER_AVAILABLE:
                elapsed_ms = int((time.time() - start_time) * 1000)
                print_logger.log_print_complete(
                    self.waybill_data,
                    success=False,
                    error=error_msg,
                    elapsed_ms=elapsed_ms
                )

            print(f"Ошибка прямой печати: {e}")

    def print_with_win32(self, file_path, printer_name):
        """Печать через win32print (только для Windows)"""
        try:
            import win32api
            import win32print

            # Если принтер не указан, используем принтер по умолчанию
            if not printer_name:
                printer_name = win32print.GetDefaultPrinter()

            # Команда печати
            win32api.ShellExecute(
                0,
                "print",
                file_path,
                f'/d:"{printer_name}"',
                ".",
                0
            )
            return True

        except ImportError:
            print("win32api не установлен")
            return False
        except Exception as e:
            print(f"Ошибка win32 печати: {e}")
            return False

    def print_with_shell(self, file_path):
        """Печать через стандартный диалог печати"""
        try:
            if sys.platform == "win32":
                # Открываем диалог печати
                os.startfile(file_path, "print")
                self.status_label.setText("✅ Открыт диалог печати")
            else:
                # Для Linux/Mac
                import subprocess
                subprocess.call(['lp', file_path])
                self.status_label.setText("✅ Отправлено на печать (lp)")

            QTimer.singleShot(3000, lambda: self.enable_buttons())

        except Exception as e:
            print(f"Ошибка shell печати: {e}")
            # Открываем файл как запасной вариант
            os.startfile(file_path)
            self.status_label.setText("✅ Файл открыт. Печатайте вручную")
            self.enable_buttons()

    def open_preview(self):
        """Создание и открытие предпросмотра"""
        import time
        start_time = time.time()

        try:
            self.status_label.setText("⏳ Создаю файл...")

            # Создаем файл с уникальным именем
            file_path = self.create_simple_excel()

            if file_path and os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                self.status_label.setText(f"✅ Файл создан: {file_name}")

                # Логируем создание файла
                if PRINT_LOGGER_AVAILABLE:
                    print_logger.log_file_created(self.waybill_data, file_path)
                    elapsed_ms = int((time.time() - start_time) * 1000)
                    print_logger.log_operation(
                        operation="Предпросмотр создан",
                        waybill_data=self.waybill_data,
                        success=True,
                        file_path=file_path,
                        elapsed_ms=elapsed_ms,
                        module="PrintPreviewDialog"
                    )

                # Открываем файл для просмотра
                if sys.platform == "win32":
                    os.startfile(file_path)
                else:
                    import subprocess
                    subprocess.call(['xdg-open', file_path])
            else:
                self.status_label.setText("❌ Не удалось создать файл")

                # Логируем ошибку
                if PRINT_LOGGER_AVAILABLE:
                    elapsed_ms = int((time.time() - start_time) * 1000)
                    print_logger.log_operation(
                        operation="Не удалось создать файл для предпросмотра",
                        waybill_data=self.waybill_data,
                        success=False,
                        elapsed_ms=elapsed_ms,
                        module="PrintPreviewDialog"
                    )

        except Exception as e:
            error_msg = str(e)
            self.status_label.setText(f"❌ Ошибка: {error_msg[:50]}")

            # Логируем ошибку
            if PRINT_LOGGER_AVAILABLE:
                elapsed_ms = int((time.time() - start_time) * 1000)
                print_logger.log_operation(
                    operation="Ошибка создания предпросмотра",
                    waybill_data=self.waybill_data,
                    success=False,
                    error=error_msg,
                    elapsed_ms=elapsed_ms,
                    module="PrintPreviewDialog"
                )

            print(f"Ошибка предпросмотра: {e}")

    def save_to_file(self):
        """Сохранение в выбранную папку"""
        try:
            default_name = f"Путевой_лист_{self.waybill_data.get('number', '')}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Сохранить путевой лист",
                default_name,
                "Excel Files (*.xlsx);;All Files (*.*)"
            )

            if file_path:
                self.status_label.setText("⏳ Сохранение...")

                # Копируем временный файл или создаем новый
                temp_file = self.save_to_temp_file()
                if temp_file and os.path.exists(temp_file):
                    import shutil
                    shutil.copy2(temp_file, file_path)
                    self.status_label.setText(f"✅ Сохранено: {os.path.basename(file_path)}")
                    QMessageBox.information(self, "Сохранение", f"Файл сохранен:\n{file_path}")
                else:
                    # Создаем простой Excel
                    self.create_simple_excel(file_path)

        except Exception as e:
            self.status_label.setText(f"❌ Ошибка: {str(e)[:50]}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка сохранения: {str(e)}")

    def save_to_temp_file(self):
        """Сохранение во временный файл с уникальным именем"""
        try:
            # Пробуем использовать существующий принтер
            if SIMPLE_PRINTER_AVAILABLE:
                try:
                    print(f"DEBUG: Вызываем print_waybill_simple для ПЛ №{self.waybill_data.get('number', '')}")

                    # КРИТИЧЕСКО ИСПРАВЛЕНИЕ: Собираем ПОЛНЫЕ данные
                    complete_data = {}

                    # 1. Базовые данные из waybill_data
                    if self.waybill_data:
                        complete_data.update(self.waybill_data)

                    # 2. Данные организации (ВСЕ поля!)
                    if hasattr(self.parent_app, 'company_combo_waybill'):
                        company_id = self.parent_app.company_combo_waybill.currentData()
                        if company_id:
                            company = self.db.get_company(company_id)
                            if company:
                                # ВСЕ поля организации
                                org_fields = {
                                    # Основные поля
                                    'org_name': company.get('name', ''),
                                    'org_address': company.get('address', ''),
                                    'org_phone': company.get('phone', ''),  # ТЕЛЕФОН
                                    'org_inn': company.get('inn', ''),
                                    'org_kpp': company.get('kpp', ''),
                                    'org_ogrn': company.get('ogrn', ''),  # ОГРН
                                    'org_director': company.get('director', ''),

                                    # Альтернативные ключи для совместимости
                                    'company_name': company.get('name', ''),
                                    'company_address': company.get('address', ''),
                                    'company_phone': company.get('phone', ''),
                                    'company_ogrn': company.get('ogrn', ''),
                                    'organization_name': company.get('name', ''),
                                    'organization_address': company.get('address', ''),
                                    'organization_phone': company.get('phone', ''),
                                    'телефон_организации': company.get('phone', ''),
                                    'огрн_организации': company.get('ogrn', ''),
                                    'адрес_организации': company.get('address', '')
                                }
                                complete_data.update(org_fields)

                                # Отладка
                                print(f"DEBUG: Данные организации:")
                                for key, value in org_fields.items():
                                    if value:
                                        print(f"  {key}: {value}")

                    # 3. Данные водителя
                    if hasattr(self.parent_app, 'driver_combo'):
                        driver_id = self.parent_app.driver_combo.currentData()
                        if driver_id:
                            driver = self.db.get_driver(driver_id)
                            if driver:
                                complete_data.update({
                                    'driver_fio': driver.get('fio', ''),
                                    'driver_snils': driver.get('snils', ''),
                                    'driver_license': driver.get('license', ''),
                                    'driver_license_class': driver.get('license_class', ''),
                                    'driver_license_date': driver.get('license_issue_date', ''),
                                })

                    # 4. Данные автомобиля
                    if hasattr(self.parent_app, 'vehicle_combo'):
                        vehicle_id = self.parent_app.vehicle_combo.currentData()
                        if vehicle_id:
                            vehicle = self.db.get_vehicle(vehicle_id)
                            if vehicle:
                                complete_data.update({
                                    'vehicle_brand': vehicle.get('brand', ''),
                                    'vehicle_plate': vehicle.get('plate', ''),
                                    'vehicle_model': vehicle.get('model', ''),
                                })

                    # 5. Данные механика
                    if hasattr(self.parent_app, 'mechanic_combo'):
                        mechanic_id = self.parent_app.mechanic_combo.currentData()
                        if mechanic_id:
                            mechanic = self.db.get_mechanic(mechanic_id)
                            if mechanic:
                                complete_data.update({
                                    'mechanic_fio': mechanic.get('fio', ''),
                                })

                    # 6. Данные из формы (текущие значения)
                    if self.parent_app:
                        form_data = {
                            # Одометр
                            'odo_start': self.parent_app.vehicle_odo_start.value(),
                            'odo_end': self.parent_app.vehicle_odo_end.value(),
                            # Топливо
                            'fuel_start': self.parent_app.fuel_start.value(),
                            'fuel_issued': self.parent_app.fuel_issued.value(),
                            'fuel_end': self.parent_app.fuel_end.value(),
                            'fuel_type': self.parent_app.fuel_type.currentText(),
                            # Медосмотры
                            'medical_pre_date': self.parent_app.medical_pre_date.dateTime().toString(
                                "yyyy-MM-dd HH:mm"),
                            'medical_pre_doctor': self.parent_app.medical_pre_doctor.text(),
                            'medical_post_date': self.parent_app.medical_post_date.dateTime().toString(
                                "yyyy-MM-dd HH:mm"),
                            'medical_post_doctor': self.parent_app.medical_post_doctor.text(),
                            # Маршрут и примечания
                            'route': self.parent_app.route_text.toPlainText(),
                            'notes': self.parent_app.notes_text.toPlainText(),
                            # Вычисляемые поля
                            'mileage': max(0,
                                           self.parent_app.vehicle_odo_end.value() - self.parent_app.vehicle_odo_start.value())
                        }
                        complete_data.update(form_data)

                    # 7. СПЕЦИАЛЬНО ДЛЯ ЯЧЕЙКИ O29
                    # Определяем что писать в "распоряжение/заказчик/примечания"
                    o29_value = ""

                    # Вариант 1: Из примечаний в форме
                    if self.parent_app.notes_text.toPlainText():
                        o29_value = self.parent_app.notes_text.toPlainText()
                        print(f"DEBUG: O29 - берем из примечаний: {o29_value[:50]}...")

                    # Вариант 2: Из поля customer в данных
                    elif 'customer' in complete_data and complete_data['customer']:
                        o29_value = str(complete_data['customer'])
                        print(f"DEBUG: O29 - берем из customer: {o29_value[:50]}...")

                    # Вариант 3: Создаем из маршрута
                    elif 'route' in complete_data and complete_data['route']:
                        o29_value = f"Маршрут: {complete_data['route']}"
                        print(f"DEBUG: O29 - создаем из маршрута: {o29_value[:50]}...")

                    # Вариант 4: Стандартный текст
                    else:
                        o29_value = "Согласно утвержденному графику"
                        print("DEBUG: O29 - используем стандартный текст")

                    # Добавляем в данные под разными ключами
                    if o29_value:
                        complete_data.update({
                            'customer': o29_value,  # Основной ключ
                            'заказчик': o29_value,  # Русский
                            'распоряжение': o29_value,  # Прямой
                            'примечания': o29_value  # Альтернатива
                        })

                    # 8. ДЕБАГ: Выводим все собранные данные
                    print(f"DEBUG: ПОЛНЫЕ ДАННЫЕ ДЛЯ ПЕЧАТИ ({len(complete_data)} полей):")

                    # Сортируем по важности
                    important_fields = [
                        'org_name', 'org_address', 'org_phone', 'org_ogrn',
                        'driver_fio', 'driver_license', 'driver_license_date',
                        'vehicle_brand', 'vehicle_plate',
                        'odo_start', 'odo_end', 'mileage',
                        'route', 'customer', 'notes',
                        'mechanic_fio'
                    ]

                    for field in important_fields:
                        if field in complete_data:
                            value = complete_data[field]
                            if value or field in ['org_phone', 'org_ogrn', 'customer']:
                                print(f"  {field}: {value}")

                    # 9. Проверяем обязательные поля
                    required_fields = [
                        ('org_name', 'Название организации'),
                        ('org_address', 'Адрес организации'),
                        ('org_phone', 'Телефон организации'),
                        ('driver_fio', 'ФИО водителя'),
                        ('vehicle_plate', 'Гос. номер автомобиля'),
                        ('odo_start', 'Показания одометра (выезд)'),
                        ('odo_end', 'Показания одометра (возврат)')
                    ]

                    missing = []
                    for field_key, field_name in required_fields:
                        if not complete_data.get(field_key):
                            missing.append(field_name)

                    if missing:
                        print(f"⚠️  ВНИМАНИЕ! Отсутствуют: {', '.join(missing)}")

                    # 10. Вызов функции печати
                    print(f"DEBUG: Передаем данные в print_waybill_simple...")
                    result = print_waybill_simple(
                        waybill_data=complete_data,
                        db_manager=self.db
                    )

                    print(f"DEBUG: print_waybill_simple вернула: {result}")

                    if result and isinstance(result, str) and os.path.exists(result):
                        print(f"✅ Файл создан: {os.path.basename(result)}")
                        return result

                    print("DEBUG: print_waybill_simple вернула None или неверный путь")

                except Exception as e:
                    print(f"❌ Ошибка simple_printer: {e}")
                    import traceback
                    traceback.print_exc()

            # Запасной вариант
            return self.create_simple_excel()

        except Exception as e:
            print(f"❌ Ошибка сохранения во временный файл: {e}")
            import traceback
            traceback.print_exc()
            return None

    def try_direct_print(self):
        """Прямая печать без открытия Excel"""
        try:
            print("DEBUG: Прямая печать...")
            self.status_label.setText("⏳ Создаю файл...")

            # 1. Создаем файл
            file_path = self.save_to_temp_file()
            if not file_path or not os.path.exists(file_path):
                print("DEBUG: Не удалось создать файл")
                self.status_label.setText("❌ Не удалось создать файл")
                return False

            print(f"DEBUG: Файл создан: {os.path.basename(file_path)}")

            # 2. Печать на Windows
            if sys.platform == "win32":
                return self.print_windows_direct(file_path)
            else:
                return self.print_other_os_direct(file_path)

        except Exception as e:
            print(f"Ошибка прямой печати: {e}")
            self.status_label.setText(f"❌ Ошибка: {str(e)[:50]}")
            return False

    def print_windows_direct(self, file_path):
        """Прямая печать на Windows"""
        try:
            self.status_label.setText("⏳ Отправляю на печать...")

            # Получаем параметры
            printer_name = self.printer_combo.currentText()
            copies = self.copies_spin.value()

            print(f"DEBUG: Печать на {printer_name}, {copies} копий")

            # СПОСОБ 1: Через os.startfile с параметром "print"
            # Это откроет стандартный диалог печати Windows
            os.startfile(file_path, "print")

            print("DEBUG: Диалог печати Windows открыт")
            self.status_label.setText("✅ Диалог печати открыт")
            return True

        except Exception as e:
            print(f"Ошибка печати Windows: {e}")
            # Если не получилось, открываем файл
            try:
                os.startfile(file_path)
                self.status_label.setText("✅ Файл открыт. Печатайте вручную")
                return True
            except:
                return False

    def print_other_os_direct(self, file_path):
        """Прямая печать на Linux/Mac"""
        try:
            import subprocess

            copies = self.copies_spin.value()

            # Для Linux
            command = ['lp']
            if copies > 1:
                command.extend(['-n', str(copies)])
            command.append(file_path)

            result = subprocess.run(command, capture_output=True, text=True)

            if result.returncode == 0:
                print("DEBUG: Печать через lp успешна")
                self.status_label.setText("✅ Отправлено на печать")
                return True
            else:
                print(f"DEBUG: Ошибка lp: {result.stderr}")
                # Пробуем открыть файл
                subprocess.call(['xdg-open', file_path])
                self.status_label.setText("✅ Файл открыт. Печатайте вручную")
                return True

        except Exception as e:
            print(f"Ошибка печати Linux/Mac: {e}")
            return False

    def create_simple_excel(self, file_path=None):
        """Создание простого Excel файла с уникальным именем"""
        try:
            import pandas as pd
            from datetime import datetime
            import uuid

            # Создаем данные для Excel
            data = {
                'Путевой лист №': [self.waybill_data.get('number', '')],
                'Дата': [self.waybill_data.get('date', '')],
                'Водитель': [self.waybill_data.get('driver_fio', '')],
                'Автомобиль': [self.waybill_data.get('vehicle_info', '')],
                'Механик': [self.waybill_data.get('mechanic_fio', '')],
                'Пробег (км)': [self.waybill_data.get('mileage', 0)],
                'Одометр (выезд)': [self.waybill_data.get('odo_start', 0)],
                'Одометр (возврат)': [self.waybill_data.get('odo_end', 0)],
                'Топливо выдано (л)': [self.waybill_data.get('fuel_issued', 0)],
                'Тип топлива': [self.waybill_data.get('fuel_type', '')],
                'Статус': [self.waybill_data.get('status', '')]
            }

            df = pd.DataFrame(data)

            # Если путь не указан, создаем временный файл с уникальным именем
            if not file_path:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                save_dir = os.path.join(script_dir, "prints")
                os.makedirs(save_dir, exist_ok=True)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                unique_id = str(uuid.uuid4())[:8]
                waybill_num = self.waybill_data.get('number', 'unknown').replace('/', '_')

                file_path = os.path.join(save_dir,
                                         f"Путевой_лист_{waybill_num}_{timestamp}_{unique_id}.xlsx")

            df.to_excel(file_path, index=False)

            print(f"DEBUG: Создан файл: {os.path.basename(file_path)}")
            return file_path

        except Exception as e:
            print(f"Ошибка создания Excel: {e}")
            return None

    def closeEvent(self, event):
        """Обработка закрытия диалога"""
        if PRINT_LOGGER_AVAILABLE:
            print_logger.log_print_dialog_closed(self.waybill_data)
        super().closeEvent(event)

    def enable_buttons(self):
        """Включить кнопки после операции"""
        self.btn_print.setEnabled(True)
        self.btn_print_direct.setEnabled(True)


def main():
    """Точка входа в приложение"""
    app = QApplication(sys.argv)
    app.setStyle('Windows')

    window = WaybillManagerApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
