import sys
import csv
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, 
                            QVBoxLayout, QWidget, QMessageBox, QTextEdit,
                            QComboBox, QSpinBox, QHBoxLayout, QGroupBox,
                            QTabWidget, QFileDialog, QCheckBox, QLineEdit,
                            QColorDialog, QStyleFactory, QInputDialog)
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice, QTimer, Qt, QUrl
from PyQt5.QtChart import QChart, QChartView, QLineSeries, QValueAxis
from PyQt5.QtGui import QPainter, QColor, QFont
from PyQt5.QtMultimedia import QSoundEffect

class WeightScaleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Программа для весового прибора МИДЛ МИ ВДА/12Я")
        self.setGeometry(100, 100, 1000, 700)
        
        self.serial = QSerialPort()
        self.weight_history = []
        self.max_history_points = 100
        self.current_unit = 'kg'
        self.units = {'kg': 1.0, 'g': 1000.0, 'lb': 2.20462}
        self.protocols = ["Auto", "MIDL-MI-VDA", "A&D", "Sartorius", "Ohaus"]
        self.current_protocol = None
        self.target_weight = None
        self.sound_effect = QSoundEffect()
        self.sound_effect.setSource(QUrl.fromLocalFile("beep.wav"))
        
        self.init_ui()
        self.init_serial_settings()
        self.init_chart()
        self.load_settings()
        
    def init_ui(self):
        # Главный виджет и табы
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Создание вкладок
        self.main_tab = QWidget()
        self.settings_tab = QWidget()
        self.chart_tab = QWidget()
        self.protocol_tab = QWidget()
        self.appearance_tab = QWidget()
        
        self.tabs.addTab(self.main_tab, "Основное")
        self.tabs.addTab(self.settings_tab, "Настройки")
        self.tabs.addTab(self.chart_tab, "График")
        self.tabs.addTab(self.protocol_tab, "Протокол")
        self.tabs.addTab(self.appearance_tab, "Внешний вид")
        
        # Инициализация вкладок
        self.init_main_tab()
        self.init_settings_tab()
        self.init_chart_tab()
        self.init_protocol_tab()
        self.init_appearance_tab()
        
        # Меню
        self.init_menu()
        
    def init_menu(self):
        menubar = self.menuBar()
        
        # Меню Файл
        file_menu = menubar.addMenu('Файл')
        
        export_action = file_menu.addAction('Экспорт данных')
        export_action.triggered.connect(self.export_data)
        
        exit_action = file_menu.addAction('Выход')
        exit_action.triggered.connect(self.close)
        
        # Меню Настройки
        settings_menu = menubar.addMenu('Настройки')
        
        theme_menu = settings_menu.addMenu('Тема')
        for theme in QStyleFactory.keys():
            action = theme_menu.addAction(theme)
            action.triggered.connect(lambda _, t=theme: self.set_theme(t))
        
    def init_main_tab(self):
        layout = QVBoxLayout()
        
        # Группа статуса
        status_group = QGroupBox("Статус подключения")
        status_layout = QVBoxLayout()
        
        self.status_label = QLabel("Статус: Не подключено")
        self.protocol_label = QLabel("Протокол: Не определен")
        self.port_label = QLabel("Порт: Не выбран")
        self.settings_label = QLabel("Параметры: Не установлены")
        
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.protocol_label)
        status_layout.addWidget(self.port_label)
        status_layout.addWidget(self.settings_label)
        status_group.setLayout(status_layout)
        
        # Группа веса
        weight_group = QGroupBox("Измерения")
        weight_layout = QVBoxLayout()
        
        self.weight_label = QLabel("Вес: ---")
        self.weight_label.setStyleSheet("font-size: 32px; font-weight: bold;")
        
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(self.units.keys())
        self.unit_combo.currentTextChanged.connect(self.change_unit)
        
        # Настройка целевого веса
        target_layout = QHBoxLayout()
        self.target_weight_edit = QLineEdit()
        self.target_weight_edit.setPlaceholderText("Целевой вес")
        self.set_target_button = QPushButton("Установить")
        self.set_target_button.clicked.connect(self.set_target_weight)
        self.clear_target_button = QPushButton("Сбросить")
        self.clear_target_button.clicked.connect(self.clear_target_weight)
        
        target_layout.addWidget(QLabel("Целевой вес:"))
        target_layout.addWidget(self.target_weight_edit)
        target_layout.addWidget(self.set_target_button)
        target_layout.addWidget(self.clear_target_button)
        
        weight_layout.addWidget(self.weight_label)
        weight_layout.addWidget(self.unit_combo)
        weight_layout.addLayout(target_layout)
        weight_group.setLayout(weight_layout)
        
        # Группа управления
        control_group = QGroupBox("Управление")
        control_layout = QHBoxLayout()
        
        self.connect_button = QPushButton("Подключить")
        self.connect_button.clicked.connect(self.toggle_connection)
        
        self.zero_button = QPushButton("Тара")
        self.zero_button.clicked.connect(self.send_zero_command)
        self.zero_button.setEnabled(False)
        
        self.calibrate_button = QPushButton("Калибровка")
        self.calibrate_button.clicked.connect(self.start_calibration)
        self.calibrate_button.setEnabled(False)
        
        control_layout.addWidget(self.connect_button)
        control_layout.addWidget(self.zero_button)
        control_layout.addWidget(self.calibrate_button)
        control_group.setLayout(control_layout)
        
        # Лог
        log_group = QGroupBox("Лог")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        self.save_log_button = QPushButton("Сохранить лог в файл")
        self.save_log_button.clicked.connect(self.save_log_to_file)
        
        log_layout.addWidget(self.log_text)
        log_layout.addWidget(self.save_log_button)
        log_group.setLayout(log_layout)
        
        # Добавление всех групп на вкладку
        layout.addWidget(status_group)
        layout.addWidget(weight_group)
        layout.addWidget(control_group)
        layout.addWidget(log_group)
        
        self.main_tab.setLayout(layout)
        
        # Таймер для периодического опроса весов
        self.timer = QTimer()
        self.timer.timeout.connect(self.read_data)
        self.timer.setInterval(500)
    
    def init_settings_tab(self):
        layout = QVBoxLayout()
        
        # Настройки порта
        port_group = QGroupBox("Настройки порта")
        port_layout = QVBoxLayout()
        
        # Выбор COM-порта
        self.port_combo = QComboBox()
        self.refresh_ports()
        
        self.refresh_button = QPushButton("Обновить список портов")
        self.refresh_button.clicked.connect(self.refresh_ports)
        
        # Настройки соединения
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "4800", "2400", "1200", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText("9600")
        
        self.data_bits_combo = QComboBox()
        self.data_bits_combo.addItems(["5", "6", "7", "8"])
        self.data_bits_combo.setCurrentText("8")
        
        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["NoParity", "EvenParity", "OddParity", "SpaceParity", "MarkParity"])
        
        self.stop_bits_combo = QComboBox()
        self.stop_bits_combo.addItems(["1", "1.5", "2"])
        
        self.flow_control_combo = QComboBox()
        self.flow_control_combo.addItems(["NoFlowControl", "HardwareControl", "SoftwareControl"])
        
        # Добавление элементов
        port_layout.addWidget(QLabel("COM-порт:"))
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(self.refresh_button)
        port_layout.addWidget(QLabel("Скорость (бод):"))
        port_layout.addWidget(self.baud_combo)
        port_layout.addWidget(QLabel("Биты данных:"))
        port_layout.addWidget(self.data_bits_combo)
        port_layout.addWidget(QLabel("Четность:"))
        port_layout.addWidget(self.parity_combo)
        port_layout.addWidget(QLabel("Стоп-биты:"))
        port_layout.addWidget(self.stop_bits_combo)
        port_layout.addWidget(QLabel("Управление потоком:"))
        port_layout.addWidget(self.flow_control_combo)
        
        port_group.setLayout(port_layout)
        
        # Настройки графика
        chart_group = QGroupBox("Настройки графика")
        chart_layout = QVBoxLayout()
        
        self.history_points_spin = QSpinBox()
        self.history_points_spin.setRange(10, 1000)
        self.history_points_spin.setValue(self.max_history_points)
        self.history_points_spin.valueChanged.connect(self.update_history_size)
        
        chart_layout.addWidget(QLabel("Количество точек на графике:"))
        chart_layout.addWidget(self.history_points_spin)
        chart_group.setLayout(chart_layout)
        
        # Настройки звука
        sound_group = QGroupBox("Настройки звука")
        sound_layout = QVBoxLayout()
        
        self.sound_checkbox = QCheckBox("Звуковое оповещение")
        self.sound_checkbox.setChecked(True)
        
        self.test_sound_button = QPushButton("Проверить звук")
        self.test_sound_button.clicked.connect(self.play_sound)
        
        sound_layout.addWidget(self.sound_checkbox)
        sound_layout.addWidget(self.test_sound_button)
        sound_group.setLayout(sound_layout)
        
        # Добавление групп на вкладку
        layout.addWidget(port_group)
        layout.addWidget(chart_group)
        layout.addWidget(sound_group)
        layout.addStretch()
        
        self.settings_tab.setLayout(layout)
    
    def init_chart_tab(self):
        layout = QVBoxLayout()
        
        # Виджет графика
        self.chart_view = QChartView()
        self.chart_view.setRenderHint(QPainter.Antialiasing)
        
        # Кнопки экспорта
        export_layout = QHBoxLayout()
        self.export_csv_button = QPushButton("Экспорт в CSV")
        self.export_csv_button.clicked.connect(lambda: self.export_data('csv'))
        self.export_excel_button = QPushButton("Экспорт в Excel")
        self.export_excel_button.clicked.connect(lambda: self.export_data('excel'))
        
        export_layout.addWidget(self.export_csv_button)
        export_layout.addWidget(self.export_excel_button)
        
        layout.addWidget(self.chart_view)
        layout.addLayout(export_layout)
        self.chart_tab.setLayout(layout)
    
    def init_protocol_tab(self):
        layout = QVBoxLayout()
        
        # Выбор протокола
        protocol_group = QGroupBox("Настройки протокола")
        protocol_layout = QVBoxLayout()
        
        self.protocol_combo = QComboBox()
        self.protocol_combo.addItems(self.protocols)
        self.protocol_combo.currentTextChanged.connect(self.change_protocol)
        
        self.detect_button = QPushButton("Автоопределение")
        self.detect_button.clicked.connect(self.detect_protocol)
        
        protocol_layout.addWidget(QLabel("Протокол весов:"))
        protocol_layout.addWidget(self.protocol_combo)
        protocol_layout.addWidget(self.detect_button)
        protocol_group.setLayout(protocol_layout)
        
        # Описание протокола
        self.protocol_info = QTextEdit()
        self.protocol_info.setReadOnly(True)
        
        layout.addWidget(protocol_group)
        layout.addWidget(QLabel("Описание протокола:"))
        layout.addWidget(self.protocol_info)
        
        self.protocol_tab.setLayout(layout)
    
    def init_appearance_tab(self):
        layout = QVBoxLayout()
        
        # Настройки цветов
        color_group = QGroupBox("Настройки цветов")
        color_layout = QVBoxLayout()
        
        # Цвет фона
        self.bg_color_button = QPushButton("Цвет фона")
        self.bg_color_button.clicked.connect(lambda: self.change_color('background'))
        
        # Цвет текста
        self.text_color_button = QPushButton("Цвет текста")
        self.text_color_button.clicked.connect(lambda: self.change_color('text'))
        
        # Цвет графика
        self.chart_color_button = QPushButton("Цвет графика")
        self.chart_color_button.clicked.connect(lambda: self.change_color('chart'))
        
        # Сброс цветов
        self.reset_colors_button = QPushButton("Сбросить цвета")
        self.reset_colors_button.clicked.connect(self.reset_colors)
        
        color_layout.addWidget(self.bg_color_button)
        color_layout.addWidget(self.text_color_button)
        color_layout.addWidget(self.chart_color_button)
        color_layout.addWidget(self.reset_colors_button)
        color_group.setLayout(color_layout)
        
        # Настройки шрифта
        font_group = QGroupBox("Настройки шрифта")
        font_layout = QVBoxLayout()
        
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 24)
        self.font_size_spin.setValue(12)
        self.font_size_spin.valueChanged.connect(self.change_font_size)
        
        self.font_family_combo = QComboBox()
        self.font_family_combo.addItems(["Arial", "Times New Roman", "Courier New", "Verdana"])
        self.font_family_combo.currentTextChanged.connect(self.change_font_family)
        
        font_layout.addWidget(QLabel("Размер шрифта:"))
        font_layout.addWidget(self.font_size_spin)
        font_layout.addWidget(QLabel("Шрифт:"))
        font_layout.addWidget(self.font_family_combo)
        font_group.setLayout(font_layout)
        
        layout.addWidget(color_group)
        layout.addWidget(font_group)
        layout.addStretch()
        
        self.appearance_tab.setLayout(layout)
    
    def init_chart(self):
        self.chart = QChart()
        self.chart.setTitle("Изменение веса во времени")
        
        # Создание серии для графика
        self.series = QLineSeries()
        self.series.setName("Вес")
        self.series.setColor(QColor(70, 130, 180))  # SteelBlue
        
        self.chart.addSeries(self.series)
        
        # Оси
        self.axisX = QValueAxis()
        self.axisX.setTitleText("Время (с)")
        self.axisX.setLabelFormat("%.1f")
        
        self.axisY = QValueAxis()
        self.axisY.setTitleText("Вес (kg)")
        
        self.chart.addAxis(self.axisX, Qt.AlignBottom)
        self.chart.addAxis(self.axisY, Qt.AlignLeft)
        
        self.series.attachAxis(self.axisX)
        self.series.attachAxis(self.axisY)
        
        self.chart_view.setChart(self.chart)
    
    def init_serial_settings(self):
        # Установка параметров по умолчанию
        self.serial.setBaudRate(QSerialPort.Baud9600)
        self.serial.setDataBits(QSerialPort.Data8)
        self.serial.setParity(QSerialPort.NoParity)
        self.serial.setStopBits(QSerialPort.OneStop)
        self.serial.setFlowControl(QSerialPort.NoFlowControl)
    
    def refresh_ports(self):
        self.port_combo.clear()
        ports = QSerialPortInfo.availablePorts()
        if ports:
            for port in ports:
                self.port_combo.addItem(port.portName(), port)
        else:
            self.port_combo.addItem("Порты не найдены")
    
    def update_history_size(self, size):
        self.max_history_points = size
        if len(self.weight_history) > size:
            self.weight_history = self.weight_history[-size:]
            self.update_chart()
    
    def toggle_connection(self):
        if self.serial.isOpen():
            self.serial.close()
            self.timer.stop()
            self.connect_button.setText("Подключить")
            self.status_label.setText("Статус: Не подключено")
            self.zero_button.setEnabled(False)
            self.calibrate_button.setEnabled(False)
            self.log_message("Отключено от весового прибора")
        else:
            if self.port_combo.currentText() == "Порты не найдены":
                QMessageBox.critical(self, "Ошибка", "Нет доступных COM-портов!")
                return
                
            port_name = self.port_combo.currentText()
            self.serial.setPortName(port_name)
            
            # Установка параметров соединения из интерфейса
            self.apply_serial_settings()
            
            if self.serial.open(QIODevice.ReadWrite):
                self.connect_button.setText("Отключить")
                self.status_label.setText(f"Статус: Подключено к {port_name}")
                self.port_label.setText(f"Порт: {port_name}")
                self.update_settings_label()
                self.zero_button.setEnabled(True)
                self.calibrate_button.setEnabled(True)
                self.timer.start()
                self.log_message(f"Подключено к {port_name}")
                self.log_message(f"Параметры: {self.settings_label.text()}")
                
                # Попытка автоопределения протокола
                if self.protocol_combo.currentText() == "Auto":
                    self.detect_protocol()
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось открыть порт!")
                self.log_message(f"Ошибка подключения к {port_name}")
    
    def apply_serial_settings(self):
        # Битрейт
        baud_rate = int(self.baud_combo.currentText())
        self.serial.setBaudRate(baud_rate)
        
        # Биты данных
        data_bits = int(self.data_bits_combo.currentText())
        if data_bits == 5:
            self.serial.setDataBits(QSerialPort.Data5)
        elif data_bits == 6:
            self.serial.setDataBits(QSerialPort.Data6)
        elif data_bits == 7:
            self.serial.setDataBits(QSerialPort.Data7)
        else:
            self.serial.setDataBits(QSerialPort.Data8)
        
        # Четность
        parity = self.parity_combo.currentText()
        if parity == "NoParity":
            self.serial.setParity(QSerialPort.NoParity)
        elif parity == "EvenParity":
            self.serial.setParity(QSerialPort.EvenParity)
        elif parity == "OddParity":
            self.serial.setParity(QSerialPort.OddParity)
        elif parity == "SpaceParity":
            self.serial.setParity(QSerialPort.SpaceParity)
        elif parity == "MarkParity":
            self.serial.setParity(QSerialPort.MarkParity)
        
        # Стоп-биты
        stop_bits = self.stop_bits_combo.currentText()
        if stop_bits == "1":
            self.serial.setStopBits(QSerialPort.OneStop)
        elif stop_bits == "1.5":
            self.serial.setStopBits(QSerialPort.OneAndHalfStop)
        elif stop_bits == "2":
            self.serial.setStopBits(QSerialPort.TwoStop)
        
        # Управление потоком
        flow_control = self.flow_control_combo.currentText()
        if flow_control == "NoFlowControl":
            self.serial.setFlowControl(QSerialPort.NoFlowControl)
        elif flow_control == "HardwareControl":
            self.serial.setFlowControl(QSerialPort.HardwareControl)
        elif flow_control == "SoftwareControl":
            self.serial.setFlowControl(QSerialPort.SoftwareControl)
    
    def update_settings_label(self):
        settings_text = (
            f"Параметры: {self.baud_combo.currentText()} бод, "
            f"{self.data_bits_combo.currentText()} бит, "
            f"{self.parity_combo.currentText()}, "
            f"{self.stop_bits_combo.currentText()} стоп-бит, "
            f"{self.flow_control_combo.currentText()}"
        )
        self.settings_label.setText(settings_text)
    
    def read_data(self):
        if self.serial.isOpen() and self.serial.canReadLine():
            data = self.serial.readLine().data().decode().strip()
            self.process_weight_data(data)
    
    def process_weight_data(self, data):
        try:
            # Обработка данных в зависимости от протокола
            if self.current_protocol == "MIDL-MI-VDA":
                # Пример обработки для МИДЛ МИ ВДА/12Я
                if data.startswith("W"):
                    parts = data.split()
                    weight_kg = float(parts[1])
                    unit = parts[2]
                    self.process_weight_value(weight_kg, unit, data)
            
            elif self.current_protocol == "A&D":
                # Пример обработки для весов A&D
                if data.startswith("+"):
                    weight_kg = float(data[1:8]) / 1000
                    self.process_weight_value(weight_kg, "kg", data)
            
            elif self.current_protocol == "Sartorius":
                # Пример обработки для весов Sartorius
                if len(data) >= 7:
                    weight_kg = float(data) / 1000
                    self.process_weight_value(weight_kg, "kg", data)
            
            elif self.current_protocol == "Ohaus":
                # Пример обработки для весов Ohaus
                if data.startswith("ST,"):
                    parts = data.split(',')
                    weight_kg = float(parts[1])
                    self.process_weight_value(weight_kg, "kg", data)
            
            else:
                # Попытка автоопределения формата
                self.try_auto_detect_protocol(data)
        
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")
    
    def process_weight_value(self, weight_kg, unit, raw_data):
        # Конвертация в выбранную единицу измерения
        converted_weight = weight_kg * self.units[self.current_unit]
        
        self.weight_label.setText(f"Вес: {converted_weight:.3f} {self.current_unit}")
        
        # Проверка на достижение целевого веса
        if self.target_weight is not None and abs(weight_kg - self.target_weight) < 0.001:
            self.notify_target_weight_reached()
        
        # Добавление в историю для графика
        timestamp = len(self.weight_history) * 0.5  # Предполагаем интервал 0.5 сек
        self.weight_history.append((timestamp, weight_kg))
        
        if len(self.weight_history) > self.max_history_points:
            self.weight_history.pop(0)
        
        self.update_chart()
        
        self.log_message(f"Получены данные: {raw_data}")
    
    def update_chart(self):
        self.series.clear()
        
        if not self.weight_history:
            return
            
        min_x = self.weight_history[0][0]
        max_x = self.weight_history[-1][0]
        min_y = min(w[1] for w in self.weight_history)
        max_y = max(w[1] for w in self.weight_history)
        
        # Добавляем небольшой зазор по Y для лучшего отображения
        y_gap = (max_y - min_y) * 0.1 if max_y != min_y else 1.0
        min_y = max(0, min_y - y_gap)
        max_y = max_y + y_gap
        
        for x, y in self.weight_history:
            self.series.append(x, y)
        
        self.axisX.setRange(min_x, max_x)
        self.axisY.setRange(min_y, max_y)
    
    def send_zero_command(self):
        if self.current_protocol == "MIDL-MI-VDA":
            command = "Z\r\n"
        elif self.current_protocol == "A&D":
            command = "Z\r\n"
        elif self.current_protocol == "Sartorius":
            command = "T\r\n"
        elif self.current_protocol == "Ohaus":
            command = "Z\r\n"
        else:
            command = "Z\r\n"  # Команда по умолчанию
        
        if self.serial.isOpen():
            self.serial.write(command.encode())
            self.log_message(f"Отправлена команда тары: {command.strip()}")
    
    def start_calibration(self):
        weight, ok = QInputDialog.getDouble(
            self, 'Калибровка', 
            'Введите эталонный вес (кг):', 
            decimals=3
        )
        
        if ok:
            reply = QMessageBox.question(
                self, 'Калибровка', 
                f'Установите эталонный вес {weight} кг и нажмите OK',
                QMessageBox.Ok | QMessageBox.Cancel
            )
            
            if reply == QMessageBox.Ok:
                if self.current_protocol == "MIDL-MI-VDA":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "A&D":
                    command = f"C {weight}\r\n"
                elif self.current_protocol == "Sartorius":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "Ohaus":
                    command = f"CAL {weight}\r\n"
                else:
                    command = f"CAL {weight}\r\n"  # Команда по умолчанию
                
                self.serial.write(command.encode())
                self.log_message(f"Начата процедура калибровки с весом {weight} кг")
    
    def change_unit(self, unit):
        self.current_unit = unit
        self.log_message(f"Изменена единица измерения на {unit}")
    
    def set_target_weight(self):
        try:
            weight = float(self.target_weight_edit.text())
            self.target_weight = weight
            self.log_message(f"Установлен целевой вес: {weight} кг")
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Введите корректное значение веса")
    
    def clear_target_weight(self):
        self.target_weight = None
        self.target_weight_edit.clear()
        self.log_message("Целевой вес сброшен")
    
    def notify_target_weight_reached(self):
        if self.sound_checkbox.isChecked():
            self.play_sound()
        
        QMessageBox.information(
            self, "Целевой вес достигнут", 
            f"Достигнут целевой вес: {self.target_weight} кг"
        )
        self.clear_target_weight()
    
    def play_sound(self):
        if not self.sound_effect.isPlaying():
            self.sound_effect.play()
    
    def save_log_to_file(self):
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Сохранить лог", "", "Текстовые файлы (*.txt);;Все файлы (*)"
        )
        
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                self.log_message(f"Лог сохранен в файл: {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {str(e)}")
    
    def export_data(self, format=None):
        if not self.weight_history:
            QMessageBox.warning(self, "Ошибка", "Нет данных для экспорта")
            return
        
        if format is None:
            formats = "CSV (*.csv);;Excel (*.xlsx);;Все файлы (*)"
            file_name, selected_filter = QFileDialog.getSaveFileName(
                self, "Экспорт данных", "", formats
            )
            
            if not file_name:
                return
                
            if "CSV" in selected_filter:
                format = 'csv'
            elif "Excel" in selected_filter:
                format = 'excel'
            else:
                # По умолчанию CSV
                format = 'csv'
                if not file_name.endswith('.csv'):
                    file_name += '.csv'
        else:
            default_name = "weights_export." + format
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Экспорт данных", default_name, 
                f"{format.upper()} (*.{format})" if format else "Все файлы (*)"
            )
            
            if not file_name:
                return
        
        try:
            if format == 'csv':
                self.export_to_csv(file_name)
            elif format == 'excel':
                self.export_to_excel(file_name)
            
            self.log_message(f"Данные экспортированы в {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать данные: {str(e)}")
    
    def export_to_csv(self, file_name):
        with open(file_name, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Время (с)", "Вес (kg)"])
            
            for time, weight in self.weight_history:
                writer.writerow([time, weight])
    
    def export_to_excel(self, file_name):
        try:
            import openpyxl
        except ImportError:
            QMessageBox.critical(
                self, "Ошибка", 
                "Для экспорта в Excel требуется модуль openpyxl. Установите его командой: pip install openpyxl"
            )
            return
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Данные весов"
        
        # Заголовки
        ws.append(["Время (с)", "Вес (kg)"])
        
        # Данные
        for time, weight in self.weight_history:
            ws.append([time, weight])
        
        # Сохранение
        wb.save(file_name)
    
    def change_protocol(self, protocol):
        if protocol == "Auto":
            self.current_protocol = None
            self.protocol_info.setText("Режим автоопределения протокола. Программа будет пытаться автоматически определить формат данных.")
        elif protocol == "MIDL-MI-VDA":
            self.current_protocol = "MIDL-MI-VDA"
            self.protocol_info.setText(
                "Протокол МИДЛ МИ ВДА/12Я\n\n"
                "Формат данных:\n"
                "W +123.45 kg\n\n"
                "Команды:\n"
                "Z - тара\n"
                "CAL <вес> - калибровка с указанием эталонного веса"
            )
        elif protocol == "A&D":
            self.current_protocol = "A&D"
            self.protocol_info.setText(
                "Протокол весов A&D\n\n"
                "Формат данных:\n"
                "+00123.45\r\n\n"
                "Команды:\n"
                "Z - тара\n"
                "C <вес> - калибровка"
            )
        elif protocol == "Sartorius":
            self.current_protocol = "Sartorius"
            self.protocol_info.setText(
                "Протокол весов Sartorius\n\n"
                "Формат данных:\n"
                "00123.45\r\n\n"
                "Команды:\n"
                "T - тара\n"
                "CAL <вес> - калибровка"
            )
        elif protocol == "Ohaus":
            self.current_protocol = "Ohaus"
            self.protocol_info.setText(
                "Протокол весов Ohaus\n\n"
                "Формат данных:\n"
                "ST,+00123.45,kg\r\n\n"
                "Команды:\n"
                "Z - тара\n"
                "CAL <вес> - калибровка"
            )
        
        self.protocol_label.setText(f"Протокол: {protocol}")
        self.log_message(f"Установлен протокол: {protocol}")
    
    def detect_protocol(self):
        if not self.serial.isOpen():
            QMessageBox.warning(self, "Ошибка", "Сначала подключитесь к весам")
            return
        
        # Попытка определить протокол по ответу весов
        self.log_message("Попытка автоопределения протокола...")
        
        # Отправка команды запроса веса (может потребоваться адаптация)
        self.serial.write("W\r\n".encode())
        
        # Ждем ответ (в реальной реализации нужно использовать таймауты)
        if self.serial.waitForReadyRead(1000):
            data = self.serial.readLine().data().decode().strip()
            self.try_auto_detect_protocol(data)
        else:
            self.log_message("Не получен ответ от весов для автоопределения")
    
    def try_auto_detect_protocol(self, data):
        if not data:
            return
            
        # Попытка определить протокол по формату данных
        if data.startswith("W +") and "kg" in data:
            self.change_protocol("MIDL-MI-VDA")
            self.log_message("Автоопределен протокол: MIDL-MI-VDA")
        elif data.startswith("+") and len(data) >= 7 and data[1:].replace(".", "").isdigit():
            self.change_protocol("A&D")
            self.log_message("Автоопределен протокол: A&D")
        elif data.replace(".", "").isdigit() and len(data) >= 5:
            self.change_protocol("Sartorius")
            self.log_message("Автоопределен протокол: Sartorius")
        elif data.startswith("ST,") and "," in data:
            self.change_protocol("Ohaus")
            self.log_message("Автоопределен протокол: Ohaus")
        else:
            self.log_message("Не удалось определить протокол автоматически")
    
    def change_color(self, element):
        color = QColorDialog.getColor()
        if color.isValid():
            if element == 'background':
                self.setStyleSheet(f"background-color: {color.name()};")
                self.log_message(f"Изменен цвет фона на {color.name()}")
            elif element == 'text':
                self.setStyleSheet(f"color: {color.name()};")
                self.log_message(f"Изменен цвет текста на {color.name()}")
            elif element == 'chart':
                self.series.setColor(color)
                self.log_message(f"Изменен цвет графика на {color.name()}")
    
    def reset_colors(self):
        self.setStyleSheet("")
        self.series.setColor(QColor(70, 130, 180))  # SteelBlue
        self.log_message("Цвета сброшены к значениям по умолчанию")
    
    def change_font_size(self, size):
        font = self.font()
        font.setPointSize(size)
        self.setFont(font)
        self.log_message(f"Изменен размер шрифта на {size}")
    
    def change_font_family(self, family):
        font = self.font()
        font.setFamily(family)
        self.setFont(font)
        self.log_message(f"Изменен шрифт на {family}")
    
    def set_theme(self, theme_name):
        app = QApplication.instance()
        app.setStyle(QStyleFactory.create(theme_name))
        self.log_message(f"Установлена тема: {theme_name}")
    
    def load_settings(self):
        # Здесь можно реализовать загрузку настроек из файла
        pass
    
    def save_settings(self):
        # Здесь можно реализовать сохранение настроек в файл
        pass
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def closeEvent(self, event):
        if self.serial.isOpen():
            self.serial.close()
        self.save_settings()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Современный стиль по умолчанию
    
    window = WeightScaleApp()
    window.show()
    sys.exit(app.exec_())