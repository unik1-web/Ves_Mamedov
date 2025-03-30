import sys
import csv
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, 
                            QVBoxLayout, QWidget, QMessageBox, QTextEdit,
                            QComboBox, QSpinBox, QHBoxLayout, QGroupBox,
                            QTabWidget, QFileDialog, QCheckBox)
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice, QTimer, Qt
from PyQt5.QtChart import QChart, QChartView, QLineSeries, QValueAxis
from PyQt5.QtGui import QPainter

class WeightScaleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Программа для весового прибора МИДЛ МИ ВДА/12Я")
        self.setGeometry(100, 100, 800, 600)
        
        self.serial = QSerialPort()
        self.weight_history = []
        self.max_history_points = 100
        self.current_unit = 'kg'
        self.units = {'kg': 1.0, 'g': 1000.0, 'lb': 2.20462}
        
        self.init_ui()
        self.init_serial_settings()
        self.init_chart()
        
    def init_ui(self):
        # Главный виджет и табы
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Вкладка основного интерфейса
        self.main_tab = QWidget()
        self.tabs.addTab(self.main_tab, "Основное")
        
        # Вкладка настроек
        self.settings_tab = QWidget()
        self.tabs.addTab(self.settings_tab, "Настройки")
        
        # Вкладка графика
        self.chart_tab = QWidget()
        self.tabs.addTab(self.chart_tab, "График")
        
        # Инициализация вкладок
        self.init_main_tab()
        self.init_settings_tab()
        self.init_chart_tab()
        
    def init_main_tab(self):
        layout = QVBoxLayout()
        
        # Группа статуса
        status_group = QGroupBox("Статус подключения")
        status_layout = QVBoxLayout()
        
        self.status_label = QLabel("Статус: Не подключено")
        self.port_label = QLabel("Порт: Не выбран")
        self.settings_label = QLabel("Параметры: Не установлены")
        
        status_layout.addWidget(self.status_label)
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
        
        weight_layout.addWidget(self.weight_label)
        weight_layout.addWidget(self.unit_combo)
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
        
        # Добавление групп на вкладку
        layout.addWidget(port_group)
        layout.addWidget(chart_group)
        layout.addStretch()
        
        self.settings_tab.setLayout(layout)
    
    def init_chart_tab(self):
        layout = QVBoxLayout()
        
        # Виджет графика
        self.chart_view = QChartView()
        self.chart_view.setRenderHint(QPainter.Antialiasing)
        
        layout.addWidget(self.chart_view)
        self.chart_tab.setLayout(layout)
    
    def init_chart(self):
        self.chart = QChart()
        self.chart.setTitle("Изменение веса во времени")
        
        # Создание серии для графика
        self.series = QLineSeries()
        self.series.setName("Вес")
        
        self.chart.addSeries(self.series)
        
        # Оси
        self.axisX = QValueAxis()
        self.axisX.setTitleText("Время (с)")
        self.axisX.setLabelFormat("%.1f")
        
        self.axisY = QValueAxis()
        self.axisY.setTitleText("Вес")
        
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
            # Пример обработки данных - адаптируйте под ваш прибор
            if data.startswith("W"):
                parts = data.split()
                weight_kg = float(parts[1])
                unit = parts[2]
                
                # Конвертация в выбранную единицу измерения
                converted_weight = weight_kg * self.units[self.current_unit]
                
                self.weight_label.setText(f"Вес: {converted_weight:.3f} {self.current_unit}")
                
                # Добавление в историю для графика
                timestamp = len(self.weight_history) * 0.5  # Предполагаем интервал 0.5 сек
                self.weight_history.append((timestamp, weight_kg))
                
                if len(self.weight_history) > self.max_history_points:
                    self.weight_history.pop(0)
                
                self.update_chart()
                
                self.log_message(f"Получены данные: {data}")
        
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")
    
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
        command = "Z\r\n"  # Пример команды - адаптируйте под ваш прибор
        if self.serial.isOpen():
            self.serial.write(command.encode())
            self.log_message(f"Отправлена команда тары: {command.strip()}")
    
    def start_calibration(self):
        reply = QMessageBox.question(
            self, 'Калибровка', 
            'Подготовьте эталонный вес и нажмите OK для начала калибровки',
            QMessageBox.Ok | QMessageBox.Cancel
        )
        
        if reply == QMessageBox.Ok:
            # Отправка команды калибровки - адаптируйте под ваш прибор
            command = "CAL\r\n"
            self.serial.write(command.encode())
            self.log_message("Начата процедура калибровки")
    
    def change_unit(self, unit):
        self.current_unit = unit
        self.log_message(f"Изменена единица измерения на {unit}")
    
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
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def closeEvent(self, event):
        if self.serial.isOpen():
            self.serial.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WeightScaleApp()
    window.show()
    sys.exit(app.exec_())