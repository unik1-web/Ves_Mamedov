import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, 
                            QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                            QComboBox, QTextEdit, QLineEdit, QSpinBox, 
                            QGroupBox, QScrollArea, QMessageBox, QGridLayout, QInputDialog)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice

class WeightScaleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Программа для весовых приборов")
        self.setGeometry(100, 100, 800, 600)
        
        self.serial = QSerialPort()
        self.protocols = ["Auto", "MIDL-MI-VDA", "ТОКВЕС SH-50", "Микросим М0601", "Ньютон 42"]
        self.current_protocol = "Auto"
        self.protocol_info = {
            "MIDL-MI-VDA": {"zero_cmd": "Z\r\n", "cal_cmd": "CAL {}\r\n"},
            "ТОКВЕС SH-50": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Микросим М0601": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Ньютон 42": {"zero_cmd": "Z\r\n", "cal_cmd": "C {}\r\n"}
        }
        
        self.init_ui()
        self.init_serial()
        
    def init_ui(self):
        # Создаем главный виджет с вкладками
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Вкладка "Основное"
        self.main_tab = QWidget()
        self.init_main_tab()
        
        # Вкладка "Настройки"
        self.settings_tab = QWidget()
        self.init_settings_tab()
        
        self.tabs.addTab(self.main_tab, "Основное")
        self.tabs.addTab(self.settings_tab, "Настройки")
        
        # Таймер для опроса порта
        self.timer = QTimer()
        self.timer.timeout.connect(self.read_data)
        
    def init_main_tab(self):
        layout = QVBoxLayout()
        
        # Группа статуса
        status_group = QGroupBox("Статус подключения")
        status_layout = QVBoxLayout()
        self.status_label = QLabel("Статус: Не подключено")
        status_layout.addWidget(self.status_label)
        status_group.setLayout(status_layout)
        
        # Группа веса
        weight_group = QGroupBox("Измерения")
        weight_layout = QVBoxLayout()
        self.weight_label = QLabel("Вес: ---")
        self.weight_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        weight_layout.addWidget(self.weight_label)
        weight_group.setLayout(weight_layout)
        
        # Группа управления
        control_group = QGroupBox("Управление")
        control_layout = QHBoxLayout()
        self.connect_btn = QPushButton("Подключить")
        self.connect_btn.clicked.connect(self.toggle_connection)
        self.zero_btn = QPushButton("Тара")
        self.zero_btn.clicked.connect(self.send_zero_command)
        self.calibrate_btn = QPushButton("Калибровка")
        self.calibrate_btn.clicked.connect(self.start_calibration)
        control_layout.addWidget(self.connect_btn)
        control_layout.addWidget(self.zero_btn)
        control_layout.addWidget(self.calibrate_btn)
        control_group.setLayout(control_layout)
        
        # Лог
        log_group = QGroupBox("Лог")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout = QVBoxLayout()
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        
        # Добавляем все группы на вкладку
        layout.addWidget(status_group)
        layout.addWidget(weight_group)
        layout.addWidget(control_group)
        layout.addWidget(log_group)
        
        self.main_tab.setLayout(layout)
        
    def init_settings_tab(self):
        layout = QVBoxLayout()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        scroll.setWidget(content)
        layout.addWidget(scroll)
        
        settings_layout = QVBoxLayout(content)
        
        # Настройки порта
        port_group = QGroupBox("Настройки порта")
        port_layout = QGridLayout()
        
        port_layout.addWidget(QLabel("COM-порт:"), 0, 0)
        self.port_combo = QComboBox()
        self.refresh_ports()
        port_layout.addWidget(self.port_combo, 0, 1)
        
        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.clicked.connect(self.refresh_ports)
        port_layout.addWidget(self.refresh_btn, 0, 2)
        
        port_layout.addWidget(QLabel("Скорость (бод):"), 1, 0)
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "57600"])
        self.baud_combo.setCurrentText("9600")
        port_layout.addWidget(self.baud_combo, 1, 1)
        
        port_layout.addWidget(QLabel("Биты данных:"), 2, 0)
        self.databits_combo = QComboBox()
        self.databits_combo.addItems(["8", "7"])
        self.databits_combo.setCurrentText("8")
        port_layout.addWidget(self.databits_combo, 2, 1)
        
        port_layout.addWidget(QLabel("Четность:"), 3, 0)
        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["None", "Even", "Odd"])
        self.parity_combo.setCurrentText("None")
        port_layout.addWidget(self.parity_combo, 3, 1)
        
        port_layout.addWidget(QLabel("Стоп-биты:"), 4, 0)
        self.stopbits_combo = QComboBox()
        self.stopbits_combo.addItems(["1", "2"])
        self.stopbits_combo.setCurrentText("1")
        port_layout.addWidget(self.stopbits_combo, 4, 1)
        
        port_group.setLayout(port_layout)
        settings_layout.addWidget(port_group)
        
        # Настройки протокола
        protocol_group = QGroupBox("Настройки протокола")
        protocol_layout = QVBoxLayout()
        
        protocol_layout.addWidget(QLabel("Протокол весов:"))
        self.protocol_combo = QComboBox()
        self.protocol_combo.addItems(self.protocols)
        protocol_layout.addWidget(self.protocol_combo)
        
        protocol_group.setLayout(protocol_layout)
        settings_layout.addWidget(protocol_group)
        
        settings_layout.addStretch()
        self.settings_tab.setLayout(layout)
        
    def init_serial(self):
        self.serial.readyRead.connect(self.read_data)
        
    def refresh_ports(self):
        self.port_combo.clear()
        ports = QSerialPortInfo.availablePorts()
        if ports:
            for port in ports:
                self.port_combo.addItem(port.portName())
        else:
            self.port_combo.addItem("Не найдены")
    
    def toggle_connection(self):
        if self.serial.isOpen():
            self.disconnect()
        else:
            self.connect()
    
    def connect(self):
        port_name = self.port_combo.currentText()
        if port_name == "Не найдены":
            QMessageBox.critical(self, "Ошибка", "Нет доступных COM-портов!")
            return
            
        self.serial.setPortName(port_name)
        self.serial.setBaudRate(int(self.baud_combo.currentText()))
        self.serial.setDataBits(int(self.databits_combo.currentText()))
        
        parity = {
            "None": QSerialPort.NoParity,
            "Even": QSerialPort.EvenParity,
            "Odd": QSerialPort.OddParity
        }.get(self.parity_combo.currentText(), QSerialPort.NoParity)
        self.serial.setParity(parity)
        
        self.serial.setStopBits(int(self.stopbits_combo.currentText()))
        
        if self.serial.open(QIODevice.ReadWrite):
            self.connect_btn.setText("Отключить")
            self.status_label.setText(f"Подключено к {port_name}")
            self.zero_btn.setEnabled(True)
            self.calibrate_btn.setEnabled(True)
            self.timer.start(500)  # Опрос каждые 500 мс
            self.log_message(f"Подключено к {port_name}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось открыть порт!")
            self.log_message(f"Ошибка подключения к {port_name}")
    
    def disconnect(self):
        self.timer.stop()
        self.serial.close()
        self.connect_btn.setText("Подключить")
        self.status_label.setText("Не подключено")
        self.zero_btn.setEnabled(False)
        self.calibrate_btn.setEnabled(False)
        self.log_message("Отключено от весового прибора")
    
    def read_data(self):
        if self.serial.isOpen() and self.serial.canReadLine():
            data = self.serial.readLine().data().decode().strip()
            self.process_weight_data(data)
    
    def process_weight_data(self, data):
        protocol = self.protocol_combo.currentText()
        try:
            weight = None
            
            if protocol == "MIDL-MI-VDA" and data.startswith("W"):
                weight = float(data.split()[1])
            elif protocol == "ТОКВЕС SH-50" and data.startswith("ST,GS,"):
                weight = float(data.split(',')[2].strip().split()[0])
            elif protocol == "Микросим М0601" and data.startswith(('+', '-')):
                weight = float(data.split()[0])
            elif protocol == "Ньютон 42" and data.startswith('N'):
                weight = float(data[1:].split()[0])
            
            if weight is not None:
                self.weight_label.setText(f"Вес: {weight:.3f} кг")
                self.log_message(f"Получены данные: {data}")
        
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")
    
    def send_zero_command(self):
        protocol = self.protocol_combo.currentText()
        cmd = self.protocol_info.get(protocol, {}).get("zero_cmd", "Z\r\n")
        if self.serial.isOpen():
            self.serial.write(cmd.encode())
            self.log_message(f"Отправлена команда тары: {cmd.strip()}")
    
    def start_calibration(self):
        weight, ok = QInputDialog.getDouble(
            self, 'Калибровка', 
            'Введите эталонный вес (кг):', 
            decimals=3
        )
        
        if ok:
            protocol = self.protocol_combo.currentText()
            cmd = self.protocol_info.get(protocol, {}).get("cal_cmd", "CAL {}\r\n").format(weight)
            self.serial.write(cmd.encode())
            self.log_message(f"Начата калибровка с весом {weight} кг")
    
    def log_message(self, message):
        self.log_text.append(message)
    
    def closeEvent(self, event):
        if self.serial.isOpen():
            self.serial.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WeightScaleApp()
    window.show()
    sys.exit(app.exec_())