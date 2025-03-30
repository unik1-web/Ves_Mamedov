import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, 
                             QVBoxLayout, QWidget, QMessageBox, QTextEdit)
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice, QTimer

class WeightScaleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Программа для весового прибора МИДЛ МИ ВДА/12Я")
        self.setGeometry(100, 100, 600, 400)
        
        self.serial = QSerialPort()
        self.init_ui()
        self.init_serial()
        
    def init_ui(self):
        # Главный виджет и layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # Элементы интерфейса
        self.status_label = QLabel("Статус: Не подключено")
        self.weight_label = QLabel("Вес: ---")
        self.weight_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        self.connect_button = QPushButton("Подключить")
        self.connect_button.clicked.connect(self.toggle_connection)
        
        self.zero_button = QPushButton("Тара")
        self.zero_button.clicked.connect(self.send_zero_command)
        self.zero_button.setEnabled(False)
        
        # Добавление элементов в layout
        layout.addWidget(self.status_label)
        layout.addWidget(self.weight_label)
        layout.addWidget(self.connect_button)
        layout.addWidget(self.zero_button)
        layout.addWidget(QLabel("Лог:"))
        layout.addWidget(self.log_text)
        
        main_widget.setLayout(layout)
        
        # Таймер для периодического опроса весов
        self.timer = QTimer()
        self.timer.timeout.connect(self.read_data)
        self.timer.setInterval(500)  # Опрос каждые 500 мс
        
    def init_serial(self):
        # Поиск доступных COM-портов
        ports = QSerialPortInfo.availablePorts()
        if not ports:
            QMessageBox.warning(self, "Ошибка", "Не найдены доступные COM-порты!")
            return
            
        # Попробуем автоматически найти весы (для примера берем первый порт)
        port_info = ports[0]
        self.serial.setPort(port_info)
        self.serial.setBaudRate(QSerialPort.Baud9600)
        self.serial.setDataBits(QSerialPort.Data8)
        self.serial.setParity(QSerialPort.NoParity)
        self.serial.setStopBits(QSerialPort.OneStop)
        self.serial.setFlowControl(QSerialPort.NoFlowControl)
        
    def toggle_connection(self):
        if self.serial.isOpen():
            self.serial.close()
            self.timer.stop()
            self.connect_button.setText("Подключить")
            self.status_label.setText("Статус: Не подключено")
            self.zero_button.setEnabled(False)
            self.log_message("Отключено от весового прибора")
        else:
            if self.serial.open(QIODevice.ReadWrite):
                self.connect_button.setText("Отключить")
                self.status_label.setText(f"Статус: Подключено к {self.serial.portName()}")
                self.zero_button.setEnabled(True)
                self.timer.start()
                self.log_message(f"Подключено к {self.serial.portName()}")
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось открыть порт!")
                self.log_message(f"Ошибка подключения к {self.serial.portName()}")
    
    def read_data(self):
        if self.serial.isOpen() and self.serial.canReadLine():
            data = self.serial.readLine().data().decode().strip()
            self.process_weight_data(data)
    
    def process_weight_data(self, data):
        # Здесь нужно реализовать парсинг данных от вашего весового прибора
        # Это пример - вам нужно адаптировать под протокол вашего устройства
        
        try:
            # Пример данных: "W +123.45 kg"
            if data.startswith("W"):
                parts = data.split()
                weight = parts[1]
                unit = parts[2]
                self.weight_label.setText(f"Вес: {weight} {unit}")
                self.log_message(f"Получены данные: {data}")
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")
    
    def send_zero_command(self):
        # Отправка команды тары (зависит от протокола вашего устройства)
        command = "Z\r\n"  # Пример команды - нужно уточнить для вашего устройства
        if self.serial.isOpen():
            self.serial.write(command.encode())
            self.log_message(f"Отправлена команда тары: {command.strip()}")
    
    def log_message(self, message):
        self.log_text.append(message)
    
    def closeEvent(self, event):
        if self.serial.isOpen():
            self.serial.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WeightScaleApp()
    window.show()
    sys.exit(app.exec_())