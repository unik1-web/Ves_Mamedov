from kivy.app import App
from kivy.uix.tabbedpanel import TabbedPanel
from kivy.properties import (StringProperty, NumericProperty, 
                           ListProperty, BooleanProperty)
from kivy.core.window import Window
from kivy.clock import Clock
from serial import Serial, SerialException
from threading import Thread
from queue import Queue
from datetime import datetime

class WeightScaleApp(TabbedPanel):
    current_weight = StringProperty("---")
    status = StringProperty("Не подключено")
    log_text = StringProperty("")
    target_weight = NumericProperty(0)
    protocols = ListProperty(["Auto", "MIDL-MI-VDA", "ТОКВЕС SH-50", "Микросим М0601", "Ньютон 42"])
    is_connected = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.serial = None
        self.data_queue = Queue()
        self.protocol_info = {
            "MIDL-MI-VDA": {"zero_cmd": "Z\r\n", "cal_cmd": "CAL {}\r\n"},
            "ТОКВЕС SH-50": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Микросим М0601": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Ньютон 42": {"zero_cmd": "Z\r\n", "cal_cmd": "C {}\r\n"}
        }
        Clock.schedule_interval(self.process_queue, 0.1)

    def process_queue(self, dt):
        while not self.data_queue.empty():
            data = self.data_queue.get()
            self.process_weight_data(data)

    def toggle_connection(self):
        if self.is_connected:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        try:
            self.serial = Serial(
                port=self.ids.port_spinner.text,
                baudrate=int(self.ids.baud_spinner.text),
                bytesize=int(self.ids.databits_spinner.text),
                parity=self.ids.parity_spinner.text[0],
                stopbits=float(self.ids.stopbits_spinner.text),
                timeout=0.1
            )
            self.is_connected = True
            self.status = f"Подключено к {self.ids.port_spinner.text}"
            Thread(target=self.read_serial, daemon=True).start()
            self.log_message(f"Подключено к {self.ids.port_spinner.text}")
        except Exception as e:
            self.log_message(f"Ошибка подключения: {str(e)}")

    def disconnect(self):
        if self.serial:
            self.serial.close()
        self.is_connected = False
        self.status = "Не подключено"
        self.log_message("Отключено от весового прибора")

    def read_serial(self):
        while self.is_connected and self.serial:
            try:
                line = self.serial.readline().decode('ascii', errors='ignore').strip()
                if line:
                    self.data_queue.put(line)
            except:
                pass

    def process_weight_data(self, data):
        try:
            protocol = self.ids.protocol_spinner.text
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
                self.current_weight = f"{weight:.3f} кг"
                self.check_target_weight(weight)
                self.log_message(f"Вес: {weight:.3f} кг")
        except Exception as e:
            self.log_message(f"Ошибка обработки: {str(e)}")

    def send_zero_command(self):
        if self.is_connected:
            protocol = self.ids.protocol_spinner.text
            cmd = self.protocol_info.get(protocol, {}).get("zero_cmd", "Z\r\n")
            self.serial.write(cmd.encode())
            self.log_message("Команда тары отправлена")

    def start_calibration(self):
        content = BoxLayout(orientation='vertical', spacing=10)
        weight_input = TextInput(hint_text='Эталонный вес (кг)', size_hint_y=None, height=40)
        
        def calibrate(_):
            try:
                weight = float(weight_input.text)
                protocol = self.ids.protocol_spinner.text
                cmd = self.protocol_info.get(protocol, {}).get("cal_cmd", "CAL {}\r\n").format(weight)
                self.serial.write(cmd.encode())
                self.log_message(f"Калибровка: {weight} кг")
                popup.dismiss()
            except ValueError:
                pass
                
        content.add_widget(weight_input)
        content.add_widget(Button(text='Калибровать', size_hint_y=None, height=40, on_press=calibrate))
        
        popup = Popup(title='Калибровка', content=content, size_hint=(0.8, 0.4))
        popup.open()

    def check_target_weight(self, weight):
        if self.target_weight > 0 and abs(weight - self.target_weight) < 0.001:
            self.log_message(f"Достигнут целевой вес: {self.target_weight} кг")
            self.target_weight = 0
            self.ids.target_input.text = ""

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text += f"[{timestamp}] {message}\n"

    def refresh_ports(self):
        ports = [f"COM{i}" for i in range(1, 17)]
        available_ports = []
        for port in ports:
            try:
                s = Serial(port)
                s.close()
                available_ports.append(port)
            except:
                pass
        self.ids.port_spinner.values = available_ports or ["Не найдены"]

class MainApp(App):
    def build(self):
        Window.clearcolor = (0.95, 0.95, 0.95, 1)
        return WeightScaleApp()

if __name__ == '__main__':
    MainApp().run()