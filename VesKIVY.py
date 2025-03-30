import sys
import csv
from datetime import datetime
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.tabbedpanel import TabbedPanel
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout
from kivy.clock import Clock
from kivy.graphics import Color, Rectangle
from kivy.core.window import Window
from kivy.properties import ObjectProperty, StringProperty, NumericProperty, ListProperty
from serial import Serial, SerialException
from threading import Thread
from queue import Queue

from kivy.uix.spinner import Spinner
from kivy.uix.label import Label
from kivy.uix.button import Button

# Стили для компактных элементов
class CompactSpinner(Spinner):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_color = (0.9, 0.9, 0.9, 1)
        self.size_hint_y = None
        self.height = 30
        self.font_size = 12

class CompactLabel(Label):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint_y = None
        self.height = 30
        self.font_size = 12
        self.halign = 'right'
        self.valign = 'middle'
        self.text_size = self.size
        self.padding_x = 5

class CompactButton(Button):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint_y = None
        self.height = 30
        self.font_size = 12
        self.background_normal = ''
        self.background_color = (0.8, 0.8, 0.8, 1)

class WeightScaleApp(TabbedPanel):
    serial_port = ObjectProperty(None)
    protocol = StringProperty("Auto")
    current_weight = StringProperty("---")
    status = StringProperty("Не подключено")
    log_text = StringProperty("")
    weight_history = ListProperty([])
    target_weight = NumericProperty(0)
    is_connected = False

    protocols = [
        "Auto", 
        "MIDL-MI-VDA", 
        "A&D", 
        "Sartorius", 
        "Ohaus", 
        "ТОКВЕС SH-50",
        "Микросим М0601",
        "Ньютон 42"
    ]

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.serial = None
        self.data_queue = Queue()
        self.protocol_info = {
            "MIDL-MI-VDA": {
                "description": "Протокол МИДЛ МИ ВДА/12Я\nФормат: W +123.45 kg\nКоманды: Z - тара, CAL - калибровка",
                "zero_cmd": "Z\r\n",
                "cal_cmd": "CAL {}\r\n"
            },
            "ТОКВЕС SH-50": {
                "description": "Протокол ТОКВЕС SH-50\nФормат: ST,GS,  +1.234 kg\nКоманды: T - тара, CAL - калибровка",
                "zero_cmd": "T\r\n",
                "cal_cmd": "CAL {}\r\n"
            },
            "Микросим М0601": {
                "description": "Протокол Микросим М0601\nФормат: +0001.234 kg\nКоманды: T - тара, CAL - калибровка",
                "zero_cmd": "T\r\n",
                "cal_cmd": "CAL {}\r\n"
            },
            "Ньютон 42": {
                "description": "Протокол Ньютон 42\nФормат: N+00012.345 kg\nКоманды: Z - тара, C - калибровка",
                "zero_cmd": "Z\r\n",
                "cal_cmd": "C {}\r\n"
            }
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
        port = self.ids.port_spinner.text
        if not port or port == "Не найдены":
            self.show_popup("Ошибка", "Не выбран COM-порт")
            return

        try:
            self.serial = Serial(
                port=port,
                baudrate=int(self.ids.baud_spinner.text),
                bytesize=int(self.ids.databits_spinner.text),
                parity=self.ids.parity_spinner.text[0],
                stopbits=float(self.ids.stopbits_spinner.text),
                timeout=0.1
            )
            self.is_connected = True
            self.status = f"Подключено к {port}"
            Thread(target=self.read_serial, daemon=True).start()
            self.ids.connect_btn.text = "Отключить"
            self.log_message(f"Подключено к {port}")
        except SerialException as e:
            self.show_popup("Ошибка", f"Не удалось подключиться: {str(e)}")
            self.log_message(f"Ошибка подключения: {str(e)}")

    def disconnect(self):
        if self.serial and self.serial.is_open:
            self.serial.close()
        self.is_connected = False
        self.status = "Не подключено"
        self.ids.connect_btn.text = "Подключить"
        self.log_message("Отключено от весового прибора")

    def read_serial(self):
        while self.is_connected and self.serial and self.serial.is_open:
            try:
                line = self.serial.readline().decode('ascii', errors='ignore').strip()
                if line:
                    self.data_queue.put(line)
            except:
                pass

    def process_weight_data(self, data):
        try:
            weight = None
            unit = "kg"
            
            if self.protocol == "MIDL-MI-VDA" and data.startswith("W"):
                parts = data.split()
                weight = float(parts[1])
            elif self.protocol == "ТОКВЕС SH-50" and data.startswith("ST,GS,"):
                parts = data.split(',')
                weight_str = parts[2].strip().split()[0]
                weight = float(weight_str)
            elif self.protocol == "Микросим М0601" and data.startswith(('+', '-')):
                weight_str = data.split()[0]
                weight = float(weight_str)
            elif self.protocol == "Ньютон 42" and data.startswith('N'):
                weight_str = data[1:].split()[0]
                weight = float(weight_str)
            
            if weight is not None:
                self.current_weight = f"{weight:.3f} {unit}"
                self.weight_history.append(weight)
                if len(self.weight_history) > 100:
                    self.weight_history.pop(0)
                self.log_message(f"Получено: {data}")
                
                if self.target_weight and abs(weight - self.target_weight) < 0.001:
                    self.notify_target_weight()
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")

    def send_zero_command(self):
        if not self.is_connected:
            return
            
        cmd = self.protocol_info.get(self.protocol, {}).get("zero_cmd", "Z\r\n")
        try:
            self.serial.write(cmd.encode())
            self.log_message(f"Отправлена команда тары: {cmd.strip()}")
        except:
            self.log_message("Ошибка отправки команды тары")

    def start_calibration(self):
        content = BoxLayout(orientation='vertical')
        weight_input = TextInput(hint_text='Эталонный вес (кг)', input_filter='float')
        
        def calibrate(instance):
            try:
                weight = float(weight_input.text)
                cmd = self.protocol_info.get(self.protocol, {}).get("cal_cmd", "CAL {}\r\n").format(weight)
                self.serial.write(cmd.encode())
                self.log_message(f"Начата калибровка с весом {weight} кг")
                popup.dismiss()
            except ValueError:
                pass
                
        content.add_widget(weight_input)
        content.add_widget(Button(text='Калибровать', on_press=calibrate))
        
        popup = Popup(title='Калибровка', content=content, size_hint=(0.8, 0.4))
        popup.open()

    def set_target_weight(self):
        try:
            self.target_weight = float(self.ids.target_input.text)
            self.log_message(f"Установлен целевой вес: {self.target_weight} кг")
        except ValueError:
            self.show_popup("Ошибка", "Введите корректное значение веса")

    def notify_target_weight(self):
        self.log_message(f"Достигнут целевой вес: {self.target_weight} кг")
        self.target_weight = 0
        self.ids.target_input.text = ""
        self.show_popup("Уведомление", "Достигнут целевой вес!")

    def show_popup(self, title, message):
        content = Button(text='OK', size_hint=(1, 0.2))
        popup = Popup(title=title, content=content, size_hint=(0.8, 0.4))
        content.bind(on_press=popup.dismiss)
        popup.open()

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text += f"[{timestamp}] {message}\n"

    def refresh_ports(self):
        ports = ["COM{}".format(i + 1) for i in range(256)]
        available_ports = []
        for port in ports:
            try:
                s = Serial(port)
                s.close()
                available_ports.append(port)
            except:
                pass
        self.ids.port_spinner.values = available_ports if available_ports else ["Не найдены"]

class MainApp(App):
    def build(self):
        Window.clearcolor = (0.95, 0.95, 0.95, 1)
        return WeightScaleApp()

if __name__ == '__main__':
    MainApp().run()