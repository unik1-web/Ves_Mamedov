from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.core.audio import SoundLoader
from kivy.lang import Builder
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.checkbox import CheckBox
import serial
import serial.tools.list_ports
from datetime import datetime

# Увеличиваем максимальное количество итераций для Clock
Clock.max_iteration = 200

# Загружаем KV файл один раз на уровне модуля
kv = '''
#:kivy 2.3.1

# Общие стили
<Label>:
    color: 0, 0, 0, 1
    font_size: '16sp'
    text_size: self.size
    halign: 'left'
    valign: 'middle'
    size_hint_y: None
    height: '44dp'

<Button>:
    background_color: 0.3, 0.6, 0.9, 1
    color: 1, 1, 1, 1
    font_size: '16sp'
    size_hint_y: None
    height: '44dp'

<TextInput>:
    background_color: 1, 1, 1, 1
    foreground_color: 0, 0, 0, 1
    cursor_color: 0.3, 0.6, 0.9, 1
    font_size: '16sp'
    padding: '10dp'
    multiline: False
    size_hint_y: None
    height: '44dp'

<Spinner>:
    background_color: 0.3, 0.6, 0.9, 1
    color: 1, 1, 1, 1
    font_size: '16sp'
    size_hint_y: None
    height: '44dp'

<Popup>:
    background: 'atlas://data/images/defaulttheme/button'
    title_color: 0, 0, 0, 1
    title_size: '18sp'
    separator_color: 0.3, 0.6, 0.9, 1
    size_hint_y: None
    height: '200dp'

<TabbedPanelItem>:
    background_color: 0.3, 0.6, 0.9, 1
    color: 1, 1, 1, 1
    font_size: '16sp'

# Основной интерфейс
<WeightScaleRoot>:
    do_default_tab: False
    tab_pos: 'top_left'
    tab_width: 200
    
    TabbedPanelItem:
        text: 'Основное'
        BoxLayout:
            orientation: 'vertical'
            padding: '10dp'
            spacing: '10dp'

            # Верхняя панель с информацией
            GridLayout:
                cols: 4
                size_hint_y: None
                height: '44dp'
                spacing: '10dp'
                
                Label:
                    text: 'Статус: Не подключено'
                    id: status_label
                    
                Label:
                    text: 'Протокол: Не определен'
                    id: protocol_label
                    
                Label:
                    text: 'Порт: Не выбран'
                    id: port_label
                    
                Button:
                    text: 'Подключить'
                    id: connect_button
                    on_press: app.toggle_connection()

            # Основная область с весом
            BoxLayout:
                orientation: 'vertical'
                size_hint_y: None
                height: '150dp'
                spacing: '10dp'
                
                Label:
                    text: 'Вес: ---'
                    id: weight_label
                    font_size: '48sp'
                    size_hint_y: None
                    height: '80dp'
                    
                GridLayout:
                    cols: 4
                    size_hint_y: None
                    height: '44dp'
                    spacing: '10dp'
                    
                    Spinner:
                        id: unit_spinner
                        text: 'kg'
                        values: ['kg', 'g', 'lb']
                        on_text: app.change_unit(self.text)
                        
                    TextInput:
                        id: target_weight
                        hint_text: 'Целевой вес'
                        
                    Button:
                        text: 'Установить'
                        on_press: app.set_target_weight()
                        
                    Button:
                        text: 'Сбросить'
                        on_press: app.clear_target_weight()

            # Панель управления
            GridLayout:
                cols: 2
                size_hint_y: None
                height: '44dp'
                spacing: '10dp'
                
                Button:
                    text: 'Тара'
                    id: zero_button
                    disabled: True
                    on_press: app.send_zero_command()
                    
                Button:
                    text: 'Калибровка'
                    id: calibrate_button
                    disabled: True
                    on_press: app.start_calibration()

            # Лог
            BoxLayout:
                orientation: 'vertical'
                spacing: '10dp'
                
                ScrollView:
                    TextInput:
                        id: log_text
                        readonly: True
                        multiline: True
                        background_color: 1, 1, 1, 1
                        foreground_color: 0, 0, 0, 1
                        
                Button:
                    text: 'Сохранить лог'
                    size_hint_y: None
                    height: '44dp'
                    on_press: app.save_log_to_file()

    TabbedPanelItem:
        text: 'Настройки'
        BoxLayout:
            orientation: 'vertical'
            padding: '10dp'
            spacing: '10dp'

            # Настройки порта
            GridLayout:
                cols: 2
                size_hint_y: None
                height: '220dp'
                spacing: '10dp'
                
                Label:
                    text: 'COM-порт:'
                Spinner:
                    id: port_spinner
                    text: 'Выберите порт'
                    values: []
                
                Label:
                    text: 'Скорость (бод):'
                Spinner:
                    id: baud_spinner
                    text: '9600'
                    values: ['1200', '2400', '4800', '9600', '19200', '38400', '57600', '115200']
                
                Label:
                    text: 'Биты данных:'
                Spinner:
                    id: databits_spinner
                    text: '8'
                    values: ['5', '6', '7', '8']
                
                Label:
                    text: 'Четность:'
                Spinner:
                    id: parity_spinner
                    text: 'Нет'
                    values: ['Нет', 'Чет', 'Нечет']
                
                Label:
                    text: 'Стоп-биты:'
                Spinner:
                    id: stopbits_spinner
                    text: '1'
                    values: ['1', '1.5', '2']

            # Настройки протокола
            GridLayout:
                cols: 2
                size_hint_y: None
                height: '88dp'
                spacing: '10dp'
                
                Label:
                    text: 'Протокол:'
                Spinner:
                    id: protocol_spinner
                    text: 'Auto'
                    values: ['Auto', 'MIDL-MI-VDA', 'A&D', 'Sartorius', 'Ohaus']
                    on_text: app.change_protocol(self.text)
                
                Label:
                    text: 'Звук:'
                CheckBox:
                    id: sound_checkbox
                    active: True

            # Кнопки управления настройками
            BoxLayout:
                size_hint_y: None
                height: '44dp'
                spacing: '10dp'
                
                Button:
                    text: 'Обновить список портов'
                    on_press: app.refresh_ports()
                
                Button:
                    text: 'Применить настройки'
                    on_press: app.apply_settings()
'''

Builder.load_string(kv)

class WeightScaleRoot(TabbedPanel):
    pass  # The layout is defined in the KV string

class WeightScaleApp(App):
    def get_application_config(self):
        return None  # Prevents loading of default config/kv file
    
    def build_config(self, config):
        pass  # Prevents loading of default config/kv file
        
    def __init__(self):
        super().__init__()
        self.serial = None
        self.weight_history = []
        self.max_history_points = 100
        self.current_unit = 'kg'
        self.units = {'kg': 1.0, 'g': 1000.0, 'lb': 2.20462}
        self.protocols = ["Auto", "MIDL-MI-VDA", "A&D", "Sartorius", "Ohaus"]
        self.current_protocol = None
        self.target_weight = None
        try:
            self.sound = SoundLoader.load('beep.wav')
            if not self.sound:
                print("Warning: Could not load beep.wav")
        except:
            print("Warning: Could not load beep.wav")
            self.sound = None
        self.update_event = None

    def build(self):
        Window.size = (800, 600)
        Window.clearcolor = (0.95, 0.95, 0.95, 1)
        return WeightScaleRoot()

    def on_start(self):
        # Инициализация таймера для обновления данных
        self.update_event = Clock.schedule_interval(self.read_data, 0.5)
        self.update_event.cancel()  # Начинаем с отключенным таймером
        self.refresh_ports()  # Перемещаем сюда вызов refresh_ports

    def refresh_ports(self):
        ports = list(serial.tools.list_ports.comports())
        self.root.ids.port_spinner.values = [port.device for port in ports]
        if ports:
            self.root.ids.port_spinner.text = ports[0].device
        else:
            self.root.ids.port_spinner.text = 'Нет портов'

    def apply_settings(self):
        if self.serial and self.serial.is_open:
            self.show_popup("Ошибка", "Сначала отключите устройство")
            return

        try:
            # Сохраняем настройки для использования при следующем подключении
            self.port = self.root.ids.port_spinner.text
            self.baudrate = int(self.root.ids.baud_spinner.text)
            self.databits = int(self.root.ids.databits_spinner.text)
            self.parity = {
                'Нет': serial.PARITY_NONE,
                'Чет': serial.PARITY_EVEN,
                'Нечет': serial.PARITY_ODD
            }[self.root.ids.parity_spinner.text]
            self.stopbits = float(self.root.ids.stopbits_spinner.text)
            
            self.log_message("Настройки сохранены")
        except Exception as e:
            self.show_popup("Ошибка", f"Ошибка применения настроек: {str(e)}")

    def change_protocol(self, protocol):
        self.current_protocol = None if protocol == "Auto" else protocol
        self.root.ids.protocol_label.text = f"Протокол: {protocol}"
        self.log_message(f"Установлен протокол: {protocol}")

    def read_data(self, dt):
        if self.serial and self.serial.is_open:
            try:
                if self.serial.in_waiting:
                    data = self.serial.readline().decode().strip()
                    self.process_weight_data(data)
            except Exception as e:
                self.log_message(f"Ошибка чтения данных: {str(e)}")

    def process_weight_data(self, data):
        try:
            if self.current_protocol == "MIDL-MI-VDA":
                if data.startswith("W"):
                    parts = data.split()
                    weight_kg = float(parts[1])
                    unit = parts[2]
                    self.process_weight_value(weight_kg, unit, data)
            
            elif self.current_protocol == "A&D":
                if data.startswith("+"):
                    weight_kg = float(data[1:8]) / 1000
                    self.process_weight_value(weight_kg, "kg", data)
            
            elif self.current_protocol == "Sartorius":
                if len(data) >= 7:
                    weight_kg = float(data) / 1000
                    self.process_weight_value(weight_kg, "kg", data)
            
            elif self.current_protocol == "Ohaus":
                if data.startswith("ST,"):
                    parts = data.split(',')
                    weight_kg = float(parts[1])
                    self.process_weight_value(weight_kg, "kg", data)
            
            else:
                self.try_auto_detect_protocol(data)
        
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")

    def process_weight_value(self, weight_kg, unit, raw_data):
        converted_weight = weight_kg * self.units[self.current_unit]
        self.root.ids.weight_label.text = f"Вес: {converted_weight:.3f} {self.current_unit}"
        
        if self.target_weight is not None and abs(weight_kg - self.target_weight) < 0.001:
            self.notify_target_weight_reached()
        
        self.log_message(f"Получены данные: {raw_data}")

    def change_unit(self, unit):
        self.current_unit = unit
        self.log_message(f"Изменена единица измерения на {unit}")

    def set_target_weight(self):
        try:
            weight = float(self.root.ids.target_weight.text)
            self.target_weight = weight
            self.log_message(f"Установлен целевой вес: {weight} кг")
        except ValueError:
            self.show_popup("Ошибка", "Введите корректное значение веса")

    def clear_target_weight(self):
        self.target_weight = None
        self.root.ids.target_weight.text = ""
        self.log_message("Целевой вес сброшен")

    def notify_target_weight_reached(self):
        if self.sound:
            self.sound.play()
        
        self.show_popup(
            "Целевой вес достигнут",
            f"Достигнут целевой вес: {self.target_weight} кг"
        )
        self.clear_target_weight()

    def toggle_connection(self):
        if self.serial and self.serial.is_open:
            self.serial.close()
            self.root.ids.connect_button.text = 'Подключить'
            self.root.ids.status_label.text = 'Статус: Не подключено'
            self.root.ids.zero_button.disabled = True
            self.root.ids.calibrate_button.disabled = True
            self.update_event.cancel()
            self.log_message("Отключено от весового прибора")
        else:
            try:
                # Используем настройки из интерфейса
                port = self.root.ids.port_spinner.text
                if port == 'Нет портов':
                    self.show_popup("Ошибка", "Нет доступных COM-портов!")
                    return

                baudrate = int(self.root.ids.baud_spinner.text)
                databits = int(self.root.ids.databits_spinner.text)
                parity = {
                    'Нет': serial.PARITY_NONE,
                    'Чет': serial.PARITY_EVEN,
                    'Нечет': serial.PARITY_ODD
                }[self.root.ids.parity_spinner.text]
                stopbits = float(self.root.ids.stopbits_spinner.text)

                self.serial = serial.Serial(
                    port=port,
                    baudrate=baudrate,
                    bytesize=databits,
                    parity=parity,
                    stopbits=stopbits,
                    timeout=1
                )

                self.root.ids.connect_button.text = 'Отключить'
                self.root.ids.status_label.text = f'Статус: Подключено к {port}'
                self.root.ids.port_label.text = f'Порт: {port}'
                self.root.ids.zero_button.disabled = False
                self.root.ids.calibrate_button.disabled = False
                self.update_event.start()
                
                settings_text = (
                    f"{baudrate} бод, "
                    f"{databits} бит, "
                    f"{self.root.ids.parity_spinner.text}, "
                    f"{stopbits} стоп-бит"
                )
                self.log_message(f"Подключено к {port} ({settings_text})")

            except Exception as e:
                self.show_popup("Ошибка", f"Не удалось подключиться: {str(e)}")

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
            command = "Z\r\n"
        
        if self.serial and self.serial.is_open:
            self.serial.write(command.encode())
            self.log_message(f"Отправлена команда тары: {command.strip()}")

    def start_calibration(self):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        weight_input = TextInput(
            multiline=False,
            input_filter='float',
            hint_text='Введите эталонный вес (кг)'
        )
        content.add_widget(Label(text='Введите эталонный вес:'))
        content.add_widget(weight_input)
        
        def calibrate(instance):
            try:
                weight = float(weight_input.text)
                if self.current_protocol == "MIDL-MI-VDA":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "A&D":
                    command = f"C {weight}\r\n"
                elif self.current_protocol == "Sartorius":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "Ohaus":
                    command = f"CAL {weight}\r\n"
                else:
                    command = f"CAL {weight}\r\n"
                
                if self.serial and self.serial.is_open:
                    self.serial.write(command.encode())
                    self.log_message(f"Начата процедура калибровки с весом {weight} кг")
                popup.dismiss()
            except ValueError:
                self.show_popup("Ошибка", "Введите корректное значение веса")
        
        popup = Popup(
            title='Калибровка',
            content=content,
            size_hint=(None, None),
            size=(300, 200)
        )
        content.add_widget(Button(
            text='Калибровать',
            size_hint_y=None,
            height='44dp',
            on_press=calibrate
        ))
        popup.open()

    def save_log_to_file(self):
        try:
            filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.root.ids.log_text.text)
            self.log_message(f"Лог сохранен в файл: {filename}")
        except Exception as e:
            self.show_popup("Ошибка", f"Не удалось сохранить файл: {str(e)}")

    def try_auto_detect_protocol(self, data):
        if not data:
            return
            
        if data.startswith("W +") and "kg" in data:
            self.current_protocol = "MIDL-MI-VDA"
            self.root.ids.protocol_label.text = "Протокол: MIDL-MI-VDA"
            self.log_message("Автоопределен протокол: MIDL-MI-VDA")
        elif data.startswith("+") and len(data) >= 7 and data[1:].replace(".", "").isdigit():
            self.current_protocol = "A&D"
            self.root.ids.protocol_label.text = "Протокол: A&D"
            self.log_message("Автоопределен протокол: A&D")
        elif data.replace(".", "").isdigit() and len(data) >= 5:
            self.current_protocol = "Sartorius"
            self.root.ids.protocol_label.text = "Протокол: Sartorius"
            self.log_message("Автоопределен протокол: Sartorius")
        elif data.startswith("ST,") and "," in data:
            self.current_protocol = "Ohaus"
            self.root.ids.protocol_label.text = "Протокол: Ohaus"
            self.log_message("Автоопределен протокол: Ohaus")
        else:
            self.log_message("Не удалось определить протокол автоматически")

    def show_popup(self, title, content):
        popup = Popup(
            title=title,
            content=Label(text=content),
            size_hint=(None, None),
            size=(300, 200)
        )
        popup.open()

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.root.ids.log_text.text += f"[{timestamp}] {message}\n"
        # Прокрутка к последнему сообщению
        self.root.ids.log_text.cursor = (0, len(self.root.ids.log_text.text))

    def on_stop(self):
        if self.serial and self.serial.is_open:
            self.serial.close()
        if self.update_event:
            self.update_event.cancel()

if __name__ == '__main__':
    WeightScaleApp().run() 