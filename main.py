from kivy.app import App
from kivy.uix.tabbedpanel import TabbedPanel
from kivy.properties import (StringProperty, NumericProperty, 
                           ListProperty, BooleanProperty)
from kivy.core.window import Window
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from serial import Serial, SerialException
from threading import Thread, Lock
from queue import Queue
from datetime import datetime
import os
import time

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
        self.log_lock = Lock()
        self.log_file = None
        self.init_log_file()
        self.protocol_info = {
            "MIDL-MI-VDA": {"zero_cmd": "Z\r\n", "cal_cmd": "CAL {}\r\n"},
            "ТОКВЕС SH-50": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Микросим М0601": {"zero_cmd": "T\r\n", "cal_cmd": "CAL {}\r\n"},
            "Ньютон 42": {"zero_cmd": "Z\r\n", "cal_cmd": "C {}\r\n"}
        }
        Clock.schedule_interval(self.process_queue, 0.1)

    def init_log_file(self):
        # Создаем директорию для логов, если её нет
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Создаем имя файла с текущей датой и временем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(log_dir, f"weight_scale_{timestamp}.log")
        
        # Записываем заголовок в лог
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write(f"=== Лог весового прибора ===\n")
            f.write(f"Начало записи: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 30 + "\n\n")

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
            if self.serial and self.serial.is_open:
                self.log_message("Закрытие существующего соединения")
                self.serial.close()
                self.serial = None
                time.sleep(0.5)  # Даем время на освобождение порта

            port = self.ids.port_spinner.text
            baudrate = int(self.ids.baud_spinner.text)
            bytesize = int(self.ids.databits_spinner.text)
            parity = self.ids.parity_spinner.text[0]
            stopbits = float(self.ids.stopbits_spinner.text)
            
            self.log_message(f"Попытка подключения к {port}")
            self.log_message(f"Параметры: {baudrate} бод, {bytesize} бит, {parity}, {stopbits} стоп-бит")
            
            # Проверяем, доступен ли порт
            try:
                test_serial = Serial(port)
                test_serial.close()
                time.sleep(0.5)  # Даем время на освобождение порта
            except Exception as e:
                self.log_message(f"Порт {port} недоступен: {str(e)}")
                return
            
            self.serial = Serial(
                port=port,
                baudrate=baudrate,
                bytesize=bytesize,
                parity=parity,
                stopbits=stopbits,
                timeout=1.0,
                write_timeout=1.0,
                exclusive=True
            )
            
            # Очистка буфера
            self.serial.reset_input_buffer()
            self.serial.reset_output_buffer()
            
            # Проверяем, что порт действительно открыт
            if not self.serial.is_open:
                raise Exception("Не удалось открыть порт")
            
            self.is_connected = True
            self.status = f"Подключено к {port}"
            self.log_message(f"Подключено к {port}")
            
            # Запускаем чтение в отдельном потоке
            self.read_thread = Thread(target=self.read_serial, daemon=True)
            self.read_thread.start()
            
            # Отправляем только один запрос веса при подключении
            if self.ids.protocol_spinner.text == "Ньютон 42":
                try:
                    self.serial.write(b'\x80P\r\n')
                    self.log_message("Отправлен запрос веса (двоичный формат)")
                except Exception as e:
                    self.log_message(f"Ошибка отправки запроса веса: {str(e)}")
                
        except Exception as e:
            self.log_message(f"Ошибка подключения: {str(e)}")
            self.is_connected = False
            self.status = "Ошибка подключения"
            if self.serial:
                try:
                    self.serial.close()
                except:
                    pass
                self.serial = None

    def disconnect(self):
        self.log_message("Начало отключения")
        self.is_connected = False
        if self.serial:
            try:
                self.serial.close()
                self.log_message("Порт закрыт")
            except Exception as e:
                self.log_message(f"Ошибка при закрытии порта: {str(e)}")
            finally:
                self.serial = None
        self.status = "Не подключено"
        self.log_message("Отключено от весового прибора")

    def read_serial(self):
        last_request_time = 0
        request_interval = 0.5  # Интервал между запросами в секундах
        
        while self.is_connected and self.serial and self.serial.is_open:
            try:
                current_time = time.time()
                
                if self.serial.in_waiting:
                    # Читаем все доступные данные
                    data = self.serial.read(self.serial.in_waiting)
                    self.log_message(f"Получены сырые данные (hex): {data.hex()}")
                    
                    # Для Ньютон 42 обрабатываем двоичные данные
                    if self.ids.protocol_spinner.text == "Ньютон 42":
                        try:
                            # Проверяем, что это двоичные данные (должны начинаться с бита Head = 1)
                            if len(data) >= 4 and (data[0] & 0x80) == 0x80:  # Проверяем бит Head
                                # Извлекаем информацию из заголовочного байта
                                header = data[0]
                                stable = (header & 0x40) != 0  # Бит 6 - стабильность
                                overload = (header & 0x20) != 0  # Бит 5 - перегрузка
                                is_zero_proc = (header & 0x10) != 0  # Бит 4 - обнуление
                                dpoints = (header & 0x0C) >> 2  # Бит 3-2 - десятичные знаки
                                channels = (header & 0x03) + 1  # Бит 1-0 - количество каналов
                                
                                self.log_message(f"Заголовок: стабильность={stable}, перегрузка={overload}, обнуление={is_zero_proc}, десятичных знаков={dpoints}, каналов={channels}")
                                
                                # Проверяем размер данных
                                expected_size = 1 + 3 * channels  # Заголовок + 3 байта на канал
                                if len(data) >= expected_size:
                                    # Обрабатываем данные каждого канала
                                    for channel in range(channels):
                                        channel_data = data[1 + channel*3:4 + channel*3]
                                        weight = self.decode_newton42_weight(channel_data, dpoints)
                                        if weight is not None:
                                            self.current_weight = f"{weight:.3f} кг"
                                            if stable:
                                                self.status = "Стабильно"
                                            elif overload:
                                                self.status = "Перегрузка"
                                            elif is_zero_proc:
                                                self.status = "Обнуление"
                                            else:
                                                self.status = "Нестабильно"
                                            self.check_target_weight(weight)
                                            self.log_message(f"Канал {channel+1}, вес: {weight:.3f} кг")
                        except Exception as e:
                            self.log_message(f"Ошибка обработки двоичных данных: {str(e)}")
                    else:
                        # Для других протоколов используем ASCII
                        try:
                            lines = data.decode('ascii', errors='ignore').split('\r\n')
                            for line in lines:
                                line = line.strip()
                                if line:
                                    self.log_message(f"Обработка строки: {line}")
                                    self.data_queue.put(line)
                        except Exception as e:
                            self.log_message(f"Ошибка декодирования ASCII: {str(e)}")
                else:
                    # Если нет данных и прошло достаточно времени с последнего запроса
                    if current_time - last_request_time >= request_interval:
                        if self.ids.protocol_spinner.text == "Ньютон 42":
                            try:
                                self.serial.write(b'\x80P\r\n')
                                self.log_message("Отправлен запрос веса (двоичный формат)")
                                last_request_time = current_time
                            except Exception as e:
                                self.log_message(f"Ошибка отправки запроса веса: {str(e)}")
                        time.sleep(0.1)
            except Exception as e:
                self.log_message(f"Ошибка чтения: {str(e)}")
                self.disconnect()  # Используем метод disconnect вместо прямого закрытия
                break

    def decode_newton42_weight(self, weight_bytes, dpoints):
        try:
            if len(weight_bytes) != 3:
                self.log_message(f"Неверный размер данных веса: {len(weight_bytes)} байт")
                return None
                
            # Проверяем, что старшие биты всех байт равны 0
            if any(b & 0x80 for b in weight_bytes):
                self.log_message("Ошибка: старшие биты данных не равны 0")
                return None
                
            # Проверяем знак (бит 6 старшего байта)
            is_negative = (weight_bytes[2] & 0x40) != 0
            
            # Очищаем знаковый бит
            weight_bytes = bytearray(weight_bytes)
            weight_bytes[2] &= 0x3F
            
            # Собираем значение
            value = 0
            for i in range(3):
                value += weight_bytes[i] << (7 * i)
            
            # Применяем знак
            if is_negative:
                value = -value
                
            # Применяем десятичную точку
            weight = value / (10 ** dpoints)
            
            self.log_message(f"Декодирование веса: байты={weight_bytes.hex()}, знак={is_negative}, значение={value}, десятичных знаков={dpoints}, вес={weight}")
            return weight
        except Exception as e:
            self.log_message(f"Ошибка декодирования веса: {str(e)}")
            return None

    def process_weight_data(self, data):
        try:
            protocol = self.ids.protocol_spinner.text
            weight = None
            
            self.log_message(f"Обработка данных для протокола {protocol}: {data}")
            
            if protocol == "MIDL-MI-VDA" and data.startswith("W"):
                weight = float(data.split()[1])
            elif protocol == "ТОКВЕС SH-50" and data.startswith("ST,GS,"):
                weight = float(data.split(',')[2].strip().split()[0])
            elif protocol == "Микросим М0601" and data.startswith(('+', '-')):
                weight = float(data.split()[0])
            elif protocol == "Ньютон 42":
                # Проверяем различные форматы данных Ньютон 42
                if data.startswith('N'):
                    try:
                        # Формат N+00012.345 kg
                        weight_str = data[1:].split()[0]  # Убираем 'N' и берем первое число
                        weight = float(weight_str)
                        self.log_message(f"Распознан вес Ньютон 42 (формат N): {weight}")
                    except Exception as e:
                        self.log_message(f"Ошибка парсинга веса Ньютон 42 (формат N): {str(e)}")
                elif data.startswith('+') or data.startswith('-'):
                    try:
                        # Альтернативный формат +00012.345
                        weight = float(data.split()[0])
                        self.log_message(f"Распознан вес (альтернативный формат): {weight}")
                    except Exception as e:
                        self.log_message(f"Ошибка парсинга альтернативного формата: {str(e)}")
                elif data.startswith('P'):
                    # Игнорируем эхо команды P
                    self.log_message("Получено эхо команды P")
                else:
                    self.log_message(f"Неизвестный формат данных Ньютон 42: {data}")
            
            if weight is not None:
                self.current_weight = f"{weight:.3f} кг"
                self.check_target_weight(weight)
                self.log_message(f"Вес: {weight:.3f} кг")
        except Exception as e:
            self.log_message(f"Ошибка обработки: {str(e)}")

    def send_zero_command(self):
        if self.is_connected:
            protocol = self.ids.protocol_spinner.text
            if protocol == "Ньютон 42":
                cmd = b'\x80Z\r\n'  # Заголовок (Head=1) + Z + CR + LF
            else:
                cmd = self.protocol_info.get(protocol, {}).get("zero_cmd", "Z\r\n").encode()
            self.serial.write(cmd)
            self.log_message("Команда тары отправлена")

    def start_calibration(self):
        content = BoxLayout(orientation='vertical', spacing=10)
        weight_input = TextInput(hint_text='Эталонный вес (кг)', size_hint_y=None, height=40)
        
        def calibrate(_):
            try:
                weight = float(weight_input.text)
                protocol = self.ids.protocol_spinner.text
                if protocol == "Ньютон 42":
                    # Преобразуем вес в двоичный формат
                    weight_int = int(weight * 1000)  # Переводим в граммы
                    weight_bytes = weight_int.to_bytes(3, byteorder='little')
                    cmd = b'\x80C' + weight_bytes + b'\r\n'  # Заголовок (Head=1) + C + вес + CR + LF
                else:
                    cmd = self.protocol_info.get(protocol, {}).get("cal_cmd", "CAL {}\r\n").format(weight).encode()
                self.serial.write(cmd)
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
        log_entry = f"[{timestamp}] {message}\n"
        
        with self.log_lock:
            # Обновляем GUI
            Clock.schedule_once(lambda dt: self._update_log(log_entry))
            
            # Записываем в файл
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(log_entry)
            except Exception as e:
                print(f"Ошибка записи в лог-файл: {str(e)}")

    def _update_log(self, message):
        self.log_text += message

    def save_log_to_file(self):
        try:
            # Создаем имя файла с текущей датой и временем
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"weight_scale_export_{timestamp}.log"
            
            # Сохраняем текущий лог в новый файл
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.log_text)
            
            self.log_message(f"Лог сохранен в файл: {filename}")
        except Exception as e:
            self.log_message(f"Ошибка сохранения лога: {str(e)}")

    def on_stop(self):
        self.log_message("Завершение работы программы...")
        self.disconnect()  # Используем метод disconnect
        self.log_message("Программа завершена")

    def exit_app(self):
        self.log_message("Завершение работы программы...")
        if self.serial and self.serial.is_open:
            self.serial.close()
        App.get_running_app().stop()

    def refresh_ports(self):
        try:
            ports = []
            for i in range(1, 17):
                port = f"COM{i}"
                try:
                    s = Serial(port)
                    s.close()
                    ports.append(port)
                except:
                    pass
            self.ids.port_spinner.values = ports or ["Не найдены"]
            self.log_message(f"Найдены порты: {', '.join(ports) if ports else 'нет'}")
        except Exception as e:
            self.log_message(f"Ошибка обновления списка портов: {str(e)}")

class MainApp(App):
    def build(self):
        Window.clearcolor = (0.95, 0.95, 0.95, 1)
        return WeightScaleApp()

if __name__ == '__main__':
    MainApp().run()