import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import time
from datetime import datetime

class WeighingScaleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Весовой прибор МИДЛ МИ ВДА/12я")
        self.root.geometry("800x600")
        
        # Переменные
        self.port_var = tk.StringVar()
        self.weight_var = tk.StringVar(value="0.000")
        self.status_var = tk.StringVar(value="Не подключено")
        self.gross_net_var = tk.StringVar(value="Брутто")
        
        self.serial_port = None
        self.is_connected = False
        self.decimal_places = 3  # По умолчанию 3 знака после запятой
        
        self.create_widgets()
        
    def create_widgets(self):
        # Фрейм для подключения
        connection_frame = ttk.LabelFrame(self.root, text="Подключение", padding="5")
        connection_frame.pack(fill="x", padx=5, pady=5)
        
        # Выпадающий список для выбора порта
        ttk.Label(connection_frame, text="Порт:").pack(side="left", padx=5)
        self.port_combo = ttk.Combobox(connection_frame, textvariable=self.port_var)
        self.port_combo.pack(side="left", padx=5)
        
        # Кнопка обновления списка портов
        ttk.Button(connection_frame, text="Обновить", command=self.update_ports).pack(side="left", padx=5)
        
        # Кнопка подключения
        self.connect_button = ttk.Button(connection_frame, text="Подключиться", command=self.toggle_connection)
        self.connect_button.pack(side="left", padx=5)
        
        # Статус подключения
        ttk.Label(connection_frame, textvariable=self.status_var).pack(side="left", padx=5)
        
        # Фрейм для отображения веса
        weight_frame = ttk.LabelFrame(self.root, text="Вес", padding="5")
        weight_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Отображение режима (Брутто/Нетто)
        ttk.Label(weight_frame, textvariable=self.gross_net_var, font=("Arial", 24)).pack()
        
        # Отображение веса
        weight_label = ttk.Label(weight_frame, textvariable=self.weight_var, font=("Arial", 48))
        weight_label.pack(expand=True)
        
        # Кнопки управления
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(control_frame, text="Тара", command=self.tare).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Ноль", command=self.zero).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Записать", command=self.save_weight).pack(side="left", padx=5)
        
        # Обновляем список портов при запуске
        self.update_ports()
        
    def update_ports(self):
        try:
            ports = [port.device for port in serial.tools.list_ports.comports()]
            if not ports:
                messagebox.showwarning("Предупреждение", "Не найдено доступных COM-портов!")
                self.status_var.set("Нет доступных портов")
                return
                
            self.port_combo['values'] = ports
            self.port_combo.set(ports[0])
            self.status_var.set(f"Доступно портов: {len(ports)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении списка портов: {str(e)}")
            self.status_var.set("Ошибка обновления портов")
            
    def toggle_connection(self):
        if not self.is_connected:
            try:
                port = self.port_var.get()
                if not port:
                    messagebox.showwarning("Предупреждение", "Выберите COM-порт!")
                    return
                    
                self.serial_port = serial.Serial(
                    port=port,
                    baudrate=9600,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=1
                )
                
                # Проверяем подключение
                if self.serial_port.is_open:
                    self.is_connected = True
                    self.connect_button.config(text="Отключиться")
                    self.status_var.set("Подключено")
                    self.start_weight_update()
                else:
                    raise Exception("Не удалось открыть порт")
                    
            except serial.SerialException as e:
                messagebox.showerror("Ошибка", f"Ошибка подключения к порту {port}: {str(e)}")
                self.status_var.set(f"Ошибка подключения: {str(e)}")
                if self.serial_port and self.serial_port.is_open:
                    self.serial_port.close()
                self.is_connected = False
                self.connect_button.config(text="Подключиться")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Неизвестная ошибка: {str(e)}")
                self.status_var.set(f"Ошибка: {str(e)}")
        else:
            self.disconnect()
            
    def disconnect(self):
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.close()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при отключении: {str(e)}")
        self.is_connected = False
        self.connect_button.config(text="Подключиться")
        self.status_var.set("Не подключено")
        self.weight_var.set("0.000")
        
    def send_command(self, command):
        if self.is_connected and self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.write(bytes([command]))
                time.sleep(0.1)  # Небольшая задержка для обработки команды
            except Exception as e:
                self.status_var.set(f"Ошибка отправки команды: {str(e)}")
                messagebox.showerror("Ошибка", f"Ошибка отправки команды: {str(e)}")
                
    def read_weight(self):
        if self.is_connected and self.serial_port and self.serial_port.is_open:
            try:
                # Отправляем команду получения веса (0ah)
                self.send_command(0x0A)
                
                # Читаем ответ (20 байт)
                response = self.serial_port.read(20)
                
                if len(response) == 20:
                    # Первые 6 байт - вес
                    weight_bytes = response[0:6]
                    # Преобразуем байты в число
                    weight = 0
                    for i, byte in enumerate(weight_bytes):
                        weight += byte * (10 ** (5-i))
                    
                    # Преобразуем в килограммы с учетом десятичных знаков
                    weight = weight / (10 ** self.decimal_places)
                    
                    # Проверяем статус (7-й байт)
                    status = response[6]
                    # Проверяем режим брутто/нетто (бит D0)
                    is_net = bool(status & 0x01)
                    self.gross_net_var.set("Нетто" if is_net else "Брутто")
                    
                    # Проверяем знак (бит D1)
                    is_negative = bool(status & 0x02)
                    if is_negative:
                        weight = -weight
                        
                    # Проверяем перегрузку (бит D2)
                    is_overload = bool(status & 0x04)
                    if is_overload:
                        self.status_var.set("Перегрузка!")
                    else:
                        self.status_var.set("Подключено")
                    
                    # Форматируем и отображаем вес
                    self.weight_var.set(f"{weight:.{self.decimal_places}f}")
                else:
                    self.status_var.set(f"Неверный ответ от весов: {len(response)} байт")
                    
            except Exception as e:
                self.status_var.set(f"Ошибка чтения: {str(e)}")
                messagebox.showerror("Ошибка", f"Ошибка чтения данных: {str(e)}")
                
    def start_weight_update(self):
        if self.is_connected:
            try:
                self.read_weight()
            except Exception as e:
                self.status_var.set(f"Ошибка чтения: {str(e)}")
            finally:
                self.root.after(1000, self.start_weight_update)
                
    def tare(self):
        if self.is_connected:
            try:
                self.send_command(0x0C)  # Команда тарирования
                self.status_var.set("Тарирование выполнено")
            except Exception as e:
                self.status_var.set(f"Ошибка тарирования: {str(e)}")
                messagebox.showerror("Ошибка", f"Ошибка тарирования: {str(e)}")
                
    def zero(self):
        if self.is_connected:
            try:
                self.send_command(0x0D)  # Команда установки нуля
                self.status_var.set("Установлен ноль")
            except Exception as e:
                self.status_var.set(f"Ошибка установки нуля: {str(e)}")
                messagebox.showerror("Ошибка", f"Ошибка установки нуля: {str(e)}")
                
    def save_weight(self):
        if self.is_connected:
            try:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                weight = self.weight_var.get()
                mode = self.gross_net_var.get()
                with open("weights.txt", "a", encoding='utf-8') as f:
                    f.write(f"{timestamp} - {weight} кг ({mode})\n")
                self.status_var.set("Вес записан")
            except Exception as e:
                self.status_var.set(f"Ошибка записи: {str(e)}")
                messagebox.showerror("Ошибка", f"Ошибка записи в файл: {str(e)}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = WeighingScaleApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Программа не может быть запущена: {str(e)}")
