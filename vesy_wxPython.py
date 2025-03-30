import sys
import csv
from datetime import datetime
import wx
import wx.adv
import serial
import serial.tools.list_ports
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import wx.lib.agw.aui as aui

class WeightScaleApp(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Программа для весового прибора МИДЛ МИ ВДА/12Я", size=(1000, 700))
        
        self.serial_port = None
        self.weight_history = []
        self.max_history_points = 100
        self.current_unit = 'kg'
        self.units = {'kg': 1.0, 'g': 1000.0, 'lb': 2.20462}
        self.protocols = [
            "Auto", 
            "MIDL-MI-VDA", 
            "A&D", 
            "Sartorius", 
            "Ohaus", 
            "ТОКВЕС SH-50",
            "Микросим М0601",
            "Ньютон 42"
        ]
        self.current_protocol = None
        self.target_weight = None
        self.timer = None
        
        self.init_ui()
        self.init_serial_settings()
        self.init_chart()
        
    def init_ui(self):
        # Create a notebook (tab control)
        self.notebook = aui.AuiNotebook(self)
        
        # Create tabs
        self.main_panel = self.create_main_tab()
        self.settings_panel = self.create_settings_tab()
        self.chart_panel = self.create_chart_tab()
        self.protocol_panel = self.create_protocol_tab()
        self.appearance_panel = self.create_appearance_tab()
        
        # Add tabs to notebook
        self.notebook.AddPage(self.main_panel, "Основное")
        self.notebook.AddPage(self.settings_panel, "Настройки")
        self.notebook.AddPage(self.chart_panel, "График")
        self.notebook.AddPage(self.protocol_panel, "Протокол")
        self.notebook.AddPage(self.appearance_panel, "Внешний вид")
        
        # Create menu
        self.init_menu()
        
        # Set up the main sizer
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.notebook, 1, wx.EXPAND)
        self.SetSizer(sizer)
        
    def init_menu(self):
        menubar = wx.MenuBar()
        
        # File menu
        file_menu = wx.Menu()
        export_item = file_menu.Append(wx.ID_ANY, "Экспорт данных")
        exit_item = file_menu.Append(wx.ID_EXIT, "Выход")
        menubar.Append(file_menu, "Файл")
        
        # Settings menu
        settings_menu = wx.Menu()
        theme_menu = wx.Menu()
        
        themes = ["Default", "Classic", "Modern"]
        for theme in themes:
            item = theme_menu.Append(wx.ID_ANY, theme)
            self.Bind(wx.EVT_MENU, lambda evt, t=theme: self.set_theme(t), item)
        
        settings_menu.AppendSubMenu(theme_menu, "Тема")
        menubar.Append(settings_menu, "Настройки")
        
        self.SetMenuBar(menubar)
        
        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_export_data, export_item)
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)
    
    def create_main_tab(self):
        panel = wx.Panel(self.notebook)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Status group
        status_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Статус подключения")
        
        self.status_label = wx.StaticText(panel, label="Статус: Не подключено")
        self.protocol_label = wx.StaticText(panel, label="Протокол: Не определен")
        self.port_label = wx.StaticText(panel, label="Порт: Не выбран")
        self.settings_label = wx.StaticText(panel, label="Параметры: Не установлены")
        
        status_group.Add(self.status_label, 0, wx.ALL, 5)
        status_group.Add(self.protocol_label, 0, wx.ALL, 5)
        status_group.Add(self.port_label, 0, wx.ALL, 5)
        status_group.Add(self.settings_label, 0, wx.ALL, 5)
        
        # Weight group
        weight_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Измерения")
        
        self.weight_label = wx.StaticText(panel, label="Вес: ---")
        font = self.weight_label.GetFont()
        font.SetPointSize(24)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        self.weight_label.SetFont(font)
        
        self.unit_combo = wx.ComboBox(panel, choices=list(self.units.keys()), style=wx.CB_READONLY)
        self.unit_combo.Bind(wx.EVT_COMBOBOX, self.on_change_unit)
        
        # Target weight controls
        target_sizer = wx.BoxSizer(wx.HORIZONTAL)
        target_sizer.Add(wx.StaticText(panel, label="Целевой вес:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.target_weight_edit = wx.TextCtrl(panel)
        self.set_target_button = wx.Button(panel, label="Установить")
        self.clear_target_button = wx.Button(panel, label="Сбросить")
        
        target_sizer.Add(self.target_weight_edit, 1, wx.ALL, 5)
        target_sizer.Add(self.set_target_button, 0, wx.ALL, 5)
        target_sizer.Add(self.clear_target_button, 0, wx.ALL, 5)
        
        self.set_target_button.Bind(wx.EVT_BUTTON, self.on_set_target_weight)
        self.clear_target_button.Bind(wx.EVT_BUTTON, self.on_clear_target_weight)
        
        weight_group.Add(self.weight_label, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 10)
        weight_group.Add(self.unit_combo, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)
        weight_group.Add(target_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        # Control group
        control_group = wx.StaticBoxSizer(wx.HORIZONTAL, panel, "Управление")
        
        self.connect_button = wx.Button(panel, label="Подключить")
        self.zero_button = wx.Button(panel, label="Тара")
        self.calibrate_button = wx.Button(panel, label="Калибровка")
        
        self.connect_button.Bind(wx.EVT_BUTTON, self.on_toggle_connection)
        self.zero_button.Bind(wx.EVT_BUTTON, self.on_send_zero_command)
        self.calibrate_button.Bind(wx.EVT_BUTTON, self.on_start_calibration)
        
        self.zero_button.Disable()
        self.calibrate_button.Disable()
        
        control_group.Add(self.connect_button, 0, wx.ALL, 5)
        control_group.Add(self.zero_button, 0, wx.ALL, 5)
        control_group.Add(self.calibrate_button, 0, wx.ALL, 5)
        
        # Log group
        log_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Лог")
        
        self.log_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)
        self.save_log_button = wx.Button(panel, label="Сохранить лог в файл")
        self.save_log_button.Bind(wx.EVT_BUTTON, self.on_save_log)
        
        log_group.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        log_group.Add(self.save_log_button, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        
        # Add all groups to main sizer
        main_sizer.Add(status_group, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(weight_group, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(control_group, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(log_group, 1, wx.EXPAND | wx.ALL, 5)
        
        panel.SetSizer(main_sizer)
        return panel
    
    def create_settings_tab(self):
        panel = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Port settings
        port_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки порта")
        
        # Port selection
        port_sizer = wx.BoxSizer(wx.HORIZONTAL)
        port_sizer.Add(wx.StaticText(panel, label="COM-порт:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.port_combo = wx.ComboBox(panel, style=wx.CB_READONLY)
        self.refresh_ports()
        
        self.refresh_button = wx.Button(panel, label="Обновить список портов")
        self.refresh_button.Bind(wx.EVT_BUTTON, self.on_refresh_ports)
        
        port_sizer.Add(self.port_combo, 1, wx.ALL, 5)
        port_sizer.Add(self.refresh_button, 0, wx.ALL, 5)
        
        # Baud rate
        baud_sizer = wx.BoxSizer(wx.HORIZONTAL)
        baud_sizer.Add(wx.StaticText(panel, label="Скорость (бод):"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.baud_combo = wx.ComboBox(panel, choices=["9600", "4800", "2400", "1200", "19200", "38400", "57600", "115200"], 
                                    style=wx.CB_READONLY)
        self.baud_combo.SetValue("9600")
        
        baud_sizer.Add(self.baud_combo, 1, wx.ALL, 5)
        
        # Data bits
        data_bits_sizer = wx.BoxSizer(wx.HORIZONTAL)
        data_bits_sizer.Add(wx.StaticText(panel, label="Биты данных:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.data_bits_combo = wx.ComboBox(panel, choices=["5", "6", "7", "8"], style=wx.CB_READONLY)
        self.data_bits_combo.SetValue("8")
        
        data_bits_sizer.Add(self.data_bits_combo, 1, wx.ALL, 5)
        
        # Parity
        parity_sizer = wx.BoxSizer(wx.HORIZONTAL)
        parity_sizer.Add(wx.StaticText(panel, label="Четность:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.parity_combo = wx.ComboBox(panel, choices=["NoParity", "EvenParity", "OddParity", "SpaceParity", "MarkParity"], 
                                      style=wx.CB_READONLY)
        parity_sizer.Add(self.parity_combo, 1, wx.ALL, 5)
        
        # Stop bits
        stop_bits_sizer = wx.BoxSizer(wx.HORIZONTAL)
        stop_bits_sizer.Add(wx.StaticText(panel, label="Стоп-биты:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.stop_bits_combo = wx.ComboBox(panel, choices=["1", "1.5", "2"], style=wx.CB_READONLY)
        stop_bits_sizer.Add(self.stop_bits_combo, 1, wx.ALL, 5)
        
        # Flow control
        flow_control_sizer = wx.BoxSizer(wx.HORIZONTAL)
        flow_control_sizer.Add(wx.StaticText(panel, label="Управление потоком:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.flow_control_combo = wx.ComboBox(panel, choices=["NoFlowControl", "HardwareControl", "SoftwareControl"], 
                                            style=wx.CB_READONLY)
        flow_control_sizer.Add(self.flow_control_combo, 1, wx.ALL, 5)
        
        # Add all port settings to group
        port_group.Add(port_sizer, 0, wx.EXPAND | wx.ALL, 5)
        port_group.Add(baud_sizer, 0, wx.EXPAND | wx.ALL, 5)
        port_group.Add(data_bits_sizer, 0, wx.EXPAND | wx.ALL, 5)
        port_group.Add(parity_sizer, 0, wx.EXPAND | wx.ALL, 5)
        port_group.Add(stop_bits_sizer, 0, wx.EXPAND | wx.ALL, 5)
        port_group.Add(flow_control_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        # Chart settings
        chart_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки графика")
        
        points_sizer = wx.BoxSizer(wx.HORIZONTAL)
        points_sizer.Add(wx.StaticText(panel, label="Количество точек на графике:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.history_points_spin = wx.SpinCtrl(panel, min=10, max=1000, initial=self.max_history_points)
        self.history_points_spin.Bind(wx.EVT_SPINCTRL, self.on_update_history_size)
        
        points_sizer.Add(self.history_points_spin, 1, wx.ALL, 5)
        chart_group.Add(points_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        # Sound settings
        sound_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки звука")
        
        self.sound_checkbox = wx.CheckBox(panel, label="Звуковое оповещение")
        self.sound_checkbox.SetValue(True)
        
        self.test_sound_button = wx.Button(panel, label="Проверить звук")
        self.test_sound_button.Bind(wx.EVT_BUTTON, self.on_play_sound)
        
        sound_group.Add(self.sound_checkbox, 0, wx.ALL, 5)
        sound_group.Add(self.test_sound_button, 0, wx.ALL, 5)
        
        # Add all groups to main sizer
        sizer.Add(port_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(chart_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(sound_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.AddStretchSpacer()
        
        panel.SetSizer(sizer)
        return panel
    
    def create_chart_tab(self):
        panel = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Chart canvas
        self.figure = Figure()
        self.canvas = FigureCanvas(panel, -1, self.figure)
        
        # Export buttons
        export_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.export_csv_button = wx.Button(panel, label="Экспорт в CSV")
        self.export_excel_button = wx.Button(panel, label="Экспорт в Excel")
        
        self.export_csv_button.Bind(wx.EVT_BUTTON, lambda e: self.on_export_data('csv'))
        self.export_excel_button.Bind(wx.EVT_BUTTON, lambda e: self.on_export_data('excel'))
        
        export_sizer.Add(self.export_csv_button, 0, wx.ALL, 5)
        export_sizer.Add(self.export_excel_button, 0, wx.ALL, 5)
        
        sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(export_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        
        panel.SetSizer(sizer)
        return panel
    
    def create_protocol_tab(self):
        panel = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Protocol settings
        protocol_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки протокола")
        
        protocol_sizer = wx.BoxSizer(wx.HORIZONTAL)
        protocol_sizer.Add(wx.StaticText(panel, label="Протокол весов:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.protocol_combo = wx.ComboBox(panel, choices=self.protocols, style=wx.CB_READONLY)
        self.protocol_combo.Bind(wx.EVT_COMBOBOX, self.on_change_protocol)
        
        self.detect_button = wx.Button(panel, label="Автоопределение")
        self.detect_button.Bind(wx.EVT_BUTTON, self.on_detect_protocol)
        
        protocol_sizer.Add(self.protocol_combo, 1, wx.ALL, 5)
        protocol_sizer.Add(self.detect_button, 0, wx.ALL, 5)
        
        protocol_group.Add(protocol_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        # Protocol info
        self.protocol_info = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        
        sizer.Add(protocol_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(wx.StaticText(panel, label="Описание протокола:"), 0, wx.ALL, 5)
        sizer.Add(self.protocol_info, 1, wx.EXPAND | wx.ALL, 5)
        
        panel.SetSizer(sizer)
        return panel
    
    def create_appearance_tab(self):
        panel = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Color settings
        color_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки цветов")
        
        self.bg_color_button = wx.Button(panel, label="Цвет фона")
        self.text_color_button = wx.Button(panel, label="Цвет текста")
        self.chart_color_button = wx.Button(panel, label="Цвет графика")
        self.reset_colors_button = wx.Button(panel, label="Сбросить цвета")
        
        self.bg_color_button.Bind(wx.EVT_BUTTON, lambda e: self.on_change_color('background'))
        self.text_color_button.Bind(wx.EVT_BUTTON, lambda e: self.on_change_color('text'))
        self.chart_color_button.Bind(wx.EVT_BUTTON, lambda e: self.on_change_color('chart'))
        self.reset_colors_button.Bind(wx.EVT_BUTTON, self.on_reset_colors)
        
        color_group.Add(self.bg_color_button, 0, wx.ALL, 5)
        color_group.Add(self.text_color_button, 0, wx.ALL, 5)
        color_group.Add(self.chart_color_button, 0, wx.ALL, 5)
        color_group.Add(self.reset_colors_button, 0, wx.ALL, 5)
        
        # Font settings
        font_group = wx.StaticBoxSizer(wx.VERTICAL, panel, "Настройки шрифта")
        
        font_size_sizer = wx.BoxSizer(wx.HORIZONTAL)
        font_size_sizer.Add(wx.StaticText(panel, label="Размер шрифта:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.font_size_spin = wx.SpinCtrl(panel, min=8, max=24, initial=12)
        self.font_size_spin.Bind(wx.EVT_SPINCTRL, self.on_change_font_size)
        
        font_size_sizer.Add(self.font_size_spin, 1, wx.ALL, 5)
        
        font_family_sizer = wx.BoxSizer(wx.HORIZONTAL)
        font_family_sizer.Add(wx.StaticText(panel, label="Шрифт:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        
        self.font_family_combo = wx.ComboBox(panel, choices=["Arial", "Times New Roman", "Courier New", "Verdana"], 
                                           style=wx.CB_READONLY)
        self.font_family_combo.Bind(wx.EVT_COMBOBOX, self.on_change_font_family)
        
        font_family_sizer.Add(self.font_family_combo, 1, wx.ALL, 5)
        
        font_group.Add(font_size_sizer, 0, wx.EXPAND | wx.ALL, 5)
        font_group.Add(font_family_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        sizer.Add(color_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(font_group, 0, wx.EXPAND | wx.ALL, 5)
        sizer.AddStretchSpacer()
        
        panel.SetSizer(sizer)
        return panel
    
    def init_chart(self):
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Изменение веса во времени")
        self.ax.set_xlabel("Время (с)")
        self.ax.set_ylabel("Вес (kg)")
        self.line, = self.ax.plot([], [], 'b-')
        self.figure.tight_layout()
    
    def init_serial_settings(self):
        # Default settings
        self.serial_settings = {
            'baudrate': 9600,
            'bytesize': 8,
            'parity': 'N',
            'stopbits': 1,
            'timeout': 1
        }
    
    def refresh_ports(self):
        self.port_combo.Clear()
        ports = serial.tools.list_ports.comports()
        if ports:
            for port in ports:
                self.port_combo.Append(port.device)
        else:
            self.port_combo.Append("Порты не найдены")
    
    def on_refresh_ports(self, event):
        self.refresh_ports()
    
    def on_update_history_size(self, event):
        self.max_history_points = event.GetPosition()
        if len(self.weight_history) > self.max_history_points:
            self.weight_history = self.weight_history[-self.max_history_points:]
            self.update_chart()
    
    def on_toggle_connection(self, event):
        if self.serial_port and self.serial_port.is_open:
            # Disconnect
            self.serial_port.close()
            if self.timer:
                self.timer.Stop()
            self.connect_button.SetLabel("Подключить")
            self.status_label.SetLabel("Статус: Не подключено")
            self.zero_button.Disable()
            self.calibrate_button.Disable()
            self.log_message("Отключено от весового прибора")
        else:
            # Connect
            if self.port_combo.GetValue() == "Порты не найдены":
                wx.MessageBox("Нет доступных COM-портов!", "Ошибка", wx.OK | wx.ICON_ERROR)
                return
                
            port_name = self.port_combo.GetValue()
            
            try:
                # Apply settings from UI
                self.apply_serial_settings()
                
                self.serial_port = serial.Serial(
                    port=port_name,
                    baudrate=self.serial_settings['baudrate'],
                    bytesize=self.serial_settings['bytesize'],
                    parity=self.serial_settings['parity'],
                    stopbits=self.serial_settings['stopbits'],
                    timeout=self.serial_settings['timeout']
                )
                
                self.connect_button.SetLabel("Отключить")
                self.status_label.SetLabel(f"Статус: Подключено к {port_name}")
                self.port_label.SetLabel(f"Порт: {port_name}")
                self.update_settings_label()
                self.zero_button.Enable()
                self.calibrate_button.Enable()
                
                # Start timer for reading data
                self.timer = wx.Timer(self)
                self.Bind(wx.EVT_TIMER, self.on_read_data, self.timer)
                self.timer.Start(500)  # 500 ms interval
                
                self.log_message(f"Подключено к {port_name}")
                self.log_message(f"Параметры: {self.settings_label.GetLabel()}")
                
                # Try auto-detect protocol
                if self.protocol_combo.GetValue() == "Auto":
                    self.on_detect_protocol(None)
                    
            except Exception as e:
                wx.MessageBox(f"Не удалось открыть порт: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
                self.log_message(f"Ошибка подключения к {port_name}: {str(e)}")
    
    def apply_serial_settings(self):
        # Baud rate
        self.serial_settings['baudrate'] = int(self.baud_combo.GetValue())
        
        # Data bits
        data_bits = int(self.data_bits_combo.GetValue())
        if data_bits == 5:
            self.serial_settings['bytesize'] = serial.FIVEBITS
        elif data_bits == 6:
            self.serial_settings['bytesize'] = serial.SIXBITS
        elif data_bits == 7:
            self.serial_settings['bytesize'] = serial.SEVENBITS
        else:
            self.serial_settings['bytesize'] = serial.EIGHTBITS
        
        # Parity
        parity = self.parity_combo.GetValue()
        if parity == "NoParity":
            self.serial_settings['parity'] = serial.PARITY_NONE
        elif parity == "EvenParity":
            self.serial_settings['parity'] = serial.PARITY_EVEN
        elif parity == "OddParity":
            self.serial_settings['parity'] = serial.PARITY_ODD
        elif parity == "SpaceParity":
            self.serial_settings['parity'] = serial.PARITY_SPACE
        elif parity == "MarkParity":
            self.serial_settings['parity'] = serial.PARITY_MARK
        
        # Stop bits
        stop_bits = self.stop_bits_combo.GetValue()
        if stop_bits == "1":
            self.serial_settings['stopbits'] = serial.STOPBITS_ONE
        elif stop_bits == "1.5":
            self.serial_settings['stopbits'] = serial.STOPBITS_ONE_POINT_FIVE
        elif stop_bits == "2":
            self.serial_settings['stopbits'] = serial.STOPBITS_TWO
        
        # Flow control is not directly supported by pySerial in the same way
        # So we'll just store it for display purposes
        self.flow_control = self.flow_control_combo.GetValue()
    
    def update_settings_label(self):
        settings_text = (
            f"Параметры: {self.serial_settings['baudrate']} бод, "
            f"{self.data_bits_combo.GetValue()} бит, "
            f"{self.parity_combo.GetValue()}, "
            f"{self.stop_bits_combo.GetValue()} стоп-бит, "
            f"{self.flow_control_combo.GetValue()}"
        )
        self.settings_label.SetLabel(settings_text)
    
    def on_read_data(self, event):
        if self.serial_port and self.serial_port.in_waiting:
            try:
                data = self.serial_port.readline().decode().strip()
                self.process_weight_data(data)
            except Exception as e:
                self.log_message(f"Ошибка чтения данных: {str(e)}")
    
    def process_weight_data(self, data):
        try:
            # Process data based on protocol
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
                    
            elif self.current_protocol == "TOKVES-SH50":
                if data.startswith("ST,GS,"):
                    try:
                        parts = data.split(',')
                        weight_str = parts[2].strip().split()[0]
                        weight_kg = float(weight_str)
                        self.process_weight_value(weight_kg, "kg", data)
                    except (IndexError, ValueError) as e:
                        self.log_message(f"Ошибка разбора данных ТОКВЕС: {str(e)}")
                        
            elif self.current_protocol == "MIKROSIM-M0601":
                if data.startswith(('+', '-')) and 'kg' in data:
                    try:
                        weight_str = data.split()[0]
                        weight_kg = float(weight_str)
                        self.process_weight_value(weight_kg, "kg", data)
                    except (IndexError, ValueError) as e:
                        self.log_message(f"Ошибка разбора данных Микросим: {str(e)}")

            elif self.current_protocol == "NEWTON-42":
                if data.startswith('N') and 'kg' in data:
                    try:
                        weight_str = data[1:].split()[0]
                        weight_kg = float(weight_str)
                        self.process_weight_value(weight_kg, "kg", data)
                    except (IndexError, ValueError) as e:
                        self.log_message(f"Ошибка разбора данных Ньютон: {str(e)}")
            else:
                # Try auto-detect
                self.try_auto_detect_protocol(data)
        
        except Exception as e:
            self.log_message(f"Ошибка обработки данных: {str(e)}")
    
    def process_weight_value(self, weight_kg, unit, raw_data):
        # Convert to selected unit
        converted_weight = weight_kg * self.units[self.current_unit]
        
        self.weight_label.SetLabel(f"Вес: {converted_weight:.3f} {self.current_unit}")
        
        # Check for target weight
        if self.target_weight is not None and abs(weight_kg - self.target_weight) < 0.001:
            self.notify_target_weight_reached()
        
        # Add to history for chart
        timestamp = len(self.weight_history) * 0.5  # Assume 0.5 sec interval
        self.weight_history.append((timestamp, weight_kg))
        
        if len(self.weight_history) > self.max_history_points:
            self.weight_history.pop(0)
        
        self.update_chart()
        
        self.log_message(f"Получены данные: {raw_data}")
    
    def update_chart(self):
        if not self.weight_history:
            return
            
        times = [x[0] for x in self.weight_history]
        weights = [x[1] for x in self.weight_history]
        
        self.line.set_data(times, weights)
        self.ax.relim()
        self.ax.autoscale_view()
        
        # Add some padding to Y axis
        y_min, y_max = self.ax.get_ylim()
        y_pad = (y_max - y_min) * 0.1 if y_max != y_min else 1.0
        self.ax.set_ylim(max(0, y_min - y_pad), y_max + y_pad)
        
        self.canvas.draw()
    
    def on_send_zero_command(self, event):
        if not self.serial_port or not self.serial_port.is_open:
            return
            
        if self.current_protocol == "MIDL-MI-VDA":
            command = "Z\r\n"
        elif self.current_protocol == "A&D":
            command = "Z\r\n"
        elif self.current_protocol == "Sartorius":
            command = "T\r\n"
        elif self.current_protocol == "Ohaus":
            command = "Z\r\n"
        elif self.current_protocol == "TOKVES-SH50":
            command = "T\r\n"
        elif self.current_protocol == "MIKROSIM-M0601":
            command = "T\r\n"
        elif self.current_protocol == "NEWTON-42":
            command = "Z\r\n"
        else:
            command = "Z\r\n"  # Default command
        
        try:
            self.serial_port.write(command.encode())
            self.log_message(f"Отправлена команда тары: {command.strip()}")
        except Exception as e:
            self.log_message(f"Ошибка отправки команды тары: {str(e)}")
    
    def on_start_calibration(self, event):
        dlg = wx.NumberEntryDialog(
            self, 
            'Введите эталонный вес (кг):', 
            'Калибровка', 
            'Эталонный вес', 
            0, 0, 10000
        )
        
        if dlg.ShowModal() == wx.ID_OK:
            weight = dlg.GetValue()
            
            confirm = wx.MessageBox(
                f'Установите эталонный вес {weight} кг и нажмите OK', 
                'Калибровка', 
                wx.OK | wx.CANCEL
            )
            
            if confirm == wx.OK:
                if self.current_protocol == "MIDL-MI-VDA":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "A&D":
                    command = f"C {weight}\r\n"
                elif self.current_protocol == "Sartorius":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "Ohaus":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "TOKVES-SH50":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "MIKROSIM-M0601":
                    command = f"CAL {weight}\r\n"
                elif self.current_protocol == "NEWTON-42":
                    command = f"C {weight}\r\n"
                else:
                    command = f"CAL {weight}\r\n"  # Default command
                
                try:
                    self.serial_port.write(command.encode())
                    self.log_message(f"Начата процедура калибровки с весом {weight} кг")
                except Exception as e:
                    self.log_message(f"Ошибка отправки команды калибровки: {str(e)}")
        
        dlg.Destroy()
    
    def on_change_unit(self, event):
        self.current_unit = event.GetString()
        self.log_message(f"Изменена единица измерения на {self.current_unit}")
    
    def on_set_target_weight(self, event):
        try:
            weight = float(self.target_weight_edit.GetValue())
            self.target_weight = weight
            self.log_message(f"Установлен целевой вес: {weight} кг")
        except ValueError:
            wx.MessageBox("Введите корректное значение веса", "Ошибка", wx.OK | wx.ICON_ERROR)
    
    def on_clear_target_weight(self, event):
        self.target_weight = None
        self.target_weight_edit.Clear()
        self.log_message("Целевой вес сброшен")
    
    def notify_target_weight_reached(self):
        if self.sound_checkbox.GetValue():
            self.on_play_sound(None)
        
        wx.MessageBox(
            f"Достигнут целевой вес: {self.target_weight} кг", 
            "Целевой вес достигнут", 
            wx.OK | wx.ICON_INFORMATION
        )
        self.on_clear_target_weight(None)
    
    def on_play_sound(self, event):
        wx.Bell()  # Simple system beep - for more advanced sound you'd need additional libraries
    
    def on_save_log(self, event):
        with wx.FileDialog(
            self, "Сохранить лог", wildcard="Текстовые файлы (*.txt)|*.txt", 
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        ) as dlg:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            
            file_name = dlg.GetPath()
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.GetValue())
                self.log_message(f"Лог сохранен в файл: {file_name}")
            except Exception as e:
                wx.MessageBox(f"Не удалось сохранить файл: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
    
    def on_export_data(self, format=None, event=None):
        if not self.weight_history:
            wx.MessageBox("Нет данных для экспорта", "Ошибка", wx.OK | wx.ICON_WARNING)
            return
        
        if format is None:
            with wx.FileDialog(
                self, "Экспорт данных", wildcard="CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx", 
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
            ) as dlg:
                if dlg.ShowModal() == wx.ID_CANCEL:
                    return
                
                file_name = dlg.GetPath()
                selected_filter = dlg.GetFilterIndex()
                
                if selected_filter == 0:
                    format = 'csv'
                    if not file_name.endswith('.csv'):
                        file_name += '.csv'
                else:
                    format = 'excel'
                    if not file_name.endswith('.xlsx'):
                        file_name += '.xlsx'
        else:
            default_name = "weights_export." + format
            with wx.FileDialog(
                self, "Экспорт данных", defaultFile=default_name, 
                wildcard=f"{format.upper()} (*.{format})|*.{format}", 
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
            ) as dlg:
                if dlg.ShowModal() == wx.ID_CANCEL:
                    return
                
                file_name = dlg.GetPath()
        
        try:
            if format == 'csv':
                self.export_to_csv(file_name)
            elif format == 'excel':
                self.export_to_excel(file_name)
            
            self.log_message(f"Данные экспортированы в {file_name}")
        except Exception as e:
            wx.MessageBox(f"Не удалось экспортировать данные: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
    
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
            wx.MessageBox(
                "Для экспорта в Excel требуется модуль openpyxl. Установите его командой: pip install openpyxl", 
                "Ошибка", 
                wx.OK | wx.ICON_ERROR
            )
            return
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Данные весов"
        
        # Headers
        ws.append(["Время (с)", "Вес (kg)"])
        
        # Data
        for time, weight in self.weight_history:
            ws.append([time, weight])
        
        # Save
        wb.save(file_name)
    
    def on_change_protocol(self, event):
        protocol = event.GetString()
        if protocol == "Auto":
            self.current_protocol = None
            self.protocol_info.SetValue(
                "Режим автоопределения протокола. Программа будет пытаться автоматически определить формат данных."
            )
        elif protocol == "MIDL-MI-VDA":
            self.current_protocol = "MIDL-MI-VDA"
            self.protocol_info.SetValue(
                "Протокол МИДЛ МИ ВДА/12Я\n\n"
                "Формат данных:\n"
                "W +123.45 kg\n\n"
                "Команды:\n"
                "Z - тара\n"
                "CAL <вес> - калибровка с указанием эталонного веса"
            )
        elif protocol == "A&D":
            self.current_protocol = "A&D"
            self.protocol_info.SetValue(
                "Протокол весов A&D\n\n"
                "Формат данных:\n"
                "+00123.45\r\n\n"
                "Команды:\n"
                "Z - тара\n"
                "C <вес> - калибровка"
            )
        elif protocol == "Sartorius":
            self.current_protocol = "Sartorius"
            self.protocol_info.SetValue(
                "Протокол весов Sartorius\n\n"
                "Формат данных:\n"
                "00123.45\r\n\n"
                "Команды:\n"
                "T - тара\n"
                "CAL <вес> - калибровка"
            )
        elif protocol == "Ohaus":
            self.current_protocol = "Ohaus"
            self.protocol_info.SetValue(
                "Протокол весов Ohaus\n\n"
                "Формат данных:\n"
                "ST,+00123.45,kg\r\n\n"
                "Команды:\n"
                "Z - тара\n"
                "CAL <вес> - калибровка"
            )
        elif protocol == "ТОКВЕС SH-50":
            self.current_protocol = "TOKVES-SH50"
            self.protocol_info.SetValue(
                "Протокол весов ТОКВЕС SH-50\n\n"
                "Формат данных:\n"
                "ST,GS,   0.000 kg\r\n"
                "ST,GS,  +1.234 kg\r\n"
                "ST,GS, -12.345 kg\r\n\n"
                "Команды:\n"
                "T - тара\n"
                "CAL - калибровка (требуется ввести эталонный вес)"
            )
        elif protocol == "Микросим М0601":
            self.current_protocol = "MIKROSIM-M0601"
            self.protocol_info.SetValue(
                "Протокол весов Микросим М0601\n\n"
                "Формат данных:\n"
                "+0001.234 kg\r\n"
                "-0000.123 kg\r\n\n"
                "Команды:\n"
                "T - тара\n"
                "CAL - калибровка (требуется ввести эталонный вес)"
            )
        elif protocol == "Ньютон 42":
            self.current_protocol = "NEWTON-42"
            self.protocol_info.SetValue(
                "Протокол весов Ньютон 42\n\n"
                "Формат данных:\n"
                "N+00012.345 kg\r\n"
                "N-00001.234 kg\r\n\n"
                "Команды:\n"
                "Z - тара\n"
                "C - калибровка (требуется ввести эталонный вес)"
            )
        
        self.protocol_label.SetLabel(f"Протокол: {protocol}")
        self.log_message(f"Установлен протокол: {protocol}")
    
    def on_detect_protocol(self, event):
        if not self.serial_port or not self.serial_port.is_open:
            wx.MessageBox("Сначала подключитесь к весам", "Ошибка", wx.OK | wx.ICON_WARNING)
            return
        
        self.log_message("Попытка автоопределения протокола...")
        
        try:
            # Send weight request command
            self.serial_port.write("W\r\n".encode())
            
            # Wait for response (in a real app you'd want to implement proper timeout)
            data = self.serial_port.readline().decode().strip()
            self.try_auto_detect_protocol(data)
        except Exception as e:
            self.log_message(f"Ошибка автоопределения протокола: {str(e)}")
    
    def try_auto_detect_protocol(self, data):
        if not data:
            return
            
        if data.startswith("W +") and "kg" in data:
            self.protocol_combo.SetValue("MIDL-MI-VDA")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: MIDL-MI-VDA")
        elif data.startswith("+") and len(data) >= 7 and data[1:].replace(".", "").isdigit():
            self.protocol_combo.SetValue("A&D")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: A&D")
        elif data.replace(".", "").isdigit() and len(data) >= 5:
            self.protocol_combo.SetValue("Sartorius")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: Sartorius")
        elif data.startswith("ST,") and "," in data:
            self.protocol_combo.SetValue("Ohaus")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: Ohaus")
        elif data.startswith("ST,GS,") and ("kg" in data or "Kg" in data):
            self.protocol_combo.SetValue("ТОКВЕС SH-50")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: ТОКВЕС SH-50")
        elif data.startswith(('+', '-')) and 'kg' in data and len(data.split()[0]) == 8:
            self.protocol_combo.SetValue("Микросим М0601")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: Микросим М0601")
        elif data.startswith('N') and ('+000' in data or '-000') and 'kg' in data:
            self.protocol_combo.SetValue("Ньютон 42")
            self.on_change_protocol(wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.protocol_combo.GetId()))
            self.log_message("Автоопределен протокол: Ньютон 42")
        else:
            self.log_message("Не удалось определить протокол автоматически")
    
    def on_change_color(self, element):
        dlg = wx.ColourDialog(self)
        if dlg.ShowModal() == wx.ID_OK:
            color = dlg.GetColourData().GetColour()
            
            if element == 'background':
                self.SetBackgroundColour(color)
                self.Refresh()
                self.log_message(f"Изменен цвет фона на {color.GetAsString(wx.C2S_HTML_SYNTAX)}")
            elif element == 'text':
                self.SetForegroundColour(color)
                self.Refresh()
                self.log_message(f"Изменен цвет текста на {color.GetAsString(wx.C2S_HTML_SYNTAX)}")
            elif element == 'chart':
                self.line.set_color(color.GetAsString(wx.C2S_HTML_SYNTAX))
                self.canvas.draw()
                self.log_message(f"Изменен цвет графика на {color.GetAsString(wx.C2S_HTML_SYNTAX)}")
        
        dlg.Destroy()
    
    def on_reset_colors(self, event):
        self.SetBackgroundColour(wx.NullColour)
        self.SetForegroundColour(wx.NullColour)
        self.line.set_color('b')  # Blue
        self.canvas.draw()
        self.Refresh()
        self.log_message("Цвета сброшены к значениям по умолчанию")
    
    def on_change_font_size(self, event):
        font = self.GetFont()
        font.SetPointSize(event.GetPosition())
        self.SetFont(font)
        self.log_message(f"Изменен размер шрифта на {event.GetPosition()}")
    
    def on_change_font_family(self, event):
        font = self.GetFont()
        font.SetFaceName(event.GetString())
        self.SetFont(font)
        self.log_message(f"Изменен шрифт на {event.GetString()}")
    
    def set_theme(self, theme_name):
        # wxPython has limited built-in theme support
        # This is a simplified implementation
        if theme_name == "Modern":
            self.SetBackgroundColour(wx.WHITE)
            self.SetForegroundColour(wx.BLACK)
        elif theme_name == "Classic":
            self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))
            self.SetForegroundColour(wx.BLACK)
        else:  # Default
            self.SetBackgroundColour(wx.NullColour)
            self.SetForegroundColour(wx.NullColour)
        
        self.Refresh()
        self.log_message(f"Установлена тема: {theme_name}")
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.AppendText(f"[{timestamp}] {message}\n")
    
    def on_exit(self, event):
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.Close()
    
    def on_export_data_menu(self, event):
        self.on_export_data(None)

if __name__ == "__main__":
    app = wx.App(False)
    frame = WeightScaleApp()
    frame.Show()
    app.MainLoop()