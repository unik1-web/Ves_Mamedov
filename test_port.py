import serial
import time
import sys

def test_port(port_name):
    print(f"Тестирование порта {port_name}")
    print("=" * 50)
    
    # Сначала попробуем просто послушать порт
    print("\nТест 1: Прослушивание порта")
    print("-" * 50)
    try:
        ser = serial.Serial(
            port=port_name,
            baudrate=9600,
            bytesize=8,
            parity='N',
            stopbits=1,
            timeout=1.0
        )
        print("Порт открыт, ожидание данных...")
        print("Нажмите Ctrl+C для остановки")
        
        start_time = time.time()
        while True:
            if ser.in_waiting:
                data = ser.read(ser.in_waiting)
                print(f"Получено (hex): {data.hex()}")
                try:
                    print(f"Получено (ASCII): {data.decode('ascii', errors='replace')}")
                except:
                    pass
            time.sleep(0.1)
            
    except KeyboardInterrupt:
        print("\nПрослушивание остановлено")
    except Exception as e:
        print(f"Ошибка: {str(e)}")
    finally:
        if 'ser' in locals():
            ser.close()
            print("Порт закрыт")
    
    # Тестируем разные параметры порта
    print("\nТест 2: Проверка разных параметров")
    print("=" * 50)
    
    port_configs = [
        {"baudrate": 9600, "bytesize": 8, "parity": 'N', "stopbits": 1, "desc": "Стандартные параметры"},
        {"baudrate": 4800, "bytesize": 8, "parity": 'N', "stopbits": 1, "desc": "Пониженная скорость"},
        {"baudrate": 19200, "bytesize": 8, "parity": 'N', "stopbits": 1, "desc": "Повышенная скорость"},
        {"baudrate": 9600, "bytesize": 7, "parity": 'E', "stopbits": 1, "desc": "7 бит, четность"},
        {"baudrate": 9600, "bytesize": 8, "parity": 'O', "stopbits": 2, "desc": "8 бит, нечетность, 2 стоп-бита"}
    ]
    
    # Тестируем разные команды
    commands = [
        (b'\x80P\r\n', "Запрос веса (двоичный)"),
        (b'Z\r\n', "Команда тары"),
        (b'?\r\n', "Запрос статуса"),
        (b'P\r\n', "Запрос веса (ASCII)"),
        (b'\x02P\r\n', "Запрос веса (STX)"),
        (b'\x02\r\n', "STX"),
        (b'\x03\r\n', "ETX"),
        (b'\x80\r\n', "Заголовок"),
        (b'\x80P', "Запрос веса без CR/LF"),
        (b'P', "P без CR/LF")
    ]
    
    for config in port_configs:
        print(f"\nТестирование с параметрами: {config['desc']}")
        print(f"Скорость: {config['baudrate']}, Биты: {config['bytesize']}, Четность: {config['parity']}, Стоп-биты: {config['stopbits']}")
        print("-" * 50)
        
        try:
            # Пробуем открыть порт
            print("1. Попытка открыть порт...")
            ser = serial.Serial(
                port=port_name,
                baudrate=config['baudrate'],
                bytesize=config['bytesize'],
                parity=config['parity'],
                stopbits=config['stopbits'],
                timeout=1.0,
                write_timeout=1.0
            )
            print("✓ Порт успешно открыт")
            
            # Очищаем буферы
            print("\n2. Очистка буферов...")
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            print("✓ Буферы очищены")
            
            # Отправляем команды
            print("\n3. Отправка команд...")
            
            for cmd, desc in commands:
                print(f"\nОтправка: {desc}")
                print(f"Команда (hex): {cmd.hex()}")
                ser.write(cmd)
                print("✓ Команда отправлена")
                
                # Ждем ответ
                print("Ожидание ответа...")
                time.sleep(0.5)
                
                if ser.in_waiting:
                    data = ser.read(ser.in_waiting)
                    print(f"Получено (hex): {data.hex()}")
                    try:
                        print(f"Получено (ASCII): {data.decode('ascii', errors='replace')}")
                    except:
                        pass
                else:
                    print("Нет ответа")
            
            # Закрываем порт
            print("\n4. Закрытие порта...")
            ser.close()
            print("✓ Порт закрыт")
            
        except serial.SerialException as e:
            print(f"Ошибка: {str(e)}")
            continue
        except Exception as e:
            print(f"Неожиданная ошибка: {str(e)}")
            continue
    
    print("\nТестирование завершено")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        port = sys.argv[1]
    else:
        port = input("Введите имя COM-порта (например, COM4): ")
    
    test_port(port) 