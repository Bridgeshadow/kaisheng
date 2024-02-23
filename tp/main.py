import serial
import time
import keyboard
import configparser
import re
import logging
from logging.handlers import TimedRotatingFileHandler
import os

config = configparser.ConfigParser()
config.read('D:/tp/config/config.ini')

log_folder = config.get('LogSettings', 'log_folder')
# 如果文件夹不存在，则创建
if not os.path.exists(log_folder):
    os.makedirs(log_folder)

# 设置按日期滚动的TimedRotatingFileHandler
log_file_path = os.path.join(log_folder, 'logfile.log')
log_handler = TimedRotatingFileHandler(log_file_path, when="midnight", interval=1, backupCount=7)
log_handler.setLevel(logging.DEBUG)
log_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] - %(message)s'))

# 设置根记录器（root logger）并加入TimedRotatingFileHandler
logging.basicConfig(level=logging.DEBUG, handlers=[log_handler])



def send_string(s):
    for c in s:
        if c.isdigit() or c == '.':
            keyboard.press_and_release(c)

def read_balance_data(serial_port, baud_rate, parity, data_bits, stop_bits):
    ser = serial.Serial(port=serial_port, baudrate=baud_rate, parity=parity,
                        stopbits=stop_bits, bytesize=data_bits, timeout=2)  # 设置timeout为2秒

    try:
        ser.flushInput()  # 刷新输入缓存
        ser.flushOutput()  # 刷新输出缓存
        while True:
            time.sleep(0.5)  # 等待一段时间，让缓冲区填充

            data_bytes = ser.readline()  # 读取一行数据，直到遇到换行符

            try:
                # 检查数据是否是有效的UTF-8编码
                data = data_bytes.decode('utf-8', errors='strict').strip()
                # 使用正则表达式从字符串中提取数值部分
                numeric_values = re.findall(r'\b(\d+\.\d+)\b', data)
                if numeric_values:
                    weight = numeric_values[0]
                    send_string(weight)
            except UnicodeDecodeError as e:
                logging.error(f"解码数据时出错: {str(e)}")

    except serial.SerialException as e:
        logging.error(f"tp_weight: {str(e)}")

    finally:
        ser.close()

def main():
    serial_port = config.get('Serial', 'Port')
    baud_rate = int(config.get('Serial', 'BaudRate'))
    parity = config.get('Serial', 'Parity')
    data_bits = int(config.get('Serial', 'DataBits'))
    stop_bits = int(config.get('Serial', 'StopBits'))
    try:
        read_balance_data(serial_port, baud_rate, parity, data_bits, stop_bits)
    except Exception as e:
        logging.error(f"连接到串口时出错: {str(e)}")

if __name__ == "__main__":
    main()
