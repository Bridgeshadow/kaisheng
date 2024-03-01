import os
import socket
import time
import configparser
import re
import logging
from mysql.connector import pooling
import threading

config = configparser.ConfigParser()
config.read('config.ini')

log_folder = config.get('LogSettings', 'log_folder')
if not os.path.exists(log_folder):
    os.makedirs(log_folder)

log_file_path = os.path.join(log_folder, 'logfile.log')
log_handler = logging.FileHandler(log_file_path)
log_handler.setLevel(logging.DEBUG)
log_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] - %(message)s'))

logging.basicConfig(level=logging.DEBUG, handlers=[log_handler])

# 创建数据库连接池
db_connection_pool = pooling.MySQLConnectionPool(
    pool_name="my_pool",
    pool_size=5,
    host=config.get('Database', 'DB_Host'),
    database=config.get('Database', 'DB_Name'),
    user=config.get('Database', 'DB_User'),
    password=config.get('Database', 'DB_Password')
)

def read_balance_data(device_ip, device_port, device_name, regex_pattern):
    while True:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.connect((device_ip, device_port))
                while True:
                    data = sock.recv(1024)
                    if data:
                        parsed_data = re.findall(regex_pattern, data.decode())
                        if parsed_data:
                            save_to_database(device_ip, device_port, device_name, parsed_data)
                    else:
                        break
        except Exception as e:
            logging.error(f"连接到设备时出错: {str(e)}")
            time.sleep(2)  # 如果连接失败，等待一段时间后重试

def save_to_database(device_ip, device_port, device_name, parsed_data):
    try:
        connection = db_connection_pool.get_connection()
        cursor = connection.cursor()
        for data in parsed_data:
            query = "INSERT INTO your_table_name (device_ip, device_port, device_name, data) VALUES (%s, %s, %s, %s)"
            cursor.execute(query, (device_ip, device_port, device_name, data))
        connection.commit()
        cursor.close()
        connection.close()
    except Exception as e:
        logging.error(f"保存到数据库时出错: {str(e)}")

def handle_device(device_ip, device_port, device_name, regex_pattern):
    thread_name = threading.current_thread().name
    logging.info(f"线程 {thread_name} 正在处理设备 {device_name}")
    read_balance_data(device_ip, device_port, device_name, regex_pattern)

def main():
    threads = []
    for section in config.sections():
        if section.startswith('Device'):
            device_ip = config.get(section, 'Device_IP')
            device_port = int(config.get(section, 'Device_Port'))
            device_name = config.get(section, 'Device_Name')
            regex_pattern = config.get(section, 'Device_Regex')
            thread = threading.Thread(target=handle_device, args=(device_ip, device_port, device_name, regex_pattern))
            threads.append(thread)
            thread.start()

    # 等待所有线程结束
    for thread in threads:
        thread.join()

if __name__ == "__main__":
    main()
