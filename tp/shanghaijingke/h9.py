import socket
import time
import configparser
import re
import logging
from logging.handlers import TimedRotatingFileHandler
import os

config = configparser.ConfigParser()
config.read('config.ini')

log_folder = config.get('LogSettings', 'log_folder')
if not os.path.exists(log_folder):
    os.makedirs(log_folder)

log_file_path = os.path.join(log_folder, 'logfile.log')
log_handler = TimedRotatingFileHandler(log_file_path, when="midnight", interval=1, backupCount=7)
log_handler.setLevel(logging.DEBUG)
log_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] - %(message)s'))

logging.basicConfig(level=logging.DEBUG, handlers=[log_handler])



def read_balance_data(server_ip, server_port):
    while True:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.connect((server_ip, server_port))
                while True:
                    data = sock.recv(1024)

                    print(data)
                        
        except Exception as e:
            logging.error(f"连接到服务器时出错: {str(e)}")
            time.sleep(5)  # 如果连接失败，等待一段时间后重试

def main():
    # server_ip = config.get('Server', 'IP')
    # server_port = int(config.get('Server', 'Port'))
    server_ip = '192.168.1.200'
    server_port = 2000
    try:
        read_balance_data(server_ip, server_port)
    except Exception as e:
        logging.error(f"连接到服务器时出错: {str(e)}")

if __name__ == "__main__":
    main()
