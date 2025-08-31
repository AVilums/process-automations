import psutil
import datetime
import time
import os
import subprocess

log_file = "server_log.txt"
backup_folder = "log_backup"
downtime_flag_file = "downtime_flag.txt"
target_server = " "
max_consecutive_failures = 3


def log(message):
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"{current_time} - {message}"
    print(log_message)
    with open(log_file, "a") as f:
        f.write(log_message + "\n")


def backup_log():
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    backup_filename = f"{backup_folder}/server_log_backup_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
    os.rename(log_file, backup_filename)
    log("Log file backed up.")
    print(f"-----------------------------------------------------------")


def check_downtime():
    consecutive_failures = 0

    while True:
        with open(os.devnull, 'w') as DEVNULL:
            try:
                subprocess.check_call(
                    ['ping', '-n', '2', target_server],
                    stdout=DEVNULL,  # suppress output
                    stderr=DEVNULL
                )
                log(f"Server status: {True}")
                return False
            except subprocess.CalledProcessError:
                log(f"Server status: {False}")
                consecutive_failures += 1
                response = False
        if consecutive_failures >= max_consecutive_failures:
            return True


while True:
    try:
        if check_downtime():
            log("Downtime detected!")
            backup_log()
            with open(downtime_flag_file, "w") as flag:
                flag.write("Downtime detected!")

        current_hour = datetime.datetime.now().hour
        if current_hour == 0:
            backup_log()

        cpu_usage = psutil.cpu_percent(interval=1)
        memory_usage = psutil.virtual_memory().percent
        network_sent = psutil.net_io_counters().bytes_sent
        network_received = psutil.net_io_counters().bytes_recv

        log(f"CPU Usage: {cpu_usage}%")
        log(f"Memory Usage: {memory_usage}%")
        log(f"Network Sent: {network_sent} bytes")
        log(f"Network Recv: {network_received} bytes")
        print(f"-----------------------------------------------------------")

    except Exception as e:
        log(f"Error: {str(e)}")

    time.sleep(10)