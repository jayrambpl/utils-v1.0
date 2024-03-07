import sys
import os
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication, QMainWindow, QStackedWidget
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from PyQt5.QtWidgets import QMenu, QMenuBar, QAction

from PyQt5.QtGui import QPixmap
from configparser import ConfigParser
import datetime
from datetime import timedelta
import threading
import time
import subprocess
import platform
import psutil
import wmi
import win32service
import win32serviceutil


class LoginWindows(QDialog):
    def __init__(self, widget):
        super(LoginWindows, self).__init__()
        loadUi("login.ui", self)
        self.setWindowTitle("Server Util v1.0")
        self.userPasswordInput.setEchoMode(QtWidgets.QLineEdit.Password)
        self.loginbutton.clicked.connect(self.gotologin)
        self.widget = widget  # Store the widget object

    def gotologin(self):
        username = self.userNameInput.text()
        password = self.userPasswordInput.text()
        if username == "" or password == "":
            self.err_label.setText("Please enter username and password")
        else:
            self.err_label.setText("")

        if username == "Rajesh" and password == "admin99":
            Utilwindow = ServerUtilwindows(self.widget)
            self.widget.addWidget(Utilwindow)
            self.widget.setCurrentIndex(self.widget.currentIndex()+1)
        else:
            self.err_label.setText("Invalid username or password.")

class ServerUtilwindows(QMainWindow):
    def __init__(self, widget):
        super(ServerUtilwindows, self).__init__()
        loadUi("main.ui", self)
        self.setWindowTitle("Server Util v1.0")
        self.PINGbutton.clicked.connect(self.ping_servers)
        self.ServiceStatusButton.clicked.connect(self.CheckServiceStatus)
        self.ServiceStartButton.clicked.connect(self.ServiceStart)
        self.actionLoad_Server_List.triggered.connect(self.load_server_list)
        self.actionLoad_Service_List.triggered.connect(self.load_service_list)
        self.actionLoad_Server_List.setShortcut("Ctrl+A")
        self.actionLoad_Service_List.setShortcut("Ctrl+S")
        self.actionPING.triggered.connect(self.ping_servers)
        self.actionPING.setShortcut("Ctrl+I")
        self.menuExit.triggered.connect(self.closeApp)
        self.Logs_text.setReadOnly(True)
        self.progressBar.setValue(0)
        self.log("Server Util v1.0 - Started ...")
        self.ServiceRegistered = []
        self.ServiceNotRegistered = []
        self.ServiceNotRuning = []
                 
        self.widget = widget  # Store the widget object
        self.load_init_data()
# ---------------------- 
    
    def log(self, message):
        current_text = self.Logs_text.toPlainText()
        time_stamp = f"{datetime.datetime.now().strftime('%d-%b-%Y %I:%M %p')} :"
        new_text = f"{time_stamp} {message}\n{current_text}"
        self.Logs_text.setPlainText(new_text)

    def closeApp(self):
        QApplication.quit()

# ----------------------     
    def load_server_list(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Text Files (*.txt)")
        if file_path:
            server_list = self.read_unique_list(file_path)
            self.list_servers.clear()
            self.list_failed.clear()
            self.list_passed.clear()
            self.list_servers.addItems(server_list)
            self.log(f"Loaded {len(server_list)} servers from file." )
        else:
            self.log(f"No servers list in file." )
# ---------------------- 
    def load_service_list(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Text Files (*.txt)")
        if file_path:
            service_list = self.read_unique_list(file_path)
            self.list_serviceStatus.clear()
            self.list_StartServiceStatus.clear()
            self.list_service.clear()
            self.list_service.addItems(service_list)
            self.log(f"Loaded {len(service_list)} services from file." )
        else:
            self.log(f"No service list in file." )    
# ---------------------- 
    def load_init_data(self):
        conf = ConfigParser()
        
        try:
            if (os.path.isfile('config.ini')):
                conf.read('config.ini')
                server_list_file = conf.get('Default', 'server_list')
                service_list_file = conf.get('Default', "service_list")
                server_list_path = os.path.join(os.getcwd(), server_list_file)
                service_list_path = os.path.join(os.getcwd(), service_list_file) 
                server_list = self.read_unique_list(server_list_path)
                self.list_servers.addItems(server_list)

                service_list = self.read_unique_list(service_list_path)
                self.list_service.addItems(service_list)

                # self.read_unique_list(service_list_path, self.list_servers)
             
            else:
                raise Exception("Config.ini file not found in current directory.")
        except Exception as e:
            QMessageBox.information(self, "Error", f"Error loading config.ini file: {e}")   
            # return
        
# ---------------------- 
    def read_unique_list(self, file_path):
        try:
            with open(file_path, "r") as file:
                line_items = file.readlines()
         
            unique_list = sorted(set(map(str.strip, line_items)))
            return unique_list
            
        except Exception as e:
            QMessageBox.information(self, "Error", f"Error loading config.ini file: {e}")   
# ---------------------- 
    def ping_servers(self):
        start_time = time.time()
        self.log("Pinging servers...")
        
        ping_results = []  # Clear existing ping results
        self.progressBar.setValue(0)
        
        item_list = [self.list_servers.item(i).text() for i in range(self.list_servers.count())]
        total_servers = len(item_list)

        if item_list:

            for i, server in enumerate(item_list, start=1):
                try:
                    
                    ping_cmd = ["ping", "-n", "1"] if platform.system().lower() == "windows" else ["ping", "-c", "1"]
                    startupinfo = None
                    if platform.system().lower() == "windows":
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    result = subprocess.run(
                        ping_cmd + [server],
                        capture_output=True,
                        text=True,
                        timeout=5,
                        startupinfo=startupinfo  
                    )
                    
                    success = "Destination host unreachable" not in result.stdout and "Request timed out" not in result.stdout
                    if success:
                        self.list_passed.addItem(server)
                        self.log(f"{server} : {success}")
                        
                    else:
                        self.list_failed.addItem(server)
                        self.log(f"{server} : {success}")

                except Exception as e:
                    self.list_failed.addItem(server)
                    self.log(f"{server} : {success}")
                    
                self.progressBar.setValue(int((i / total_servers) * 100)) 
                self.statusbar.showMessage(f"Status: {i} out of {total_servers} completed.")
                
                end_time = time.time()  
                elapsed_time = end_time - start_time
                formatted_time = str(timedelta(seconds=elapsed_time)).split(".")[0]
                # self.log(f"Time Taken: {formatted_time}")
        else:
            QMessageBox.information(self, "Error", "No servers to ping.")
            return
           
        end_time = time.time()  
        elapsed_time = end_time - start_time
        formatted_time = str(timedelta(seconds=elapsed_time)).split(".")[0]
        self.log(f"Time Taken: {formatted_time}")
        self.log("Ping Completed")
        # self.generate_reports()
        # QMessageBox.information(self, "Ping Completed", f"Ping completed in {formatted_time}")
        

# -----------------------------
    def CheckServiceStatus(self):
        start_time = time.time()
        
        # ping_results = []  # Clear existing ping results
        self.progressBar.setValue(0)
        item_list = [self.list_passed.item(i).text() for i in range(self.list_passed.count())]
        service_list = [self.list_service.item(i).text() for i in range(self.list_service.count())]
        total_service = len(service_list)
        total_servers = len(item_list)
        self.log(f"Checking {total_service} services {total_servers} servers.")
        self.list_serviceRunning.clear()
        self.list_ServiceNOTRunning.clear()
        
        if item_list:
            for i, server_ip in enumerate(item_list, start=1):
                try:
                    c = wmi.WMI(server_ip)
                    services = {service.Name: service.State for service in c.Win32_Service()}
                    
                    ServiceNotRegistered = [item for item in service_list if item not in services]
                    ServiceRegistered = [item for item in service_list if item in services]
                    
                    if ServiceNotRegistered:
                        self.log(f"{server_ip} : {ServiceNotRegistered} not registered.")
                        # continue

                    for service_name in ServiceRegistered:
                        
                        line_item = ""
                        if services[service_name] == "Running":
                            line_item=(f"{server_ip}-{service_name}-{services[service_name]}")
                            self.list_serviceRunning.addItem(line_item)
                            self.log(line_item)

                        elif services[service_name] == "Stopped":
                            line_item = (f"{server_ip}-{service_name}-{services[service_name]}")
                            StoppedItems = (f"{server_ip}-{service_name}")
                            self.ServiceNotRuning.append(StoppedItems)
                            self.list_ServiceNOTRunning.addItem(line_item)
                            self.log(line_item)
                            
                        else:
                            line_item = (f"{server_ip}-{service_name}-{services[service_name]}")
                            self.list_ServiceNOTRunning.addItem(line_item)
                            self.log(line_item)
                    
                except wmi.x_wmi as e:
                    self.log(f"WMI Error on {server_ip}: {str(e)}\n")

                except Exception as e:
                    self.log(f"Error on {server_ip}: {str(e)}\n")
                
                self.progressBar.setValue(int((i / total_servers) * 100))
                
        
        end_time = time.time()  
        elapsed_time = end_time - start_time
        formatted_time = str(timedelta(seconds=elapsed_time)).split(".")[0]
        self.log(f"Time Taken: {formatted_time}")
        # QMessageBox.information(self, "Service Status Check Completed", f"Service Status Check completed in {formatted_time}") 

    # ----------------------------      
   
    def ServiceStart(self):
        
        for item in self.ServiceNotRuning:
            ip_address, service_name = item.split("-")
            self.log(f"{ip_address} : '{service_name}' - starting...")
            # c = wmi.WMI(ip_address, user="Jayram5G", password="J10aysinH@")
            try:
                c = wmi.WMI(ip_address)
            except Exception as e:
                self.log(f"Error connecting to {ip_address}: {str(e)}")
                continue
            try:
                service = c.Win32_Service(Name=service_name)
            except Exception as e:
                self.log(f"Error getting service '{service_name}' on {ip_address}: {str(e)}")
                continue

            if service:
                try:
                    # win32service.StartService(ip_address, service_name)
                    # win32serviceutil.StartService(service_name, machine=ip_address)
                    c.Win32_Service.StartService(Name=service_name)

                    self.log(f"{ip_address} : {service_name} - Started.")
                except win32service.error as e:
                    self.log(f"Error starting service '{service_name}' on {ip_address}: {str(e)}")
                    continue    
                # service[0].StartService()
                self.log(f"{ip_address} : {service_name} - Started.")
            else:
                self.log(f"{ip_address} :{service_name} - Not able to start.")

                     
# ------------------------
def main():
    app = QApplication(sys.argv)
    widget = QStackedWidget()
    login = LoginWindows(widget)
    
    widget.addWidget(login)
    widget.setFixedHeight(600)
    widget.setFixedWidth(950)
    widget.show()

    try:
        sys.exit(app.exec_())
    except:
        print("Exiting")

if __name__ == "__main__":
    main()
