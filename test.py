import wmi
import win32service
import win32service, win32api, win32con, winerror
import sys, pywintypes, os


def ServiceStart(ip_address, service_name):
    
    c = wmi.WMI(ip_address)
    # c = wmi.WMI(ip_address, user="Jayram5G", password="J10aysinH@")
    service = c.Win32_Service(Name=service_name)
    print(service[0])
    print(service[0].State)
    
    try:
        
        # services_all = {service.Name: service.State for service in c.Win32_Service()}
        # print(services_all)
        # print(services_all[service_name])
        # print(services_all[service_name] == "Running")
        # print(services_all[service_name] == "Stopped")

        if service:

            if service[0].State == 'Running':
                print(f"{ip_address} : {service_name} is already running.")
            else:
                # Start the service
                try:
                    # result, = service[0].StartService()
                    win32service.StartService(service_name)
                except Exception as e:
                    print(e)
                
        else:
            print(f"{ip_address} : {service_name} - Service not found.")
    except Exception as e:
        print(f"{ip_address} : {service_name} - {e}.")

def SmartOpenService(hscm, name, access):

    return win32service.OpenService(hscm, name, access)
        
    # name = win32service.GetServiceKeyName(hscm, name)
    # return win32service.OpenService(hscm, name, access)

def StartService(serviceName, args = None, machine = None):
    hscm = win32service.OpenSCManager(machine,None,win32service.SC_MANAGER_ALL_ACCESS)
    try:

        hs = SmartOpenService(hscm, serviceName, win32service.SERVICE_ALL_ACCESS)
        try:
            win32service.StartService(hs, args)
        finally:
            win32service.CloseServiceHandle(hs)
    finally:
        win32service.CloseServiceHandle(hscm)

StartService("com.docker.service", "192.168.29.11")
# ServiceStart("192.168.29.11", "com.docker.service")
