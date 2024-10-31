# -*- coding: utf-8 -*-
import subprocess
import time
from configuration import red, green, yellow, log_col, debug_mode
from typing import List, Tuple, Dict, Any, Optional, Iterator, IO
import abc, glob, os, gzip, ctypes, sys, socket
from collections import defaultdict
from datetime import datetime
from typing import Optional
import urllib.request
import urllib.error
import html


# Section eight left side
# WSL details
def wsl():
    command_wsl = "(Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux).State"
    result_wsl = subprocess.run(["powershell", command_wsl], shell=True, stdout=subprocess.PIPE)
    output_wsl = result_wsl.stdout.decode("utf-8").strip()

    report_wsl=f"""<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        Windows Subsystem for Linux (WSL)
                    </h2>
                    <div class="reportSectionBody">"""
    result_wsl_list = subprocess.run(["powershell", "wsl -l"], shell=True, stdout=subprocess.PIPE)
    output_wsl_list = result_wsl_list.stdout.replace(b'\x00', b'').decode('ascii')
    wsl_list = output_wsl_list.split("\n")
    wsl_list = [i for i in wsl_list if i]
    if output_wsl == "Enabled" or wsl_list[0].strip() == 'Windows Subsystem for Linux Distributions:':
        if wsl_list[0].strip() == 'Windows Subsystem for Linux Distributions:':
            for item in wsl_list[1:]:
                if "(Default)" in item:
                    report_wsl+=f"""<b>{item.strip()}</b><br>"""
                else:
                    report_wsl+=f"""{item.strip()}<br>"""
        else:
            report_wsl+="""Windows Subsystem for Linux (WSL) is <b>Enabled</b> on this system but No WSL Distributions are found."""
    else:
        report_wsl+="""Windows Subsystem for Linux Distributions is not enabled."""

    report_wsl+=f"""</div>
                </div>"""
    return report_wsl

# WSL
def wsl_report(show_verbose):
    report_wsl = ""
    print(yellow, end='')
    print('[!] Fetching WSL details')
    if show_verbose:
        st_time_wsl = time.localtime()
        print(log_col,end="")
        print('[v] WSL details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_wsl)))
    try:
        report_wsl = wsl()
        print(green, end='')
        print('[√] WSL details fetched successfully')
    except Exception as e:
        report_wsl = f"""<div class="reportSection rsRight">
        <h2 class="reportSectionHeader">
            Windows Subsystem for Linux (WSL)
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching WSL Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching WSL details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_wsl = time.localtime()
        print(log_col,end="")
        print('[v] WSL details fetched at '+str(time.strftime("%H:%M:%S", ed_time_wsl)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.fromtimestamp(time.mktime(ed_time_wsl)) - datetime.fromtimestamp(time.mktime(st_time_wsl))).split(":"))))
    return {'report_wsl': report_wsl}


# Usb details
if sys.platform == "win32" :
    import winreg

def convert_binary_to_ascii_string(binary_string: bytes) -> str:
    return ''.join([chr(byte) for byte in binary_string if 0 < byte < 128])

def convert_windows_time_to_unix(windows_timestamp: int) -> int:
    windows_tick = 10000000
    seconds_to_unix_epoch = 11644473600
    return int(windows_timestamp / windows_tick - seconds_to_unix_epoch)

def parse_windows_log_file(filepath: str) -> Iterator[List[str]]:
    with open(filepath, 'r', encoding="iso-8859-1") as log_file:
        log_section = []
        for line in log_file:
            if len(log_section) == 0 and line.startswith('>>>'):
                next_line = next(log_file)
                if next_line.startswith('>>>'):
                    log_section.append(line)
                    log_section.append(next_line)
                    continue
            if len(log_section) > 0 and line.startswith('<<<'):
                next_line = next(log_file)
                if next_line.startswith('<<<'):
                    log_section.append(line)
                    log_section.append(next_line)
                    yield log_section
                    log_section.clear()
                    continue
            if len(log_section) > 0:
                log_section.append(line)

def get_device_info_from_web(vendor_id: str, product_id: str, max_attempts: int = 3) -> Tuple[Optional[str], Optional[str]]:
    if max_attempts <= 0:
        return None, None
    url = f'https://devicehunt.com/view/type/usb/vendor/{vendor_id}/device/{product_id}'
    try:
        with urllib.request.urlopen(url) as response:
            html = response.read().decode('utf-8')
    except urllib.error.URLError as e:
        html = '{ "vendor" : "Error to fetch",  "device" : "Error to fetch"}'
    start_vendor = html.find('--type-vendor')
    if start_vendor == -1:
        vendor_name = 'N/A'
    else:
        start_vendor = html.find('details__heading', start_vendor) + 17
        end_vendor = html.find('<', start_vendor)
        vendor_name = (html[start_vendor:end_vendor].strip()).replace('>\n', '')
    start_device = html.find('--type-device')
    if start_device == -1:
        device_description = 'N/A'
    else:
        start_device = html.find('details__heading', start_device) + 17
        end_device = html.find('<', start_device)
        device_description = (html[start_device:end_device].strip()).replace('>\n', '')
    return vendor_name, device_description

def open_linux_log_file(filepath: str) -> IO:
    if filepath.endswith('.gz'):
        return gzip.open(filepath, 'rt')
    else:
        return open(filepath, 'r')

def parse_linux_log_file(filepath: str) -> Iterator[List[str]]:
    with open_linux_log_file(filepath) as log_file:
        section = []
        for line in log_file:
            if 'New USB device found' in line:
                if len(section) == 0:
                    section.append(line)
                    continue
                else:
                    section.clear()
                    section.append(line)
            if len(section) > 0:
                section.append(line)
                if 'Mounted /dev/sd' in line:
                    yield section
                    section.clear()



class USBDevice:
    def __init__(
        self,
        version: Optional[str] = None,
        serial_number: Optional[str] = None,
        friendly_name: Optional[str] = None,
        vendor_id: Optional[str] = None,
        product_id: Optional[str] = None,
        first_connect_date: Optional[datetime] = None,
        last_connect_date: Optional[datetime] = None,
        vendor_name: Optional[str] = None,
        product_description: Optional[str] = None,
    ):
        self.version = version
        self.serial_number = serial_number
        self.friendly_name = friendly_name
        self.vendor_id = vendor_id
        self.product_id = product_id
        self.first_connect_date = first_connect_date
        self.last_connect_date = last_connect_date
        self.vendor_name = vendor_name
        self.product_description = product_description

    def get_details(self) -> str:
        details = f'Device: {self.friendly_name}\n'
        details += f'Vendor Name: {self.vendor_name}\n'
        details += f'Product Description: {self.product_description}\n'
        details += f'First Connect Date: {self.first_connect_date}\n'
        details += f'Last Connect Date: {self.last_connect_date}\n'
        details += f'Serial Number: {self.serial_number}\n'
        details += f'Vendor ID: {self.vendor_id}\n'
        details += f'Product ID: {self.product_id}\n'
        details += f'Version: {self.version}\n'
        return details


class USBDeviceWindows(USBDevice):
    def __init__(
        self,
        version: Optional[str] = None,
        serial_number: Optional[str] = None,
        friendly_name: Optional[str] = None,
        vendor_id: Optional[str] = None,
        product_id: Optional[str] = None,
        first_connect_date: Optional[datetime] = None,
        last_connect_date: Optional[datetime] = None,
        vendor_name: Optional[str] = None,
        product_description: Optional[str] = None,
        usbstor_vendor: Optional[str] = None,
        usbstor_product: Optional[str] = None,
        parent_prefix_id: Optional[str] = None,
        guid: Optional[str] = None,
        drive_letter: Optional[str] = None,
    ):
        super().__init__(
            version=version,
            serial_number=serial_number,
            friendly_name=friendly_name,
            vendor_id=vendor_id,
            product_id=product_id,
            first_connect_date=first_connect_date,
            last_connect_date=last_connect_date,
            vendor_name=vendor_name,
            product_description=product_description,
        )
        self.usbstor_vendor = usbstor_vendor
        self.usbstor_product = usbstor_product
        self.parent_prefix_id = parent_prefix_id
        self.guid = guid
        self.drive_letter = drive_letter

    def get_details(self) -> str:
        details = super().get_details()
        details += f'USBSTOR Vendor: {self.usbstor_vendor}\n'
        details += f'USBSTOR Product: {self.usbstor_product}\n'
        details += f'Drive Letter: {self.drive_letter}\n'
        details += f'GUID: {self.guid}\n'
        details += f'Parent Prefix ID: {self.parent_prefix_id}\n'
        return details


class USBDeviceLinux(USBDevice):
    def __init__(
        self,
        version: Optional[str] = None,
        serial_number: Optional[str] = None,
        friendly_name: Optional[str] = None,
        vendor_id: Optional[str] = None,
        product_id: Optional[str] = None,
        first_connect_date: Optional[datetime] = None,
        last_connect_date: Optional[datetime] = None,
        vendor_name: Optional[str] = None,
        product_description: Optional[str] = None,
        syslog_manufacturer: Optional[str] = None,
        syslog_product: Optional[str] = None,
        device_size: Optional[str] = None,
    ):
        super().__init__(
            version=version,
            serial_number=serial_number,
            friendly_name=friendly_name,
            vendor_id=vendor_id,
            product_id=product_id,
            first_connect_date=first_connect_date,
            last_connect_date=last_connect_date,
            vendor_name=vendor_name,
            product_description=product_description,
        )
        self.syslog_manufacturer = syslog_manufacturer
        self.syslog_product = syslog_product
        self.device_size = device_size

    def get_details(self) -> str:
        details = super().get_details()
        details += f'SYSLOG Manufacturer: {self.syslog_manufacturer}\n'
        details += f'SYSLOG Product: {self.syslog_product}\n'
        details += f'Device size: {self.device_size}\n'
        return details

# @dataclass(init=True, repr=True, eq=False)
# class USBDevice:
#     version: Optional[str] = None
#     serial_number: Optional[str] = None
#     friendly_name: Optional[str] = None
#     vendor_id: Optional[str] = None
#     product_id: Optional[str] = None
#     first_connect_date: Optional[datetime] = None
#     last_connect_date: Optional[datetime] = None
#     vendor_name: Optional[str] = None
#     product_description: Optional[str] = None
#     def get_details(self) -> str:
#         details = f'Device: {self.friendly_name}\n'
#         details += f'Vendor Name: {self.vendor_name}\n'
#         details += f'Product Description: {self.product_description}\n'
#         details += f'First Connect Date: {self.first_connect_date}\n'
#         details += f'Last Connect Date: {self.last_connect_date}\n'
#         details += f'Serial Number: {self.serial_number}\n'
#         details += f'Vendor ID: {self.vendor_id}\n'
#         details += f'Product ID: {self.product_id}\n'
#         details += f'Version: {self.version}\n'
#         return details

# @dataclass(init=True, repr=True, eq=False)
# class USBDeviceWindows(USBDevice):
#     usbstor_vendor: Optional[str] = None
#     usbstor_product: Optional[str] = None
#     parent_prefix_id: Optional[str] = None
#     guid: Optional[str] = None
#     drive_letter: Optional[str] = None

#     def get_details(self) -> str:
#         details = super().get_details()
#         details += f'USBSTOR Vendor: {self.usbstor_vendor}\n'
#         details += f'USBSTOR Product: {self.usbstor_product}\n'
#         details += f'Drive Letter: {self.drive_letter}\n'
#         details += f'GUID: {self.guid}\n'
#         details += f'Parent Prefix ID: {self.parent_prefix_id}\n'
#         return details

# @dataclass(init=True, repr=True, eq=False)
# class USBDeviceLinux(USBDevice):
#     syslog_manufacturer: Optional[str] = None
#     syslog_product: Optional[str] = None
#     device_size: Optional[str] = None

#     def get_details(self) -> str:
#         details = super().get_details()
#         details += f'SYSLOG Manufacturer: {self.syslog_manufacturer}\n'
#         details += f'SYSLOG Product: {self.syslog_product}\n'
#         details += f'Device size: {self.device_size}\n'
#         return details

class BaseViewer(abc.ABC):
    @abc.abstractmethod
    def get_usb_devices(self) -> List[USBDevice]:
        pass
    @staticmethod
    def _set_devices_info(usb_devices: List[USBDevice]) -> None:
        for device in usb_devices:
            if device.vendor_id is not None and device.product_id is not None:
                vendor_name, product_description = get_device_info_from_web(device.vendor_id, device.product_id)
                device.vendor_name = vendor_name
                device.product_description = product_description

if sys.platform == "win32":
    class WindowsViewer(BaseViewer):
        __USBSTOR_PATH = r'SYSTEM\CurrentControlSet\Enum\USBSTOR'
        __USB_PATH = r'SYSTEM\CurrentControlSet\Enum\USB'
        __MOUNTED_DEVICES_PATH = r'SYSTEM\MountedDevices'
        __PORTABLE_DEVICES_PATH = r'SOFTWARE\Microsoft\Windows Portable Devices\Devices'
        __MOUNT_POINTS_PATH = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2'
        def __init__(self):
            self.__machine_registry = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
            self.__user_registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        def __del__(self):
            self.__machine_registry.Close()
            self.__user_registry.Close()
        def get_usb_devices(self) -> List[USBDevice]:
            usb_devices = self.__get_base_device_info()
            self.__set_vendor_and_product_ids(usb_devices)
            self.__set_guids(usb_devices)
            self.__set_drive_letters(usb_devices)
            self.__set_first_connect_dates(usb_devices)
            self.__set_last_connect_dates(usb_devices)
            self._set_devices_info(usb_devices)
            return usb_devices
        def __get_base_device_info(self) -> List[USBDeviceWindows]:
            root_key = winreg.OpenKey(self.__machine_registry, WindowsViewer.__USBSTOR_PATH)
            usbstor_keys = self.__get_registry_keys(root_key)
            usb_devices = []
            for key_str in usbstor_keys:
                device_attributes = self.__parse_device_name(key_str)
                if device_attributes is None:
                    continue
                vendor, product, version = device_attributes
                usb_path = rf'{WindowsViewer.__USBSTOR_PATH}\{key_str}'
                usb_key = winreg.OpenKey(self.__machine_registry, usb_path)
                devices_keys = self.__get_registry_keys(usb_key)
                for device in devices_keys:
                    device_key = winreg.OpenKey(self.__machine_registry, rf'{usb_path}\{device}')
                    device_values = self.__get_registry_values(device_key)
                    serial_number = device.split('&')[0]
                    friendly_name = device_values['FriendlyName']
                    if 'ParentPrefixId' in device_values:
                        parent_prefix_id = device_values['ParentPrefixId']
                    else:
                        parent_prefix_id = device
                    usb_device = USBDeviceWindows(
                        usbstor_vendor=vendor,
                        usbstor_product=product,
                        version=version,
                        serial_number=serial_number,
                        friendly_name=friendly_name,
                        parent_prefix_id=parent_prefix_id
                    )
                    usb_devices.append(usb_device)
            return usb_devices
        def __set_vendor_and_product_ids(self, usb_devices: List[USBDeviceWindows]) -> None:
            root_key = winreg.OpenKey(self.__machine_registry, WindowsViewer.__USB_PATH)
            device_ids = self.__get_registry_keys(root_key)
            device_dict = {}
            for device_id in device_ids:
                if 'VID' not in device_id or 'PID' not in device_id:
                    continue
                device_key = winreg.OpenKey(self.__machine_registry, rf'{WindowsViewer.__USB_PATH}\{device_id}')
                serial_number = self.__get_registry_keys(device_key)[0]
                device_dict[serial_number] = device_id
            for device in usb_devices:
                for serial_number, device_id in device_dict.items():
                    if device.serial_number != serial_number:
                        continue
                    device_info = device_id.split('&')
                    device.vendor_id = device_info[0].replace('VID_', '')
                    device.product_id = device_info[1].replace('PID_', '')
        def __set_guids(self, usb_devices: List[USBDeviceWindows]) -> None:
            root_key = winreg.OpenKey(self.__machine_registry, WindowsViewer.__MOUNTED_DEVICES_PATH)
            registry_values = self.__get_registry_values(root_key)
            for device in usb_devices:
                for key, value in registry_values.items():
                    value = convert_binary_to_ascii_string(value)
                    if device.parent_prefix_id not in value:
                        continue
                    if r'\Volume' in key:
                        guid_start_index = key.index('{')
                        device.guid = key[guid_start_index:]
        def __set_drive_letters(self, usb_devices: List[USBDeviceWindows]) -> None:
            root_key = winreg.OpenKey(self.__machine_registry, WindowsViewer.__PORTABLE_DEVICES_PATH)
            registry_keys = self.__get_registry_keys(root_key)
            for device in usb_devices:
                for key in registry_keys:
                    if device.parent_prefix_id not in key:
                        continue
                    device_key = winreg.OpenKey(self.__machine_registry, rf'{WindowsViewer.__PORTABLE_DEVICES_PATH}\{key}')
                    values = self.__get_registry_values(device_key)
                    device.drive_letter = values['FriendlyName']
        @staticmethod
        def __set_first_connect_dates(usb_devices: List[USBDeviceWindows]) -> None:
            time_dict = {}
            for log_path in glob.glob(r'C:\Windows\inf\setupapi.dev*.log'):  # There could be multiple files in system
                for section in parse_windows_log_file(log_path):
                    if 'Device Install ' in section[0] and 'SUCCESS' in section[-1]:
                        install_time = section[-2].split()[-2:]  # Get only date and time from string
                        install_time = ' '.join(install_time)
                        install_time = install_time.split('.')[0]  # Remove milliseconds
                        install_time = datetime.strptime(install_time, '%Y/%m/%d %H:%M:%S')
                        time_dict[section[0]] = install_time
            for device in usb_devices:
                for key, install_time in time_dict.items():
                    if device.serial_number in key:
                        device.first_connect_date = install_time
        def __set_last_connect_dates(self, usb_devices: List[USBDeviceWindows]) -> None:
            root_key = winreg.OpenKey(self.__user_registry, WindowsViewer.__MOUNT_POINTS_PATH)
            guids = self.__get_registry_keys(root_key)
            for device in usb_devices:
                for guid in guids:
                    if device.guid != guid:
                        continue
                    device_key = winreg.OpenKey(self.__user_registry, rf'{WindowsViewer.__MOUNT_POINTS_PATH}\{guid}')
                    timestamp = self.__get_registry_timestamp(device_key)
                    device.last_connect_date = datetime.fromtimestamp(timestamp)
        @staticmethod
        def __parse_device_name(device_name: str) -> Optional[Tuple[str, str, str]]:
            name_split = device_name.split('&')
            if len(name_split) != 4 or name_split[0] != 'Disk':
                return None
            vendor = name_split[1].replace('Ven_', '')
            product = name_split[2].replace('Prod_', '')
            version = name_split[3].replace('Rev_', '')
            return vendor, product, version
        @staticmethod
        def __get_registry_keys(root_key: winreg.HKEYType) -> List[str]:
            key_info = winreg.QueryInfoKey(root_key)
            keys_length = key_info[0]
            return [winreg.EnumKey(root_key, index) for index in range(keys_length)]
        @staticmethod
        def __get_registry_values(root_key: winreg.HKEYType) -> Dict[str, Any]:
            key_info = winreg.QueryInfoKey(root_key)
            values_length = key_info[1]
            values_dict = defaultdict(lambda: None)
            for index in range(values_length):
                name, value, _ = winreg.EnumValue(root_key, index)
                values_dict[name] = value
            return values_dict
        @staticmethod
        def __get_registry_timestamp(root_key: winreg.HKEYType) -> int:
            key_info = winreg.QueryInfoKey(root_key)
            return convert_windows_time_to_unix(key_info[2])

if sys.platform == 'linux' :
    class LinuxViewer(BaseViewer):
        def __init__(self):
            self.__hostname = socket.gethostname()
        def get_usb_devices(self) -> List[USBDevice]:
            usb_devices = self.__get_base_device_info()
            self._set_devices_info(usb_devices)
            return usb_devices
        def __get_base_device_info(self) -> List[USBDeviceLinux]:
            usb_devices = []
            for section, year in self.__get_log_sections():
                serial_number = self.__get_device_info_from_section(section, 'SerialNumber:')
                connect_time = self.__get_device_connect_time(section[0], year)
                device = self.__get_device_if_exist(usb_devices, serial_number)
                if device is not None:
                    device.last_connect_date = connect_time
                    continue
                device = USBDeviceLinux(serial_number=serial_number, first_connect_date=connect_time, last_connect_date=connect_time)
                device.vendor_id = self.__get_device_id_by_type(section[0], 'idVendor')
                device.product_id = self.__get_device_id_by_type(section[0], 'idProduct')
                device.version = self.__get_device_id_by_type(section[0], 'bcdDevice')
                device.syslog_product = self.__get_device_info_from_section(section, 'Product:')
                device.syslog_manufacturer = self.__get_device_info_from_section(section, 'Manufacturer:')
                device.serial_number = self.__get_device_info_from_section(section, 'SerialNumber:')
                device.friendly_name = self.__get_device_info_from_section(section, 'Direct-Access')
                device.device_size = self.__get_device_size(section)
                usb_devices.append(device)
            return usb_devices
        def __get_log_sections(self) -> Iterator[Tuple[List[str], int]]:
            for log_path in sorted(glob.glob('/var/log/syslog*'), reverse=True):
                year = self.__get_file_last_modification_year(log_path)
                for section in parse_linux_log_file(log_path):
                    yield section, year
        @staticmethod
        def __get_file_last_modification_year(filepath: str) -> int:
            timestamp = os.stat(filepath).st_mtime
            return datetime.fromtimestamp(timestamp).year
        def __get_device_connect_time(self, string: str, year: int) -> datetime:
            end_index = string.index(self.__hostname)
            connect_time = string[:end_index].strip()
            connect_time = f'{year}-{connect_time}'
            return datetime.strptime(connect_time, '%Y-%b %d %H:%M:%S')
        @staticmethod
        def __get_device_id_by_type(string: str, id_type: str) -> str:
            index_start = string.index(id_type)
            if ',' in string[index_start:]:
                index_end = string.index(',', index_start)
            else:
                index_end = len(string)
            device_id = string[index_start:index_end]
            device_id = device_id.replace(f'{id_type}=', '')
            device_id = device_id.strip()
            return device_id
        @staticmethod
        def __get_device_info_from_section(section: List[str], info_type: str) -> Optional[str]:
            for line in section:
                if info_type in line:
                    start_index = line.index(info_type)
                    return line[start_index + len(info_type):].split(maxsplit=1)[-1].strip()
            return None
        @staticmethod
        def __get_device_size(section: List[str]) -> Optional[str]:
            for line in section:
                if 'logical blocks' in line:
                    index_start = line.rindex(']')
                    return line[index_start + 1:].strip()
            return None
        @staticmethod
        def __get_device_if_exist(usb_devices: List[USBDeviceLinux], serial_number: str) -> Optional[USBDeviceLinux]:
            for device in usb_devices:
                if device.serial_number == serial_number:
                    return device
            return None

def get_usb_viewer() -> Optional[BaseViewer]:
    if sys.platform == "win32":
        return WindowsViewer()
    elif sys.platform == "linux":
        return LinuxViewer()
    else:
        return None


def usb():
    viewer = get_usb_viewer()
    devices = viewer.get_usb_devices()
    usb_report="""<div class="reportSection rsLeft" id="usbstorage">
				<h2 class="reportSectionHeader">
					USB Storage Use
				</h2>
				<div class="reportSectionBody">"""
    
    Device = ""
    Vendor_Name = ""
    Product_Description = ""
    First_Connect = ""
    Last_Connect = ""
    SN = ""
    VID = ""
    PID = ""
    Ver = ""
    USBSTOR_Vendor = ""
    USBSTOR_Product = ""
    DriveName = ""
    GUID = ""
    ParID = ""
    len_dev = 0
    total_dev = len(devices)
    for device in devices:
        for line in device.get_details().split('\n'):
            if "Device: " in line:
                # usb_report += """<p class="hasInfo" title="First Used: 27-09-2022 15:26:02">"""
                Device = line[8:].strip()
            elif "Vendor Name: " in line:
                Vendor_Name = line[13:].strip()
            elif "Product Description: " in line:
                Product_Description = line[21:].strip()
            elif "First Connect Date: " in line:
                First_Connect = line[20:].strip()
            elif "Last Connect Date: " in line:
                Last_Connect = line[19:].strip()
            elif "Serial Number: " in line:
                SN = line[15:].strip()
            elif "Vendor ID: " in line:
                VID = line[11:].strip()
            elif "Product ID: " in line:
                PID = line[12:].strip()
            elif "Version: " in line:
                Ver = line[9:].strip()
            elif "USBSTOR Vendor: " in line:
                USBSTOR_Vendor = line[16:].strip()
            elif "USBSTOR Product: " in line:
                USBSTOR_Product = line[17:].strip()
            elif "Drive Letter: " in line:
                DriveName = line[14:].strip()
            elif "GUID: " in line:
                GUID = line[6:].strip()
            elif "Parent Prefix ID: " in line:
                ParID = line[18:].strip()
    # print (Device, Vendor_Name, Product_Description, First_Connect, Last_Connect, SN, VID, PID, Ver, USBSTOR_Vendor, USBSTOR_Product, DriveName, GUID, ParID)
        usb_report += f"""<span class="hasInfo" title="First Connect Date: {First_Connect}
Last Connect Date: {Last_Connect}
Serial Number: {SN}
Vendor ID: {VID}
Product ID: {PID}
Version: {Ver}
USBSTOR Vendor: {USBSTOR_Vendor}
USBSTOR Product: {USBSTOR_Product}
GUID: {GUID}
Parent Prefix ID: {ParID}"><b>Device: {Device}</b></span> <a target="_blank" style="text-decoration: none" href="https://devicehunt.com/view/type/usb/vendor/{VID}/device/{PID}">(Link)</a><br>
                    Vendor Name: {Vendor_Name}<br>
                    Product Description: {Product_Description}<br>
                    Drive Name: {DriveName}<br>"""
        len_dev = len_dev+1
        if total_dev > 1 and len_dev != total_dev:
            usb_report += "<br>"
    usb_report+="""</div>
                </div>"""
    return usb_report

# usb
def usb_report(show_verbose):
    report_usb = ""
    print(yellow, end='')
    print('[!] Fetching USB details')
    if show_verbose:
        st_time_usb = time.localtime()
        print(log_col,end="")
        print('[v] USB details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_usb)))
    try:
        report_usb = usb()
        print(green, end='')
        print('[√] USB details fetched successfully')
    except FileNotFoundError:
        report_usb = f"""<div class="reportSection rsLeft">
        <h2 class="reportSectionHeader">
            USB Storage Use
        </h2>
        <div class="reportSectionBody">
        No USB history found
        </div>
        </div>"""
        print(green, end='')
        print('[√] USB details fetched successfully')
    except Exception as e:
        report_usb = f"""<div class="reportSection rsLeft">
        <h2 class="reportSectionHeader">
            USB Storage Use
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching USB Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching USB details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_usb = time.localtime()
        print(log_col,end="")
        print('[v] USB details fetched at '+str(time.strftime("%H:%M:%S", ed_time_usb)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.fromtimestamp(time.mktime(ed_time_usb)) - datetime.fromtimestamp(time.mktime(st_time_usb))).split(":"))))
    return {'report_usb': report_usb}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")