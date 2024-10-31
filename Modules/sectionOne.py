# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Left of First Section
# Get OS Details
def os_details():
    report_os=""
    command_OS = '(wmic os get caption /value | findstr "=").Substring((wmic os get caption /value | findstr "=").IndexOf("=")+1); (wmic os get buildnumber /value | findstr "=").Substring((wmic os get buildnumber /value | findstr "=").IndexOf("=")+1); (wmic os get version /value | findstr "=").Substring((wmic os get version /value | findstr "=").IndexOf("=")+1); (wmic os get muilanguages /value | findstr "=").Substring((wmic os get muilanguages /value | findstr "=").IndexOf("=")+1)'
    result_OS = subprocess.run(
        ["powershell", command_OS], shell=True, stdout=subprocess.PIPE)
    output_OS = result_OS.stdout.decode("utf-8")
    # Get Original Install Date
    command_idt = "(Get-CimInstance -ClassName Win32_OperatingSystem).InstallDate"
    result_idt = subprocess.run(
        ["powershell", command_idt], shell=True, stdout=subprocess.PIPE)
    output_idt = result_idt.stdout.decode('utf-8')
    os_name = ""
    os_version = ""
    sys_locale = ""
    install_date = output_idt
    line=output_OS.split("\n")
    os_name = line[0].strip()
    os_version = str(line[1].strip())+" "+str(line[2].strip())
    sys_locale = ((line[3].strip())[:-1])[1:]
    # Get Secure Boot details
    command_SB = "Confirm-SecureBootUEFI"
    result_SB = subprocess.run(
        ["powershell", command_SB], shell=True, stdout=subprocess.PIPE)
    output_SB = result_SB.stdout.decode("utf-8")
    # Adding to the report
    report_os += f"""<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Operating System
                    </h2>
                    <div class="reportSectionBody">
                        {os_name} {os_version}<br>System Locale: {sys_locale}<br>Installed: {install_date}<br>Boot Mode: """
    if (output_SB.strip() == "True"):
        report_os += """UEFI with Secure Boot <span
                            style="color:#049304; font-weight:bold;">enabled</span>"""
    elif (output_SB.strip() == "False"):
        report_os += """UEFI with Secure Boot <span
                                style="color:#ff0000; font-weight:bold;">disabled</span>"""
    else:
        report_os += """<span
                        style="color:#ff0000; font-weight:bold;">Failed to Determined</span>"""
            
    command_win = "cscript.exe C:\Windows\System32\slmgr.vbs /dlv"
    result_win = subprocess.run(["powershell", command_win], shell=True, stdout=subprocess.PIPE)
    output_win = result_win.stdout.decode('utf-8').strip()
    li_status = None
    key_chanel = ""
    for line in output_win.split('\n'):
        if "License Status: " in line:
            li_status = line.split(":")[1].strip()
        if "Product Key Channel: " in line:
            key_chanel = line[21:].strip()
    if li_status == "Licensed":
        report_os += f"""<br>Windows Licensed: <span style="color:#049304; font-weight:bold;">Genuine Windows</span>"""
    else:
        report_os += f"""<br>Windows Licensed: <span style="color:#ff0000; font-weight:bold;">{li_status}</span>"""
    if li_status == None:
        report_os += f"""<br>Windows Licensed: <span style="color:#ff0000; font-weight:bold;">Failed to verify</span>"""
    report_os += f"""<br>Product Key Chanel: {key_chanel}"""
    report_os += f"""</div>
            </div>"""
    return report_os

# Right of First Section
# Get System Model
def system_model():
    report_sm=""
    command_sm = "wmic csproduct get vendor; wmic csproduct get Name; wmic diskdrive get serialnumber; wmic bios get serialnumber"
    result_sm = subprocess.run(
        ["powershell", command_sm], shell=True, stdout=subprocess.PIPE)
    output_sm = result_sm.stdout.decode('utf-8')
    sys_manufacturer = (output_sm.split("\n")[1])
    sys_model = (output_sm.split("\n")[4])
    sys_serialNo = (output_sm.split("\n")[7])
    sys_bios_serialNo = (output_sm.split("\n")[11])
    ChassisTypes_DESC = {1: 'Other', 2: 'Unknown', 3: 'Desktop', 4: 'Low Profile Desktop', 5: 'Pizza Box', 6: 'Mini Tower', 7: 'Tower', 8: 'Portable', 9: 'Laptop', 10: 'Notebook', 11: 'Hand Held', 12: 'Docking Station', 13: 'All in One', 14: 'Sub Notebook', 15: 'Space-Saving', 16: 'Lunch Box', 17: 'Main System Chassis', 18: 'Expansion Chassis', 19: 'SubChassis', 20: 'Bus Expansion Chassis', 21: 'Peripheral Chassis', 22: 'Storage Chassis', 23: 'Rack Mount Chassis', 24: 'Sealed-Case PC', 30: 'Tablet', 31: 'Convertible', 32: 'Detachable'}
    ChassisTypes_no_cmd = "Get-CimInstance -ClassName Win32_SystemEnclosure -Property * | Select-Object ChassisTypes"
    ChassisTypes_result = subprocess.run(
        ["powershell", ChassisTypes_no_cmd], shell=True, stdout=subprocess.PIPE)
    ChassisTypes_output = ChassisTypes_result.stdout.decode('utf-8')
    ChassisTypes_ID = ChassisTypes_output.split('\n')[3].strip().strip('{}')
    ChassisTypes_ID_DESC = ChassisTypes_DESC.get(int(ChassisTypes_ID))
    # UEFI version
    uefi_version_cmd = "Get-WmiObject -Class Win32_BIOS | Select-Object Name; $bios = Get-WmiObject -Class Win32_BIOS; $update = $bios.ReleaseDate; echo $update"
    uefi_version_result = subprocess.run(
        ["powershell", uefi_version_cmd], shell=True, stdout=subprocess.PIPE)
    uefi_version_output = uefi_version_result.stdout.decode('utf-8')
    year = uefi_version_output.split('\n')[4].strip()[:4]
    month = uefi_version_output.split('\n')[4].strip()[4:6]
    day = uefi_version_output.split('\n')[4].strip()[6:8]
    uefi_version = uefi_version_output.split('\n')[3]
    last_update = day+'/'+month+'/'+year
    #TPM vesion
    tpm_version_cmd= "(Get-WmiObject -Namespace \'root\\cimv2\\security\\microsofttpm\' -class win32_tpm).SpecVersion"
    tpm_version_result = subprocess.run(["powershell", tpm_version_cmd], shell=True, stdout=subprocess.PIPE)
    tpm_version_output = tpm_version_result.stdout.decode('utf-8')
    # Adding to the report
    report_sm += f"""<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        System Model
                    </h2>
                    <div class="reportSectionBody">
                        System Model: {sys_model}<br>System Manufacturer: {sys_manufacturer}<br>Serial Number: {sys_serialNo}<br>BIOS Serial Number: {sys_bios_serialNo}
                        <br>Chassis Type: {ChassisTypes_ID_DESC}<br>UEFI version: {uefi_version}<br>UEFI Last Update: {last_update}<br>"""
    if (tpm_version_output!=""):
        tpm_ver=float(tpm_version_output.split(",")[0])
        if(tpm_ver>=2.0):
            report_sm += f"""TPM version: {tpm_ver} <span style="color:#049304; font-weight:bold;">(Secured)</span><br>"""
        else:
            report_sm += f"""TPM version: {tpm_ver} <span style="color:#ff0000; font-weight:bold;">(Not Secured)</span><br>"""
    else:
        report_sm += f"""TPM Not Enabled. <span style="color:#ff0000; font-weight:bold;">(Not Secured)</span><br>"""
    report_sm += f"""</div>
                </div>"""
    return report_sm


# ---Operating System---
def osdetails_Report(show_verbose):
    report_os = ""
    print(yellow, end='')
    print('[!] Getting OS information')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] OS information gathering starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_os = os_details()
        print(green, end='')
        print('[√] OS information gathered successfully')
    except Exception as e:
        report_os = f"""<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Operating System
                    </h2>
                    <div class="reportSectionBody">
                        <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching OS Details</span>
                    </div>
                </div>"""
        print(red, end='')
        print('[×] Error in gathering OS information')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] OS information gathering finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_os': report_os}


# ---System Model---
def system_model_Report(show_verbose):
    report_sm = ""
    print(yellow, end='')
    print('[!] Getting System information')
    if show_verbose:
        st_timeSM = time.localtime()
        print(log_col,end="")
        print('[v] System information gathering starting at '+str(time.strftime("%H:%M:%S", st_timeSM)))
    try:
        report_sm = system_model()
        print(green, end='')
        print('[√] System information gathered successfully')
    except Exception as e:
        report_sm = f"""<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        System Model
                    </h2>
                    <div class="reportSectionBody">
                        <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching System Model Details</span>
                    </div>
                </div>"""
        print(red, end='')
        print('[×] Error in gathering System information')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] System information gathering finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_timeSM))).split(":"))))
    return {'report_sm': report_sm}


if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")