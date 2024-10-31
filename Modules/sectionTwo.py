# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Left of Second Section
# Get Processor Details
def processor_detail():
    report_pd=""
    command_PD = "Get-WmiObject -Class Win32_Processor | Select-Object Caption, Manufacturer, MaxClockSpeed, Name, SocketDesignation"
    result_PD = subprocess.run(
        ["powershell", command_PD], shell=True, stdout=subprocess.PIPE)
    output_PD = result_PD.stdout.decode("utf-8")
    p_name = ""
    p_cap = ""
    p_manuf = ""
    p_clock = ""
    p_socdesig = ""
    for line in output_PD.split("\n"):
        if "Caption" in line:
            p_cap = line.split(":")[1].strip()
        elif "Manufacturer" in line:
            p_manuf = line.split(":")[1].strip()
        elif "MaxClockSpeed" in line:
            p_clock = line.split(":")[1].strip()
        elif "Name" in line:
            p_name = line.split(":")[1].strip()
        elif "SocketDesignation" in line:
            p_socdesig = line.split(":")[1].strip()
    # Adding to the report
    report_pd += f"""<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Processor
                    </h2>
                    <div class="reportSectionBody">
                        Name: {p_name}<br>Caption: {p_cap}<br>Manufacturer: {p_manuf}<br>Max Clock Speed: {p_clock}<br>Socket Designation: {p_socdesig}
                    </div>
                </div>"""
    return report_pd

# Right of Second Section
# Get Physical Drive details
def drives_detail():
    report_dd=""
    command_DD = "Get-PhysicalDisk | Select-Object DeviceId, FriendlyName, SerialNumber, MediaType, HealthStatus, Size | Sort-Object -Property DeviceId"
    result_DD = subprocess.run(
        ["powershell", command_DD], shell=True, stdout=subprocess.PIPE)
    output_DD = result_DD.stdout.decode("utf-8")
    d_id = ""
    d_name = ""
    d_serial = ""
    d_media = ""
    d_health = ""
    d_size = ""
    # Adding to the report
    report_dd += """<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        Drives
                    </h2>
                    <div class="reportSectionBody">
                        <table role="presentation">
                            <colgroup>
                                <col style="width: 5%">
                                <col style="width: 35%">
                                <col style="width: 20%">
                                <col style="width: 15%">
                                <col style="width: 15%">
                                <col style="width: 10%">
                            </colgroup>
                            <tr>
                                <th>ID</th>
                                <th style="text-align: center">Name</th>
                                <th style="text-align: center">Serial Number</th>
                                <th style="text-align: center">Media Type</th>
                                <th style="text-align: center">Health Status</th>
                                <th style="text-align: center">Total Size</th>
                            </tr>"""
    for line in output_DD.split("\n"):
        if "DeviceId" in line:
            d_id = line.split(":")[1].strip()
            report_dd += f"""<tr>
                        <td>{d_id}</td>"""
        elif "FriendlyName" in line:
            d_name = line.split(":")[1].strip()
            report_dd += f"""<td>{d_name}</td>"""
        elif "SerialNumber" in line:
            d_serial = line.split(":")[1].strip()
            report_dd += f"""<td>{d_serial}</td>"""
        elif "MediaType" in line:
            d_media = line.split(":")[1].strip()
            report_dd += f"""<td style="text-align: center">{d_media}</td>"""
        elif "HealthStatus" in line:
            d_health = line.split(":")[1].strip()
            report_dd += f"""<td style="text-align: center">{d_health}</td>"""
        elif "Size" in line:
            d_size = line.split(":")[1].strip()
            d_size = round(int(d_size)/(1024*1024*1024), 2)
            report_dd += f"""<td>{d_size} GB</td>
            <td></td>
            <td><span></span></td>
            </tr>"""
    report_dd += """</table>
            </div>
            </div>"""
    return report_dd

# ---Processor---
def processor_Report(show_verbose):
    report_pd = ""
    print(yellow, end='')
    print('[!] Getting Processor information')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Processor information gathering starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_pd = processor_detail()
        print(green, end='')
        print('[√] Processor information gathered successfully')
    except Exception as e:
        report_pd = f"""<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Processor
                    </h2>
                    <div class="reportSectionBody">
                        <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Processor Details</span>
                    </div>
                </div>"""
        print(red, end='')
        print('[×] Error in gathering Processor information')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Processor information gathering finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_pd': report_pd}

# Physical Drives
def driver_Report(show_verbose):
    report_dd = ""
    print(yellow, end='')
    print('[!] Getting Physical Drives information')
    if show_verbose:
        st_timedd = time.localtime()
        print(log_col,end="")
        print('[v] Physical Drives information gathering starting at '+str(time.strftime("%H:%M:%S", st_timedd)))
    try:
        report_dd = drives_detail()
        print(green, end='')
        print('[√] Physical Drives information gathered successfully')
    except Exception as e:
        report_dd = f"""<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        Drives
                    </h2>
                    <div class="reportSectionBody">
                        <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Drives Details</span>
                    </div>
                </div>"""
        print(red, end='')
        print('[×] Error in gathering Physical Drives information')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Physical Drives information gathering finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_timedd))).split(":"))))
    return {'report_dd': report_dd}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")