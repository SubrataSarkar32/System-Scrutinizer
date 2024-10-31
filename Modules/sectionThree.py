# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Section Three left side
# Get Memory module details
def memory_detail():
    report_mem=""
    slot_count=0
    report_mem += """<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Memory Details
                    </h2>
                    <div class="reportSectionBody">"""
    # Usable Memory
    command_mem = "(Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory/1GB"
    result_mem = subprocess.run(["powershell", command_mem], shell=True, stdout=subprocess.PIPE)
    output_mem = result_mem.stdout.decode("utf-8")
    totalUsable_mem = str(round(float(output_mem.strip()), 2))
    report_mem+=f"""{totalUsable_mem} Gigabytes Usable Installed Memory.<br><br>"""
    report_mem+="""<table role="presentation">
                            <colgroup>
                                <col style="width: 35%">
                                <col style="width: 20%">
                                <col style="width: 15%">
                                <col style="width: 15%">
                                <col style="width: 15%">
                            </colgroup>
                            <tr>
                                <th style="text-align: left">Manufacturer</th>
                                <th style="text-align: center">Clock Speed</th>
                                <th style="text-align: center">Channel</th>
                                <th style="text-align: center">Capacity</th>
                                <th style="text-align: center">Serial Number</th>
                            </tr>"""
    # Memory Module Details
    command_mem = "Get-WmiObject win32_physicalmemory | Select-Object Manufacturer, Configuredclockspeed, Devicelocator,Capacity,Serialnumber | Format-List"
    result_mem = subprocess.run(["powershell", command_mem], shell=True, stdout=subprocess.PIPE)
    output_mem = result_mem.stdout.decode("utf-8")
    command_mem_count = "(Get-WmiObject -Class Win32_PhysicalMemoryArray).MemoryDevices"
    result_mem_count = subprocess.run(["powershell", command_mem_count], shell=True, stdout=subprocess.PIPE)
    output_mem_count = result_mem_count.stdout.decode("utf-8")
    mem_manuf=""
    mem_speed=""
    mem_device=""
    mem_capacity=""
    mem_serial=""
    total_mem_slots=output_mem_count.strip()
    for line in output_mem.split("\n"):
        if "Manufacturer" in line:
            mem_manuf = line.split(":")[1].strip()
            slot_count+=1
            report_mem+=f"""<tr>
                        <td style="text-align: left">{mem_manuf}</td>"""
        elif "Configuredclockspeed" in line:
            mem_speed = line.split(":")[1].strip()
            report_mem+=f"""<td style="text-align: center">{mem_speed}</td>"""
        elif "Devicelocator" in line:
            mem_device = line.split(":")[1].strip()
            report_mem+=f"""<td style="text-align: center">{mem_device}</td>"""
        elif "Capacity" in line:
            mem_capacity = line.split(":")[1].strip()
            mem_capacity_conv = (int(mem_capacity)/1024/1024/1024)
            report_mem+=f"""<td style="text-align: center">{mem_capacity_conv} GB</td>"""
        elif "Serialnumber" in line:
            mem_serial = line.split(":")[1].strip()
            if mem_serial== "00000000":
                report_mem+=f"""<td style="text-align: center">N/A</td><td></td>
            <td><span></span></td>
            </tr>"""
            else:
                report_mem+=f"""<td style="text-align: center">{mem_serial}</td><td></td>
            <td><span></span></td>
            </tr>"""
    report_mem += """</table>"""
    # Total system memory capacity
    command_mem_Total = "Get-WmiObject Win32_PhysicalMemoryArray | Where {$_.Use -eq 3} | Foreach {($_.MaxCapacity*1KB)/1GB}"
    result_mem_Total = subprocess.run(["powershell", command_mem_Total], shell=True, stdout=subprocess.PIPE)
    output_mem_Total = (result_mem_Total.stdout.decode("utf-8")).strip()
    report_mem+=f"""<br>{output_mem_Total} Gigabytes of Maximum System Memory Capacity."""
    report_mem+=f"""<br>{slot_count} out of {total_mem_slots} memory slots used."""
    report_mem += """</div>
                </div>"""
    return report_mem

# Right of Third Section
# Get Local Drive Volumes
def local_drive_detail():
    report_l_drive=""
    report_l_drive += """<div class="reportSection rsRight">
                <h2 class="reportSectionHeader">
					Local Drive Volumes
				</h2>
				<div class="reportSectionBody">
					<table role="presentation">
						<colgroup>
							<col style="width: 10%">
							<col style="width: 45%">
							<col style="width: 15%">
							<col style="width: 15%">
							<col style="width: 15%">
						</colgroup>
						<tr>
                            <th style="text-align: left">ID</th>
                            <th style="text-align: center">Volume Name</th>
                            <th style="text-align: center">File System</th>
                            <th style="text-align: center">Total Size</th>
                            <th style="text-align: center">Free Space</th>
                        </tr>"""
    # Local Drive Details
    command_l_drive = "Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace | Format-List"
    result_l_drive = subprocess.run(["powershell", command_l_drive], shell=True, stdout=subprocess.PIPE)
    output_l_drive = result_l_drive.stdout.decode("utf-8")
    l_drive_DeviceID = ""
    l_drive_VolumeName = ""
    l_drive_FileSystem = ""
    l_drive_Size = ""
    l_drive_FreeSpace = ""
    for line in output_l_drive.split("\n"):
        if "DeviceID" in line:
            l_drive_DeviceID = line.split(":")[1].strip()
            report_l_drive+=f"""<tr>
                        <td>{l_drive_DeviceID}</td>"""
        elif "VolumeName" in line:
            l_drive_VolumeName = line.split(":")[1].strip()
            if l_drive_VolumeName == "":
                l_drive_VolumeName  = "N/A"
            report_l_drive+=f"""<td>{l_drive_VolumeName}</td>"""
        elif "FileSystem" in line:
            l_drive_FileSystem = line.split(":")[1].strip()
            if l_drive_FreeSpace != "":
                report_l_drive+=f"""<td style="text-align: center">{l_drive_FileSystem}</td>"""
            else:
                report_l_drive+=f"""<td style="text-align: center">N/A</td>"""
        elif "Size" in line:
            l_drive_Size = line.split(":")[1].strip()
            if l_drive_Size != "":
                l_drive_Size_conv = round((int(l_drive_Size)/1024/1024/1024),2)
                report_l_drive+=f"""<td style="text-align: center">{l_drive_Size_conv} GB</td>"""
            else:
                l_drive_Size_conv = "N/A"
                report_l_drive+=f"""<td style="text-align: center">{l_drive_Size_conv}</td>"""
        elif "FreeSpace" in line:
            l_drive_FreeSpace = line.split(":")[1].strip()
            if l_drive_FreeSpace != "":
                l_drive_FreeSpace_conv = round((int(l_drive_FreeSpace)/1024/1024/1024),2)
                report_l_drive+=f"""<td style="text-align: center">{l_drive_FreeSpace_conv} GB</td><td></td><td><span></span></td></tr>"""
            else:
                l_drive_FreeSpace_conv = "N/A"
                report_l_drive+=f"""<td style="text-align: center">{l_drive_FreeSpace_conv}</td><td></td>
                <td><span></span></td>
                </tr>"""
    report_l_drive += """</table>"""
    # Where windows is installed
    command_winLoc = "(Get-CimInstance -ClassName Win32_OperatingSystem).SystemDrive"
    result_winLoc = subprocess.run(["powershell", command_winLoc], shell=True, stdout=subprocess.PIPE)
    output_winLoc = result_winLoc.stdout.decode("utf-8").strip()
    report_l_drive+=f"""<br>Operating System is installed on {output_winLoc}.
            </div>
            </div>"""
    return report_l_drive



# Memory modules
def memory_Report(show_verbose):
    report_mem = ""
    print(yellow, end='')
    print('[!] Fetching Memory Modules information')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Memory Modules information fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_mem = memory_detail()
        print(green, end='')
        print('[√] Memory modules information fetched successfully')
    except Exception as e:
        report_mem = f"""<div class="reportSection rsLeft">
                    <h2 class="reportSectionHeader">
                        Memory Details
                    </h2>
                    <div class="reportSectionBody">
                        <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Memory Details</span>
                    </div>
                </div>"""
        print(red, end='')
        print('[×] Error in Fetching Memory Modules information')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Memory Modules information fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_mem': report_mem}

# Local drive volumes
def local_Driver_Report(show_verbose):
    report_l_drive = ""
    print(yellow, end='')
    print('[!] Fetching Local Drive details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Local Drive details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_l_drive = local_drive_detail()
        print(green, end='')
        print('[√] Local Drive details fetched successfully')
    except Exception as e:
        report_l_drive = f"""<div class="reportSection rsRight">
            <h2 class="reportSectionHeader">
                Local Drive Volumes
            </h2>
            <div class="reportSectionBody">
                <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Local Drive Volumes Details</span>
            </div>
        </div>"""
        print(red, end='')
        print('[×] Error in Fetching Local Drive details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Local Drive details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_l_drive': report_l_drive}


if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")