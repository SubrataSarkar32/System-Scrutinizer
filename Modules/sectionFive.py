# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Left side of section Five
# Controller details
def controller_details():
    report_controller=""
    commands_controllers =  "Get-WmiObject -Class Win32_IDEController | Select-Object DeviceID | ForEach-Object {Get-PnpDevice -InstanceId $_.DeviceID | Select-Object -ExpandProperty FriendlyName}; (Get-CimInstance -ClassName Win32_SCSIController).Name; (Get-CimInstance -ClassName Win32_USBController).Caption"
    result_controllers = subprocess.run(["powershell", commands_controllers], shell=True, stdout=subprocess.PIPE)
    output_controllers = result_controllers.stdout.decode("utf-8")
    report_controller+="""<div class="reportSection rsLeft">
				<h2 class="reportSectionHeader">
					Controllers
				</h2>
				<div class="reportSectionBody">"""
    for i in output_controllers.strip().split("\n") :
        report_controller+=i+"""<br>"""
    report_controller+="""</div>
			</div>"""
    return report_controller

# Right side of section Five
# Audio details
def audio_details():
    report_audio=""
    command_Audio_drive = "Get-CimInstance -ClassName Win32_SoundDevice | Select-Object Manufacturer, Name, Status | Format-List"
    result_Audio_drive = subprocess.run(["powershell", command_Audio_drive], shell=True, stdout=subprocess.PIPE)
    output_Audio_drive = result_Audio_drive.stdout.decode("utf-8")
    report_audio+="""<div class="reportSection rsRight">
                <h2 class="reportSectionHeader">
					Multimedia
				</h2>
				<div class="reportSectionBody">
					<table role="presentation">
						<colgroup>
							<col style="width: 30%">
							<col style="width: 40%">
							<col style="width: 30%">
						</colgroup>
						<tr>
                            <th style="text-align: left">Manufacturer</th>
                            <th style="text-align: center">Driver Name</th>
                            <th style="text-align: center">Status</th>
                        </tr>"""
    Audio_drive_manuf = ""
    Audio_drive_name = ""
    Audio_drive_status = ""
    for line in output_Audio_drive.split("\n"):
        if "Manufacturer" in line:
            Audio_drive_manuf = line.split(":")[1].strip()
            # print(Audio_drive_manuf)
            report_audio+=f"""<tr>
                        <td>{Audio_drive_manuf}</td>"""
        elif "Name" in line:
            Audio_drive_name = line.split(":")[1].strip()
            # print(Audio_drive_name)
            report_audio+=f"""<td>{Audio_drive_name}</td>"""
        elif "Status" in line:
            Audio_drive_status = line.split(":")[1].strip()
            # print(Audio_drive_status)
            report_audio+=f"""<td style="text-align: center">{Audio_drive_status}</td>"""
    report_audio+="""</table>
                </div>
            </div>"""
    return report_audio

# Controller details
def controller_details_Report(show_verbose):
    report_controller = ""
    print(yellow, end='')
    print('[!] Fetching Controller details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Controller details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_controller = controller_details()
        print(green, end='')
        print('[√] Controller details fetched successfully')
    except Exception as e:
        report_controller = f"""<div class="reportSection rsLeft">
            <h2 class="reportSectionHeader">
                Controllers
            </h2>
            <div class="reportSectionBody">
                <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Controller Details</span>
            </div>
        </div>"""
        print(red, end='')
        print('[×] Error in Fetching Controller details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Controller details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_controller': report_controller}


# Audio Details
def audio_details_Report(show_verbose):
    report_audio = ""
    print(yellow, end='')
    print('[!] Fetching Audio device details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Audio device details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_audio = audio_details()
        print(green, end='')
        print('[√] Audio device details fetched successfully')
    except Exception as e:
        report_audio = f"""<div class="reportSection rsRight">
            <h2 class="reportSectionHeader">
                Multimedia
            </h2>
            <div class="reportSectionBody">
                <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Multimedia Details</span>
            </div>
        </div>"""
        print(red, end='')
        print('[×] Error in Fetching Audio device details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Audio device details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_audio': report_audio}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")