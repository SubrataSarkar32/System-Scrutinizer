# -*- coding: utf-8 -*-
import subprocess
import datetime
import time
from configuration import red, green, yellow, log_col, debug_mode

# Left side of Section Five
# Antivirus details
def antivirus_details():
    report_virus=""
    report_virus="""<div class="reportSection rsLeft" id="antivirus">
				<h2 class="reportSectionHeader">
					Virus Protection
				</h2>
				<div class="reportSectionBody">
					<table role="presentation">"""
    
    # for windows defender

    commands_windows_defender_state =  "(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Where-Object {$_.displayName -eq \"Windows Defender\"}).productState"
    result_windows_defender_state = subprocess.run(["powershell", commands_windows_defender_state], shell=True, stdout=subprocess.PIPE)
    output_windows_defender_state = result_windows_defender_state.stdout.decode("utf-8").strip()

    # print("Windows Defender State")

    state=int(output_windows_defender_state)
    # print(state)
    state_hex=hex(state)
    digits=str(state_hex[-4:])

    final_state=""

    if digits[0]=="0":
        final_state="off"
    elif digits[0]=="1":
        final_state="on"
    elif digits[0]=="2":
        final_state="snoozed"
    elif digits[0]=="3":
        final_state="expired"
    else:
        final_state="unknown"

    # print(final_state)


    commands_win_def_details = "(Get-MpComputerStatus).AMEngineVersion;(Get-MpComputerStatus).AMProductVersion;(Get-MpComputerStatus).AntivirusSignatureLastUpdated -replace \"\\n\", \"\";(Get-MpComputerStatus).QuickScanStartTime -replace \"\\n\", \"\";(Get-MpComputerStatus).RealTimeProtectionEnabled"
    result_win_def_details= subprocess.run(["powershell", commands_win_def_details], shell=True, stdout=subprocess.PIPE)
    output_win_def_details= result_win_def_details.stdout.decode("utf-8").strip()

    # print("windows defender details")

    scan_eng_ver=""
    prod_ver=""
    def_lst_updt=""
    lst_scan=""
    rtp=""

    line=output_win_def_details.split("\n")
    scan_eng_ver=line[0].strip()
    prod_ver=line[1].strip()
    def_lst_updt=line[2].strip()
    lst_scan=line[3].strip()
    rtp=line[4].strip()
    # exclusions
    extension=""
    path=""
    process=""

    commands_win_def_excl_ext="(Get-MpPreference).ExclusionExtension"
    result_win_def_excl_ext=subprocess.run(["powershell", commands_win_def_excl_ext], shell=True, stdout=subprocess.PIPE)
    output_win_def_excl_ext= result_win_def_excl_ext.stdout.decode("utf-8").strip()

    extension = output_win_def_excl_ext

    commands_win_def_excl_path="(Get-MpPreference).ExclusionPath"
    result_win_def_excl_path=subprocess.run(["powershell", commands_win_def_excl_path], shell=True, stdout=subprocess.PIPE)
    output_win_def_excl_path= result_win_def_excl_path.stdout.decode("utf-8").strip()

    path=output_win_def_excl_path

    commands_win_def_excl_proc="(Get-MpPreference).ExclusionProcess"
    result_win_def_excl_proc=subprocess.run(["powershell", commands_win_def_excl_proc], shell=True, stdout=subprocess.PIPE)
    output_win_def_excl_proc= result_win_def_excl_proc.stdout.decode("utf-8").strip()

    process=output_win_def_excl_proc

    extension=extension.replace("\r","")
    if extension == "":
        extension = None
    path=path.replace("\r","")
    if path == "":
        path = None
    process=process.replace("\r","")
    if process == "":
        process = None

    report_virus+=f"""<tr>
							<td class="hasInfo" title="Exclusion Extension: 
{extension}
Exclusion Path: 
{path}
Exclusion Process: 
{process}
State: {final_state.capitalize()}"><b>Windows Defender</b> Version {prod_ver}</td>
						</tr>
						<tr>
							<td>&nbsp;&nbsp;&nbsp;&nbsp;Scan Engine Version {scan_eng_ver}</td>
						</tr>
						<tr>
							<td>&nbsp;&nbsp;&nbsp;&nbsp;Virus Definitions Version {def_lst_updt}</td>
						</tr>
						<tr>
							<td>&nbsp;&nbsp;&nbsp;&nbsp;Last Disk Scan on {lst_scan}</td>
						</tr>
						<tr>
							<td>&nbsp;&nbsp;&nbsp;&nbsp;Realtime File Scanning {rtp}</td>
						</tr>"""

    # for other antivirus versions
    command_other_antivirus = "(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Where-Object {$_.displayName -ne \"Windows Defender\"}).displayName"
    result_other_antivirus = subprocess.run(["powershell", command_other_antivirus], shell=True, stdout=subprocess.PIPE)
    output_other_antivirus = result_other_antivirus.stdout.decode("utf-8").strip()

    name=[]

    name=output_other_antivirus.split("\n")

    # to remove blank values
    name=[i for i in name if i]

    # print(len(name))
    if (len(name)!=0):
        for antivirus in name:

            # print(antivirus)

            command_prdct_state="(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Where-Object {$_.displayName -like \""+antivirus+"\"}).productState"
            result_prdct_state = subprocess.run(["powershell", command_prdct_state], shell=True, stdout=subprocess.PIPE)
            output_prdct_state = result_prdct_state.stdout.decode("utf-8").strip()
            
            state=int(output_prdct_state)
            # print(state)
            state_hex=hex(state)
            digits=str(state_hex[-4:])

            final_state_oav=""

            if digits[0]=="0":
                final_state_oav="off"
            elif digits[0]=="1":
                final_state_oav="on"
            elif digits[0]=="2":
                final_state_oav="snoozed"
            elif digits[0]=="3":
                final_state_oav="expired"
            else:
                final_state_oav="unknown"
            # print(output_prdct_state)

            command_av_ver_dt="Get-WmiObject Win32_Product | select Name, Version, InstallDate | where {$_.Name -like \""+antivirus+"\"} | Format-List"
            result_av_ver_dt = subprocess.run(["powershell", command_av_ver_dt], shell=True, stdout=subprocess.PIPE)
            output_av_ver_dt = result_av_ver_dt.stdout.decode("utf-8").strip()

            # print(output_av_ver_dt)

            version=""
            date=""

            for line in output_av_ver_dt.split("\n"):
                if "Version" in line:
                    version=line.split(":")[1].strip()
                if "InstallDate" in line:
                    date=line.split(":")[1].strip()
            report_virus+=f"""<tr>
							<td><b>{antivirus}</b> Version {version}</td>
						</tr>
						<tr>
							<td>&nbsp;&nbsp;&nbsp;&nbsp;Antivirus State {final_state_oav}</td>
						</tr>"""
            if date !="":
                date = datetime.datetime.strptime(date, "%Y%m%d").strftime("%d %B %y")
                report_virus+=f"""<tr>
                                <td>&nbsp;&nbsp;&nbsp;&nbsp;Install Date {date}</td>
                            </tr>"""

    report_virus+="""</table>
				</div>
			</div>"""
    return report_virus

# Right side of Section Five
# Display details
def display_details():
    report_display=""
    commands_display =  "(Get-PnpDevice -Class Display).Caption"
    result_display = subprocess.run(["powershell", commands_display], shell=True, stdout=subprocess.PIPE)
    output_display = result_display.stdout.decode("utf-8")
    report_display+="""<div class="reportSection rsRight">
				<h2 class="reportSectionHeader">
					Display
				</h2>
				<div class="reportSectionBody">"""
    for i in output_display.strip().split("\n") :
        report_display+=i+"""<br>"""
    report_display+="""</div>
			</div>"""
    return report_display

# Antivirus Details
def antivirus_details_Report(show_verbose):
    report_virus = ""
    print(yellow, end='')
    print('[!] Fetching Antivirus details')
    if show_verbose:
        st_time_av = time.localtime()
        print(log_col,end="")
        print('[v] Antivirus details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_av)))
    try:
        report_virus = antivirus_details()
        print(green, end='')
        print('[√] Antivirus details fetched successfully')
    except Exception as e:
        report_virus = f"""<div class="reportSection rsLeft">
        <h2 class="reportSectionHeader">
            Virus Protection
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching Antivirus Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching Antivirus details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_av = time.localtime()
        print(log_col,end="")
        print('[v] Antivirus details fetched at '+str(time.strftime("%H:%M:%S", ed_time_av)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time_av)) - datetime.datetime.fromtimestamp(time.mktime(st_time_av))).split(":"))))
    return {'report_virus': report_virus}

# Display Details
def display_details_Report(show_verbose):
    report_display = ""
    print(yellow, end='')
    print('[!] Fetching Display details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Display details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_display = display_details()
        print(green, end='')
        print('[√] Display details fetched successfully')
    except Exception as e:
        report_display = f"""<div class="reportSection rsRight">
        <h2 class="reportSectionHeader">
            Displays
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching Display Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching Display details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Display details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_display': report_display}


if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")