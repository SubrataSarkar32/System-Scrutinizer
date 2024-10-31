# -*- coding: utf-8 -*-
import subprocess
import time
from configuration import red, green, yellow, log_col, debug_mode
import datetime



# Left section of the Section Nine
# Hotfix details
def hotfix_details():
    report_hotfix=""
    command_hotfix = "Get-HotFix | Select-Object HotFixID, Description, InstalledOn | Format-List"
    result_hotfix = subprocess.run(["powershell", command_hotfix], shell=True, stdout=subprocess.PIPE)
    output_hotfix = result_hotfix.stdout.decode("utf-8")

    report_hotfix+="""<div class="reportSection rsLeft">
                <h2 class="reportSectionHeader">
					Hotfix
				</h2>
				<div class="reportSectionBody">
					<table role="presentation">
						<colgroup>
							<col style="width: 30%">
							<col style="width: 40%">
							<col style="width: 30%">
						</colgroup>
						<tr>
                            <th style="text-align: left">Version</th>
                            <th style="text-align: center">Description</th>
                            <th style="text-align: center">Install date</th>
                        </tr>"""

    hotfix_ver = ""
    hotfix_desc = ""
    hotfix_install_date = ""

    for line in output_hotfix.split("\n"):
        if "HotFixID" in line:
            hotfix_ver = line.split(":")[1].strip()
            # print(hotfix_ver)
            report_hotfix+=f"""<tr>
                        <td>{hotfix_ver}</td>"""
        elif "Description" in line:
            hotfix_desc = line.split(":")[1].strip()
            # print(hotfix_desc)
            report_hotfix+=f"""<td style="text-align: center">{hotfix_desc}</td>"""
        elif "InstalledOn" in line:
            hotfix_install_date = (line.split(":")[1].strip()).split(" ")[0]
            # print(hotfix_install_date)
            report_hotfix+=f"""<td style="text-align: center">{hotfix_install_date}</td>"""
    
    report_hotfix+="""</table>
                </div>
            </div>"""
    return report_hotfix


# Hotfix details
def hotfix_details_report(show_verbose):
    report_hotfix = ""
    print(yellow, end='')
    print('[!] Fetching Hotfix details')
    if show_verbose:
        st_time_hf = time.localtime()
        print(log_col,end="")
        print('[v] Hotfix details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_hf)))
    try:
        report_hotfix = hotfix_details()
        print(green, end='')
        print('[√] Hotfix details fetched successfully')
    except Exception as e:
        report_hotfix = f"""<div class="reportSection rsLeft">
        <h2 class="reportSectionHeader">
            Hotfix
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching Hotfix Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching Hotfix details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_hf = time.localtime()
        print(log_col,end="")
        print('[v] Hotfix details fetched at '+str(time.strftime("%H:%M:%S", ed_time_hf)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time_hf)) - datetime.datetime.fromtimestamp(time.mktime(st_time_hf))).split(":"))))
    return {'report_hotfix': report_hotfix}


if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")