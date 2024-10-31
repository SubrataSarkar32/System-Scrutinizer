# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Left section of the Section Seven
# Network details
def network_details():
    command_network = "(Get-NetAdapter).Name"
    result_network = subprocess.run(["powershell", command_network], shell=True, stdout=subprocess.PIPE)
    output_network = result_network.stdout.decode("utf-8")
    report_network=f"""<div class="reportSection rsLeft">
				<h2 class="reportSectionHeader">
					Communications
				</h2>
				<div class="reportSectionBody">
					<!-- SummaryNet => Summary -->
					<table role="presentation">
						<colgroup>
							<col style="width: 15%">
							<col style="width: 35%">
							<col style="width: 50%">
						</colgroup>"""
    inf_name=output_network.split("\n")
    inf_name=[i for i in inf_name if i]

    command_network_default = "(Get-NetIPConfiguration | Where-Object {$_.IPv4DefaultGateway -ne $null}).InterfaceAlias"
    result_network_default = subprocess.run(["powershell", command_network_default], shell=True, stdout=subprocess.PIPE)
    output_network_default = result_network_default.stdout.decode("utf-8")
    default=output_network_default.split("\n")
    def_inf=[]
    for i in range(len(default)):
        def_inf.append(default[i].strip())
    def_inf=[i for i in def_inf if i]

    name=""
    desc=""
    mac=""
    status=""
    speed=""
    ipv4=""
    gateway=""

    primary_items = []
    non_primary_items = []
    # Separate items into primary and non-primary lists
    for i in range(len(inf_name)):
        name = inf_name[i].strip()
        if name in def_inf:
            primary_items.append(name)
        else:
            non_primary_items.append(name)
    primary_items.sort()
    non_primary_items.sort()
    all_items = primary_items + non_primary_items

    len_name = 0
    for name in all_items:
        if name in primary_items:
            report_network+=f"""<tr>
							<td colspan="3"><span><b>"""
            if len_name >= 1:
                report_network += """\r<br>\r"""
            report_network+=f"""{name}</b> (<span style="color:#049304; font-weight:bold;">Primary</span>)</span>
							</td>
						</tr>"""
        else:
            report_network+=f"""<tr>
                <td colspan="3"><span><b>"""
            if len_name >= 1:
                report_network += """\r<br>\r"""
            report_network+=f"""{name}</span>
							</td>
						</tr>"""
        len_name = len_name+1
        command_network_status = "Get-NetAdapter | Where-Object {$_.Name -eq \""+name.strip()+"\"} | Select-Object InterfaceDescription, MacAddress, Status, LinkSpeed | Format-List"
        result_network_status = subprocess.run(["powershell", command_network_status], shell=True, stdout=subprocess.PIPE)
        output_network_status = result_network_status.stdout.decode("utf-8")
        inf_desc=output_network_status.split("\n")
        inf_desc = list(map(lambda x: x.replace('\r', ''), inf_desc))
        inf_desc=[i for i in inf_desc if i]

        command_network_ip = "Get-NetIPConfiguration | Where-Object {$_.InterfaceAlias -eq \""+name.strip()+"\"}"
        result_network_ip = subprocess.run(["powershell", command_network_ip], shell=True, stdout=subprocess.PIPE)
        output_network_ip = result_network_ip.stdout.decode("utf-8")
        inf_ip=output_network_ip.split("\n")
        inf_ip = list(map(lambda x: x.replace('\r', ''), inf_ip))
        inf_ip=[i for i in inf_ip if i]

        command_ipv6_IP = "if ($ipv6Addresses = (Get-NetIPAddress -InterfaceAlias \""+name.strip()+"\" -AddressFamily IPv6 -ErrorAction SilentlyContinue).IPAddress) { $ipv6Addresses } else { \"None\" }"
        result_ipv6_IP = subprocess.run(["powershell", command_ipv6_IP], shell=True, stdout=subprocess.PIPE)
        output_ipv6_IP = result_ipv6_IP.stdout.decode("utf-8").strip()
        
        for lines in inf_ip:
            if "IPv4Address          : " in lines:
                ipv4=lines[23:].strip()
            if "IPv4DefaultGateway   : " in lines:
                gateway=lines[23:].strip()
        
        for line in inf_desc:
            if "InterfaceDescription : " in line:
                desc=line[23:].strip()
            if "MacAddress           : " in line:
                mac=line[23:].strip()
            if "Status               : " in line:
                status=line[23:].strip()
            if "LinkSpeed            : " in line:
                speed=line[23:].strip()

        report_network+=f"""<tr>
							<td></td>
							<td>Interface Description:</td>
							<td>{desc}</td>
						</tr>
                        <tr>
							<td></td>
							<td>Mac Address:</td>
							<td>{mac}</td>
						</tr>
                        <tr>
							<td></td>
							<td>Status</td>
							<td>{status + str("&nbsp;&uarr;") if status.strip().lower() == "up" else (status + str("&nbsp;&darr;") if status.strip().lower() == "down" else status)}</td>
						</tr>
                        <tr>
							<td></td>
							<td>Link Speed:</td>
							<td>{speed}</td>
						</tr>
                        <tr>
							<td></td>
							<td>IPv4:</td>
							<td>{ipv4}</td>
						</tr>
                        <tr>
							<td></td>
							<td>IPv6:</td>
							<td>{output_ipv6_IP}</td>
						</tr>
                        <tr>
							<td></td>
							<td>Gateway:</td>
							<td>{gateway}</td>
						</tr>"""
    report_network+=f"""</table>
				</div>
			</div>"""
    report_network = report_network.encode('utf-8').decode('utf-8')
    return report_network

# Right section of the Section Seven
# Other devices details
def other_devices():
    report_other_devices=""
    command_otherDev = 'Get-PnpDevice -PresentOnly | Where-Object { $_.Class -eq "AudioEndpoint" -or $_.Class -eq "Battery" -or $_.Class -eq "Biometric" -or $_.Class -eq "Bluetooth" -or $_.Class -eq "Camera" -or $_.Class -eq "HIDClass" -or $_.Class -eq "Keyboard" -or $_.Class -eq "Mouse" -or $_.Class -eq "SecurityDevices" -or $_.Class -eq "USB" -or $_.Class -eq "MEDIA"} | Select-Object Class, FriendlyName | Sort-Object Class | Format-List'
    result_otherDev = subprocess.run(["powershell", command_otherDev], shell=True, stdout=subprocess.PIPE)
    output_otherDev = result_otherDev.stdout.decode("utf-8").strip()
    report_other_devices=f"""<div class="reportSection rsRight">
                    <h2 class="reportSectionHeader">
                        Other Devices
                    </h2>
                    <div class="reportSectionBody">"""
    # Class List
    class_list = []
    friendly_list = []
    seen = {}
    index_to_delete = []
    for line in output_otherDev.split("\n"):
        if "Class" in line.split(":")[0]:
            class_list.append(line.split(":")[1].strip())
        if "FriendlyName" in line.split(":")[0]:
            friendly_list.append(line.split(":")[1].strip())
    for i, num in enumerate(friendly_list):
        if num in seen:
            index_to_delete.append(i)
        else:
            seen[num] = True
    for i in sorted(index_to_delete, reverse=True):
        del class_list[i]
        del friendly_list[i]
    for i in range(len(class_list)):
        report_other_devices+=f"""{friendly_list[i] + " - <i> ( " + (class_list[i]).upper() +" )"} </i> <br>"""
    report_other_devices+=f"""</div>
                </div>"""
    return report_other_devices

# Network interface details
def network_report(show_verbose):
    report_network = ""
    print(yellow, end='')
    print('[!] Fetching Network details')
    if show_verbose:
        st_time_network = time.localtime()
        print(log_col,end="")
        print('[v] Network details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_network)))
    try:
        report_network = network_details()
        print(green, end='')
        print('[√] Network details fetched successfully')
    except Exception as e:
        report_network = f"""<div class="reportSection rsRight">
        <h2 class="reportSectionHeader">
            Communications
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching Network Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching Network details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_network = time.localtime()
        print(log_col,end="")
        print('[v] Network details fetched at '+str(time.strftime("%H:%M:%S", ed_time_network)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time_network)) - datetime.datetime.fromtimestamp(time.mktime(st_time_network))).split(":"))))
    return {'report_network': report_network}

#  Other devices
def other_devices_report(show_verbose):
    report_other_devices = ""
    print(yellow, end='')
    print('[!] Fetching Other Device details')
    if show_verbose:
        st_time_od = time.localtime()
        print(log_col,end="")
        print('[v] Other Device details fetching starting at '+str(time.strftime("%H:%M:%S", st_time_od)))
    try:
        report_other_devices = other_devices()
        print(green, end='')
        print('[√] Other Device details fetched successfully')
    except Exception as e:
        report_other_devices = f"""<div class="reportSection rsRight">
        <h2 class="reportSectionHeader">
            Other Devices
        </h2>
        <div class="reportSectionBody">
        <span style="font-size: 1.28em; color: #ff0000;">Error in fetching Other Device Details</span>
        </div>
        </div>"""
        print(red, end='')
        print('[×] Error in fetching Other Device details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time_od = time.localtime()
        print(log_col,end="")
        print('[v] Other Device details fetched at '+str(time.strftime("%H:%M:%S", ed_time_od)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time_od)) - datetime.datetime.fromtimestamp(time.mktime(st_time_od))).split(":"))))
    return {'report_other_devices': report_other_devices}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")