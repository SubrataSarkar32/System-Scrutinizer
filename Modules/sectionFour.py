# -*- coding: utf-8 -*-
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode

# Left side of section four
# Get User Details
def user_details():
    report_user=""
    report_user+="""<div class="reportSection rsLeft">
				<h2 class="reportSectionHeader">
					Users
					<span class="rshNote">(mouse over user name for details)</span>
				</h2>
				<div class="reportSectionBody">
					<!-- SummaryLogins => Summary -->
					<table>
						<tr>"""
    # local user accounts
    command_user_name ="(Get-LocalUser | Where-Object {$_.Enabled -eq $true}).Name"
    result_user_name = subprocess.run(['powershell.exe', command_user_name], capture_output=True)
    output_user_name=result_user_name.stdout.decode()
    users=output_user_name.strip().split("\n")
    full_name=""
    status=""
    last_logon=""
    sid=""
    logons=""
    pass_age=""
    priv=""
    report_user+="""<th colspan="2" class="hasInfo"
								title="User accounts defined locally on this computer">local user
								accounts</th>
							<th>last logon</th>
						</tr>"""
    for i in range(len(users)):
        # full name
        command_user_fullname ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).FullName"
        result_user_fullname = subprocess.run(['powershell.exe', command_user_fullname], capture_output=True)
        output_user_fullname = result_user_fullname.stdout.decode()
        full_name=users[i].strip()+"-"+output_user_fullname.strip()
        # status
        command_user_status ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).Enabled"
        result_user_status = subprocess.run(['powershell.exe', command_user_status], capture_output=True)
        output_user_status = result_user_status.stdout.decode()
        status=output_user_status.strip()
        # last logon
        command_user_last_logon ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).LastLogon"
        result_user_last_logon = subprocess.run(['powershell.exe', command_user_last_logon], capture_output=True)
        output_user_last_logon = result_user_last_logon.stdout.decode().strip()
        if (output_user_last_logon==""):
            last_logon="never"
        else:
            last_logon=output_user_last_logon
        # SID
        command_user_SID ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).SID.Value"
        result_user_SID = subprocess.run(['powershell.exe', command_user_SID], capture_output=True)
        output_user_SID = result_user_SID.stdout.decode()
        sid=output_user_SID.strip()
        # no of logons
        command_user_logon="(Get-EventLog -LogName Security -InstanceId 4624 -ErrorAction SilentlyContinue | Where-Object {$_.ReplacementStrings[5] -eq '"+users[i].strip()+"'}).Count"
        result_user_logon = subprocess.run(['powershell.exe', command_user_logon], capture_output=True)
        output_user_logon = result_user_logon.stdout.decode()
        logons=output_user_logon.strip()
        # password age
        command_user_pass_age="(New-TimeSpan -Start (Get-LocalUser -Name '"+users[i].strip()+"').PasswordLastSet -End (Get-Date)).Days"
        result_user_pass_age = subprocess.run(['powershell.exe', command_user_pass_age], capture_output=True)
        output_user_pass_age = result_user_pass_age.stdout.decode().strip()
        if output_user_pass_age=="":
            pass_age="0"
        else:
            pass_age=output_user_pass_age.strip()
        # Privilege Level
        # Administrators
        command_admin ="(Get-LocalGroupMember -Name 'Administrators').Name"
        result_admin = subprocess.run(['powershell.exe', command_admin], capture_output=True)
        output_admin = result_admin.stdout.decode()
        # Users
        command_users ="(Get-LocalGroupMember -Name 'Users').Name"
        result_users = subprocess.run(['powershell.exe', command_users], capture_output=True)
        output_users = result_users.stdout.decode()
        user=users[i].strip()
        if user in output_admin:
            priv="admin"
        elif user in output_users:
            priv="user"
        else:
            priv="guest"
        report_user+="""<tr>
                        <td style="width:1em"> </td>
                        <td class="hasInfo" title=\""""+full_name+"""
        Privilege: """+priv+"""
        Enabled: """+status+"""
        Number of Logons: """+logons+"""
        Password Age: """+pass_age+""" days
        SID: """+sid+"""\">"""+user+"""</td>
                        <td> """+last_logon+"""</td>
                        <td>("""+priv+""")</td>
                    </tr>"""
    # local system accounts
    command_user_name ="(Get-LocalUser | Where-Object {$_.Enabled -eq $false}).Name"
    result_user_name = subprocess.run(['powershell.exe', command_user_name], capture_output=True)
    output_user_name=result_user_name.stdout.decode()
    users=output_user_name.strip().split("\n")
    full_name=""
    status=""
    last_logon=""
    sid=""
    logons=""
    pass_age=""
    priv=""
    report_user+="""<th colspan="2" class="hasInfo"
								title="Local accounts not associated with a recent user">local system
								accounts</th>
							<th>last logon</th>
						</tr>"""
    for i in range(len(users)):
        # full name
        command_user_fullname ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).FullName"
        result_user_fullname = subprocess.run(['powershell.exe', command_user_fullname], capture_output=True)
        output_user_fullname = result_user_fullname.stdout.decode()
        full_name=users[i].strip()+"-"+output_user_fullname.strip()
        # status
        command_user_status ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).Enabled"
        result_user_status = subprocess.run(['powershell.exe', command_user_status], capture_output=True)
        output_user_status = result_user_status.stdout.decode()
        status=output_user_status.strip()
        # last logon
        command_user_last_logon ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).LastLogon"
        result_user_last_logon = subprocess.run(['powershell.exe', command_user_last_logon], capture_output=True)
        output_user_last_logon = result_user_last_logon.stdout.decode().strip()
        if (output_user_last_logon==""):
            last_logon="never"
        else:
            last_logon=output_user_last_logon
        # SID
        command_user_SID ="(Get-LocalUser | Where-Object {$_.Name -eq '"+users[i].strip()+"'}).SID.Value"
        result_user_SID = subprocess.run(['powershell.exe', command_user_SID], capture_output=True)
        output_user_SID = result_user_SID.stdout.decode()
        sid=output_user_SID.strip()
        # no of logons
        command_user_logon="(Get-EventLog -LogName Security -InstanceId 4624 -ErrorAction SilentlyContinue | Where-Object {$_.ReplacementStrings[5] -eq '"+users[i].strip()+"'}).Count"
        result_user_logon = subprocess.run(['powershell.exe', command_user_logon], capture_output=True)
        output_user_logon = result_user_logon.stdout.decode()
        logons=output_user_logon.strip()
        # password age
        command_user_pass_age="(New-TimeSpan -Start (Get-LocalUser -Name '"+users[i].strip()+"').PasswordLastSet -End (Get-Date)).Days"
        result_user_pass_age = subprocess.run(['powershell.exe', command_user_pass_age], capture_output=True)
        output_user_pass_age = result_user_pass_age.stdout.decode().strip()
        if output_user_pass_age=="":
            pass_age="0"
        else:
            pass_age=output_user_pass_age.strip()
        # Privilege Level
        # Administrators
        command_admin ="(Get-LocalGroupMember -Name 'Administrators').Name"
        result_admin = subprocess.run(['powershell.exe', command_admin], capture_output=True)
        output_admin = result_admin.stdout.decode()
        # Users
        command_users ="(Get-LocalGroupMember -Name 'Users').Name"
        result_users = subprocess.run(['powershell.exe', command_users], capture_output=True)
        output_users = result_users.stdout.decode()
        user=users[i].strip()
        if user in output_admin:
            priv="admin"
        elif user in output_users:
            priv="user"
        else:
            priv="guest"
        report_user+="""<tr>
                        <td style="width:1em"> </td>
                        <td class="hasInfo" title=\""""+full_name+"""
        Privilege: """+priv+"""
        Enabled: """+status+"""
        Number of Logons: """+logons+"""
        Password Age: """+pass_age+""" days
        SID: """+sid+"""\">"""+user+"""</td>
                        <td> """+last_logon+"""</td>
                        <td>("""+priv+""")</td>
                    </tr>"""
    report_user+="""</table>
				</div>
			</div>"""
    return report_user

# Right side of Section Four
# Get Printer Details
def printer_details():
    report_printer = ""
    report_printer+="""<div class="reportSection rsRight">
				<h2 class="reportSectionHeader">
					Printers
				</h2>
				<div class="reportSectionBody">
					<!-- SummaryPrinters => Summary -->
					<table role="presentation">"""
    command_printer_name = "(Get-Printer).DriverName"
    result_printer_name = subprocess.run(["powershell", command_printer_name], shell=True, stdout=subprocess.PIPE)
    output_printer_name = result_printer_name.stdout.decode("utf-8")
    printers=output_printer_name.split("\n")
    printers = list(filter(lambda x: x.strip(), printers))
    command_printer_port = "(Get-Printer).PortName"
    result_printer_port = subprocess.run(["powershell", command_printer_port], shell=True, stdout=subprocess.PIPE)
    output_printer_port = result_printer_port.stdout.decode("utf-8")
    ports=output_printer_port.split("\n")
    ports = list(filter(lambda x: x.strip(), ports))
    for i in range(len(printers)):
        report_printer+="""<tr>
                        <td>"""+printers[i].strip()+"""</td>
                        <td>on """+ports[i].strip()+"""</td>
                    </tr>"""
    report_printer+="""</table>
				</div>
			</div>"""
    return report_printer

# User details
def user_details_Report(show_verbose):
    report_user = ""
    print(yellow, end='')
    print('[!] Fetching User details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] User details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_user = user_details()
        print(green, end='')
        print('[√] User details fetched successfully')
    except Exception as e:
        report_user = f"""<div class="reportSection rsLeft">
            <h2 class="reportSectionHeader">
                User
            </h2>
            <div class="reportSectionBody">
                <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching User Details</span>
            </div>
        </div>"""
        print(red, end='')
        print('[×] Error in Fetching User details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] User details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_user': report_user}

# Printer details
def printer_details_Report(show_verbose):
    report_printer = ""
    print(yellow, end='')
    print('[!] Fetching Printer details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Printer details fetching starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report_printer = printer_details()
        print(green, end='')
        print('[√] Printer details fetched successfully')
    except Exception as e:
        report_printer = f"""<div class="reportSection rsRight">
            <h2 class="reportSectionHeader">
                Printers
            </h2>
            <div class="reportSectionBody">
                <span style="font-size: 1.28em; color: #ff0000;">Error In Fetching Printers Details</span>
            </div>
        </div>"""
        print(red, end='')
        print('[×] Error in Fetching Printer details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Printer details fetched at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'report_printer': report_printer}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")