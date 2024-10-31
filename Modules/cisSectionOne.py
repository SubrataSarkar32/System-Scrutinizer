# -*- coding: utf-8 -*-
import os
import time
import datetime
import subprocess
from configuration import passFlag, whiteFlag, failFlag, strikethrough
from configuration import red, green, yellow, log_col, debug_mode, referenceLink

# Enforce password history
check_1_1_1 = "N/A"
# Maximum password age
check_1_1_2 = "N/A"
# Minimum password age
check_1_1_3 = "N/A"
# Minimum password length
check_1_1_4 = "N/A"
# Password must meet complexity requirements
check_1_1_5 = "N/A"
# Relax minimum password length limits
check_1_1_6 = "N/A"
# Store passwords using reversible encryption
check_1_1_7 = "N/A"

# Account lockout duration
check_1_2_1 = "N/A"
# Account lockout threshold
check_1_2_2 = "N/A"
# Reset account lockout counter after
check_1_2_3 = "N/A"

def accountReporter():
    global check_1_1_1, check_1_1_2, check_1_1_3, check_1_1_4, check_1_1_5, check_1_1_6, check_1_1_7, check_1_2_1, check_1_2_2, check_1_2_3
    variables_list = [check_1_1_1, check_1_1_2, check_1_1_3, check_1_1_4, check_1_1_5, check_1_1_6, check_1_1_7, check_1_2_1, check_1_2_2, check_1_2_3]
    count = sum(variable == True for variable in variables_list if variable != "N/A" and variable is not None)
    total_variables = len(variables_list)
    non_default_variables = total_variables - (variables_list.count("N/A") + variables_list.count(None))
    score = round((count / non_default_variables) * 100)

    accountRep = f'''<!-- Account Policy's -->
			<table class="ReportSection" id="section1">
				<colgroup>
					<col style="width: 2em">
					<col style="width: 2em">
					<col style="width: 2em">
					<col style="width: 90%">
					<col>
				</colgroup>
				<thead>
					<tr>
						<td class="noshed">
							<img {passFlag if score >= 50 else failFlag if score < 35 else whiteFlag} title="Score for this section is {score} out of 100" />
						</td>
						<td><button class="Ctl" title="Expand" onclick="toggleTableBody(this)"><img src="./assets/plus.png" alt="Expand" /></button></td>
						<th colspan="2">
							Account Policy's Settings
						</th>
						<td>Section Result:&ensp;&ensp;{count}&ensp;of&ensp;{total_variables} settings pass</td>
					</tr>
				</thead>
                <tbody style="display:none;">
					<tr{strikethrough if check_1_1_1 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_1 == True else failFlag if check_1_1_1 == False else whiteFlag }></td>
						<th>1.1.1</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Enforce password history' is must be set to '24 or more password(s)'</a>
						</td>
					</tr>
					<tr{strikethrough if check_1_1_2 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_2 == True else failFlag if check_1_1_2 == False else whiteFlag }></td>
						<th>1.1.2</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Maximum password age' is must set to '365 or fewer days, but not 0'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_1_3 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_3 == True else failFlag if check_1_1_3 == False else whiteFlag }></td>
						<th>1.1.3</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Minimum password age' is must set to '1 or more day(s)'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_1_4 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_4 == True else failFlag if check_1_1_4 == False else whiteFlag }></td>
						<th>1.1.4</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Minimum password length' is must set to '14 or more character(s)'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_1_5 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_5 == True else failFlag if check_1_1_5 == False else whiteFlag }></td>
						<th>1.1.5</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Password must meet complexity requirements' is must set to 'Enabled'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_1_6 == None else ''}><td></td>
                        <td><img {passFlag if check_1_1_6 == True else failFlag if check_1_1_6 == False else whiteFlag }></td>
						<th>1.1.6</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Relax minimum password length limits' is must set to 'Enabled'</a>
						</td>
					</tr>
					<tr{strikethrough if check_1_1_7 == None else ''}><td></td>
						<td><img {passFlag if check_1_1_7 == True else failFlag if check_1_1_7 == False else whiteFlag }></td>
						<th>1.1.7</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Store passwords using reversible encryption' is must set to 'Disabled'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_2_1 == None else ''}><td></td>
						<td><img {passFlag if check_1_2_1 == True else failFlag if check_1_2_1 == False else whiteFlag }></td>
						<th>1.2.1</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Account lockout duration' is must set to '15 or more minute(s)'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_2_2 == None else ''}><td></td>
						<td><img {passFlag if check_1_2_2 == True else failFlag if check_1_2_2 == False else whiteFlag }></td>
						<th>1.2.2</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Account lockout threshold' is must set to '5 or fewer invalid logon attempt(s), but not 0'</a>
						</td>
					</tr>
                    <tr{strikethrough if check_1_2_3 == None else ''}><td></td>
						<td><img {passFlag if check_1_2_3 == True else failFlag if check_1_2_3 == False else whiteFlag }></td>
						<th>1.2.3</th>
						<td colspan="2">
							<a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Reset account lockout counter after' is must set to '15 or more minute(s)'</a>
						</td>
					</tr>
                    '''
    accountRep += '</tbody></table>'
    return accountRep, score


# Check Password Policy
# CIS Benchmark 1.1.1, 1.1.2, 1.1.3, 1.1.4
def checkPasswordPolicy_1to4():
    global check_1_1_1, check_1_1_2, check_1_1_3, check_1_1_4
    command = 'net accounts'
    result = subprocess.run( ["powershell", command], shell=True, stdout=subprocess.PIPE)
    output = result.stdout.decode("utf-8")
    for line in output.split('\n'):
        try:
            if "Length of password history maintained: " in line:
                password_history =  int(line.split(":")[1].strip())
                if password_history >= 24 :
                    check_1_1_1 = True
                else:
                    check_1_1_1 = False
        except:
            check_1_1_1 = False
        try:
            if "Maximum password age (days): " in line:
                max_password_age =  int(line.split(":")[1].strip())
                if max_password_age <= 365 and  max_password_age != 0:
                    check_1_1_2 = True
                else:
                    check_1_1_2 = False
        except:
            pass
        try:
            if "Minimum password age (days): " in line:
                min_password_age =  int(line.split(":")[1].strip())
                if min_password_age >= 1 :
                    check_1_1_3 = True
                else:
                    check_1_1_3 = False
        except:
            pass
        try:
            if "Minimum password length: " in line:
                min_password_len =  int(line.split(":")[1].strip())
                if min_password_len >= 14 :
                    check_1_1_4 = True
                else:
                    check_1_1_4 = False
        except:
            pass
# check_1_1_5, check_1_1_6, heck_1_1_7, check_1_2_1, check_1_2_2, check_1_2_3
def checkPolicy_5and7and1_2_1to1_2_3():
    global check_1_1_5, check_1_1_6, check_1_1_7, check_1_2_1, check_1_2_2, check_1_2_3
    command = "secedit /export /areas securitypolicy /cfg 'C:\\secpol.cfg' > $null; (Get-Content 'C:\\secpol.cfg') | Select-String 'PasswordComplexity','ClearTextPassword','LockoutDuration','LockoutBadCount','ResetLockoutCount'"
    result = subprocess.run(["powershell", command], shell=True, capture_output=True)
    output = result.stdout.decode("utf-8").strip()
    for line in output.split('\n'):
        if "PasswordComplexity" in line:
            try:
                password_complexity = int(line.split("=")[1].strip())
                if password_complexity == 1:
                    check_1_1_5 = True
                else:
                    check_1_1_5 = False
            except:
                check_1_1_5 = False
        elif "RelaxMinimumPasswordLengthLimits" in line:
            try:
                password_complexity = int(line.split(",")[1].strip())
                if password_complexity == 1:
                    check_1_1_6 = True
                else:
                    check_1_1_6 = False
            except:
                check_1_1_6 = False
        elif "ClearTextPassword" in line:
            try:
                clearTextPassword = int(line.split("=")[1].strip())
                if clearTextPassword == 0:
                    check_1_1_7 = True
                else:
                    check_1_1_7 = False
            except:
                check_1_1_7 = False
        elif "LockoutDuration" in line:
            try:
                LockoutDuration = int(line.split("=")[1].strip())
                if LockoutDuration >= 15:
                    check_1_2_1 = True
                else:
                    check_1_2_1 = False
            except:
                check_1_2_1 = False
        elif "LockoutBadCount" in line:
            try:
                LockoutBadCount = int(line.split("=")[1].strip())
                if LockoutBadCount <=5 and LockoutBadCount!=0:
                    check_1_2_2 = True
                else:
                    check_1_2_2 = False
            except:
                check_1_2_2 = False
        elif "ResetLockoutCount" in line:
            try:
                ResetLockoutCount = int(line.split("=")[1].strip())
                if ResetLockoutCount >= 15:
                    check_1_2_3 = True
                else:
                    check_1_2_3 = False
            except:
                check_1_2_3 = False
    if check_1_1_6 == "N/A":
        check_1_1_6 = False
    if os.path.exists("C:\\secpol.cfg"):
        os.remove("C:\\secpol.cfg")

# ---Account Policies---
def AccountPolicy_Report(show_verbose):
    print(yellow, end='')
    print('[!] Getting Account Policy\'s details')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Account Policy\'s details gathering starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        try:
            checkPasswordPolicy_1to4()
        except Exception as e:
            if debug_mode:
                print(red, end='')
                print('[×] Error in gathering Password Policy\'s details 1.1.1 to 1.1.4')
                print('Debug : '+str(e))
        try:
            checkPolicy_5and7and1_2_1to1_2_3()
        except Exception as e:
            if debug_mode:
                print(red, end='')
                print('[×] Error in gathering Password Policy\'s details 1.1.1 to 1.1.4')
                print('Debug : '+str(e))
        CISsection1HTML, score = accountReporter()
        print(green, end='')
        print('[√] Account Policy\'s details gathered successfully')
    except Exception as e:
        print(red, end='')
        print('[×] Error in gathering Account Policy\'s details')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Account Policy\'s details gathering finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    return {'CISsection1Score' : score,'CISsection1HTML': CISsection1HTML}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")