# -*- coding: utf-8 -*-
import os
import time
import datetime
import subprocess
from commonFunctions import get_username_from_sid
from configuration import passFlag, whiteFlag, failFlag, strikethrough
from configuration import red, green, yellow, log_col, debug_mode, referenceLink


# Access Credential Manager as a trusted caller
check_2_2_1 = "N/A"
# Access this computer from the network
check_2_2_2 = "N/A"
# Act as part of the operating system
check_2_2_3 = "N/A"
# Adjust memory quotas for a process
check_2_2_4 = "N/A"
# Allow log on locally
check_2_2_5 = "N/A"
# Allow log on through Remote Desktop Services
check_2_2_6 = "N/A"
# Back up files and directories
check_2_2_7 = "N/A"
# Change the system time
check_2_2_8 = "N/A"
# Change the time zone
check_2_2_9 = "N/A"
# Create a pagefile
check_2_2_10 = "N/A"
# Create a token object
check_2_2_11 = "N/A"
# Create global objects
check_2_2_12 = "N/A"
# Create permanent shared objects
check_2_2_13 = "N/A"
# Create symbolic links
check_2_2_14 = "N/A"
# Debug programs
check_2_2_15 = "N/A"
# Deny access to this computer from the network
check_2_2_16 = "N/A"
# Deny log on as a batch job
check_2_2_17 = "N/A"
# Deny log on as a service
check_2_2_18 = "N/A"
# Deny log on locally
check_2_2_19 = "N/A"
# Deny log on through Remote Desktop Services
check_2_2_20 = "N/A"
# Enable computer and user accounts to be trusted for delegation
check_2_2_21 = "N/A"
# Force shutdown from a remote system
check_2_2_22 = "N/A"
# Generate security audits'
check_2_2_23 = "N/A"
# Impersonate a client after authentication
check_2_2_24 = "N/A"
# Increase scheduling priority
check_2_2_25 = "N/A"
# Load and unload device drivers
check_2_2_26 = "N/A"
# Lock pages in memory
check_2_2_27 = "N/A"
# Manage auditing and security log
check_2_2_30 = "N/A"
# Modify an object label
check_2_2_31 = "N/A"
# Modify firmware environment values
check_2_2_32 = "N/A"
# Perform volume maintenance tasks
check_2_2_33 = "N/A"
# Profile single process
check_2_2_34 = "N/A"
# Profile system performance
check_2_2_35 = "N/A"
# Replace a process level token
check_2_2_36 = "N/A"
# Restore files and directories
check_2_2_37 = "N/A"
# Shut down the system
check_2_2_38 = "N/A"
# Take ownership of files or other objects
check_2_2_39 = "N/A"

# Accounts: Administrator account status
check_2_3_1_1 = "N/A"
# Accounts: Block Microsoft accounts
check_2_3_1_2 = "N/A"
# Accounts: Guest account status
check_2_3_1_3 = "N/A"
# Accounts: Limit local account use of blank passwords to console logon only
check_2_3_1_4 = "N/A"
# Accounts: Rename administrator account
check_2_3_1_5 = "N/A"
# Accounts: Rename guest account
check_2_3_1_6 = "N/A"

# Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings
check_2_3_2_1 = "N/A"
# Audit: Shut down system immediately if unable to log security audits
check_2_3_2_2 = "N/A"

# Interactive logon: Do not require CTRL+ALT+DELETE
check_2_3_7_1 = "N/A"
# Interactive logon: Don't display last signed-in
check_2_3_7_2 = "N/A"
# Interactive logon: Machine inactivity limit
check_2_3_7_4 = "N/A"
# 'Interactive logon: Message text for users attempting to log on
check_2_3_7_5 = "N/A"
# 'Interactive logon: Message title for users attempting to log on
check_2_3_7_6 = "N/A"
# Interactive logon: Prompt user to change password before expiration
check_2_3_7_7 = "N/A"
# Interactive logon: Smart card removal behavior
check_2_3_7_8 = "N/A"

# Microsoft network client: Digitally sign communications (always)
check_2_3_8_1 = "N/A"
# Microsoft network client: Digitally sign communications (if server agrees)
check_2_3_8_2 = "N/A"
# Microsoft network client: Send unencrypted password to third-party SMB servers
check_2_3_8_3 = "N/A"

# Microsoft network server: Amount of idle time required before suspending session
check_2_3_9_1 = "N/A"
# Microsoft network server: Digitally sign communications (always)
check_2_3_9_2 = "N/A"
# Microsoft network server: Digitally sign communications (if client agrees)
check_2_3_9_3 = "N/A"
# Microsoft network server: Disconnect clients when logon hours expire
check_2_3_9_4 = "N/A"
# Microsoft network server: Server SPN target name validation level
check_2_3_9_5 = "N/A"


def LocalPolicies_Reporter():
    global check_2_3_1_1, check_2_3_1_2, check_2_3_1_3, check_2_3_1_4, check_2_3_1_5, check_2_3_1_6, check_2_3_2_1, check_2_3_2_2
    global check_2_3_7_1, check_2_3_7_2, check_2_3_7_4, check_2_3_7_5, check_2_3_7_6, check_2_3_7_7, check_2_3_7_8
    global_vars = globals()
    for i in range(1, 40):
        if i == 28 or i == 29:
            continue        
        variable_name = f"check_2_2_{i}"
        global_vars[variable_name]
    variables_listLP = [global_vars["check_2_2_1"], global_vars["check_2_2_2"], global_vars["check_2_2_3"], global_vars["check_2_2_4"], global_vars["check_2_2_5"], global_vars["check_2_2_6"], global_vars["check_2_2_7"], global_vars["check_2_2_8"], global_vars["check_2_2_9"], global_vars["check_2_2_10"], global_vars["check_2_2_11"], global_vars["check_2_2_12"], global_vars["check_2_2_13"], global_vars["check_2_2_14"], global_vars["check_2_2_15"], global_vars["check_2_2_16"], global_vars["check_2_2_17"], global_vars["check_2_2_18"], global_vars["check_2_2_19"], global_vars["check_2_2_20"], global_vars["check_2_2_21"], global_vars["check_2_2_22"], global_vars["check_2_2_23"], global_vars["check_2_2_24"], global_vars["check_2_2_25"], global_vars["check_2_2_26"], global_vars["check_2_2_27"], global_vars["check_2_2_30"], global_vars["check_2_2_31"], global_vars["check_2_2_32"], global_vars["check_2_2_33"], global_vars["check_2_2_34"], global_vars["check_2_2_35"], global_vars["check_2_2_36"], global_vars["check_2_2_37"], global_vars["check_2_2_38"], global_vars["check_2_2_39"], check_2_3_1_1, check_2_3_1_2, check_2_3_1_3, check_2_3_1_4, check_2_3_1_5, check_2_3_1_6, check_2_3_2_1, check_2_3_2_2, check_2_3_7_1, check_2_3_7_2, check_2_3_7_4, check_2_3_7_5, check_2_3_7_6, check_2_3_7_7, check_2_3_7_8]
    countLP = sum(variableLP == True for variableLP in variables_listLP if variableLP != "N/A" and variableLP is not None)
    total_variablesLP = len(variables_listLP)
    non_default_variablesLP = total_variablesLP - (variables_listLP.count("N/A") + variables_listLP.count(None))
    scoreLP = round((countLP / non_default_variablesLP) * 100)
    localPoliciesRep = f'''<!-- Local Policy's -->
			<table class="ReportSection" id="section2">
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
                            <img {passFlag if scoreLP >= 50 else failFlag if scoreLP < 35 else whiteFlag} title="Score for this section is {scoreLP} out of 100" />
                        </td>
                        <td><button class="Ctl" title="Expand" onclick="toggleTableBody(this)"><img src="./assets/plus.png" alt="Expand" /></button></td>
						<th colspan="2">Local Policy's Settings
                        </th>
                        <td>Section Result:&ensp;&ensp;{countLP}&ensp;of&ensp;{total_variablesLP} settings pass</td>
                    </tr>
				</thead>
                <tbody style="display:none;">
					<tr{strikethrough if global_vars["check_2_2_1"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_1"] == True else failFlag if global_vars["check_2_2_1"] == False else whiteFlag }></td>
                        <th>2.2.1</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Access Credential Manager as a trusted caller' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_2"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_2"] == True else failFlag if global_vars["check_2_2_2"] == False else whiteFlag }></td>
                        <th>2.2.2</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);"> 'Access this computer from the network' is must set to 'Administrators, Remote Desktop Users'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_3"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_3"] == True else failFlag if global_vars["check_2_2_3"] == False else whiteFlag }></td>
                        <th>2.2.3</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Act as part of the operating system' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_4"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_4"] == True else failFlag if global_vars["check_2_2_4"] == False else whiteFlag }></td>
                        <th>2.2.4</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Adjust memory quotas for a process' is must set to 'Administrators, LOCAL SERVICE, NETWORK SERVICE'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_5"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_5"] == True else failFlag if global_vars["check_2_2_5"] == False else whiteFlag }></td>
                        <th>2.2.5</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Allow log on locally' is must set to 'Administrators, Users'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_6"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_6"] == True else failFlag if global_vars["check_2_2_6"] == False else whiteFlag }></td>
                        <th>2.2.6</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Allow log on through Remote Desktop Services' is must set to 'Administrators, Remote Desktop Users'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_7"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_7"] == True else failFlag if global_vars["check_2_2_7"] == False else whiteFlag }></td>
                        <th>2.2.7</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Back up files and directories' is must set to 'Administrators'</a>
                        </td>
                    </tr>      
                    <tr{strikethrough if global_vars["check_2_2_8"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_8"] == True else failFlag if global_vars["check_2_2_8"] == False else whiteFlag }></td>
                        <th>2.2.8</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Change the system time' is must set to 'Administrators, LOCAL SERVICE'</a>
                        </td>
                    </tr>    
                    <tr{strikethrough if global_vars["check_2_2_9"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_9"] == True else failFlag if global_vars["check_2_2_9"] == False else whiteFlag }></td>
                        <th>2.2.9</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Change the time zone' is must set to 'Administrators, LOCAL SERVICE, Users'</a>
                        </td>
                    </tr>    
                    <tr{strikethrough if global_vars["check_2_2_10"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_10"] == True else failFlag if global_vars["check_2_2_10"] == False else whiteFlag }></td>
                        <th>2.2.10</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Create a pagefile' is must set to 'Administrators'</a>
                        </td>
                    </tr>                                                                              
                    <tr{strikethrough if global_vars["check_2_2_11"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_11"] == True else failFlag if global_vars["check_2_2_11"] == False else whiteFlag }></td>
                        <th>2.2.11</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Create a token object' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_12"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_12"] == True else failFlag if global_vars["check_2_2_12"] == False else whiteFlag }></td>
                        <th>2.2.12</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Create global objects' is must set to 'Administrators, LOCAL SERVICE, NETWORK SERVICE, SERVICE'</a>
                        </td>
                    </tr>                    
                    <tr{strikethrough if global_vars["check_2_2_13"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_13"] == True else failFlag if global_vars["check_2_2_13"] == False else whiteFlag }></td>
                        <th>2.2.13</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Create permanent shared objects' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_14"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_14"] == True else failFlag if global_vars["check_2_2_14"] == False else whiteFlag }></td>
                        <th>2.2.14</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">Configure 'Create symbolic links'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_15"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_15"] == True else failFlag if global_vars["check_2_2_15"] == False else whiteFlag }></td>
                        <th>2.2.15</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Debug programs' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_16"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_16"] == True else failFlag if global_vars["check_2_2_16"] == False else whiteFlag }></td>
                        <th>2.2.16</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Deny access to this computer from the network' must include 'Guests'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_17"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_17"] == True else failFlag if global_vars["check_2_2_17"] == False else whiteFlag }></td>
                        <th>2.2.17</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Deny log on as a batch job' must include 'Guests'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_18"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_18"] == True else failFlag if global_vars["check_2_2_18"] == False else whiteFlag }></td>
                        <th>2.2.18</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Deny log on as a service' must include 'Guests'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_19"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_19"] == True else failFlag if global_vars["check_2_2_19"] == False else whiteFlag }></td>
                        <th>2.2.19</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Deny log on locally' must include 'Guests'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_20"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_20"] == True else failFlag if global_vars["check_2_2_20"] == False else whiteFlag }></td>
                        <th>2.2.20</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Deny log on through Remote Desktop Services' must include 'Guests'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_21"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_21"] == True else failFlag if global_vars["check_2_2_21"] == False else whiteFlag }></td>
                        <th>2.2.21</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Enable computer and user accounts to be trusted for delegation' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_22"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_22"] == True else failFlag if global_vars["check_2_2_22"] == False else whiteFlag }></td>
                        <th>2.2.22</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Force shutdown from a remote system' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_23"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_23"] == True else failFlag if global_vars["check_2_2_23"] == False else whiteFlag }></td>
                        <th>2.2.23</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Generate security audits' is must set to 'LOCAL SERVICE, NETWORK SERVICE'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_24"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_24"] == True else failFlag if global_vars["check_2_2_24"] == False else whiteFlag }></td>
                        <th>2.2.24</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Impersonate a client after authentication' is must set to 'Administrators, LOCAL SERVICE, NETWORK SERVICE, SERVICE'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_25"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_25"] == True else failFlag if global_vars["check_2_2_25"] == False else whiteFlag }></td>
                        <th>2.2.25</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Increase scheduling priority' is must set to 'Administrators, Window Manager\Window Manager Group'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_26"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_26"] == True else failFlag if global_vars["check_2_2_26"] == False else whiteFlag }></td>
                        <th>2.2.26</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Load and unload device drivers' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_27"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_27"] == True else failFlag if global_vars["check_2_2_27"] == False else whiteFlag }></td>
                        <th>2.2.27</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Lock pages in memory' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_30"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_30"] == True else failFlag if global_vars["check_2_2_30"] == False else whiteFlag }></td>
                        <th>2.2.30</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Manage auditing and security log' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_31"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_31"] == True else failFlag if global_vars["check_2_2_31"] == False else whiteFlag }></td>
                        <th>2.2.31</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Modify an object label' is must set to 'No One'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_32"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_32"] == True else failFlag if global_vars["check_2_2_32"] == False else whiteFlag }></td>
                        <th>2.2.32</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Modify firmware environment values' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_33"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_33"] == True else failFlag if global_vars["check_2_2_33"] == False else whiteFlag }></td>
                        <th>2.2.33</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Perform volume maintenance tasks' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_34"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_34"] == True else failFlag if global_vars["check_2_2_34"] == False else whiteFlag }></td>
                        <th>2.2.34</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Profile single process' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_35"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_35"] == True else failFlag if global_vars["check_2_2_35"] == False else whiteFlag }></td>
                        <th>2.2.35</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Profile system performance' is must set to 'Administrators, NT SERVICE\WdiServiceHost'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_36"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_36"] == True else failFlag if global_vars["check_2_2_36"] == False else whiteFlag }></td>
                        <th>2.2.36</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);"> 'Replace a process level token' is must set to 'LOCAL SERVICE, NETWORK SERVICE'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_37"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_37"] == True else failFlag if global_vars["check_2_2_37"] == False else whiteFlag }></td>
                        <th>2.2.37</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Restore files and directories' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_38"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_38"] == True else failFlag if global_vars["check_2_2_38"] == False else whiteFlag }></td>
                        <th>2.2.38</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Shut down the system' is must set to 'Administrators, Users'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if global_vars["check_2_2_39"] == None else ''}><td></td>
                        <td><img {passFlag if global_vars["check_2_2_39"] == True else failFlag if global_vars["check_2_2_39"] == False else whiteFlag }></td>
                        <th>2.2.39</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Take ownership of files or other objects' is must set to 'Administrators'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_1  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_1 == True else failFlag if check_2_3_1_1 == False else whiteFlag }></td>
                        <th>2.3.1.1</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Administrator account status' is must set to 'Disabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_2  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_2 == True else failFlag if check_2_3_1_2 == False else whiteFlag }></td>
                        <th>2.3.1.2</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Block Microsoft accounts' is must set to 'Users can't add or log on with Microsoft accounts'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_3  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_3 == True else failFlag if check_2_3_1_3 == False else whiteFlag }></td>
                        <th>2.3.1.3</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Guest account status' is must set to 'Disabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_4  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_4 == True else failFlag if check_2_3_1_4 == False else whiteFlag }></td>
                        <th>2.3.1.4</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Limit local account use of blank passwords to console logon only' is must set to 'Enabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_5  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_5 == True else failFlag if check_2_3_1_5 == False else whiteFlag }></td>
                        <th>2.3.1.5</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Rename administrator account'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_1_6  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_1_6 == True else failFlag if check_2_3_1_6 == False else whiteFlag }></td>
                        <th>2.3.1.6</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Accounts: Rename guest account'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_2_1  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_2_1 == True else failFlag if check_2_3_2_1 == False else whiteFlag }></td>
                        <th>2.3.2.1</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings' is must set to 'Enabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_2_2  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_2_2 == True else failFlag if check_2_3_2_2 == False else whiteFlag }></td>
                        <th>2.3.2.2</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Audit: Shut down system immediately if unable to log security audits' is must set to 'Disabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_1  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_1 == True else failFlag if check_2_3_7_1 == False else whiteFlag }></td>
                        <th>2.3.7.1</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Do not require CTRL+ALT+DEL' is must set to 'Disabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_2  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_2 == True else failFlag if check_2_3_7_2 == False else whiteFlag }></td>
                        <th>2.3.7.2</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Don't display last signed-in' is must set to 'Enabled'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_4  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_4 == True else failFlag if check_2_3_7_4 == False else whiteFlag }></td>
                        <th>2.3.7.4</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Machine inactivity limit' is must set to 900 or fewer second(s), but not 0</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_5  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_5 == True else failFlag if check_2_3_7_5 == False else whiteFlag }></td>
                        <th>2.3.7.5</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Message text for users attempting to log on' must be configured</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_6  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_6 == True else failFlag if check_2_3_7_6 == False else whiteFlag }></td>
                        <th>2.3.7.6</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Message title for users attempting to log on' must be configured</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_7  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_7 == True else failFlag if check_2_3_7_7 == False else whiteFlag }></td>
                        <th>2.3.7.7</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Prompt user to change password before expiration' is must set to 'between 5 and 14 days'</a>
                        </td>
                    </tr>
                    <tr{strikethrough if check_2_3_7_8  == None else ''}><td></td>
                        <td><img {passFlag if check_2_3_7_8 == True else failFlag if check_2_3_7_8 == False else whiteFlag }></td>
                        <th>2.3.7.8</th>
                        <td colspan="2">
                            <a href="{referenceLink}" target="BkHelpWindow" onclick="createHelpWindow(event);">'Interactive logon: Smart card removal behavior' is must be configured</a>
                        </td>
                    </tr>
    '''
    localPoliciesRep += '</tbody></table>'
    return localPoliciesRep, scoreLP


# Check User Rights Assignment
# CIS Benchmark 2.2.1 to 2.2.39
def checkUserRightsAssignment():
    global_vars = globals()
    for i in range(1, 40):
        if i == 28 or i == 29:
            continue        
        variable_name = f"check_2_2_{i}"
        global_vars[variable_name]
    sec_titles = ['SeNetworkLogonRight', 'SeTcbPrivilege', 'SeBackupPrivilege', 'SeSystemtimePrivilege', 'SeCreatePagefilePrivilege', 'SeCreateTokenPrivilege', 'SeCreatePermanentPrivilege', 'SeDebugPrivilege', 'SeRemoteShutdownPrivilege', 'SeAuditPrivilege', 'SeIncreaseQuotaPrivilege', 'SeIncreaseBasePriorityPrivilege', 'SeLoadDriverPrivilege', 'SeLockMemoryPrivilege', 'SeBatchLogonRight', 'SeServiceLogonRight', 'SeInteractiveLogonRight', 'SeSecurityPrivilege', 'SeSystemEnvironmentPrivilege', 'SeProfileSingleProcessPrivilege', 'SeSystemProfilePrivilege', 'SeAssignPrimaryTokenPrivilege', 'SeRestorePrivilege', 'SeShutdownPrivilege', 'SeTakeOwnershipPrivilege', 'SeDenyNetworkLogonRight', 'SeDenyBatchLogonRight', 'SeDenyServiceLogonRight', 'SeDenyInteractiveLogonRight', 'SeEnableDelegationPrivilege', 'SeManageVolumePrivilege', 'SeRemoteInteractiveLogonRight', 'SeDenyRemoteInteractiveLogonRight', 'SeImpersonatePrivilege', 'SeCreateGlobalPrivilege', 'SeTrustedCredManAccessPrivilege', 'SeRelabelPrivilege', 'SeTimeZonePrivilege', 'SeCreateSymbolicLinkPrivilege']
    command = "secedit /export /areas user_rights /cfg 'C:\\user_rights.cfg' > $null; (Get-Content 'C:\\user_rights.cfg') | Select-String {}".format(', '.join(f"'{item}'" for item in sec_titles))+" -ErrorAction SilentlyContinue"
    result = subprocess.run(["powershell", command], shell=True, capture_output=True)
    output = result.stdout.decode("utf-8").strip()
    presentList = [line.split(' = ')[0] for line in output.split('\n') if line and line[:2] == 'Se']
    if 'SeTrustedCredManAccessPrivilege' in presentList:
        global_vars["check_2_2_1"] = False
    else:
        global_vars["check_2_2_1"] = True
    
    if 'SeTcbPrivilege' in presentList:
        global_vars["check_2_2_3"] = False
    else:
        global_vars["check_2_2_3"] = True
    
    if 'SeCreateTokenPrivilege' in presentList:
        global_vars["check_2_2_11"] = False
    else:
        global_vars["check_2_2_11"] = True
    
    if 'SeCreatePermanentPrivilege' in presentList:
        global_vars["check_2_2_13"] = False
    else:
        global_vars["check_2_2_13"] = True
    
    if 'SeEnableDelegationPrivilege' in presentList:
        global_vars["check_2_2_21"] = False
    else:
        global_vars["check_2_2_21"] = True
    
    if 'SeLockMemoryPrivilege' in presentList:
        global_vars["check_2_2_27"] = False
    else:
        global_vars["check_2_2_27"] = True
    
    if 'SeRelabelPrivilege' in presentList:
        global_vars["check_2_2_31"] = False
    else:
        global_vars["check_2_2_31"] = True
    for line in output.split('\n'):
        if "SeNetworkLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'Remote Desktop Users'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_2"] = True
            else:
                global_vars["check_2_2_2"] = False
            
        elif "SeIncreaseQuotaPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid, True).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'LOCAL SERVICE'.lower(), 'NETWORK SERVICE'.lower()] for item in nameList) and len(nameList) == 3:
                global_vars["check_2_2_4"] = True
            else:
                global_vars["check_2_2_4"] = False
            
        elif "SeInteractiveLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid, True).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'Users'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_5"] = True
            else:
                global_vars["check_2_2_5"] = False
        
        elif "SeRemoteInteractiveLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'Remote Desktop Users'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_6"] = True
            else:
                global_vars["check_2_2_6"] = False
            
        elif "SeBackupPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_7"] = True
            else:
                global_vars["check_2_2_7"] = False
        
        elif "SeSystemtimePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'LOCAL SERVICE'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_8"] = True
            else:
                global_vars["check_2_2_8"] = False
        
        elif "SeTimeZonePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'LOCAL SERVICE'.lower(), 'Users'.lower()] for item in nameList) and len(nameList) == 3:
                global_vars["check_2_2_9"] = True
            else:
                global_vars["check_2_2_9"] = False
        
        elif "SeCreatePagefilePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_10"] = True
            else:
                global_vars["check_2_2_10"] = False
                
        elif "SeCreateGlobalPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_12"] = True
            else:
                global_vars["check_2_2_12"] = False
                
        elif "SeCreateSymbolicLinkPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 4:
                global_vars["check_2_2_14"] = True
            else:
                global_vars["check_2_2_14"] = False
        elif "SeDebugPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_15"] = True
            else:
                global_vars["check_2_2_15"] = False
        
        elif "SeDenyNetworkLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if 'Guest'.lower() in nameList:
                global_vars["check_2_2_16"] = True
            else:
                global_vars["check_2_2_16"] = False
        
        elif "SeDenyBatchLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if 'Guest'.lower() in nameList:
                global_vars["check_2_2_17"] = True
            else:
                global_vars["check_2_2_17"] = False
        
        elif "SeDenyServiceLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if 'Guest'.lower() in nameList:
                global_vars["check_2_2_18"] = True
            else:
                global_vars["check_2_2_18"] = False
        
        elif "SeDenyInteractiveLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if 'Guest'.lower() in nameList:
                global_vars["check_2_2_19"] = True
            else:
                global_vars["check_2_2_19"] = False
        
        elif "SeDenyRemoteInteractiveLogonRight" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if 'Guest'.lower() in nameList:
                global_vars["check_2_2_20"] = True
            else:
                global_vars["check_2_2_20"] = False
        
        elif "SeRemoteShutdownPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_22"] = True
            else:
                global_vars["check_2_2_22"] = False
        
        elif "SeAuditPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['LOCAL SERVICE'.lower(), 'NETWORK SERVICE'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_23"] = True
            else:
                global_vars["check_2_2_23"] = False
        
        elif "SeImpersonatePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['LOCAL SERVICE'.lower(), 'NETWORK SERVICE'.lower(), 'SERVICE'.lower()] for item in nameList) and len(nameList) == 3:
                global_vars["check_2_2_24"] = True
            else:
                global_vars["check_2_2_24"] = False
        
        elif "SeIncreaseBasePriorityPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid, True).lower() for sid in userList]
            if all(item in ['BUILTIN\Administrators'.lower(), 'Window Manager\Window Manager Group'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_25"] = True
            else:
                global_vars["check_2_2_25"] = False
        
        elif "SeLoadDriverPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_26"] = True
            else:
                global_vars["check_2_2_26"] = False
        
        elif "SeSecurityPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_30"] = True
            else:
                global_vars["check_2_2_30"] = False
        
        elif "SeSystemEnvironmentPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_32"] = True
            else:
                global_vars["check_2_2_32"] = False
            
        elif "SeManageVolumePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_33"] = True
            else:
                global_vars["check_2_2_33"] = False
            
        elif "SeProfileSingleProcessPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_34"] = True
            else:
                global_vars["check_2_2_34"] = False
            
        elif "SeSystemProfilePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid, True).lower() for sid in userList]
            if all(item in ['BUILTIN\Administrators'.lower(), 'NT SERVICE\WdiServiceHost'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_35"] = True
            else:
                global_vars["check_2_2_35"] = False
            
        elif "SeAssignPrimaryTokenPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['LOCAL SERVICE'.lower(), 'NETWORK SERVICE'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_36"] = True
            else:
                global_vars["check_2_2_36"] = False
            
        elif "SeRestorePrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_37"] = True
            else:
                global_vars["check_2_2_37"] = False
            
        elif "SeShutdownPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower(), 'Users'.lower()] for item in nameList) and len(nameList) == 2:
                global_vars["check_2_2_38"] = True
            else:
                global_vars["check_2_2_38"] = False
            
        elif "SeTakeOwnershipPrivilege" in line:
            userList = line.split("=")[1].strip().split(',')
            nameList = [get_username_from_sid(sid).lower() for sid in userList]
            if all(item in ['Administrators'.lower()] for item in nameList) and len(nameList) == 1:
                global_vars["check_2_2_39"] = True
            else:
                global_vars["check_2_2_39"] = False
    for i in range(1, 40):
        if i == 28 or i == 29:
            continue
        variable_name = f"check_2_2_{i}"
        if global_vars[variable_name] == "N/A" :
            global_vars[variable_name] = False
    
    if os.path.exists("C:\\user_rights.cfg"):
        os.remove("C:\\user_rights.cfg")


# Check Security Options > Accounts & Security Options
# CIS Benchmark 2.3.1.1 to 2.3.1.6 & 2.3.2.1, 2.3.2.2
def check_Accounts_and_Audit_Policies():
    global check_2_3_1_1, check_2_3_1_2, check_2_3_1_3, check_2_3_1_4, check_2_3_1_5, check_2_3_1_6, check_2_3_2_1, check_2_3_2_2
    commandAccounts = "secedit /export /areas securitypolicy /cfg 'C:\\accounts_and_audit.cfg' > $null; (Get-Content 'C:\\accounts_and_audit.cfg') | Select-String 'EnableAdminAccount','EnableGuestAccount','NewAdministratorName','NewGuestName','LimitBlankPasswordUse','SCENoApplyLegacyAuditPolicy','CrashOnAuditFail' -ErrorAction SilentlyContinue"
    resultAccounts = subprocess.run(["powershell", commandAccounts], shell=True, capture_output=True)
    outputAccounts = resultAccounts.stdout.decode("utf-8").strip().split('\n')
    if "SCENoApplyLegacyAuditPolicy" not in outputAccounts:
        check_2_3_2_1 = False
    for line in outputAccounts:
        if "EnableAdminAccount" in line:
            try:
                EnableAdminAccount = int(line.split("=")[1].strip())
                if EnableAdminAccount == 0:
                    check_2_3_1_1 = True
                else:
                    check_2_3_1_1 = False
            except:
                check_2_3_1_1 = False
        elif "EnableGuestAccount" in line:
            try:
                EnableGuestAccount = int(line.split("=")[1].strip())
                if EnableGuestAccount == 0:
                    check_2_3_1_3 = True
                else:
                    check_2_3_1_3 = False
            except:
                check_2_3_1_3 = False
        elif "LimitBlankPasswordUse" in line:
            try:
                LimitBlankPasswordUse = int(line.split(",")[1].strip())
                if LimitBlankPasswordUse == 1:
                    check_2_3_1_4 = True
                else:
                    check_2_3_1_4 = False
            except:
                check_2_3_1_4 = False
        elif "NewAdministratorName" in line:
            NewAdministratorName = line.split("=")[1].strip()
            try:
                if NewAdministratorName.lower() == "administrator":
                    check_2_3_1_5 = True
                else:
                    check_2_3_1_5 = False
            except:
                check_2_3_1_5 = True
        elif "NewGuestName" in line:
            NewGuestName = line.split("=")[1].strip()
            try:
                if NewGuestName.lower() == "administrator":
                    check_2_3_1_6 = True
                else:
                    check_2_3_1_6 = False
            except:
                check_2_3_1_6 = True
        elif "SCENoApplyLegacyAuditPolicy" in line:
            try:
                SCENoApplyLegacyAuditPolicy = int(line.split(",")[1].strip())
                if SCENoApplyLegacyAuditPolicy == 1:
                    check_2_3_2_1 = True
                else:
                    check_2_3_2_1 = False
            except:
                check_2_3_2_1 = False
        elif "CrashOnAuditFail" in line:
            try:
                CrashOnAuditFail = int(line.split(",")[1].strip())
                if CrashOnAuditFail == 0:
                    check_2_3_2_2 = True
                else:
                    check_2_3_2_2 = False
            except:
                check_2_3_2_2 = False
    try:
        commandMicrosoftAccounts = "try { (Get-ItemProperty -Path \"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\" -Name \"NoConnectedUser\" -ErrorAction Stop).NoConnectedUser } catch { \"NotSet\" }"
        resultMicrosoftAccounts = subprocess.run(["powershell", commandMicrosoftAccounts], shell=True, capture_output=True)
        outputMicrosoftAccounts = resultMicrosoftAccounts.stdout.decode("utf-8").strip()
        if int(outputMicrosoftAccounts) == 3:
            check_2_3_1_2 = True
        else:
            check_2_3_1_2 = False
    except:
        check_2_3_1_2 = False 

    if os.path.exists("C:\\accounts_and_audit.cfg"):
        os.remove("C:\\accounts_and_audit.cfg")


# Check Security Options > Accounts & Security Options
# CIS Benchmark 2.3.7.1 to 2.3.7.8 (Except 2.3.7.3)
def check_Interactive_logon_and_():
    global check_2_3_7_1, check_2_3_7_2, check_2_3_7_4, check_2_3_7_5, check_2_3_7_6, check_2_3_7_7, check_2_3_7_8
    commandLogon = "secedit /export /areas securitypolicy /cfg 'C:\\Logon_and.cfg' > $null; (Get-Content 'C:\\Logon_and.cfg') | Select-String 'DisableCAD', 'DontDisplayLastUserName', 'InactivityTimeoutSecs', 'LegalNoticeText', 'LegalNoticeCaption', 'PasswordExpiryWarning', 'ScRemoveOption' -ErrorAction SilentlyContinue"
    resultLogon = subprocess.run(["powershell", commandLogon], shell=True, capture_output=True)
    outputLogon = resultLogon.stdout.decode("utf-8").strip().split('\n')
    if "DisableCAD" not in outputLogon:
        check_2_3_7_1 = False
    if "InactivityTimeoutSecs" not in outputLogon:
        check_2_3_7_4 = False
    for line in outputLogon:
        if "DisableCAD" in line:
            try:
                DisableCAD = int(line.split(",")[1].strip())
                if DisableCAD == 0:
                    check_2_3_7_1 = True
                else:
                    check_2_3_7_1 = False
            except:
                check_2_3_7_1 = False
        
        elif "DontDisplayLastUserName" in line:
            try:
                DontDisplayLastUserName = int(line.split(",")[1].strip())
                if DontDisplayLastUserName == 1:
                    check_2_3_7_2 = True
                else:
                    check_2_3_7_2 = False
            except:
                check_2_3_7_2 = False
        
        elif "InactivityTimeoutSecs" in line:
            try:
                InactivityTimeoutSecs = int(line.split(",")[1].strip())
                if InactivityTimeoutSecs != 0 and InactivityTimeout <= 900:
                    check_2_3_7_4 = True
                else:
                    check_2_3_7_4 = False
            except:
                check_2_3_7_4 = False
        
        elif "LegalNoticeText" in line:
            try:
                LegalNoticeText = line.split(",")[1].strip()
                if LegalNoticeText != "" or LegalNoticeText is not None:
                    check_2_3_7_5 = True
                else:
                    check_2_3_7_5 = False
            except:
                check_2_3_7_5 = False
        
        elif "LegalNoticeCaption" in line:
            try:
                LegalNoticeCaption = line.split(",")[1].strip()
                if LegalNoticeCaption != "" or LegalNoticeCaption is not None:
                    check_2_3_7_6 = True
                else:
                    check_2_3_7_6 = False
            except:
                check_2_3_7_6 = False
        
        elif "PasswordExpiryWarning" in line:
            try:
                PasswordExpiryWarning = int(line.split(",")[1].strip())
                if PasswordExpiryWarning == 5 or PasswordExpiryWarning == 14 or (PasswordExpiryWarning > 5 and PasswordExpiryWarning < 14):
                    check_2_3_7_7 = True
                else:
                    check_2_3_7_7 = False
            except:
                check_2_3_7_7 = False
        
        elif "ScRemoveOption" in line:
            try:
                ScRemoveOption = int(line.split(",")[1].strip())
                if ScRemoveOption >= 1 and ScRemoveOption <= 3:
                    check_2_3_7_8 = True
                else:
                    check_2_3_7_8 = False
            except:
                check_2_3_7_8 = False
    
    if os.path.exists("C:\\Logon_and.cfg"):
        os.remove("C:\\Logon_and.cfg")

# check_Interactive_logon_and_()

# ---Local Policies---
def LocalPolicies_Report(show_verbose):
    print(yellow, end='')
    print('[!] Getting Local Policy\'s details')
    if show_verbose:
        ed_timeLP = time.localtime()
        print(log_col,end="")
        print('[v] Local Policy\'s details gathering starting at '+str(time.strftime("%H:%M:%S", ed_timeLP)))
    try:
        try:
            checkUserRightsAssignment()
        except Exception as UA:
            if debug_mode:
                print(red, end='')
                print('[×] Error in gathering User Rights Assignment\'s details')
                print('Debug : '+str(UA))
        try:
            check_Accounts_and_Audit_Policies()
        except Exception as AA:
            if debug_mode:
                print(red, end='')
                print('[×] Error in gathering Account Policy\'s and Audit details')
                print('Debug : '+str(AA))
        try:
            check_Interactive_logon_and_()
        except Exception as IL:
            if debug_mode:
                print(red, end='')
                print('[×] Error in gathering Interactive logon Policy\'s and Audit details')
                print('Debug : '+str(IL))
    except Exception as e:
        print(red, end='')
        print('[×] Error in gathering Local Policy\'s details')
        if debug_mode:
            print('Debug : '+str(e))
    CISsection2HTML, scoreLPFinal = LocalPolicies_Reporter()
    print(green, end='')
    print('[√] Local Policy\'s details gathered successfully')
    if show_verbose:
        ed_timeLP = time.localtime()
        print(log_col,end="")
        print('[v] Local Policy\'s details gathering finished at '+str(time.strftime("%H:%M:%S", ed_timeLP)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_timeLP)) - datetime.datetime.fromtimestamp(time.mktime(ed_timeLP))).split(":"))))
    return {'CISsection2Score' : scoreLPFinal,'CISsection2HTML': CISsection2HTML}

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")