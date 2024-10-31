# -*- coding: utf-8 -*-
import getpass
import socket
import subprocess
import time
import datetime
from configuration import red, green, yellow, log_col, debug_mode, __version__

header_HTML = """<!DOCTYPE html>
            <html lang="en">
            <head>
                <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>ChangeTitle</title>
                <meta name="description" content="changeMe">
                <meta name="author" content="Hrishikesh (Hrishikesh7665) & Shabdik (ninja-hattori)">
                <link rel="shortcut icon" href="./assets/favicon.ico" type="image/x-icon">
                <link rel="icon" href="./assets/favicon.ico" type="image/x-icon">
                <link rel="stylesheet" href="./assets/style.css">
                <script src="./assets/script.js"></script>
            </head>
            <body>
                <div class="watermark">
                    <img height="375" width="375" src="./assets/logo.png" alt="System Scrutinizer">
                </div>
                <section>
                    <div class="header">
                        <h1>System Scrutinizer<p style="font-size: 20px !important; margin-top: 10px;">Version : """+__version__+"""</p>
                        </h1>
                    </div>
                </section>
                <section>
                    <div class="report">
                        <h1 class="reportTitle">Computer Profile Summary</h1>
                        <div class="reportTitleDesc">
                            <p class="reportTitleDescP"><b>Computer Profile Summary</b> provides a detailed overview of the technical specifications and configuration of this computer system. It typically includes information such as the operating system, processor type and speed, memory capacity, storage capacity, network and other relevant details about the computer's hardware and software. This summary can be used for inventory management, troubleshooting, and technical support purposes, as well as for identifying and purchasing compatible hardware and software for the computer. Additionally, it can be useful for tracking and analyzing the performance of the computer over time.</p>
                        </div>
                        <!--				[Header]					-->
                        <table class="reportHeader">
                            <colgroup>
                                <col style="width: 50%">
                                <col style="width: 50%">
                            </colgroup>
                            """
# Header section
def header():
    report_header = ""
    computer_name = ""
    # Get computer name
    computer_name = socket.gethostname()
    # Get current date and time
    current_time = time.strftime("%Y-%m-%d %H:%M:%S")
    # Get current user
    username = getpass.getuser()
    #Uptime
    result = subprocess.run(['powershell.exe', '(Get-Date) - (gcim Win32_OperatingSystem).LastBootUpTime'], capture_output=True)
    output=result.stdout.decode()
    days=""
    hrs=""
    mins=""
    secs=""
    uptime=""
    for line in output.split("\n"):
        if "Days              :" in line:
            days=line.split(":")[1].strip()
        if "Hours             :" in line:
            hrs=line.split(":")[1].strip()
        if "Minutes           :" in line:
            mins=line.split(":")[1].strip()
        if "Seconds           : " in line:
            secs=line.split(":")[1].strip()
    uptime = days+":"+hrs+":"+mins+":"+secs
    # Header
    report_header += f"""<tr>
                        <th>Computer Name:</th>
                        <td>{computer_name}</td>
                    </tr>
                    <tr>
                        <th>Profile Date:</th>
                        <td>{current_time}</td>
                    </tr>
                    <tr>
                        <th>Windows Logon:</th>
                        <td>{username}</td>
                    </tr>
                    <tr>
                        <th>Uptime:</th>
                        <td>{uptime}</td>
                    </tr>
                </table>"""
    return report_header, computer_name

# Header Report Generation
def header_Report (show_verbose):
    report = ""
    report_header = ""
    computer_name = ""
    print(yellow, end='')
    print('[!] Generating Report Header')
    if show_verbose:
        st_time = time.localtime()
        print(log_col,end="")
        print('[v] Generating Report Header starting at '+str(time.strftime("%H:%M:%S", st_time)))
    try:
        report = header_HTML
        report_header, computer_name = header()
        print(green, end='')
        print('[√] Report Header generated successfully')
    except Exception as e:
        report_header = f"""<tr>
                        <th></th>
                        <td ><span style="margin-left: -110px; font-size: 1.28em; color: #ff0000;">Error In Fetching Header Details</span></td>
                    </tr>
                </table>"""
        print(red, end='')
        print('[×] Error in generating Report Header')
        if debug_mode:
            print('Debug : '+str(e))
    if show_verbose:
        ed_time = time.localtime()
        print(log_col,end="")
        print('[v] Generating Report Header finished at '+str(time.strftime("%H:%M:%S", ed_time)),end="")
        print(" [{:.0f}:{:02.0f}:{:02.0f}]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(ed_time)) - datetime.datetime.fromtimestamp(time.mktime(st_time))).split(":"))))
    
    output_dict = {"report":report, "report_header": report_header, "computer_name": computer_name}
    return output_dict


if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")