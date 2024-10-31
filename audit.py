# -*- coding: utf-8 -*-
import concurrent.futures
import multiprocessing
import webbrowser
import argparse
import datetime
import hashlib
import ctypes
import time
import html
import sys
import os
import re

folder_hash = "5324fa1f7b59f244f132151917ef07722c120003cb198398c144ff3a0975bc1e"

show_verbose = False
scanStart_T = None

def current_path():
    CurrentPath = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(CurrentPath)
    if sys.platform == "win32":
        return path.replace(os.sep, '\\')
    else:
        return path.replace(os.sep, '/')
mei_path = current_path()
if sys.platform == "win32":
    modules_path = mei_path+"\\Modules\\"
    base_path = os.getcwd().replace(os.sep, '\\')
    report_path = base_path+"\\System Report\\"
    assets_path = report_path+"\\assets\\"
else:
    modules_path = mei_path+"/Modules/"
    base_path = os.getcwd().replace(os.sep, '/')
    report_path = base_path+"/System Report/"
    assets_path = report_path+"/assets/"

sys.path.insert(1, modules_path)

try:
    from configuration import *
    from configuration import __version__
except Exception as e:
    print ('[×] Internal error occurred')
    print ('Debug : '+str(e))
    print ('[*] Exiting the program')
    exit ()

# Check for file integrity
def check_integrity(folder, hash_value):
    if not os.path.exists(folder):
        return False
    sha256 = hashlib.sha256()
    for root, dirs, files in os.walk(folder):
        for file in files:
            file_path = os.path.join(root, file)
            with open(file_path, 'rb') as f:
                while True:
                    data = f.read(4096)
                    if not data:
                        break
                    sha256.update(data)
    calculated_hash = sha256.hexdigest()
    if calculated_hash == hash_value:
        return True
    else:
        return False

# Terminal banner
def terminal_banner():
    main_str= "                                                                                  "
    text = "Version : " + __version__
    new_string = main_str[len(text):]
    new = "|"+str(new_string[ : len(new_string)//2] + text + new_string[len(new_string)//2 : ])+"|"
    os.system('cls')
    print(ascii_col, end='')
    print(' __________________________________________________________________________________ ')
    print('|  __           _                   __                 _   _       _               |')
    print('| / _\_   _ ___| |_ ___ _ __ ___   / _\ ___ _ __ _   _| |_(_)_ __ (_)_______ _ __  |')
    print('| \ \| | | / __| __/ _ \ \'_ ` _ \  \ \ / __| \'__| | | | __| | \'_ \| |_  / _ \ \'__| |')
    print('| _\ \ |_| \__ \ ||  __/ | | | | | _\ \ (__| |  | |_| | |_| | | | | |/ /  __/ |    |')
    print('| \__/\__, |___/\__\___|_| |_| |_| \__/\___|_|   \__,_|\__|_|_| |_|_/___\___|_|    |')
    print('|     |___/                                                                        |')
    print('|       Developed By : Hrishikesh (Hrishikesh7665) & Shabdik (ninja-hattori)       |')
    print (new)
    print('|__________________________________________________________________________________|')

# # Get Network Drive details
# def network_drive():
#     global report
#     report += """<div class="reportSection rsLeft">
#                 <h2 class="reportSectionHeader">
# 					Network Drives
# 				</h2>
# 				<div class="reportSectionBody">
# 					<i>None detected</i>
# 				</div>
# 			</div>"""

if __name__ == "__main__" :
    os.system("")
    terminal_banner()
    try :
        parser = argparse.ArgumentParser()
        parser = argparse.ArgumentParser(prog='PROG', usage=str(os.path.basename(__file__))+' [options] (-h or --help for help)',add_help=False)
        parser.add_argument('-h', '--help', action='store_true')
        parser.add_argument("-t", "--thread", type=int, help="Number of thread program will use")
        parser.add_argument("-f", "--file", type=str, help="File name of final report")
        parser.add_argument("-v", "--verbose", action='store_true', help="Show some extra logs")
        parser.add_argument("--about", action='store_true', help="Show 'System Scrutinizer' description")
        args = parser.parse_args()
    except SystemExit as e:
        print(nocolor)
        exit()
    if args.help:
        print(yellow,end="")
        print("-h or --help           Show this help message")
        print("--about                Show 'System Scrutinizer' description")
        print("-t or --thread         Specify the number's of thread to use else best will be used")
        print("-f or --file           Specify the output file name else use the default")
        print("-v or --verbose        Specify the output file name else use the default")
        print(nocolor)
        exit()
    if args.about:
        print(nocolor)
        print("System Scrutinizer Description goes here\n")
        exit()
    if args.verbose:
        show_verbose = True
    if args.file != None:
        f_filename = args.file
        try:
            _, file_extension = os.path.splitext(f_filename)
            if file_extension.lower() != ".html":
                # If no extension found
                if not file_extension:
                    f_filename += ".html"
                else:
                    f_filename = f_filename.replace(file_extension, ".html")
        except:
            f_filename = None
    else:
        f_filename = None
    if args.thread == None:
        num_threads = int((multiprocessing.cpu_count()/2)+1)
        print(yellow,end="")
        print("\n[!] No thread number specified (-t / --thread)")
    else:
        if args.thread > multiprocessing.cpu_count() :
            print(yellow,end="")
            print("[!] Reducing number of threads for avoiding Thread Trashing and System Failure")
            num_threads = int(multiprocessing.cpu_count()-1)
        else:
            num_threads = args.thread
    print(cyan, end="")
    print("[*] This program will use "+str(num_threads)+" threads\n")
    # for windows
    if sys.platform == "win32":
        print(green, end="")
        print("[√] Windows system detected")
        print(yellow, end='')
        print('[!] Checking for Administrator rights')
        if ctypes.windll.shell32.IsUserAnAdmin() == False:
            print(red, end='')
            print('[×] Administrator rights not found')
            print(light_red)
            print('[*] Please run this script as Administrator')
        else:
            print(green, end='')
            print('[√] Administrator rights found')
            print(yellow, end="")
            print("[!] Performing Integrity Check")
            report_path_status = os.path.exists(report_path)
            assets_path_status = os.path.exists(assets_path)
            assets_path_integrity = False
            if assets_path_status:
                assets_path_integrity = check_integrity(assets_path, folder_hash)
            if report_path_status == True and assets_path_status == True and assets_path_integrity == True:
                print(green, end="")
                print("[√] Integrity check succeeded")
            else:
                print(light_red, end="")
                print("[×] Integrity check failed or missing folder/files")
                if  report_path_status == False:
                    os.mkdir(report_path)
                print(yellow, end="")
                print("[!] Trying to extracting files")
                try:
                    from extractor import extractAssets
                    extractAssets(report_path)
                    print(green, end="")
                    print("[√] All files extracted successfully")
                except Exception as e:
                    print(red, end='')
                    print('[×] Error in extracting files')
                    if debug_mode:
                        print('Debug : '+str(e))
                    print('[*] Please clone the repository and try again')
                    print('[*] Exiting the program')
                    print(nocolor)
                    exit()
            if show_verbose:
                print(log_col, end="")
                scanStart_T = time.localtime()
                print("[v] Scan stared at "+str(time.strftime("%H:%M:%S", scanStart_T)))
            try :
                from headerSection import header_Report
                from sectionOne import osdetails_Report, system_model_Report
                from sectionTwo import processor_Report, driver_Report
                from sectionThree import memory_Report, local_Driver_Report
                from sectionFour import user_details_Report, printer_details_Report
                from sectionFive import controller_details_Report, audio_details_Report
                from sectionSix import antivirus_details_Report, display_details_Report
                from sectionSeven import network_report, other_devices_report
                from sectionEight import wsl_report, usb_report
                from sectionNine import hotfix_details_report
                from softwaresSection import startupSoftware_Report, installSoftware_Report
                from cisHeaderSection import makeCISHeader
                from cisSectionOne import AccountPolicy_Report
                from cisSectionTwo import LocalPolicies_Report
            except Exception as e:
                print (red, end='')
                print ('[×] Internal error occurred')
                if debug_mode:
                    print ('Debug : '+str(e))
                print ('[*] Exiting the program')
                print (nocolor)
                exit ()
            report = None
            computer_name = None
            report_header = None
            report_os = None
            report_sm = None
            report_pd = None
            report_dd = None
            report_mem = None
            report_l_drive = None
            report_user = None
            report_printer = None
            report_controller = None
            report_audio = None
            report_display = None
            report_virus = None
            report_hotfix = None
            report_other_devices = None
            report_wsl = None
            report_usb = None
            report_network= None
            # for software sections
            report_StartupSoftware = None
            report_installedSoftware = None
            # For Cis Sections
            CISsection1Score = 0
            CISsection1HTML = None
            CISsection2Score = 0
            CISsection2HTML = None
            # Put all the functions name in sequential order
            functions_list = [header_Report, osdetails_Report, user_details_Report, system_model_Report, processor_Report, driver_Report, memory_Report, local_Driver_Report, printer_details_Report, controller_details_Report, audio_details_Report, antivirus_details_Report, display_details_Report, hotfix_details_report, other_devices_report, wsl_report, usb_report, startupSoftware_Report, installSoftware_Report, network_report, AccountPolicy_Report, LocalPolicies_Report]
            output_collection = []
            with concurrent.futures.ThreadPoolExecutor(max_workers=num_threads) as executor:
                for func in functions_list:
                    future = executor.submit(func, show_verbose)
                    output_collection.append(future)
            for future in output_collection:
                if future.result() is not None:
                    if isinstance(future.result(), dict):
                        if len(future.result()) == 3:
                            report = future.result().get("report")
                            report_header = future.result().get("report_header")
                            computer_name = future.result().get("computer_name")
                        elif len(future.result()) == 2:
                            if "CISsection1Score" in future.result() or "CISsection1HTML" in future.result():
                                CISsection1Score = future.result().get("CISsection1Score")
                                CISsection1HTML = future.result().get("CISsection1HTML")
                            if "CISsection2Score" in future.result() or "CISsection2HTML" in future.result():
                                CISsection2Score = future.result().get("CISsection2Score")
                                CISsection2HTML = future.result().get("CISsection2HTML")
                        else :
                            if "report_os" in future.result():
                                report_os = future.result().get("report_os")
                            elif "report_sm" in future.result():
                                report_sm = future.result().get("report_sm")
                            elif "report_pd" in future.result():
                                report_pd = future.result().get("report_pd")
                            elif "report_dd" in future.result():
                                report_dd = future.result().get("report_dd")
                            elif "report_mem" in future.result():
                                report_mem = future.result().get("report_mem")
                            elif "report_l_drive" in future.result():
                                report_l_drive = future.result().get("report_l_drive")
                            elif "report_user" in future.result():
                                report_user = future.result().get("report_user")
                            elif "report_printer" in future.result():
                                report_printer = future.result().get("report_printer")
                            elif "report_controller" in future.result():
                                report_controller = future.result().get("report_controller")
                            elif "report_audio" in future.result():
                                report_audio = future.result().get("report_audio")
                            elif "report_virus" in future.result():
                                report_virus = future.result().get("report_virus")
                            elif "report_display" in future.result():
                                report_display = future.result().get("report_display")
                            elif "report_hotfix" in future.result():
                                report_hotfix = future.result().get("report_hotfix")
                            elif "report_other_devices" in future.result():
                                report_other_devices = future.result().get("report_other_devices")
                            elif "report_wsl" in future.result():
                                report_wsl = future.result().get("report_wsl")
                            elif "report_usb" in future.result():
                                report_usb = future.result().get("report_usb")
                            elif "report_StartupSoftware" in future.result():
                                report_StartupSoftware = future.result().get("report_StartupSoftware")
                            elif "report_installedSoftware" in future.result():
                                report_installedSoftware = future.result().get("report_installedSoftware")
                            elif "report_network" in future.result():
                                report_network = future.result().get("report_network")
            # End of the report
            print(yellow, end='')
            print('[!] Adding finishing touches')
            try:
                report = report + report_header + report_os + report_sm + report_pd + report_dd + report_mem + report_l_drive + report_user + report_printer + report_controller + report_audio + report_virus + report_display + report_network + report_other_devices + report_usb + report_wsl + report_hotfix
                # Software
                report += report_StartupSoftware + report_installedSoftware
                #Cis
                totalScore = (lambda s: round(s[0]/(100/10),1) + round(s[1]/(100/10),1))( (CISsection1Score, CISsection2Score) )
                report += makeCISHeader(totalScore, 2) + CISsection1HTML + CISsection2HTML
            except Exception as e:
                print(red, end='')
                print('[×] Error in merging report')
                if debug_mode:
                    print('Debug : '+str(e))
                print (nocolor)
                exit()
            report += """
                    </div>
            </section>
            <div class="blurb" style="padding-top:0"></div>
                        <div id="pageFooter" role="region">
                        Copyright under <a style="text-decoration: none;" href="test.com", target="_blank">MIT Licenses</a><br>
                        Made by Hrishikesh (<a style="text-decoration: none;" href="https://www.github.com/Hrishikesh7665", target="_blank">Hrishikesh7665</a>) & Shabdik (<a style="text-decoration: none;" href="https://www.github.com/ninja-hattori", target="_blank">ninja-hattori</a>)
                        <button onclick="scrollToTop()" id="scrollToTopBtn" title="Go to top">ADD L</button>
                        </div>
                    </div>
                </body>
            </html>"""
            if show_verbose:
                print(log_col, end="")
                scanEnd_T = time.localtime()
                print("[v] Scan completed at "+str(time.strftime("%H:%M:%S", scanEnd_T)))
            print(yellow, end='')
            print('[!] Generating final report')
            try:
                current_timeStamp = time.localtime()
                current_time = time.strftime("%H-%M-%S-%d-%m-%Y", current_timeStamp)
                if f_filename == None:
                    final_report_Name = report_path+str(re.sub(r'[^\w\s]', '_', computer_name.strip()))+"_final_report_"+str(current_time)+".html"
                else:
                    final_report_Name = report_path+f_filename
                report = report.replace("<title>ChangeTitle</title>", "<title>System Report of "+computer_name+"</title>")
                report = report.replace('<meta name="description" content="changeMe">', '<meta name="description" content="System Report of '+computer_name+' generated using System Scrutinizer tool on '+current_time+'">')
                with open(final_report_Name, "w") as f:
                    f.write(report)
                print(green, end='')
                print('[√] Final report generated successfully')
                print(cyan, end='')
                print('\n[*] Final report Saved at : ' +final_report_Name)
                if show_verbose:
                    print(log_col, end="")
                    print("\n[v] Report generated at "+str(time.strftime("%H:%M:%S", current_timeStamp)),end="")
                    print(" [time taken {:.0f} hour(s) {:02.0f} minute(s) {:02.0f} second(s)]".format(*map(int, str(datetime.datetime.fromtimestamp(time.mktime(current_timeStamp)) - datetime.datetime.fromtimestamp(time.mktime(scanStart_T))).split(":"))))
                    print("[v] Opening report in browser")
                webbrowser.open(final_report_Name)
            except Exception as e:
                print(red, end='')
                print('[×] Error in generating final report')
                if debug_mode:
                    print('Debug : '+str(e))
    # Uncomment the two lines below before make it executable
    print(cyan, end='')
    input("\nPress Enter to close the console...")
    print(nocolor)