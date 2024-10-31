# -*- coding: utf-8 -*-
import subprocess
import re

def get_username_from_sid(sid, total_name=False):
    formatted_string = sid.lstrip('*')
    if re.match(r'^S-\d-\d+-\d+-\d+$', formatted_string) or re.match(r'^[Ss]-\d-\d+(-\d+)+$', formatted_string):
        if formatted_string == "S-1-5-32-551":
            return "Backup Operators"
        elif formatted_string == "S-1-5-32-555":
            return "Builtin\\Remote Desktop Users" if total_name else "Remote Desktop Users"
        else:
            command = f'$username = (New-Object System.Security.Principal.SecurityIdentifier "{formatted_string}").Translate([System.Security.Principal.NTAccount]).Value; Write-Output "Username: $username"'
            result = subprocess.run(['powershell', '-Command', command], capture_output=True, text=True)
            output = result.stdout.strip()
            if output.startswith("Username: "):
                username = output.split(": ")[1]
                if total_name:
                    return username.strip()
                else:
                    try:
                        username_trimmed = username.split('\\')[-1]
                        return username_trimmed.strip()
                    except:
                        return username.strip()
    else:
        return sid

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")