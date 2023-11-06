import subprocess
import sys
import os

# Check if a URL parameter is provided
if len(sys.argv) < 2:
    print("Usage: python script.py <URL>")
    sys.exit(1)

# Get the URL from the command line argument
url = sys.argv[1]

# Determine the user's temporary folder path
temp_folder = os.path.join(os.environ['userprofile'], 'AppData', 'Local', 'Temp')

# Path to the VBScript file in the user's temporary folder
vbs_script_path = os.path.join(temp_folder, 'script.vbs')

# VBScript code
vbs_script = f'''
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.ToolBar = False
ie.Navigate "{url}"
'''

# Save the VBScript code to the temporary folder
with open(vbs_script_path, 'w') as vbs_file:
    vbs_file.write(vbs_script)

# Run the VBScript code using Python
subprocess.call(['cscript.exe', vbs_script_path])

# Remove the .vbs file from the temporary folder
#os.remove(vbs_script_path)

