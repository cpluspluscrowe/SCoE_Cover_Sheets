import os
import subprocess
path = r"C:\SCoE"
for root,dirs,files in os.walk(path):
    for file in files:
        if file == "Cover Sheet.html":
            fullpath = os.path.join(root,file)
            fullPdfPath = fullpath.replace(".html",".pdf")
            command = r'"C:\Users\CCrowe\AppData\Local\Google\Chrome SxS\Application\chrome.exe" --headless --disable-gpu --landscape  --print-to-pdf="{0}" "{1}"'.format(fullPdfPath,fullpath)
            subprocess.check_output(command)
            
