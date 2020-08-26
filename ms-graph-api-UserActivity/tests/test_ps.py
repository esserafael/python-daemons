import os
import subprocess

# Current script path
current_wdpath = os.path.dirname(__file__)

ps_script_path = os.path.join(current_wdpath, "ConvertTo-ExcelCustomReportHTML.ps1")
ps_html_path = os.path.join(current_wdpath, "teste.html")
ps_xlsx_path = os.path.join(current_wdpath, "teste.xlsx")

ps_arg = f"{ps_script_path} -HtmlPath {ps_html_path} -XlsxPath {ps_xlsx_path}"

subprocess.Popen([
    "powershell.exe",
    f"{ps_arg}"
    ])