"""
The configuration file would look like this (sans those // comments):

{
    "authority": "https://login.microsoftonline.com/Enter_the_Tenant_Name_Here",
    "scope": ["https://graph.microsoft.com/.default"],
        // For more information about scopes for an app, refer:
        // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow#second-case-access-token-request-with-a-certificate"
}

You can then run this sample with a JSON configuration file:

    python app.py parameters.json
"""

import os
import sys  # For simplicity, we'll read config file from 1st CLI param sys.argv[1]
import subprocess, sys
import json
import logging
import datetime
import uuid
import pathlib
import glob
import uuid
import re
import csv
from shutil import copyfile
from dateutil import parser, tz

import aiohttp
import msal

import asyncio
import time



async def call_ps(ps_args):
    logging.info(f"Calling PS.")
    try:
        p = subprocess.Popen([
            "powershell.exe", "-NoProfile", f"{ps_args}"],
            stdout=subprocess.PIPE)
        output, err = p.communicate()
        return output

    except Exception as e:
        logging.error(f"Exception while calling PowerShell: {str(e)}")
        return None


async def gather_files(workers):

    header_columns = [
        "Nome",
        "E-mailUniasselvi",
        "DataDeEntrada",
        "AplicativoMicrosoft",
        "AplicativoClienteUtilizado",
        "Navegador",
        "SistemaOperacional",
        "IPAddress",
        "Cidade",
        "Estado",
        "País"
    ]

    logging.info(f"Creating column headers in HTML file '{html_file_path}'.")

    with open(html_file_path, "a", newline='', encoding='utf-8') as html_file:
        html_file.write("<tr class=header>")    
        for header_name in header_columns:
            html_file.write(f"<td>{header_name}</td>")
        html_file.write("</tr>")

    logging.info(f"Creating header row in CSV file '{csv_file_path}'.")    

    with open(csv_file_path, "w", newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(header_columns)

    # Appending files.
    logging.info(f"Consolidating files into a single one of each type.")
    try:  
        html_file = open(html_file_path, 'a+')
        csv_file = open(csv_file_path, 'a+')
        for worker in range(workers):
            temp_file = open(html_file_path.replace(".html", f"_worker{worker}.html"), 'r') 
            html_file.write(temp_file.read())
            temp_file.close()

            temp_file = open(csv_file_path.replace(".csv", f"_worker{worker}.csv"), 'r') 
            csv_file.write(temp_file.read())
            temp_file.close()

        html_file.close()  
        csv_file.close()
    except Exception as e:
        logging.error(f"Exception consolidating files: {e}")

    # Close HTML file tags
    with open(html_file_path, "a", newline='', encoding='utf-8') as html_file:
        html_file.write("</table></body></html>")
        
    logging.info(f"Finished getting results and consolidating data, everything exported to CSV and HTML files.")

    ps_script_path = os.path.join(current_wdpath, "ConvertTo-ExcelCustomReportHTML.ps1")
    #ps_html_path = os.path.join(current_wdpath, "teste.html")
    ps_xlsx_path = os.path.join(current_wdpath, output_files_fname, f"AuditoriaEntrada_{yesterday.strftime('%d-%m-%Y')}_Completo_{str(uuid.uuid4())}.xlsx")

    ps_args = f"{ps_script_path} -HtmlPath {html_file_path} -XlsxPath {ps_xlsx_path}"
    await call_ps(ps_args)       


async def main():

    workers = 24
    await gather_files(workers)


if __name__ == "__main__":

    try:
        config = json.load(open(sys.argv[1]))

        # Current script path
        current_wdpath = os.path.dirname(__file__)
        output_files_fname = "output-files"

        # Creates dir if does not exist.
        pathlib.Path(os.path.join(current_wdpath, output_files_fname)).mkdir(exist_ok=True)

        yesterday = datetime.datetime.today() - datetime.timedelta(days=1)

        # HTML File
       # html_template_path = os.path.join(current_wdpath, "template.html")
        html_file_path = os.path.join(current_wdpath, output_files_fname, "auditSignIns_2021-03-10_generated_2021-03-11_143714.html")

        #copyfile(html_template_path, html_file_path)

        # CSV File
        csv_file_path = html_file_path.replace(".html", ".csv")

        # Logging
        log_filename_datetime = datetime.datetime.now()
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(os.path.join(current_wdpath, f"debug_gather_files_{log_filename_datetime.strftime('%Y-%m-%d_%H%M%S')}.log")),
                logging.StreamHandler()
            ]
        )   

        s = time.perf_counter()
        asyncio.get_event_loop().run_until_complete(main())
        elapsed = time.perf_counter() - s
        logging.info(f"Script finished, executed in {elapsed:0.2f} seconds.")
    except Exception as e:
        logging.error(f"Error: {str(e)}")
