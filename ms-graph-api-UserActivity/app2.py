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


async def get_token():

    client_id = "335e1303-eec5-4f9b-b251-adf9b63c8c72"

    client_secret_var = "daemon_client_secret"
    client_secret = os.getenv(client_secret_var)
    if not client_secret:
        errmsg = f"Define {client_secret_var} environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)
    else:
        logging.info(f"{client_secret_var} found.")    

    # Create a preferably long-lived app instance which maintains a token cache.
    app = msal.ConfidentialClientApplication(
        client_id, authority=config["authority"],
        client_credential=client_secret,
        # token_cache=...  # Default cache is in memory only.
                        # You can learn how to use SerializableTokenCache from
                        # https://msal-python.rtfd.io/en/latest/#msal.SerializableTokenCache
        )

    # The pattern to acquire a token looks like this.
    result = None

    # Since we are looking for token for the current app, NOT for an end user,
    # notice we give account parameter as None.
    result = app.acquire_token_silent(config["scope"], account=None)

    if not result:
        logging.info("No token exists in cache. Getting a new one from AzureAD.")
        result = app.acquire_token_for_client(scopes=config["scope"])

    result["renew_datetime"]  = datetime.datetime.now() + datetime.timedelta(seconds=result["expires_in"])
    
    return result

async def renew_token():
    logging.info("Getting new token from AzureAD.")
    result = await get_token()
    while not "access_token" in result:
        logging.error(f"Error getting new token: {result.get('error')}: {result.get('error_description')} (correlation_ID: {result.get('correlation_id')})")                
        time.sleep(30)
        result = await get_token()
    
    return result

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


async def save_to_html(row):
    try:
        with open(html_file_path, "a", newline='', encoding='utf-8') as html_file:
            if row:
                here_tz = tz.tzlocal()
                converted_dt = parser.parse(row["createdDateTime"])
                html_file.write(
                    f"""
<tr>
<td>{row['userDisplayName']}</td> 
<td>{row['userPrincipalName']}</td> 
<td class=date>{converted_dt.astimezone(here_tz).strftime("%Y-%m-%d %H:%M:%S,%f")}</td>
<td>{row['appDisplayName']}</td> 
<td>{row['clientAppUsed']}</td> 
<td>{row['deviceDetail']['browser']}</td> 
<td>{row['deviceDetail']['operatingSystem']}</td> 
<td>&nbsp;{row['ipAddress']}</td> 
<td>{row['location']['city']}</td> 
<td>{row['location']['state']}</td> 
<td>{row['location']['countryOrRegion']}</td> 
</tr>
""")
            else:
                logging.error("HTML: Row doesn't contain data.")
                   
    except Exception as e:
        logging.error(f"Exception while generating HTML file: {str(e)}")


async def save_to_csv(row):
    try:
        with open(csv_file_path, "a", newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            if row:
                csv_writer.writerow((
                    row["userDisplayName"],
                    row["userPrincipalName"],
                    row["createdDateTime"],
                    row["appDisplayName"],
                    row["clientAppUsed"],
                    row["deviceDetail"]["browser"],
                    row["deviceDetail"]["operatingSystem"],
                    row["ipAddress"],
                    row["location"]["city"],
                    row["location"]["state"],
                    row["location"]["countryOrRegion"]
                ))
            else:
                logging.error("CSV: Row doesn't contain data.")

    except Exception as e:
        logging.error(f"Exception while generating CSV file: {str(e)}")


async def get_graph_data(endpoint, token, session):
    async with session.get(
        endpoint,
        headers={
            'Authorization': 'Bearer ' + token['access_token']
    }) as response:
        
        if response.status == 200:
            return json.loads(await response.read())
        else:
            logging.error(f"Error getting sigins {response.status} {response.reason}")
            if "Retry-After" in response.headers:
                logging.warning(f"Request response has Retry-After header, probably we're being throttled, waiting {response.headers['Retry-After']} second(s).")
                time.sleep(int(response.headers["Retry-After"]))
            return await get_graph_data(endpoint, token, session)


async def get_data(endpoint, token, session):
    try:
        request_filter = f"filter=createdDateTime ge {yesterday.strftime('%Y-%m-%d')}T03:00:00Z and createdDateTime le {datetime.datetime.today().strftime('%Y-%m-%d')}T03:00:00Z"
        request_order = "orderby=createdDateTime"
        endpoint_signIns = f"{endpoint}?&${request_filter}&${request_order}"

        logging.info(f"Endpoint set as: '{endpoint_signIns}'")

        page_counter = 1

        logging.info(f"Getting page {page_counter}")

        graph_data = await get_graph_data(endpoint_signIns, token, session)
        while "error" in graph_data or None == graph_data:  
            logging.error(f"graph data with error or empty. {graph_data['error']['code']} {graph_data['error']['message']}")          
            if graph_data['error']['code'] == "InvalidAuthenticationToken":
                token = await renew_token()
            graph_data = await get_graph_data(endpoint_signIns, token, session)

        tasks = []
        for row in graph_data["value"]:
            if row:
                tasks.append(asyncio.ensure_future(save_to_html(row)))
                tasks.append(asyncio.ensure_future(save_to_csv(row)))

            #if len(tasks) == 50:
                #logging.info(f"Gathering page tasks")
                await asyncio.gather(*tasks, return_exceptions=True)
                tasks = []

            
        #logging.info(f"Gathering page last tasks")
        #await asyncio.gather(*tasks, return_exceptions=True)

        logging.info("HTML and CSV file appended.")

        page_counter += 1

        while "@odata.nextLink" in graph_data:
            next_link = graph_data["@odata.nextLink"]

            if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
                token = await renew_token() 

            logging.info(f"Getting page {page_counter}")               

            graph_data = await get_graph_data(next_link, token, session)
            while "error" in graph_data or None == graph_data:  
                logging.error(f"graph data with error or empty. {graph_data['error']['code']} {graph_data['error']['message']}")          
                if graph_data['error']['code'] == "InvalidAuthenticationToken":
                    token = await renew_token()
                graph_data = await get_graph_data(next_link, token, session)
            
            tasks = []
            for row in graph_data["value"]:
                if row:
                    tasks.append(asyncio.ensure_future(save_to_html(row)))
                    tasks.append(asyncio.ensure_future(save_to_csv(row)))
                
                #if len(tasks) == 50:
                    #logging.info(f"Gathering page tasks")
                    await asyncio.gather(*tasks, return_exceptions=True)
                    tasks = []
                
            
            #logging.info(f"Gathering page last tasks")
            #await asyncio.gather(*tasks, return_exceptions=True)

            logging.info("HTML and CSV file appended.")

            page_counter += 1
        
        return True
        
    except Exception as e:
        logging.error(f"Exception while generating CSV file: {str(e)}")
        return False     

async def start_report_gathering(token):    

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


    async with aiohttp.ClientSession() as session:
        #await get_data(token, session)
        #tasks = []
        #tasks.append(asyncio.ensure_future(get_data(config["endpoint_signIns"], token, session)))
        #result = await asyncio.gather(*tasks, return_exceptions=True)
        result = await get_data(config["endpoint_signIns"], token, session)        

        if result:
            # Close html file.
            with open(html_file_path, "a", newline='', encoding='utf-8') as html_file:
                html_file.write("</table></body></html>")
            
            logging.info(f"Finished getting result pages, everything exported to CSV and HTML files '{csv_file_path}'.")

            ps_script_path = os.path.join(current_wdpath, "ConvertTo-ExcelCustomReportHTML.ps1")
            #ps_html_path = os.path.join(current_wdpath, "teste.html")
            ps_xlsx_path = os.path.join(current_wdpath, output_files_fname, f"AuditoriaEntrada_{yesterday.strftime('%d-%m-%Y')}_Completo_{str(uuid.uuid4())}.xlsx")

            ps_args = f"{ps_script_path} -HtmlPath {html_file_path} -XlsxPath {ps_xlsx_path}"
            await call_ps(ps_args)      


async def main():

    token = await get_token()

    if "access_token" in token:
        await start_report_gathering(token)
    else:
        logging.error(f"{token.get('error')}: {token.get('error_description')} (correlation_ID: {token.get('correlation_id')})")
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))  # You may need this when reporting a bug


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
        html_template_path = os.path.join(current_wdpath, "template.html")
        html_file_path = os.path.join(current_wdpath, output_files_fname, f"auditSignIns_{yesterday.strftime('%Y-%m-%d')}_generated_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.html")

        copyfile(html_template_path, html_file_path)

        # CSV File
        csv_file_path = html_file_path.replace(".html", ".csv")

        # Logging
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(os.path.join(current_wdpath, f"debug_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.log")),
                logging.StreamHandler()
            ]
        )   

        s = time.perf_counter()
        asyncio.get_event_loop().run_until_complete(main())
        elapsed = time.perf_counter() - s
        logging.info(f"Script finished, executed in {elapsed:0.2f} seconds.")
    except Exception as e:
        logging.error(f"Error: {str(e)}")
