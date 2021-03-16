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


async def save_to_html(row, worker):
    try:
        logging_worker = f" [Worker{worker}] "
        with open(html_file_path.replace(".html", f"_worker{worker}.html"), "a", newline='', encoding='utf-8') as html_file:
            if row:
                here_tz = tz.tzlocal()
                converted_dt = parser.parse(row["createdDateTime"])
                html_file.write(
                    f"""
<tr>
<td>{row['userDisplayName']}</td> 
<td>{row['userPrincipalName']}</td> 
<td class=date>{converted_dt.astimezone(here_tz).strftime("%Y-%m-%d %H:%M:%S")}</td>
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
                logging.error(f"{logging_worker}HTML: Row doesn't contain data.")
                   
    except Exception as e:
        logging.error(f"{logging_worker}Exception while generating HTML file: {e}, trying again.")
        await save_to_html(row)


async def save_to_csv(row, worker):
    try:
        logging_worker = f" [Worker{worker}] "
        with open(csv_file_path.replace(".csv", f"_worker{worker}.csv"), "a", newline='', encoding='utf-8-sig') as csv_file:
            csv_writer = csv.writer(csv_file)
            if row:
                here_tz = tz.tzlocal()
                converted_dt = parser.parse(row["createdDateTime"])
                csv_writer.writerow((
                    row["userDisplayName"],
                    row["userPrincipalName"],
                    converted_dt.astimezone(here_tz).strftime("%Y-%m-%d %H:%M:%S"),
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
                logging.error(f"{logging_worker}CSV: Row doesn't contain data.")

    except Exception as e:
        logging.error(f"{logging_worker}Exception while generating CSV file: {e}, trying again.")
        await save_to_csv(row)


async def get_graph_data(endpoint:str, token, session):
    async with session.get(
        endpoint,
        headers={
            'Authorization': 'Bearer ' + token['access_token']
    }) as response:
        
        if response.status == 200:
            return json.loads(await response.read())
        else:
            logging.error(f"Error getting signins {response.status} {response.reason} - The request endpoint was: '{endpoint}'")
            if "Retry-After" in response.headers:
                logging.warning(f"Request response has Retry-After header, probably we're being throttled, waiting {response.headers['Retry-After']} second(s).")
                time.sleep(int(response.headers["Retry-After"]))
            return await get_graph_data(endpoint, token, session)


async def get_data(endpoint, start_date, end_date, worker, token, session):

    async def get_graph_data_inner(endpoint_inner: str, token, session):
        graph_data_status = False
        while not graph_data_status:
            try:
                logging.info(f"{logging_worker}Getting page {page_counter}. Endpoint: '{endpoint_inner}'")
                try:
                    graph_data = await get_graph_data(endpoint_inner, token, session)
                except Exception as e:
                    logging.error(f"{logging_worker}Error getting graph data (1): {e}")
                    #graph_data = None
                if graph_data:
                    if "error" in graph_data:
                        logging.error(f"{logging_worker}graph_data with error. {graph_data['error']['code']} {graph_data['error']['message']}")
                        if graph_data['error']['code'] == "InvalidAuthenticationToken":
                            token = await renew_token()
                    else:
                        graph_data_status = True

            except Exception as e:
                logging.error(f"{logging_worker}Error getting graph data (2): {e}")

            if not graph_data_status:
                asyncio.sleep(10)
        
        return graph_data


    async def save_to_files_inner(graph_data, worker):
        try:
            tasks = []
            previous_row_str = None
            if graph_data:
                for row in graph_data["value"]:
                    if row:
                        this_row_str = json.dumps({k: row[k] for k in (
                            'userPrincipalName',
                            'appDisplayName',
                            'clientAppUsed',
                            'ipAddress'
                            )})
                        try:
                            if ("@tutor.uniasselvi.com.br" in row["userPrincipalName"] or 
                                "@aluno.uniasselvi.com.br" in row["userPrincipalName"]) and (this_row_str != previous_row_str):

                                tasks.append(asyncio.ensure_future(save_to_html(row, worker)))
                                tasks.append(asyncio.ensure_future(save_to_csv(row, worker)))

                                #if len(tasks) == 50:
                                #logging.info(f"Gathering page tasks")

                                previous_row_str = this_row_str

                                await asyncio.gather(*tasks, return_exceptions=True)

                        except Exception as e:
                            logging.error(f"{logging_worker}Error writing row data to files: {e}")

                        tasks = []
            else:
                logging.warning(f"{logging_worker}graph_data returned None somehow? (3)")

                
            #logging.info(f"Gathering page last tasks")
            #await asyncio.gather(*tasks, return_exceptions=True)

            logging.info(f"{logging_worker}HTML and CSV file appended.")
        
        except Exception as e:
            logging.error(f"{logging_worker}Error writing graph data in files: {e}")


    logging_worker = f" [Worker{worker}] "
    try:        
        request_filter = f"filter=createdDateTime ge {start_date.strftime('%Y-%m-%dT%H:%M:%SZ')} and createdDateTime le {end_date.strftime('%Y-%m-%dT%H:%M:%S.999999Z')}"
        request_order = "orderby=createdDateTime"
        endpoint_signIns = f"{endpoint}?&${request_filter}&${request_order}"

        logging.info(f"{logging_worker}Endpoint set as: '{endpoint_signIns}'")

        page_counter = 1

        '''
        graph_data_status = False
        while not graph_data_status:
            try:
                logging.info(f"{logging_worker}Getting page {page_counter}. Endpoint: '{endpoint_signIns}'")
                try:
                    graph_data = await get_graph_data(endpoint_signIns, token, session)
                except Exception as e:
                    logging.error(f"{logging_worker}Error getting graph data (1): {e}")
                    #graph_data = None
                if graph_data:
                    if "error" in graph_data:
                        logging.error(f"{logging_worker}graph_data with error. {graph_data['error']['code']} {graph_data['error']['message']}")
                        if graph_data['error']['code'] == "InvalidAuthenticationToken":
                            token = await renew_token()
                    else:
                        graph_data_status = True

            except Exception as e:
                logging.error(f"{logging_worker}Error getting graph data (2): {e}")

        try:
            tasks = []
            previous_row_str = None
            if graph_data:
                for row in graph_data["value"]:
                    if row:
                        this_row_str = json.dumps({k: row[k] for k in (
                            'userPrincipalName',
                            'appDisplayName',
                            'clientAppUsed',
                            'ipAddress'
                            )})
                        try:
                            if ("@tutor.uniasselvi.com.br" in row["userPrincipalName"] or 
                                "@aluno.uniasselvi.com.br" in row["userPrincipalName"]) and (this_row_str != previous_row_str):

                                tasks.append(asyncio.ensure_future(save_to_html(row, worker)))
                                tasks.append(asyncio.ensure_future(save_to_csv(row, worker)))

                                #if len(tasks) == 50:
                                #logging.info(f"Gathering page tasks")

                                previous_row_str = this_row_str

                                await asyncio.gather(*tasks, return_exceptions=True)

                        except Exception as e:
                            logging.error(f"{logging_worker}Error writing row data to files: {e}")

                        tasks = []
            else:
                logging.warning(f"{logging_worker}graph_data returned None somehow? (3)")

                
            #logging.info(f"Gathering page last tasks")
            #await asyncio.gather(*tasks, return_exceptions=True)

            logging.info(f"{logging_worker}HTML and CSV file appended.")

            page_counter += 1
        
        except Exception as e:
            logging.error(f"{logging_worker}Error writing graph data in files: {e}")
        '''
        
        graph_data = await get_graph_data_inner(endpoint_signIns, token, session)
        await save_to_files_inner(graph_data, worker)

        page_counter += 1

        while "@odata.nextLink" in graph_data:
            next_link = str(graph_data["@odata.nextLink"])

            if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
                token = await renew_token() 

            graph_data = await get_graph_data_inner(next_link, token, session)
            await save_to_files_inner(graph_data, worker)

            page_counter += 1

            '''
            logging.info(f"{logging_worker}Getting page {page_counter}. Endpoint: '{next_link}'")              
            try:
                graph_data = await get_graph_data(next_link, token, session)
            except Exception as e:
                logging.error(f"{logging_worker}Error getting graph data: {e}")
                graph_data = None

            while "error" in graph_data or None == graph_data:  
                logging.error(f"{logging_worker}graph data with error or empty. {graph_data['error']['code']} {graph_data['error']['message']}")          
                if graph_data['error']['code'] == "InvalidAuthenticationToken":
                    token = await renew_token()
                try:
                    graph_data = await get_graph_data(next_link, token, session)
                except Exception as e:
                    logging.error(f"{logging_worker}Error getting graph data: {e}")
                    graph_data = None
            

            try:
                tasks = []
                previous_row_str = None
                if graph_data:
                    for row in graph_data["value"]:
                        if row:
                            this_row_str = json.dumps({k: row[k] for k in (
                                'userPrincipalName',
                                'appDisplayName',
                                'clientAppUsed',
                                'ipAddress'
                                )})
                            try:
                                if ("@tutor.uniasselvi.com.br" in row["userPrincipalName"] or 
                                    "@aluno.uniasselvi.com.br" in row["userPrincipalName"]) and (this_row_str != previous_row_str):
                                    
                                    tasks.append(asyncio.ensure_future(save_to_html(row, worker)))
                                    tasks.append(asyncio.ensure_future(save_to_csv(row, worker)))

                                    previous_row_str = this_row_str

                                    #if len(tasks) == 50:
                                        #logging.info(f"Gathering page tasks")

                                    await asyncio.gather(*tasks, return_exceptions=True) 

                            except Exception as e:
                                logging.error(f"{logging_worker}Error writing row data to files: {e}")

                            tasks = []
                else:
                    logging.warning(f"{logging_worker}graph_data returned None. Possibly correlated with [WinError 121] The semaphore timeout period has expired")
                    
                
                #logging.info(f"Gathering page last tasks")
                #await asyncio.gather(*tasks, return_exceptions=True)

                logging.info(f"{logging_worker}HTML and CSV file appended.")

                page_counter += 1
            
            except Exception as e:
                logging.error(f"{logging_worker}Error writing graph data in files: {e}")
            '''

        
        logging.info(f"{logging_worker}Finished getting and writing data.")
        return True
        
    except Exception as e:
        logging.error(f"{logging_worker}Exception getting report data: {e}")
        return False 


async def gather_files(workers):

    header_columns = [
        "Nome",
        "E-mail",
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

    with open(csv_file_path, "w", newline='', encoding='utf-8-sig') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(header_columns)

    # Appending files.
    logging.info(f"Consolidating files into a single one of each type.")
    try:  
        html_file = open(html_file_path, 'a+')
        csv_file = open(csv_file_path, 'a+', encoding='utf-8-sig')
        for worker in range(workers):
            worker_html_file = html_file_path.replace(".html", f"_worker{worker}.html")
            temp_file = open(worker_html_file, 'r') 
            html_file.write(temp_file.read())
            temp_file.close()
            os.remove(worker_html_file)

            worker_csv_file = csv_file_path.replace(".csv", f"_worker{worker}.csv")
            temp_file = open(worker_csv_file, 'r') 
            csv_file.write(temp_file.read())
            temp_file.close()
            os.remove(worker_csv_file)

        html_file.close()  
        csv_file.close()
    except Exception as e:
        logging.error(f"Exception consolidating files: {e}")

    # Close HTML file tags
    with open(html_file_path, "a", newline='', encoding='utf-8') as html_file:
        html_file.write("</table></body></html>")
        
    logging.info(f"Finished getting results and consolidating data, everything exported to CSV and HTML files.")


    #ps_script_path = os.path.join(current_wdpath, "ConvertTo-ExcelCustomReportHTML.ps1")
    #ps_xlsx_path = os.path.join(current_wdpath, output_files_fname, f"AuditoriaEntrada_{yesterday.strftime('%d-%m-%Y')}_Completo_{str(uuid.uuid4())}.xlsx")
    #ps_args = f"{ps_script_path} -HtmlPath {html_file_path} -XlsxPath {ps_xlsx_path}"

    ps_script_path = os.path.join(current_wdpath, "Send-Reports.ps1")
    ps_args = f"{ps_script_path} -CsvPath {csv_file_path}"
    await call_ps(ps_args)




async def start_report_gathering(token):
    try:

        workers = 24

        tasks = []
        start_datetime = yesterday.replace(hour=3, minute=0, second=0)

        async with aiohttp.ClientSession() as session:
            for i in range(workers):
                try:
                    if i == 0:
                        task_start_datetime = start_datetime
                        next_task_start_datetime = start_datetime + datetime.timedelta(hours=(24 / workers))
                        task_end_datetime = next_task_start_datetime - datetime.timedelta(seconds=1)
                    else:
                        task_start_datetime = next_task_start_datetime
                        next_task_start_datetime = next_task_start_datetime + datetime.timedelta(hours=(24 / workers))
                        task_end_datetime = next_task_start_datetime - datetime.timedelta(seconds=1)

                    logging.info(f"Worker number: {i} - {task_start_datetime.strftime('%Y-%m-%dT%H:%M:%SZ')} to {task_end_datetime.strftime('%Y-%m-%dT%H:%M:%S.999999Z')}")
                    tasks.append(asyncio.ensure_future(get_data(config["endpoint_signIns"], task_start_datetime, task_end_datetime, i, token, session)))
                except Exception as e:
                    logging.error(f"Error during worker/task creation. Worker #{i} - Error: {e}")

            results = await asyncio.gather(*tasks, return_exceptions=True)

        #if not False in results:
        await gather_files(workers)
        #print(f"{start_datetime2.strftime('%Y-%m-%dT%H:%M:%SZ')}")
    except Exception as e:
        logging.error(f"Error in start_report_gathering function: {e}")       


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

        yesterday = datetime.datetime.today() - datetime.timedelta(days=6)

        # HTML File
        html_template_path = os.path.join(current_wdpath, "template.html")
        html_file_path = os.path.join(current_wdpath, output_files_fname, f"auditSignIns_{yesterday.strftime('%Y-%m-%d')}_generated_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.html")

        copyfile(html_template_path, html_file_path)

        # CSV File
        csv_file_path = html_file_path.replace(".html", ".csv")

        # Logging
        log_filename_datetime = datetime.datetime.now()
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(os.path.join(current_wdpath, f"debug_{log_filename_datetime.strftime('%Y-%m-%d_%H%M%S')}.log")),
                logging.StreamHandler()
            ]
        )   

        s = time.perf_counter()
        asyncio.get_event_loop().run_until_complete(main())
        elapsed = time.perf_counter() - s
        logging.info(f"Script finished, executed in {elapsed:0.2f} seconds.")
    except Exception as e:
        logging.error(f"Error: {str(e)}")
