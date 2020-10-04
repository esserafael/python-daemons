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
import csv
import logging
import datetime
import uuid
import pathlib
import glob
import uuid

import aiohttp
import msal

import asyncio
import time

    
async def call_ps(ps_args):
    logging.info(f"Calling PS command: {ps_args}")
    try:
        p = subprocess.Popen([
            "powershell.exe", "-NoProfile", f"{ps_args}"],
            stdout=subprocess.PIPE)
        output, err = p.communicate()
        return output

    except Exception as e:
        logging.error(f"Exception while calling PowerShell: {str(e)}")
        return None

async def call_zabbix_sender(host, key, value):
    logging.info(f"Calling Zabbix sender.")
    try:
        p = subprocess.Popen([
            f"{zabbix_sender_path}", "-z", f"{config['zabbix_server']}", "-s", f"{host}", "-k", f"{key}", "-o", f"{value}"],
            stdout=subprocess.PIPE)
        output, err = p.communicate()
        return output

    except Exception as e:
        logging.error(f"Exception while calling Zabbix sender: {str(e)}")
        return None

async def get_graph_data(endpoint, token, session):
    async with session.get(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']}, ) as graph_data:
            #print(graph_data)
            if graph_data.status == 200:
                return await graph_data.read()
            else:
                return None

async def get_token():
    client_id = os.getenv("daemon_client_id3")
    if not client_id:
        errmsg = "Define daemon_client_id3 environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)    
    else:
        logging.info("client_id found -> '{0}'.".format(client_id))

    client_secret = os.getenv("daemon_client_secret3")
    if not client_secret:
        errmsg = "Define daemon_client_secret3 environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)
    else:
        logging.info("client_secret found.")


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
    
    return result

async def get_org_data(token, session):
    try:
        logging.info(config["endpoint_org"])
        org_graph_data = json.loads(await get_graph_data(config["endpoint_org"], token, session))
        
        if None != org_graph_data:

            org_object_limit = org_graph_data["value"][0]["directorySizeQuota"]["total"]
            org_used_objects = org_graph_data["value"][0]["directorySizeQuota"]["used"]

            output = await call_zabbix_sender("azuread", "windowsazuread.objectlimit", org_object_limit)
            logging.info(f"Zabbix sender return: {output}")
            output = await call_zabbix_sender("azuread", "windowsazuread.usedobjects", org_used_objects)
            logging.info(f"Zabbix sender return: {output}")

            if org_used_objects >= ((config['azuread_usedquota_alert_percentage'] / 100) * org_object_limit):

                alert_title = "\"Warning: We\'re approaching max AAD objects limit.\""
                alert_content = f"""
                <tr><td style=`\"padding: 0 0 20px`\">We are above <b>95%</b> of the Azure AD objects quota used.</td></tr>
                <tr><td style=`\"padding: 0 0 20px`\">Current number of used objects: <b>{org_used_objects}</b></td></tr>
                <tr><td style=`\"padding: 0 0 20px`\">Total objects quota (limit): <b>{org_object_limit}</b></td></tr>
                <tr><td style=`\"padding: 0 0 20px`\"><b>Suggested actions:</b> Remove stale/unused objects or request Microsoft Support to increase the limit.</td></tr>"""

                logging.info(alert_title)

                ps_args = f"{ps_alert_script_path} -To {config['alerts_recipient']} -Title {alert_title} -Content \"{alert_content}\""
                await call_ps(ps_args)
            
    except Exception as e:
        logging.error(f"Exception while calling get_org_data: {str(e)}")  


async def start_org_monitor(token):
    async with aiohttp.ClientSession() as session:
        tasks = []
        task = asyncio.ensure_future(get_org_data(token, session))
        tasks.append(task)

        await asyncio.gather(*tasks, return_exceptions=True)


async def main():

    token = await get_token()

    if "access_token" in token:
        await start_org_monitor(token)
    else:
        logging.error(f"{token.get('error')}: {token.get('error_description')} (correlation_ID: {token.get('correlation_id')})")
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))  # You may need this when reporting a bug


if __name__ == "__main__":

    import time

    try:
        config = json.load(open(sys.argv[1]))

        # Current script path
        current_wdpath = os.path.dirname(__file__)
        ps_alert_script_path = os.path.join(current_wdpath, "Send-AlertMailMessage.ps1")
        zabbix_sender_path = os.path.join(current_wdpath, "zabbix_sender.exe") 

        # Logging
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(os.path.join(current_wdpath, "debug.log")),
                logging.StreamHandler()
            ]
        )   

        s = time.perf_counter()
        asyncio.get_event_loop().run_until_complete(main())
        elapsed = time.perf_counter() - s
        logging.info(f"Script finished, executed in {elapsed:0.2f} seconds.")
    except Exception as e:
        logging.error(f"Error: {str(e)}")
