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
    client_id = os.getenv("daemon_client_id2")
    if not client_id:
        errmsg = "Define daemon_client_id2 environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)    
    else:
        logging.info("client_id found -> '{0}'.".format(client_id))

    client_secret = os.getenv("daemon_client_secret2")
    if not client_secret:
        errmsg = "Define daemon_client_secret2 environment variable"
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


async def set_user_pic(user, token, session):
    endpoint_ProfilePic = f"{config['endpoint_ProfilePic']}/{user['UserPrincipalName']}/photos/{config['pic_size']}/$value"
    graph_data = await get_graph_data(endpoint_ProfilePic, token, session)

    # Check if AD attribute is empty or user changed the profile pic.
    if (None != graph_data and None == user['thumbnailPhoto']) or (bytearray(user['thumbnailPhoto']) != bytearray(graph_data)):
        logging.info(f"Setting picture for user {user['UserPrincipalName']}.")
        try:
            cache_file_name = None
            cache_file_name = f"{os.path.join(cache_file_folder, str(uuid.uuid4()))}.jpg"
        except Exception as e:
            logging.error(f"Exception while generating cache file name: {str(e)}")

        if None != cache_file_name:

            with (open(cache_file_name, "wb")) as cache_file:
                cache_file.write(graph_data)

            ps_args = f"Set-ADUser -Identity \"{user['SamAccountName']}\" -Replace @{{thumbnailPhoto=([byte[]](Get-Content \"{cache_file_name}\" -Encoding Byte))}} -Server {config['prefered_dc']}\n"

            with open(ps_file_name, "a") as ps_file:
                ps_file.write(ps_args)     


async def set_all_users_pics(ad_users, token):
    async with aiohttp.ClientSession() as session:
        tasks, temp_ad_users = [], []
        max_usersget, current_usercount = 500, 0
        temp_all_sam = ""
        for user in ad_users:
            current_usercount += 1
            if current_usercount < max_usersget:
                temp_all_sam += f"\"{user['SamAccountName']}\","
            else:                
                temp_all_sam += f"\"{user['SamAccountName']}\""
                print("Chegou")
                temp_ad_users = json.loads(await call_ps(f"{temp_all_sam} | Get-ADUser -Server {config['prefered_dc']} -Properties thumbnailPhoto | Select-Object SamAccountName, UserPrincipalName, thumbnailPhoto | Sort-Object UserPrincipalName | ConvertTo-Json -Compress"))

                for inside_user in temp_ad_users:
                    #task = asyncio.ensure_future(set_user_pic(user, token, session))
                    task = asyncio.ensure_future(set_user_pic(inside_user, token, session))
                    tasks.append(task)

                current_usercount = 0
                temp_all_sam = ""            
            
        await asyncio.gather(*tasks, return_exceptions=True)

        if os.path.exists(ps_file_name):
            logging.info("Starting PowerShell processing, this may take a while.")
            await call_ps(f"Get-Content \"{ps_file_name}\" | ForEach-Object {{Invoke-Expression $_}}")
            logging.info("Finished PowerShell processing.")


async def main():

    if(os.path.exists(cache_file_folder)):
        # Clears the cache folder
        files = glob.glob(f"{cache_file_folder}/*")
        for f in files:
            os.remove(f)
    else:
        # Creates the cache folder if does not exist.
        pathlib.Path(cache_file_folder).mkdir(exist_ok=True)

    token = await get_token()

    if "access_token" in token:
        try:
            #teste1 = await call_ps(f"Get-ADUser -Server {config['prefered_dc']} -Filter {{Enabled -eq $true}} -SearchBase \"{config['base_search']}\" -Properties thumbnailPhoto | Select-Object SamAccountName, UserPrincipalName, thumbnailPhoto | Sort-Object UserPrincipalName | ConvertTo-Json -Compress")
            ad_users = json.loads(await call_ps(f"Get-ADUser -Server {config['prefered_dc']} -Filter {{Enabled -eq $true}} -SearchBase \"{config['base_search']}\" | Select-Object SamAccountName | Sort-Object SamAccountName | ConvertTo-Json -Compress"))
            #teste1 = await call_ps(f"Invoke-Command -ComputerName {config['prefered_dc']} -Session (Get-PSSession -Name {ps_session['Name']}) -ScriptBlock {{Get-ADUser -Filter {{Enabled -eq $true}} -SearchBase \"{config['base_search']}\" | Select-Object SamAccountName | Sort-Object SamAccountName}} | ConvertTo-Json -Compress")
        except Exception as e:
            logging.error(f"Exception while getting AD Users: {str(e)}")
        #ad_users = json.loads(await call_ps(f"Get-ADUser -Server {config['prefered_dc']} -Filter {{Enabled -eq $true}} -SearchBase \"{config['base_search']}\" -Properties thumbnailPhoto | Select-Object SamAccountName, UserPrincipalName, thumbnailPhoto | Sort-Object UserPrincipalName | ConvertTo-Json -Compress"))
        #ad_users = json.loads(teste1)
        await set_all_users_pics(ad_users, token)
    else:
        logging.error("{0}: {1} (correlation_id: {3})".format(token.get("error"), token.get("error_description"), token.get("correlation_id")))
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))  # You may need this when reporting a bug


if __name__ == "__main__":
    import time

    try:

        config = json.load(open(sys.argv[1]))

        # Current script path
        current_wdpath = os.path.dirname(__file__)
        cache_file_folder = os.path.join(current_wdpath, "cache-files") 
        ps_file_name = f"{os.path.join(cache_file_folder, str(uuid.uuid4()))}"

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
