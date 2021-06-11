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

import aiohttp
import msal

import asyncio
import time


async def get_token():

    client_id = "d7b889ea-d93d-48bf-8589-196c00c48421"

    client_secret = ""


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

async def get_team(id, token, session):
    async with session.get(
        f"https://graph.microsoft.com/beta/teams/{id}",
        headers={
            'Authorization': 'Bearer ' + token['access_token']
    }) as response_team:

        if response_team.status == 200:
            #team_data = json.loads(await response_team.read())
            #if team_data["specialization"] == "educationClass" and team_data["isMembershipLimitedToOwners"]:
            #    return team_data
            return json.loads(await response_team.read())
        elif response_team.status == 404:
            logging.warning(f"Team not found {response_team.status} {response_team.reason}. id: {id}")
            #return None
        else:
            logging.error(f"Some errors getting a team {response_team.status} {response_team.reason}")

async def process_groups_page(groups_data, token, session):
    teams_not_activated = []
    for group in groups_data["value"]:
        tasks = []
        task = asyncio.ensure_future(get_team(group['id'], token, session))
        tasks.append(task)        
       
        teams_not_activated.append(await asyncio.gather(*tasks, return_exceptions=True))
        #teams_not_activated2 = [team for team in teams_not_activated if team]
        #teams_not_activated.append(await get_team(group['id'], token, session))
    
    print(f"Found: {len(teams_not_activated)}")

    with open(csv_file_path, "a", newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        for team in teams_not_activated:
            if team[0]:
                csv_writer.writerow((
                    team[0]["id"],                
                    team[0]["displayName"],
                    team[0]["description"],
                    team[0]["specialization"],
                    team[0]["isMembershipLimitedToOwners"],
                    team[0]["visibility"],
                    team[0]["isArchived"],
                    team[0]["createdDateTime"]
                ))
    #return teams_not_activated

async def get_groups(endpoint, token, session):
    async with session.get(
        endpoint,
        headers={
            'Authorization': 'Bearer ' + token['access_token']
    }) as response:
        if response.status == 200:
            return json.loads(await response.read())
        else:
            logging.error(f"Some errors getting groups {response.status} {response.reason}")
            return None


async def get_data(token, session):

    page_counter = 1
    teams_not_activated = []

    logging.info(f"Getting page {page_counter}")
    page_counter += 1

    groups_data = await get_groups("https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')", token, session)
    await process_groups_page(groups_data, token, session)


    while "@odata.nextLink" in groups_data:

        if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
            token = await get_token()
            if not"access_token" in token:
                logging.error(f"{token.get('error')}: {token.get('error_description')} (correlation_ID: {token.get('correlation_id')})")                


        logging.info(f"Getting page {page_counter}")
        page_counter += 1

        groups_data = await get_groups(groups_data["@odata.nextLink"], token, session)
        await process_groups_page(groups_data, token, session)        

    return
    #print(len(teams_not_activated))   


async def start_report_gathering(token):

    header_columns = [
        "id",                
        "displayName",
        "description",
        "specialization",
        "isMembershipLimitedToOwners",
        "visibility",
        "isArchived",
        "createdDateTime"
    ]

    logging.info("Creating header row in CSV file '{0}'.".format(csv_file_path))

    with open(csv_file_path, "w", newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(header_columns)

    async with aiohttp.ClientSession() as session:
        await get_data(token, session)
        tasks = []
        task = asyncio.ensure_future(get_data(token, session))
        tasks.append(task)

        await asyncio.gather(*tasks, return_exceptions=True)


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

    import time

    try:
        config = json.load(open(sys.argv[1]))

        # Current script path
        current_wdpath = os.path.dirname(__file__)
        csv_file_path = os.path.join(current_wdpath, "result.csv")

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
