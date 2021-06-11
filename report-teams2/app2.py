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

async def get_events(user, token, session):
    try:
        #logging.info(user)
        async with session.get(
            f"https://graph.microsoft.com/v1.0/users/{user}/calendarView/?endDateTime=2021-03-06T03:00:00Z&startDateTime=2021-03-01T01:00:00Z",
            headers={
                'Authorization': 'Bearer ' + token['access_token']
        }) as response_events:

            if response_events.status == 200:
                event_data = json.loads(await response_events.read())
                for event in event_data["value"]:           
                    #if not re.search("^Encontro com o.+", event["subject"]):
                    with open(csv_file_path, "a", newline='', encoding='utf-8') as csv_file:
                        csv_writer = csv.writer(csv_file)
                        if event:
                            csv_writer.writerow((
                                #user["id"],
                                #user["displayName"],
                                #user["userPrincipalName"],
                                user,
                                event["subject"],
                                f"{event['start']['dateTime']}Z",
                                f"{event['end']['dateTime']}Z",
                                event["organizer"]["emailAddress"]["name"],
                                event["organizer"]["emailAddress"]["address"],
                                event["createdDateTime"],
                                event["lastModifiedDateTime"],
                                event["type"],
                                event["isOnlineMeeting"],
                                event["id"],
                                event["iCalUId"]                            
                            ))
                        '''
                        async with session.delete(
                            f"https://graph.microsoft.com/v1.0/users/{user['id']}/events/{event['id']}",
                            headers={
                                'Authorization': 'Bearer ' + token['access_token']
                        }) as response_delete:  
                            print("Sei la")  
                        '''


            elif response_events.status == 404:
                logging.warning(f"Events not found {response_events.status} {response_events.reason}.")
                #return None
            else:
                logging.error(f"Some errors getting events {response_events.status} {response_events.reason}")
        
    except Exception as e:
        logging.error(f"Error in get_events function: {str(e)}")

async def process_users_page(users_data, token, session):
    #users_events = []
    if "value" in users_data:
        for user in users_data["value"]:
            tasks = []
            task = asyncio.ensure_future(get_events(user, token, session))
            tasks.append(task)        
        
        await asyncio.gather(*tasks, return_exceptions=True)
            #teams_not_activated2 = [team for team in teams_not_activated if team]
            #teams_not_activated.append(await get_team(group['id'], token, session))
    else:        
        #tasks = []
        #task = asyncio.ensure_future(get_events(users_data, token, session))
        #tasks.append(task)        
    
        #await asyncio.gather(*tasks, return_exceptions=True)

        await get_events(users_data, token, session)
        #teams_not_activated2 = [team for team in teams_not_activated if team]
        #teams_not_activated.append(await get_team(group['id'], token, session))
    
    #print(f"Found: {len(users_events)}")

    
    #return teams_not_activated

async def get_users(endpoint, token, session):
    async with session.get(
        endpoint,
        headers={
            'Authorization': 'Bearer ' + token['access_token']
    }) as response:
        if response.status == 200:
            #return json.loads(await response.read())
            await process_users_page(json.loads(await response.read()), token, session)
        else:
            logging.error(f"Some errors getting users {response.status} {response.reason}")
            return None


async def get_data(token, session):

    page_counter = 1
    teams_not_activated = []

    logging.info(f"Getting page {page_counter}")
    page_counter += 1

    csv_file_read = os.path.join(os.path.dirname(__file__), config["csv_read_name"])

    task_counter = 0
    task_counter_total = 0
    tasks = []

    with open(csv_file_read, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:

            if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
                token = await get_token()
                if not"access_token" in token:
                    logging.error(f"{token.get('error')}: {token.get('error_description')} (correlation_ID: {token.get('correlation_id')})")                

            #print(f"{row['Email']}")

            task = asyncio.ensure_future(get_events(row["EMAIL"], token, session))
            tasks.append(task)

            task_counter += 1

            if task_counter == 50:
            
                logging.info(f"Sending {task_counter} tasks(users).")

                await asyncio.gather(*tasks, return_exceptions=True)

                task_counter_total += task_counter
                logging.info(f"{task_counter_total} users processed.")

                task_counter = 1
                tasks = []

                    
            #users_data = await get_users(f"https://graph.microsoft.com/v1.0/users/{row['Email']}", token, session)
            #await process_users_page(users_data, token, session)  

    '''
    users_data = await get_users("https://graph.microsoft.com/v1.0/users?$top=10", token, session)
    #users_data = await get_users("https://graph.microsoft.com/v1.0/users/5b45e36b-4a84-4260-8c38-2605098189b7", token, session)
    #await process_users_page(users_data, token, session)


    while "@odata.nextLink" in users_data:

        if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
            token = await get_token()
            if not"access_token" in token:
                logging.error(f"{token.get('error')}: {token.get('error_description')} (correlation_ID: {token.get('correlation_id')})")                


        logging.info(f"Getting page {page_counter}")
        page_counter += 1

        users_data = await get_users(users_data["@odata.nextLink"], token, session)
        await process_users_page(users_data, token, session)        
    '''
    return
    #print(len(teams_not_activated))   


async def start_report_gathering(token):

    header_columns = [
        #"Userid",
        #"UserdisplayName",
        "userPrincipalName",
        "EventSubject",
        "start_dateTime",
        "end_dateTime",
        "organizer_name",
        "organizer_emailAddress",
        "createdDateTime",
        "lastModifiedDateTime",
        "type",
        "isOnlineMeeting",
        "Eventid",
        "iCalUId"
    ]

    logging.info("Creating header row in CSV file '{0}'.".format(csv_file_path))

    with open(csv_file_path, "w", newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(header_columns)

    async with aiohttp.ClientSession() as session:
        #await get_data(token, session)
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
        csv_file_path = os.path.join(current_wdpath, config["csv_result_name"])

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
