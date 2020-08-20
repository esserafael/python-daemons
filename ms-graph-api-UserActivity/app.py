"""
The configuration file would look like this (sans those // comments):

{
    "authority": "https://login.microsoftonline.com/Enter_the_Tenant_Name_Here",
    "client_id": "your_client_id",
    "scope": ["https://graph.microsoft.com/.default"],
        // For more information about scopes for an app, refer:
        // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow#second-case-access-token-request-with-a-certificate"

    "secret": "The secret generated by AAD during your confidential app registration",
        // For information about generating client secret, refer:
        // https://github.com/AzureAD/microsoft-authentication-library-for-python/wiki/Client-Credentials#registering-client-secrets-using-the-application-registration-portal

    "endpoint": "https://graph.microsoft.com/v1.0/users"

}

You can then run this sample with a JSON configuration file:

    python sample.py parameters.json
"""

import os
import sys  # For simplicity, we'll read config file from 1st CLI param sys.argv[1]
import json
import csv
import logging

import requests
import msal

import pandas as pd

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler()
    ]
)

client_id = os.getenv("daemon_client_id")
if not client_id:
    errmsg = "Define daemon_client_id environment variable"
    logging.error(errmsg)
    raise ValueError(errmsg)    
else:
    logging.info("client_id found -> '{0}'.".format(client_id))

client_secret = os.getenv("daemon_client_secret")
if not client_secret:
    errmsg = "Define daemon_client_secret environment variable"
    logging.error(errmsg)
    raise ValueError(errmsg)
else:
    logging.info("client_secret found.".format(client_id))   

config = json.load(open(sys.argv[1]))

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

# Firstly, looks up a token from cache
# Since we are looking for token for the current app, NOT for an end user,
# notice we give account parameter as None.
result = app.acquire_token_silent(config["scope"], account=None)

if not result:
    logging.info("No token exists in cache. Getting a new one from AzureAD.")
    result = app.acquire_token_for_client(scopes=config["scope"])

if "access_token" in result:
    # Calling graph using the access token
    graph_data = requests.get(  # Use token to call downstream service
        config["endpoint_test2"],
        headers={'Authorization': 'Bearer ' + result['access_token']}, ).json()
    if "error" in graph_data:
        logging.error("{0}: {1}".format(graph_data["error"]["code"], graph_data["error"]["message"]))
    else:
        print("Graph API call result: ")
        print(json.dumps(graph_data, indent=2))
        with open('graph_data.json', 'w', encoding='utf-8') as f_json:
            json.dump(graph_data, f_json, ensure_ascii=False, indent=4)
        #df = pd.read_json(r"graph_data.json")
        #df.to_csv("test.csv", encoding='utf-8', index=False)

        #json_data = json.loads(graph_data)

        f_csv = csv.writer(open("graph_data.csv", "w", encoding='utf-8'))
        f_csv.writerow(["Nome", "E-mailUniasselvi", "E-mailPessoal", "Celular", "ÚltimoLogon"])

        for graph_data in graph_data:
            f_csv.writerow(
                graph_data["displayName"],
                graph_data["mail"],
                graph_data["otherMails"][0],
                graph_data["mobilePhone"],
                graph_data["value"][0]["signInActivity"]["lastSignInDateTime"]
            )
        
else:
    logging.error("{0}: {1} (correlation_id: {3})".format(result.get("error"), result.get("error_description"), result.get("correlation_id")))
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))  # You may need this when reporting a bug

