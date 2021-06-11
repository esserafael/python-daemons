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
import datetime
import json
import logging
import re

import aiohttp
import msal

import asyncio
import time

import keyring
import uuid


async def get_token():

    '''
    client_secret_env = "py_usermgt_secret"
    client_secret = os.getenv(client_secret_env)
    if not client_secret:
        errmsg = f"Define {client_secret_env} environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)
    else:
        logging.info(f"Secret {client_secret_env} found.")
    '''

    client_secret = keyring.get_password(config["keyring_app"], config["client_id"])

    # Create a preferably long-lived app instance which maintains a token cache.
    app = msal.ConfidentialClientApplication(
        config["client_id"], authority=config["authority"],
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

    result.update({
        "renew_datetime": datetime.datetime.now() + datetime.timedelta(seconds=result["expires_in"]),
        "gioconda_access_token": keyring.get_password(config["keyring_app"], "gioconda_auth")
        })
    
    return result


async def renew_token():
    logging.info("Getting new AzureAD token.")
    result = await get_token()
    while not "access_token" in result:
        logging.error(f"Error getting new token: {result.get('error')}: {result.get('error_description')} (correlation_ID: {result.get('correlation_id')})")                
        time.sleep(30)
        result = await get_token()

    return result

'''
async def get_gioconda_token():    
    client_secret_env = "py_usermgt_gioconda_secret"
    client_secret = os.getenv(client_secret_env)
    if not client_secret:
        errmsg = f"Define {client_secret_env} environment variable"
        logging.error(errmsg)
        raise ValueError(errmsg)
    else:
        logging.info(f"Gioconda Secret {client_secret_env} found.")

    return client_secret
'''

async def add_license(user, token, session, *args):
    try:
        log_message_begin = f"User: {user['emailAD']}"
        license_sku = args[0]

        license = {
            "addLicenses": [
                {
                "disabledPlans": [],
                "skuId": license_sku
                }
            ],
            "removeLicenses": []
        }

        async with session.post(
            f"{config['endpoint_users']}/{user['emailAD']}/assignLicense",
            headers={
                'Authorization': 'Bearer ' + token['access_token'],
                'Content-type': 'application/json'
            }, json=license
        ) as response_lic:

            log_message_end = f"({response_lic.status} {response_lic.reason})."
            if response_lic.status == 200:
                logging.info(f"{log_message_begin} - User licensed {log_message_end}")
                user.update({"result_status": "OK", "result_msg": "Usuário criado e licenciado com sucesso."})
            elif response_lic.status == 404:
                logging.error(f"{log_message_begin} - User has not been found {log_message_end}")
                await asyncio.sleep(10)
                user.update(await add_license(user, token, session, license_sku))
                #user.update({"result_msg": f"Usuário não encontrado (Problema durante licenciamento). {log_message_end}"})
            else:
                logging.error(f"{log_message_begin} - User has not been licensed {log_message_end}")
                user.update({"result_msg": f"Erro ao atribuir licença ao usuário. {log_message_end}"})

    except Exception as e:
        logging.error(f"{log_message_begin} - Error licensing user: {e}")
        user.update({"result_msg": f"Erro ao atribuir licença ao usuário. {e}"})

    return user


async def create_user(user, token, session, *args):

    log_message_begin = f"User: {user['emailAD']}"
    user.update({"result_status": "NOK", "result_msg": ""}) 

    try:
        if "@aluno.uniasselvi.com.br" in user['emailAD']:
            license_sku = config['sku_aluno']
            department = "Aluno"
        else:
            license_sku = config['sku_docente']
            department = "Docente"

        user_post = {
            "accountEnabled": True,
            "displayName": user['nome'],
            "givenName": user['primeiro_nome'],
            "surname": user['ultimo_nome'],
            "mailNickname": re.sub('@.*', '', user['emailAD']),
            "userPrincipalName": user['emailAD'],
            "mail": user['emailAD'],
            "department": department,
            "companyName": "Uniasselvi",
            "country": "BR",
            "usageLocation": "BR",
            "preferredLanguage": "pt-BR",
            "passwordProfile" : {
                "forceChangePasswordNextSignIn": False,
                "password": user['senhaAD']
            },
            "passwordPolicies": "DisablePasswordExpiration",
            "onPremisesExtensionAttributes": {
                "extensionAttribute3": user['cpf']
            }
        }

        if user['email_recuperacao']:
            user_post.update({"otherMails": [user['email_recuperacao']]})
        
        if user['telefone_recuperacao']:
            user_post.update({"mobilePhone": user['telefone_recuperacao']})

        async with session.post(
            config['endpoint_users'],
            headers={
                'Authorization': 'Bearer ' + token['access_token'],
                'Content-type': 'application/json'
            }, json=user_post
        ) as response:

            log_message_end = f"({response.status} {response.reason})."

            if response.status == 201:
                logging.info(f"{log_message_begin} - User created {log_message_end}")
                user.update(await add_license(user, token, session, license_sku))

            elif response.status == 400:
                error = json.loads(await response.read())
                log_message_end = f"{log_message_end} - {error['error']['message']}"

                if "userPrincipalName already exists" in error['error']['message']:                    
                    # Try to assign license
                    logging.info(f"{log_message_begin} - User already exists, will try to assign license. {log_message_end}")
                    user.update(await add_license(user, token, session, license_sku))                    
                else:
                    logging.warning(f"{log_message_begin} - User has not been created. Checking if it already exists. {log_message_end}")

                    async with session.get(
                        f"{config['endpoint_users']}/{user['emailAD']}",
                        headers={
                            'Authorization': 'Bearer ' + token['access_token']
                        }
                    ) as response_get_user:
                        if response_get_user.status == 200:
                            logging.info(f"{log_message_begin} - User already exists.")
                            user.update(await add_license(user, token, session, license_sku))
                        else:
                            logging.error(f"{log_message_begin} - User has not been created {log_message_end}")
                            user.update({"result_msg": f"Erro ao criar usuário. {log_message_end}"})                  

            else:
                logging.error(f"{log_message_begin} - User has not been created {log_message_end}")
                user.update({"result_msg": f"Erro ao criar usuário. {log_message_end}"})
        
    except Exception as e:
        logging.error(f"{log_message_begin} - Error creating user: {e}")
        user['result_msg'] = f"Erro ao criar usuário. {e}"

    try:
        status = {
            "emailAD": user['emailAD'],
            "status": user['result_status'],
            "message": user['result_msg']
        }
        
        async with session.post(
            args[0],
            headers={
               'Authorization': 'Basic ' + token['gioconda_access_token'],
                'Content-type': 'application/json'
            }, json=status
        ) as response_status:
            log_message_end = f"({response_status.status} {response_status.reason})."
            logging.info(f"{log_message_begin} - Status update response: {log_message_end}")

    except Exception as e:
        logging.error(f"{log_message_begin} - Error sending user status: {e}")


async def start_user_creation(token):

    if config["prod"]:
        endpoint_gioconda = config["endpoint_gioconda_prod"]
        endpoint_gioconda_status = config["endpoint_gioconda_status_prod"]
    else:
        endpoint_gioconda = config["endpoint_gioconda_homo"]
        endpoint_gioconda_status = config["endpoint_gioconda_status_homo"]
        created_json_file = os.path.join(current_wdpath, f"CreatedUsers_{uuid.uuid4()}.json")

    async with aiohttp.ClientSession() as session:

        while True:

            try:

                if token["renew_datetime"] <= (datetime.datetime.now() + datetime.timedelta(minutes=5)):
                    token = await renew_token() 

                async with session.get(
                    endpoint_gioconda,
                    headers={
                        'Authorization': 'Basic ' + token['gioconda_access_token'],
                    }
                ) as response:
                    if response.status == 200:
                        tasks = []
                        users = json.loads(await response.read())
                        for user in users:
                            #logging.info(json.dumps(user))
                            tasks.append(asyncio.ensure_future(create_user(user, token, session, endpoint_gioconda_status)))

                            if not config['prod']:
                                with open(created_json_file, 'a', encoding='utf-8') as json_file:
                                    json_file.write(json.dumps(user))
                        
                        await asyncio.gather(*tasks, return_exceptions=True)
                    
                    elif response.status >= 500:
                        logging.error(f"Getting error response while trying to request Gioconda API - ({response.status} {response.reason}).")
                        time.sleep(30)
            
            except Exception as e:
                logging.error(f"Error(2): {str(e)}")
                time.sleep(10)

        
        #task = asyncio.ensure_future(get_org_data(token, session))
        #task = asyncio.ensure_future(create_user(user, token, session))
        #tasks.append(task)

        #await asyncio.gather(*tasks, return_exceptions=True)


async def main():

    token = await get_token()

    if "access_token" in token:
        await start_user_creation(token)
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
        logging.info(f"Script started. Production: {config['prod']}")
        asyncio.get_event_loop().run_until_complete(main())
        elapsed = time.perf_counter() - s
        logging.info(f"Script finished, executed in {elapsed:0.2f} seconds.")
    except Exception as e:
        logging.error(f"Error(1): {str(e)}")
