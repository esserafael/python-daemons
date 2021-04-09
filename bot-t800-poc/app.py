# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio
import sys
import traceback
import uuid
import json
import re
import keyring
import sqlite3 as sl
import aiohttp
from datetime import datetime
from http import HTTPStatus
from urllib.parse import unquote

from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import (
    BotFrameworkAdapterSettings,
    TurnContext,
    BotFrameworkAdapter,
)
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.schema import Activity, ActivityTypes

from skills import TeamsConversation, Call
from config import DefaultConfig

from jose import jwt
from urllib.request import urlopen

CONFIG = DefaultConfig()


# Create adapter.
# See https://aka.ms/about-bot-adapter to learn more about how bots work.
SETTINGS = BotFrameworkAdapterSettings(CONFIG.APP_ID, CONFIG.APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)


# Catch-all for errors.
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The bot encountered an error or bug.")
    await context.send_activity(
        "To continue to run this bot, please fix the bot source code."
    )
    # Send a trace activity if we're talking to the Bot Framework Emulator
    if context.activity.channel_id == "emulator":
        # Create a trace activity that contains the error object
        trace_activity = Activity(
            label="TurnError",
            name="on_turn_error Trace",
            timestamp=datetime.utcnow(),
            type=ActivityTypes.trace,
            value=f"{error}",
            value_type="https://www.botframework.com/schemas/error",
        )
        # Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.send_activity(trace_activity)


ADAPTER.on_turn_error = on_error

# If the channel is the Emulator, and authentication is not in use, the AppId will be null.
# We generate a random AppId for this case only. This is not required for production, since
# the AppId will have a value.
APP_ID = SETTINGS.app_id if SETTINGS.app_id else uuid.uuid4()

# Create the Bot
BOT = TeamsConversation(CONFIG.APP_ID, CONFIG.APP_PASSWORD)

# Listen for incoming requests on /api/messages.
async def messages(req: Request) -> Response:
    # Main bot message handler.
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)

    activity = Activity().deserialize(body)
    auth_header = req.headers["Authorization"] if "Authorization" in req.headers else ""

    response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if response:
        return json_response(data=response.body, status=response.status)
    return Response(status=HTTPStatus.OK)


async def validate_incoming_token(req):
    auth = req.headers.get("Authorization", None)
    if not auth:
        return False
    
    auth_parts = auth.split() 
    if auth_parts[0].lower() != "bearer" or len(auth_parts) == 1 or len(auth_parts) > 2:
        return False
    
    token = auth_parts[1]
    agent = req.headers.get("User-Agent", None)
    if agent and "Microsoft-Skype" in agent:
        oidconfig = json.loads(urlopen("https://api.aps.skype.com/v1/.well-known/OpenIdConfiguration").read())
        jwks = json.loads(urlopen(oidconfig.get("jwks_uri", {})).read())
        unverified_header = jwt.get_unverified_header(token)
        rsa_key = {}
        for key in jwks["keys"]:
            if key["kid"] == unverified_header["kid"]:
                rsa_key = {
                    "kty": key["kty"],
                    "kid": key["kid"],
                    "use": key["use"],
                    "n": key["n"],
                    "e": key["e"]
                }
        if rsa_key:
            try:
                payload = jwt.decode(
                    token,
                    rsa_key,
                    algorithms=unverified_header.get("alg", {}),
                    audience=CONFIG.APP_ID,
                    issuer=oidconfig.get("issuer", {}),
                    options={
                        'verify_signature': True,
                        'verify_aud': True,
                        'verify_iat': True,
                        'verify_exp': True,
                        'verify_nbf': True,
                        'verify_iss': True
                    }
                )

                if payload.get("aud", {}) == CONFIG.APP_ID and payload.get("iss", {}) == oidconfig.get("issuer", {}):
                    return True

            except jwt.ExpiredSignatureError:
                return False
            except jwt.JWTClaimsError:
                return False
            except jwt.JWTError:
                return False
            except Exception:
                return False
        else:
            False
    
    else:
        try:
            payload = jwt.decode(token, keyring.get_password("T800_JWT", "TokenSecret"))
            return True
        except jwt.JWTError:
            return False
        except Exception:
            return False


async def get_token():
    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"{CONFIG.APP_AUTHORITY}/b0e7335f-fd1f-46ad-98c7-55e6e4e222ea/oauth2/v2.0/token",
            headers={
                'Content-type': 'application/x-www-form-urlencoded'
            }, data={
                "client_id": CONFIG.APP_ID,
                "scope": CONFIG.APP_SCOPE,
                "grant_type": "client_credentials",
                "client_secret": CONFIG.APP_PASSWORD
            }
        ) as response:
            if response.status == 200:
                return json.loads(await response.read()).get("access_token", {})


async def calling(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"] or "application/json" == req.content_type:

        if not await validate_incoming_token(req):
            print(f"Token inválido: {req.headers} - {req}")
            return web.Response(body=b'{"error": "Invalid authorization"}', status=HTTPStatus.UNAUTHORIZED)
        
        body = await req.json()

        async with aiohttp.ClientSession() as session:

            token = await get_token()

            if token:
                for notification in body.get('value', {}):

                    call = Call(
                        re.search("^\/.*?\/.*?\/(.*?)(/.*)?$", notification['resource']).group(1),
                        None
                    )

                    if 'state' in notification['resourceData']:
                        print(f"changeType: {notification['changeType']} - state: {notification['resourceData']['state']}")

                        call.state = notification.get('resourceData', {}).get('state', {})

                        await call.write_call_to_db("t-800.db")

                        if notification.get('resourceData', {}).get('state', {}) == "incoming":                                    
                            await call.answer(token, session)
                        elif notification.get('resourceData', {}).get('state', {}) == "established" and CONFIG.PLAY_PROMPT:
                            if notification.get('resourceData', {}).get('meetingInfo', {}) or notification.get('resourceData', {}).get('mediaState', {}).get('audio', {}) == "active":
                                await call.play_prompt("I can also change my voice.", token, session)

                    elif 'status' in notification['resourceData']:
                        print(f"changeType: {notification['changeType']} - state: {notification['resourceData']['status']}")
                        if CONFIG.PLAY_PROMPT_HANGUP_AFTER and (notification.get('resourceData', {}).get('status', {}) == "completed" or notification.get('resourceData', {}).get('status', {}) == "failed"):
                            await call.hang_up(token, session)   
                            if notification.get('resourceData', {}).get('status', {}) == "failed":
                                print(f"Failure info: {notification.get('resourceData', {}).get('resultInfo', {}).get('message', {})}")

                    return Response(status=HTTPStatus.OK)         
                        
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)


async def _join_call(url, token, session):
    join_weburl_search = re.search('.*meetup-join\/((?:(?!\/).)*)\/((?:(?!\?).)*).*context=(.*)', unquote(url))
    context = json.loads(join_weburl_search.group(3))

    json_data = {
        "@odata.type": "#microsoft.graph.call",
        "callbackUri": f"{CONFIG.APP_NGROK_ADDR}/calling",
        "requestedModalities": [
            "audio"
        ],
        "mediaConfig": {
            "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
        },
        "chatInfo": {
            "@odata.type": "#microsoft.graph.chatInfo",
            "threadId": join_weburl_search.group(1),
            "messageId": join_weburl_search.group(2)
        },
        "meetingInfo": {
            "@odata.type": "#microsoft.graph.organizerMeetingInfo",
            "organizer": {
                "@odata.type": "#microsoft.graph.identitySet",
                "user": {
                    "@odata.type": "#microsoft.graph.identity",
                    "id": context.get("Oid", {}),
                    "tenantId": context.get("Tid", {})
                }
            }
        },
        "tenantId": context.get("Tid", {})
    }

    async with session.post(
        "https://graph.microsoft.com/beta/app/calls",
        headers={
            'Authorization': 'Bearer ' + token,
            'Content-type': 'application/json'
        }, json=json_data
    ) as response:
        if response.status == 201:
            print("Call joined.")
            call = json.loads(await response.read())
            con = sl.connect(CONFIG.DBFILE)
            with con:
                sql = 'INSERT INTO Calls (id_call, state, join_weburl) values(?, ?, ?)'
                data = [                                    
                    call.get('id', {}),
                    call.get('state', {}),
                    url
                ]
                con.execute(sql, data)
            con.close()

            return call
        else:
            print("Problem joining call.")


async def join(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"] or "application/json" == req.content_type:

        if not await validate_incoming_token(req):
            print(f"Token inválido: {req.headers} - {req}")
            return web.Response(body=b'{"error": "Invalid authorization"}', status=HTTPStatus.UNAUTHORIZED)

        body = await req.json()
        async with aiohttp.ClientSession() as session:
            token = await get_token()
            if token:
                await _join_call(body.get("webJoinUrl", {}), token, session)

        return Response(status=HTTPStatus.OK)
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)


async def hangup(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"] or "application/json" == req.content_type:
        
        if not await validate_incoming_token(req):
            print(f"Token inválido: {req.headers} - {req}")
            return web.Response(body=b'{"error": "Invalid authorization"}', status=HTTPStatus.UNAUTHORIZED)
        
        async with aiohttp.ClientSession() as session:
            token = await get_token()
            if token:                
                con = sl.connect(CONFIG.DBFILE)
                with con:
                    results = con.execute("SELECT * FROM Calls WHERE state = 'established'").fetchall()
                    hang_calls = []
                    if results:                        
                        for result in results:                   
                            call = Call(
                                result[1],
                                result[3]
                            )

                            await call.hang_up(token, session)
                            hang_calls.append(call.id)
                con.close()
                json_data = {"calls": hang_calls}
                return json_response(data=json_data, status=HTTPStatus.OK)
            else:
                return Response(body=b'{"error": "Could not get token"}', status=HTTPStatus.INTERNAL_SERVER_ERROR)
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)


async def participants(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"] or "application/json" == req.content_type:
        
        if not await validate_incoming_token(req):
            print(f"Token inválido: {req.headers} - {req}")
            return web.Response(body=b'{"error": "Invalid authorization"}', status=HTTPStatus.UNAUTHORIZED)
            
        body = await req.json()
        if not "webJoinUrl" in body:
            return Response(body=b'{"error": "Missing webJoinUrl property in body"}', status=HTTPStatus.BAD_REQUEST)
        else:
            url = body.get("webJoinUrl", None)
            async with aiohttp.ClientSession() as session:
                token = await get_token()
                if token:
                    con = sl.connect(CONFIG.DBFILE)
                    with con:
                        data = [(url)]
                        result = con.execute("SELECT * FROM Calls WHERE join_weburl = ? AND state = 'established'", data).fetchone()
                        if result:
                            print("Call já estabelecida.")                     
                            call = Call(result[1], None)
                            result = await call.get_participants(token, session)
                        else:
                            print("Ingressando na call.")   
                            join_result = await _join_call(url, token, session)
                            if join_result:
                                call = Call(join_result.get('id', None), None)                                      
                                data = [(call.id)]
                                result = None
                                while not result:
                                    with con:
                                        result = con.execute("SELECT * FROM Calls WHERE id_call = ? AND state = 'established'", data).fetchall()
                                        await asyncio.sleep(0.01)                           
                                result = await call.get_participants(token, session)

                        #await call.hang_up(token, session)
                    con.close()
                    return web.Response(body=result)
                else:
                    return Response(body=b'{"error": "Could not get token"}', status=HTTPStatus.INTERNAL_SERVER_ERROR)
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)


async def broadcast(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"] or "application/json" == req.content_type:
        
        if not await validate_incoming_token(req):
            print(f"Token inválido: {req.headers} - {req}")
            return web.Response(body=b'{"error": "Invalid authorization"}', status=HTTPStatus.UNAUTHORIZED)
            
        body = await req.json()
        if not "speech" in body:
            return Response(body=b'{"error": "Missing speech property in body"}', status=HTTPStatus.BAD_REQUEST)
        else:
            async with aiohttp.ClientSession() as session:
                token = await get_token()
                if token:
                    con = sl.connect(CONFIG.DBFILE)
                    with con:
                        results = con.execute("SELECT * FROM Calls WHERE state = 'established'").fetchall()
                        sent_calls = []
                        if results:                    
                            for result in results:
                                call = Call(result[1], None)
                                await call.play_prompt(body.get("speech", {}), token, session, body.get("lang", {}))
                                sent_calls.append(call.id)
                    con.close()
                    json_data = {"calls": sent_calls}
                    return json_response(data=json_data, status=HTTPStatus.OK)
                
                else:
                    return Response(status=HTTPStatus.INTERNAL_SERVER_ERROR)



APP = web.Application(middlewares=[aiohttp_error_middleware])
APP.router.add_post("/api/messages", messages)
APP.router.add_post("/calling", calling)
APP.router.add_post("/join", join)
APP.router.add_post("/hangup", hangup)
APP.router.add_post("/participants", participants)
APP.router.add_post("/broadcast", broadcast)

APP.router.add_static('/static', 'static')

if __name__ == "__main__":
    try:
        web.run_app(APP, host="localhost", port=CONFIG.PORT)
    except Exception as error:
        raise error
