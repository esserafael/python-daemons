#!/usr/bin/env python3
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import os
import keyring

""" Bot Configuration """


class DefaultConfig:
    """ Bot Configuration """

    #PORT = 3978
    PORT = 12345
    DBFILE = "t-800.db"
    APP_AUTHORITY = "https://login.microsoftonline.com"
    APP_SCOPE = [ "https://graph.microsoft.com/.default" ]
    APP_ID = os.environ.get("MicrosoftAppId", "95dc3706-fe69-4ee2-9879-750660f64753")
    #APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")
    APP_PASSWORD = keyring.get_password("T800", "AppSecret")
    APP_NGROK_ADDR = "https://34aae6ba15b8.ngrok.io"
    PLAY_PROMPT = False
    PLAY_PROMPT_HANGUP_AFTER = False
