import uuid
import random
import sqlite3 as sl
from gtts import gTTS
from pydub import AudioSegment

from config import DefaultConfig

CONFIG = DefaultConfig()

class Call:
    def __init__(self, id, state):
        self.id = id
        self.state = state


    async def answer(self, token, session):
        print(f"Answering call...")
                                
        endpoint = f"https://graph.microsoft.com/beta/app/calls/{self.id}/answer"
        json_data = {
            "callbackUri": f"{CONFIG.APP_NGROK_ADDR}/calling",
            "acceptedModalities": [ "audio" ],
            "mediaConfig": {
                "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
            }
        }

        async with session.post(
            endpoint,
            headers={
                'Authorization': 'Bearer ' + token,
                'Content-type': 'application/json'
            }, json=json_data
        ) as response:
            if response.status == 202:
                print("Call answered.")
            else:
                print("Problem answering call.")
    

    async def hang_up(self, token, session):
        print("Hanging up...")

        endpoint = f"https://graph.microsoft.com/beta/app/calls/{self.id}"
        
        async with session.delete(
            endpoint,
            headers={
                'Authorization': 'Bearer ' + token
            }
        ) as response:
            if response.status == 204:
                print("Call hung up")
            else:
                print(f"Problem hanging up call: {await response.read()}")

    
    async def get_participants(self, token, session):
        endpoint = f"https://graph.microsoft.com/beta/app/calls/{self.id}/participants"
        async with session.get(
            endpoint,
            headers={
                'Authorization': 'Bearer ' + token
            }
        ) as response:
            if response.status == 200:
                print("Participants list gathered.")
                return await response.read()
            else:
                print(f"Problem getting participants. {await response.read()}")

    
    async def play_prompt(self, text, token, session, language="en"):
        print(f"Sending prompt to call...")

        if not language:
            language = "en"

        await self._text_to_speech(text, language)
            
        endpoint = f"https://graph.microsoft.com/beta/app/calls/{self.id}/playPrompt"
        json_data = {
            #"clientContext": f"{str(uuid.uuid4())}",
            "clientContext": "t800-context",
            "prompts": [
                {
                    "@odata.type": "#microsoft.graph.mediaPrompt",
                    "mediaInfo": {
                        "@odata.type": "#microsoft.graph.mediaInfo",
                        "uri": f"{CONFIG.APP_NGROK_ADDR}/static/speech_16.wav",
                        #"uri": f"{CONFIG.APP_NGROK_ADDR}/static/6.wav",
                        "resourceId": f"{str(uuid.uuid4())}"
                    }
                }
            ]
        }

        async with session.post(
            endpoint,
            headers={
                'Authorization': 'Bearer ' + token,
                'Content-type': 'application/json'
            }, json=json_data
        ) as response:
            if response.status == 200:
                print("Prompt sent.")
            else:
                print("Problem sending prompt.")


    async def _text_to_speech(self, text, language):
        #language = 'en'

        myobj = gTTS(text=text, lang=language, slow=False)
        outputfile = "static\speech"
        myobj.save(f"{outputfile}.mp3")

        sound = AudioSegment.from_mp3(f"{outputfile}.mp3")
        converted_sound = sound.set_frame_rate(16000).set_channels(1)
        converted_sound.export(f"{outputfile}_16.wav", format="wav", bitrate=16)


    async def write_call_to_db(self, dbfilename):
        con = sl.connect(dbfilename)
        with con:
            data = [(self.id)]

            result = con.execute("SELECT * FROM Calls WHERE id_call = ?", data).fetchall()
            if result:
                sql = 'UPDATE Calls SET state = ? WHERE id_call = ?'
                data = [  
                    self.state,
                    self.id
                ]            
            else:
                sql = 'INSERT INTO Calls (id_call, state) values(?, ?)'
                data = [                                    
                    self.id,
                    self.state,
                ]                                    

            con.execute(sql, data)

        con.close()
