# Import the required module for text 
# to speech conversion
from gtts import gTTS
from pydub import AudioSegment
  
# This module is imported so that we can 
# play the converted audio
import os
  
# The text that you want to convert to audio
mytext = 'Olá meus amigos.'
  
# Language in which you want to convert
language = 'en'
  
# Passing the text and language to the engine, 
# here we have marked slow=False. Which tells 
# the module that the converted audio should 
# have a high speed
myobj = gTTS(text=mytext, lang=language, slow=False)
  
# Saving the converted audio in a mp3 file named
# welcome 
myobj.save("welcome.mp3")

sound = AudioSegment.from_mp3("welcome.mp3")
sound.export("welcome.wav", format="wav", bitrate=16)
  
# Playing the converted file
#os.system("mpg321 welcome.mp3")