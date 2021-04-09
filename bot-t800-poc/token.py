import keyring
from jose import jwt

token = jwt.encode({"AppName": "T-800", "iss": "esserafael@gmail.com"}, keyring.get_password("T800_JWT", "TokenSecret"), algorithm="HS256")
print(token)
#keyring.set_password("T800_JWT", "Token", token)
token2 = jwt.decode(token, keyring.get_password("T800_JWT", "TokenSecret"))
print(token2)