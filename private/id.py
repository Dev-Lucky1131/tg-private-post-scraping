
from telethon.sync import TelegramClient

api_id =   
api_hash = ""  
channel_link = ""  
client = TelegramClient("session_name", api_id, api_hash)
client.start()

try:
    chat = client.get_entity(channel_link)  
    print(f"Channel Name: {chat.title}")
    print(f"Channel ID: {chat.id}")
except Exception as e:
    print(f"Error: {e}")

client.disconnect()
