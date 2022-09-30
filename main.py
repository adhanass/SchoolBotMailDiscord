import win32com.client
import pyttsx3
import os
from http import client
import discord
from dotenv import load_dotenv
from datetime import datetime, timedelta
import time

load_dotenv(dotenv_path="config")
default_intents = discord.Intents.default()
default_intents.members= True
client = discord.Client(intents = default_intents)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()

@client.event
async def on_ready():
    print("Cher Ecamien, le bot est de nouveau en ligne sur notre serveur !")
    
@client.event
async def on_message(cmddiscord):
    if cmddiscord.content.lower() == "!mail":
        while True:
            for message in inbox.Items:
                if message.UnRead == True:
                    print(message)
                    await cmddiscord.channel.send(f"@everyone Nouveau mail de la part de : {message.SenderName} !")
                    await cmddiscord.channel.send(f"Sujet du mail | '{message.subject}'")
                    await cmddiscord.channel.send(f"'{message.body}'")
                    message.Unread = False
            time.sleep(10)

        
client.run(os.getenv("TOKEN"))