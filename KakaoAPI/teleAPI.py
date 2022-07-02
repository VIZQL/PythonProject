from telethon import TelegramClient, events, sync
import asyncio


# Remember to use your own values from my.telegram.org!
api_id = '14078381'
api_hash = '1c92f649600a09317f38c1df4f9e1ee4'
client = TelegramClient('KKW', api_id, api_hash)

@client.on(events.NewMessage(chats= "https://t.me/kikawoTest")) # 내 채널
@client.on(events.NewMessage(chats= "https://t.me/mainnews")) # main news 채널
@client.on(events.NewMessage(chats= "https://t.me/sedaily_worldwide_news")) # 국제뉴스 채널 서울경제
@client.on(events.NewMessage(chats= "https://t.me/rassiro_channel")) # AI 찌라시로
@client.on(events.NewMessage(chats= "https://t.me/davidstocknew")) # 주식소리통 NEO

async def my_event_handler(event):
    print(event.raw_text)

async def TeleAPI():

    await client.start()
    await client.run_until_disconnected()


if __name__ == "__main__":
    client.loop.run_until_complete(TeleAPI())




'''
from telethon import TelegramClient
import asyncio
import nest_asyncio
 
api_id = '14078381'
api_hash = '1c92f649600a09317f38c1df4f9e1ee4'
client = TelegramClient('session_file', api_id, api_hash)

async def Run_Telethon():
    await client.start()
    print('okay')
    await client.disconnect()
    
nest_asyncio.apply()
asyncio.run(Run_Telethon())

'''


##########################################################################
'''
# Remember to use your own values from my.telegram.org!
api_id = "14078381"
api_hash = "1c92f649600a09317f38c1df4f9e1ee4"
client = TelegramClient('anon', api_id, api_hash)

channel_name = "AI라씨로-주식투자속보"

@client.on(events.NewMessage(chats=channel_name))
async def my_event_handler(event):
    print(event.raw_text)

client.start()
client.run_until_disconnected()


token = "5392144349:AAEX1mKIHRh3gpCuEwMTfheYuogIoETWFrM"
bot = telegram.Bot(token)
updates = bot.getUpdates()

# print(updates.keys())
# print(type(updates))
for i in updates:
    print(i)
    # print(type(i))
# print(updates[0].message.chat_id)

# bot.sendMessage(chat_id='986047981', text="안녕하세요. 저는 봇입니다.")
'''