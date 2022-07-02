from telethon.sync import TelegramClient
# from telethon import TelegramClient
from telethon.tl.functions.messages import (GetHistoryRequest)

import telethon
import asyncio


api_id = '14078381'
api_hash = '1c92f649600a09317f38c1df4f9e1ee4'
chat = "https://t.me/mainnews"

offset_id = 0
'''
"https://t.me/kikawoTest")) # 내 채널
"https://t.me/mainnews")) # main n
"https://t.me/sedaily_worldwide_ne
"https://t.me/rassiro_channel")) #
"https://t.me/davidstocknew")) # 주식
'''

# entity = await client.get_entity(chat)

client = TelegramClient('KKW', api_id, api_hash)

client.start()

history =  client(GetHistoryRequest(
            peer="https://t.me/kikawoTest",
            offset_id=offset_id,
            offset_date=None,
            add_offset=0,
            limit=5,
            max_id=0,
            min_id=0,
            hash=0
))


messages = history.messages[0].message

# offset_id = messages[len(messages) - 1].id
# total_messages = len(all_messages)

# next_post = client.iter_messages(
#                 chat,
#                 limit=5,
#                 min_id=your_post_id,
#                 reverse=True
#             )

print(messages)


# with TelegramClient('KKW', api_id, api_hash) as client:
    # print(client.iter_messages(chat))

    # for message in client.iter_messages(chat):
    #     # print()
    #     print(message.sender_id, ':', message.text)