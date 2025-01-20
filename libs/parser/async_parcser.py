import time
import asyncio

from aiohttp import ClientSession

from links import links


async def get_page(url):
    async with ClientSession() as session:
        async with session.get(url=url) as response:
            return response.status


async def main(links):
    tasks = []
    for url in links:
        tasks.append(asyncio.create_task(get_page(url)))

    for task in tasks:
        await task


a = time.strftime("%X")

asyncio.run(main(links))

b = time.strftime("%X")
print("\nTime", a, "-", b, "\n")
