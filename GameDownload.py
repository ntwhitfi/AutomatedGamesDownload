from bs4 import BeautifulSoup
from pyppeteer import launch
from pathlib import Path
from pyunpack import Archive
import json
import pandas as pd 
import shutil
import os
import time
import requests
import asyncio

# vimm's lair URL - source for roms/isos
URL = "https://vimm.net"
# URL params for advanced search
SEARCH_PATH = "/vault/?mode=adv&p=list&system={0}&q={1}&players=%3E%3D&playersValue=1&simultaneous=Any&publisher=&year=%3D&yearValue=&cart=%3D&cartValue=&rating=%3D&ratingValue=&sort=Title&sortOrder=ASC"
# Download URL
DOWNLOAD_URL = "https://download4.vimm.net/download/?mediaId={0}"
# Local download path
LOCAL_DOWNLOAD_PATH = 'C:\\Users\\User\\Downloads\\'
# Game system path map
GAME_PATH_MAP = {
    "PS1": "C:\\Users\\User\\Documents\\ROMS\\PS1\\",
    "Dreamcast": "C:\\Users\\User\\Documents\\ROMS\\Dreamcast\\",
    "GameCube": "C:\\Users\\User\\Documents\\ROMS\\Gamecube\\",
    "N64": "C:\\Users\\User\\Documents\\ROMS\\N64\\",
    "PS2": "C:\\Users\\User\\Documents\\ROMS\\PS2 ISOs\\",
    "PSP": "C:\\Users\\User\\Documents\\ROMS\\PSP\\",
}
# Game requests file
GAME_REQUESTS_FILE_PATH = 'C:\\Users\\User\\OneDrive\\GameRequests.xlsx'
# Completed game requests file
COMPLETED_REQUESTS_FILE = "C:\\Users\\User\\Documents\\automation\\completedGames.txt"


# Retrieve game page from search
def get_page_from_search(system, query):
    response = requests.get(URL + SEARCH_PATH.format(system, query))
    # retrieve table of search results
    html_parser = BeautifulSoup(response.text, features="html.parser")
    game_table = html_parser.find("table", attrs={"class":"rounded centered cellpadding1 hovertable"})
    game_list = []
    
    for row in game_table.find_all("a"):
        print(row['href'])
        game_page = str(row['href'])
        if "?p=" not in game_page:
            game_list.append(game_page)
    
    return game_list


# Retrieve download link for the game requested
def get_game_download_link(game_urls):
    media_id_list = []
    for game_link in game_urls:
        response = requests.get(URL + game_link)
        html_parser = BeautifulSoup(response.text, features="html.parser")
        page_inputs = html_parser.find_all("input")

        for input_tag in page_inputs:
            print(input_tag)
            if 'name' in str(input_tag) and input_tag['name'] == "mediaId":
                media_id_list.append(input_tag['value'])

    return media_id_list


#Use Puppeteer to download the files to avoid bot checks
async def download_games(games):
    downloaded_file_list = []
    browser = await launch({'headless': False})
    page = await browser.newPage()
    await page.goto(URL)

    file_name = ""
    for game_link in games:
        await page.goto(URL + game_link)
        form = await page.querySelector('#download_form')
        await page.evaluate('(form) => form.submit()', form)
        while file_name is "" or file_name.name.endswith(".crdownload"):
            time.sleep(5)
            file_list = sorted(Path(LOCAL_DOWNLOAD_PATH).iterdir(), key=os.path.getmtime, reverse=True)
            if len(file_list) > 0:
                file_name = file_list[0]
                print(file_name)
        downloaded_file_list.append(file_name.name)
    await browser.close()

    return downloaded_file_list


#Extract downloaded rom archive to the approprate folder
def extract_game(system, archive_file_path):
    Archive(archive_file_path).extractall(GAME_PATH_MAP[system])


#Get list of games to download from excel sheet
def get_game_requests():
    game_requests = []
    xls = pd.ExcelFile(GAME_REQUESTS_FILE_PATH)
    data_frame = pd.read_excel(xls, 'Requests')
    print(data_frame)
    for index, row in data_frame.iterrows():
        game_requests.append({"title": row['Title'], "system": row['System']})
    return game_requests


#Record completed request
def record_completed_request(request):
    with open(COMPLETED_REQUESTS_FILE, "a") as file:
        file.write(request + "\n")


#Get previously completed requests
def get_completed_requests():
    completed_requests = []
    with open(COMPLETED_REQUESTS_FILE, "r") as file:
        completed_requests = file.read().split('\n')
    return completed_requests


for request in get_game_requests():
    completed_requests = get_completed_requests()
    print(str(completed_requests))
    if not json.dumps(request) in get_completed_requests():
        games_to_download = get_page_from_search(request['system'], request['title'])
        print("games: " + str(games_to_download))

        game_archive_files = asyncio.get_event_loop().run_until_complete(download_games(games_to_download))

        for game_archive in game_archive_files:
            print("Extracting archive")
            extract_game(request['system'], LOCAL_DOWNLOAD_PATH + game_archive)
    
        print("Finished request " + json.dumps(request))

        record_completed_request(json.dumps(request))
    else:
        print("Request already completed " + json.dumps(request))
print("Done!")
