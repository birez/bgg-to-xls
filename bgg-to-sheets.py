import time
import xml.etree.ElementTree as ET
import requests
import xlsxwriter
import json
import validators
import html
from enum import Enum, IntEnum
from io import BytesIO

baseUrl="https://boardgamegeek.com/xmlapi2/"
gameUrl="https://boardgamegeek.com/boardgame/"
backoff_delay = 10

class Endpoints(Enum):
    COLLECTION = baseUrl + "collection?own=1&want=0&wishlist=0&username="
    THING = baseUrl + "thing?id="

    def __add__(self, other):
        return self.value + other

class Columns(IntEnum):
    IMAGE = 0
    NAME = 1
    GRADE = 2
    PLAYERS = 3
    DESCRIPTION = 4

# So not to kill API while testing
def delay():
    time.sleep(0.7) # There is unknown rate limit enforced, trying to limit request rate


def getXML(url, id):
    global backoff_delay
    delay()
    r = requests.get(url + id)
    print("GET: {}, code: {}".format(r.url, r.status_code))
    if r.status_code == 429:
        print("Rate limit exceeded.... waiting...")
        time.sleep(backoff_delay)
        backoff_delay = backoff_delay * 1.2
        print("Retrying...")
        return getXML(url, id)
    return r.text

def collectioToList(collectionXML):
    gamesList = []
    tree = ET.fromstring(collectionsXML)
    for item in tree:
        game = {}
        print(item.attrib)
        game["id"] = item.attrib["objectid"]
        game["name"] = item.find('name').text
        try:
            game["imgUrl"] = item.find('image').text
            game["thumb"] = item.find('thumbnail').text
        except:
            print("Game: {} Seems to not have images".format(item.find('name').text))
            game["imgUrl"] = ""
            game["thumb"] = ""

        gamesList.append(game)
    return gamesList

def addGamesDetails(gamelist):
    detailedLib = []
    for game in gamelist:
        thingXML = getXML(Endpoints.THING, game['id'])
        tree = ET.fromstring(thingXML)
        game["description"] = html.unescape(tree[0].find('description').text)
        game["minplayers"] = tree[0].find('minplayers').attrib['value']
        game["maxplayers"] = tree[0].find('maxplayers').attrib['value']
        detailedLib.append(game)
        # break
    return detailedLib


def makeLibrary(collectionsXML):
    gameslist = collectioToList(collectionsXML)
    print(json.dumps(gameslist, indent=1))
    detailedLib = addGamesDetails(gameslist)
    print(json.dumps(detailedLib, indent=1))
    return detailedLib

def getImage(url):
    if validators.url(url):
        r = requests.get(url)
        print("GET: {}, code: {}".format(r.url, r.status_code))
        if r.status_code == 200:
            return BytesIO(r.content)
    return None

def createSheet(lib, filename):
    headers = ["Image", "Name", "Grade", "Players", "Description"]
    workbook = xlsxwriter.Workbook(filename + ".xlsx")
    worksheet = workbook.add_worksheet()
    bold_format = workbook.add_format({'bold': True})
    wrap_format = workbook.add_format({'text_wrap': True})
    shrink_format = workbook.add_format({'shrink': True})

    worksheet.set_column(Columns.IMAGE, Columns.IMAGE, width=30)
    worksheet.set_column(Columns.NAME, Columns.NAME, width=20)
    worksheet.set_column(Columns.PLAYERS, Columns.PLAYERS, width=8)
    worksheet.set_column(Columns.GRADE, Columns.GRADE, width=8)
    worksheet.set_column(Columns.DESCRIPTION, Columns.DESCRIPTION, width=130)

    worksheet.set_row(0, height=None, cell_format=bold_format)
    worksheet.write_row(0,0,headers)
    row = 1
    for game in lib:
        image = getImage(game['thumb'])
        worksheet.insert_image(row, Columns.IMAGE, "image", {"image_data": image, 'object_position': 1})
        worksheet.write_url(row, Columns.NAME, gameUrl + game['id'], string=game['name'], cell_format=wrap_format)
        worksheet.write_string(row, Columns.PLAYERS, "{} - {}".format(
            game['minplayers'], game['maxplayers']
        ))
        worksheet.write_string(row, Columns.DESCRIPTION, game['description'], cell_format=wrap_format)
        worksheet.set_row(row, 120)
        row += 1

    workbook.close()


if __name__ == "__main__":
    username = "birez"
    collectionsXML = getXML(Endpoints.COLLECTION, username)
    library = makeLibrary(collectionsXML)
    createSheet(library, username)

