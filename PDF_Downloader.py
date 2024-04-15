import sys, urllib, openpyxl
import urllib.request
import urllib.error
from itertools import islice

BRNUM = 0
FIRST_DOWNLOAD_LINK = 37
SECOND_DOWNLOAD_LINK = 38
SHEET_NAME = "0"


wbSave = None


def generator(filename: str, amount: int):
    wb = openpyxl.load_workbook(filename=filename, read_only=True)
    ws = wb[SHEET_NAME]
    i = 0
    for row in islice(ws.iter_rows(values_only=True), 1, None):
        if amount > 0:
            i += 1
            if i >= amount + 1:
                break
        downloader([row[BRNUM], row[FIRST_DOWNLOAD_LINK], row[SECOND_DOWNLOAD_LINK]])


def downloader(row):
    url = row[1]
    downloaded = False
    savefile = f"{row[0]}.pdf"
    try:
        urllib.request.urlretrieve(url, savefile)
        downloaded = True
    except (
        urllib.error.HTTPError,
        urllib.error.URLError,
        ConnectionResetError,
        Exception,
    ) as e:
        url = row[2]
        if url != "":
            try:
                url = row[2]
                urllib.request.urlretrieve(url, savefile)
                downloaded = True
            except (
                urllib.error.HTTPError,
                urllib.error.URLError,
                ConnectionResetError,
                Exception,
            ) as e2:
                print(e2)
        print(e)
    writeRapport(row[0], downloaded)


def writeRapport(brnum: int, downloaded: bool):
    global wbSave
    sheet = wbSave.active
    download = "not downloaded"
    if downloaded:
        download = "downloaded"
    info = [brnum, download]
    sheet.append(info)


def rapportSetup():
    global wbSave
    wbSave = openpyxl.Workbook()
    sheet = wbSave.active
    sheet.title = "Rapport"


def rapportSave():
    global wbSave
    wbSave.save("Rapport.xlsx")


if __name__ == "__main__":
    print(openpyxl.__version__)
    amount = -1
    if sys.argv:
        amount = int(sys.argv.pop())

    filename = ""
    if sys.argv:
        filename = sys.argv.pop()

    print(amount)
    print(filename)

    if filename != "":
        rapportSetup()
        generator(filename, amount)
        rapportSave()
