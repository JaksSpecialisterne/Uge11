import sys, urllib, openpyxl, socket, time, urllib.request, urllib.error
from itertools import islice

BRNUM = 0
NAME = 2
SECTOR = 7
COUNTRY = 8
DATE_ADDED = 11
TITLE = 12
PUBLICATION_YEAR = 13
FIRST_DOWNLOAD_LINK = 37
SECOND_DOWNLOAD_LINK = 38
SHEET_NAME = "0"

TIME_TXT = "Singlethread/elapsedTime.txt"
timeList = []

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
        downloader(
            [
                row[BRNUM],
                row[FIRST_DOWNLOAD_LINK],
                row[SECOND_DOWNLOAD_LINK],
                row[NAME],
                row[SECTOR],
                row[COUNTRY],
                row[DATE_ADDED],
                row[TITLE],
                row[PUBLICATION_YEAR],
            ]
        )


def downloader(row):
    url = row[1]
    downloaded = False
    savefile = f"Singlethread/{row[0]}.pdf"
    error = ""
    error2 = ""

    socket.setdefaulttimeout(15)

    start = time.time()

    try:
        urllib.request.urlretrieve(url, savefile)
        downloaded = True
    except (
        urllib.error.HTTPError,
        urllib.error.URLError,
        ConnectionResetError,
        Exception,
    ) as e:
        error = f"{e}"
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
                error2 = f"{e2}"
        else:
            error2 = "No link"

    end = time.time()
    time1 = end - start

    writeRapport(
        row[0],
        downloaded,
        row[3],
        row[4],
        row[5],
        row[6],
        row[7],
        row[8],
        error,
        error2,
    )

    global timeList
    timeList.append([row[0], time1])


def writeRapport(
    brnum: int,
    downloaded: bool,
    name: str,
    sector: str,
    country: str,
    dateAdded: str,
    title: str,
    publicationYear: int,
    error: str,
    error2: str,
):
    global wbSave
    sheet = wbSave.active
    download = "not downloaded"
    if downloaded:
        download = "downloaded"
    info = [
        brnum,
        download,
        name,
        sector,
        country,
        dateAdded,
        title,
        publicationYear,
        error,
        error2,
    ]
    sheet.append(info)


def rapportSetup():
    global wbSave
    wbSave = openpyxl.Workbook()
    sheet = wbSave.active
    sheet.title = "Rapport"
    info = [
        "BRNum",
        "Download Status",
        "Name",
        "Sector",
        "Country",
        "Date Added",
        "Title",
        "Publication Year",
        "Primary Link Error",
        "Secondary Link Error",
    ]
    sheet.append(info)


def rapportSave():
    global wbSave
    wbSave.save("Singlethread/Metadata2017_2020.xlsx")


def writeTime(downloadTime):
    global timeList
    file = open(TIME_TXT, "w")
    for entry in timeList:
        if entry[1] == 0.0:
            file.write(f"BRNum: {entry[0]}, download failed\n")
        else:
            file.write(f"BRNum: {entry[0]}, download time elapsed: {entry[1]}\n")
    file.write(f"Total download time elapsed: {downloadTime}\n")
    file.close()


if __name__ == "__main__":

    amount = -1
    if len(sys.argv) == 3:
        amount = int(sys.argv.pop())

    filename = sys.argv.pop()

    print(f"amount: {amount}")
    print(f"filename: {filename}")

    if filename != "":
        rapportSetup()
        start = time.time()
        generator(filename, amount)
        end = time.time()
        rapportSave()
        writeTime(end - start)
