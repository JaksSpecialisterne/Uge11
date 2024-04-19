import sys, urllib, openpyxl, socket, time, urllib.request, urllib.error, concurrent.futures, json
from itertools import islice
from operator import itemgetter

SHEET_NAME = "0"
TIME_TXT = "Multithread/elapsedTime.txt"

timeList = []

wbSave = None


def generator(filename: str, amount: int, workers: int, timeoutTime: int):
    wb = openpyxl.load_workbook(filename=filename, read_only=True)
    ws = wb[SHEET_NAME]
    i = 0

    pool = concurrent.futures.ThreadPoolExecutor(max_workers=workers)
    with open("constants.json", "r") as f:
        data = json.load(f)
        for row in islice(ws.iter_rows(values_only=True), 1, None):
            if amount > 0:
                i += 1
                if i >= amount + 1:
                    break
            con = list(data.values())
            row2 = itemgetter(*con[0:9])(row)
            pool.submit(
                downloader,
                row2,
                timeoutTime,
            )
    pool.shutdown(wait=True)


def downloader(row, timeoutTime):
    url = row[1]
    downloaded = False
    savefile = f"Multithread/{row[0]}.pdf"
    error = ""
    error2 = ""
    timeElapsed = 0.0

    socket.setdefaulttimeout(timeoutTime)

    start = time.time()
    try:
        urllib.request.urlretrieve(url, savefile)
        downloaded = True
    except (
        urllib.error.HTTPError,
        urllib.error.URLError,
        ConnectionResetError,
        socket.error,
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
                socket.error,
                Exception,
            ) as e2:
                error2 = f"{e2}"
        else:
            error2 = "No link"

    end = time.time()
    if downloaded:
        timeElapsed = end - start

    writeRapport(
        row[0],
        downloaded,
        error,
        error2,
        row[3],
        row[4],
        row[5],
        row[6],
        row[7],
        row[8],
    )

    global timeList
    timeList.append([row[0], timeElapsed])


def writeRapport(
    brnum: int,
    downloaded: bool,
    error: str,
    error2: str,
    name: str,
    sector: str,
    country: str,
    dateAdded: str,
    title: str,
    publicationYear: int,
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
    wbSave.save("Multithread/Metadata2017_2020.xlsx")


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

    timeoutTime = int(sys.argv.pop())
    workers = int(sys.argv.pop())
    amount = int(sys.argv.pop())
    filename = sys.argv.pop()

    print(f"timeoutTime: {timeoutTime}")
    print(f"worker: {workers}")
    print(f"amount: {amount}")
    print(f"filename: {filename}")

    if filename != "":
        rapportSetup()
        start = time.time()
        generator(filename, amount, workers, timeoutTime)
        end = time.time()
        rapportSave()
        writeTime(end - start)
