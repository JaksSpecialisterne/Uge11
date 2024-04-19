# requirements
First you need to make sure you have all needed requirements, you can type in your command console: "pip install -r requirements.txt"

The multithreaded version is the most up to date version, as everything made after threading was implemented was only made on the multithreaded version

# running the program single threaded
To run the program the following should be typed in the command console: "PDF_Downloader.py filename amount", where filename should be replaced by the filename of the excel that contains the pdf links and amount should be replaced by the amount of pdf's that should be downloaded, amount can be omitted and left empty and doing so will download everything.


# running the program multithreaded
To run the program the following should be typed in the command console: "PDF_Downloader.py filename amount workers timeoutTime".
"filename" should be replaced by the filename of the excel that contains the pdf links. 
"amount" should be replaced by the amount of pdf's that should be downloaded, amount can be set to "-1" to download everything.
"workers" should be replaced by the amount of threads the program shall attempt to use while running to program, more workers can potential decrease the overall time but if set too high can just as easily increase it. Default is 5.
"timeoutTime" should be replaced by the maximum amount of time that the program will wait for each url to respond. Default is 15.