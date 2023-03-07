# Accessing-Bunny.Net-CDN-via-Classic-ASP
How to access Bunny CDN using Classic ASP and create a file server

Requirents:
Classic ASP
Bunny CDN account
ServerXMLHTTP Component

Steps:
1) Edit the fileserver.asp to configure the STORAGE_ZONE to match your own Storage Zone name.

2) Add your secret key. This is typically the FTP password for the zone.
In the code it is located here
Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "GET", BUNNY_CDN, False
httpRequest.setRequestHeader "AccessKey", "INSERT_API_KEY_HERE"

3) You are done. The program will read your directory and allow users to navigate the folders. It will also display a
DOWNLOAD button for files to be downloaded. This uses a file named "DOWNLOAD.ASP". I did this because I wanted to prevent
people from scraping content.

4) In the DOWNLOAD.ASP file, the number of downloads are stored in a database file. This can be removed from your code if you want.

See it in action at CommodoreGames.Net

c64milasoft@gmail.com
