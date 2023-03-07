<html>
<head>
<title>Commodore File Server</title>
</head>

<%
' this code is written in classic asp and will connect to
' your bunny.net cdn. it will read the list of folders and
' files and allow you to display them and download the
' files. it works as an effective file server
%>


<br>
<table class="table my-0 mx-0" >
  <thead class="thead-light">
    <tr>
      <th scope="col">Folder</th>
      <th scope="col">Last Modified</th>
      <th scope="col">Download</th>
      <th scope="col">Downloads</th>
      </tr>
  </thead>
  <tbody>

<%
' Define the STORAGE_ZONE constant variable. change this to your own
' storage zone name
'
' change the BUNNY_CDN as well to your location-specific Bunny CDN
'


Const STORAGE_ZONE = "scene"

Dim httpRequest, responseText, json, folder, BUNNY_CDN

' has a folder name been passed via querystring? if so
' add it to the folder url otherwise we'll just use root folder
' if your folder name is greater than 100 chars change line below
' as we restrict folder to 100 chars

folder = Request.QueryString("folder")
If Len(folder) > 30 Then folder = Left(folder, 100)

BUNNY_CDN = "https://ny.storage.bunnycdn.com/" & STORAGE_ZONE & "/" & folder & "/"

Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "GET", BUNNY_CDN, False
httpRequest.setRequestHeader "AccessKey", "KEY GOES HERE FROM BUNNY.NET"
httpRequest.setRequestHeader "accept", "application/json"

httpRequest.Send
responseText = httpRequest.responseText

' was the request successful?
If httpRequest.Status = 200 Then

    Set json = Server.CreateObject("MSScriptControl.ScriptControl")
    json.Language = "JScript"
    json.AddCode("var data = " & responseText)

  ' -- parse the path to find out where we are in the folder hierarchy
  ' -- also prevent errors that crop up with any bad URL   

    On Error Resume Next 
    path = json.eval("data[0].Path") 
    On Error Goto 0

' -- open SQL to show counts of downloads --

Set objconn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("MyConnectionString")

    ' split the path into an array of folder names

  response.write "<a href=""fileserver.asp"">HOME/</a>"
  path=replace(path,STORAGE_ZONE & "/","")  ' remove STORAGE ZONE name from the string
  folders = Split(path, "/")     ' split the path to drill down folders
    
    ' create a link for each level of the folder hierarchy
    Dim currentFolder
    currentFolder = ""
    For Each foldername In folders
        If foldername <> "" Then ' skip empty folder names
            currentFolder = currentFolder & foldername & "/"
            Response.Write "<a href=""fileserver.asp?folder=" & currentFolder & """>" & foldername & "</a>/"
        End If
    Next
    Response.Write("<br><br>")
    
  
    For Each obj In json.eval("data")
 
        isdir = obj.isdirectory  ' directory or file?
        name = obj.Objectname	 ' folder or file name
        fdate = obj.LastChanged  ' last changed date
    

' Is this a Directory?
' ====================

 If isdir = "True" and lcase(name) <> "sites" and lcase(name) <> "download" Then 
            uname = ucase(name)
            Response.Write "<tr><td><a href=""fileserver.asp?folder=" & path & name & """><img src=""/img/folder.png""> " & uname & "</a></td>"
             Response.Write "<td>" & left(fdate, 10) & "</td>"
            Response.Write("</tr>")
 End if

' Is this a file?
' ================

        If isdir <> "True" Then   ' file not directory
safe = path&name
safe= replace(safe,"'","")
safe = replace(safe,";","")
safe = replace(safe,"--","")
safe = replace(safe,"%","")
safe = replace(safe,"union","")
sql = "select gamedownloads from fileserver where gametitle = '"&safe&"'"
set objrs = objconn.execute(sql)
hits = 0
if not objrs.eof then hits = objrs("gamedownloads")

            Response.Write "<tr><td><img src=""/img/file.png""> " & name & "</a></td>"
            Response.Write "<td>" & left(fdate, 10) & "</td>"
	    Response.Write "<td>"
	    Response.Write "<form action=""download.asp"" method=""post"" rel=""nofollow"">"
	    Response.Write "<input type=""hidden"" name=""fileurl"" value=""" & name & """>"
	    Response.Write "<input type=""hidden"" name=""path"" value=""" & path & """>"
	    Response.Write "<input type=""hidden"" name=""download"" value=""c:\" & path & "55"">"
            Response.Write "<button type=""submit"">Download</button>"
            Response.Write "</form>"
	    Response.Write "</td>"
	    Response.Write "<td>"&hits&"</td>"
	    Response.Write("</tr>")
	End If
    Next    
    Set json = Nothing

Else
    Response.Write "<p>Error retrieving folder data. HTTP Status: " & httpRequest.Status & "</p>"
End If

httprequest.abort()
objconn.close
Set objconn = nothing
Set httpRequest = Nothing
response.write "</tbody></table>"
%>

