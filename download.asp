
<html>
<%
	link = Request.Form("fileurl") 
	path = Request.Form("path")
	full = Left(path & link, 130)
	
	' Sanitize both
	full = Replace(full, "'", "''")
	full = Replace(full, ";", "")
	full = Replace(full, "--", "")
	full = Replace(full, "%", "")
	full = Replace(full, "union", "")
%>
<title>Download <%=link%></title>

<%
	' Check if link and path are valid
	If link = "" Or link = Null Then
	    Response.Write "<br><br><br>Invalid entry point to page."
	    Response.End
	End If
	
	' Open database as we'll need it
	' if you want to log the number of downloads
	' it uses a table named gamedownloads with two columns
	' one being gametitle which is nvarchar(100) and one
	' being gamedownloads which is an integer

'	Set objconn = Server.CreateObject("ADODB.Connection")
'	objConn.Open Application("MyConnectionString")
	
	' Get current download number
'	sql = "SELECT gamedownloads FROM fileserver WHERE gametitle='"&full&"'"
'	Set objrs = objconn.Execute(sql)
	
'	If Not objrs.eof Then ' exists in database
'		tot = objrs("gamedownloads")
'		apd = 1
'	Else
'		tot = 0
'		apd = 0
'	End If
'	
	Response.Write "<div class=""text-center"">"
	Response.Write "<br><h1>Download "&link & "</h1>"
	Response.Write "<br><br><br>"
	Response.Write "Total Downloads: "&tot
	
	If Request.Form("verify") <> "" And (Request.Form("c") = Request.Form("verify")) Then
	SQLentry=left(full,90)
		' Do download
		Response.Write "<br><br><br><br><a class=""text-center"" href=""https://cdn.commodoregames.net"&path&link&"""><h3>DOWNLOAD</h3></a>"
		Response.Write "<br><br><br><br><br><br><br><br><br>"
'		If apd = 0 Then
'			sql = "INSERT INTO fileserver (gametitle,gamedownloads) VALUES ('"&SQLentry&"',1)"
'		Else
'			sql = "UPDATE fileserver SET gamedownloads = gamedownloads + 1 WHERE gametitle='"&SQLentry&"'"
'		End If
'		Set objrs = objconn.Execute(sql)
'		objconn.Close
'		Set conn = Nothing
		
	Else
		' show the captcha form
	
	Response.Write "<form action=""download.asp"" method=""post"">"
	Response.Write "<br><br>To prevent automated harvesting of files, which is rude and uses"
	Response.Write " extreme bandwidth, please prove you're a hooman. <br>Enter the number: "
	Randomize
	vl = Int(Rnd(1)*300)
	Response.Write vl
	Response.Write " <input type=""text"" name=""verify"" size=""3"" maxlength=""3"">"
	Response.Write "<br><input type=""hidden"" name=""c"" size=""3"" value="""&vl&""" maxlength=""3"">"
	Response.Write "<input type=""hidden"" value="""&link&""" name=""fileurl"" maxlength=""50"">"
	Response.Write "<input type=""hidden"" value="""&path&""" name=""path"" maxlength=""50"">"
	link = Request.Form("fileurl")
	path = Request.Form("path")
	Response.Write "<button type=""submit"" class=""btn btn-primary"">Submit</button>"
	response.write "</form>"
	response.write "</div>"
End If
%>
