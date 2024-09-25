<%

Option Explicit


Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName
Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


Set objRS = objDB.Execute("select * from DATA")

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SCIENTIFICROOTS</title>
</head>

<body topmargin="0" leftmargin="2" link="#996600" vlink="#996600" alink="#996600" text="#996600" bgcolor="#FFFFFF">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="800" id="AutoNumber1">
  <tr>
    <td><!---#include file = "inc/head.asp"----></td>
  </tr>
  <tr>
    <td>
    <%

If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If

Response.Write("<table border=0 cellpadding=3 cellspacing=4>")
Response.Write("<tr bgcolor=ffffff>")

Response.Write("<th filter=ALL><font face=Verdana size=2>Title</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>First Name</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Surname</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Id Number</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Student Number</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Training Group</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Client</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Project</font></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana size=2>" & objRS("P_title") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Fname") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Sname") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Id_num") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Student_num") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Training_group") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Client") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Project") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
%>
 </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>
<%


objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>