<%

Option Explicit
Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName

Dim dbname
Dim Cnpath

Dim Project

Project = request.form("project")

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName



Set objRS = objDB.Execute("SELECT DATA.Project, Count(DATA.Student_NUM) AS CountOfStudent_NUM FROM DATA GROUP BY DATA.Project HAVING Project='"& project &"' ; ")

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
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
Response.Write("<Blockquote>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=ffffff>")

Response.Write("<th filter=ALL><font face=Verdana>Project</font></th>")
Response.Write("<th filter=ALL><font face=Verdana>Student Count</font></th>")

Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana>" & Project & "</font></td>")
	Response.Write("<td align=right><font face=Verdana>" & objRS("CountOfStudent_NUM") & "</font></td>")
	

	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</Blockquote>")
Response.Write("<br>")
Response.Write("<br>")
Response.Write("<br>")
Response.Write("<br>")
Response.Write("<br>")

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