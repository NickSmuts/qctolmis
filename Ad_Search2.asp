<%

Option Explicit


Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName
Dim SQL
Dim dbname
Dim Cnpath
Dim Fname
Dim Lname
Dim IDNUMBER

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


Fname = Request.form("Fname")
Lname = Request.form("Lname")
IDNUMBER = Request.form("IDNUMBER")

'SQL= (" Select FirstName, LastName ")

SQL=SQL +("Select * from Data")

If Fname <>"" then
			SQL= SQL +(" where FName like  '" & Fname & "%' ")
	Elseif Lname <> ""then
			
			SQL = SQL +(" where SName like  '" & Lname & "%' ")
	Elseif IDNUMBER <> ""then
			SQL = SQL +(" where Id_num like  '" & IDNUMBER & "%' ")
End if


If Lname <> ""then

	SQL = SQL +(" and SName like  '" & Lname & "%' ")
end if

If IDNUMBER <> ""then

	SQL = SQL +(" and Id_num like  '" & IDNUMBER & "%' ")
end if
	
If Fname= "" and Lname ="" then
	SQL = SQL +(" ORDER BY FName ASC")
End if

 
 'Response.write SQL
 'Response.end
 
Set objRS = objDB.Execute(SQL)



'Set objRS = objDB.Execute("select * from DATA")

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
Response.Write("<blockquote>")
	Response.Write("<font face=Verdana size=2><b>No matching records found.</b></font>")
	Response.Write("<p><font face=Verdana size=2><a href=nameSearch2.asp><b>Back to Search</b></a></font></p>")
Response.Write("</blockquote>")

Response.Write("<p>")
Response.Write("<p>")
Response.Write("<p>")
Response.Write("<p>")
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
Response.Write("<th filter=ALL><font face=Verdana size=2></font></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana size=2>" & objRS("P_title") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Fname") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Sname") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Id_num") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2><a href=QSearch1.asp?StudentNum=" & objRS("Student_num") & ">" & objRS("Student_num") & "</a></font></td>")
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