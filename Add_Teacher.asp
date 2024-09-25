<%

Option Explicit

Function SQLQuote(var)
	If InStr(var, "'") <> 0 Then
		var = Replace(var, "'", "''")
	End If

	SQLQuote = var
End Function


Dim objDB
Dim objRS
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError


Dim Tname
Dim Tsname
Dim Tnoid

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


sAction = Request("action")



Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
html = html & "<font face=Verdana>Please add the Assessor details then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=Add_teacher.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>First Name:</font></td><td><input size=20 name=Tname value=" & Chr(34) & Tname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Surname:</font></td><td><input size=20 name=Tsname value=" & Chr(34) & Tsname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Assessor Number:</font></td><td><input size=20 name=Tnoid value=" & Chr(34) & Tnoid & Chr(34) & "></td></tr>"
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
	
	
Set objRS = objDB.Execute("select * from teacher")



If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If
Response.Write("<blockquote>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=ffffff>")
Response.Write("<th filter=ALL><font face=Verdana>First Name</font></th>")
Response.Write("<th filter=ALL><font face=Verdana>Surname</font></th>")
Response.Write("<th filter=ALL><font face=Verdana>Assessor Number</font></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana>" & objRS("Tname") & "</font</td>")
	Response.Write("<td><font face=Verdana>" & objRS("Tsname") & "</font></td>")
	Response.Write("<td><font face=Verdana>" & objRS("Tnoid") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")



	
	
	
	
	
	
	
End Sub

Sub ValidateData()
	Tname = Trim(Request.Form("Tname"))
	Tsname = Trim(Request.Form("Tsname"))
	Tnoid = Trim(Request.Form("Tnoid"))

	

	If Tname = "" Then
		sError = sError & "Tname is a required field.<br>"
	End If 

	If Tsname = "" Then
		sError = sError & "Tsname is a required field.<br>"
	End If 

	If Tnoid = "" Then
		sError = sError & "Tnoid is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into teacher ("
		sql = sql & "Tname,"
		sql = sql & "Tsname,"
		sql = sql & "Tnoid"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(Tname) & "',"
		sql = sql & "'" & SqlQuote(Tsname) & "',"
		sql = sql & "'" & SqlQuote(Tnoid) & "'"
		sql = sql & ");"


		'response.write sql
		ObjDB.Execute(sql)

				If Err = 0 Then
			Response.Write "<Blockquote>"
			Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			Response.Write "<p><font face=Verdana><a href=Admin.asp>Administration</a></font></P>"
			Response.Write "</blockquote>"
			PageEnd()
			Response.End
		End If
	End If
	
	objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

End Sub



Sub PageStart()
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
    <td><%
    
End Sub
Sub PageEnd()
%> </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>
<%
End Sub


Select Case sAction
	Case ""
		PageStart()
		DisplayForm()
		PageEnd()

	Case "Update"
	    PageStart()
		ValidateData()
        PageEnd()
	

End Select
%>