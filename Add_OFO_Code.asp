<%
'New OFO Code page created by Rodney Addo
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


Dim OFOCode
Dim OFODesc


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
html = html & "<font face=Verdana>Please add OFO Code and OFO Description then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=Add_OFO_Code.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>OFO Code:</font></td><td><input size=20 name=OFOCode value=" & Chr(34) & OFOCode & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>OFO Description:</font></td><td><input size=20 name=OFODesc value=" & Chr(34) & OFODesc & Chr(34) & "></td></tr>"
	

	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana></font></td><td></td></tr>"
	
	
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
	
	
Set objRS = objDB.Execute("select * from OFOCode")



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
Response.Write("<th width='200' align='left' filter=ALL><font face=Verdana>OFO Code</font></th>")
Response.Write("<th width='400' align='left' filter=ALL><font face=Verdana>OFO Description</font></th>")

Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td width='200'><font face=Verdana>" & objRS("OFOCode") & "</font</td>")
	Response.Write("<td width='400'><font face=Verdana>" & objRS("OFODesc") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")



	
	
	
	
	
	
	
End Sub

Sub ValidateData()
	OFOCode = Trim(Request.Form("OFOCode"))
	OFODesc = Trim(Request.Form("OFODesc"))
	

	

	If OFOCode = "" Then
		sError = sError & "OFOCode is a required field.<br>"
	End If 

	If OFODesc = "" Then
		sError = sError & "OFODesc is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into OFOCode ("
		sql = sql & "OFOCode,"
		sql = sql & "OFODesc"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(OFOCode) & "',"
		sql = sql & "'" & SqlQuote(OFODesc) & "'"
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