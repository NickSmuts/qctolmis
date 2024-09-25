<%
'New Client page created by Nick Smuts
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


Dim CompanyName
Dim TrainingManager
Dim CNumber
Dim SICCode
Dim SSUCode



Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


sAction = Request("action")



Sub DisplayForm()
	
	
	
Set objRS = objDB.Execute("select * from Client")



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
Response.Write("<th width='30' align='left'><font face=Verdana size='2'>ID No</font></th>")
Response.Write("<th width='300' align='left'><font face=Verdana size='2'>Company Name</font></th>")
Response.Write("<th width='200' align='left'><font face=Verdana size='2'>Training Manager</font></th>")
Response.Write("<th width='150' align='left'><font face=Verdana size='2'>Contact Number</font></th>")
Response.Write("<th width='100' align='left'><font face=Verdana size='2'>SIC Code</font></th>")
Response.Write("<th width='100' align='left'><font face=Verdana size='2'>SSU Code</font></th>")


Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td width='30' align='left'><font face=Verdana size='2'>" & objRS("ID_no") & "</font</td>")
	Response.Write("<td width='300' align='left'><font face=Verdana size='2'>" & objRS("CompanyName") & "</font</td>")
	Response.Write("<td width='200' align='left'><font face=Verdana size='2'>" & objRS("TrainingManager") & "</font></td>")
	Response.Write("<td width='150' align='left'><font face=Verdana size='2'>" & objRS("CNumber") & "</font</td>")
	Response.Write("<td width='100' align='left'><font face=Verdana size='2'>" & objRS("SICCode") & "</font></td>")
	Response.Write("<td width='100' align='left'><font face=Verdana size='2'>" & objRS("SSUCode") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")



	
	
	
	
	
	
	
End Sub

Sub ValidateData()
 CompanyName = Trim(Request.Form("CompanyName"))
 TrainingManager = Trim(Request.Form("TrainingManager"))
 CNumber = Trim(Request.Form("CNumber"))
 SICCode = Trim(Request.Form("SICCode"))
 SSUCode = Trim(Request.Form("SSUCode"))
	

	

	If CompanyName = "" Then
		sError = sError & "CompanyName is a required field.<br>"
	End If 

	If TrainingManager = "" Then
		sError = sError & "TrainingManager is a required field.<br>"
	End If 

		If CNumber = "" Then
		sError = sError & "CNumber is a required field.<br>"
	End If 

		If SICCode = "" Then
		sError = sError & "SICCode is a required field.<br>"
	End If 

		If SSUCode = "" Then
		sError = sError & "SSUCode is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into Client ("
		sql = sql & "CompanyName,"
		sql = sql & "TrainingManager,"
		sql = sql & "CNumber,"
		sql = sql & "SICCode,"
		sql = sql & "SSUCode"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(CompanyName) & "',"
		sql = sql & "'" & SqlQuote(TrainingManager) & "',"
		sql = sql & "'" & SqlQuote(CNumber) & "',"
		sql = sql & "'" & SqlQuote(SICCode) & "',"
		sql = sql & "'" & SqlQuote(SSUCode) & "'"
		sql = sql & ");"


		'response.write sql
		ObjDB.Execute(sql)

				If Err = 0 Then
			Response.Write "<Blockquote>"
			Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			Response.Write "<p><font face=Verdana><a href=Default.asp>Administration</a></font></P>"
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

	Case ""
	    PageStart()
		ValidateData()
        PageEnd()
	

End Select
%>