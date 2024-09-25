<%

Option Explicit


Function SQLQuote(var)
	If InStr(var, "'") <> 0 Then
		var = Replace(var, "'", "''")
	End If

	SQLQuote = var
End Function

'framework variables...
Dim objDB
Dim objRS
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError

'database variables...
Dim No_sec_id
Dim Fname
Dim Security_num

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
	sRowColor="#dddddd"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana>Please add Administration names and numbers then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=Add_Security.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>First Name:</font></td><td><input size=20 name=Fname value=" & Chr(34) & Fname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Security Number:</font></td><td><input size=20 name=Security_num value=" & Chr(34) & Security_num & Chr(34) & "></td></tr>"
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
	
	Fname = Trim(Request.Form("Fname"))
	Security_num = Trim(Request.Form("Security_num"))

	'TODO: Modify/Delete the field validations below to match your situation...

	

	If Fname = "" Then
		sError = sError & "Fname is a required field.<br>"
	End If 

	If Security_num = "" Then
		sError = sError & "Security_num is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into SECURITY ("
	
		sql = sql & "Fname,"
		sql = sql & "Security_num"
		sql = sql & ") "
		sql = sql & "Values ("

		sql = sql & "'" & SqlQuote(Fname) & "',"
		sql = sql & "'" & SqlQuote(Security_num) & "'"
		sql = sql & ");"

	
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