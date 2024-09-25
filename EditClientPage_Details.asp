<%
'New Edit Client page created by Rodney Addo
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
Dim ID_no
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



Sub GetData()



	Set objRS = objDB.Execute("select * from Client where ID_no=" & request.form("ID") & "")




	If objRS.EOF Then
	 ID_no = ""
	 CompanyName = ""
	 TrainingManager = ""
	 CNumber = ""
	 SICCode = ""
	 SSUCode = ""
	Else
		ID_no = objRS("ID_no")
	 	CompanyName = objRS("CompanyName")
	 	TrainingManager = objRS("TrainingManager")
	 	CNumber = objRS("CNumber")
	 	SICCode = objRS("SICCode")
	 	SSUCode = objRS("SSUCode")
	End If
End Sub



Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana>Please edit client details then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=EditClientPage_Details.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"




	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Company Name:</font></td><td><input size=35 name=CompanyName value=" & Chr(34) & CompanyName & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Training Manager:</font></td><td><input size=35 name=TrainingManager value=" & Chr(34) & TrainingManager & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Contact Number:</font></td><td><input size=35 name=CNumber value=" & Chr(34) & CNumber & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>SICCode:</font></td><td><input size=35 name=SICCode value=" & Chr(34) & SICCode & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>SSUCode:</font></td><td><input size=35 name=SSUCode value=" & Chr(34) & SSUCode & Chr(34) & "></td></tr>"
	html = html & "</table><p>"
	html = html & "<input type=hidden name=ID_no value =" & Chr(34) & ID_no & Chr(34) & ">"

	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
    ID_no = Request.Form("ID_no")
	CompanyName = Trim(Request.Form("CompanyName"))
	TrainingManager = Trim(Request.Form("TrainingManager"))
	CNumber = Trim(Request.Form("CNumber"))
	SICCode = Trim(Request.Form("SICCode"))
	SSUCode = Trim(Request.Form("SSUCode"))





	If CompanyName = "" Then
		sError = sError & "Company Name is a required field.<br>"
	End If

	If TrainingManager = "" Then
		sError = sError & "Training Manager is a required field.<br>"
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

		sql = "UPDATE Client SET "

		sql = sql & "CompanyName='" & SqlQuote(CompanyName) & "',"
		sql = sql & "TrainingManager='" & SqlQuote(TrainingManager) & "',"
		sql = sql & "CNumber='" & SqlQuote(CNumber) & "',"
		sql = sql & "SICCode='" & SqlQuote(SICCode) & "',"
		sql = sql & "SSUCode='" & SqlQuote(SSUCode) & "'"
		sql = sql & "where ID_no=" & ID_no & ";"

		ObjDB.Execute(sql)

		If Err = 0 Then
			Response.Write "<Blockquote>"
			Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			Response.Write "<p><font face=Verdana><a href=Default.asp>Main page</a></font></P>"
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
    <td>
<%

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
		GetData()
		DisplayForm()
		PageEnd()

	Case "Update"
	    PageStart()
		ValidateData()
        PageEnd()



End Select
%>
