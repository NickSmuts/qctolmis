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
Dim Idnostan
Dim Stitle
Dim Snumber
Dim Ctype
Dim Scredits
Dim Standardnum

Dim dbname
Dim cnpath


dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Standardnum = request.form("Standardnum")




sAction = Request("action")

Sub GetData()
	
	Set objRS = objDB.Execute("select * from Standards where Snumber = '" & Standardnum & "' ")

	If objRS.EOF Then
		Idnostan = ""
		Stitle = ""
		Snumber = ""
		CType = ""
		Scredits = ""
	Else
		Idnostan = objRS("Idnostan")
		Stitle = objRS("Stitle")
		Snumber = objRS("Snumber")
		CType = objRS("CType")
		Scredits = objRS("Scredits")
	End If
End Sub

Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana size=2>Please make your changes then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"

	html = html & "<form name=form1 method=Post action=Edit_Standards1.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	'html = html & "<tr bgcolor=" & sRowColor &"><td>Idnostan:</td><td><input size=20 name=Idnostan value=" & Chr(34) & Idnostan & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana size=2>Standard Title:</font></td><td><input size=50 name=Stitle value=" & Chr(34) & Stitle & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana size=2>Standard Number:</font></td><td><input size=20 name=Snumber value=" & Chr(34) & Snumber & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana size=2>CType:</font></td><td><input size=20 name=CType value=" & Chr(34) & CType & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana size=2>Standard Credits:</font></td><td><input size=20 name=Scredits value=" & Chr(34) & Scredits & Chr(34) & "></td></tr>"
	html = html & "</table><p>"
	html = html & "<input type=hidden name=Idnostan value =" & Chr(34) & Idnostan & Chr(34) & ">"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
	Idnostan = Request.Form("Idnostan")
	Stitle = Trim(Request.Form("Stitle"))
	Snumber = Trim(Request.Form("Snumber"))
	CType = Trim(Request.Form("CType"))
	Scredits = Trim(Request.Form("Scredits"))

	



	If Stitle = "" Then
		sError = sError & "Stitle is a required field.<br>"
	End If 

	If Snumber = "" Then
		sError = sError & "Snumber is a required field.<br>"
	End If 

	If Scredits = "" Then
		sError = sError & "Scredits is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else


		
		sql = "UPDATE Standards SET "
		sql = sql & "Stitle='" & SqlQuote(Stitle) & "',"
		sql = sql & "Snumber='" & SqlQuote(Snumber) & "',"
		sql = sql & "Ctype='" & SqlQuote(CType) & "',"
		sql = sql & "Scredits='" & SqlQuote(Scredits) & "'"
		sql = sql & " where Idnostan = " & Idnostan & ";"

		'response.write sql
		ObjDB.Execute(sql)

		If Err = 0 Then
			'if we get here without an error, the data was updated...
			Response.Write "<font face=Verdana>Update Successful!</font>"
			Response.Write "<p><font face=Verdana><a href=Admin.asp>Administration</a></font></P>"
			Response.End
		End If
	End If
End Sub

Sub PageStart()
%><html>

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
End Sub


Select Case sAction
	Case ""
		PageStart()
		GetData()
		DisplayForm()
		Pageend()

	Case "Update"
		PageStart()
		ValidateData()
		Pageend()
	

End Select
%>