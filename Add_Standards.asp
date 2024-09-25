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
Dim STitle
Dim Snumber
Dim Ctype
Dim SCredits

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
	html = html & "<font face=Verdana>Please add the Standards then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"

	html = html & "<form name=form1 method=Post action=Add_Standards.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Title:</font></td><td><input size=20 name=STitle value=" & Chr(34) & STitle & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Number:</font></td><td><input size=20 name=Snumber value=" & Chr(34) & Snumber & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>CType:</font></td><td><input size=20 name=CType value=" & Chr(34) & CType & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Credits:</font></td><td><input size=20 name=SCredits value=" & Chr(34) & SCredits & Chr(34) & "></td></tr>"
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
	
	
	Set objRS = objDB.Execute("select * from Standards")



If objRS.EOF Then
	Response.Write("<b><font face=Verdana> No matching records found.</font></b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If

Response.Write("<blockquote>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")

Response.Write("<tr bgcolor=cccccc>")
Response.Write("<td filter=ALL><font face=Verdana>Title</font></td>")
Response.Write("<td filter=ALL><font face=Verdana>Number</font></td>")
Response.Write("<td filter=ALL><font face=Verdana>Credits</font></td>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana>" & objRS("STitle") & "</font></td>")
	Response.Write("<td><font face=Verdana>" & objRS("Snumber") & "</font></td>")
	Response.Write("<td><font face=Verdana>" & objRS("Ctype") & "</font></td>")
	Response.Write("<td><font face=Verdana>" & objRS("SCredits") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")

objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

	
	
	
	
End Sub

Sub ValidateData()
	STitle = Trim(Request.Form("STitle"))
	Snumber = Trim(Request.Form("Snumber"))
	CType = Trim(Request.Form("CType"))
	SCredits = Trim(Request.Form("SCredits"))

	'TODO: Modify/Delete the field validations below to match your situation...

	If STitle = "" Then
		sError = sError & "Title is a required field.<br>"
	End If 

	If Snumber = "" Then
		sError = sError & "Number is a required field.<br>"
	End If 
	
	If Ctype = "" Then
		sError = sError & "Catergory Type is a required field.<br>"
	End If 

	If SCredits = "" Then
		sError = sError & "Credits is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		
		sql = "Insert Into Standards ("
		sql = sql & "STitle,"
		sql = sql & "Snumber,"
		sql = sql & "Ctype,"
		sql = sql & "SCredits"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(STitle) & "',"
		sql = sql & "'" & SqlQuote(Snumber) & "',"
		sql = sql & "'" & SqlQuote(CType) & "',"
		sql = sql & "'" & SqlQuote(SCredits) & "'"
		sql = sql & ")"

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
		DisplayForm()
		PageEnd()

	Case "Update"
	    PageStart()
		ValidateData()
        PageEnd()
	

End Select
%>