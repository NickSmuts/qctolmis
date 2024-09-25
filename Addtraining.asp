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
Dim objRS1
Dim objRS2
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError

'database variables...
Dim Student_num
Dim Stitle
Dim Scompetent
Dim Studentnum
Dim AssessorID

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Studentnum = request.form("StudentNum")

Set objRS1 = objDB.Execute("select * from [Standards]")


sAction = Request("action")



Sub DisplayForm()
		html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana>Please add Standard to learner then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	
	html = html & "<form name=form1 method=Post action=Addtraining.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Student Number:</font></td><td>"& Studentnum &"</td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Standard Title:</font></td>"
	
	
	
			Html = html & "<td><select  name=Stitle>"
			Do While Not objRS1.EOF
			html = html & "<option "
			Html = html &"value=" & Chr(34) & objRS1("Stitle") & Chr(34) & ">" &objRS1("Stitle")
			objRS1.MoveNext
			Loop
			Html = html & "</select></td></tr>"

    

	html = html & "</table><p>"
	html = html & "<input type=hidden name=Student_num value =" & Chr(34) & Studentnum & Chr(34) & ">"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
	Student_num = Trim(Request.Form("Student_num"))
	Stitle = Trim(Request.Form("Stitle"))
	Scompetent = "Not_Yet"
	



	If Student_num = "" Then
		sError = sError & "Student_num is a required field.<br>"
	End If 

	If Stitle = "" Then
		sError = sError & "Stitle is a required field.<br>"
	End If 

	If Scompetent = "" Then
		sError = sError & "Scompetent is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into LearnerData ("
		sql = sql & "Student_num,"
		sql = sql & "Stitle,"
		sql = sql & "Scompetent"
		
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(Student_num) & "',"
		sql = sql & "'" & SqlQuote(Stitle) & "',"
		sql = sql & "'" & SqlQuote(Scompetent) & "'"
		
		sql = sql & ");"

	

		'response.write sql
		ObjDB.Execute(sql)

		If Err = 0 Then
			Response.Write "<Blockquote>"
			Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			Response.Write "<p><font face=Verdana><a href=Default.asp>Main page</a></font></P>"
			Response.Write "<p><font face=Verdana><a href=Admin.asp>Admin page</a></font></P>"

			Response.Write "</blockquote>"
			Response.redirect "Qsearchtraining.asp?StudentID="& Student_num &" "
			
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

	Case "Update"
	    PageStart()
		ValidateData()
        PageEnd()



End Select
%>