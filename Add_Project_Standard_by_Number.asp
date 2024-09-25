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
Dim dbname
Dim Cnpath
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError


Dim Project
Dim Standard

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

dim objRS1
dim objRS2

Set objRS1 = objDB.Execute("select * from [Project]")
Set objRS2 = objDB.Execute("select * from [Standards]")





sAction = Request("action")

Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "Please make your changes then press the Update button.<p>"
	html = html & "<font color=red>" & sError & "</font><p>"

	html = html & "<form name=form1 method=Post action=Add_Project_Standard.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Project:</font></td>"
    
    Html = html & "<td><select  name=Project>"
			Do While Not objRS1.EOF
			html = html & "<option "
			If Project = (objRS1("Projectname")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS1("Projectname") & Chr(34) & ">" &objRS1("Projectname")
	
			objRS1.MoveNext
			Loop
				   		
    	Html = html & "</select></td></tr>"
	
	
	
	'html = html & "<tr bgcolor=" & sRowColor &"><td>Standard:</td><td><input size=20 name=Standard value=" & Chr(34) & Standard & Chr(34) & "></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Standards:</font></td>"
    
    Html = html & "<td><select  name=Standard>"
			Do While Not objRS2.EOF
			html = html & "<option "
			If Project = (objRS2("SNumber")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS2("SNumber") & Chr(34) & ">" &objRS2("SNumber")
	
			objRS2.MoveNext
			Loop
				   		
    	Html = html & "</select></td></tr>"
	
	
	
	
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
	Project = Trim(Request.Form("Project"))
	Standard = Trim(Request.Form("Standard"))

	

	If Project = "" Then
		sError = sError & "Project is a required field.<br>"
	End If 

	If Standard = "" Then
		sError = sError & "Standard is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
		sql = "Insert Into Project_Standard ("
		sql = sql & "Project,"
		sql = sql & "Standard"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(Project) & "',"
		sql = sql & "'" & SqlQuote(Standard) & "'"
		sql = sql & ");"



		'response.write sql
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