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
Dim Projectname
Dim D1
Dim D2
Dim D3
Dim EnrolldateY
Dim EnrolldateM
Dim EnrolldateD



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
	html = html & "<font face=Verdana>Please add new project details then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=Add_project.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Project Name:</font></td><td><input size=20 name=Projectname value=" & Chr(34) & Projectname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Date</font></td>"
	
	html = html & "<td>"
	
	
	 
html = html & "	 <p><select size=1 name=D1>"
html = html & "  <option value=2005>2005</option>"
html = html & "  <option value=2006>2006</option>"
html = html & "  <option value=2007>2007</option>"
html = html & "  <option value=2008>2008</option>"
html = html & "  <option value=2009>2009</option>"
html = html & "  <option value=2010>2010</option>"
html = html & "  <option value=2011>2011</option>"
html = html & "  <option value=2011>2012</option>"
html = html & "  <option value=2011>2013</option>"
'add 06/03/2011 by Nick Smuts
html = html & "  </select><select size=1 name=D2>"
html = html & "  <option value=01>01</option>"
html = html & "  <option value=02>02</option>"
html = html & "  <option value=03>03</option>"
html = html & "  <option value=04>04</option>"
html = html & "  <option value=05>05</option>"
html = html & "  <option value=06>06</option>"
html = html & "  <option value=07>07</option>"
html = html & "  <option value=08>08</option>"
html = html & "  <option value=09>09</option>"
html = html & "  <option value=10>10</option>"
html = html & "  <option value=11>11</option>"
html = html & "  <option value=12>12</option>"
html = html & "  </select>"
html = html & "  <select size=1 name=D3>"
html = html & "  <option value=01>01</option>"
html = html & "  <option value=02>02</option>"
html = html & "  <option value=03>03</option>"
html = html & "  <option value=04>04</option>"
html = html & "  <option value=05>05</option>"
html = html & "  <option value=06>06</option>"
html = html & "  <option value=07>07</option>"
html = html & "  <option value=08>08</option>"
html = html & "  <option value=09>09</option>"
html = html & "  <option value=10>10</option>"
html = html & "  <option value=11>11</option>"
html = html & "  <option value=12>12</option>"
html = html & "  <option value=13>13</option>"
html = html & "  <option value=14>14</option>"
html = html & "  <option value=15>15</option>"
html = html & "  <option value=16>16</option>"
 html = html & " <option value=17>17</option>"
html = html & "  <option value=18>18</option>"
html = html & "  <option value=19>19</option>"
html = html & "  <option value=20>20</option>"
html = html & "  <option value=21>21</option>"
html = html & "  <option value=22>22</option>"
html = html & "  <option value=23>23</option>"
html = html & "  <option value=24>24</option>"
html = html & "  <option value=25>25</option>"
html = html & "  <option value=26>26</option>"
html = html & "  <option value=27>27</option>"
html = html & "  <option value=28>28</option>"
html = html & "  <option value=29>29</option>"
html = html & "  <option value=30>30</option>"
html = html & "  <option value=31>31</option>"
 html = html & " </select> </p></td></tr>"	
	
	
	
	html = html & "</table><p>"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
	
	
	
	Set objRS = objDB.Execute("select * from Project")
	


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
Response.Write("<tr>")
Response.Write("<th filter=ALL><font face=Verdana>Project Name</font></th>")
Response.Write("<th filter=ALL><font face=Verdana>Date</font></th>")
Response.Write("</tr>")
sRowColor = "ffffff"
Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana>" & objRS("Projectname") & "</font></td>")
	Response.Write("<td><font face=Verdana>"& objRS("EnrolldateY") & "/"& objRS("EnrolldateM") & "/"& objRS("EnrolldateD") & "</font></td>")
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
	Projectname = Trim(Request.Form("Projectname"))
Enrolldatey = Trim(Request.Form("D1"))
Enrolldatem = Trim(Request.Form("D2"))
Enrolldated = Trim(Request.Form("D3"))



		
		



	If Projectname = "" Then
		sError = sError & "Projectname is a required field.<br>"
	End If 

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		'Code to add a new record...
	sql = "Insert Into Project ("
		sql = sql & "Projectname,"
		sql = sql & "Enrolldatey,"
		sql = sql & "Enrolldatem,"
		sql = sql & "Enrolldated"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(Projectname) & "',"
		sql = sql & "'" & SqlQuote(Enrolldatey) & "',"
		sql = sql & "'" & SqlQuote(Enrolldatem) & "',"
		sql = sql & "'" & SqlQuote(Enrolldated) & "'"
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