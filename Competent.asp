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
Dim objrs2
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError


Dim Student_num
Dim Stitle
Dim Scompetent
Dim EnrolldateY
Dim EnrolldateM
Dim EnrolldateD
Dim AssessorID


Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


Student_num = Request.form("Studentnum")
Stitle = Request.form("Stitle")

Set objRS2 = objDB.Execute("select * from [teacher]")

sAction = Request("action")

Sub GetData()
	
	Set objRS = objDB.Execute("select * from LearnerData where Student_num = '"& Student_num &"' and Stitle = '"& Stitle &"' ")

	If objRS.EOF Then
	
		Student_num = ""
		Stitle = ""
		Scompetent = ""
	Else
	
		Student_num = objRS("Student_num")
		Stitle = objRS("Stitle")
		Scompetent = objRS("Scompetent")
	End If
End Sub

Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana>Please select ""Competent"" or ""Not Yet""  and completion Date then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"

	html = html & "<form name=form1 method=Post action=Competent.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Student Number:</font></td><td>"& Student_num & "</td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Standard Title:</font></td><td>"& Stitle &"</td></tr>"
		html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Competent:</font></td>"
	 	Html = html & "<td><select  name=Scompetent>"
		Html = html & "<option value=Competent>Competent</option>"
		Html = html & "<option value=Not_Yet>Not Yet</option>"
			
		Html = html & "</select></td></tr>"

html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Assessor:</font></td>"
	
			Html = html & "<td><select  name=AssessorID>"
			Do While Not objRS2.EOF
			html = html & "<option "
			Html = html &"value=" & Chr(34) & objRS2("Tnoid") & Chr(34) & ">" &objRS2("Tname")
			objRS2.MoveNext
			Loop
			Html = html & "</select></td></tr>"	
	
	
html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Date</font></td>"	
	 html = html & "<td>"
html = html & "	 <p><select size=1 name=D1>"
html = html & "  <option value=2004>2004</option>"
html = html & "  <option value=2005>2005</option>"
html = html & "  <option value=2006>2006</option>"
html = html & "  <option value=2007>2007</option>"
html = html & "  <option value=2008>2008</option>"
html = html & "  <option value=2009>2009</option>"
html = html & "  <option value=2010>2010</option>"
html = html & "  <option value=2011>2011</option>"
html = html & "  <option value=2012>2012</option>"
html = html & "  <option value=2013>2013</option>"
html = html & "  <option value=2014>2014</option>"
html = html & "  <option value=2015>2015</option>"
html = html & "  <option value=2016>2016</option>"
html = html & "  <option value=2017>2017</option>"
html = html & "  <option value=2018>2018</option>"
html = html & "  <option value=2019>2019</option>"
html = html & "  <option value=2020>2020</option>"
html = html & "  <option value=2021>2021</option>"
html = html & "  <option value=2022>2022</option>"
html = html & "  <option value=2023>2023</option>"
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
	
	html = html & "<input type=hidden name=Student_num value =" & Chr(34) & Student_num & Chr(34) & ">"
	
	html = html & "<input type=hidden name=Stitle value =" & Chr(34) & Stitle & Chr(34) & ">"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
	Student_num = Trim(Request.Form("Student_num"))

	Stitle = Trim(Request.Form("Stitle"))
	Scompetent = Trim(Request.Form("Scompetent"))
	Enrolldatey = Trim(Request.Form("D1"))
	AssessorID = Trim(Request.form("AssessorID"))
Enrolldatem = Trim(Request.Form("D2"))
Enrolldated = Trim(Request.Form("D3"))

	

	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
		


		
		sql = "UPDATE LearnerData SET "
		sql = sql & "Student_num='" & SqlQuote(Student_num) & "',"
		sql = sql & "Stitle='" & SqlQuote(Stitle) & "',"
		sql = sql & "Scompetent='" & SqlQuote(Scompetent) & "',"
		sql = sql & "AssessorID='" & SqlQuote(AssessorID) & "',"
		sql = sql & "Enrolldatey='" & SqlQuote(Enrolldatey) & "',"
		sql = sql & "Enrolldatem='" & SqlQuote(Enrolldatem) & "',"
		sql = sql & "Enrolldated='" & SqlQuote(Enrolldated) & "'"
		sql = sql & " WHERE Student_num = '"& Student_num &"' and Stitle = '"& Stitle &"';"
	
		

		'response.write sql
		'response.end
		ObjDB.Execute(sql)

	
objDB.Close

Set objDB = Nothing

If Err = 0 Then
			'Response.Write "<Blockquote>"
			'Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			'Response.Write "<p><font face=Verdana><a href=Default.asp>Main page</a></font></P>"
			'Response.Write "</blockquote>"
			'PageEnd()
			Response.redirect "Qsearchtraining.asp?StudentID="& Student_num &" "
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