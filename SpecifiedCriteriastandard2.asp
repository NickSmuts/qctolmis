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
Dim sRowColor
Dim objDB
Dim objRS

Dim sDBName
Dim SQL

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


Stan = request.form("D1")
Proj = request.form("D2")


'response.write Stan
'response.write Proj

'response.end



if proj = "All" then
response.redirect ("SpecifiedCriteriastandard.asp")
end if

if Stan = "All" then
response.redirect ("SpecifiedCriteriastandard.asp")
end if
	
Title = title + ("<font face=""Verdana"" size=""2"">Project: " & Proj & "  </font><BR>") 	



 'SQL =("SELECT Project_Standard.Project, Project_Standard.Standard, LearnerData.Student_NUM, LearnerData.SCompetent, DATA.FName, DATA.Sname, DATA.ID_NUM, DATA.Contact_NUM FROM DATA INNER JOIN (Project_Standard INNER JOIN LearnerData ON Project_Standard.Standard = LearnerData.STitle) ON DATA.Student_NUM = LearnerData.Student_NUM WHERE (((Project_Standard.Project)='" & Proj & "') AND ((Project_Standard.Standard)='" & Stan & "'))ORDER BY LearnerData.Student_NUM ")

'SQL =("SELECT Project_Standard.Project, Project_Standard.Standard, LearnerData.Student_NUM, LearnerData.SCompetent, DATA.FName, DATA.Sname, DATA.ID_NUM, DATA.Contact_NUM, DATA.P_Title FROM DATA INNER JOIN (Project_Standard INNER JOIN LearnerData ON Project_Standard.Standard = LearnerData.STitle) ON (DATA.Project = Project_Standard.Project) AND (DATA.Student_NUM = LearnerData.Student_NUM) WHERE Project_Standard.Project='" & Proj & "' AND Project_Standard.Standard='" & Stan & "' ORDER BY LearnerData.Student_NUM ")

SQL =("SELECT  * FROM DATA INNER JOIN Project_Standard INNER JOIN LearnerData ON Project_Standard.Standard = LearnerData.STitle ON DATA.Student_NUM = LearnerData.Student_NUM AND DATA.Project = Project_Standard.Project WHERE Project_Standard.Project='" & Proj & "' AND Project_Standard.Standard='" & Stan & "' ")

Set objRS = objDB.Execute(SQL)

response.write title
response.write SQL

'response.end




If objRS.EOF Then
Response.Write("<br>")
Response.Write("<br>")
	Response.Write("<font face=""Verdana"" size=""2""><b>No matching records found.</b></font>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If



response.write icount
Response.Write("<blockquote>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=silver>")


Response.Write("<th filter=ALL><font size=2 face=Verdana>Student Number</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>Title</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>Id Number</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>First Name</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>Surname</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>Contact Number</font></th>")
Response.Write("<th filter=ALL><font size=2 face=Verdana>Competent</font></th>")

Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Student_num") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("P_title") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Id_num") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Fname") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Sname") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Contact_num") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS("Scompetent") & "</font></td>")
	
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")
response.write ("<font face=""Verdana"" size=""2"">" & icount & " records found.</font>")


objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>

 </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>