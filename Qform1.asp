<%

Option Explicit

Dim sRowColor
Dim objDB
Dim objRS
dim objrs1

Dim sDBName
dim Html
Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Set objRS = objDB.Execute("SELECT DATA.Sname, DATA.FName, DATA.ID_NUM, Project.EnrolldateY, Project.EnrolldateM, Project.EnrolldateD, DATA.NATQUA, DATA.Student_NUM FROM (DATA INNER JOIN LearnerData ON DATA.Student_NUM = LearnerData.Student_NUM) INNER JOIN Project ON DATA.Project = Project.ProjectName WHERE (((LearnerData.Student_NUM)='"& request.form("StudentNum") & "'))")

Set objRS1 = objDB.Execute("SELECT LearnerData.Student_NUM, LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD FROM LearnerData WHERE (((LearnerData.Student_NUM)='"& request.form("StudentNum") & "')) ORDER BY LearnerData.EnrolldateY DESC , LearnerData.EnrolldateM DESC , LearnerData.EnrolldateD DESC ")
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
    
If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If


Response.Write("<Blockquote>")
Response.Write("<table border=0 cellpadding=2 bordercolor=#111111 width=700 >")


sRowColor = "ffffff"


Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=2><b><font size=3 face=Verdana>Name</font></b></td>")
Response.Write("    <td rowspan=4 align=right valign=top></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=2><font size=3 face=Verdana>&nbsp;" & objRS("Fname") & "&nbsp;" & objRS("Sname") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana>ID Number</font></b></td>")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana>Student Number</font></b></td>")

Response.Write("  </tr>")
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Id_num") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Student_num") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana>National Qualification</font></b></td>")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana></font></b></td>")

Response.Write("  </tr>")
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("NATQUA") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana></font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana>Enroll Date</font></b></td>")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana></font></b></td>")

Response.Write("  </tr>")
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("EnrolldateY") & "/" & objRS("Enrolldatem") & "/" & objRS("Enrolldated") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana></font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana>Achievement Date</font></b></td>")
Response.Write("    <td width=""33%""><b><font size=3 face=Verdana></font></b></td>")

Response.Write("  </tr>")
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS1("EnrolldateY") & "/" & objRS1("Enrolldatem") & "/" & objRS1("Enrolldated") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana></font></td>")
Response.Write("  </tr>")

Response.Write("</table>")
Response.Write("</Blockquote><BR>")
%>
<Blockquote>
<form method="POST" action="QForm_Qualification.asp">
  
  <input type="hidden" name="Firstname" value="<%=objRS("Fname")%>">
  <input type="hidden" name="NOID" value="<%=objRS("ID_Num")%>">
   <input type="hidden" name="Surname" value="<%=objRS("Sname")%>"> 
  <input type="hidden" name="NQUA" value="<%=objRS("NATQUA")%>"> 
<%  

dim enrol
dim Achieve

 Enrol =   objRS("EnrolldateY")& objRS("Enrolldatem") & objRS("Enrolldated")
 Achieve =   objRS1("EnrolldateY")& objRS1("Enrolldatem") & objRS1("Enrolldated")
  %>
 <input type="hidden" name="Enroll" value="<%=enrol%>"> 
 <input type="hidden" name="Achieve" value="<%=Achieve%>"> 
 
 <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1">
    <tr>
      <td><font face="Verdana" size="3">Learner Achievement<br>
      Status ID</font></td>
      <td>
  <select size="1" name="D1">
  <option value="01">UNKNOWN</option>
  <option value="02">ACHIEVED</option>
  <option value="03">ENROLLED</option>
  <option value="04">DE-ENROLLED</option>
  <option value="05">RE-ENROLLED</option>
  <option value="06">OTHER</option>
  </select></td>
    </tr>
    <tr>
      <td><font face="Verdana" size="3">Learner Achievement<br>
      Type ID</font></td>
      <td>
  
    <select size="1" name="D2">
  <option value="01">UNKNOWN</option>
  <option value="02">PRIOR LEARNING</option>
  <option value="03">DISTANCE LEARNING</option>
  <option value="04">CONTACT MODEL</option>
  <option value="05">WORK PLACE LEARNING</option>
  <option value="06">OTHER</option>
  <option value="07">MIXED MODE</option>
  </select></td>
    </tr>
  </table><BR>
  
 <input type="submit" value="Quick Form" name="B1"></p>
</form>
</Blockquote>




 </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>

<%

objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>