<%

Option Explicit

Dim sRowColor
Dim objDB
Dim objRS
Dim objRS1
Dim objRS2
Dim sDBName
Dim dbname
Dim Cnpath
dim STUDENTID
dim Studentnum
dim STUDID
Dim EnrolldateY
Dim EnrolldateM
Dim EnrolldateD
Dim Project



dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName



StudentNum = Request.form("Studentnum")

Set objRS = objDB.Execute("select * from DATA where Student_num = '"& StudentNum & "'")

Set objRS1 = objDB.Execute("SELECT LearnerData.Student_NUM, LearnerData.STitle, LearnerData.SCompetent, Standards.SNumber, Standards.SCredits, LearnerData.AssessorID, LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle where Student_num = '"& request.form("StudentNum") & "'")

Set objRS2 = objDB.Execute("SELECT LearnerData.Student_NUM, Sum(Standards.SCredits) AS SumOfSCredits FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle GROUP BY LearnerData.Student_NUM HAVING (((LearnerData.Student_NUM)='"& request.form("StudentNum") & "'))")

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

Do While Not objRS.EOF
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=2><b><font size=3 face=Verdana>Name</font></b></td>")
Response.Write("    <td rowspan=4 align=right valign=top></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=2><font size=3 face=Verdana>" & objRS("P_title") & "&nbsp;" & objRS("Fname") & "&nbsp;" & objRS("Sname") & "</font></td>")
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
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</Blockquote>")

   Response.Write("<Blockquote>")

Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr>")
Response.Write("<td><font size=3 face=Verdana>To add new standards to the learner.</font></td>")
Response.Write("</tr>")
Response.Write("<tr>")
Response.Write("<td><font size=3 face=Verdana>")
Response.Write("<form method=POST action=Addtraining.asp><input type=hidden name=StudentNum value=" & Studentnum & ">  <input type=hidden name=Project value=" & Project & "><input type=submit value=Standards name=""B1"" ></form>")
Response.Write("</font></td>")
Response.Write("</tr>")
Response.Write("</table>")
Response.Write("</Blockquote>")	

Response.Write("<Blockquote>")
If objRS1.EOF Then
	Response.Write("<font size=2 face=Verdana>No Standards records found.</font>")
	objRS1.Close
	Set objRS1 = Nothing
	Response.End
End If

Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Title</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Number</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Credits</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Assessor Number</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Competent</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana>Completed Date</font></td>")
Response.Write("<td filter=ALL><font size=3 face=Verdana></font></td>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS1.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Stitle") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Snumber") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("SCredits") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("AssessorID") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Scompetent") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("EnrolldateY") & "/" & objRS1("EnrolldateM") & "/" & objRS1("Enrolldated") & "</font></td>")
	Response.Write("<td><form method=POST action=Competent.asp><input type=hidden name=Stitle value='" & objRS1("Stitle") & "'><input type=hidden name=StudentNum value=" & objRS1("Student_num") & "><input type=submit value=Edit name=""B1"" style=""font-size: 8pt; font-family: Verdana""></form></td>")
	Response.Write("</tr>")
	objRS1.MoveNext
Loop
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font size=2 face=Verdana></font></td>")
	Response.Write("<td><font size=2 face=Verdana><u>Total Credits</u></font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS2("SumOfSCredits") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana></font></td>")
	Response.Write("<td filter=ALL><font size=3 face=Verdana></font></td>")
	Response.Write("</tr>")
Response.Write("</table>")
Response.Write("</Blockquote>")


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

objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>