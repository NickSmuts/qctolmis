<%

Option Explicit

Dim sRowColor
Dim objDB
Dim objRS
Dim objRS1
Dim objRS2
Dim Snum
Dim Competent
Dim sDBName
Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Snum = request.form("Student_num")
Competent ="Competent"




Set objRS = objDB.Execute("select * from DATA where Student_num = '"& request("StudentNum") & "'")
Set objRS1 = objDB.Execute("SELECT DATA.P_Title, DATA.FName, DATA.Sname, DATA.Student_NUM, LearnerData.SCompetent, Standards.SNumber,Standards.Stitle,Standards.CType,Standards.Scredits, DATA.Project FROM (DATA INNER JOIN LearnerData ON DATA.Student_NUM = LearnerData.Student_NUM) INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE DATA.Student_NUM = '"& request("StudentNum") & "' AND SCompetent ='Competent' ")
Set objRS2 = objDB.Execute("SELECT Sum(Standards.SCredits) AS SumOfSCredits, LearnerData.Student_NUM, LearnerData.SCompetent FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle GROUP BY LearnerData.Student_NUM, LearnerData.SCompetent HAVING LearnerData.Student_NUM = '"& request("StudentNum") & "' AND SCompetent ='Competent' ")




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

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
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

'Response.Write("<Blockquote>")
Response.Write("<table border=0 cellpadding=1 bordercolor=#111111 width=700 >")


sRowColor = "ffffff"




Do While Not objRS.EOF
Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=2><b><font size=3 face=Verdana>Name</font></b></td>")
Response.Write("    <td rowspan=4 align=left valign=top><img border=0 src=Photo/" & objRS("Photo") & ".jpg width=140 height=100></td>")
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


Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><b><font size=3 face=Verdana>Address</font></b></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><font size=3 face=Verdana>" & objRS("Addres") & "<BR>")
Response.Write(" " & objRS("Address") & "<BR>")
Response.Write(" " & objRS("P_code") & "<BR>")
Response.Write(" " & objRS("City") & "<BR>")
Response.Write(" " & objRS("Province") & "<BR>")

Response.Write("</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Contact Number</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Contact Cell</font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana>Language</font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Contact_num") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Contact_cell") & "</font></td>")
Response.Write("    <td width=""34%""><font size=3 face=Verdana>" & objRS("Language") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Sex</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Disability</font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana>Age</font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Sex") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Disability") & "</font></td>")
Response.Write("    <td width=""34%""><font size=3 face=Verdana>" & objrs("Age") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Status</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Highest Education</font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana>Year Achieved</font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Marital_status") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Education") & "</font></td>")
Response.Write("    <td width=""34%""><font size=3 face=Verdana>" & objRS("Year") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Training Group</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Client</font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Training_group") & "</Font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Client") & "</Font></td>")
Response.Write("    <td width=""34%"">&nbsp;</td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Banking Details</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Bank_name") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Bank_branch") & "</font></td>")
Response.Write("    <td width=""34%""><font size=3 face=Verdana>" & objRS("Bank_ibt") & "</font></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Bank_account") & "</font></td>")
Response.Write("    <td width=""33%"">&nbsp;</td>")
Response.Write("    <td width=""34%"">&nbsp;</td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>National Qualifications</font></B></td>")
Response.Write("    <td width=""33%""><B><font size=3 face=Verdana>Project Name</font></B></td>")
Response.Write("    <td width=""34%""><B><font size=3 face=Verdana></font></B></td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objrs("NATQUA") & "</font></td>")
Response.Write("    <td width=""33%""><font size=3 face=Verdana>" & objRS("Project") & "</font></td>")
Response.Write("    <td width=""34%"">&nbsp;</td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("    <td width=""33%""></td>")
Response.Write("    <td width=""33%"">&nbsp;</td>")
Response.Write("    <td width=""34%"">&nbsp;</td>")
Response.Write("  </tr>")

Response.Write("<tr bgcolor=" & sRowColor & ">")
Response.Write("   <td colspan=3><hr color=#996600 size=1></td>")
Response.Write("  </tr>")

	
	objRS.MoveNext
Loop

Response.Write("</table>")

Response.Write(" <br>")

If objRS1.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS1.Close
	
	Set objRS1 = Nothing

	Response.End
End If

Response.Write("   <b><font size=3 face=Verdana>Standards Completed.</font></b><br>")
Response.Write(" <br>")
Response.Write("<table border=1 cellpadding=4 cellspacing=0 width=700>")
Response.Write("<tr bgcolor=ffffff>")
Response.Write("<td><font size=3 face=Verdana><B>Standard Number</b></font></td>")
Response.Write("<td><font size=3 face=Verdana><B>Standard Title</b></font></td>")
Response.Write("<td><font size=3 face=Verdana><B>Course Type</b></font></td>")
Response.Write("<td><font size=3 face=Verdana><B>Standard Credits</b></font></td>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS1.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Snumber") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Stitle") & "</font></td>")
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("CType") & "</font></td>")
	
	Response.Write("<td><font size=2 face=Verdana>" & objRS1("Scredits") & "</font></td>")
	Response.Write("</tr>")
	objRS1.MoveNext
Loop



	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font size=2 face=Verdana>&nbsp;</font></td>")
	Response.Write("<td><font size=2 face=Verdana>Total Credits</font></td>")
	
	Response.Write("<td><font size=2 face=Verdana>" & objRS2("SumOfSCredits") & "</font></td>")
	Response.Write("</tr>")


Response.Write("</table>")

Response.Write(" <br>")
		
			
'Response.Write("</Blockquote>")

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