
<%

Option Explicit


Dim sRowColor
Dim objDB
Dim objRS
Dim objRS2
Dim objRs1
Dim sDBName
Dim dbname
Dim Cnpath
Dim Snum
Dim Competent

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Snum = request.form("Student_NUM")
Competent ="Competent"



Set objRS2 = objDB.Execute("SELECT DATA.FName, DATA.Sname, DATA.Student_NUM FROM DATA WHERE (((DATA.Student_NUM)='"& SNUM & "'))")
Set objRS = objDB.Execute("SELECT LearnerData.AssessorID, Standards.SNumber, Standards.STitle, Standards.SCredits, LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD, LearnerData.Student_NUM, LearnerData.SCompetent FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE (((LearnerData.Student_NUM)='"& SNUM & "')) AND (((LearnerData.SCompetent)='" & Competent & "')) ;")

Set objRS1 = objDB.Execute("SELECT Sum(Standards.SCredits) AS SumOfSCredits, LearnerData.Student_NUM, LearnerData.SCompetent FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle GROUP BY LearnerData.Student_NUM, LearnerData.SCompetent HAVING (((LearnerData.Student_NUM)='"& SNUM & "')) AND (((LearnerData.SCompetent)='" & Competent & "')) ;")


Response.Write("<html>")
Response.Write("<head>")

Response.Write("</head>")
Response.Write("<body bgcolor=white>")



If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber1">
  <tr>
    <td colspan="4" width="600"><b><font face="Verdana">LEARNER STATUS FORM <br>
    FOR UNIT STANDARDS</font></b></td>
    <td width="26">&nbsp;</td>
  </tr>
  <tr>
    <td width="47">&nbsp;</td>
    <td width="417">&nbsp;</td>
    <td width="229">&nbsp;</td>
    <td width="225">&nbsp;</td>
    <td width="26">&nbsp;</td>
  </tr>
  <tr>
    <td width="47">&nbsp;</td>
    <td width="417" rowspan="2" align="left" valign="top">
    <font face="Verdana" size="1">This form has been designed, according to SAQA 
    specifications, to load achievements that have been assessed against NQF-compliant 
    unit standards. As with course and qualifications both enrolments and 
    achievements will be tracked. Providers are urged to complete and submit 
    this form when sending in Learner Achievements to the SETA</font></td>
    <td width="229">&nbsp;</td>
    <td width="251" colspan="2" rowspan="2" align="left" valign="top">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#CCFFFF" colspan="2"><b>
        <font face="Verdana" size="1">OFFICIAL USE ONLY</font></b></td>
      </tr>
      <tr>
        <td width="50%" bgcolor="#CCFFFF"><font face="Verdana" size="1">Date 
        Received</font></td>
        <td width="50%">&nbsp;</td>
      </tr>
      <tr>
        <td width="50%" bgcolor="#CCFFFF"><font face="Verdana" size="1">Date 
        Captured</font></td>
        <td width="50%">&nbsp;</td>
      </tr>
      <tr>
        <td width="50%" bgcolor="#CCFFFF"><font face="Verdana" size="1">
        Signature</font></td>
        <td width="50%">&nbsp;</td>
      </tr>
    </table>
       </td>
  </tr>
  <tr>
    <td width="47">&nbsp;</td>
    <td width="229">&nbsp;</td>
  </tr>
  <tr>
    <td width="47">&nbsp;</td>
    <td width="417">&nbsp;</td>
    <td width="229">&nbsp;</td>
    <td width="225">&nbsp;</td>
    <td width="26">&nbsp;</td>
  </tr>
</table>
<blockquote>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="61%" id="AutoNumber1">
  <tr>
    <td width="24%"><font face="Verdana" size="2">Name:</font></td>
    <td width="76%"><font face="Verdana" size="2"><%=objRS2("FName")%>&nbsp;<%=objRS2("SName")%></font></td>
  </tr>
  <tr>
    <td width="24%"><font face="Verdana" size="2">Student number:</font></td>
    <td width="76%"><font face="Verdana" size="2"><%=objRS2("Student_NUM")%></font></td>
  </tr>
</table>
<BR>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="600" >

<%
Response.Write("<tr bgcolor=ccffff>")


Response.Write("<th filter=ALL><font face=Verdana size=2>Assessor Number</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Standard Title</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Standard Number</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Credits</font></th>")
Response.Write("<th filter=ALL><font face=Verdana size=2>Achievement Date</font></th>")

Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Assessorid") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("Stitle") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("SNumber") & "</font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS("SCredits") & "</font></td>")
	
	Response.Write("<td><font face=Verdana size=2>" & objRS("Enrolldatey") & "/" & objRS("Enrolldatem") & "/" & objRS("Enrolldated") & "</font></td>")
	

	Response.Write("</tr>")
	objRS.MoveNext
Loop
Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td></td>")
	Response.Write("<td></td>")
	Response.Write("<td><font face=Verdana size=2><b>Total Credits</b></font></td>")
	Response.Write("<td><font face=Verdana size=2>" & objRS1("SumOfSCredits") & "</font></td>")
	
	Response.Write("<td></td>")
	

	Response.Write("</tr>")
Response.Write("</table>")
Response.Write("<blockquote>")
Response.Write("</body>")
Response.Write("</html>")

objRS.Close
objrs1.close
objrs2.close
objDB.Close
Set objRS = Nothing
Set objrs1 = nothing
Set objrs2 = nothing
Set objDB = Nothing

%>