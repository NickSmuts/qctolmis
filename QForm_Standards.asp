<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title></title>
</head>

<body>


<%
Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Snum = request.form("Student_NUM")
Surname = request.form("Surname")
Fname = request.form("Firstname")
IDNUM = request.form("NOID")
Standards = request.form("standard")

enroll = request.form("Enroll")
ProCode = "P2SCI8782-605"
D1 = request.form("D1")
D2 = request.form("D2")	

Set objRS = objDB.Execute("SELECT Standards.SNumber, LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD, LearnerData.Student_NUM FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE (((LearnerData.Student_NUM)='"& SNUM & "') AND ((Standards.SNumber)='"& STANDARDS & "'))") 

		
%>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="950" id="AutoNumber1">
  <tr>
    <td colspan="4" width="921"><b><font face="Verdana">LEARNER STATUS FORM <br>
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
    this form when sending in Learner Achievements to the PAETA</font></td>
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
    <td width="417"><b><font face="Verdana" size="2">Student Number:<%=Snum%></font></b></td>
    <td width="229">&nbsp;</td>
    <td width="225">&nbsp;</td>
    <td width="26">&nbsp;</td>
  </tr>
</table>
<BR>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="950" id="AutoNumber3" height="420">
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner Surname</font></b></td>
  	<%
		For x = 1 To 16 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(Surname,x,1)) & "</B></font></td>"
		Next
	%>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19">&nbsp;</td>
    <td width="36" height="19" align="center"><font face="Verdana" size="2"></font></td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="47" height="19">&nbsp;</td>
    <td width="41" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="43" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td colspan="6" bgcolor="#CCFFFF" width="228" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner First Name</font></b></td>
  	<%
		For x = 1 To 16
	 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(Fname,x,1)) & "</B></font></td>"
		Next
	%>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19">&nbsp;</td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="47" height="19">&nbsp;</td>
    <td width="41" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="43" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="228" colspan="6" bgcolor="#CCFFFF" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner 2nd Name</font></b></td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="47" height="19">&nbsp;</td>
    <td width="41" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="43" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="39" height="19">&nbsp;</td>
    <td width="39" height="19">&nbsp;</td>
    <td width="37" height="19">&nbsp;</td>
    <td width="37" height="19">&nbsp;</td>
    <td width="34" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19">&nbsp;</td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="47" height="19">&nbsp;</td>
    <td width="41" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="43" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="228" colspan="6" bgcolor="#CCFFFF" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner National ID</font></b></td>
 	<%
		For x = 1 To 15 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(IDNum,x,1)) & "</B></font></td>"
		Next
	%>
    <td width="108" colspan="3" bgcolor="#CCFFFF" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Number</font></b></td>
    <td width="673" colspan="16" bgcolor="#CCFFFF" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner Alternate ID</font></b></td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="47" height="19">&nbsp;</td>
    <td width="41" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="43" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="42" height="19">&nbsp;</td>
    <td width="39" height="19">&nbsp;</td>
    <td width="39" height="19">&nbsp;</td>
    <td width="37" height="19">&nbsp;</td>
    <td width="37" height="19">&nbsp;</td>
    <td width="34" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Number</font></b></td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="493" colspan="12" bgcolor="#CCFFFF" height="19"><i>
    <font size="1" face="Verdana">This is for Learner who are not a citizen of 
    SA or do not have a National ID</font></i></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Alternative ID Type</font></b></td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="541" colspan="13" bgcolor="#CCFFCC" height="19"><i>
    <font size="1" face="Verdana">A Unique identifier for an alternate id number 
    – see the SAQA code guide</font></i></td>
  </tr>
  
  <tr>
    <td bgcolor="#CCFFFF" width="409" colspan="4" height="19"><b><font face="Verdana" size="2">&nbsp;Learner Achievement 
    Status ID</font></b></td>
   <%
		For x = 1 To 2 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(D1,x,1)) & "</B></font></td>"
		Next
	%>
    <td width="447" colspan="11" bgcolor="#CCFFCC" height="19"><i>
    <font size="1" face="Verdana">See the SAQA code guide</font></i></td>
  </tr>
  
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="20"><b><font face="Verdana" size="2">&nbsp;Learner Achievement 
    Type ID</font></b></td>
    <%
		For x = 1 To 2 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(D2,x,1)) & "</B></font></td>"
		Next
	%>
    <td width="192" height="20" colspan="4" bgcolor="#CCFFFF">&nbsp;</td>
    <td width="400" height="20" colspan="10" bgcolor="#CCFFFF"><b>
    <font face="Verdana" size="2"></font></b></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19">&nbsp;</td>
    <td width="273" height="19" colspan="6" bgcolor="#CCFFCC"><i>
    <font face="Verdana" size="1">See the SAQA code guide</font></i></td>
  
  <%
		For x = 1 To 8
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(Enrolld,x,1)) & "</B></font></td>"
		Next
	%>

  
  
    <td width="71" height="19" bgcolor="#CCFFFF" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Learner Enrolled Date</font></b></td>
    <%
		For x = 1 To 8 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(Enroll,x,1)) & "</B></font></td>"
		Next
	%>

    <td width="317" height="19" bgcolor="#CCFFFF" colspan="8">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19"><b><font face="Verdana" size="2">&nbsp;Provider Code</font></b></td>
  	<%
		For x = 1 To 16 
			Response.Write "<td align=center><font face=Verdana size=2><B>" & UCase(Mid(ProCode,x,1)) & "</B></font></td>"
		Next
	%>
  </tr>
  <tr>
    <td bgcolor="#CCFFFF" width="277" height="19">&nbsp;</td>
    <td width="36" height="19">&nbsp;</td>
    <td width="45" height="19">&nbsp;</td>
    <td width="51" height="19">&nbsp;</td>
    <td width="48" height="19">&nbsp;</td>
    <td width="46" height="19">&nbsp;</td>
    <td width="447" height="19" bgcolor="#CCFFFF" colspan="11">&nbsp;</td>
  </tr>
</table>

</body>

</html>