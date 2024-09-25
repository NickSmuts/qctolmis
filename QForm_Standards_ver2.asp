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

iniFname = Left (Fname, 1)
iniSurname = Left (Surname, 1)

INITIALS = iniFname & iniSurname

IDNUM = request.form("NOID")



StNUM = request.form("Studnet_NUM")

RACE = request.form("Race")

Standards = request.form("standard")

SEX = request.form("Sex")

Language  = request.form("Language")
Addres  = request.form("Addres")
Address  = request.form("Address")
Contact_NUM  = request.form("Contact_NUM")
Company = request.form("Client")
Course_attending = request.form("Project")
Standard_NUM = request.form("Standard_NUM")
SCompetent = request.form("SCompetent") 
AssessorID = request.form("AssessorID")
OFOCode = request.form("OFOCode")
OFODesc = request.form("OFODesc")

CompanyName = request.form("CompanyName")
TrainingManager = request.form("TrainingManager")
CNumber = request.form("CNumber")
SICCode = request.form("SICCode")
SSUCode = request.form("SSUCode")

  





AssessorTname = request.form("AssessorTname")
AssessorTSname = request.form("AssessorTSname")

enroll = request.form("Enroll")
ProCode = "P2SCI8782-605"
D1 = request.form("D1")
D2 = request.form("D2")	

Set objRS = objDB.Execute("SELECT Standards.SNumber, LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD, LearnerData.Student_NUM FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE (((LearnerData.Student_NUM)='"& SNUM & "') AND ((Standards.SNumber)='"& STANDARDS & "'))") 


Set objRS1 = objDB.Execute("SELECT LearnerData.EnrolldateY, LearnerData.EnrolldateM, LearnerData.EnrolldateD, LearnerData.STitle, Standards.SNumber, LearnerData.Student_NUM, LearnerData.SCompetent, LearnerData.AssessorID FROM LearnerData INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE (((LearnerData.Student_NUM)='"& request.form("Student_Num") & "'))")

		
%>
<table width="950" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="687">&nbsp;</td>
    <td width="263" rowspan="2"><img src="images/logo_agriseta.jpg" width="257" height="193"  alt=""/></td>
  </tr>
  <tr>
    <td valign="bottom"><p><strong>PROVIDER NAME</strong><br>
      <strong>LEARNER IFORMATION FORM</strong></p></td>
  </tr>
</table>

<p>
  <!--                UPDATING FROM HERE               -->
  
  
  
</p>
<table width="950" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td colspan="14"><strong>LEARNER DETAILS</strong></td>
  </tr>
  <tr>
    <td width="258">Learner  Surname</td>
    	<%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Surname,x,1)) & "</B></font></td>"
		Next
	%>
  </tr>
  <tr>
    <td>Learner  Name</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Fname,x,1)) & "</B></font></td>"
		Next
	%>

  </tr>
  <tr>
    <td>Initials</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(INITIALS,x,1)) & "</B></font></td>"
		Next
	%>
 
  </tr>
  <tr>
    <td>Learner  South African Id number</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(IDNUM,x,1)) & "</B></font></td>"
		Next
	%>

  </tr>
  <tr>
    <td>Race</td>
       <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(RACE,x,1)) & "</B></font></td>"
		Next
	%>
   
  </tr>
  <tr>
    <td>Gender</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(SEX,x,1)) & "</B></font></td>"
		Next
	%>	

  </tr>
  <tr>
    <td>Home Language</td>
      <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Language,x,1)) & "</B></font></td>"
		Next
	%>	
  </tr>
  <tr>
    <td>Learner  Home Address</td>
      <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Addres,x,1)) & "</B></font></td>"
		Next
	%>	
  </tr>
  <tr>
    <td>Learner  Postal Address</td>
      <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Address,x,1)) & "</B></font></td>"
		Next
	%>	
  </tr>
  <tr>
    <td>Learner  Phone number</td>
      <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(Contact_NUM,x,1)) & "</B></font></td>"
		Next
	%>	
  </tr>
  <tr>
    <td>OFO  Code</td>
<%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(OFOCode,x,1)) & "</B></font></td>"
    Next
  %>  
  </tr>
  <tr>
    <td>OFO  Description </td>
    <%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(OFODesc,x,1)) & "</B></font></td>"
    Next
  %> 
  </tr>
  <tr>
    <td colspan="14"><strong>COMPANY DETAILS</strong></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Name  of Company</td>
<td colspan="13"><font face=Verdana size=2><B><% Response.Write UCase(CompanyName) %></B></font></td>	
   
  </tr>
  <tr>
    <td rowspan="3" valign="top">Postal  Address</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Training  Manager</td>
    <%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(TrainingManager,x,1)) & "</B></font></td>"
    Next
  %>
  </tr>
  <tr>
    <td>Contact  Number/s</td>
    <%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(CNumber,x,1)) & "</B></font></td>"
    Next
  %>
  </tr>
  <tr>
    <td>SIC  Code</td>
    <%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(SICCode,x,1)) & "</B></font></td>"
    Next
  %>
  </tr>
  <tr>
    <td>SSU  Code</td>
    <%
    For x = 1 To 13 
      Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(SSUCode,x,1)) & "</B></font></td>"
    Next
  %>
  </tr>
  <tr>
    <td colspan="14"><strong>COURSE INFORMATION</strong></td>
  </tr>
  <tr>
    <td>NSDS  Target (Employed / Unemployed)</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Name  of Course attending</td>
    
<td colspan="13"><font face=Verdana size=2><B><% Response.Write UCase(Course_attending) %></B></font></td>	
  
  </tr>
  <tr>
    <td valign="top">Unit Standard Number</td>
  <td colspan="13"><font face=Verdana size=2><B><%
Do While Not objRS1.EOF
Response.Write objRS1("SNumber") & "  "
objRS1.MoveNext
Loop
%></B></font></td>
  </tr>
  <tr>
    <td>Date  of Course</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Competent  / Not Competent</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(SCompetent,x,1)) & "</B></font></td>"
		Next
	%>
  </tr>
  <tr>
    <td>Name  of Assessor</td>
    <%
		For x = 1 To 13 
			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(AssessorTname & " " &AssessorTSname,x,1)) & "</B></font></td>"
		Next
	%>
    
  </tr>
  <tr>
    <td>Assessor  number</td> 
     <%
		'For x = 1 To 1
'			Response.Write "<td width='40' align=center><font face=Verdana size=2><B>" & UCase(Mid(AssessorID,x,1)) & "</B></font></td>"
'		Next

Response.Write "<td colspan='7'><font face=Verdana size=2><B>" & (AssessorID) & "</B></font></td>"
	%>
    <td>Date</td>
    <td colspan="5">&nbsp;</td>
  </tr>
  <tr>
    <td>Signature  of Assessor</td>
    <td colspan="7">&nbsp;</td>
    <td>Date</td>
    <td colspan="5">&nbsp;</td>
  </tr>
  <tr>
    <td>Signature  of Learner</td>
    <td colspan="13">&nbsp;</td>
  </tr>
</table>

</body>

</html>