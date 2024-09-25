<%

Option Explicit

Dim sRowColor
Dim objDB
Dim objRS
Dim objrs1
Dim sDBName
Dim dbname
Dim Cnpath
Dim studentnum
Dim iCount

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Studentnum = request.form("studentnum")

'Set objRS = objDB.Execute("SELECT DATA.P_Title, DATA.FName, DATA.Sname,DATA.ID_Num, DATA.Student_NUM FROM DATA WHERE DATA.Student_NUM='" & studentnum &"' ")
Set objRS = objDB.Execute("select * from DATA WHERE DATA.Student_NUM='" & studentnum &"'")
Set objRS1 = objDB.Execute("SELECT DATA.P_Title, DATA.FName, DATA.Sname, DATA.Student_NUM, LearnerData.SCompetent, Standards.SNumber,Standards.Stitle, DATA.Project FROM (DATA INNER JOIN LearnerData ON DATA.Student_NUM = LearnerData.Student_NUM) INNER JOIN Standards ON LearnerData.STitle = Standards.STitle WHERE DATA.Student_NUM ='" & studentnum &"' AND SCompetent ='Competent' ")
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
Response.Write("<blockquote>")
Response.Write("<font face=Verdana>"& objRS("P_title") & "&nbsp;" & objRS("Fname") & "&nbsp;" & objRS("Sname") & "</font><br>")
Response.Write("<font face=Verdana>" & objRS("Student_num") & "</font>")
Response.Write("</blockquote>")
%>
<blockquote>




 <form method="POST" action="cert/certnew.asp">
 

    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber2">
        <tr>
          <td><%
          
          
			Do While Not objRS1.EOF
			
			iCount = iCount + 1
			
			<!-- Response.Write("<p><input type=checkbox name=C"& iCount &" value=" & Chr(34) & objRS1("Snumber") & Chr(34) & ">" & objRS1("Snumber") &" / " & objRS1("Stitle") &"</p>") -->
      Response.Write("<p><input type=checkbox name=C"& iCount &" value=" & Chr(34) & objRS1("Stitle") & Chr(34) & ">" & objRS1("Snumber") &" / " & objRS1("Stitle") &"</p>")

		
			'response.write icount
			
			objRS1.MoveNext
			Loop
			
			
			%></td>
          <td></td>
        </tr>
      
      </table>
      </center>
    </div>
     <input type="hidden" name="icount" value="<%=icount%>">
    <input type="hidden" name="Project" value="<%=objRS("Project")%>">
      <input type="hidden" name="Firstname" value="<%=objRS("Fname")%>">
  <input type="hidden" name="NOID" value="<%=objRS("ID_Num")%>">
   <input type="hidden" name="Surname" value="<%=objRS("Sname")%>"> 
 
    <input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
</blockquote>
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
objRS1.Close
objDB.Close
Set objRS = Nothing
Set objRS1 = Nothing
Set objDB = Nothing

%>