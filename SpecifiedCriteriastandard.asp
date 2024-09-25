<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SCIENTIFICROOTS</title>
</head>


<%

'framework variables...
Dim objDB
Dim objRS
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError


Dim objRS1


Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


Set objRS1 = objDB.Execute("select * from [Project]")


%>


<body topmargin="0" leftmargin="2" link="#996600" vlink="#996600" alink="#996600" text="#996600" bgcolor="#FFFFFF">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="800" id="AutoNumber1">
  <tr>
    <td><!---#include file = "inc/head.asp"----></td>
  </tr>
  <tr>
    <td>
   
<form method="POST" action="SpecifiedCriteriastandard1.asp">
  <center>
    <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="700" id="AutoNumber2" height="200">
      <tr>
        <td width="100%" colspan="2" height="13">
        <p align="center"><u><font face="Verdana" size="2">Search Learn Data on 
        the following fields.</font></u></p>
        <p align="center">&nbsp;</td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Project</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D1">
         <option value="All">All</option>
        <%
        	Do While Not objRS1.EOF
			html = html & "<option "
			If Project = (objRS1("Projectname")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS1("Projectname") & Chr(34) & ">" &objRS1("Projectname")
	
			objRS1.MoveNext
			Loop
			Response.Write html
        %>
        
        </select></font></td>
      </tr>
 
      <tr>
        <td width="33%" height="19" align="right">&nbsp;</td>
        <td width="33%" height="19">&nbsp;</td>
      </tr>
      <tr>
        <td width="33%" height="26">&nbsp;</td>
        <td width="33%" height="26"><font face="Verdana"><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></font></td>
      </tr>
    </table>
    </center>
  </div>
  <p>&nbsp;</p>
</form>
<p>&nbsp;</td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>
<%
objRS1.Close
objDB.Close
Set objRS1 = Nothing
Set objDB = Nothing
%>