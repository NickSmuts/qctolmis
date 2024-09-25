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
Dim objRS2

Dim objRS4
Dim objRS5

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Set objRS1 = objDB.Execute("select * from [Disability]")
Set objRS2 = objDB.Execute("select * from [Education]")

Set objRS4 = objDB.Execute("select * from [Natqua]")
Set objRS5 = objDB.Execute("select * from [Project]")

%>


<body topmargin="0" leftmargin="2" link="#996600" vlink="#996600" alink="#996600" text="#996600" bgcolor="#FFFFFF">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="800" id="AutoNumber1">
  <tr>
    <td><!---#include file = "inc/head.asp"----></td>
  </tr>
  <tr>
    <td>
   
<form method="POST" action="SpecifiedCriteria1.asp">
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
        	Do While Not objRS5.EOF
			html = html & "<option "
			If Project = (objRS5("Projectname")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS5("Projectname") & Chr(34) & ">" &objRS5("Projectname")
	
			objRS5.MoveNext
			Loop
			Response.Write html
        %>
        
        </select></font></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Gender</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D2">
        <option value="None">All</option>
        <option value="Male">Male</option>
        <option value="Female">Female</option>
        </select></font></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Nat Qualification</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D3">
        <%	Do While Not objRS4.EOF
			htm = htm & "<option "
			If Natqua = (objRS4("NQname")) then
   			htm = htm & "selected "
  			end if
			Htm = htm &"value=" & Chr(34) & objRS4("NQname") & Chr(34) & ">" &objRS4("NQname")
	
			objRS4.MoveNext
			Loop
			Response.Write htm
        %>
        </select></font></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Highest Education</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D4">
         <option value="None">All</option>
        <%
        Do While Not objRS2.EOF
			ht = ht & "<option "
			If Education = (objRS2("EducationName")) then
   			ht = ht & "selected "
  			end if
			Ht = ht &"value=" & Chr(34) & objRS2("EducationName") & Chr(34) & ">" &objRS2("EducationName")
	
			objRS2.MoveNext
			Loop
			Response.Write ht
			%>
        </select></font></td>
      </tr>

      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">ID Number</font></td>
        <td width="33%" height="19"><input type="text" name="D5" size="20"></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Race</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D6">
        <option value="None">All</option>
        <option value="Black">Black</option>
        <option value="White">White</option>
        <option value="Coloured">Coloured</option>
        <option value="Indian">Indian</option>
        <option value="Asian">Asian</option>
        </select></font></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Marital Status</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D7">
        
        <option value="None">All</option>
        
        <option value=Single>Single</option>
		<option value=Married>Married</option>
		<option value=Divorced>Divorced</option>
		<option value=Widowed>Widowed</option>
        </select></font></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">disability</font></td>
        <td width="33%" height="19"><font face="Verdana"><select size="1" name="D8">
        
     <%   Do While Not objRS1.EOF
			htmls = htmls & "<option "
			If Disability = (objRS1("Disability")) then
   			htmls = htmls & "selected "
  			end if
			Htmls = htmls &"value=" & Chr(34) & objRS1("Disability") & Chr(34) & ">" &objRS1("Disability")
			objRS1.MoveNext
			Loop
			Response.Write htmls
        %>
        </select></font></td>
      </tr>
  
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Province</select></font></td>
        <td width="33%" height="19"><input type="text" name="D9" size="20"></td>
      </tr>
      <tr>
        <td width="33%" height="19" align="right"><font face="Verdana" size="2">Client</font></td>
        <td width="33%" height="19"><input type="text" name="D10" size="20"></td>
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