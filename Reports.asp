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
    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber2">
        <tr>
          <td width="116"><u><b><font face="Verdana" size="2">REPORTS</font></b></u></td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116"><font face="Verdana">[1]</font></td>
          <td width="484"><font face="Verdana" color="#009933">
          <%

Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName
Dim Html
Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Set objRS = objDB.Execute("select * from Project")

If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If
Response.Write("<form method=POST action=report_progress.asp>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=ffffff>")
Response.Write("<th filter=ALL>Learner Count by Project</th>")
Response.Write("<th filter=ALL></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

    Html = html & "<td><select  name=Project>"
			Do While Not objRS.EOF
			html = html & "<option "
			
			Html = html &"value=" & Chr(34) & objRS("Projectname") & Chr(34) & ">" &objRS("Projectname")
	
			objRS.MoveNext
			Loop
				   		
    	Html = html & "</select></td>"
    	Response.Write html

Response.Write("<TD> <input type=submit value=Report name=B1></TD>")    	
Response.Write("</tr>")
Response.Write("</table>")
Response.Write("</form>")
objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>
</font></td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
      </table>
      </center>
    </div>
    <p>&nbsp;</p>
    <p>&nbsp;</td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>