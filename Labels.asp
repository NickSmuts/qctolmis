<%

Option Explicit


Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Set objRS = objDB.Execute("select * from DATA")

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

Do While Not objRS.EOF
Response.Write("<Blockquote>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr>")
Response.Write("   <td><font size=3 face=Verdana>" & objRS("P_title") & "&nbsp;" & objRS("Fname") & "&nbsp;" & objRS("Sname") & "</font><br>")
Response.Write("   <font size=3 face=Verdana>" & objRS("Addres") & "<BR>")
If objRS("Address") = "N/A" then
Response.Write("")
else 
Response.Write(" " & objRS("Address") & "<BR>")
end if


Response.Write(" " & objRS("P_code") & "<BR>")
Response.Write(" " & objRS("City") & "<BR>")
Response.Write(" " & objRS("Province") & "<BR>")
Response.Write("</font></td>")
Response.Write("  </tr>")

	

Response.Write("</table>")
Response.Write("</Blockquote>")
Response.Write("<Br>")
objRS.MoveNext
Loop
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