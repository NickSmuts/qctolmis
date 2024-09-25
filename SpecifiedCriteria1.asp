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



Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName
Dim SQL
Dim SQL2

Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName


D1 = request.form("D1")
D2 = request.form("D2")
D3 = request.form("D3")
D4 = request.form("D4")
D5 = request.form("D5")
D6 = request.form("D6")
D7 = request.form("D7")
D8 = request.form("D8")
D9 = request.form("D9")
D10 = request.form("D10")

C1 = request.form("C1")


'response.write D1 & "<br>"
'response.write D2 & "<br>"
'response.write D3 & "<br>" 
'response.write D4 & "<br>"
'response.write D5 & "<br>"
'response.write D6 & "<br>"
'response.write D7 & "<br>"
'response.write D8 & "<br>"
'response.write D9 & "<br>"
'response.write D10 & "<br>"

SQL = ("select * from DATA where ")


if D1 = "All" then
SQL= SQL +("No_id > 1 ")

else
SQL = SQL +("project = '" & D1 & "' ")	
Title = title + ("<font face=""Verdana"" size=""2"">Project: " & D1 & "  </font><BR>") 	
end if

If D2 <> "None" then		
SQL = SQL +("and Sex = '" & D2 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">Gender: " & D2 & "  </font><BR>") 
end if

If D3 <> "None" then		
SQL = SQL +("and NATQUA = '" & D3 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">NAT Qualification: " & D3 & "  </font><BR>") 
end if

If D4 <> "None" then		
SQL = SQL +("and Education = '" & D4 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">Highest Qualification: " & D4 & "  </font><BR>")
end if

If D5 <> "" then		
SQL = SQL +("and ID_NUM like '" & D5 & "%' ")
Title = title + ("<font face=""Verdana"" size=""2"">ID Number: " & D5 & "  </font><BR>")
end if

If D6 <> "None" then		
SQL = SQL +("and Race = '" & D6 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">Race: " & D6 & "  </font><BR>")
end if

If D7 <> "None" then		
SQL = SQL +("and Marital_Status = '" & D7 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">Marital Status: " & D7 & "  </font><BR>")
end if

If D8 <> "None" then		
SQL = SQL +("and Disability = '" & D8 & "' ")
Title = title + ("<font face=""Verdana"" size=""2"">Disability: " & D8 & "  </font><BR>")
end if

If D9 <> "" then		
SQL = SQL +("and Province LIKE '" & D9 & "%' ")
Title = title + ("<font face=""Verdana"" size=""2"">Province: " & D9 & "  </font><BR>")
end if

If D10 <> "" then		
SQL = SQL +("and Client LIKE '" & D10 & "%' ")
Title = title + ("<font face=""Verdana"" size=""2"">Client: " & D10 & "  </font><BR>")
end if



response.write title
Set objRS = objDB.Execute(SQL)


If objRS.EOF Then
Response.Write("<br>")
Response.Write("<br>")
	Response.Write("<font face=""Verdana"" size=""2""><b>No matching records found.</b></font>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If
response.write icount
Response.Write("<blockquote>")
Response.Write("<table border=0 cellpadding=4 cellspacing=0>")
Response.Write("<tr bgcolor=e9e0c4>")



'Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">No_id</font></th>")
Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">Title</font></th>")
Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">First Name</font></th>")
Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">Surname</font></th>")
Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">Id Number</font></th>")
Response.Write("<th filter=ALL><font face=""Verdana"" size=""2"">Contact Number</font></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

Do While Not objRS.EOF

	iCount = iCount + 1
	
	
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	'Response.Write("<td>" & objRS("Student_num") & "</td>")
	Response.Write("<td><font face=""Verdana"" size=""2"">" & objRS("P_title") & "</font></td>")
	Response.Write("<td><font face=""Verdana"" size=""2"">" & objRS("Fname") & "</font></td>")
	Response.Write("<td><font face=""Verdana"" size=""2""><a href=QSearch.asp?Student_num=" & objRS("Student_num") & ">" & objRS("Sname") & "</a></font></td>")
	Response.Write("<td><font face=""Verdana"" size=""2"">" & objRS("Id_num") & "</font></td>")
	Response.Write("<td><font face=""Verdana"" size=""2"">" & objRS("Contact_num") & "</font></td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</blockquote>")
response.write ("<font face=""Verdana"" size=""2"">" & icount & " records found.</font>")


objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>

 </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>