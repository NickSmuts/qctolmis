<%

Option Explicit



Dim sRowColor
Dim objDB
Dim objRS
Dim sDBName
Dim Fname
Dim Lname
Dim SQL
Dim C1
Dim C2
Dim C3
Dim C4
Dim C5
Dim D1
Dim D2
Dim HTMl
Dim cnpath
Dim dbname
Dim VFAX
Dim icount



dbname="data/db2.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Fname = Request.form("Fname")
Lname = Request.form("Lname")
C1 =Request.form("C1")
C2 =Request.form("C2")
C3 =Request.form("C3")
C4 =Request.form("C4")
C5 =Request.form("C5")
D1 =Request.form("D1")
D2 =Request.form("D2")

SQL= (" Select FirstName, LastName ")

If C1 = "ON" then
SQL=SQL +(", Extno ")
End if
If C2 = "ON" then
SQL=SQL +(", Cellnumber")
End if
If C3 = "ON" then
SQL=SQL +(", Email")
End if
If C4 = "ON" then
SQL=SQL +(", Faxnumber ,VFAX")
End if
If C5 = "ON" then
SQL=SQL +(", OfficeNO")
End if
If d1 = "ON" then
SQL=SQL +(", Office")
End if
If d2 = "ON" then
SQL=SQL +(", Occupation")
End if

SQL=SQL +(" from HREmployee")
If Fname <>"" then
SQL= SQL +(" where FirstName like  '" & Fname & "%' ")
Else
	If Lname <> ""then
	SQL = SQL +(" where LastName like  '" & Lname & "%' ")
	end if
End if
	If Lname <> ""then
	SQL = SQL +(" and LastName like  '" & Lname & "%' ORDER BY LastName ASC")
	end if
	
If Fname= "" and Lname ="" then
SQL = SQL +(" ORDER BY FirstName ASC")
End if

 
 'Response.write SQL
 'Response.end
 
Set objRS = objDB.Execute(SQL)

Response.Write("<html>")
Response.Write("<head>")
Response.Write("<title>Search Results</title>")
Response.Write("</head>")
Response.Write("<body bgcolor=white >")

Response.Write("<table border=0 cellpadding=0 cellspacing=0 width=800>")
Response.Write("    <tr>")
Response.Write("      <img border=0 src=images/NEW_NOSANet_Test_05.jpg><td>")
Response.Write("      &nbsp;</td>")
Response.Write("    </tr>")
Response.Write("  </table>")


If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If

HTML = ""

html = html & "<table border=0 cellpadding=0 cellspacing=0>"
html = html & "<tr bgcolor=#eeeeee>"
html = html & "<th filter=ALL><font face=Arial size=2>First Name.</font></th>"
html = html & "<th filter=ALL><font face=Arial size=2>Surname.</font></th>"

If C1 ="ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Ext. Number.</font></th>"
End if
If C2 ="ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Cell Number.</font></th>"
End if
If C3 ="ON" then
html = html & "<th filter=ALL><font face=Arial size=2>E-Mail Address.</font></th>"
End if
If C4 ="ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Fax Number. &  VFAX.</font></th>"
End if
If C5 ="ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Office Number.</font></th>"
End if
If D1 = "ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Office / Region.</font></th>"
End if
If D2 = "ON" then
html = html & "<th filter=ALL><font face=Arial size=2>Occupation. </font></th>"
End if
html = html & "</tr>"
sRowColor = "#ffffff"
Do While Not objRS.EOF

iCount = iCount + 1
	If iCount Mod 2 = 0 Then
		sRowColor = "ffffff"
	Else
		sRowColor = "#efffee"
	End If

html = html & "<tr bgcolor=" & sRowColor & ">"
html = html & "<td> <font face=Arial size=1>" & objRS("Firstname") & "</font>&nbsp;</td>"
html = html & "<td> <font face=Arial size=1>" & objRS("Lastname") & "</font>&nbsp;</td>"

If C1 = "ON" then
html = html & "<td> <font face=Arial size=1>" & objRS("Extno") & "</font>&nbsp;</td>"
End if
If C2 = "ON" then
html = html & "<td> <font face=Arial size=1>" & objRS("Cellnumber") & "</font>&nbsp;</td>"
End if
If C3 = "ON" then
html = html & "<td> <font face=Arial size=1><a href=mailto:" & objRS("Email") & ">" & objRS("Email") & "</a></font>&nbsp;</td>"
End if
If C4 = "ON"	then
html = html & "<td> <font face=Arial size=1>" & objRS("Faxnumber") & " </font>"

VFAX = objrs("VFAX")

			If VFAX <> "" Then
					html = html & "<font face=Arial size=1>  VFAX: " & objRS("VFax") & "</font>&nbsp;</td>"
			Else
 					html = html & "&nbsp;</td>"
 			end if
End if	
If C5 = "ON"	then
html = html & "<td> <font face=Arial size=1>" & objRS("OfficeNO") & "</font>&nbsp;</td>"
End if
If D1 = "ON"	then
html = html & "<td> <font face=Arial size=1>" & objRS("Office") & "</font>&nbsp;</td>"
End if
If D2 = "ON"	then
html = html & "<td> <font face=Arial size=1>" & objRS("Occupation") & "</font>&nbsp;</td>"
End if
html = html & "</tr>"

html = html & "<tr>"

html = html & "<td><hr noshade size=1 color=#008000> </td>"
html = html & "<td><hr noshade size=1 color=#008000> </td>"
If C1 = "ON" then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
If C2 = "ON" then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
If C3 = "ON" then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
If C4 = "ON"	then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if	
If C5 = "ON"	then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
If D1 = "ON"	then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
If D2 = "ON"	then
html = html & "<td><hr noshade size=1 color=#008000> </td>"
End if
html = html & "</tr>"

objRS.MoveNext
Loop
html = html & "</table>"


Response.Write html

	
	
	
	




Response.Write("</body>")
Response.Write("</html>")

objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>