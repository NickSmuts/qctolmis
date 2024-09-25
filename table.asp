
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

Response.Write("<html>")
Response.Write("<head>")
Response.Write("<title>ASP Table Wizard</title>")
Response.Write("</head>")
Response.Write("<body bgcolor=white>")

Response.Write("<h3>ASP Table Wizard</h3>")

If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If

Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=silver>")

'COOL TIP: the <filter> tag is used by Excel 97 and later...
'if your users save this file from the browser and open it in XL, XL will
'parse all the table cells into XL ranges and turn on filtering...

Response.Write("<th filter=ALL>No_id</th>")
Response.Write("<th filter=ALL>Fname</th>")
Response.Write("<th filter=ALL>Sname</th>")
Response.Write("<th filter=ALL>Id_num</th>")
Response.Write("<th filter=ALL>Student_num</th>")
Response.Write("<th filter=ALL>Address</th>")
Response.Write("<th filter=ALL>P_code</th>")
Response.Write("<th filter=ALL>Contact_num</th>")
Response.Write("<th filter=ALL>Contact_cell</th>")
Response.Write("<th filter=ALL>Training_group</th>")
Response.Write("<th filter=ALL>Sex</th>")
Response.Write("<th filter=ALL>Disability</th>")
Response.Write("<th filter=ALL>B_date</th>")
Response.Write("<th filter=ALL>Marital_status</th>")
Response.Write("<th filter=ALL>Language</th>")
Response.Write("<th filter=ALL>Education</th>")
Response.Write("<th filter=ALL>Year</th>")
Response.Write("<th filter=ALL>Unit_name</th>")
Response.Write("<th filter=ALL>Unit_num</th>")
Response.Write("<th filter=ALL>Client</th>")
Response.Write("<th filter=ALL>Credits</th>")
Response.Write("<th filter=ALL>Bank_name</th>")
Response.Write("<th filter=ALL>Bank_branch</th>")
Response.Write("<th filter=ALL>Bank_ibt</th>")
Response.Write("<th filter=ALL>Bank_account</th>")
Response.Write("<th filter=ALL>N_qualification_award</th>")
Response.Write("<th filter=ALL>N_qualification_name</th>")
Response.Write("<th filter=ALL>Project</th>")
Response.Write("<th filter=ALL>Photo</th>")
Response.Write("</tr>")

sRowColor = "lightblue"

Do While Not objRS.EOF
	Response.Write("<tr bgcolor=" & sRowColor & ">")
	Response.Write("<td>" & objRS("No_id") & "</td>")
	Response.Write("<td>" & objRS("Fname") & "</td>")
	Response.Write("<td>" & objRS("Sname") & "</td>")
	Response.Write("<td>" & objRS("Id_num") & "</td>")
	Response.Write("<td>" & objRS("Student_num") & "</td>")
	Response.Write("<td>" & objRS("Address") & "</td>")
	Response.Write("<td>" & objRS("P_code") & "</td>")
	Response.Write("<td>" & objRS("Contact_num") & "</td>")
	Response.Write("<td>" & objRS("Contact_cell") & "</td>")
	Response.Write("<td>" & objRS("Training_group") & "</td>")
	Response.Write("<td>" & objRS("Sex") & "</td>")
	Response.Write("<td>" & objRS("Disability") & "</td>")
	Response.Write("<td>" & objRS("B_date") & "</td>")
	Response.Write("<td>" & objRS("Marital_status") & "</td>")
	Response.Write("<td>" & objRS("Language") & "</td>")
	Response.Write("<td>" & objRS("Education") & "</td>")
	Response.Write("<td>" & objRS("Year") & "</td>")
	Response.Write("<td>" & objRS("Unit_name") & "</td>")
	Response.Write("<td>" & objRS("Unit_num") & "</td>")
	Response.Write("<td>" & objRS("Client") & "</td>")
	Response.Write("<td>" & objRS("Credits") & "</td>")
	Response.Write("<td>" & objRS("Bank_name") & "</td>")
	Response.Write("<td>" & objRS("Bank_branch") & "</td>")
	Response.Write("<td>" & objRS("Bank_ibt") & "</td>")
	Response.Write("<td>" & objRS("Bank_account") & "</td>")
	Response.Write("<td>" & objRS("N_qualification_award") & "</td>")
	Response.Write("<td>" & objRS("N_qualification_name") & "</td>")
	Response.Write("<td>" & objRS("Project") & "</td>")
	Response.Write("<td>" & objRS("Photo") & "</td>")
	Response.Write("</tr>")
	objRS.MoveNext
Loop

Response.Write("</table>")
Response.Write("</body>")
Response.Write("</html>")

objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>