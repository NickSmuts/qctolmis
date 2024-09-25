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

Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=silver>")



Response.Write("<th filter=ALL>Projectname</th>")

Response.Write("</tr>")

sRowColor = "ccffff"



    Html = html & "<td><select  name=Project>"
			Do While Not objRS.EOF
			html = html & "<option "
			
			Html = html &"value=" & Chr(34) & objRS("Projectname") & Chr(34) & ">" &objRS("Projectname")
	
			objRS.MoveNext
			Loop
				   		
    	Html = html & "</select></td></tr>"
    	Response.Write html

Response.Write("</table>")
objRS.Close
objDB.Close
Set objRS = Nothing
Set objDB = Nothing

%>