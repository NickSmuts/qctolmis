<%



Option Explicit


Function SQLQuote(var)
	If InStr(var, "'") <> 0 Then
		var = Replace(var, "'", "''")
	End If

	SQLQuote = var
End Function

'framework variables...
Dim objDB
Dim objRS
Dim objRS1
Dim objRS2
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError

'database variables...
Dim Standard
Dim Project
Dim Project1
Dim Student_num
Dim Stitle
Dim Scompetent
Dim Studentnum

Dim arrName, iCounter


Dim dbname
Dim Cnpath

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName

Studentnum = request.form("StudentNum")
Project1 = request.form("Project")







sAction = Request("action")

Sub GetData()
	
	Set objRS = objDB.Execute("select * from Project_Standard where Project ='"& request.form("Project") & "' ")

If objRS.EOF Then
Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"" >")
  
 Response.Write("<tr><td><a href=""default.asp""><img border=""0"" src=""images/main.jpg""></a></td></tr></table>")
 Response.Write("<br>")
	Response.Write("<b>No matching Project with Standards found. Please add standards to the Project</b>")
 Response.Write("<br>")
	Response.Write("<form method=POST action=Project_learner.asp><input type=hidden name=StudentNum value=" & Studentnum & "> <input type=submit value=Standards name=""B1"" ></form>")

	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If


arrName = objRS.GetRows


Student_num = studentnum
'Stitle = "test"
Scompetent = "Not_Yet"




For iCounter = 0 to UBound(arrName,2)		
		
		sql = "Insert Into LearnerData ("
		sql = sql & "Student_num,"
		sql = sql & "Stitle,"
		sql = sql & "Scompetent"
		sql = sql & ") "
		sql = sql & "Values ("
		sql = sql & "'" & SqlQuote(Student_num) & "',"
		sql = sql & "'" & arrName(2,iCounter) & "',"
		sql = sql & "'" & SqlQuote(Scompetent) & "'"
		sql = sql & ");"



		'response.write sql
		

		ObjDB.Execute(sql)
Next


		If Err = 0 Then
			'if we get here without an error, the data was updated...
				Response.redirect "Qsearchtraining.asp?StudentID="& Student_num &" "
			Response.End
		End If
	
End Sub

Sub PageStart()

End Sub





Select Case sAction
	Case ""
		PageStart()
		GetData()

End Select
%>