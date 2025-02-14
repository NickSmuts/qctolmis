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
Dim sDBName
Dim sAction
Dim sRowColor
Dim html
Dim sql
Dim sError

'database variables...
Dim No_id
Dim P_title
Dim Fname
Dim Sname
Dim Id_num
Dim Age
Dim Student_num
Dim Addres
Dim Address
Dim City
Dim P_code
Dim Province
Dim Contact_num
Dim Contact_cell
Dim Training_group
Dim Sex
Dim Race
Dim Disability
Dim Marital_status
Dim Language
Dim Education
Dim Year
Dim Natqua
Dim Client
Dim Bank_name
Dim Bank_branch
Dim Bank_ibt
Dim Bank_account
Dim Project
Dim Photo

dim Genders(2)
dim j
dim Status(4)
Dim i
Dim Sir(3)
Dim k

Dim objRS1
Dim objRS2
Dim objRS5
Dim objRS4

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
sAction = Request("action")



Sub GetData()
	
	
	Set objRS = objDB.Execute("select * from DATA where Student_num = '"& request.form("StudentNum") & "'")

	If objRS.EOF Then
	No_id = ""
		P_title = ""
		Fname = ""
		Sname = ""
		Id_num = ""
		Age = ""
		Student_num = ""
		Addres = ""
		Address = ""
		City = ""
		P_code = ""
		Province = ""
		Contact_num = ""
		Contact_cell = ""
		Training_group = ""
		Sex = ""
		Race = ""
		Disability = ""
		Marital_status = ""
		Language = ""
		Education = ""
		Year = ""
		Natqua = ""
		Client = ""
		Bank_name = ""
		Bank_branch = ""
		Bank_ibt = ""
		Bank_account = ""
		Project = ""
		Photo = ""
	Else
	    No_id = objRS("No_id")
		P_title = objRS("P_title")
		Fname = objRS("Fname")
		Sname = objRS("Sname")
		Id_num = objRS("Id_num")
		Age = objRS("Age")
		Student_num = objRS("Student_num")
		Addres = objRS("Addres")
		Address = objRS("Address")
		City = objRS("City")
		P_code = objRS("P_code")
		Province = objRS("Province")
		Contact_num = objRS("Contact_num")
		Contact_cell = objRS("Contact_cell")
		Training_group = objRS("Training_group")
		Sex = objRS("Sex")
		Race = objRS("Race")
		Disability = objRS("Disability")
		Marital_status = objRS("Marital_status")
		Language = objRS("Language")
		Education = objRS("Education")
		Year = objRS("Year")
		Natqua = objRS("Natqua")
		Client = objRS("Client")
		Bank_name = objRS("Bank_name")
		Bank_branch = objRS("Bank_branch")
		Bank_ibt = objRS("Bank_ibt")
		Bank_account = objRS("Bank_account")
		Project = objRS("Project")
		Photo = objRS("Photo")
	End If
End Sub



Sub DisplayForm()
	html=""
	sRowColor="#ffffff"
	html = html & "<blockquote>"
	html = html & "<font face=Verdana>Please add new learner details then press the Update button.</font><p>"
	html = html & "<font color=red>" & sError & "</font><p>"
	html = html & "<form name=form1 method=Post action=Edit_learner1.asp>"
	html = html & "<table cellpadding=2 cellspacing=2>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Title:</font></td>"
	 	Html = html & "<td><select  name=P_title>"
	 	Sir(1)="Mr"
	 	Sir(2)="Mrs"
	 	Sir(3)="Miss"
	 	For k = 1 to 3
	 	html = html & "<option "
	 				if P_title = Sir(k) then 
	 				html = html & "selected "
	 				end if
	 	html = html & "value=" & Chr(34) & Sir(k) & Chr(34) & "> " & Sir(k) & "</option>"
		next
		Html = html & "</select></td></tr>"
	
	
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>First Name:</font></td><td><input size=35 name=Fname value=" & Chr(34) & Fname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Surname:</font></td><td><input size=35 name=Sname value=" & Chr(34) & Sname & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Id Number:</font></td><td><input size=35 name=Id_num value=" & Chr(34) & Id_num & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Age:</font></td><td><input size=20 name=Age value=" & Chr(34) & Age & Chr(34) & "></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Student Number:</font></td><td><input size=35 name=Student_num value=" & Chr(34) & Student_num & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Address:</font></td><td><input size=35 name=Addres value=" & Chr(34) & Addres & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Address:</font></td><td><input size=35 name=Address value=" & Chr(34) & Address & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>City:</font></td><td><input size=35 name=City value=" & Chr(34) & City & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Postal Code:</font></td><td><input size=35 name=P_code value=" & Chr(34) & P_code & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Province:</font></td><td><input size=35 name=Province value=" & Chr(34) & Province & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Contact Number:</font></td><td><input size=35 name=Contact_num value=" & Chr(34) & Contact_num & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Contact Cellular:</font></td><td><input size=35 name=Contact_cell value=" & Chr(34) & Contact_cell & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Training Group:</font></td><td><input size=35 name=Training_group value=" & Chr(34) & Training_group & Chr(34) & "></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Sex:</font></td>"
	
	    Html = html & "<td><select  name=Sex>"
		Genders(1) = "Male"
		Genders(2) = "Female"
		For j = 1 to 2
		html = html & "<option "
  				if Sex = Genders(j) then
   						html = html & "selected "
  				end if
  		html = html & "value=" & Chr(34) & Genders(j) & Chr(34) & "> " & Genders(j) & "</option>"
		next
		Html = html & "</select></td></tr>"
	
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Race:</font></td><td><input size=20 name=Race value=" & Chr(34) & Race & Chr(34) & "></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Disability:</font></td>"
	
			Html = html & "<td><select  name=Disability>"
			Do While Not objRS1.EOF
			html = html & "<option "
			If Disability = (objRS1("Disability")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS1("Disability") & Chr(34) & ">" &objRS1("Disability")
			objRS1.MoveNext
			Loop	   		
    		Html = html & "</select></td></tr>"

	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Marital Status:</font></td>"
	
		 Html = html & "<td><select  name=Marital_status>"
		Status(1) = "Single"
		Status(2) = "Married"
		Status(3) = "Divorced"
		Status(4) = "Widowed"
		For i = 1 to 4
		html = html & "<option "
  				if Marital_status = Status(i) then
   						html = html & "selected "
  				end if
  		html = html & "value=" & Chr(34) & Status(i) & Chr(34) & "> " & Status(i) & "</option>"
		next
		Html = html & "</select></td></tr>"	


	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Language:</font></td><td><input size=35 name=Language value=" & Chr(34) & Language & Chr(34) & "></td></tr>"
		
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Highest Education:</font></td>"
	
			Html = html & "<td><select  name=Education>"
			Do While Not objRS2.EOF
			html = html & "<option "
			If Education = (objRS2("EducationName")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS2("EducationName") & Chr(34) & ">" &objRS2("EducationName")
	
			objRS2.MoveNext
			Loop
			Html = html & "</select></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Education Year:</font></td><td><input size=35 name=Year value=" & Chr(34) & Year & Chr(34) & "></td></tr>"
	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>National Qualification Name:</font></td>"
	
	Html = html & "<td><select  name=Natqua>"
			Do While Not objRS4.EOF
			html = html & "<option "
			If Natqua = (objRS4("NQname")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS4("NQname") & Chr(34) & ">" &objRS4("NQname")
	
			objRS4.MoveNext
			Loop
				   		
    	Html = html & "</select></td></tr>"

	
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Client:</font></td><td><input size=35 name=Client value=" & Chr(34) & Client & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Bank Name:</font></td><td><input size=35 name=Bank_name value=" & Chr(34) & Bank_name & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Bank Branch Name:</font></td><td><input size=35 name=Bank_branch value=" & Chr(34) & Bank_branch & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Bank IBT Number:</font></td><td><input size=35 name=Bank_ibt value=" & Chr(34) & Bank_ibt & Chr(34) & "></td></tr>"
	html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Bank Account:</font></td><td><input size=35 name=Bank_account value=" & Chr(34) & Bank_account & Chr(34) & "></td></tr>"
	
   html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Project:</font></td>"
    
    Html = html & "<td><select  name=Project>"
			Do While Not objRS5.EOF
			html = html & "<option "
			If Project = (objRS5("Projectname")) then
   			html = html & "selected "
  			end if
			Html = html &"value=" & Chr(34) & objRS5("Projectname") & Chr(34) & ">" &objRS5("Projectname")
	
			objRS5.MoveNext
			Loop
				   		
    	Html = html & "</select></td></tr>"
html = html & "<tr bgcolor=" & sRowColor &"><td><font face=Verdana>Photo:</font></td><td><input size=35 name=Photo value=" & Chr(34) & Photo & Chr(34) & "></td></tr>"
	
	html = html & "</table><p>"
	html = html & "<input type=hidden name=No_id value =" & Chr(34) & No_id & Chr(34) & ">"
	html = html & "<input type=submit name=action value=Update>"
	html = html & "</form>"
	html = html & "</blockquote>"
	Response.Write html
End Sub

Sub ValidateData()
    No_id = Request.Form("No_id")
	P_title = Trim(Request.Form("P_title"))
	Fname = Trim(Request.Form("Fname"))
	Sname = Trim(Request.Form("Sname"))
	Id_num = Trim(Request.Form("Id_num"))
	Age = Trim(Request.Form("Age"))
	Student_num = Trim(Request.Form("Student_num"))
	Addres = Trim(Request.Form("Addres"))
	Address = Trim(Request.Form("Address"))
	City = Trim(Request.Form("City"))
	P_code = Trim(Request.Form("P_code"))
	Province = Trim(Request.Form("Province"))
	Contact_num = Trim(Request.Form("Contact_num"))
	Contact_cell = Trim(Request.Form("Contact_cell"))
	Training_group = Trim(Request.Form("Training_group"))
	Sex = Trim(Request.Form("Sex"))
	Race = Trim(Request.Form("Race"))
	Disability = Trim(Request.Form("Disability"))
	Marital_status = Trim(Request.Form("Marital_status"))
	Language = Trim(Request.Form("Language"))
	Education = Trim(Request.Form("Education"))
	Year = Trim(Request.Form("Year"))
	Natqua = Trim(Request.Form("Natqua"))
	Client = Trim(Request.Form("Client"))
	Bank_name = Trim(Request.Form("Bank_name"))
	Bank_branch = Trim(Request.Form("Bank_branch"))
	Bank_ibt = Trim(Request.Form("Bank_ibt"))
	Bank_account = Trim(Request.Form("Bank_account"))
	Project = Trim(Request.Form("Project"))
	Photo = Trim(Request.Form("Photo"))

	

	If P_title = "" Then
		sError = sError & "P_title is a required field.<br>"
	End If 

	If Fname = "" Then
		sError = sError & "Fname is a required field.<br>"
	End If 

	If Sname = "" Then
		sError = sError & "Sname is a required field.<br>"
	End If 

	If Id_num = "" Then
		sError = sError & "Id_num is a required field.<br>"
	End If 

	dim age1
	
	Age1 = datepart("yyyy",date)
	Age = Age1 - (1900 + (left(id_num,2)))
	
	

	If Student_num = "" Then
		sError = sError & "Student_num is a required field.<br>"
	End If 

	If Addres = "" Then
		sError = sError & "Addres is a required field.<br>"
	End If 

	If Address = "" Then
		sError = sError & "Address is a required field.<br>"
	End If 

	If City = "" Then
		sError = sError & "City is a required field.<br>"
	End If 

	If P_code = "" Then
		sError = sError & "P_code is a required field.<br>"
	End If 

	If Province = "" Then
		sError = sError & "Province is a required field.<br>"
	End If 

	If Contact_num = "" Then
		sError = sError & "Contact_num is a required field.<br>"
	End If 

	If Contact_cell = "" Then
		sError = sError & "Contact_cell is a required field.<br>"
	End If 

	If Training_group = "" Then
		sError = sError & "Training_group is a required field.<br>"
	End If 

	If Sex = "" Then
		sError = sError & "Sex is a required field.<br>"
	End If 

	If Race = "" Then
		sError = sError & "Race is a required field.<br>"
	End If 

	If Disability = "" Then
		sError = sError & "Disability is a required field.<br>"
	End If 

	If Marital_status = "" Then
		sError = sError & "Marital_status is a required field.<br>"
	End If 

	If Language = "" Then
		sError = sError & "Language is a required field.<br>"
	End If 

	If Education = "" Then
		sError = sError & "Education is a required field.<br>"
	End If 

	If Year = "" Then
		sError = sError & "Year is a required field.<br>"
	End If 

	If Natqua = "" Then
		sError = sError & "Natqua is a required field.<br>"
	End If 

	If Client = "" Then
		sError = sError & "Client is a required field.<br>"
	End If 

	 

	If Bank_name = "" Then
		Bank_name = "N/A"
	End If 

	If Bank_branch = "" Then
		Bank_branch = "N/A"
	End If 

	If Bank_ibt = "" Then
		Bank_ibt = "N/A"
	End If 

	If Bank_account = "" Then
		Bank_account = "N/A"
	End If  

	If Project = "" Then
		Project = "N/A"
	End If 

	If Photo = "" Then
		Photo = "noimage"
	End If 


	If sError <> "" Then
		DisplayForm()
		Response.End
	Else
	
		sql = "UPDATE DATA SET "
		sql = sql & "P_title='" & SqlQuote(P_title) & "',"
		sql = sql & "Fname='" & SqlQuote(Fname) & "',"
		sql = sql & "Sname='" & SqlQuote(Sname) & "',"
		sql = sql & "Id_num='" & SqlQuote(Id_num) & "',"
		sql = sql & "Age='" & SqlQuote(Age) & "',"
		sql = sql & "Student_num='" & SqlQuote(Student_num) & "',"
		sql = sql & "Addres='" & SqlQuote(Addres) & "',"
		sql = sql & "Address='" & SqlQuote(Address) & "',"
		sql = sql & "City='" & SqlQuote(City) & "',"
		sql = sql & "P_code='" & SqlQuote(P_code) & "',"
		sql = sql & "Province='" & SqlQuote(Province) & "',"
		sql = sql & "Contact_num='" & SqlQuote(Contact_num) & "',"
		sql = sql & "Contact_cell='" & SqlQuote(Contact_cell) & "',"
		sql = sql & "Training_group='" & SqlQuote(Training_group) & "',"
		sql = sql & "Sex='" & SqlQuote(Sex) & "',"
		sql = sql & "Race='" & SqlQuote(Race) & "',"
		sql = sql & "Disability='" & SqlQuote(Disability) & "',"
		sql = sql & "Marital_status='" & SqlQuote(Marital_status) & "',"
		sql = sql & "Language='" & SqlQuote(Language) & "',"
		sql = sql & "Education='" & SqlQuote(Education) & "',"
		sql = sql & "Year='" & SqlQuote(Year) & "',"
		sql = sql & "Natqua='" & SqlQuote(Natqua) & "',"
		sql = sql & "Client='" & SqlQuote(Client) & "',"
		sql = sql & "Bank_name='" & SqlQuote(Bank_name) & "',"
		sql = sql & "Bank_branch='" & SqlQuote(Bank_branch) & "',"
		sql = sql & "Bank_ibt='" & SqlQuote(Bank_ibt) & "',"
		sql = sql & "Bank_account='" & SqlQuote(Bank_account) & "',"
		sql = sql & "Project='" & SqlQuote(Project) & "',"
		sql = sql & "Photo='" & SqlQuote(Photo) & "'"
		sql = sql & " where NO_ID = " & NO_ID & ";"
		

		'response.write sql
		ObjDB.Execute(sql)

		If Err = 0 Then
			Response.Write "<Blockquote>"
			Response.Write "<P><font face=Verdana>Update Successful!</font></P><BR>"
			Response.Write "<p><font face=Verdana><a href=Default.asp>Main page</a></font></P>"
			Response.Write "</blockquote>"
			PageEnd()
			Response.End
		End If
	End If
End Sub

Sub PageStart()
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
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
    <td><%
    
End Sub
Sub PageEnd()
%> </td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>
<%
End Sub


Select Case sAction
	Case ""
		PageStart()
		GetData()
		DisplayForm()
		PageEnd()

	Case "Update"
	    PageStart()
		ValidateData()
        PageEnd()



End Select
%>