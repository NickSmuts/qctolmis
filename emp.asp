<%@ Language = VBscript %>
<% Option Explicit %>

<% Response.Buffer = True %>

<%
'***************************************
' This is downloaded from www.plus2net.com //
' You can distribute this code with the link to www.plus2net.com ///
'  Please don't  remove the link to www.plus2net.com ///
' This is for your learning only not for commercial use. ///////
' The author is not responsible for any type of loss or problem or damage on using this script.//
' You can use it at your own risk. /////
' *****************************************


%>
<html>
<head>
<title>Database Search</title>
<SCRIPT language=JavaScript>
function reload(form){
var val=form.dept.options[form.dept.options.selectedIndex].value;
self.location='emp.asp?dept=' + val ;
}
</script>
</head><body>

<%
Dim objconn,objRS,strSQL,dept
dept=Request.QueryString("dept")
Set objconn = Server.CreateObject("ADODB.Connection")
objconn.ConnectionString = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("emp.mdb")
objconn.Open

Set objRs = Server.CreateObject("ADODB.Recordset")

'''''First drop down list starts here'''''

strSQL = "SELECT distinct dept from emp_m"
objRS.Open strSQL, objconn
Response.Write "<form method=post name=f1 action=''><select name=dept onchange='reload(this.form)'><option value=''>Select dept</option>"
Do While Not objRS.EOF 
	Response.Write "<option value=" & objRs("dept") & ">" & objRs("dept") & "</option>"
     objRS.MoveNext
 Loop
objRs.Close
Response.Write "</select>"
Response.Write "<br>----<br>"

''' Second drop down list starts here ''''

If len(dept) > 1 Then

strSQL = "SELECT  * FROM emp_m where dept='" & dept &"'"
objRS.Open strSQL, objconn

Do While Not objRS.EOF 
	Response.Write objRs("emp_no") & " " & objRs("name") & "  "  & objRs("dept") & "<br>"
     objRS.MoveNext
 Loop
Response.Write "</form>"
objRs.Close

objconn.Close
end if 

     %>
</body>
</html>

