<HTML>
<HEAD>
<!--TITLE certificate TITLE-->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>


<!-- ImageReady Slices (certificate.psd) -->

<%


  Firstname = request.form("firstname")
  Surname = request.form("surname")
  noid = Request.form("noid")
  Project = Request.form("Project")
icount = Request.form("icount")
 
  
a = request.form("C1")
b = request.form("C2")
c = request.form("C3")
d = request.form("C4")
e = request.form("C5")
f = request.form("C6")
g = request.form("C7")
h = request.form("C8")
i = request.form("C9")
j = request.form("C10")
k = request.form("C11")
l = request.form("C12")
m = request.form("C13")
n = request.form("C14")
o = request.form("C15")
p = request.form("C16")
q = request.form("C17")
r = request.form("C18")

a123 = request.form("C19")
b123 = request.form("C20")
c123 = request.form("C21")
d123 = request.form("C22")
e123 = request.form("C23")
f123 = request.form("C24")
g123 = request.form("C25")
h123 = request.form("C26")
i123 = request.form("C27")
j123 = request.form("C28")
k123 = request.form("C29")
l123 = request.form("C30")
m123 = request.form("C31")
n123 = request.form("C32")
o123 = request.form("C33")
p123 = request.form("C34")
q123 = request.form("C35")
r123 = request.form("C36")


if a <> "" then
Acount = Acount + 1
end if

If b <> "" then
Acount = Acount + 1
End if

If c <> "" then
Acount = Acount + 1
End if




'Response.write icount






%>

<TABLE WIDTH=595 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD>
			<p align="center"><img border="0" src="images/Certify_01.jpg"></TD>
	</TR>
	<TR>
		<TD>
			<p align="center"><b><font face="Verdana"><%=Firstname%>&nbsp;<%=surname%></font></b><br>
			<p align="center"><b><font face="Verdana"><%=NOID%></font></b>
			
			</TD>
	</TR>
	<TR>
		<TD>
			<p align="center"><br>
            <img border="0" src="images/Certify_03.jpg"></TD>
	</TR>
  
	<TR>
		<TD align="left" valign="top">
			<div align="center">
              <center>
<%
              
If acount = 1 then

Response.Write("<table border=""0"" cellpadding=""2"" cellspacing=""2"" style=""border-collapse: collapse"" bordercolor=""#111111"" width=""100%"" id=""AutoNumber1"">")
Response.Write("  <tr>")
Response.Write("    <td width=""33%"" align=""center""><b><font face=""Verdana"" size=""1""></font></b>&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center""><b><font face=""Verdana"" size=""1"">" & a & "</font></b></td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("    <td width=""33%"" align=""center"">&nbsp;</td>")
Response.Write("  </tr>")
Response.Write("</table>")

else

Response.Write("<table border=""0"" cellpadding=""2"" cellspacing=""2"" style=""border-collapse: collapse"" bordercolor=""#111111"" width=""400"" id=""AutoNumber2"">")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1""></font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1""></font></b></td>")
Response.Write("  </tr>")

Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & a & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & b & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & c & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & d & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & e & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & f & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & g & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & h & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & i & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & j & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & k & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & l & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & m & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & n & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & o & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & p & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & q & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & r & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & a123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & b123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & c123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & d123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & e123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & f123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & g123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & h123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & i123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & j123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & k123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & l123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & m123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & n123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & o123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & p123 & "</font></b></td>")
Response.Write("  </tr>")
Response.Write("  <tr>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & q123 & "</font></b></td>")
Response.Write("    <td width=""50%"" align=""center""><b><font face=""Verdana"" size=""1"">" & r123 & "</font></b></td>")
Response.Write("  </tr>") 
Response.Write("</table>")
end if
%>             
			
              </center>
            </div>
        </TD>
	</TR>
	<TR>
		<TD align="left" valign="top">
			<table border="0" cellpadding="4" cellspacing="4" style="border-collapse: collapse" bordercolor="#111111" width="627" id="AutoNumber2">
              <tr>
                <td width="611" colspan="3">
                <p align="center">
                <img border="0" src="images/certificateNew_06.jpg"></td>
              </tr>
              <tr>
                <td width="84">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                <td width="233"><p align="center"><b><font face="Verdana"><%=DATE%></font></b>&nbsp;</td>
                <td width="270">
                <p align="left"><b><font face="Verdana">&nbsp;&nbsp;&nbsp;PAET 2907</font></b>&nbsp;</td>
              </tr>
              <tr>
                <td width="611" colspan="3">
                <p align="center">
                <img border="0" src="images/certificate_08.jpg"></td>
              </tr>
            </table>
        </TD>
	</TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>