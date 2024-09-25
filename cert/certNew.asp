<HTML>
<HEAD>
<!--TITLE certificate TITLE-->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
</HEAD>
<BODY BGCOLOR=#FFFFFF TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 style="width: 646px; margin-left: 30px">


<!-- ImageReady Slices (certificate.psd) -->

<%


  Firstname = request.form("firstname")
  Surname = request.form("surname")
  noid = Request.form("noid")
  Project = Request.form("Project")
icount = Request.form("icount")
 
  
cc1 = request.form("C1")
cc2 = request.form("C2")
cc3 = request.form("C3")
cc4 = request.form("C4")
cc5 = request.form("C5")
cc6 = request.form("C6")
cc7 = request.form("C7")
cc8 = request.form("C8")
cc9 = request.form("C9")
cc10 = request.form("C10")
cc11 = request.form("C11")
cc12 = request.form("C12")
cc13 = request.form("C13")
cc14 = request.form("C14")
cc15 = request.form("C15")
cc16 = request.form("C16")
cc17 = request.form("C17")
cc18 = request.form("C18")
cc19 = request.form("C19")
cc20 = request.form("C20")
cc21 = request.form("C21")
cc22 = request.form("C22")
cc23 = request.form("C23")
cc24 = request.form("C24")
cc25 = request.form("C25")
cc26 = request.form("C26")
cc27 = request.form("C27")
cc28 = request.form("C28")
cc29 = request.form("C29")
cc30 = request.form("C30")
cc31 = request.form("C31")
cc32 = request.form("C32")
cc33 = request.form("C33")
cc34 = request.form("C34")
cc35 = request.form("C35")
cc36 = request.form("C36")


'if a <> "" then
'Acount = Acount + 1
'end if

'If b <> "" then
'Acount = Acount + 1
'End if

'If c <> "" then
'Acount = Acount + 1
'End if



Dim Selectedstandards, Standard
Selectedstandards = Array(cc1,cc2,cc3,cc4,cc5,cc6,cc7,cc8,cc9,cc10,cc11,cc12,cc13,cc14,cc15,cc16,cc17,cc18,cc19,cc20,cc1,cc22,cc23,cc24,cc25,cc26,cc27,cc28,cc29,cc30,cc31,cc32,cc33,cc34,cc35,cc36)


%>

<TABLE WIDTH=595 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD>
			<p align="center"><img border="0" src="images/Certify_01.jpg"></TD>
		<TD>
			&nbsp;</TD>
		<TD>
			&nbsp;</TD>
	</TR>
	<TR>
		<TD>
			<p align="center"><b><font face="Verdana"><%=Firstname%>&nbsp;<%=surname%></font></b><br>
			<p align="center"><b><font face="Verdana"><%=NOID%></font></b>			
		</TD>
		<TD>
			&nbsp;</TD>
		<TD>
			&nbsp;</TD>
	</TR>
	<TR>
		<TD>
			<p align="center"><br>
            <img border="0" src="images/Certify_03.jpg"></TD>
		<TD>
			&nbsp;</TD>
		<TD>
			&nbsp;</TD>
	</TR>
  
	<TR>
		<TD align="left" valign="top">
			<div align="center">
              <center style="width: 510px; height: 58px">
	<%
	For Each Standard In SelectedStandards
	If Standard <> "" then
		Response.Write("<b><font face=""Verdana"" Margin-left=""2"" size=""1"">" & Standard  & "; " & "</font></b>")
	End if
	Next 	
	%> 			
              </center>
           

	</TR>
	<TR>
		<TD align="left" valign="top">
			<table border="0" cellpadding="4" cellspacing="4" style="border-collapse: collapse" bordercolor="#111111" width="627" id="AutoNumber2">
              <tr>
                <td width="611">
                <p align="center">
                &nbsp;</td>
              </tr>

              <tr>
                <td width="611">
                <p align="center">
                &nbsp;</td>
              </tr>
            </table>
        </TD>
		<TD align="left" valign="top">
			&nbsp;</TD>
		<TD align="left" valign="top">
			&nbsp;</TD>
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
                <td width="84" style="height: 31px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                <td width="233" style="height: 31px"><p align="center"><b><font face="Verdana"><%=DATE%></font></b>&nbsp;</td>
                <td width="270" style="height: 31px">
                <p align="left"><b><font face="Verdana">&nbsp;&nbsp;&nbsp;PAET 2907</font></b>&nbsp;</td>
              </tr>
              <tr>
                <td width="611" colspan="3">
                <p align="center">
                <img border="0" src="images/certificate_08.jpg"></td>
              </tr>
            </table>
        </TD>
		<TD align="left" valign="top">
			&nbsp;</TD>
		<TD align="left" valign="top">
			&nbsp;</TD>
	</TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>