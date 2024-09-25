<%@ Page Language="C#" AutoEventWireup="true" CodeFile="LearnerCVSearch.aspx.cs" Inherits="LearnerCV" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Language" content="en-za">
    <meta name="GENERATOR" content="Microsoft FrontPage 12.0">
    <meta name="ProgId" content="FrontPage.Editor.Document">
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <title>SCIENTIFICROOTS</title>

    <style type="text/css">
        .auto-style1 {
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            font-size: x-small;
        }
    </style>
</head>
<body runat="server" topmargin="0" leftmargin="2" link="#996600" vlink="#996600" alink="#996600" text="#996600" bgcolor="#FFFFFF">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
        <tbody>
            <tr>
                <td><a href="default.asp">
                    <img border="0" src="images/main.jpg"></a></td>
            </tr>
        </tbody>
    </table>
    <blockquote>
        <u><b><font face="Verdana">Learner CV</font></b></u>
        <p><font color="red"></font></p>

        <form name="form1" runat="server">
            <table cellpadding="2" cellspacing="2">
                <tbody>
                    <tr bgcolor="#ffffff">
                        <td><font face="Verdana">Student Number:</font></td>
                        <td>
                            <asp:TextBox ID="txtStudentNumber" runat="server"></asp:TextBox>
                            <asp:Label ID="lblError" ForeColor="red" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>
                            <asp:Button ID="BtnSubmit" runat="server" Text="Submit" OnClick="BtnSubmit_Click" />
                            <asp:Button ID="BtnReset" runat="server" Text="Reset" OnClick="BtnReset_Click" />
                        </td>
                    </tr>
                </tbody>
            </table>
            <p>
            </p>

            
        </form>
    </blockquote>

</body>
</html>
