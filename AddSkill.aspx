<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AddSkill.aspx.cs" Inherits="AddSkill" %>

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
        <font face="Verdana">Please add the skill then press the Update button.</font>
        <p><font color="red"></font></p>
        
        <form name="form1" runat="server">
            <table cellpadding="2" cellspacing="2">
                <tbody>
                    <tr bgcolor="#ffffff">
                        <td><font face="Verdana">Skill:</font></td>
                        <td>
                            <asp:TextBox ID="txtSkill" runat="server"></asp:TextBox>
                            <asp:Label ID="lblError" ForeColor="red" runat="server" Text=""></asp:Label>
                    </tr>
                </tbody>
            </table>
            <p>
                <asp:Button ID="BtnUpdate" runat="server" Text="Update" OnClick="BtnUpdate_Click" />
                </p>
        </form>
        
        <div>
            <table>
                <tr>
                    <th style="font-size:20px;">Skills Fields
                    </th>
                </tr>
                <asp:Label ID="lblSkills" runat="server" Text=""></asp:Label>
            </table>
        </div>
    </blockquote>

</body>
</html>
