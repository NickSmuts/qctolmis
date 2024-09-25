<%@ Page Language="C#" AutoEventWireup="true" CodeFile="learnerCVSkills.aspx.cs" Inherits="learnerCVSkills" %>

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
        <u><b><font face="Verdana">Learner CV Skills</font></b></u>
        <p><font color="red"></font></p>

        <form name="form1" runat="server">
            <asp:Panel ID="learnerPanel" Visible="false" runat="server">
                <div>
                    <table>
                        <tr>
                            <td>
                                <asp:Label ID="lblSName" runat="server" Text=""></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="lblSNum" runat="server" Text=""></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <th style="font-size: 20px;text-align:left;">Skills Selection
                            </th>
                        </tr>
                        <tr>
                            <td>
                                Skill : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:DropDownList ID="DDSkills" runat="server">

                                </asp:DropDownList>
                            </td>
                           <%-- <td>
                                
                            </td>--%>
                            </tr>
                        <tr>
                            <td>
                                Experiance : &nbsp;<asp:TextBox ID="txtSkillYears" runat="server"></asp:TextBox>
                                <asp:Button ID="BtnAddSkill" runat="server" Width="70px" Text="Add" OnClick="BtnAddSkill_Click" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
                                <asp:ListBox ID="skillListBox" Width="170px" runat="server" AutoPostBack="true" OnSelectedIndexChanged="skillListBox_SelectedIndexChanged"></asp:ListBox>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
                                <asp:Button ID="BtnDeleteSkill" runat="server" OnClick="BtnDeleteSkill_Click" Text="Delete" />
                            </td>
                        </tr>
                        <%--<asp:CheckBoxList ID="checkBoxSkills" RepeatDirection="Horizontal" AutoPostBack="true" OnSelectedIndexChanged="checkBoxSkills_SelectedIndexChanged" RepeatColumns="6" CellSpacing="2" CellPadding="12" runat="server"></asp:CheckBoxList>
                    <asp:Label ID="lblSkills" runat="server" Text=""></asp:Label>--%>
                    </table>
                    <br />
                    <asp:Button ID="BtnSubmit" runat="server" Text="Submit" OnClick="BtnSubmit_Click"/>
                            <asp:Button ID="BtnReset" runat="server" Text="Reset" OnClick="BtnReset_Click" />
                </div>
            </asp:Panel>
        </form>
    </blockquote>

</body>
</html>
