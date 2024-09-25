<%@ Page Language="C#" AutoEventWireup="true" CodeFile="LearnerCV.aspx.cs" Inherits="LearnerCV" %>

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
        
        .cvdots {
            text-align: right;
            display: inline;
        }

        .cvlabel {
            text-align: left;
        }

        table.blueTable {
            /*width: 70%;*/
            text-align: left;
            margin-left: 60px;
        }

            /*table.blueTable {
            text-align: left;
            left:20px;
        }*/

            table.blueTable td, table.blueTable th {
                padding: 4px 4px;
            }

            table.blueTable tbody td {
                font-size: 16px;
                color: black;
            }

            table.blueTable thead {
                color: black;
                font-size: 17px;
            }

                table.blueTable thead th {
                    font-size: 17px;
                    font-weight: bold;
                    border-left: 0px solid #D0E4F5;
                }

                    table.blueTable thead th:first-child {
                        border-left: none;
                    }

            table.blueTable tfoot {
                font-weight: bold;
            }

        @media print {
            .hideprint {
                visibility: hidden;
            }

            @page {
                margin: 0;
            }

            body {
                margin: .1cm;
            }

            html, body {
                height: 99%;
                page-break-after: avoid;
                page-break-before: avoid;
            }
            #learnerPanel
            {
                margin-top:-150px;
            }
            .footer {
           position: fixed;
           left: 0;
           bottom: 0;
           width: 100%;
           text-align: left;
            }
        }
    </style>
</head>
<body runat="server" topmargin="0" leftmargin="2" link="#996600" vlink="#996600" alink="#996600" bgcolor="#FFFFFF">
    <table border="0" class="hideprint" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
        <tbody>
            <tr>
                <td><a href="default.asp">
                    <img border="0" src="images/main.jpg"></a></td>
                <td>
                    <a href="#" class="hideprint">
                        <img alt="" width="50" style="margin-left: 100%;" src="images/34441.png" onclick="printMe()" />
                    </a>

                </td>
            </tr>
        </tbody>
    </table>
    <form name="form1" runat="server">
        <asp:Panel ID="learnerPanel" Visible="true" runat="server">

            <table class="blueTable">
                <thead>
                    <td colspan="2">
                        <b>Personal Details
                                <hr />
                        </b>
                    </td>

                </thead>
                <tr>
                    <td>Surname 
                        <p style="padding-left: 159px; padding-right: 50px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblSurname" runat="server" CssClass="cvlabel" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Name
                        <p style="padding-left: 178px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblName" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Identity Number
                        <p style="padding-left: 110px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblId" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Gender
                        <p style="padding-left: 169px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblGender" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Marital Status
                        <p style="padding-left: 125px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblmaritalStatus" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Disability
                        <p style="padding-left: 153px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblDisability" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Contact No
                        <p style="padding-left: 143px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblContactNum" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <br />
                <thead>

                    <td colspan="2">
                        <b>Education & Qualifications
                                <hr />
                        </b>
                    </td>
                </thead>
                <tr>
                    <td>Highest Education
                        <p style="padding-left: 94px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblHighestEducation" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>Year
                        <p style="padding-left: 183px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblYear" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td><b>Quaifications</b>
                    </td>
                    <td>
                        <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                    </td>
                </tr>

            </table>
            <table class="blueTable">
                <tr>
                    <td></td>
                    <td>Institute<p style="padding-left: 154px; padding-right: 50px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblInstitute" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td>Qualifications<p style="padding-left: 115px; display: inline;">:</p>
                    </td>
                    <td>
                        <asp:Label ID="lblQualifications" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
            </table>
            <table class="blueTable">
                <tr>
                    <td><b>Skills</b>
                    </td>
                    <td></td>
                </tr>
                <tr>
                    <td></td>
                    <td>
                        <asp:Label ID="lblSkills" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
            </table>
            <br />

            <footer class="footer">
                <table>
                    <tr>
                        <td>Client Name</td>
                    </tr>
                    <tr>
                        <td>CompanyName</td>
                    </tr>
                    <tr>
                        <td>EmailAddress</td>
                    </tr>
                    <tr>
                        <td>ContactNum</td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblDate" runat="server" Text=""></asp:Label></td>
                    </tr>
                </table>
            </footer>
        </asp:Panel>
    </form>


</body>
</html>
<script>
    function printMe() {
        window.print();
    }
</script>
