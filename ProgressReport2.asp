<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
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
    <td>
        <blockquote>
    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber2">
        <tr>
          <td width="116"><u><b><font face="Verdana" size="2">REPORTS</font></b></u></td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116"><font face="Verdana">[1]</font></td>
          <td width="484"><font face="Verdana" color="#009933">
          <%

Dim sRowColor
Dim objDB
Dim objRS
Dim objRS2
Dim sDBName
Dim Html
Dim dbname
Dim Cnpath
Dim sAction

dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName
             
Set objRS = objDB.Execute("select n.NQName +',' + d.project as NQName from NatQua as n left join DATA as d on d.NATQUA=n.NQName where not n.NQName='none' and not d.project='' group by n.NQName,d.project ")

If objRS.EOF Then
	Response.Write("<b>No matching records found.</b>")
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	Response.End
End If

 sAction = Request("action")
Response.Write("<form method=POST>")
Response.Write("<table border=0 cellpadding=2 cellspacing=2>")
Response.Write("<tr bgcolor=ffffff>")
Response.Write("<th filter=ALL>Qualification Progress Report</th>")
Response.Write("<th filter=ALL></th>")
Response.Write("</tr>")

sRowColor = "ffffff"

    Html = html & "<td><select  name=qualification>"
			Do While Not objRS.EOF
			html = html & "<option "
			
			Html = html &"value=" & Chr(34) & objRS("NQName") & Chr(34) & ">" &objRS("NQName")
	
			objRS.MoveNext
			Loop
				   		
    	Html = html & "</select></td>"
    	Response.Write html

Response.Write("<TD> <input type=submit name=action value=Report></TD>")    	
Response.Write("</tr>")
Response.Write("</table>")
Response.Write("</form>")
objRS.Close
objDB.Close



Set objRS = Nothing
Set objDB = Nothing
              Select Case sAction
	            Case "Report"
              set objExcel = CreateObject("Excel.Application")
              'Create a new workbook.
                set objWorkbook= objExcel.Workbooks.Add
                 objWorkbook.Sheets.Add
                'Select the first sheet
                Sheet = 1
              Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(Sheet)
               objWorkSheet.Name = "Trainers"
              
                objWorkSheet.cells(1,1).value = "Koue Bokkeveld Opleidingsentrum BK"
                objWorkSheet.cells(2,1).value = "P.O.Box 56"
                objWorkSheet.cells(3,1).value = "Koue Bokkeveld"
                objWorkSheet.cells(4,1).value = "'6836'"
                objWorkSheet.cells(5,1).value = "Phone : 0233170983 /023 3170588"
                objWorkSheet.cells(6,1).value = "Fax : 023 3170597/0867160325"
                objWorkSheet.cells(7,1).value = "E-pos : jacob@kbos.co.za"
              objWorkSheet.cells(9,1).value = "PROGRESS REPOR REF No."+Request.Form("qualification")
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(1, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(1,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(2, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(2,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(3, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(3,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(4, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(4,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(5, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(5 ,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(6, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(6,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(7, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(7,3)).Merge
               objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(8, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(8 ,3)).Merge
              objExcel.ActiveWorkbook.ActiveSheet.Range(objExcel.ActiveWorkbook.ActiveSheet.Cells(9, 1), objExcel.ActiveWorkbook.ActiveSheet.Cells(9 ,3)).Merge

              objWorkSheet.cells(11,1).value = "Trainer Name"
              objWorkSheet.cells(11,2).value = "Trainer Contact Number"
              objWorkSheet.cells(11,3).value = "Company"

              Set objWorkSheet2 = objWorkbook.Worksheets(2)
              objWorkSheet2.Range(objWorkSheet2.Cells(1, 4), objWorkSheet2.Cells(1 ,7)).Merge
              objWorkSheet2.Name = "Learner Standard"
              'Put the first row in bold
               objWorkSheet.Range("A11:C11").Font.Bold = True
              objWorkSheet.Range("A11:C11").Font.Size = 12
              objWorkSheet2.Range("A1:E1").Font.Bold = True
              objWorkSheet2.Range("A1:E1").Font.Size = 12
              objWorkSheet2.Range("A2:G2").Font.Bold = True
              objWorkSheet2.Range("A2:G2").Font.Size = 12
              'Freeze the panes
             '   objWorkSheet2.Range("A1:C1").Select
              '  objExcel.ActiveWindow.FreezePanes = True
              objWorkSheet2.cells(1,4).value = "Unitstandard numbers"
              objWorkSheet2.cells(2,1).value = "No"
              objWorkSheet2.cells(2,2).value = "Surname"
              objWorkSheet2.cells(2,3).value = "Name"
              objWorkSheet2.cells(2,4).value = "Electives"
              objWorkSheet2.cells(2,5).value = "Fundamentals"
              objWorkSheet2.cells(2,6).value = "Core"
              objWorkSheet2.cells(2,7).value = "N/A"
              For column = 1 to 12
              objExcel.Columns(column).AutoFit()
              Next
              objWorkSheet.Columns(1).ColumnWidth = 30
              objWorkSheet.Columns(2).ColumnWidth = 20
              objWorkSheet.Columns(3).ColumnWidth = 25
              objWorkSheet2.Columns(2).ColumnWidth = 20
              objWorkSheet2.Columns(3).ColumnWidth = 20
              For column =4 to 7
              objWorkSheet2.Columns(column).AutoFit()
              Next
dbname="data/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName
          
Set objRS = objDB.Execute("SELECT t.Tname,T.TSname FROM (DATA inner join learnerdata as l on data.Student_NUM=l.Student_NUM ) inner join teacher as t on t.tnoid=l.AssessorID WHERE data.NATQUA='"+split(Request.Form("qualification"),",")(0)+"' and Data.project='"+split(Request.Form("qualification"),  ",")(1)+"' group by t.Tname,T.TSname")

If objRS.EOF Then
	objExcel.cells(12,1).value = "No Trainers found for Qualification"
	objRS.Close
	objDB.Close
	Set objRS = Nothing
	Set objDB = Nothing
	'Response.End
              else
              value = 12 'value + 1
              
              ColumnCounter = 1 'value + 1
              Do While Not objRS.EOF
			'Html = html &"value=" & Chr(34) & objRS("NQName") & Chr(34) & ">" &objRS("NQName")
	        objWorkSheet.cells(value,1).value = objRS("Tname") & " " & objRS("TSname")
            objWorkSheet.cells(value,2).value = "-"
            objWorkSheet.cells(value,3).value = "-"
            rowCounter = rowCounter+1
            objRS.MoveNext
            value = value +1
			Loop
            a = Now()
            End If
              rowCounter = 3 'value + 1
              sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath
            Set objDB = Server.CreateObject("ADODB.Connection")
            objDB.Open sDBName
             ' Response.Write("SELECT Data.SName,Data.FName,s.SNumber,s.CType FROM (DATA inner join learnerdata as l on data.Student_NUM=l.Student_NUM ) inner join standards as s on s.stitle=l.stitle WHERE l.SCompetent='Competent' and data.NATQUA='"+split(Request.Form("qualification"),",")(0)+"' and Data.project='"+split(Request.Form("qualification"),  ",")(1)+"'")   
			   Set objRS2 = objDB.Execute("SELECT Data.SName,Data.FName,s.SNumber,s.CType FROM (DATA inner join learnerdata as l on data.Student_NUM=l.Student_NUM ) inner join standards as s on s.stitle=l.stitle WHERE l.SCompetent='Competent' and data.NATQUA='"+split(Request.Form("qualification"),",")(0)+"' and Data.project='"+split(Request.Form("qualification"),  ",")(1)+"'")
               'Set objRS2 = objDB.Execute("select * from DATA as d left join learnerdata as l on l.Student_NUM=d.Student_NUM and l.SCompetent='Competent' left join standards as s on s.stitle=l.stitle  where d.NATQUA='"+split(Request.Form("qualification"),",")(0)+"'")
               If objRS2.EOF Then
	           with objWorkSheet2
              .Cells(rowCounter, 1) = rowCounter - 2
              .Cells(rowCounter, 2) ="No Learners Found"
              .Cells(rowCounter, 3) = "-"
              .Cells(rowCounter, 4) = "-"
              .Cells(rowCounter, 5) = "-"
              .Cells(rowCounter, 6) = "-"
              .Cells(rowCounter, 7) = "-"
              END With
	            objRS2.Close
	            Set objRS2 = Nothing
	            'Response.End
              else
               sNumberType = ""
              sNumberTypeOld = ""
              sNumberOld=""
              StudentE=""
              StudentF=""
              StudentC=""
              StudentNA=""
              Do While Not objRS2.EOF
              if sNumberOld = objRS2("Sname") & "," & objRS2("Fname") OR sNumberOld = ""  then
                  Select Case objRS2("CType")
                  Case "C"     StudentC = StudentC & " , " & objRS2("snumber")
                  Case "F"     StudentF = StudentF & " , " & objRS2("snumber")
                  Case "E"     StudentE = StudentE & " , " & objRS2("snumber")
                  Case Else      StudentNA = StudentNA & " , " & objRS2("snumber")
                  End Select
              else
                  With objWorkSheet2
                  .Cells(rowCounter, 1) = rowCounter - 2
                  .Cells(rowCounter, 2) = split(sNumberOld,",")(0)' objRS2("Sname")
                  .Cells(rowCounter, 3) = split(sNumberOld,",")(1)
                  .Cells(rowCounter, 4) = StudentE
                  .Cells(rowCounter, 5) = StudentF
                  .Cells(rowCounter, 6) = StudentC
                  .Cells(rowCounter, 7) = StudentNA
                  END With
                  rowCounter = rowCounter+1
                  sNumberType = ""
                  sNumberTypeOld = ""
                  sNumberOld=""
                  StudentE=""
                  StudentF=""
                  StudentC=""
                  StudentNA=""
              Select Case objRS2("CType")
                  Case "C"     StudentC = StudentC & " , " & objRS2("snumber")
                  Case "F"     StudentF = StudentF & " , " & objRS2("snumber")
                  Case "E"     StudentE = StudentE & " , " & objRS2("snumber")
                  Case Else      StudentNA = StudentNA & " , " & objRS2("snumber")
               End Select
              End If
              sNumberOld = objRS2("Sname") & "," & objRS2("Fname")
              objRS2.MoveNext
              Loop
              With objWorkSheet2
                  .Cells(rowCounter, 1) = rowCounter - 2
                  .Cells(rowCounter, 2) = split(sNumberOld,",")(0)' objRS2("Sname")
                  .Cells(rowCounter, 3) = split(sNumberOld,",")(1)
                  .Cells(rowCounter, 4) = StudentE
                  .Cells(rowCounter, 5) = StudentF
                  .Cells(rowCounter, 6) = StudentC
                  .Cells(rowCounter, 7) = StudentNA
              END WITH
            End If
             
              'Quit Excel
              ' objWorkbook.Saveas "c:\kbosreport\testXLS" &Second(Now()) &".xlsx"
              '  objWorkbook.Close
              '  objExcel.workbooks.close
            objExcel.Application.Quit
 
            'Clean Up
            
              Set objWorkSheet = Nothing
            Set objWorkSheet2 = Nothing
            Set objExcel = Nothing
               ' objExcel.quit
                set objExcel = nothing 
           
               End Select
               
%>
</font></td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
        <tr>
          <td width="116">&nbsp;</td>
          <td width="484">&nbsp;</td>
        </tr>
      </table>
      </center>
    </div>
            </blockquote>
    <p>&nbsp;</p>
    <p>&nbsp;</td>
  </tr>
  <tr>
    <td><!---#include file = "inc/Foot.asp"----></td>
  </tr>
</table>

</body>

</html>