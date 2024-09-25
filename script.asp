<%

dim dbname
dim cnpath
dim sDBName
dim value
dim qr

dbname="DATA/learner.mdb"
cnpath="DBQ=" & server.mappath(dbname)
sDBName = "driver={Microsoft Access Driver (*.mdb)}; " & cnpath


Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open sDBName
if request.QueryString("qr") then
	qr = request.QueryString("qr")
	Set objRS9 = objDB.Execute("select * from [OFOCode] where OFOCODE='"&qr&"'")
     response.write objRS9("OFODesc")
end if 

if request.QueryString("company") then
	  value = request.QueryString("company")
	  Set objRS7 = objDB.Execute("select * from [Client] where ID_no="&CInt(value))
	  response.write objRS7("TrainingManager")&";"
      response.write objRS7("CNumber")&";"
      response.write objRS7("SICCode")&";"
      response.write objRS7("SSUCode")&";"
end if




%>


