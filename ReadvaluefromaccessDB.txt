Dim objCon,objRs
Set objCon=createobject("ADODB.connection")
Set objRs=createobject("ADODB.recordset")
objCon.provider="Microsoft.ACE.OLEDB.12.0"
objCon.open "C:\Users\shbharti\Documents\DBChkpoint.accdb"
objRs.open "select * from Employee",objCon


Do Until  objRs.EOF =true
fname = objRs.fields("FirstName")
lname= objRs.Fields("Lastname")

msgbox fname

SystemUtil.Run "C:\Program Files (x86)\HPE\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe" 
WpfWindow("HPE MyFlight Sample Applicatio").WpfEdit("agentName").Set fname
WpfWindow("HPE MyFlight Sample Applicatio").WpfEdit("password").Set lname
WpfWindow("HPE MyFlight Sample Applicatio").WpfButton("OK").Click
WpfWindow("HPE MyFlight Sample Applicatio").Close



objRs.movenext
Loop

objRs.close
objCon.close