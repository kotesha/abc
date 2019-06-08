Dim objConnection 
'Set Adodb Connection Object
Set objConnection = CreateObject("ADODB.Connection")     
Dim objRecordSet 
 
'Create RecordSet Object
Set objRecordSet = CreateObject("ADODB.Recordset")     
 
Dim DBQuery 'Query to be Executed
DBQuery = "select * from vnv_employee"
 
'Connecting using SQL OLEDB Driver
objConnection.Open "Provider = sqloledb.1; Server =.\SQLEXPRESS;User Id = System; Password=Pass123 ; Database = Trial"

msgbox "read data" 

'Execute the Query
objRecordSet.Open DBQuery,objConnection
 
'Return the Result Set
Value = objRecordSet.fields.item(0)				
msgbox Value
 
' Release the Resources
objRecordSet.Close        
objConnection.Close		
 
Set objConnection = Nothing
Set objRecordSet = Nothing