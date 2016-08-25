<%
' Declare variables
Dim oConnection, oRecordSet, sSql
Dim lStationId

' Open the database
Set oConnection = Server.CreateObject("ADODB.Connection")
oConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("App_Data/eudia.mdb")

' Check the API key
Set oRecordSet = Server.CreateObject("ADODB.Recordset") 	
sSql = "SELECT Station.Id FROM Station WHERE Station.Key = '" & Replace(Request.Form("key"), "'", "''") & "'"
oRecordSet.Open sSql, oConnection
If oRecordSet.EOF Then
	Response.Status = "403 Forbidden"
Else
	lStationId = oRecordSet.Fields("Id").Value 
End If

' If we have an API key, store the values
If lStationId > 0 Then
	sSql = "INSERT INTO Observation (StationId, Recorded, Temperature1, Temperature2, Temperature3, Pressure, Humidity) VALUES (" & lStationId & ", " & CDbl(Request.Form("recorded")) & ", " & CDbl(Request.Form("temperature1")) & ", " & CDbl(Request.Form("temperature2")) & ", " & CDbl(Request.Form("temperature3")) & ", " & CDbl(Request.Form("pressure")) & ", " & CDbl(Request.Form("humidity")) & ")"
	oConnection.Execute sSql
End If

' Clean up
Set oRecordSet = Nothing
Set oConnection = Nothing
%>