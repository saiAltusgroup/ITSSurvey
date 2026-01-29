<%

' DK, Apr 01, 2008 - version 1.0

Function getXMLFromDB(myXmlID)

	Dim objComm, rs

	Set objComm = Server.CreateObject("ADODB.Command")
	Set rs = Server.CreateObject("ADODB.Recordset")

	With objComm

		.ActiveConnection = its_conn
		.CommandText = "usp_Get_XML"
		.CommandType = adCmdStoredProc

		'Define stored procedure params and append to command'

		.Parameters.Append(.CreateParameter("@xmlID", adInteger, adParamInput, 0))

	End With

	objComm("@xmlID") = myXmlID
	rs.Open = objComm.Execute

	If not rs.EOF Then
		xmlStr = rs("xmlStr")
	End If

	rs.Close

	Set rs = Nothing
	Set objComm = Nothing

	getXMLfromDB = xmlStr

End Function


Function putXmlToDB(myXmlID, xmlStr)

	Dim objComm, thisSession

	Set objComm = Server.CreateObject("ADODB.Command")
	thisSession = Session.SessionID

	With objComm

		.ActiveConnection = its_conn
		.CommandText = "usp_Put_XML"
		.CommandType = adCmdStoredProc

		'Define stored procedure params and append to command'

		.Parameters.Append(.CreateParameter("ReturnStatus", adInteger, adParamReturnValue))
		.Parameters.Append(.CreateParameter("@xmlID", adInteger, adParamInput, 0))
		.Parameters.Append(.CreateParameter("@thisSession", adVarChar, adParamInput, 50))
		.Parameters.Append(.CreateParameter("@xmlStr", adLongVarChar, adParamInput, len(xmlStr)))

	End With

	objComm("@xmlID") = myXmlID
	objComm("@thisSession") = CStr(thisSession)
	objComm("@xmlStr") = xmlStr
	objComm.Execute

	myXmlID = objComm("ReturnStatus")

	Set objComm = Nothing

	putXmlToDB = myXmlID

End Function


Function saveXmlInDB(xmlStr)

	saveXmlinDB = putXmlToDB(0, xmlStr)

End Function

%>
