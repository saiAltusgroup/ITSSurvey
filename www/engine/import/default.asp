
<!--#INCLUDE virtual="/sys_include/db.asp"-->

<%

Dim UpdateDb
Dim WriteFile
Dim PeriodID
Dim strQuestionTable
Dim strFile
Redim aryError(0)



UpdateDB = false
WriteFile = false

Server.ScriptTimeout = 600 '10 minutes

'strTablePrefix="dat_Test_"
strTablePrefix="dat_"

'---------------------- You must set this ----------------------
' PeriodId 144 is Q1 2018 => file: "data/batch_44/ITS Q1 2018 Datafile.dat"
' PeriodId 145 is Q2 2018 => file: "data/batch_45/ITS Q2 2018 Datafile.dat"
'---------------------------------------------------------------
PeriodID = 156
strFile = "data/batch_56/ITS Q1 2021 Datafile.dat"

' ==============================================================

Dim objRs
Dim strSQL


if(Request.Form("submit") = "import") then
	Dim intLine

	if(ValidateRequired("Period",PeriodID)) Then
		Call ValidateNumber("Period",PeriodID)
	end if
	Call ValidateRequired("Path to File",strFile)

	if(Len(aryError(0)) = 0) then

		PeriodID = Cint(PeriodID)
		aryTemp = Split(GetFileAsString(server.mappath(strFile)), vbCrlf)

		Set objRs = Server.CreateObject("ADODB.Recordset")
		Set objComm = Server.CreateObject("ADODB.Command")

		objComm.ActiveConnection=its_conn

		Response.Write "UBound(aryTemp)=" & UBound(aryTemp) & "<br />"

		for i=0 to UBound(aryTemp)

			Response.Write "<hr />"
			Response.Write "Record " & i & "<br />"

			aryTemp2 = Split(aryTemp(i), Chr(9))

			intColumns = UBound(aryTemp2)

			Response.Write "intColumns=" & intColumns & "<br />"

			if(i = 0) then 'Column Headers

				Redim aryColumns(intColumns,4)

				for j=0 to UBound(aryTemp2)

					aryTemp2(j)=fixQuestionLabel(aryTemp2(j))

					strQuestionLabel=trim(aryTemp2(j))
					strSQL = "SELECT QuestionID,ISNULL(MapTo,LabelShort) AS 'LabelShort',QuestionTypeID,Ignored FROM " & strTablePrefix & "Question WHERE LabelShort = '" & strQuestionLabel & "'"

					objRs.Open strSQL, its_conn

					if(objRs.EOF) then

						if isIgnoredQuestion(strQuestionLabel) then
							aryColumns(j,0) = 0
							aryColumns(j,1) = strQuestionLabel
							aryColumns(j,2) = -1
							aryColumns(j,3) = 0.0
						else
							ProduceError(aryTemp2(j) & " is not found in the columns table.")
						end if

					else
						aryColumns(j,0) = objRs("QuestionID")
						aryColumns(j,1) = trim(objRs("LabelShort"))
						aryColumns(j,2) = Cint(objRs("QuestionTypeID"))
						aryColumns(j,3) = objRs("Ignored")
					end if

					objRs.Close()

				Next

				if WriteFile then Call WriteLine(server.mappath(strFile & ".out"), aryTemp(i))


			else

				if(Len(aryError(0)) <> 0) then exit for
				if(UBound(aryTemp2)<0) then exit for
				if(Len(aryTemp2(0))<1) then exit for

				objComm.CommandText="INSERT INTO " & strTablePrefix & "Response(PeriodID) VALUES("& PeriodID &")"
				If (UpdateDB) Then objComm.Execute
				Response.Write objComm.CommandText & "<br />"


				strSQL = "SELECT Max(ResponseID) AS 'ResponseID' FROM " & strTablePrefix & "Response"

				objRs.Open strSQL, its_conn

				if( not objRs.EOF) then
					intResponse = objRS("ResponseID")
				else
					intResponse = 0
				end if

				objRs.Close()


				Response.Write "intResponse=" & intResponse & "<br />"

				line=""
				line_part1=""
				line_part2=""

				for j=0 to UBound(aryTemp2)




					if(aryColumns(j,2) = 0) then

						if aryColumns(j,1)="CONTRIB" then aryColumns(j,1)="CONTRIBUT"

						objComm.CommandText = "UPDATE " & strTablePrefix & "Response SET " & aryColumns(j,1) & "='"&EscapeSQL(aryTemp2(j))&"' WHERE ResponseID =" & intResponse
						Response.Write objComm.CommandText & "<br />"
						'if(aryTemp2(j) <> aryColumns(j,3)) then
							If (UpdateDB) Then objComm.Execute
							Response.Write "Command executed<br />"
						'end if

					elseif(aryColumns(j,2) = 1) then

						Response.Write "aryColumns=[" & aryColumns(j,1) & "]<br />"
						Response.Write "aryTemp2=[" & Trim(aryTemp2(j)) & "]<br />"

						if Trim(aryTemp2(j))="" then aryTemp2(j)=aryColumns(j,3) ' added by DK on Jun 25, 2008 to accept NULL values

						objComm.CommandText = "INSERT INTO " & strTablePrefix & "ResponseItemInt(QuestionID,ResponseID,Value) VALUES("&aryColumns(j,0)&","&intResponse&","&CLng(Trim(aryTemp2(j)))&")"
						if(CLng(aryTemp2(j)) <> CLng(aryColumns(j,3))) then
							If (UpdateDB) Then objComm.Execute
							Response.write "<strong>Imported</strong> int value " & aryTemp2(j) & "<br />"
						else
							Response.write "Ignored int value " & aryTemp2(j) & " or " & aryColumns(j,3) & "<br />"
						end if

					elseif(aryColumns(j,2) = 2) then

						'Response.Write "aryColumns=[" & aryColumns(j,1) & "]<br />"
						'Response.Write "aryTemp2=[" & Trim(aryTemp2(j)) & "]<br />"

						if Trim(aryTemp2(j))="" then aryTemp2(j)=aryColumns(j,3) ' added by DK on Jun 25, 2008 to accept NULL values

						
						'lacfre 2020-06-04 check CDbl error
						' if(Not IsNumeric(Trim(aryTemp2(j)))) then
							' Response.write  Trim(aryTemp2(j)) & "   IS NOT NUMERIC*********************************"  & "<br />"
						' end if
						
						
						objComm.CommandText = "INSERT INTO " & strTablePrefix & "ResponseItemFloat(QuestionID,ResponseID,Value) VALUES("&aryColumns(j,0)&","&intResponse&","&CDbl(Trim(aryTemp2(j)))&")"
						
						
						if(CDbl(aryTemp2(j)) <> CDbl(aryColumns(j,3))) then
							Response.write "<strong>Imported</strong> int value " & aryTemp2(j) & "<br />"
							If (UpdateDB) Then objComm.Execute
						else
							Response.write "Ignored int value " & aryTemp2(j) & " or " & aryColumns(j,3) & "<br />"
						end if

					elseif(aryColumns(j,2) = 3) then

						objComm.CommandText = "INSERT INTO " & strTablePrefix & "ResponseItemTxt(QuestionID,ResponseID,Value) VALUES("&aryColumns(j,0)&","&intResponse&",'"&EscapeSQL(aryTemp2(j))&"')"
						if(Trim(aryTemp2(j)) <> aryColumns(j,3)) then
							If (UpdateDB) Then objComm.Execute
						end if

					elseif(aryColumns(j,2) = -1) then

							Response.write "Ignored question <b>" & aryColumns(j,1) & "</b><br />"
					end if



					if (Len(line)>0) then line = line & Chr(9)
					line = line & aryTemp2(j)

					if j<499 then

						if (Len(line_part1)>0) then line_part1 = line_part1 & Chr(9)
						line_part1 = line_part1 & aryTemp2(j)

					else

						if (Len(line_part2)>0) then line_part2 = line_part2 & Chr(9)
						line_part2 = line_part2 & aryTemp2(j)

					end if


				next

				Response.Write "End Record<br />"



				if WriteFile then
					Call WriteLine(server.mappath(strFile & ".out"), line)
					Call WriteLine(server.mappath(strFile & ".part1.out"), line_part1)
					Call WriteLine(server.mappath(strFile & ".part2.out"), line_part2)
				end if


			end if

		Next

	end if

 end if




If(Request.Form("submit") = "filter") Then

	Response.Write "Start Filter<br />"

	'if question Type is float and value count greater then 5 remove one min and one max value

	Set objRs = Server.CreateObject("ADODB.Recordset")
	Set objComm = Server.CreateObject("ADODB.Command")

	objComm.ActiveConnection=its_conn

	strSQL = "SELECT a.QuestionID,b.MaxVal,c.MinVal " &_
		"FROM	( " &_
		"		SELECT QuestionID " &_
		"		FROM " & strTablePrefix & "ResponseItemFloat rit  " &_
		"			JOIN " & strTablePrefix & "Response r " &_
		"				ON rit.ResponseID = r.ResponseID " &_
		"		WHERE r.PeriodID = " & PeriodID & " " &_
		"		GROUP BY QuestionID " &_
		"		HAVING COUNT(rit.ResponseItemFloatID) > 5 " &_
		"	) a " &_
		"	JOIN ( " &_
		"		SELECT QuestionID,MAX(Value) AS MaxVal " &_
		"		FROM " & strTablePrefix & "ResponseItemFloat rit JOIN dat_Response r ON rit.ResponseID = r.ResponseID WHERE r.PeriodID = " & PeriodID & " " &_
		"		GROUP BY QuestionID " &_
		"	) b " &_
		"		ON a.QuestionID = b.QuestionID " &_
		"	JOIN ( " &_
		"		SELECT QuestionID,MIN(Value) AS MinVal " &_
		"		FROM " & strTablePrefix & "ResponseItemFloat rit JOIN dat_Response r ON rit.ResponseID = r.ResponseID WHERE r.PeriodID = " & PeriodID & " " &_
		"		GROUP BY QuestionID " &_
		"	) c " &_
		"		ON a.QuestionID = c.QuestionID " &_
		"ORDER BY a.QuestionID"

	Response.Write strSQL & "<hr>"

	objRs.Open strSQL, its_conn

	if(objRs.EOF) then
		ProduceError("No available responses for period id " & PeriodID & " have been found.")
	else
		Do While Not objRs.EOF
			objComm.CommandText="DELETE " &_
				"FROM " & strTablePrefix & "ResponseItemFloat " &_
				"WHERE ResponseItemFloatID IN ( " &_
				" " &_
				"		SELECT Top 1 rit.ResponseItemFloatID " &_
				"		FROM " & strTablePrefix & "ResponseItemFloat rit " &_
				"			JOIN " & strTablePrefix & "Response r " &_
				"				ON rit.ResponseID = r.ResponseID " &_
				"		WHERE r.PeriodID = " & PeriodID & " " &_
				"		AND rit.QuestionID =  " & objRs("QuestionID") &_
				"		AND rit.Value =  " & objRs("MaxVal") &_
				"		ORDER BY Value DESC " &_
				" " &_
				"	) " &_
				"	OR ResponseItemFloatID IN ( " &_
				" " &_
				"		SELECT Top 1 rit.ResponseItemFloatID " &_
				"		FROM " & strTablePrefix & "ResponseItemFloat rit " &_
				"			JOIN " & strTablePrefix & "Response r " &_
				"				ON rit.ResponseID = r.ResponseID " &_
				"		WHERE r.PeriodID = " & PeriodID & " " &_
				"		AND rit.QuestionID =  " & objRs("QuestionID") &_
				"		AND rit.Value =  " & objRs("MinVal") &_
				"		ORDER BY rit.Value ASC " &_
				"	) "

			If (UpdateDB) Then objComm.Execute
			objRs.MoveNext

			Response.Write objComm.CommandText & "<hr/>"

		loop
	end if

	objRs.Close()

	Response.Write "End Filter<br />"

End If




%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
	<head>
		<title>Data Importer and Filter</title>
	</head>
	<body>

	<h1>Data Importer and Filter</h1>

<%
		If( Len(aryError(0)) > 0 ) Then
			ProduceErrorList aryError
		End If
%>
		<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
			<input type="submit" name="submit" id="import" value="import" />
			<input type="submit" name="submit" id="filter" value="filter" />
		</form>
	</body>
</html>












<%
'subroutines




Sub ProduceErrorList(a)
	If(Len(a(0)) > 0) Then
%>
	<ol class="error">
<%
	for each e in a
%>
	<li><%=e%></li>
<%
	next
%>
	</ol>
<%
	End If
End Sub


Sub WriteLine(filename, line)

	dim filesys, filetxt

	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set filetxt = filesys.OpenTextFile(filename, 8, True)

	filetxt.WriteLine(line)
	filetxt.Close

	Set filesys = nothing
	Set filetxt = nothing

End Sub



'functions

Function FileExists(byVal pathname)

	'declare variables
	Dim objFSO

	'create objects
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'return existance of file
	FileExists = objFSO.FileExists(pathname)

	'set object to nothing
	Set objFSO = Nothing

End Function

Function GetFileAsString(filename)

	'declare variables
	Dim objFSO, objFile, objTS, strReturn

	'Check to see that file exists return
	If(Not FileExists(filename) ) Then
		Call ProduceError(filename & " file does not exist")
		exit function
	End If

	'set values
	intCount = 0
	strReturn = ""

	'create objects
	Set objFSO = server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(filename)
	Set objTS = objFile.OpenAsTextStream(1,0)

	Do While Not objTS.atEndofstream
		'Trim the begining and ending spaces of file
		strLine = Trim(objTS.ReadLine)
		'test for empty line
		If(Len(strLine) > 0) Then
			'Test for comment
			If(Left(strLine,1) <> "#") Then
				'add line
				strReturn = strReturn & strLine & vbCrlf
			End If
		End If
	Loop

	'close object
	objTS.Close

	'set objects to nothing
	Set objTS = nothing
	Set objFile = nothing
	Set objFSO = nothing

	GetFileAsString = strReturn

End Function

Function ProduceError(s)
	Dim c

	c = Ubound(aryError)

	If(Len(aryError(c)) > 0) Then
		c = c + 1
		ReDim Preserve aryError(c)
	End If

	aryError(c) = s

	ProduceError = aryError

End Function

function ValidateRequired(Name,Value)

	r = true

	if(len(Value) = 0) then
		ProduceError(Name & " is required;")
		r = false
	end if

	ValidateRequired = r

end function

function ValidateNumber(Name,Value)

	r = true

	if(not isNumeric(Value)) then
		ProduceError(Name & " must be a number;")
		r = false
	end if

	ValidateNumber = r

end function

function EscapeSQL(s)

	EscapeSQL = Replace(Trim(s),"'","''")

end function


Function fixQuestionLabel(label)

		l=label

		l=Replace(l, "VO124A_", "VPO124A_")
		l=Replace(l, "VO124B_", "VPO124B_")
		l=Replace(l, "VO124C_", "VPO124C_")
		l=Replace(l, "VO124D_", "VPO124D_")
		l=Replace(l, "VO124E_", "VPO124E_")
		l=Replace(l, "VO124F_", "VPO124F_")
		l=Replace(l, "VO124G_", "VPO124G_")
		l=Replace(l, "VO124H_", "VPO124H_")

		fixQuestionLabel=l

End Function



Function isIgnoredQuestion(label)

	l=ucase(label)
	isIgnoredQuestion=true

	if (l="CD82CD1_OTHER") then exit function
	if (l="CD82CD2_OTHER") then exit function
	if (l="CD82CD3_OTHER") then exit function

	if (l="APMB19A_O") then exit function
	if (l="APMB19A_OP") then exit function
	if (l="APMB19A_D") then exit function
	if (l="APMB19A_DP") then exit function
	if (l="APMB19B_O") then exit function
	if (l="APMB19B_OP") then exit function
	if (l="APMB19B_D") then exit function
	if (l="APMB19B_DP") then exit function
	if (l="APMB19C_O") then exit function
	if (l="APMB19C_OP") then exit function
	if (l="APMB19C_D") then exit function
	if (l="APMB19C_CP") then exit function
	if (l="APMB19C_DP") then exit function
	if (l="APMB19D_O") then exit function
	if (l="APMB19D_OP") then exit function
	if (l="APMB19D_D") then exit function
	if (l="APMB19D_DP") then exit function
	if (l="APMB19E_O") then exit function
	if (l="APMB19E_OP") then exit function
	if (l="APMB19E_D") then exit function
	if (l="APMB19E_DP") then exit function
	if (l="APMB19F_O") then exit function
	if (l="APMB19F_OP") then exit function
	if (l="APMB19F_D") then exit function
	if (l="APMB19F_DP") then exit function
	if (l="APMB20") then exit function
	if (l="Q1R") then exit function
	if (l="VPO1A_V") then exit function
	if (l="VPO1A_E") then exit function
	if (l="VPO1A_C") then exit function
	if (l="VPO1A_T") then exit function
	if (l="VPO1A_O") then exit function
	if (l="VPO1A_M") then exit function
	if (l="VPO1A_H") then exit function
	if (l="VPO1B_V") then exit function
	if (l="VPO1B_E") then exit function
	if (l="VPO1B_C") then exit function
	if (l="VPO1B_T") then exit function
	if (l="VPO1B_O") then exit function
	if (l="VPO1B_M") then exit function
	if (l="VPO1B_H") then exit function
	if (l="VPO1C_V") then exit function
	if (l="VPO1C_E") then exit function
	if (l="VPO1C_C") then exit function
	if (l="VPO1C_T") then exit function
	if (l="VPO1C_O") then exit function
	if (l="VPO1C_M") then exit function
	if (l="VPO1C_H") then exit function
	if (l="VPO1D_V") then exit function
	if (l="VPO1D_E") then exit function
	if (l="VPO1D_C") then exit function
	if (l="VPO1D_T") then exit function
	if (l="VPO1D_O") then exit function
	if (l="VPO1D_M") then exit function
	if (l="VPO1D_H") then exit function
	if (l="VPO1E_V") then exit function
	if (l="VPO1E_E") then exit function
	if (l="VPO1E_C") then exit function
	if (l="VPO1E_T") then exit function
	if (l="VPO1E_O") then exit function
	if (l="VPO1E_M") then exit function
	if (l="VPO1E_H") then exit function
	if (l="VPO1F_V") then exit function
	if (l="VPO1F_E") then exit function
	if (l="VPO1F_C") then exit function
	if (l="VPO1F_T") then exit function
	if (l="VPO1F_O") then exit function
	if (l="VPO1F_M") then exit function
	if (l="VPO1F_H") then exit function
	if (l="VPO1G_V") then exit function
	if (l="VPO1G_E") then exit function
	if (l="VPO1G_C") then exit function
	if (l="VPO1G_T") then exit function
	if (l="VPO1G_O") then exit function
	if (l="VPO1G_M") then exit function
	if (l="VPO1G_H") then exit function
	if (l="VPO1H_V") then exit function
	if (l="VPO1H_E") then exit function
	if (l="VPO1H_C") then exit function
	if (l="VPO1H_T") then exit function
	if (l="VPO1H_O") then exit function
	if (l="VPO1H_M") then exit function
	if (l="VPO1H_H") then exit function
	if (l="VPO5A_V") then exit function
	if (l="VPO5A_E") then exit function
	if (l="VPO5A_C") then exit function
	if (l="VPO5A_T") then exit function
	if (l="VPO5A_O") then exit function
	if (l="VPO5A_M") then exit function
	if (l="VPO5A_H") then exit function
	if (l="VPO5B_V") then exit function
	if (l="VPO5B_E") then exit function
	if (l="VPO5B_C") then exit function
	if (l="VPO5B_T") then exit function
	if (l="VPO5B_O") then exit function
	if (l="VPO5B_M") then exit function
	if (l="VPO5B_H") then exit function
	if (l="VPO5C_V") then exit function
	if (l="VPO5C_E") then exit function
	if (l="VPO5C_C") then exit function
	if (l="VPO5C_T") then exit function
	if (l="VPO5C_O") then exit function
	if (l="VPO5C_M") then exit function
	if (l="VPO5C_H") then exit function
	if (l="VPO5D_V") then exit function
	if (l="VPO5D_E") then exit function
	if (l="VPO5D_C") then exit function
	if (l="VPO5D_T") then exit function
	if (l="VPO5D_O") then exit function
	if (l="VPO5D_M") then exit function
	if (l="VPO5D_H") then exit function
	if (l="VPO6A_V") then exit function
	if (l="VPO6A_E") then exit function
	if (l="VPO6A_C") then exit function
	if (l="VPO6A_T") then exit function
	if (l="VPO6A_O") then exit function
	if (l="VPO6A_M") then exit function
	if (l="VPO6A_H") then exit function
	if (l="VPO6B_V") then exit function
	if (l="VPO6B_E") then exit function
	if (l="VPO6B_C") then exit function
	if (l="VPO6B_T") then exit function
	if (l="VPO6B_O") then exit function
	if (l="VPO6B_M") then exit function
	if (l="VPO6B_H") then exit function
	if (l="VPO6C_V") then exit function
	if (l="VPO6C_E") then exit function
	if (l="VPO6C_C") then exit function
	if (l="VPO6C_T") then exit function
	if (l="VPO6C_O") then exit function
	if (l="VPO6C_M") then exit function
	if (l="VPO6C_H") then exit function
	if (l="VPO6D_V") then exit function
	if (l="VPO6D_E") then exit function
	if (l="VPO6D_C") then exit function
	if (l="VPO6D_T") then exit function
	if (l="VPO6D_O") then exit function
	if (l="VPO6D_M") then exit function
	if (l="VPO6D_H") then exit function
	if (l="VPO90") then exit function
	if (l="VPO105") then exit function
	if (l="VPO109_V") then exit function
	if (l="VPO109_E") then exit function
	if (l="VPO109_C") then exit function
	if (l="VPO109_T") then exit function
	if (l="VPO109_O") then exit function
	if (l="VPO109_M") then exit function
	if (l="VPO109_H") then exit function
	if (l="VPO110_V") then exit function
	if (l="VPO110_E") then exit function
	if (l="VPO110_C") then exit function
	if (l="VPO110_T") then exit function
	if (l="VPO110_O") then exit function
	if (l="VPO110_M") then exit function
	if (l="VPO110_H") then exit function
	if (l="VPO111") then exit function
	if (l="VPO112A_V") then exit function
	if (l="VPO112A_E") then exit function
	if (l="VPO112A_C") then exit function
	if (l="VPO112A_T") then exit function
	if (l="VPO112A_O") then exit function
	if (l="VPO112A_M") then exit function
	if (l="VPO112A_H") then exit function
	if (l="VPO112B_V") then exit function
	if (l="VPO112B_E") then exit function
	if (l="VPO112B_C") then exit function
	if (l="VPO112B_T") then exit function
	if (l="VPO112B_O") then exit function
	if (l="VPO112B_M") then exit function
	if (l="VPO112B_H") then exit function
	if (l="VPO112C_V") then exit function
	if (l="VPO112C_E") then exit function
	if (l="VPO112C_C") then exit function
	if (l="VPO112C_T") then exit function
	if (l="VPO112C_O") then exit function
	if (l="VPO112C_M") then exit function
	if (l="VPO112C_H") then exit function
	if (l="VPO112D_V") then exit function
	if (l="VPO112D_E") then exit function
	if (l="VPO112D_C") then exit function
	if (l="VPO112D_T") then exit function
	if (l="VPO112D_O") then exit function
	if (l="VPO112D_M") then exit function
	if (l="VPO112D_H") then exit function
	if (l="VPO114_V") then exit function
	if (l="VPO114_E") then exit function
	if (l="VPO114_C") then exit function
	if (l="VPO114_T") then exit function
	if (l="VPO114_O") then exit function
	if (l="VPO114_M") then exit function
	if (l="VPO114_H") then exit function
	if (l="VPO114A_V") then exit function
	if (l="VPO114A_E") then exit function
	if (l="VPO114A_C") then exit function
	if (l="VPO114A_T") then exit function
	if (l="VPO114A_O") then exit function
	if (l="VPO114A_M") then exit function
	if (l="VPO114A_H") then exit function
	if (l="VPO115_V") then exit function
	if (l="VPO115_E") then exit function
	if (l="VPO115_C") then exit function
	if (l="VPO115_T") then exit function
	if (l="VPO115_O") then exit function
	if (l="VPO115_M") then exit function
	if (l="VPO115_H") then exit function
	if (l="VPO115A_V") then exit function
	if (l="VPO115A_E") then exit function
	if (l="VPO115A_C") then exit function
	if (l="VPO115A_T") then exit function
	if (l="VPO115A_O") then exit function
	if (l="VPO115A_M") then exit function
	if (l="VPO115A_H") then exit function
	if (l="VPO116A_V") then exit function
	if (l="VPO116A_E") then exit function
	if (l="VPO116A_C") then exit function
	if (l="VPO116A_T") then exit function
	if (l="VPO116A_O") then exit function
	if (l="VPO116A_M") then exit function
	if (l="VPO116A_H") then exit function
	if (l="VPO116B_V") then exit function
	if (l="VPO116B_E") then exit function
	if (l="VPO116B_C") then exit function
	if (l="VPO116B_T") then exit function
	if (l="VPO116B_O") then exit function
	if (l="VPO116B_M") then exit function
	if (l="VPO116B_H") then exit function
	if (l="VPO117A") then exit function
	if (l="VPO117B") then exit function
	if (l="VPO117C") then exit function
	if (l="VPO117D") then exit function
	if (l="VPO118A") then exit function
	if (l="VPO118B") then exit function
	if (l="VPO125_V") then exit function
	if (l="VPO125_E") then exit function
	if (l="VPO125_C") then exit function
	if (l="VPO125_T") then exit function
	if (l="VPO125_O") then exit function
	if (l="VPO125_M") then exit function
	if (l="VPO125_H") then exit function
	if (l="VPO126_V") then exit function
	if (l="VPO126_E") then exit function
	if (l="VPO126_C") then exit function
	if (l="VPO126_T") then exit function
	if (l="VPO126_O") then exit function
	if (l="VPO126_M") then exit function
	if (l="VPO126_H") then exit function
	if (l="VPR28AA_V") then exit function
	if (l="VPR28AA_E") then exit function
	if (l="VPR28AA_C") then exit function
	if (l="VPR28AA_T") then exit function
	if (l="VPR28AA_O") then exit function
	if (l="VPR28AA_M") then exit function
	if (l="VPR28AA_H") then exit function
	if (l="VPR28BB_V") then exit function
	if (l="VPR28BB_E") then exit function
	if (l="VPR28BB_C") then exit function
	if (l="VPR28BB_T") then exit function
	if (l="VPR28BB_O") then exit function
	if (l="VPR28BB_M") then exit function
	if (l="VPR28BB_H") then exit function
	if (l="VPR59A_V") then exit function
	if (l="VPR59A_E") then exit function
	if (l="VPR59A_C") then exit function
	if (l="VPR59A_T") then exit function
	if (l="VPR59A_O") then exit function
	if (l="VPR59A_M") then exit function
	if (l="VPR59A_H") then exit function
	if (l="VPR59B_V") then exit function
	if (l="VPR59B_E") then exit function
	if (l="VPR59B_C") then exit function
	if (l="VPR59B_T") then exit function
	if (l="VPR59B_O") then exit function
	if (l="VPR59B_M") then exit function
	if (l="VPR59B_H") then exit function
	if (l="VPR59C_V") then exit function
	if (l="VPR59C_E") then exit function
	if (l="VPR59C_C") then exit function
	if (l="VPR59C_T") then exit function
	if (l="VPR59C_O") then exit function
	if (l="VPR59C_M") then exit function
	if (l="VPR59C_H") then exit function
	if (l="VPR59D_V") then exit function
	if (l="VPR59D_E") then exit function
	if (l="VPR59D_C") then exit function
	if (l="VPR59D_T") then exit function
	if (l="VPR59D_O") then exit function
	if (l="VPR59D_M") then exit function
	if (l="VPR59D_H") then exit function
	if (l="VPR70E_V") then exit function
	if (l="VPR70E_E") then exit function
	if (l="VPR70E_C") then exit function
	if (l="VPR70E_T") then exit function
	if (l="VPR70E_O") then exit function
	if (l="VPR70E_M") then exit function
	if (l="VPR70E_H") then exit function
	if (l="VPR70F_V") then exit function
	if (l="VPR70F_E") then exit function
	if (l="VPR70F_C") then exit function
	if (l="VPR70F_T") then exit function
	if (l="VPR70F_O") then exit function
	if (l="VPR70F_M") then exit function
	if (l="VPR70F_H") then exit function
	if (l="VPR113A") then exit function
	if (l="VPR113B") then exit function
	if (l="VPR113C") then exit function
	if (l="VPR113D") then exit function
	if (l="VPR124A_V") then exit function
	if (l="VPR124A_E") then exit function
	if (l="VPR124A_C") then exit function
	if (l="VPR124A_T") then exit function
	if (l="VPR124A_O") then exit function
	if (l="VPR124A_M") then exit function
	if (l="VPR124A_H") then exit function
	if (l="VPR124B_V") then exit function
	if (l="VPR124B_E") then exit function
	if (l="VPR124B_C") then exit function
	if (l="VPR124B_T") then exit function
	if (l="VPR124B_O") then exit function
	if (l="VPR124B_M") then exit function
	if (l="VPR124B_H") then exit function
	if (l="VPR124C_V") then exit function
	if (l="VPR124C_E") then exit function
	if (l="VPR124C_C") then exit function
	if (l="VPR124C_T") then exit function
	if (l="VPR124C_O") then exit function
	if (l="VPR124C_M") then exit function
	if (l="VPR124C_H") then exit function
	if (l="VPR124D_V") then exit function
	if (l="VPR124D_E") then exit function
	if (l="VPR124D_C") then exit function
	if (l="VPR124D_T") then exit function
	if (l="VPR124D_O") then exit function
	if (l="VPR124D_M") then exit function
	if (l="VPR124D_H") then exit function
	if (l="VPR128A") then exit function
	if (l="VPR128B") then exit function
	if (l="VPR128C") then exit function
	if (l="VPR128D") then exit function
	if (l="VPR128E") then exit function
	if (l="VPR128F") then exit function
	if (l="VPR128G") then exit function
	if (l="VPR128H") then exit function
	if (l="VPR128I") then exit function
	if (l="VPR136A_V") then exit function
	if (l="VPR136A_E") then exit function
	if (l="VPR136A_C") then exit function
	if (l="VPR136A_T") then exit function
	if (l="VPR136A_O") then exit function
	if (l="VPR136A_M") then exit function
	if (l="VPR136A_H") then exit function
	if (l="VPR136B_V") then exit function
	if (l="VPR136B_E") then exit function
	if (l="VPR136B_C") then exit function
	if (l="VPR136B_T") then exit function
	if (l="VPR136B_O") then exit function
	if (l="VPR136B_M") then exit function
	if (l="VPR136B_H") then exit function
	if (l="VPR136C_V") then exit function
	if (l="VPR136C_E") then exit function
	if (l="VPR136C_C") then exit function
	if (l="VPR136C_T") then exit function
	if (l="VPR136C_O") then exit function
	if (l="VPR136C_M") then exit function
	if (l="VPR136C_H") then exit function
	if (l="VPR137A_V") then exit function
	if (l="VPR137A_E") then exit function
	if (l="VPR137A_C") then exit function
	if (l="VPR137A_T") then exit function
	if (l="VPR137A_O") then exit function
	if (l="VPR137A_M") then exit function
	if (l="VPR137A_H") then exit function
	if (l="VPR137B_V") then exit function
	if (l="VPR137B_E") then exit function
	if (l="VPR137B_C") then exit function
	if (l="VPR137B_T") then exit function
	if (l="VPR137B_O") then exit function
	if (l="VPR137B_M") then exit function
	if (l="VPR137B_H") then exit function
	if (l="VPR137C_V") then exit function
	if (l="VPR137C_E") then exit function
	if (l="VPR137C_C") then exit function
	if (l="VPR137C_T") then exit function
	if (l="VPR137C_O") then exit function
	if (l="VPR137C_M") then exit function
	if (l="VPR137C_H") then exit function
	if (l="VPR138A_V") then exit function
	if (l="VPR138A_E") then exit function
	if (l="VPR138A_C") then exit function
	if (l="VPR138A_T") then exit function
	if (l="VPR138A_O") then exit function
	if (l="VPR138A_M") then exit function
	if (l="VPR138A_H") then exit function
	if (l="VPR138B_V") then exit function
	if (l="VPR138B_E") then exit function
	if (l="VPR138B_C") then exit function
	if (l="VPR138B_T") then exit function
	if (l="VPR138B_O") then exit function
	if (l="VPR138B_M") then exit function
	if (l="VPR138B_H") then exit function
	if (l="VPR138C_V") then exit function
	if (l="VPR138C_E") then exit function
	if (l="VPR138C_C") then exit function
	if (l="VPR138C_T") then exit function
	if (l="VPR138C_O") then exit function
	if (l="VPR138C_M") then exit function
	if (l="VPR138C_H") then exit function
	if (l="VPR140") then exit function
	if (l="VPR141_V") then exit function
	if (l="VPR141_E") then exit function
	if (l="VPR141_C") then exit function
	if (l="VPR141_T") then exit function
	if (l="VPR141_O") then exit function
	if (l="VPR141_M") then exit function
	if (l="VPR141_H") then exit function
	if (l="VPR141A_V") then exit function
	if (l="VPR141A_E") then exit function
	if (l="VPR141A_C") then exit function
	if (l="VPR141A_T") then exit function
	if (l="VPR141A_O") then exit function
	if (l="VPR141A_M") then exit function
	if (l="VPR141A_H") then exit function
	if (l="VPR142_V") then exit function
	if (l="VPR142_E") then exit function
	if (l="VPR142_C") then exit function
	if (l="VPR142_T") then exit function
	if (l="VPR142_O") then exit function
	if (l="VPR142_M") then exit function
	if (l="VPR142_H") then exit function
	if (l="VPR142A_V") then exit function
	if (l="VPR142A_E") then exit function
	if (l="VPR142A_C") then exit function
	if (l="VPR142A_T") then exit function
	if (l="VPR142A_O") then exit function
	if (l="VPR142A_M") then exit function
	if (l="VPR142A_H") then exit function
	if (l="VPR143A") then exit function
	if (l="VPR143B") then exit function
	if (l="VPI35A") then exit function
	if (l="VPI35B") then exit function
	if (l="VPI68A_V") then exit function
	if (l="VPI68A_E") then exit function
	if (l="VPI68A_C") then exit function
	if (l="VPI68A_T") then exit function
	if (l="VPI68A_O") then exit function
	if (l="VPI68A_M") then exit function
	if (l="VPI68A_H") then exit function
	if (l="VPI68B_V") then exit function
	if (l="VPI68B_E") then exit function
	if (l="VPI68B_C") then exit function
	if (l="VPI68B_T") then exit function
	if (l="VPI68B_O") then exit function
	if (l="VPI68B_M") then exit function
	if (l="VPI68B_H") then exit function
	if (l="VPI75AA_V") then exit function
	if (l="VPI75AA_E") then exit function
	if (l="VPI75AA_C") then exit function
	if (l="VPI75AA_T") then exit function
	if (l="VPI75AA_O") then exit function
	if (l="VPI75AA_M") then exit function
	if (l="VPI75AA_H") then exit function
	if (l="VPI75BB_V") then exit function
	if (l="VPI75BB_E") then exit function
	if (l="VPI75BB_C") then exit function
	if (l="VPI75BB_T") then exit function
	if (l="VPI75BB_O") then exit function
	if (l="VPI75BB_M") then exit function
	if (l="VPI75BB_H") then exit function
	if (l="VPI78A_V") then exit function
	if (l="VPI78A_E") then exit function
	if (l="VPI78A_C") then exit function
	if (l="VPI78A_T") then exit function
	if (l="VPI78A_O") then exit function
	if (l="VPI78A_M") then exit function
	if (l="VPI78A_H") then exit function
	if (l="VPI78B_V") then exit function
	if (l="VPI78B_E") then exit function
	if (l="VPI78B_C") then exit function
	if (l="VPI78B_T") then exit function
	if (l="VPI78B_O") then exit function
	if (l="VPI78B_M") then exit function
	if (l="VPI78B_H") then exit function
	if (l="VPI130A_V") then exit function
	if (l="VPI130A_E") then exit function
	if (l="VPI130A_C") then exit function
	if (l="VPI130A_T") then exit function
	if (l="VPI130A_O") then exit function
	if (l="VPI130A_M") then exit function
	if (l="VPI130A_H") then exit function
	if (l="VPI130B_V") then exit function
	if (l="VPI130B_E") then exit function
	if (l="VPI130B_C") then exit function
	if (l="VPI130B_T") then exit function
	if (l="VPI130B_O") then exit function
	if (l="VPI130B_M") then exit function
	if (l="VPI130B_H") then exit function
	if (l="VPI135_V") then exit function
	if (l="VPI135_E") then exit function
	if (l="VPI135_C") then exit function
	if (l="VPI135_T") then exit function
	if (l="VPI135_O") then exit function
	if (l="VPI135_M") then exit function
	if (l="VPI135_H") then exit function
	if (l="VPI136_V") then exit function
	if (l="VPI136_E") then exit function
	if (l="VPI136_C") then exit function
	if (l="VPI136_T") then exit function
	if (l="VPI136_O") then exit function
	if (l="VPI136_M") then exit function
	if (l="VPI136_H") then exit function
	if (l="VPI137_V") then exit function
	if (l="VPI137_E") then exit function
	if (l="VPI137_C") then exit function
	if (l="VPI137_T") then exit function
	if (l="VPI137_O") then exit function
	if (l="VPI137_M") then exit function
	if (l="VPI137_H") then exit function
	if (l="VPI137A_V") then exit function
	if (l="VPI137A_E") then exit function
	if (l="VPI137A_C") then exit function
	if (l="VPI137A_T") then exit function
	if (l="VPI137A_O") then exit function
	if (l="VPI137A_M") then exit function
	if (l="VPI137A_H") then exit function
	if (l="VPI138_V") then exit function
	if (l="VPI138_E") then exit function
	if (l="VPI138_C") then exit function
	if (l="VPI138_T") then exit function
	if (l="VPI138_O") then exit function
	if (l="VPI138_M") then exit function
	if (l="VPI138_H") then exit function
	if (l="VPI138A_V") then exit function
	if (l="VPI138A_E") then exit function
	if (l="VPI138A_C") then exit function
	if (l="VPI138A_T") then exit function
	if (l="VPI138A_O") then exit function
	if (l="VPI138A_M") then exit function
	if (l="VPI138A_H") then exit function
	if (l="VPMR36F_V") then exit function
	if (l="VPMR36F_E") then exit function
	if (l="VPMR36F_C") then exit function
	if (l="VPMR36F_T") then exit function
	if (l="VPMR36F_O") then exit function
	if (l="VPMR36F_M") then exit function
	if (l="VPMR36F_H") then exit function
	if (l="VPMR36G_V") then exit function
	if (l="VPMR36G_E") then exit function
	if (l="VPMR36G_C") then exit function
	if (l="VPMR36G_T") then exit function
	if (l="VPMR36G_O") then exit function
	if (l="VPMR36G_M") then exit function
	if (l="VPMR36G_H") then exit function
	if (l="VPMR78A_V") then exit function
	if (l="VPMR78A_E") then exit function
	if (l="VPMR78A_C") then exit function
	if (l="VPMR78A_T") then exit function
	if (l="VPMR78A_O") then exit function
	if (l="VPMR78A_M") then exit function
	if (l="VPMR78A_H") then exit function
	if (l="VPMR78B_V") then exit function
	if (l="VPMR78B_E") then exit function
	if (l="VPMR78B_C") then exit function
	if (l="VPMR78B_T") then exit function
	if (l="VPMR78B_O") then exit function
	if (l="VPMR78B_M") then exit function
	if (l="VPMR78B_H") then exit function
	if (l="VPMR80A_V") then exit function
	if (l="VPMR80A_E") then exit function
	if (l="VPMR80A_C") then exit function
	if (l="VPMR80A_T") then exit function
	if (l="VPMR80A_O") then exit function
	if (l="VPMR80A_M") then exit function
	if (l="VPMR80A_H") then exit function
	if (l="VPMR80B_V") then exit function
	if (l="VPMR80B_E") then exit function
	if (l="VPMR80B_C") then exit function
	if (l="VPMR80B_T") then exit function
	if (l="VPMR80B_O") then exit function
	if (l="VPMR80B_M") then exit function
	if (l="VPMR80B_H") then exit function
	if (l="VPMRA118_V") then exit function
	if (l="VPMRA118_E") then exit function
	if (l="VPMRA118_C") then exit function
	if (l="VPMRA118_T") then exit function
	if (l="VPMRA118_O") then exit function
	if (l="VPMRA118_M") then exit function
	if (l="VPMRA118_H") then exit function
	if (l="VPMR121F_V") then exit function
	if (l="VPMR121F_E") then exit function
	if (l="VPMR121F_C") then exit function
	if (l="VPMR121F_T") then exit function
	if (l="VPMR121F_O") then exit function
	if (l="VPMR121F_M") then exit function
	if (l="VPMR121F_H") then exit function
	if (l="VPMR121G_V") then exit function
	if (l="VPMR121G_E") then exit function
	if (l="VPMR121G_C") then exit function
	if (l="VPMR121G_T") then exit function
	if (l="VPMR121G_O") then exit function
	if (l="VPMR121G_M") then exit function
	if (l="VPMR121G_H") then exit function
	if (l="VPMR121H_V") then exit function
	if (l="VPMR121H_E") then exit function
	if (l="VPMR121H_C") then exit function
	if (l="VPMR121H_T") then exit function
	if (l="VPMR121H_O") then exit function
	if (l="VPMR121H_M") then exit function
	if (l="VPMR121H_H") then exit function
	if (l="VPMR121I_V") then exit function
	if (l="VPMR121I_E") then exit function
	if (l="VPMR121I_C") then exit function
	if (l="VPMR121I_T") then exit function
	if (l="VPMR121I_O") then exit function
	if (l="VPMR121I_M") then exit function
	if (l="VPMR121I_H") then exit function
	if (l="VPMR121J_V") then exit function
	if (l="VPMR121J_E") then exit function
	if (l="VPMR121J_C") then exit function
	if (l="VPMR121J_T") then exit function
	if (l="VPMR121J_O") then exit function
	if (l="VPMR121J_M") then exit function
	if (l="VPMR121J_H") then exit function
	if (l="VPMR121K_V") then exit function
	if (l="VPMR121K_E") then exit function
	if (l="VPMR121K_C") then exit function
	if (l="VPMR121K_T") then exit function
	if (l="VPMR121K_O") then exit function
	if (l="VPMR121K_M") then exit function
	if (l="VPMR121K_H") then exit function
	if (l="VPMR122") then exit function
	if (l="VPMR124_V") then exit function
	if (l="VPMR124_E") then exit function
	if (l="VPMR124_C") then exit function
	if (l="VPMR124_T") then exit function
	if (l="VPMR124_O") then exit function
	if (l="VPMR124_M") then exit function
	if (l="VPMR124_H") then exit function
	if (l="VPMR124A_V") then exit function
	if (l="VPMR124A_E") then exit function
	if (l="VPMR124A_C") then exit function
	if (l="VPMR124A_T") then exit function
	if (l="VPMR124A_O") then exit function
	if (l="VPMR124A_M") then exit function
	if (l="VPMR124A_H") then exit function
	if (l="VPMR125_V") then exit function
	if (l="VPMR125_E") then exit function
	if (l="VPMR125_C") then exit function
	if (l="VPMR125_T") then exit function
	if (l="VPMR125_O") then exit function
	if (l="VPMR125_M") then exit function
	if (l="VPMR125_H") then exit function
	if (l="VPMR125A_V") then exit function
	if (l="VPMR125A_E") then exit function
	if (l="VPMR125A_C") then exit function
	if (l="VPMR125A_T") then exit function
	if (l="VPMR125A_O") then exit function
	if (l="VPMR125A_M") then exit function
	if (l="VPMR125A_H") then exit function
	if (l="VPMR126A1") then exit function
	if (l="VPMR126A2") then exit function
	if (l="VPMR126A3") then exit function
	if (l="VPMR126A4") then exit function
	if (l="VPMR126B") then exit function
	if (l="VPMR128HH") then exit function
	if (l="VPMR129EE") then exit function
	if (l="VPMR139A") then exit function
	if (l="VPMR139B") then exit function
	if (l="VPMR139C") then exit function
	if (l="VPMR139D") then exit function
	if (l="VPMR139E") then exit function
	if (l="VPMR140B1") then exit function
	if (l="VPMR140B2") then exit function
	if (l="VPMR140B3") then exit function
	if (l="VPMR140S1") then exit function
	if (l="VPMR140S2") then exit function
	if (l="VPMR140S3") then exit function
	if (l="CSCM43AA") then exit function
	if (l="CSCM43BB") then exit function
	if (l="CSCM44A") then exit function
	if (l="CSCM44B") then exit function
	if (l="CSCM45") then exit function
	if (l="CSCM46A") then exit function
	if (l="CSCM46B") then exit function
	if (l="CSCM47A_C") then exit function
	if (l="CSCM48AA_R") then exit function
	if (l="CSCM48AA_C") then exit function
	if (l="CSCM48BB_R") then exit function
	if (l="CSCM48BB_C") then exit function
	if (l="CSCM49") then exit function
	if (l="CSDM84") then exit function
	if (l="CSDM85A") then exit function
	if (l="CSDM85B") then exit function
	if (l="CSCM94A") then exit function
	if (l="CSCM94B") then exit function
	if (l="MISC8") then exit function
	if (l="MISCA9") then exit function
	if (l="MISC12") then exit function
	if (l="MISC14A_V") then exit function
	if (l="MISC14A_E") then exit function
	if (l="MISC14A_C") then exit function
	if (l="MISC14A_T") then exit function
	if (l="MISC14A_O") then exit function
	if (l="MISC14A_M") then exit function
	if (l="MISC14A_H") then exit function
	if (l="MISC14B_V") then exit function
	if (l="MISC14B_E") then exit function
	if (l="MISC14B_C") then exit function
	if (l="MISC14B_T") then exit function
	if (l="MISC14B_O") then exit function
	if (l="MISC14B_M") then exit function
	if (l="MISC14B_H") then exit function
	if (l="MISC14C_V") then exit function
	if (l="MISC14C_E") then exit function
	if (l="MISC14C_C") then exit function
	if (l="MISC14C_T") then exit function
	if (l="MISC14C_O") then exit function
	if (l="MISC14C_M") then exit function
	if (l="MISC14C_H") then exit function
	if (l="MISC14D_V") then exit function
	if (l="MISC14D_E") then exit function
	if (l="MISC14D_C") then exit function
	if (l="MISC14D_T") then exit function
	if (l="MISC14D_O") then exit function
	if (l="MISC14D_M") then exit function
	if (l="MISC14D_H") then exit function
	if (l="MISC14E_V") then exit function
	if (l="MISC14E_E") then exit function
	if (l="MISC14E_C") then exit function
	if (l="MISC14E_T") then exit function
	if (l="MISC14E_O") then exit function
	if (l="MISC14E_M") then exit function
	if (l="MISC14E_H") then exit function
	if (l="MISC15A_V") then exit function
	if (l="MISC15A_E") then exit function
	if (l="MISC15A_C") then exit function
	if (l="MISC15A_T") then exit function
	if (l="MISC15A_O") then exit function
	if (l="MISC15A_M") then exit function
	if (l="MISC15A_H") then exit function
	if (l="MISC15B_V") then exit function
	if (l="MISC15B_E") then exit function
	if (l="MISC15B_C") then exit function
	if (l="MISC15B_T") then exit function
	if (l="MISC15B_O") then exit function
	if (l="MISC15B_M") then exit function
	if (l="MISC15B_H") then exit function
	if (l="MISC15C_V") then exit function
	if (l="MISC15C_E") then exit function
	if (l="MISC15C_C") then exit function
	if (l="MISC15C_T") then exit function
	if (l="MISC15C_O") then exit function
	if (l="MISC15C_M") then exit function
	if (l="MISC15C_H") then exit function
	if (l="MISC15D_V") then exit function
	if (l="MISC15D_E") then exit function
	if (l="MISC15D_C") then exit function
	if (l="MISC15D_T") then exit function
	if (l="MISC15D_O") then exit function
	if (l="MISC15D_M") then exit function
	if (l="MISC15D_H") then exit function
	if (l="MISC16A") then exit function
	if (l="MISC16B") then exit function
	if (l="MISC16C") then exit function
	if (l="MISC16D") then exit function
	if (l="MISC16E") then exit function
	if (l="MISC19A_V") then exit function
	if (l="MISC19A_E") then exit function
	if (l="MISC19A_C") then exit function
	if (l="MISC19A_T") then exit function
	if (l="MISC19A_O") then exit function
	if (l="MISC19A_M") then exit function
	if (l="MISC19A_H") then exit function
	if (l="MISC19B_V") then exit function
	if (l="MISC19B_E") then exit function
	if (l="MISC19B_C") then exit function
	if (l="MISC19B_T") then exit function
	if (l="MISC19B_O") then exit function
	if (l="MISC19B_M") then exit function
	if (l="MISC19B_H") then exit function
	if (l="MISC19C_V") then exit function
	if (l="MISC19C_E") then exit function
	if (l="MISC19C_C") then exit function
	if (l="MISC19C_T") then exit function
	if (l="MISC19C_O") then exit function
	if (l="MISC19C_M") then exit function
	if (l="MISC19C_H") then exit function
	if (l="MISC20") then exit function
	if (l="MISC92A_L") then exit function
	if (l="MISC92A_B") then exit function
	if (l="MISC92B_L") then exit function
	if (l="MISC92B_B") then exit function
	if (l="MISC92C_L") then exit function
	if (l="MISC92C_B") then exit function
	if (l="MISC92D_L") then exit function
	if (l="MISC92D_B") then exit function
	if (l="MISC92E_L") then exit function
	if (l="MISC92E_B") then exit function
	if (l="MISC97A_I1") then exit function
	if (l="MISC97A_R1") then exit function
	if (l="MISC97A_S1") then exit function
	if (l="MISC97A_F1") then exit function
	if (l="MISC97A_I2") then exit function
	if (l="MISC97A_R2") then exit function
	if (l="MISC97A_S2") then exit function
	if (l="MISC97A_F2") then exit function
	if (l="MISC97B_I1") then exit function
	if (l="MISC97B_R1") then exit function
	if (l="MISC97B_S1") then exit function
	if (l="MISC97B_F1") then exit function
	if (l="MISC97B_I2") then exit function
	if (l="MISC97B_R2") then exit function
	if (l="MISC97B_S2") then exit function
	if (l="MISC97B_F2") then exit function
	if (l="MISC109") then exit function

	if (l="RESPID") then exit function
	if (l="STATUS") then exit function
	if (l="DATE_END") then exit function
	if (l="Q423") then exit function
	if (l="Q425") then exit function
	if (l="CONTRIBU") then exit function
	if (l="CONTR0") then exit function
	if (l="CONTR1") then exit function
	if (l="CONTR2") then exit function
	if (l="CONTR3") then exit function
	if (l="CONTR4") then exit function
	if (l="CONTR5") then exit function
	if (l="CONTR6") then exit function
	if (l="CONTR7") then exit function
	if (l="COMPA0") then exit function
	if (l="COMPA1") then exit function
	if (l="COMPA2") then exit function
	if (l="COMPA3") then exit function
	if (l="COMPA4") then exit function
	if (l="COMPA5") then exit function
	if (l="COMPA6") then exit function
	if (l="COMPA7") then exit function
	if (l="PHONE0") then exit function
	if (l="PHONE1") then exit function
	if (l="PHONE2") then exit function
	if (l="PHONE3") then exit function
	if (l="PHONE4") then exit function
	if (l="PHONE5") then exit function
	if (l="PHONE6") then exit function
	if (l="PHONE7") then exit function
	if (l="SURVE0") then exit function
	if (l="SURVE1") then exit function
	if (l="SURVE2") then exit function
	if (l="SURVE3") then exit function
	if (l="SURVE4") then exit function
	if (l="SURVE5") then exit function
	if (l="SURVE6") then exit function
	if (l="SURVE7") then exit function
	if (l="ASSET_1") then exit function
	if (l="ASSET_2") then exit function
	if (l="ASSET_3") then exit function
	if (l="ASSET_4") then exit function
	if (l="ASSET_7") then exit function
	if (l="CITY_1") then exit function
	if (l="CITY_2") then exit function
	if (l="CITY_3") then exit function
	if (l="CITY_4") then exit function
	if (l="CITY_5") then exit function
	if (l="CITY_6") then exit function
	if (l="CITY_7") then exit function
	if (l="CITY_11") then exit function
	if (l="CITY_16") then exit function

	if (l="MISC11_B") then exit function
	if (l="MISC93") then exit function
	if (l="MISC93_B") then exit function


	isIgnoredQuestion=false

End Function





' SQL Queries to analyze the resulting data
'
' SELECT b.*, a.Label FROM dat_ResponseItemTxt b JOIN dat_Question a ON b.QuestionID = a.QuestionID WHERE ResponseID=63
'
'
' Delete all data associated with a single period (e.g. period 110)
'
' DELETE FROM dat_ResponseItemInt WHERE ResponseID IN (SELECT ResponseID FROM dat_Response WHERE PeriodID = 110)
' DELETE FROM dat_ResponseItemFloat WHERE ResponseID IN (SELECT ResponseID FROM dat_Response WHERE PeriodID = 110)
' DELETE FROM dat_ResponseItemTxt WHERE ResponseID IN (SELECT ResponseID FROM dat_Response WHERE PeriodID = 110)
' DELETE FROM dat_Response WHERE PeriodID = 110



' 145 = VPO99E_T - Altus Insite Investor Outlook Chart

' return data as columns (13, 9, 23)
' declare @v1 int, @v2 int, @v3 int
' select @v1=count(*) from dat_ResponseItemInt where QuestionID=145 and Value=1
' select @v2=count(*) from dat_ResponseItemInt where QuestionID=145 and Value=2
' select @v3=count(*) from dat_ResponseItemInt where QuestionID=145 and Value=3
' select @v1 as v1, @v2 as v2, @v3 as v3

' return data as rows (13, 9, 23)
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID=145 and Value=1 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID=145 and Value=2 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID=145 and Value=3


' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=1 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=2 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=3






' Step 1:
'
' select * from dat_question where label like '%VPO99E_%'
'
' Results:
'
' 139	VPO99E_V  = Vancouver
' 141	VPO99E_E  = Edmonton
' 143	VPO99E_C  = Calgary
' 145	VPO99E_T  = Toronto
' 147	VPO99E_O  = Ottawa
' 149	VPO99E_M  = Montreal
' 151	VPO99E_Q  = Quebec City
'
' 153	VPO99E_F  = Fredericton
' 155	VPO99E_N  = Moncton
' 159	VPO99E_R  = Charlottetown
' 161	VPO99E_H  = Halifax
' 164	VPO99E_S  = St. John's
' 167	VPO99E_J  = Saint John
'
'
' Step 2:
'
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=1 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=2 union all
' select count(*) AS v1 from dat_ResponseItemInt where QuestionID in (139, 141, 143, 145, 147, 149, 151) and Value=3
'
'
' Results:
'
' 57 = 27%
' 37 = 17.5%
' 117 = 55.5%
'
' Total: 211 = 100%
'



' select * from dat_question where label like '%VPO99%'

' select * from dat_ResponseItemInt where QuestionID in (144) order by Value

' select avg(Value) from dat_ResponseItemInt where QuestionID in (144)


%>
