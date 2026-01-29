<!--#INCLUDE virtual="/sys_include/db.asp"-->
<%
	Dim UpdateDb
	Dim Debug
	UpdateDB = false
	Debug = false

	Dim ArrayCloneModel(12,2)'settings must be correct on this one or it will blow
'values for the array
	ArrayCloneModel(0,0) = "VPO99E" 'model
	ArrayCloneModel(0,1) = Array("VPI33E","VPI75E","VPMR127E","VPMR36E","VPMR81E","VPO119E","VPO127E","VPO23E","VPO64E","VPR100D","VPR101D","VPR28D","VPR33D","VPR69D","VPR70D")
	ArrayCloneModel(1,0) = "VPO99B" 'model
	ArrayCloneModel(1,1) = Array("VPI33B","VPI75B","VPMR127B","VPMR36B","VPMR81B","VPO119B","VPO127B","VPO23B","VPO64B")
	ArrayCloneModel(2,0) = "VPO99A" 'model
	ArrayCloneModel(2,1) = Array("VPI33A","VPI75A","VPMR127A","VPMR36A","VPMR81A","VPO119A","VPO127A","VPO23A","VPO64A","VPR100A","VPR101A","VPR28A","VPR33A","VPR69A","VPR70A")
	ArrayCloneModel(3,0) = "VPO99C" 'model
	ArrayCloneModel(3,1) = Array("VPI33C","VPI75C","VPMR127C","VPMR36C","VPMR81C","VPO119C","VPO127C","VPO23C","VPO64C","VPR100B","VPR101B","VPR28B","VPR33B","VPR69B","VPR70B")
	ArrayCloneModel(4,0) = "VPO99D" 'model
	ArrayCloneModel(4,1) = Array("VPI33D","VPI75D","VPMR127D","VPMR36D","VPMR81D","VPO119D","VPO127D","VPO23D","VPO64D","VPR100C","VPR101C","VPR28C","VPR33C","VPR69C","VPR70C")
	ArrayCloneModel(5,0) = "VPO21A" 'model
	ArrayCloneModel(5,1) = Array("VPO122A","VPO134A","VPO134B","VPO65A")
	ArrayCloneModel(6,0) = "VPO21B" 'model
	ArrayCloneModel(6,1) = Array("VPO122B")
	ArrayCloneModel(7,0) = "VPO57A" 'model
	ArrayCloneModel(7,1) = Array("VPO67A")
	ArrayCloneModel(8,0) = "VPO57B" 'model
	ArrayCloneModel(8,1) = Array("VPO67B")
	ArrayCloneModel(9,0) = "VPO22A" 'model
	ArrayCloneModel(9,1) = Array("VPI129A","VPI134A","VPO121A","VPO123A","VPO135A","VPO66A","VPMR79A")
	ArrayCloneModel(10,0) = "VPO22B" 'model
	ArrayCloneModel(10,1) = Array("VPI129B","VPI134B","VPO121B","VPO123B","VPO135B","VPO66B")
	ArrayCloneModel(11,0) = "VPO22C" 'model
	ArrayCloneModel(11,1) = Array("VPI129C","VPI134C","VPO121C","VPO123C","VPO135C","VPO66C")
	ArrayCloneModel(12,0) = "VPO22D" 'model
	ArrayCloneModel(12,1) = Array("VPI129D","VPI134D","VPO121D","VPO123D","VPO135D","VPO66D","VPMR79B")


'do not modify with anything below this line
	Redim aryMessage(0)
'	Redim CloneArray(0)
	Dim Model
	Dim ModelSql
	Dim ModelRs
	Dim Clone
	Dim CloneSql
	Dim CloneRs
	Set ModelRs = Server.CreateObject("ADODB.Recordset")
	Set CloneRs = Server.CreateObject("ADODB.Recordset")
	Set objComm = Server.CreateObject("ADODB.Command")

	objComm.ActiveConnection=its_conn

	For i = 0 to UBound(ArrayCloneModel)

		Model = ArrayCloneModel(i,0)
		ReDim aryModelQuestion(0)
		Redim aryModelMarket(0)
		ModelCount = 0
		ModelSql = "SELECT QuestionID, MarketID, PropertyId, ChartID, ColumnID, Status FROM dat_Question WHERE Label like '" & Model & "%' ORDER BY MarketID"

		ModelRs.Open ModelSql, its_conn

		If(ModelRs.EOF) Then

			Call ProduceMessage("error","Model " & Model & "Contained no records.")
			ModelBad = true

		Else

			ModelBad = false

			Do While Not ModelRs.EOF

				If(ModelRs("Status") <> 1) Then
					Call ProduceMessage("error","QuestionID " & ModelRs("QuestionID") & " status does not equal 1.") 'Check Status
					ModelBad = true
				Else
					Call ProduceMessage("message","QuestionID " & ModelRs("QuestionID") & " Status equals " & ModelRs("Status"))
				End If

				If(ModelCount = 0) Then ' get first values to compare against

					ModelPropertyID = ModelRs("PropertyID")
					ModelChartID = ModelRs("ChartID")
					ModelColumnID = ModelRs("ColumnID")

				Else 'compare against first values

					If(ModelPropertyID <> ModelRs("PropertyID")) Then 'Check PropertyID for consistency
						Call ProduceMessage("error","QuestionID " & ModelRs("QuestionID") & " PropertyID does not match.")
						ModelBad = true
					Else
						Call ProduceMessage("message","QuestionID " & ModelRs("QuestionID") & " PropertyID " & ModelRs("PropertyID") & " does match " & ModelPropertyID & ".")
					End If

					If(ModelChartID <> ModelRs("ChartID")) Then 'Check ChartID for consistency
						Call ProduceMessage("error","QuestionID " & ModelRs("QuestionID") & " ChartID does not match.")
						ModelBad = true
					Else
						Call ProduceMessage("message","QuestionID " & ModelRs("QuestionID") & " ChartID " & ModelRs("ChartID") &" does match " & ModelChartID & ".")
					End If

					If(ModelColumnID <> ModelRs("ColumnID")) Then 'Check ColumnID for consistency
						Call ProduceMessage("error","QuestionID " & ModelRs("QuestionID") & " ColumnID does not match.")
						ModelBad = true
					Else
						Call ProduceMessage("message","QuestionID " & ModelRs("QuestionID") & " ColumnID " & ModelRs("ColumnID") &" does match " & ModelColumnID & ".")
					End If

				End If

				ReDim Preserve aryModelQuestion(ModelCount)
				ReDim Preserve aryModelMarket(ModelCount)
				aryModelQuestion(ModelCount) = ModelRs("QuestionID") 'preserve QuestionID's
				aryModelMarket(ModelCount) = ModelRs("MarketID") 'preserve MarketID's

				ModelCount = ModelCount + 1 'get count of records

				ModelRs.MoveNext

			Loop

'			If(Len(aryError(0)) = 0) Then

'				Call ProduceMessage("message","Model " & Model & " has no errors.")

'			Else

'				Exit For 'model has errored so we leave

'			End If


		End If

		ModelRs.Close() ' close the recordset

		If( Not ModelBad) Then

			Call ProduceMessage("message","Model " & Model & " contains " & ModelCount & " records.")

			CloneArray = ArrayCloneModel(i,1)

			For j = 0 to UBound(CloneArray)

				CloneBad = false
				Clone = CloneArray(j)
				ReDim aryCloneQuestion(0)
				Redim aryCloneMarket(0)
				CloneCount = 0
				CloneSql = "SELECT QuestionID, MarketID, PropertyId, ChartID, ColumnID, Status FROM dat_Question WHERE Label like '" & Clone & "%' ORDER BY MarketID"

				CloneRs.Open CloneSql, its_conn

				If(CloneRs.EOF) Then
					Call ProduceMessage("error","Model " & Model & "'s Clone " & Clone & " contained no records." )
					CloneBad = true
				Else

					Do While Not CloneRs.EOF

						If(CloneCount = 0) Then ' get first values to compare against

							ClonePropertyID = CloneRs("PropertyID")

						Else 'compare against first values

							If(ClonePropertyID <> CloneRs("PropertyID")) Then 'Check PropertyID for consistency
								Call ProduceMessage("error","QuestionID " & CloneRs("QuestionID") & " PropertyID does not match.")
								CloneBad = true
							Else
								Call ProduceMessage("message","QuestionID " & CloneRs("QuestionID") & " PropertyID " & CloneRs("PropertyID") &" does match " & ClonePropertyID & ".")
							End If

						End If

						ReDim Preserve aryCloneQuestion(CloneCount)
						ReDim Preserve aryCloneMarket(CloneCount)
						aryCloneQuestion(CloneCount) = CloneRs("QuestionID") 'preserve QuestionID's
						aryCloneMarket(CloneCount) = CloneRs("MarketID") 'preserve MarketID's

						CloneCount = CloneCount + 1 'get count of records

						CloneRs.MoveNext

					Loop

					If(ModelCount <> CloneCount) Then

						Call ProduceMessage("error","Model " & Model & " count " & ModelCount & " and Clone " & Clone & " count " & CloneCount & " do not match.") 'check record set counts
						CloneBad = true

					Else

						Call ProduceMessage("message","Model " & Model & " count " & ModelCount & " and Clone " & Clone & " count " & CloneCount & " match.")

						For k = 0 to UBound(aryModelMarket)

							If(aryModelMarket(k) <> aryCloneMarket(k)) Then 'check marketID's
								Call ProduceMessage("error","Model " & Model & " MarketID " & aryModelMarket(k) & " and Clone " & Clone & " MarketID " & aryCloneMarket(k) & " do not match")
								CloneBad = true
							Else
								Call ProduceMessage("message","Model " & Model & " MarketID " & aryModelMarket(k) & " and Clone " & Clone & " MarketID " & aryCloneMarket(k) & " match")
							End If

						Next

						For l = 0 to UBound(aryModelQuestion)

							If(InArray(aryCloneQuestion,aryModelQuestion(l))) Then 'check QuestionId's
								Call ProduceMessage("error","Model " & Model & " and Clone " & Clone & " both contain QuestionID " & aryModelQuestion(l) )
								CloneBad = true
							Else
								Call ProduceMessage("message","Model " & Model & "'s Clone " & Clone & " does not contain QuestionID " & aryModelQuestion(l))
							End If

						Next

					End IF


					If(Not CloneBad) Then

						For m = 0 to UBound(aryModelQuestion)

							objComm.CommandText="UPDATE dat_Question SET ChartID = " & ModelChartID & ", ColumnID = " & ModelColumnID & ", Status = 2 WHERE QuestionID = " & aryCloneQuestion(m) & ";"
							If(UpdateDB) Then
								objComm.Execute
							Else
								Response.Write objComm.CommandText & "<br />"
							End If

						Next

					End If


				End If

				If(Not CloneBad) Then
					Call ProduceMessage("message","Clone " & Clone & " contains " & CloneCount & " records.")
				End If

				CloneRs.Close() ' close the recordset

			Next

		End If

	Next



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
	<head>
		<title>Column Chart Assignment</title>
	</head>
	<style type="text/css">
		ol li.error{ color:red; }
		ol li.message{ color:green; }
	</style>
	<body>

	<h1>Column Chart Assignment</h1>


<%
		If( Len(aryMessage(0)) > 0 ) Then
			ProduceMessageList aryMessage
		End If
%>
		<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
			<input type="submit" name="submit" id="submit" value="submit" />
		</form>
	</body>
</html>
<%
'subroutines

Sub ProduceMessageList(a)
	If(Len(a(0)) > 0) Then
%>
	<ol>
<%
	for each m in a
	 ma = split(m,",")
	 if(Not Debug) Then
	  if(Ma(0) <> "message") Then
%>
		<li class="error"><%=ma(1)%></li>
<%
	  end if
	 else
%>
		<li class="<%=ma(0)%>"><%=ma(1)%></li>
<%
	 end if
	next
%>
	</ol>
<%
	End If
End Sub

%>


<%

Function ProduceMessage(t,s)
	Dim c

	c = Ubound(aryMessage)

	If(Len(aryMessage(c)) > 0) Then
		c = c + 1
		ReDim Preserve aryMessage(c)
	End If

	aryMessage(c) = t & "," & s

	ProduceMessage = aryMessage

End Function

Function InArray(a,s)
	Dim i

	For i = 0 to UBound(a)

		If cstr(a(i)) = cstr(s) Then

			InArray = true
			Exit Function

		End If

	Next

	InArray = false

End Function

%>