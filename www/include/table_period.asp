<%

' DK, Apr 01, 2008 - version 1.0

Class cTablePeriod

	Public TableObj

	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid
	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub

	Public Function Read()

		strSQL = "SELECT PeriodID, Name FROM dat_Period WHERE Status=1 ORDER BY PeriodID"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, its_conn

		Read = TableObj.TableOpen(0)

		do while not rs.eof
			TableObj.RowID=rs("PeriodID")
			Read = Read & TableObj.RowOpen & TableObj.Cell(LangField(rs("Name"))) & TableObj.RowClose
			rs.MoveNext
		loop

		Read = Read & TableObj.TableClose

		rs.Close()
		Set rs=Nothing

	End Function

End Class

%>
