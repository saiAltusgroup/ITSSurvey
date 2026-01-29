<%

' DK, Apr 01, 2008 - version 1.0

Class cConverter

	' Constructor
	Private Sub Class_Initialize()
	End Sub

	' Destructor
	Private Sub Class_Terminate()
	End Sub


	Public Function NumToStr(lookup_type, lookup_id)

		' lookup factory

		select case lookup_type

		case "LKP01_PERIOD_NAME"
			table="dat_Period"
			field_id="PeriodID"
			field_val="Name"

		case "LKP02_PERIOD_BITVALUE"
			table="dat_Period"
			field_id="PeriodID"
			field_val="BitValue"

		case "LKP03_MARKET_NAME"
			table="dat_Market"
			field_id="MarketID"
			field_val="Name"

		end select

		lookup_id=clng(lookup_id)
		strSQL = "SELECT " & field_val & " FROM " & table & " WHERE " & field_id & " = " & lookup_id
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, its_conn

		NumToStr=""
		if not rs.eof then NumToStr=rs(field_val)

		rs.Close()
		Set rs=Nothing

	End Function

	Public Function NumToLng(lookup_type, lookup_id)
		NumToLng=0
		s=NumToStr(lookup_type, lookup_id)
		if isNumeric(s) then NumToLng=CLng(s)
	End Function


End Class

%>
