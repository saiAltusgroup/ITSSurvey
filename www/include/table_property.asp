<%

' DK, Apr 01, 2008 - version 1.0

Class cTableProperty

	Public TableObj
	Public CurrPeriodID
	Public EnableShortNames
	Public Flag_Q1234

	Private ChartFilter
	Private Filter

	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid

		CurrPeriodID=g_CURRENT_PERIOD_ID
		EnableShortNames=true

		Flag_Q1234=false

		ProductID=0
		ChartID=0
	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub


	Public Property Let ProductID(ID)
		Filter=""
		if ID>0 Then Filter="WHERE ProductID=" & CLng(ID)
	End Property


	Public Property Let ChartID(ID)
		ChartFilter=""
		if ID>0 Then ChartFilter=" AND ChartID=" & CLng(ID)
	End Property



	Public Function Read()
		Read=ReadRs(TableObj, Filter, true)
	End Function


	Public Function ReadNames(p_arr, display_mode)

		'Response.Write "p_arr=" & p_arr & "<br>"
		'Response.Write "EnableShortNames=" & EnableShortNames & "<br>"

		ReadNames=""

		if EnableShortNames then

			tmp=Split(p_arr, ",")
			for i=0 to Ubound(tmp)
				tmp(i)=CLng(tmp(i))
			next
			Call Sort(tmp, tmp, Ubound(tmp)+1)
			arr=Join(tmp, ",")

			'Response.Write arr & "<br>"

			'if arr="1,2,3,4,5,6,17" then ReadNames=Lang("all_available_office")
			'if arr="7,8,9,10,11,18" then ReadNames=Lang("all_available_retail")
			'if arr="12,13,14" then ReadNames=Lang("all_available_industrial")
			'if arr="" or arr="0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18" then ReadNames=Lang("all_available")

			if ubound(tmp)>0 then

				MatchOffice=0
				MatchRetail=0
				MatchIndustrial=0
				MatchAll=0

				for i=0 to ubound(tmp)

					if InArray(Array(1,2,3,4,5,6,17), tmp(i)) then MatchOffice=MatchOffice+1
					if InArray(Array(7,8,9,10,11,18), tmp(i)) then MatchRetail=MatchRetail+1
					if InArray(Array(12,13,14), tmp(i)) then MatchIndustrial=MatchIndustrial+1
					if InArray(Array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18), tmp(i)) then MatchAll=MatchAll+1

				next

				count_property=ubound(tmp)+1

				ReadNames=""
				if count_property=MatchOffice then
					ReadNames=Lang("all_available_office")
				elseif count_property=MatchRetail then
					ReadNames=Lang("all_available_retail")
				elseif count_property=MatchIndustrial then
					ReadNames=Lang("all_available_industrial")
				elseif count_property=MatchAll then
					ReadNames=Lang("all_available")
				end if

			end if

		end if

		if Len(ReadNames)>0 then exit function

		str_filter=""
		if Len(p_arr)>0 then str_filter="WHERE PropertyID IN (" & p_arr & ")"

		'Response.Write "str_filter=" & str_filter & "<br>"

		Set tmpObj=new cTablePlaintext
		tmpObj.DisplayMode=display_mode
		ReadNames=ReadRs(tmpObj, str_filter, false)

		'Response.Write "ReadNames=" & ReadNames & "<br>"

		Set tmpObj = nothing

	End Function


	Private Function ReadRs(table_obj, str_filter, check_period)

		strSQL="SELECT PropertyID, Name FROM dat_Property " & str_filter & " ORDER BY Name"
		'Response.Write strSQL & "<br>"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, its_conn

		ReadRs = table_obj.TableOpen(0)

		do while not rs.eof
			table_obj.RowID=rs("PropertyID")

			exist_in_period=true
			if check_period then exist_in_period=ExistsInPeriod(table_obj.RowID)

			'if exist_in_period then ReadRs = ReadRs & table_obj.RowOpen & table_obj.Cell(LangField(rs("Name"))) & table_obj.RowClose

			' to exclude REIT Units - requested by KA on June 12, 2015
			if exist_in_period and table_obj.RowID>0 then ReadRs = ReadRs & table_obj.RowOpen & table_obj.Cell(LangField(rs("Name"))) & table_obj.RowClose
			rs.MoveNext
		loop

		ReadRs = ReadRs & table_obj.TableClose

		rs.Close()
		Set rs=Nothing

	End Function


	Private Function ExistsInPeriod(property_id)

		'ExistsInPeriod=true
		'Exit Function

		ExistsInPeriod=false

		if Flag_Q1234 then
			period_bitvalue=15 ' 1+2+4+8=15 (Q1 and Q2 and Q3 and Q4)
		else
			Set c = new cConverter
			period_bitvalue=c.NumToLng("LKP02_PERIOD_BITVALUE", CurrPeriodID)
			Set c = nothing
		end if

		strSQL="SELECT QuestionID FROM dat_Question WHERE PropertyID=" & property_id & " AND ((QuarterBitMask & " & period_bitvalue & ")<>0)" & ChartFilter

		'Response.Write "period_bitvalue=" & period_bitvalue & "<br>"
		'Response.Write "strSQL=" & strSQL & "<br>"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, its_conn, 3, 1

		If (rs.RecordCount > 0) Then ExistsInPeriod=true

		rs.Close()
		Set rs=Nothing

	End Function

End Class

%>


<%

Function PropertySelector(output, period, selector_type, product, chart_id, flag_q1234)

	output.Name="property"

	Set t=new cTableProperty
	t.ProductID=product
	t.ChartID=chart_id
	t.CurrPeriodID=period
	t.Flag_Q1234=flag_q1234


	' factory of property selectors

	Set plaintext=new cTablePlaintext

	select case selector_type

	case 0

		t.ProductID=999
		Set t.TableObj=output
		PropertySelector=t.Read
		PropertySelector=output.PrependHack(PropertySelector, option_0_a, Lang("all"), false)

	case 1, 2

		Set t.TableObj=plaintext

		' --- all property types

		plaintext.DisplayMode=2 : option_0_a=t.Read ' just rowids
		'plaintext.DisplayMode=1 : option_0_b=t.Read ' just cells

		Set t.TableObj=output
		PropertySelector=t.Read
		if selector_type=1 then PropertySelector=output.PrependHack(PropertySelector, option_0_a, Lang("all"), true)

	case 3 ' --- choose one market from market groups 1 and 2

		Set t.TableObj=output
		PropertySelector=t.Read


	case else ' --- choose any number of markets from any market group (for super-administrators, the most flexibe option)

		output.Multi=true
		Set t.TableObj=output
		PropertySelector=t.Read

	end select

	Set plaintext=nothing

	Set t = nothing

End Function


Function PropertySelector_HTML(period, selector_type, product, chart_id, flag_q1234)

	Set output=new cTableListbox
	output.FullDisplay=false
	PropertySelector_HTML=PropertySelector(output, period, selector_type, product, chart_id, flag_q1234)
	Set output=nothing

End Function


Function PropertySelector_Ajax(period, selector_type, product, chart_id, flag_q1234)

	Set output=new cTablePlaintext
	output.DisplayMode=-1 ' json
	PropertySelector_Ajax=PropertySelector(output, period, selector_type, product, chart_id, flag_q1234)
	Set output=nothing

End Function

%>
