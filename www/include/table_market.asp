<%

' DK, Apr 01, 2008 - version 1.0

Class cTableMarket

	Public TableObj
	Public CurrPeriodID
	Public Flag_Q1234

	Private GroupFilter
	Private ChartFilter

	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid

		CurrPeriodID=g_CURRENT_PERIOD_ID

		Flag_Q1234=false

		GroupFilter=""
		ChartID=0
	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub

	Public Function SetFilterByGroup(group_id_arr)
		GroupFilter=""
		if group_id>=0 then GroupFilter=" WHERE GroupID IN (" & group_id_arr & ") "
	End Function

	Public Property Let ChartID(ID)
		ChartFilter=""
		if ID>0 Then ChartFilter=" AND ChartID=" & CLng(ID)
	End Property

	Public Function Read()

		strSQL = "SELECT MarketID, GroupID, Name FROM dat_Market " & GroupFilter & " ORDER BY GroupID, Name"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, its_conn

		'Response.Write strSQL & "<br>"
		'Response.Write ChartFilter & "<br>"

		Read = TableObj.TableOpen(0)

		do while not rs.eof
			TableObj.RowID=rs("MarketID")

			'Response.Write TableObj.RowID & "<br>"

			'Read = Read & TableObj.RowOpen & TableObj.Cell(LangField(rs("Name"))) & TableObj.Cell(rs("GroupID")) & TableObj.RowClose
			if ExistsInPeriod(TableObj.RowID) then Read = Read & TableObj.RowOpen & TableObj.Cell(LangField(rs("Name"))) & TableObj.RowClose
			rs.MoveNext
		loop

		Read = Read & TableObj.TableClose

		rs.Close()
		Set rs=Nothing

	End Function


	Private Function ExistsInPeriod(market_id)

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

		strSQL="SELECT QuestionID FROM dat_Question WHERE MarketID=" & market_id & " AND ((QuarterBitMask & " & period_bitvalue & ")<>0)" & ChartFilter

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

Function MarketSelector(output, period, selector_type, chart_id, flag_q1234)

	output.Name="market"

	Set t=new cTableMarket
	t.ChartID=chart_id
	t.CurrPeriodID=period
	t.Flag_Q1234=flag_q1234

	' factory of market selectors

	Set plaintext=new cTablePlaintext

	select case selector_type

	case 1, 2

		Set t.TableObj=plaintext


		' --- entire market group 0 and 1 (0 = REIT, 1=National)

		't.SetFilterByGroup("0,1") '
		t.SetFilterByGroup("1")
		plaintext.DisplayMode=2 : option_0_a=t.Read ' just rowids
		'plaintext.DisplayMode=1 : option_0_b=t.Read ' just cells

		' or

		' --- entire market group 2 (2=Atlantic)

		t.SetFilterByGroup("2")
		plaintext.DisplayMode=2 : option_1_a=t.Read ' just rowids
		'plaintext.DisplayMode=1 : option_1_b=t.Read ' just cells


		'Response.Write "option_0_a=" & option_0_a & "<br>"
		'Response.Write "option_1_a=" & option_1_a & "<br>"


		Set t.TableObj=output
		if selector_type=2 then
			t.SetFilterByGroup("1,2")
			check_flag=true
		else
			t.SetFilterByGroup("-1")
			check_flag=false
		end if

		MarketSelector=t.Read

		' Note 1: starting with Q2 2008 (period=105) no more data is collected for Atlantic markets.
		' Note 2: On July 10, 2008 Kevin asked to remove access to all data related to Altantic markets from the site.
		' if Len(option_1_a)>0 and period<105 then MarketSelector=output.PrependHack(MarketSelector, option_1_a, Lang("all_atlantic"), check_flag)
		' if Len(option_0_a)>0 then MarketSelector=output.PrependHack(MarketSelector, option_0_a, Lang("all_national"), check_flag)

		if Len(option_0_a)>0 then MarketSelector=output.PrependHack(MarketSelector, option_0_a, Lang("all_markets"), check_flag)

	case 3 ' --- choose one market from market groups 1 and 2

		Set t.TableObj=output
		t.SetFilterByGroup("1,2")
		MarketSelector=t.Read

	case 4 ' --- choose one market from market group 1

		Set t.TableObj=output
		t.SetFilterByGroup("1")
		MarketSelector=t.Read

	case 5 ' --- choose one market from market group 2

		Set t.TableObj=output
		t.SetFilterByGroup("2")
		MarketSelector=t.Read

	case 6 ' --- choose one market from market groups 1 and 3

		Set t.TableObj=output
		t.SetFilterByGroup("1,3")
		MarketSelector=t.Read

	case 7 ' --- choose one market from market groups 1 and 3

     	Set t.TableObj=output
     	t.SetFilterByGroup("7")
       	MarketSelector=t.Read
       	if Len(MarketSelector)>0 then MarketSelector=output.PrependHack(MarketSelector, option_0_a, Lang("all_markets"), check_flag)

	case else ' --- choose any number of markets from any market group (for super-administrators, the most flexibe option)

		output.Multi=true
		Set t.TableObj=output
		MarketSelector=t.Read

	end select

	Set plaintext=nothing

	Set t = nothing

End Function


Function MarketSelector_HTML(period, selector_type, chart_id, flag_q1234)

	Set output=new cTableListbox
	output.FullDisplay=false
	MarketSelector_HTML=MarketSelector(output, period, selector_type, chart_id, flag_q1234)
	Set output=nothing

End Function


Function MarketSelector_Ajax(period, selector_type, chart_id, flag_q1234)

	Set output=new cTablePlaintext
	output.DisplayMode=-1 ' json
	MarketSelector_Ajax=MarketSelector(output, period, selector_type, chart_id, flag_q1234)
	Set output=nothing

End Function

%>
