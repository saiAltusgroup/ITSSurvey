<%

' DK, Apr 01, 2008 - version 1.0

Class cTableReport

	Public TableObj
	Public CurrPeriodID
	Public GroupIdFilter
	Public ReportArr
	Public CheckExistsInPeriod

	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid

		GroupIdFilter=2
		CurrPeriodID=g_CURRENT_PERIOD_ID

		ReDim ReportArr(62)

		' report_id | property_selector_type | market_selector_type | ChartId | GroupId | report_name

		ReportArr(0)="R101|2|1|2|0|" & Lang("report_R101")
		ReportArr(1)="R100|2|1|1|0|" & Lang("report_R100")
		'ReportArr(2)="R102|2|4|2|0|" & Lang("report_R102")
		'ReportArr(2)="R111|2|1|21|0|" & Lang("report_R111")
		'ReportArr(2)="R110|2|1|20|0|" & Lang("report_R110")
		ReportArr(2)="R401|1|1|3|0|" & Lang("report_R401")
		ReportArr(3)="R400|1|1|3|0|" & Lang("report_R400")
		ReportArr(4)="R402|1|2|3|0|" & Lang("report_R402")
		ReportArr(5)="R403|1|1|3|0|" & Lang("report_R403")
		ReportArr(6)="R404|1|1|3|0|" & Lang("report_R404")
		ReportArr(7)="R500|1|1|19|0|" & Lang("report_R500")
		ReportArr(8)="R200|2|1|4|0|" & Lang("report_R200")
		ReportArr(9)="R201|2|1|5|0|" & Lang("report_R201")
		ReportArr(10)="R202|2|1|6|0|" & Lang("report_R202")
		ReportArr(11)="R203|2|1|7|0|" & Lang("report_R203")
		ReportArr(12)="R300|2|1|8|0|" & Lang("report_R300")
		ReportArr(13)="H101|2|6|2|1|" & Lang("report_H101")
		ReportArr(14)="H100|2|6|1|1|" & Lang("report_H100")
		'ReportArr(16)="H111|2|6|21|1|" & Lang("report_H111")
		'ReportArr(16)="H110|2|6|20|1|" & Lang("report_H110")
		ReportArr(15)="H401|2|4|3|1|" & Lang("report_H401")
		ReportArr(16)="H400|2|4|3|1|" & Lang("report_H400")
		ReportArr(17)="H402|2|2|3|1|" & Lang("report_H402")
		ReportArr(18)="H500|2|4|19|1|" & Lang("report_H500")
		ReportArr(19)="H200|2|4|4|1|" & Lang("report_H200")
		ReportArr(20)="H201|2|4|5|1|" & Lang("report_H201")
		ReportArr(21)="H202|2|4|6|1|" & Lang("report_H202")
		ReportArr(22)="H203|2|4|7|1|" & Lang("report_H203")
		ReportArr(23)="X100|1|0|-1|2|" & Lang("report_X100")
		ReportArr(24)="S100|1|0|10|2|" & Lang("report_S100")
		ReportArr(25)="S101|1|0|11|2|" & Lang("report_S101")
		ReportArr(26)="S102|1|0|12|2|" & Lang("report_S102")
		ReportArr(27)="S200|1|0|13|2|" & Lang("report_S200")
		ReportArr(28)="S300|1|0|14|2|" & Lang("report_S300")
		ReportArr(29)="S301|1|0|15|2|" & Lang("report_S301")
		ReportArr(30)="S302|1|0|16|2|" & Lang("report_S302")
		ReportArr(31)="S303|1|0|17|2|" & Lang("report_S303")
		ReportArr(32)="S304|1|0|18|-2|" & Lang("report_S304")	' Value -2 means that this report is currently disabled
		ReportArr(33)="S305|1|0|2|2|" & Lang("report_S305")
		ReportArr(34)="S306|1|0|2|2|" & Lang("report_S306")

		' Added after Jan 30, 2009
		' Quarter 1 Reports
		ReportArr(35)="R120|2|1|22|0|" & Lang("report_R120")
		ReportArr(36)="R121|2|1|23|0|" & Lang("report_R121")
		ReportArr(37)="R122|2|1|24|0|" & Lang("report_R122")
		ReportArr(38)="R123|1|1|25|2|" & Lang("report_R123")
		ReportArr(39)="R124|2|1|26|0|" & Lang("report_R124")
		ReportArr(40)="R125|2|1|27|0|" & Lang("report_R125")
		ReportArr(41)="R126|2|1|28|0|" & Lang("report_R126")

		ReportArr(42)="R127|1|1|30|0|" & Lang("report_R127")
		ReportArr(43)="R130|1|1|32|0|" & Lang("report_R130")
		'ReportArr(44)="R128|1|0|31|0|" & Lang("report_R128")
		'ReportArr(45)="R129|1|0|33|0|" & Lang("report_R129")

		ReportArr(44)="T100|1|1|29|2|" & Lang("report_T100")	' Value -2 means that this report is currently disabled. It needs to be changed. (CSDM83)

		' Quarter 2 Reports
		ReportArr(45)="R131|1|1|34|0|" & Lang("report_R131")
		ReportArr(46)="R132|1|1|35|0|" & Lang("report_R132")
		ReportArr(47)="R133|1|1|36|0|" & Lang("report_R133")
		ReportArr(48)="R134|1|1|37|2|" & Lang("report_R134")

		' ReportArr(49)="S307|1|0|38|2|" & Lang("report_S307") ' Removed in Q2 2010 as per Bob's request
		
		ReportArr(49)="S308|1|0|39|2|" & Lang("report_S308")
		
		' ReportArr(51)="S309|1|0|40|2|" & Lang("report_S309") ' Removed in Q2 2010 as per Bob's request

		ReportArr(50)="T101|1|1|41|-2|" & Lang("report_T101")	' Value -2 means that this report is currently disabled. It needs to be changed. (Financial Indicators)
		
		' ReportArr(53)="T102|1|1|42|2|" & Lang("report_T102")

		' Quarter 3 Reports
		ReportArr(51)="R135|1|1|43|2|" & Lang("report_R135")
		ReportArr(52)="S310|1|0|44|2|" & Lang("report_S310")
		ReportArr(53)="R136|1|1|45|2|" & Lang("report_R136")

		' Quarter 4 Reports
		ReportArr(54)="S311|1|0|46|2|" & Lang("report_S311")
		ReportArr(55)="S312|1|0|47|2|" & Lang("report_S312")
		ReportArr(56)="R137|1|1|48|2|" & Lang("report_R137")
		ReportArr(57)="R138|1|1|49|2|" & Lang("report_R138")
		ReportArr(58)="R139|1|1|50|2|" & Lang("report_R139")
		ReportArr(59)="S313|1|1|51|2|" & Lang("report_S313")
		ReportArr(60)="S314|1|1|52|2|" & Lang("report_S314")
		ReportArr(61)="S315|1|1|53|2|" & Lang("report_S315")
        ReportArr(62)="R140|1|7|55|0|" & Lang("report_R140")

    ' report_id | property_selector_type | market_selector_type | ChartId | GroupId | report_name

		CheckExistsInPeriod=true

	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub


	Public Sub DeactivateExceptionsByPeriod(period_id)

		if (period_id=108) then ReportArr(32) = replace(ReportArr(32), "S304|1|0|18|2|", "S304|1|0|18|-2|")	' Report S304 is an exception
		if (period_id=108) then ReportArr(51) = replace(ReportArr(51), "S309|1|0|40|2|", "S309|1|0|40|-2|")	' Report S309 is an exception

	End Sub

	Public Function Read()

		Read = TableObj.TableOpen(0)

		For i=0 to ubound(ReportArr)
			s=split(ReportArr(i), "|")
'			if(s(5)="Homogeneous Portfolio Effect" or s(5)="Influence d'un portefeuille homogène") then
'					a=2
'
'			else
			TableObj.RowID=s(0)

			exists_in_period=true
			if CheckExistsInPeriod then exists_in_period=ExistsInPeriod(s(3))

			if CInt(s(4))=GroupIdFilter and exists_in_period then Read = Read & TableObj.RowOpen & TableObj.Cell(s(5)) & TableObj.RowClose
'			End If
		next

		Read = Read & TableObj.TableClose

	End Function

	Public Function GetPropertySelectorType(report_id)
		GetPropertySelectorType=0
		for i=0 to ubound(ReportArr)
			s=split(ReportArr(i), "|")
			if s(0)=report_id then GetPropertySelectorType=s(1) : exit function
		next
	End Function

	Public Function GetMarketSelectorType(report_id)
		GetMarketSelectorType=0
		for i=0 to ubound(ReportArr)
			s=split(ReportArr(i), "|")
			if s(0)=report_id then GetMarketSelectorType=s(2) : exit function
		next
	End Function

	Public Function GetChartID(report_id)
		GetChartID=0
		for i=0 to ubound(ReportArr)
			s=split(ReportArr(i), "|")
			if s(0)=report_id then GetChartID=s(3) : exit function
		next
	End Function


	Public Function GetGroupId(report_id)
		GetGroupId=-1
		for i=0 to ubound(ReportArr)
			s=split(ReportArr(i), "|")
			if s(0)=report_id then GetGroupId=s(4) : exit function
		next
	End Function

	Private Function ExistsInPeriod(chart_id)

		'ExistsInPeriod=true
		'Exit Function

		ExistsInPeriod=false

		Set c = new cConverter
		period_bitvalue=c.NumToLng("LKP02_PERIOD_BITVALUE", CurrPeriodID)
		Set c = nothing

		strSQL="SELECT QuestionID FROM dat_Question WHERE ChartID=" & chart_id & " AND ((QuarterBitMask & " & period_bitvalue & ")<>0)"

		'Response.Write "CurrPeriodID=" & CurrPeriodID & "<br>"
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
