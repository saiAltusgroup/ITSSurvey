<%

' DK, Apr 01, 2008 - version 1.0

Class cChartArea_HistoricalValuationParams ' Historical Valuation Parameters Area Chart '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=1
		Title=Lang("report_H100")
		Filename="chart_20.asp?ID=0"
	End Sub

	Public Function Init

		all_period_ids=g_HISTORY_PERIOD_ID_ARR

		ChartIDArr=Array(UniqueID, UniqueID, UniqueID)
		ColumnIDArr=Array(1, 1, 1)
		PeriodIDArr=Array(all_period_ids, all_period_ids, all_period_ids)
		QuestionTypeIDArr=Array(1, 1, 1)
		FilterForValuesArr=Array(" AND ri.Value=1", " AND ri.Value=2", " AND ri.Value=3")
		GroupingTypeArr=Array(121, 121, 121)
	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=3

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("chart_item_head_d1"),3) & table_obj.HeadSpan(Lang("chart_item_head_d2"),3) & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_e1")) & table_obj.Head(Lang("chart_item_head_e2")) & table_obj.Head(Lang("chart_item_head_e3")) & table_obj.Head(Lang("chart_item_head_e1")) & table_obj.Head(Lang("chart_item_head_e2")) & table_obj.Head(Lang("chart_item_head_e3")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a7"), Lang("chart_item_unit_a7"), Lang("chart_item_unit_a7"), Lang("chart_item_unit_a8"), Lang("chart_item_unit_a8"), Lang("chart_item_unit_a8"))
		table_obj.ChartColumnDescs=Array(Lang("chart_item_desc_e1"), Lang("chart_item_desc_e2"), Lang("chart_item_desc_e3"), Lang("chart_item_desc_e1"), Lang("chart_item_desc_e2"), Lang("chart_item_desc_e3"))
		table_obj.ChartIgnoredColumns="1,2,3"


		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		s = table_obj.RowOpen

		total=0
		for i=0 to ubound(point_array)
			total = total + point_array(i)
			s = s & table_obj.Cell(point_array(i))
		next

		for i=0 to ubound(point_array)
			s = s & table_obj.Cell(FormatNumber(100*SafeDivide(point_array(i), total), 1))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartLine_HistoricalValuationParams ' Historical Valuation Parameters Line Chart '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private FlavourID
	Private CoumnSetID
	Private PricePerUnit
	Private PricePerUnitShort

	Private MarketName

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=2
		Title=Lang("report_H101")
		PricePerUnit=Lang("chart_item_persqft")
		PricePerUnitShort=Lang("chart_item_persqft_short")
		SetFlavourID(0)
		MarketName=""
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		all_period_ids=g_HISTORY_PERIOD_ID_ARR

		if CoumnSetID=1 then
			ChartIDArr=Array(UniqueID)
			ColumnIDArr=Array(0)
			PeriodIDArr=Array(all_period_ids)
			QuestionTypeIDArr=Array(2) ' ***** last number used to be 1
			FilterForValuesArr=Array(f)
			GroupingTypeArr=Array(142)
		elseif CoumnSetID=2 then
			' Created for the "Downtown Class AA Office Ground Lease" chart
			ChartIDArr=Array(UniqueID)
			ColumnIDArr=Array(1)
			PeriodIDArr=Array(all_period_ids)
			QuestionTypeIDArr=Array(2) ' ***** last number used to be 1
			FilterForValuesArr=Array(f)
			GroupingTypeArr=Array(142)
		elseif CoumnSetID=3 then
			ChartIDArr=Array(UniqueID,UniqueID,UniqueID)
			ColumnIDArr=Array(1,2,3)
			PeriodIDArr=Array(all_period_ids,all_period_ids,all_period_ids)
			QuestionTypeIDArr=Array(2,2,2) ' ***** last number used to be 1
			FilterForValuesArr=Array(f,f,f)
			GroupingTypeArr=Array(142,142,142)
		else
			ChartIDArr=Array(UniqueID,UniqueID,UniqueID,UniqueID)
			ColumnIDArr=Array(1,2,3,0)
			PeriodIDArr=Array(all_period_ids,all_period_ids,all_period_ids,all_period_ids)
			QuestionTypeIDArr=Array(2,2,2,2) ' ***** last number used to be 1
			FilterForValuesArr=Array(f,f,f,f)
			GroupingTypeArr=Array(142,142,142,142)
		end if

	End Function

	Public Function SetFlavourID(ID)

		Filename="chart_19.asp?ID=0"

		if ID=3 then
			FlavourID=3
			CoumnSetID=3
		elseif ID=2 then
			FlavourID=2
			CoumnSetID=4
			PricePerUnit=Lang("chart_item_persuite")
			PricePerUnitShort=Lang("chart_item_persuite_short")
		elseif ID=1 then
			FlavourID=1
			CoumnSetID=1
			Filename="chart_19.asp?ID=2"
		elseif ID=4 then
			FlavourID=4
			CoumnSetID=2
		else
			FlavourID=0
			CoumnSetID=4
		end if

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen

			if CoumnSetID=4 then InitCallback = InitCallback & table_obj.Head(Lang("chart_item_head_f1")) & table_obj.Head(Lang("chart_item_head_f2")) & table_obj.Head(Lang("chart_item_head_f3")) & table_obj.Head(Lang("chart_item_dollar") & " " & PricePerUnitShort)
			if CoumnSetID=3 then InitCallback = InitCallback & table_obj.Head(Lang("chart_item_head_f1")) & table_obj.Head(Lang("chart_item_head_f2")) & table_obj.Head(Lang("chart_item_head_f3"))
			if CoumnSetID=2 then InitCallback = InitCallback & table_obj.Head(Lang("chart_item_head_f1"))
			if CoumnSetID=1 then InitCallback = InitCallback & table_obj.Head(Lang("chart_item_dollar") & " " & PricePerUnitShort)

			InitCallback = InitCallback & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2="{" & CurrPeriodLabel & "}"
		table_obj.ChartSubtitle=s_property

		if CoumnSetID=1 then
			table_obj.ChartColumnUnits=Array(Lang("chart_item_dollar") & " " & PricePerUnit)
			table_obj.ChartColumnDescs=Array(Lang("chart_item_dollar") & " " & PricePerUnit)
		elseif CoumnSetID=2 then
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"))
			table_obj.ChartColumnDescs=Array(Lang("chart_item_desc_b1"))
		else
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_dollar") & " " & PricePerUnit)
			table_obj.ChartColumnDescs=Array(Lang("chart_item_desc_f1"), Lang("chart_item_desc_f2"), Lang("chart_item_desc_f3"), Lang("chart_item_dollar") & " " & PricePerUnit)
		end if
		table_obj.ChartIgnoredColumns=""

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'if TypeName(table_obj)="cTableXmlchart" then
			MarketName=GetLabelPart(table_obj.RowID, " - ", 0)
			table_obj.RowID=GetLabelPart(table_obj.RowID, " - ", 1)
		'end if

		s = table_obj.RowOpen

		dec_places=1

		for i=0 to ubound(point_array)

			point_array(i)=Round(point_array(i), 2)

			if i=ubound(point_array) and (CoumnSetID=1 or CoumnSetID=4) then dec_places=2

			point_array(i)=FormatNumber(point_array(i), dec_places)

			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		s=""
		if TypeName(table_obj)<>"cTableXmlchart" then s="<h3>" & MarketName & "</h3>" else read_head = Replace(read_head, table_obj.ChartTitle2, CurrPeriodLabel & " - " & MarketName)

		FinalCallback=s & read_head & read_body & read_foot

	End Function

End Class



Class cChartLine_HistoricalFamilyOfRates

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Public UnitForYAxis

	Private MarketName

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=4
		Title=Lang("report_R000")
		Filename="chart_19.asp?ID=1"
		UnitForYAxis=Lang("chart_item_dollarpersqft")
		MarketName=""
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		all_period_ids=g_HISTORY_PERIOD_ID_ARR

		ChartIDArr=Array(UniqueID, UniqueID)
		ColumnIDArr=Array(1, 1)
		PeriodIDArr=Array(all_period_ids, all_period_ids)
		QuestionTypeIDArr=Array(2, 2)
		FilterForValuesArr=Array(f, f)
		GroupingTypeArr=Array(145, 141)

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.HeadSpan(UnitForYAxis,3) & table_obj.Head("") & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2="{" & CurrPeriodLabel & "}"
		table_obj.ChartSubtitle=s_property

		table_obj.ChartColumnUnits=Array(UnitForYAxis, UnitForYAxis, UnitForYAxis, "")
		table_obj.ChartColumnDescs=Array("", UnitForYAxis & " " & Lang("chart_item_minmaxrange"), UnitForYAxis & " " & Lang("chart_item_average"), Lang("chart_item_count"))
		table_obj.ChartIgnoredColumns="4"

		InitCallback=""

	End Function


	Public Function PointCallback(table_obj, point_array)

		'if TypeName(table_obj)="cTableXmlchart" then
			MarketName=GetLabelPart(table_obj.RowID, " - ", 0)
			table_obj.RowID=GetLabelPart(table_obj.RowID, " - ", 1)
		'end if

		s = table_obj.RowOpen

		for i=0 to ubound(point_array)

			if i<ubound(point_array) then
				point_array(i)=Round(point_array(i), 2)
				point_array(i)=FormatNumber(point_array(i), 2)
			end if

			'if i=2 then point_array(i)=FormatNumber(point_array(i), 1)
			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function


	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		s=""
		if TypeName(table_obj)<>"cTableXmlchart" then s="<h3>" & MarketName & "</h3>" else read_head = Replace(read_head, table_obj.ChartTitle2, CurrPeriodLabel & " - " & MarketName)

		FinalCallback=s & read_head & read_body & read_foot

	End Function

End Class



Class cChartBar_HistoricalBarometer ' Historical Barometer Bar Chart '

	Public isHistorical

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private BarometerTypeID
	Private BarometerCalcId
	Private MarketName


	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		Call SetCalcId(0)
		Call SetType(0)
		BarometerType=0
		UniqueID=3
		'TempCount=0
		MarketName=""
	End Sub


	Public Function SetCalcId(calc_id)
		BarometerCalcId=calc_id
	End Function

	Public Function SetType(type_id)

		' BarometerTypeID: 0 = Product/Market Barometer, 1=Location Barometer, 2=Property Type Barometer

		BarometerTypeID=type_id

		if BarometerTypeID=1 then
			Title=Lang("report_H401")
		elseif BarometerTypeID=2 then
			Title=Lang("report_H402")
		else
			Title=Lang("report_H400")
		end if

		Filename="chart_18.asp?ID=3"

	End Function


	Public Function Init

		all_period_ids=g_HISTORY_PERIOD_ID_ARR

		if BarometerCalcId=1 then

			ChartIDArr=Array(UniqueID,UniqueID,UniqueID, UniqueID)
			ColumnIDArr=Array(1,1,1, 1)
			PeriodIDArr=Array(all_period_ids,all_period_ids,all_period_ids, all_period_ids)
			QuestionTypeIDArr=Array(1,1,1, 1)
			FilterForValuesArr=Array(" AND ri.Value=1"," AND ri.Value=2"," AND ri.Value=3", " AND (ri.Value=1 OR ri.Value=2 OR ri.Value=3)")

			if BarometerTypeID=1 then
				GroupingTypeArr=Array(141,141,141, 141)
			elseif BarometerTypeID=2 then
				GroupingTypeArr=Array(151,151,151, 151)
			else
				GroupingTypeArr=Array(161,161,161, 161)
			end if

		else

			ChartIDArr=Array(UniqueID,UniqueID, UniqueID)
			ColumnIDArr=Array(1,1, 1)
			PeriodIDArr=Array(all_period_ids,all_period_ids, all_period_ids)
			QuestionTypeIDArr=Array(1,1, 1)
			FilterForValuesArr=Array(" AND ri.Value=1"," AND ri.Value=2", " AND (ri.Value=1 OR ri.Value=2)")

			if BarometerTypeID=1 then
				GroupingTypeArr=Array(141,141, 141)
			elseif BarometerTypeID=2 then
				GroupingTypeArr=Array(151,151, 151)
			else
				GroupingTypeArr=Array(161,161, 161)
			end if

		end if

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			if BarometerCalcId=1 then
				InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head("") & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
			else
				InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
			end if
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2="{" & CurrPeriodLabel & "}"
		table_obj.ChartSubtitle=s_property

		if BarometerCalcId=1 then
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a10"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"))
			table_obj.ChartColumnDescs=Array("", Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"))
			table_obj.ChartIgnoredColumns="2,3,4"
		else
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a5"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"))
			table_obj.ChartColumnDescs=Array("", Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"))
			table_obj.ChartIgnoredColumns="2,3"
		end if

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'Response.Write BarometerTypeID

		'if TypeName(table_obj)="cTableXmlchart" then
			if BarometerTypeID=0 then
				MarketName=GetLabelPart(table_obj.RowID, " - ", 0)
				table_obj.RowID=GetLabelPart(table_obj.RowID, " - ", 2)
			elseif BarometerTypeID=1 then
				MarketName=GetLabelPart(table_obj.RowID, " - ", 0)
				table_obj.RowID=GetLabelPart(table_obj.RowID, " - ", 1)
			elseif BarometerTypeID=2 then
				MarketName=""
				table_obj.RowID=GetLabelPart(table_obj.RowID, " - ", 1)
			end if
		'end if

		last_index=ubound(point_array)

		'if TypeName(table_obj)<>"cTableXmlchart" then
		'	table_obj.Extra="(" & point_array(last_index) & chr(160) & Lang("chart_item_respondents") & ")"
		'end if

		if BarometerCalcId=1 then
			' --- phase 1 ---
			if (point_array(0)>point_array(1)) then point_array(0)=-SafeDivide(point_array(0),point_array(1)+point_array(2)) else point_array(0)=SafeDivide(point_array(1),point_array(0)+point_array(2))
			point_array(0)=FormatNumber(point_array(0), 1)
			point_array(1)="" ' unused
			point_array(2)="-" ' unused
		else
			' --- phase 1 ---
			if (point_array(0)>=point_array(1)) then point_array(0)=SafeDivide(point_array(0),point_array(1)) else point_array(0)=-SafeDivide(point_array(1),point_array(0))
			point_array(0)=FormatNumber(point_array(0), 1)
			point_array(1)="-" ' unused
		end if

		'if TypeName(table_obj)<>"cTableXmlchart" then
		'point_array(last_index-1)=NumToDir(point_array(last_index-1))
		'end if

		s = table_obj.RowOpen

		for i=0 to ubound(point_array)

			if BarometerCalcId=1 then
				if i=0 or i=3 or i=6 or i=9 then
					table_obj.Label=point_array(i)
					'if BarometerTimeID=0 then table_obj.Symbol=point_array(last_index-1)
				end if
			else
				if i=0 or i=2 or i=4 or i=6 then
					table_obj.Label=point_array(i)
					'if BarometerTimeID=0 then table_obj.Symbol=point_array(last_index-1)
				end if
			end if

			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		s=""
		if len(MarketName)>0 then
			if TypeName(table_obj)<>"cTableXmlchart" then s="<h3>" & MarketName & "</h3>" else read_head = Replace(read_head, table_obj.ChartTitle2, CurrPeriodLabel & " - " & MarketName)
		else
			if TypeName(table_obj)<>"cTableXmlchart" then s="" else read_head = Replace(read_head, table_obj.ChartTitle2, CurrPeriodLabel)
		end if

		FinalCallback=s & read_head & read_body & read_foot

	End Function

End Class


%>
