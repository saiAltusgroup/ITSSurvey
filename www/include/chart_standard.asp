<%

' DK, Apr 01, 2008 - version 1.0

Class cChartPie_ValuationParams ' Valuation Parameters Pie Chart '

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

	Public GroupingType

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=1
		Title=Lang("report_R100")
		Filename="chart_16.asp?ID=0"
		GroupingType=12
	End Sub

	Public Function Init
		ChartIDArr=Array(UniqueID)
		ColumnIDArr=Array(1)
		PeriodIDArr=Array(CurrPeriodID)
		QuestionTypeIDArr=Array(1)
		FilterForValuesArr=Array("")
		GroupingTypeArr=Array(GroupingType)
	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'table_obj.Label="(" & point_array(0) & " " & Lang("chart_item_respondents") & ")|" & ucase(table_obj.RowID) & "|{%.%}%"
		table_obj.Label="|" & ucase(table_obj.RowID) & "|{%.%}%"

		s = table_obj.RowOpen

		for i=0 to ubound(point_array)
			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartBar_DynamicPie ' same model as cChartPie_ValuationParams but for multiple columns that are displayed as bars

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

	Public GroupingType
	Public LabelsArr
	Public ColumnsArr
	Public ValuesArr

	Private Table  '(cols, rows)
	Private LastRow

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=1
		Title=Lang("report_R134")
		Filename="chart_18.asp?ID=4"
		GroupingType=16
		LabelsArr=Array("Untitled")
		ColumnsArr=Array(1)
		ValuesArr=Array(1)
		LastRow=-1
		redim Table(0, 0)
	End Sub

	Public Function Init

		j=Ubound(LabelsArr)

		Redim ChartIDArr(j)
		Redim ColumnIDArr(j)
		Redim PeriodIDArr(j)
		Redim QuestionTypeIDArr(j)
		Redim FilterForValuesArr(j)
		Redim GroupingTypeArr(j)

		for i=0 to j

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=ColumnsArr(i)
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=1
			FilterForValuesArr(i)=""
			GroupingTypeArr(i)=GroupingType

		next

		LastRow=-1
		redim Table(j, 0)

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		j=Ubound(ValuesArr)
		k=Ubound(LabelsArr)

		cols_to_add=j-k

		if TypeName(table_obj)<>"cTableXmlchart" then

			InitCallback=table_obj.RowOpen
			for i=0 to j
				InitCallback = InitCallback & table_obj.Head(ValuesArr(i))
			next
			InitCallback = InitCallback & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property


		Redim temp_array(j)

		table_obj.ChartColumnUnits=temp_array
		table_obj.ChartColumnDescs=temp_array

		for i=0 to j

			table_obj.ChartColumnUnits(i)=Lang("chart_item_head_d3")
			table_obj.ChartColumnDescs(i)=ValuesArr(i)

		next

		LastRow=-1
		redim Table(k, 0)

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		j=ubound(point_array)
		Redim temp_array(j)

		LastRow=LastRow+1
		redim preserve Table(j, LastRow)

		'table_obj.Label="(" & point_array(0) & " " & Lang("chart_item_respondents") & ")|" & ucase(table_obj.RowID) & "|{%.%}%"
		' table_obj.Label="|" & ucase(table_obj.RowID) & "|{%.%}%"

		's = table_obj.RowOpen

		for i=0 to j
			's = s & table_obj.Cell(point_array(i))
			Table(i, LastRow)=point_array(i)
		next

		'PointCallback = s & table_obj.RowClose

		PointCallback = ""

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		last_col=Ubound(LabelsArr)

		s=""

		' Loop to transpose table

		for i=0 to last_col

			table_obj.RowID=LabelsArr(i)

			s = s & table_obj.RowOpen

			sum = 0
			for j=0 to LastRow
				sum = sum + Table(i, j)
			next

			for j=0 to LastRow
				percent=100*SafeDivide(Table(i, j), sum)
				percent=FormatNumber(Round(percent, 1), 1)

				table_obj.Label=percent
				s = s & table_obj.Cell(percent)
			next

			s = s & table_obj.RowClose

		next

		read_body=s

		FinalCallback=read_head & read_body & read_foot

	End Function

End Class



Class cChartBar_ValuationParams ' Valuation Parameters Bar Chart '

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
	Private PricePerUnit
	Private PricePerUnitShort

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=2
		Title=Lang("report_R101")
		PricePerUnit=Lang("chart_item_persqft")
		PricePerUnitShort=Lang("chart_item_persqft_short")
		SetFlavourID(0)
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		prev_period_id=CurrPeriodID-PrevPeriodDelta
		complementary_id=UniqueID-1

		ChartIDArr=Array(complementary_id, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID)

		if FlavourID=1 then
			ColumnIDArr=Array(1, 0,0, 0,0, 0,0, 0,0)
		else
			ColumnIDArr=Array(1, 1,1, 2,2, 3,3, 0,0)
		end if

		PeriodIDArr=Array(CurrPeriodID, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id)
		QuestionTypeIDArr=Array(1, 2,2, 2,2, 2,2, 2,2) ' ***** last two numbers used to be 1
		FilterForValuesArr=Array(f, f,f, f,f, f,f, f,f)
		GroupingTypeArr=Array(1, 5,2, 5,2, 5,2, 2,2)

	End Function

	Public Function SetFlavourID(ID)

		if ID=2 then
			PricePerUnit=Lang("chart_item_persuite")
			PricePerUnitShort=Lang("chart_item_persuite_short")
			ID=0
		elseif ID=3 then
			PricePerUnit=""
			PricePerUnitShort=""
			ID=0
		end if

		if ID=1 then
			FlavourID=1
			Filename="chart_17.asp?ID=1"
		else
			FlavourID=0
			Filename="chart_17.asp?ID=0"
		end if
	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=-1

		if TypeName(table_obj)<>"cTableXmlchart" then
			if FlavourID=1 then
				InitCallback=table_obj.RowOpen & table_obj.HeadSpan(PricePerUnit,4) & table_obj.HeadSpan(PricePerUnit,4) & table_obj.HeadSpan(PricePerUnit,4) & table_obj.HeadSpan(PricePerUnit,2) & table_obj.RowClose
				InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_dollar")) & table_obj.Head(Lang("chart_item_change")) & table_obj.RowClose
			else
				InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("chart_item_head_a1"),4) & table_obj.HeadSpan(Lang("chart_item_head_a2"),4) & table_obj.HeadSpan(Lang("chart_item_head_a3"),4)
				if Len(PricePerUnit)>0 then InitCallback=InitCallback & table_obj.HeadSpan(PricePerUnit,2)
				InitCallback=InitCallback & table_obj.RowClose
				InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change"))
				if Len(PricePerUnit)>0 then InitCallback=InitCallback & table_obj.Head(Lang("chart_item_dollar")) & table_obj.Head(Lang("chart_item_change"))
				InitCallback=InitCallback & table_obj.RowClose
			end if
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		if FlavourID=1 then
			unit=""
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a6"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a3"), Lang("chart_item_unit_a2"))
			table_obj.ChartColumnDescs=Array("", "", "", "", "", "", "", "", "", Lang("chart_item_desc_a2"), Lang("chart_item_desc_a2"), Lang("chart_item_desc_a2"), Lang("chart_item_desc_a3"), Lang("chart_item_desc_a2"))
			table_obj.ChartIgnoredColumns="4,5,6,7,8,9,10,11,12,13,14"
		else
			table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a3"), Lang("chart_item_unit_a2"))
			table_obj.ChartColumnDescs=Array("", Lang("chart_item_desc_b1"), "", Lang("chart_item_desc_a2"), "", Lang("chart_item_desc_b2"), "", Lang("chart_item_desc_a2"), "", Lang("chart_item_desc_b3"), "", Lang("chart_item_desc_a2"), Lang("chart_item_desc_a3"), Lang("chart_item_desc_a2"))
			table_obj.ChartIgnoredColumns="4,8,12,13,14"
		end if


		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		dec_places=2
		if FlavourID=0 then dec_places=1

		point_array(4)=DiffToSgn(point_array(3), point_array(4), 1)
		point_array(8)=DiffToSgn(point_array(7), point_array(8), 1)
		point_array(12)=DiffToSgn(point_array(11), point_array(12), 1)
		point_array(14)=DiffToSgn(point_array(13), point_array(14), 1)

		ubound_point_array=ubound(point_array)

		table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"
		if Len(PricePerUnitShort)>0 then
			table_obj.Extra2=Lang("chart_item_dollar") & FormatNumber(point_array(13), 2) & " " & PricePerUnitShort & " - " & NumToDir(point_array(14))
		else
			ubound_point_array=ubound_point_array-2
		end if

		s = table_obj.RowOpen

		for i=1 to ubound_point_array
			point_array(i)=Round(point_array(i), 2)

			table_obj.Label=""
			if i=3 or i=7 or i=11 then
				point_array(i)=FormatNumber(point_array(i), dec_places) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
			elseif i=13 then
				point_array(i)=FormatNumber(point_array(i), 2)
			else
				point_array(i)=FormatNumber(point_array(i), dec_places)
			end if

			if TypeName(table_obj)<>"cTableXmlchart" then
				if i=4 or i=8 or i=12 or i=14 then point_array(i)=NumToDir(point_array(i))
			end if

			s = s & table_obj.Cell(point_array(i))

		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartBar_ValuationParams_Dynamic ' Valuation Parameters Bar Chart (Dynamic Version) '

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

	Public LabelsArr
	Public Units
	Public DecPlaces
	Public RespondentsDisplayFormat

	Public GroupingBase
	Public ColumnStartID

	Public CustomSubtitle

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=22
		Title=""
		Filename="chart_17.asp?ID=0"
		LabelsArr=Array("")
		Units=Lang("chart_item_percent")
		DecPlaces = 1
		RespondentsDisplayFormat = 0
		GroupingBase=0
		ColumnStartID=0
		CustomSubtitle=""
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		prev_period_id=CurrPeriodID-PrevPeriodDelta

		nr_of_cols=Ubound(LabelsArr)+1

		j=2*(nr_of_cols+1)

		Redim ChartIDArr(j)
		Redim ColumnIDArr(j)
		Redim PeriodIDArr(j)
		Redim QuestionTypeIDArr(j)
		Redim FilterForValuesArr(j)
		Redim GroupingTypeArr(j)

		' first column (#0) is special (number of respondents)
		i=0
		ChartIDArr(i)=UniqueID
		ColumnIDArr(i)=1
		PeriodIDArr(i)=CurrPeriodID
		QuestionTypeIDArr(i)=2
		FilterForValuesArr(i)=f
		GroupingTypeArr(i)=1+GroupingBase

		'A hack but, hey it works...
		if UniqueID=43 OR UniqueID=45 OR UniqueID=52 then
			QuestionTypeIDArr(i)=1
		end if

		k=1+ColumnStartID
		grouping_type_a=5
		grouping_type_b=2

		for i=1 to j step 2

			if k=(nr_of_cols+1+ColumnStartID) then
				' last column is special
				grouping_type_a=2
				k=0
			end if

			question_type_id=2
			'A hack but, hey it works...
			if UniqueID=30 and k=7 then
				question_type_id=1
			elseif UniqueID=43 OR UniqueID=45 OR UniqueID=52 then
				question_type_id=1
			end if

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=k
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=question_type_id
			FilterForValuesArr(i)=f
			GroupingTypeArr(i)=grouping_type_a+GroupingBase

			ChartIDArr(i+1)=UniqueID
			ColumnIDArr(i+1)=k
			PeriodIDArr(i+1)=prev_period_id
			QuestionTypeIDArr(i+1)=question_type_id
			FilterForValuesArr(i+1)=f
			GroupingTypeArr(i+1)=grouping_type_b+GroupingBase

			k=k+1
		next

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=-1
		nr_of_cols=Ubound(LabelsArr)+1

		if TypeName(table_obj)<>"cTableXmlchart" then

			InitCallback=""

			if UniqueID=46 then InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("chart_asset_value"),nr_of_cols*4)  

			InitCallback=InitCallback & table_obj.RowOpen
			for col=1 to nr_of_cols
				InitCallback = InitCallback & table_obj.HeadSpan(LabelsArr(col-1),4)
			next
			InitCallback=InitCallback & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen
			for col=1 to nr_of_cols
				InitCallback = InitCallback & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head("")
			next
			InitCallback=InitCallback & table_obj.RowClose

			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		if Len(CustomSubtitle)>0 then s_property=CustomSubtitle
		table_obj.ChartSubtitle=s_property

		j=(nr_of_cols*4+2)-1
		Redim temp_array(j)

		table_obj.ChartColumnUnits=temp_array
		table_obj.ChartColumnDescs=temp_array

		k=0
		comma=""

		for col=1 to nr_of_cols

			table_obj.ChartColumnUnits(k)=Units
			table_obj.ChartColumnDescs(k)=""
			k=k+1

			table_obj.ChartColumnUnits(k)=""
			table_obj.ChartColumnDescs(k)=LabelsArr(col-1)
			k=k+1

			table_obj.ChartColumnUnits(k)=""
			table_obj.ChartColumnDescs(k)=""
			k=k+1

			table_obj.ChartColumnUnits(k)=""
			table_obj.ChartColumnDescs(k)=""
			k=k+1

			table_obj.ChartIgnoredColumns=table_obj.ChartIgnoredColumns & comma & k
			comma=","

		next

		'A hack but, hey it works...
		if UniqueID=30 then table_obj.ChartIgnoredColumns=table_obj.ChartIgnoredColumns & comma & "17,18,19,21,22,23,25,26,27"
		if UniqueID=31 then table_obj.ChartIgnoredColumns=table_obj.ChartIgnoredColumns & comma & "13,14,15"

		'Response.Write "table_obj.ChartIgnoredColumns=" & table_obj.ChartIgnoredColumns & "<br>"

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		nr_of_cols=Ubound(LabelsArr)+1

		for col=1 to nr_of_cols
			j=col*4
			point_array(j)=DiffToSgn(point_array(j-1), point_array(j), 1)
		next

		ubound_point_array=ubound(point_array)

		if (RespondentsDisplayFormat=1) Then

			table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"

		elseif (RespondentsDisplayFormat=2) Then

			point_array(19)=FormatNumber(Round(point_array(19), 2), 2)
			point_array(23)=FormatNumber(Round(point_array(23), 1), 1)
			point_array(27)=FormatNumber(Round(point_array(27), 1), 1)

			table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ") ($" & point_array(19) & " " & Lang("chart_item_persqft_short") & ")"
			table_obj.Extra2="(" & Lang("chart_max_loan_to_value") & " " & point_array(23) & "% " & NumToDir(point_array(24)) & ") (" & Lang("chart_min_spread") & " " & point_array(27) & " " & Lang("bps") & " " & NumToDir(point_array(28)) & ")"

		elseif (RespondentsDisplayFormat=3) Then

			point_array(15)=FormatNumber(Round(point_array(15), 2), 2)

			table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"
			table_obj.Extra2=Lang("avg_dollar_ps") & " : $" & point_array(15) & " " & NumToDir(point_array(16))

		end if

		ubound_point_array=ubound_point_array-2

		s = table_obj.RowOpen

		for i=1 to ubound_point_array

			dec_places=DecPlaces
			' A hack but, hey it works...
			'if UniqueID=30 and (i>=4*4+1) then dec_places=0
			if UniqueID=30 then
				if (i=17 or i=18 or i=21 or i=22 or i=25 or i=26) then dec_places=0
				if (i=19) then dec_places=2
			elseif UniqueID=31 then
				if (i=13 or i=14 or i=15) then dec_places=2
			end if

			point_array(i)=Round(point_array(i), dec_places)

			table_obj.Label=""

			for col=1 to nr_of_cols
				j=col*4-1
				if i=j then
					point_array(i)=FormatNumber(point_array(i), dec_places) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
				else
					point_array(i)=FormatNumber(point_array(i), dec_places)
				end if
			next

			if TypeName(table_obj)<>"cTableXmlchart" then

				for col=1 to nr_of_cols
					j=col*4
					if i=j then point_array(i)=NumToDir(point_array(i))
				next

			end if

			s = s & table_obj.Cell(point_array(i))

		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartBar_Barometer ' Barometer Bar Chart '

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

	Private TempArr1
	Private TempArr2
	Private TempCount

	Private BarometerTypeID
	Private BarometerTimeID
	Private BarometerCalcId
	'Private NrOfPeriods


	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		Call SetCalcId(0)
		Call SetTypeAndTime(0, 0)
		BarometerType=0
		UniqueID=3
		TempCount=0
	End Sub


	Public Function SetCalcId(calc_id)
		BarometerCalcId=calc_id
	End Function

	Public Function SetTypeAndTime(type_id, time_id)

		' BarometerTypeID: 0 = Product/Market Barometer, 1=Location Barometer, 2=Property Type Barometer
		' BarometerTimeID: 0 = Regular (Sorted top-to-bottom), 1 = Rolling Four Quarters

		BarometerTypeID=type_id
		BarometerTimeID=time_id


		if BarometerTypeID=1 then
			Title=Lang("report_R401")
		elseif BarometerTypeID=2 then
			Title=Lang("report_R402")
		else
			Title=Lang("report_R400")
		end if


		if BarometerTimeID=1 then
			Filename="chart_18.asp?ID=1"
			Title=Title & " - " & Lang("chart_item_qtr_rolling4")
			'NrOfPeriods=4
		else
			Filename="chart_18.asp?ID=0"
			Title=Title & " - " & Lang("chart_item_qtr_current")
			'NrOfPeriods=2
		end if

	End Function


	Public Function Init

		prev_period_id1=CurrPeriodID-1*PrevPeriodDelta
		prev_period_id2=CurrPeriodID-2*PrevPeriodDelta
		prev_period_id3=CurrPeriodID-3*PrevPeriodDelta

		if BarometerTimeID=1 then

			if BarometerCalcId=1 then

				Call NotImplemented()

			else

				ChartIDArr=Array(UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID)
				ColumnIDArr=Array(1,1, 1,1, 1,1, 1,1, 1)
				PeriodIDArr=Array(prev_period_id3,prev_period_id3, prev_period_id2,prev_period_id2, prev_period_id1,prev_period_id1, CurrPeriodID,CurrPeriodID, CurrPeriodID)
				QuestionTypeIDArr=Array(1,1, 1,1, 1,1, 1,1, 1)
				FilterForValuesArr=Array(" AND ri.Value=1"," AND ri.Value=2", " AND ri.Value=1"," AND ri.Value=2", " AND ri.Value=1"," AND ri.Value=2", " AND ri.Value=1"," AND ri.Value=2", " AND (ri.Value=1 OR ri.Value=2)")

				if BarometerTypeID=1 then
					GroupingTypeArr=Array(1,1, 1,1, 1,1, 1,1, 1)
				elseif BarometerTypeID=2 then
					GroupingTypeArr=Array(21,21, 21,21, 21,21, 21,21, 21)
				else
					GroupingTypeArr=Array(81,81, 81,81, 81,81, 81,81, 81)
				end if

			end if

		else

			if BarometerCalcId=1 then

				ChartIDArr=Array(UniqueID,UniqueID,UniqueID, UniqueID,UniqueID,UniqueID, UniqueID)
				ColumnIDArr=Array(1,1,1, 1,1,1, 1)
				PeriodIDArr=Array(prev_period_id1,prev_period_id1,prev_period_id1, CurrPeriodID,CurrPeriodID,CurrPeriodID, CurrPeriodID)
				QuestionTypeIDArr=Array(1,1,1, 1,1,1, 1)
				FilterForValuesArr=Array(" AND ri.Value=1"," AND ri.Value=2"," AND ri.Value=3", " AND ri.Value=1"," AND ri.Value=2"," AND ri.Value=3", " AND (ri.Value=1 OR ri.Value=2 OR ri.Value=3)")

				if BarometerTypeID=1 then
					GroupingTypeArr=Array(1,1,1, 1,1,1, 1)
				elseif BarometerTypeID=2 then
					GroupingTypeArr=Array(21,21,21, 21,21,21, 21)
				else
					GroupingTypeArr=Array(81,81,81, 81,81,81, 81)
				end if

			else

				ChartIDArr=Array(UniqueID,UniqueID, UniqueID,UniqueID, UniqueID)
				ColumnIDArr=Array(1,1, 1,1, 1)
				PeriodIDArr=Array(prev_period_id1,prev_period_id1, CurrPeriodID,CurrPeriodID, CurrPeriodID)
				QuestionTypeIDArr=Array(1,1, 1,1, 1)
				FilterForValuesArr=Array(" AND ri.Value=1"," AND ri.Value=2", " AND ri.Value=1"," AND ri.Value=2", " AND (ri.Value=1 OR ri.Value=2)")

				if BarometerTypeID=1 then
					GroupingTypeArr=Array(1,1, 1,1, 1)
				elseif BarometerTypeID=2 then
					GroupingTypeArr=Array(21,21, 21,21, 21)
				else
					GroupingTypeArr=Array(81,81, 81,81, 81)
				end if

			end if

		end if

		redim TempArr1(0)
		redim TempArr2(0)

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then

			Set c = new cConverter
			period_0=c.NumToStr("LKP01_PERIOD_NAME", CurrPeriodID)
			period_1=c.NumToStr("LKP01_PERIOD_NAME", CurrPeriodID-1)
			period_2=c.NumToStr("LKP01_PERIOD_NAME", CurrPeriodID-2)
			period_3=c.NumToStr("LKP01_PERIOD_NAME", CurrPeriodID-3)
			Set c = nothing

			if BarometerTimeID=1 then

				if BarometerCalcId=1 then
					Call NotImplemented()
				else
					InitCallback=table_obj.RowOpen & table_obj.HeadSpan(period_3,2) & table_obj.HeadSpan(period_2,2) & table_obj.HeadSpan(period_1,2) & table_obj.HeadSpan(Lang("chart_item_current") & " (" & period_0 & ")",3) & table_obj.RowClose
					InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
				end if

			else

				if BarometerCalcId=1 then
					InitCallback=table_obj.RowOpen & table_obj.HeadSpan(period_1,3) & table_obj.HeadSpan(Lang("chart_item_current") &" (" & period_0 & ")",4) & table_obj.RowClose
					InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head("") & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head("") & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
				else
					InitCallback=table_obj.RowOpen & table_obj.HeadSpan(period_1,2) & table_obj.HeadSpan(Lang("chart_item_current") &" (" & period_0 & ")",3) & table_obj.RowClose
					InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_head_b1")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
				end if

			end if

			exit function

		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property
		if BarometerTimeID=0 and BarometerTypeID=0 then table_obj.ChartSubtitle=table_obj.ChartSubtitle & " " & Lang("chart_item_top15least15")

		Set c = new cConverter

		if BarometerTimeID=1 then
			if BarometerCalcId=1 then
				Call NotImplemented()
			else
				table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a5"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"))
				table_obj.ChartColumnDescs=Array(c.NumToStr("LKP01_PERIOD_NAME",CurrPeriodID-3), Lang("chart_item_desc_a1"), c.NumToStr("LKP01_PERIOD_NAME",CurrPeriodID-2), Lang("chart_item_desc_a1"), c.NumToStr("LKP01_PERIOD_NAME",CurrPeriodID-1), Lang("chart_item_desc_a1"), c.NumToStr("LKP01_PERIOD_NAME",CurrPeriodID), Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"))
				table_obj.ChartIgnoredColumns="2,4,6,8,9"
			end if
		else

			if BarometerCalcId=1 then
				table_obj.ChartSubtitle=Lang("report_R500_subtitle")
				table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a10"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"))
				table_obj.ChartColumnDescs=Array(Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"), "", Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"))
				table_obj.ChartIgnoredColumns="1,2,3,5,6,7"
			else
				table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a5"), Lang("chart_item_unit_a1"), Lang("chart_item_unit_a1"))
				table_obj.ChartColumnDescs=Array(Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"), "", Lang("chart_item_desc_a1"), Lang("chart_item_desc_a1"))
				table_obj.ChartIgnoredColumns="1,2,4,5"
			end if
		end if

		Set c = nothing

		TempCount=0

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'table_obj.Label="(" & point_array(0) & " " & Lang("chart_item_respondents") & ")|" & ucase(table_obj.RowID) & "|{%.%}%"

		' For BarometerCalcId=1
		' input array:    prev_a,     prev_b, prev_c, curr_a,     curr_b, curr_c
		' after phase 1:  prev_ratio, 0,      0,      curr_ratio, 0,      0
		' after phase 2:  prev_ratio, 0,      0,      curr_ratio, dir,    0

		' For BarometerCalcId=0
		' input array:    prev_a,     prev_b, curr_a,     curr_b
		' after phase 1:  prev_ratio, 0,      curr_ratio, 0
		' after phase 2:  prev_ratio, 0,      curr_ratio, dir

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

			if (point_array(3)>point_array(4)) then point_array(3)=-SafeDivide(point_array(3),point_array(4)+point_array(5)) else point_array(3)=SafeDivide(point_array(4),point_array(3)+point_array(5))
			point_array(3)=FormatNumber(point_array(3), 1)
			point_array(4)="" ' unused
			point_array(5)="-" ' unused

		else

			' --- phase 1 ---
			if (point_array(0)>=point_array(1)) then point_array(0)=SafeDivide(point_array(0),point_array(1)) else point_array(0)=-SafeDivide(point_array(1),point_array(0))
			point_array(0)=FormatNumber(point_array(0), 1)
			point_array(1)="-" ' unused

			if (point_array(2)>=point_array(3)) then point_array(2)=SafeDivide(point_array(2),point_array(3)) else point_array(2)=-SafeDivide(point_array(3),point_array(2))
			point_array(2)=FormatNumber(point_array(2), 1)
			point_array(3)="-" ' unused

		end if

		if BarometerTimeID=1 then
			if BarometerCalcId=1 then
				Call NotImplemented()
			else
				if (point_array(4)>=point_array(5)) then point_array(4)=SafeDivide(point_array(4),point_array(5)) else point_array(4)=-SafeDivide(point_array(5),point_array(4))
				point_array(4)=FormatNumber(point_array(4), 1)
				point_array(5)="-" ' unused

				if (point_array(6)>=point_array(7)) then point_array(6)=SafeDivide(point_array(6),point_array(7)) else point_array(6)=-SafeDivide(point_array(7),point_array(6))
				point_array(6)=FormatNumber(point_array(6), 1)
				point_array(7)="-" ' unused
			end if
		end if

		' --- phase 2 ---
		if BarometerCalcId=1 then
			point_array(last_index-1)=DiffToSgn(point_array(last_index-3), point_array(last_index-6), 1)
		else
			point_array(last_index-1)=DiffToSgn(point_array(last_index-2), point_array(last_index-4), 1)
		end if

		if TypeName(table_obj)<>"cTableXmlchart" then
			point_array(last_index-1)=NumToDir(point_array(last_index-1))
		end if


		s = table_obj.RowOpen

		for i=0 to ubound(point_array)

			if BarometerCalcId=1 then
				if i=0 or i=3 or i=6 or i=9 then
					table_obj.Label=point_array(i)
					if BarometerTimeID=0 then table_obj.Symbol=point_array(last_index-1)
				end if
			else
				if i=0 or i=2 or i=4 or i=6 then
					table_obj.Label=point_array(i)
					if BarometerTimeID=0 then table_obj.Symbol=point_array(last_index-1)
				end if
			end if

			s = s & table_obj.Cell(point_array(i))
		next

		s = s & table_obj.RowClose


		redim preserve TempArr1(TempCount)
		redim preserve TempArr2(TempCount)
		if BarometerCalcId=1 then
			TempArr1(TempCount)=CDbl(point_array(last_index-3))
		else
			TempArr1(TempCount)=CDbl(point_array(last_index-2))
		end if
		TempArr2(TempCount)=s
		TempCount=TempCount+1

		PointCallback=""

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		if BarometerCalcId=1 then
			max_nr_of_rows=32
		else
			max_nr_of_rows=30
		end if

		if BarometerTimeID=0 then Call Sort(TempArr1, TempArr2, TempCount)

		read_body=""

		exclude_min=-1
		exclude_max=-1
		i_from=0 : i_to=TempCount-1 : i_step=1

		if BarometerTimeID=0 then

			i_from=TempCount-1 : i_to=0 : i_step=-1

			if TypeName(table_obj)="cTableXmlchart" then
				if TempCount>max_nr_of_rows then
					exclude_min=max_nr_of_rows/2
					exclude_max=TempCount-1-max_nr_of_rows/2
					'Response.Write "TempCount=" & TempCount & ", exclude_min=" & exclude_min & ", exclude_max=" & exclude_max & "<br>"
				end if
			end if

		end if

		for i=i_from to i_to step i_step

			tmp=TempArr2(i)
			styleid=i Mod 2
			tmp=Replace(tmp, "alternate_0", "alternate_" & styleid)
			tmp=Replace(tmp, "alternate_1", "alternate_" & styleid)

			if i>exclude_max or i<exclude_min then read_body = read_body & tmp
		next


		TempCount=0

		FinalCallback=read_head & read_body & read_foot

	End Function


	Private Function NotImplemented()
		Response.Write "FEATURE NOT IMPLEMENTED (BarometerCalcId=1 when BarometerTimeID=1)<br>"
		Response.End
	End Function

End Class



Class cChartBar_FamilyOfRates

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

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=4
		Title=Lang("report_R000")
		Filename="chart_17.asp?ID=0"
		UnitForYAxis=Lang("chart_item_dollarpersqft")
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		prev_period_id=CurrPeriodID-PrevPeriodDelta

		ChartIDArr=Array(UniqueID, UniqueID, UniqueID)
		ColumnIDArr=Array(1, 1, 1)
		PeriodIDArr=Array(CurrPeriodID, CurrPeriodID, prev_period_id)
		QuestionTypeIDArr=Array(2, 2, 2)
		FilterForValuesArr=Array(f, f, f)
		GroupingTypeArr=Array(1, 5, 2)

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=-1

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.HeadSpan(UnitForYAxis,4) & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		table_obj.ChartColumnUnits=Array(UnitForYAxis, UnitForYAxis, UnitForYAxis, Lang("chart_item_unit_a2"))
		table_obj.ChartColumnDescs=Array("", "", "", Lang("chart_item_desc_a2"))
		table_obj.ChartIgnoredColumns="4"

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'response.Write "UniqueID=" & UniqueID & "<br>"

		dec_places=1
		label_prefix=""

		if UniqueID<>7 then
			dec_places=2
			label_prefix="$"
		end if

		point_array(4)=DiffToSgn(point_array(3), point_array(4), 2)

		table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"
		'table_obj.Extra2=""

		s = table_obj.RowOpen

		for i=1 to ubound(point_array)
			point_array(i)=Round(point_array(i), dec_places)
			point_array(i)=FormatNumber(point_array(i), dec_places)

			table_obj.Label=""
			if i=3 then
				table_obj.Label=label_prefix & point_array(i) : table_obj.Symbol=point_array(4)
			end if

			if TypeName(table_obj)<>"cTableXmlchart" then
				if i=4 then point_array(i)=NumToDir(point_array(i))
			end if

			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartBar_FaceRateForecast ' Face Rate Forecast Bar Chart '

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
	Public OmitCol2And3

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=8
		Title=Lang("report_R300")
		Filename="chart_17.asp?ID=0"
		OmitCol2And3=false
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		prev_period_id=CurrPeriodID-PrevPeriodDelta

		ChartIDArr=Array(UniqueID, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID)
		ColumnIDArr=Array(1, 1,1, 2,2, 3,3, 4,4)
		PeriodIDArr=Array(CurrPeriodID, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id)
		QuestionTypeIDArr=Array(2, 2,2, 2,2, 2,2, 2,2) ' ***** first number used to be 1
		FilterForValuesArr=Array(f, f,f, f,f, f,f, f,f)
		GroupingTypeArr=Array(1, 5,2, 5,2, 5,2, 5,2)

		if OmitCol2And3 then Filename="chart_17.asp?ID=1"

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=-1

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("chart_item_head_c1"), 4) & table_obj.HeadSpan(Lang("chart_item_head_c2"), 4) & table_obj.HeadSpan(Lang("chart_item_head_c3"), 4) & table_obj.HeadSpan(Lang("chart_item_head_c4"), 4) & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"))
		table_obj.ChartColumnDescs=Array("", Lang("chart_item_desc_c1"), "", Lang("chart_item_desc_a2"), "", Lang("chart_item_desc_c2"), "", Lang("chart_item_desc_a2"), "", Lang("chart_item_desc_c3"), "", Lang("chart_item_desc_a2"), "", Lang("chart_item_desc_c4"), "", Lang("chart_item_desc_a2"))
		table_obj.ChartIgnoredColumns="4,8,12,16"

		if OmitCol2And3 then table_obj.ChartIgnoredColumns=table_obj.ChartIgnoredColumns & ",5,6,7,9,10,11"

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		point_array(4)=DiffToSgn(point_array(3), point_array(4), 1)
		point_array(8)=DiffToSgn(point_array(7), point_array(8), 1)
		point_array(12)=DiffToSgn(point_array(11), point_array(12), 1)
		point_array(16)=DiffToSgn(point_array(15), point_array(16), 1)

		table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"
		'table_obj.Extra2=""

		s = table_obj.RowOpen

		for i=1 to ubound(point_array)
			point_array(i)=Round(point_array(i), 2)

			table_obj.Label=""
			if i=3 or i=7 or i=11 or i=15 then
				point_array(i)=FormatNumber(point_array(i), 1) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
			else
				point_array(i)=FormatNumber(point_array(i), 1)
			end if

			if TypeName(table_obj)<>"cTableXmlchart" then
				if i=4 or i=8 or i=12 or i=16 then point_array(i)=NumToDir(point_array(i))
			end if

			if OmitCol2And3 and i>=5 and i<=12 then point_array(i)="-"

			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class


%>
