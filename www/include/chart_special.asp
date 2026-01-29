<%

' DK, Apr 01, 2008 - version 1.0

Class cChartPie_RespondentCategories ' Respondent Categories Pie Chart '

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
		UniqueID=-1
		Title=Lang("report_X100")
		Filename="chart_16.asp?ID=0"
	End Sub

	Public Function Init
		ChartIDArr=Array(UniqueID)
		ColumnIDArr=Array(1)
		PeriodIDArr=Array(CurrPeriodID)
		QuestionTypeIDArr=Array(1)
		FilterForValuesArr=Array("")
		GroupingTypeArr=Array(-1)
	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_nrofrespondents")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle="" 's_property

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'table_obj.Label="(" & point_array(0) & " " & Lang("chart_item_respondents") & ")|" & ucase(table_obj.RowID) & "|{%.%}%"
		table_obj.Label="|{%.%}%|" & ucase(table_obj.RowID)

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



Class cChartBar_Motives ' Motives for Investment and Threats to Investment Bar Chart '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta
	Public LabelsArr

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private TempArr1_a
	Private TempArr1_b
	Private TempArr2_a
	Private TempArr2_b

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=10
		Title=Lang("report_R000")
		Filename="chart_18.asp?ID=0"
		LabelsArr=Array("")
	End Sub

	Public Function Init

		'ChartIDArr=Array(UniqueID,UniqueID,UniqueID,UniqueID,UniqueID,UniqueID,UniqueID,UniqueID,UniqueID)
		'ColumnIDArr=Array(1,2,3,4,5,6,7,8,9)
		'PeriodIDArr=Array(CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID,CurrPeriodID)
		'QuestionTypeIDArr=Array(1,1,1,1,1,1,1,1,1)
		'FilterForValuesArr=Array("","","","","","","","","")
		'GroupingTypeArr=Array(10,10,10,10,10,10,10,10,10)

		f=""
		prev_period_id=CurrPeriodID-4

		j=1+2*Ubound(LabelsArr)

		Redim ChartIDArr(j)
		Redim ColumnIDArr(j)
		Redim PeriodIDArr(j)
		Redim QuestionTypeIDArr(j)
		Redim FilterForValuesArr(j)
		Redim GroupingTypeArr(j)

		k=1
		for i=0 to j step 2

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=k
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=1
			FilterForValuesArr(i)=f
			GroupingTypeArr(i)=10

			ChartIDArr(i+1)=UniqueID
			ColumnIDArr(i+1)=k
			PeriodIDArr(i+1)=prev_period_id
			QuestionTypeIDArr(i+1)=1
			FilterForValuesArr(i+1)=f
			GroupingTypeArr(i+1)=10

			k=k+1
		next

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		redim TempArr1_a(ubound(LabelsArr))
		redim TempArr1_b(ubound(LabelsArr))
		redim TempArr2_a(ubound(LabelsArr))
		redim TempArr2_b(ubound(LabelsArr))

		if TypeName(table_obj)<>"cTableXmlchart" then
			'InitCallback=table_obj.RowOpen & table_obj.Head("Resp1") & table_obj.Head("Resp2") & table_obj.Head("Resp3") & table_obj.Head("Resp4") & table_obj.Head("Resp5") & table_obj.Head("Resp6") & table_obj.Head("Resp7") & table_obj.Head("Resp8") & table_obj.Head("Resp9") & table_obj.RowClose
			InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_d3")) & table_obj.Head(Lang("chart_item_change")) & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=Lang("report_S10X_subtitle")

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a8"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
		table_obj.ChartColumnDescs=Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
		table_obj.ChartIgnoredColumns="2"

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		'table_obj.Label="(" & point_array(0) & " " & Lang("chart_item_respondents") & ")|" & ucase(table_obj.RowID) & "|{%.%}%"

		k=0
		for i=0 to ubound(point_array) step 2

			TempArr1_a(k)=TempArr1_a(k)+point_array(i)
			TempArr1_b(k)=TempArr1_b(k)+point_array(i+1)

			if table_obj.RowID="4" or table_obj.RowID="5" then
				TempArr2_a(k)=TempArr2_a(k)+point_array(i)
				TempArr2_b(k)=TempArr2_b(k)+point_array(i+1)
			end if

			k=k+1
		next

		PointCallback = ""

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		tmp=LabelsArr

		for i=0 to ubound(TempArr1_a)
			val_a=100*SafeDivide(TempArr2_a(i),TempArr1_a(i))
			val_b=100*SafeDivide(TempArr2_b(i),TempArr1_b(i))
			TempArr1_a(i)=i & "|" & val_b
			TempArr2_a(i)=val_a
		next

		Call Sort(TempArr2_a, TempArr1_a, ubound(TempArr1_a)+1)

		read_body=""
		for i=ubound(TempArr1_a) to 0 step -1

			table_obj.RowID=""

			s=Split(TempArr1_a(i), "|")
			k=s(0)
			val_b=s(1)

			s=Split(tmp(k), "|")
			table_obj.Extra=s(0)
			if ubound(s)=1 then table_obj.Extra2=s(1)

			val_a=TempArr2_a(i)

			val_b=DiffToSgn(val_a, val_b, 1)

			if TypeName(table_obj)<>"cTableXmlchart" then
				val_b=NumToDir(val_b)
			end if

			val_a=FormatNumber(Round(val_a,1), 1)

			table_obj.Label=val_a
			table_obj.Symbol=val_b

			read_body=read_body & table_obj.RowOpen
			read_body=read_body & table_obj.Cell(val_a)
			read_body=read_body & table_obj.Cell(val_b)
			read_body=read_body & table_obj.RowClose

			TempArr1_a(i)=0
			TempArr1_b(i)=0
			TempArr2_a(i)=0
			TempArr2_b(i)=0

		next

		FinalCallback=read_head & read_body & read_foot
	End Function

End Class



Class cChartBar_CSCMFamilyA ' Availability of New Real Estate Equity, Cost of Debt Outlook, Capital Source '

	Public UniqueID
	Public Title
	Public Subtitle
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta
	Public LabelsArr

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private FlavourID

	Private	TempCount
	Private TempArr1
	Private TempArr2
	Private TempSum

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=14
		Title=Lang("report_R000")
		Subtitle=Lang("report_R000_subtitle")
		Filename="chart_18.asp"
		LabelsArr=Array("")
		SetFlavourID(0)
	End Sub

	Public Function Init

		'ChartIDArr=Array(UniqueID,UniqueID,UniqueID)
		'ColumnIDArr=Array(1,2,3)
		'PeriodIDArr=Array(CurrPeriodID,CurrPeriodID,CurrPeriodID)
		'QuestionTypeIDArr=Array(1,1,1)
		'FilterForValuesArr=Array("","","")
		'GroupingTypeArr=Array(11,11,11)

		f=""
		prev_period_id=CurrPeriodID-2

		if FlavourID=2 then
			grouping_type=11
		elseif FlavourID=1 then
			grouping_type=12
		else
			grouping_type=13
		end if

		j=1+2*Ubound(LabelsArr)

		Redim ChartIDArr(j)
		Redim ColumnIDArr(j)
		Redim PeriodIDArr(j)
		Redim QuestionTypeIDArr(j)
		Redim FilterForValuesArr(j)
		Redim GroupingTypeArr(j)

		k=1
		for i=0 to j step 2

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=k
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=1
			FilterForValuesArr(i)=f
			GroupingTypeArr(i)=grouping_type

			ChartIDArr(i+1)=UniqueID
			ColumnIDArr(i+1)=k
			PeriodIDArr(i+1)=prev_period_id
			QuestionTypeIDArr(i+1)=1
			FilterForValuesArr(i+1)=f
			GroupingTypeArr(i+1)=grouping_type

			k=k+1
		next

		TempCount=0
		Redim TempArr1(0)
		Redim TempArr2(0)
		TempSum=Array(0,0, 0,0, 0,0, 0,0)

	End Function

	Public Function SetFlavourID(ID)

		if ID=2 then
			FlavourID=2
			Filename="chart_18.asp?ID=2"
		elseif ID=1 then
			FlavourID=1
			Filename="chart_18.asp?ID=1"
		else
			FlavourID=0
			Filename="chart_18.asp?ID=0"
		end if

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		j=ubound(LabelsArr)

		if TypeName(table_obj)<>"cTableXmlchart" then

			if FlavourID=0 then
				InitCallback=table_obj.RowOpen & table_obj.Head(Lang("chart_item_head_d2")) & table_obj.Head(Lang("chart_item_change")) & table_obj.RowClose
			else

				InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("chart_item_head_d2"),2+2*j) & table_obj.RowClose & table_obj.RowOpen
				for i=0 to j
					InitCallback=InitCallback & table_obj.Head(LabelsArr(i)) & table_obj.Head(Lang("chart_item_change"))
				next
				InitCallback=InitCallback & table_obj.RowClose
			end if
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=Subtitle

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a8"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a8"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a8"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a8"), Lang("chart_item_unit_a2"))
		table_obj.ChartColumnDescs=Array("", "", "", "", "", "", "", "")
		table_obj.ChartIgnoredColumns="2,4,6,8"

		for i=0 to j
			table_obj.ChartColumnDescs(i*2)=ucase(LabelsArr(i))
		next

		InitCallback=""

		TempCount=0
		TempSum=Array(0,0, 0,0, 0,0, 0,0)

	End Function

	Public Function PointCallback(table_obj, point_array)

		redim preserve TempArr1(TempCount)
		redim preserve TempArr2(TempCount)

		if FlavourID=0 then
			col_sum_a=0
			col_sum_b=0
			for i=0 to ubound(point_array) step 2
				col_sum_a=col_sum_a+point_array(i)
				col_sum_b=col_sum_b+point_array(i+1)
			next
			for i=0 to ubound(point_array) step 2
				point_array(i)=col_sum_a
				point_array(i+1)=col_sum_b
			next
		end if

		TempArr1(TempCount)=Join(point_array, "|")
		TempArr2(TempCount)=table_obj.RowID

		for i=0 to ubound(point_array)
			TempSum(i)=TempSum(i)+point_array(i)
		next

		TempCount=TempCount+1

		PointCallback = ""

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		read_body=""

		for j=0 to TempCount-1

			point_array=Split(TempArr1(j), "|")

			if FlavourID=0 then

				val0=100*SafeDivide(point_array(0), TempSum(0))
				val1=100*SafeDivide(point_array(1), TempSum(1))

				'diff=val1
				diff=DiffToSgn(val0, val1, 1)

				TempArr1(j)=val0
				TempArr2(j)=TempArr2(j) & "|" & diff

			else

				for i=0 to ubound(point_array)
					val=point_array(i)
					val=100*SafeDivide(val, TempSum(i))
					point_array(i)=val
				next

				table_obj.RowID=TempArr2(j)

				read_body = read_body & table_obj.RowOpen

				for i=0 to ubound(point_array)

					if i=0 or i=2 or i=4 or i=6 then
						point_array(i+1)=DiffToSgn(point_array(i), point_array(i+1), 1)
						point_array(i)=FormatNumber(point_array(i), 1) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
					else
						point_array(i)=FormatNumber(point_array(i), 1)
					end if

					if TypeName(table_obj)<>"cTableXmlchart" then
						if i=1 or i=3 or i=5 or i=7 then point_array(i)=NumToDir(point_array(i))
					end if

					read_body = read_body & table_obj.Cell(point_array(i))

				next

				read_body=read_body & table_obj.RowClose

			end if

		next


		'table_obj.RowID="TOTAL"
		'read_body = read_body & table_obj.RowOpen
		'for i=0 to ubound(point_array)
		'	read_body = read_body & table_obj.Cell(TempSum(i))
		'	'Response.Write TempSum(i) & " | "
		'next
		'read_body=read_body & table_obj.RowClose


		if FlavourID=0 then

			Call Sort(TempArr1, TempArr2, ubound(TempArr2)+1)

			read_body=""

			for i=ubound(TempArr2) to 0 step -1

				val=FormatNumber(TempArr1(i), 1)
				s=Split(TempArr2(i), "|")

				table_obj.RowID=s(0)
				table_obj.Label=val
				table_obj.Symbol=s(1)

				read_body=read_body & table_obj.RowOpen
				read_body=read_body & table_obj.Cell(val)

				if TypeName(table_obj)<>"cTableXmlchart" then
					s(1)=NumToDir(s(1))
				end if

				read_body=read_body & table_obj.Cell(s(1))

'				read_body=read_body & table_obj.Cell("")
'				read_body=read_body & table_obj.Cell("")

				read_body=read_body & table_obj.RowClose

				TempArr1(i)=0
				TempArr2(i)=0

			next

		end if


		TempCount=0
		TempSum=Array(0,0, 0,0, 0,0, 0,0)

		FinalCallback=read_head & read_body & read_foot

	End Function

End Class



Class cChartBar_CSCMFamilyB ' Expected Price Performance of Real Estate Equities, Total Required Rate of Return '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta
	Public LabelsArr

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=15
		Title=Lang("report_R000")
		Filename="chart_17.asp?ID=0"
		LabelsArr=Array("")
	End Sub

	Public Function Init

		f=""
		prev_period_id=CurrPeriodID-2

		'ChartIDArr=Array(UniqueID,UniqueID, UniqueID,UniqueID, UniqueID,UniqueID)
		'ColumnIDArr=Array(1,1, 2,2, 3,3)
		'PeriodIDArr=Array(CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id)
		'QuestionTypeIDArr=Array(2,2, 2,2, 2,2)
		'FilterForValuesArr=Array(f,f, f,f, f,f)
		'GroupingTypeArr=Array(55,52, 55,52, 55,52)

		j=1+2*Ubound(LabelsArr)

		Redim ChartIDArr(j)
		Redim ColumnIDArr(j)
		Redim PeriodIDArr(j)
		Redim QuestionTypeIDArr(j)
		Redim FilterForValuesArr(j)
		Redim GroupingTypeArr(j)

		k=1
		for i=0 to j step 2

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=k
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=2
			FilterForValuesArr(i)=f
			GroupingTypeArr(i)=55

			ChartIDArr(i+1)=UniqueID
			ColumnIDArr(i+1)=k
			PeriodIDArr(i+1)=prev_period_id
			QuestionTypeIDArr(i+1)=2
			FilterForValuesArr(i+1)=f
			GroupingTypeArr(i+1)=52

			k=k+1
		next

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		j=Ubound(LabelsArr)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen
			for i=0 to j
				InitCallback=InitCallback & table_obj.HeadSpan(LabelsArr(i),4)
			next
			InitCallback=InitCallback & table_obj.RowClose & table_obj.RowOpen
			for i=0 to j
				InitCallback=InitCallback & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change"))
			next
			InitCallback=InitCallback & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=Lang("report_S30X_subtitle")

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"))
		table_obj.ChartColumnDescs=Array("", "", "", Lang("chart_item_unit_a2"), "", "", "", Lang("chart_item_unit_a2"), "", "", "", Lang("chart_item_unit_a2"))
		table_obj.ChartIgnoredColumns="4,8,12"

		for i=0 to j
			table_obj.ChartColumnDescs(1+i*4)=ucase(LabelsArr(i))
		next

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		dec_places=1

		'table_obj.Extra2=""

		s = table_obj.RowOpen

		for i=0 to ubound(point_array)

			'point_array(i)=Round(point_array(i), 2)

			table_obj.Label=""
			if i=2 or i=6 or i=10 then
				point_array(i+1)=DiffToSgn(point_array(i), point_array(i+1), 1)
				point_array(i)=FormatNumber(point_array(i), dec_places) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
			else
				point_array(i)=FormatNumber(point_array(i), dec_places)
			end if

			if TypeName(table_obj)<>"cTableXmlchart" then
				if i=3 or i=7 or i=11 then point_array(i)=NumToDir(point_array(i))
			end if

			s = s & table_obj.Cell(point_array(i))
		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)

		FinalCallback=read_head & read_body & read_foot

	End Function

End Class



Class cChartBar_OCR_IRR ' OCR and IRR Bar Chart '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta
	Public ColumnID

	Public ChartIDArr
	Public ColumnIDArr
	Public PeriodIDArr
	Public QuestionTypeIDArr
	Public FilterForValuesArr
	Public GroupingTypeArr
	Public PropertyArr

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=19
		Title=Lang("report_R000")
		Filename="chart_17.asp?ID=0"
		ColumnID=-1
	End Sub

	Public Function Init

		'f=" AND ri.Value<>0" ' **** is this necessary?
		f=""
		prev_period_id=CurrPeriodID-1

		ChartIDArr=Array(2,2, 2,2, 2,2, 2,2)
		ColumnIDArr=Array(ColumnID,ColumnID, ColumnID,ColumnID, ColumnID,ColumnID, ColumnID,ColumnID)
		PeriodIDArr=Array(CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id, CurrPeriodID,prev_period_id)
		QuestionTypeIDArr=Array(2,2, 2,2, 2,2, 2,2) ' ***** last two numbers used to be 1
		FilterForValuesArr=Array(f,f, f,f, f,f, f,f)
		GroupingTypeArr=Array(5,2, 5,2, 5,2, 5,2)
		PropertyArr=Array(3,3, 7,7, 13,13, 15,15)

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen & table_obj.HeadSpan(Lang("L_DOWNTOWN_CLASS_AA_OFFICE"),4) & table_obj.HeadSpan(Lang("L_TIER_I_REGIONAL_MALL"),4) & table_obj.HeadSpan(Lang("L_SINGLE_TENANT_INDUSTRIAL"),4) & table_obj.HeadSpan(Lang("L_SUBURBAN_MULTI_UNIT_RESIDENTIAL"),4)
			InitCallback=InitCallback & table_obj.RowClose
			InitCallback=InitCallback & table_obj.RowOpen & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change")) & table_obj.Head(Lang("chart_item_max")) & table_obj.Head(Lang("chart_item_min")) & table_obj.Head(Lang("chart_item_avg")) & table_obj.Head(Lang("chart_item_change"))
			InitCallback=InitCallback & table_obj.RowClose
			exit function
		end if

		table_obj.ChartTitle=Title
		table_obj.ChartTitle2=CurrPeriodLabel
		table_obj.ChartSubtitle=s_property

		table_obj.ChartColumnUnits=Array(Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a4"), Lang("chart_item_unit_a2"))
		table_obj.ChartColumnDescs=Array("", Lang("L_DOWNTOWN_CLASS_AA_OFFICE"), "", Lang("chart_item_desc_a2"), "", Lang("L_TIER_I_REGIONAL_MALL"), "", Lang("chart_item_desc_a2"), "", Lang("L_SINGLE_TENANT_INDUSTRIAL"), "", Lang("chart_item_desc_a2"), "", Lang("L_SUBURBAN_MULTI_UNIT_RESIDENTIAL"), "", Lang("chart_item_desc_a2"))
		table_obj.ChartIgnoredColumns="4,8,12,16"

		InitCallback=""

	End Function

	Public Function PointCallback(table_obj, point_array)

		dec_places=1

		point_array(3)=DiffToSgn(point_array(2), point_array(3), 1)
		point_array(7)=DiffToSgn(point_array(6), point_array(7), 1)
		point_array(11)=DiffToSgn(point_array(10), point_array(11), 1)
		point_array(15)=DiffToSgn(point_array(14), point_array(15), 1)

		s = table_obj.RowOpen

		for i=0 to ubound(point_array)
			point_array(i)=Round(point_array(i), 2)

			table_obj.Label=""
			if i=2 or i=6 or i=10 or i=14 then
				point_array(i)=FormatNumber(point_array(i), dec_places) : table_obj.Label=point_array(i) : table_obj.Symbol=point_array(i+1)
			else
				point_array(i)=FormatNumber(point_array(i), dec_places)
			end if

			if TypeName(table_obj)<>"cTableXmlchart" then
				if i=3 or i=7 or i=11 or i=15 then point_array(i)=NumToDir(point_array(i))
			end if

			s = s & table_obj.Cell(point_array(i))

		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class


%>
