<%

' DK/CS, Feb 05, 2009 - version 1.0

Class cChartTable_CDMParameters_Dynamic ' Conventional Debt Market Parameters Chart (Dynamic Version) '

	Public UniqueID
	Public Title
	Public Filename
	Public CurrPeriodID
	Public CurrPeriodLabel
	Public PrevPeriodDelta

	Public ChartIDArr
	Public ColumnIDArr, ColumnsArr, LabelsArr
	Public QuestionTypeIDArr, TypesArr
	Public FilterForValuesArr, FilterArr
	Public GroupingTypeArr, GroupingArr
	Public PeriodIDArr

	Public Units
	Public DecPlaces
	Public RespondentsGroupingType
	Public RespondentsDisplayFormat
	Public RespondentsQuestionType

	Private Sub Class_Initialize()
		PrevPeriodDelta=1
		UniqueID=29
		Title=""
		Filename=""
		LabelsArr=Array("")
		ColumnsArr=Array("")
		TypesArr=Array("")
		FilterArr=Array("")
		GroupingArr=Array("")
		Units=Lang("chart_item_percent")
		DecPlaces=1
		RespondentsGroupingType=0
		RespondentsQuestionType=2
		RespondentsDisplayFormat=0
	End Sub

	Public Function Init

		f=""
		nr_of_cols=Ubound(LabelsArr)

		j=(nr_of_cols+1)

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
		QuestionTypeIDArr(i)=RespondentsQuestionType
		FilterForValuesArr(i)=f
		GroupingTypeArr(i)=RespondentsGroupingType

		for i=1 to j

			ChartIDArr(i)=UniqueID
			ColumnIDArr(i)=ColumnsArr(i-1)
			PeriodIDArr(i)=CurrPeriodID
			QuestionTypeIDArr(i)=TypesArr(i-1)
			FilterForValuesArr(i)=FilterArr(i-1)
			GroupingTypeArr(i)=GroupingArr(i-1)
		next

	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)

		cols_to_add=-1
		nr_of_cols=Ubound(LabelsArr)+1

		if TypeName(table_obj)<>"cTableXmlchart" then
			InitCallback=table_obj.RowOpen
			for col=1 to nr_of_cols
				InitCallback = InitCallback & table_obj.Head(LabelsArr(col-1))
			next
			InitCallback=InitCallback & table_obj.RowClose
			exit function
		end if

	End Function

	Public Function PointCallback(table_obj, point_array)

		ubound_point_array=ubound(point_array)

		if UniqueID=29 then

			last_col=ubound_point_array
			If (CurrPeriodID >= 108) Then last_col = last_col - 2

			' express the last three columns in percentages
			point_array(last_col-3)=Round(SafeDivide(100*point_array(last_col-3), point_array(last_col)), DecPlaces)
			point_array(last_col-2)=Round(SafeDivide(100*point_array(last_col-2), point_array(last_col)), DecPlaces)
			point_array(last_col-1)=Round(SafeDivide(100*point_array(last_col-1), point_array(last_col)), DecPlaces)
			' the last column is a sum of previous two
			point_array(last_col)=point_array(last_col-1)+point_array(last_col-2)

		end if

		if RespondentsDisplayFormat=1 then
			table_obj.Extra="(" & point_array(0) & chr(160) & Lang("chart_item_respondents") & ")"
		end if

		s = table_obj.RowOpen

		for i=1 to ubound_point_array

			dec_places=DecPlaces
			'if i=2 then dec_places=2 ' Commented by DK on Sep 11, 2014 as per Kevin's info: 83CSDM - Conventional Debt Market Parameters has a column that is no longer part of the question; 

			point_array(i)=Round(point_array(i), dec_places)

			table_obj.Label=""
			point_array(i)=FormatNumber(point_array(i), dec_places)

			s = s & table_obj.Cell(point_array(i))

		next

		PointCallback = s & table_obj.RowClose

	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class

%>
