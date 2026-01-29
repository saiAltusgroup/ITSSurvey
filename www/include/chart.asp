<!--#INCLUDE file="chart_standard.asp"-->
<!--#INCLUDE file="chart_special.asp"-->
<!--#INCLUDE file="chart_historical.asp"-->
<!--#INCLUDE file="chart_tabular.asp"-->

<%

' DK, Apr 01, 2008 - version 1.0

Class cChartUnknown

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
		Title="Unknown Chart"
		Filename=""
	End Sub

	Public Function Init
		ChartIDArr=Array(UniqueID)
		ColumnIDArr=Array(0)
		PeriodIDArr=Array(CurrPeriodID)
		QuestionTypeIDArr=Array(1)
		FilterForValuesArr=Array("")
		GroupingTypeArr=Array(12)
	End Function

	Public Function InitCallback(table_obj, s_property, cols_to_add)
		InitCallback=""
	End Function

	Public Function PointCallback(table_obj, point_array)
	End Function

	Public Function FinalCallback(table_obj, read_head, read_body, read_foot)
		FinalCallback=read_head & read_body & read_foot
	End Function

End Class

Function DiffToSgn(pa, pb, dec_places)

		pa_round=Round(pa, dec_places)
		pb_round=Round(pb, dec_places)
		diff=Round(pa_round-pb_round, 4)

		'Response.Write "Current=" & pa_round & ", Previous=" & pb_round & ", Difference=" & diff
		'if Abs(diff)<0.1 then Response.Write " *** EQUAL since ABS(diff) is less than 0.1 *** " else Response.Write " *** DiffToSgn=" & Sgn(diff)
		'Response.Write "<br />"

		' NOTE: On Dec 16, 2008 Kevin requested the following change:
		'
		' "Also, for generating an up or down arrow on the charts which show direction of change,
		'  we need to add a rule that the change has be greater than one-tenth of a percent to generate
		'  a change arrow rather than a no change."
		'
		' He later confirmed that one-tenth is not a percent, just absolute value.
		' So we should insert the following lines:

		DiffToSgn=0
		if Abs(diff)<0.1 then exit function
		' The treshold used to be 0.5 but after speaking with Kevin on Dec 11, 2014 we changed it to 0.1

		' In other words, if the difference is less than 0.1 then we should treat it as zero.

		DiffToSgn=Sgn(diff)

End Function


Function NumToDir(num)

	NumToDir=Lang("dir_same")
	if num<0 then
		NumToDir=Lang("dir_down")
	elseif num>0 then
		NumToDir=Lang("dir_up")
	end if

End Function

Function GetLabelPart(label, separator, part_index)
	GetLabelPart=label
	tmp_arr=Split(label, separator)
	if part_index>=0 and part_index<=ubound(tmp_arr) then GetLabelPart=tmp_arr(part_index)
End Function

%>
